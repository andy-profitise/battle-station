"""
Main processing loop for the Vendor Command Center.

This is the brain that ties everything together:
1. Fills the queue from Gmail (live emails)
2. Accepts interrupts from Telegram/Teams
3. Processes work items one at a time
4. Falls back to proactive scanning when queue is empty
"""
from __future__ import annotations

import logging
import time
from datetime import datetime
from typing import Any

from agents.context_builder import ContextBuilder
from agents.reasoner import Reasoner
from agents.action_executor import ActionExecutor, ExecutionResult
from agents.proactive_scanner import ProactiveScanner
from core.models import VendorContext, WorkItem, WorkItemSource
from core.vendor_resolver import VendorResolver
from integrations.gmail import GmailClient
from integrations.monday import MondayClient
from integrations.calendar_api import CalendarClient
from integrations.box_api import BoxClient
from integrations.airtable import AirtableClient
from integrations.telegram_bot import TelegramListener
from integrations.teams_bot import TeamsListener
from queue.priority_queue import VendorWorkQueue
from config import settings

logger = logging.getLogger(__name__)


class CommandCenter:
    """
    The main orchestration engine.

    Lifecycle:
    1. initialize() - set up all clients and load vendor map
    2. run() - start the processing loop
    3. shutdown() - clean up
    """

    def __init__(self):
        # Clients
        self.gmail = GmailClient()
        self.monday = MondayClient()
        self.calendar = CalendarClient()
        self.box = BoxClient()
        self.airtable = AirtableClient()

        # Agents
        self.context_builder = ContextBuilder(
            self.gmail, self.monday, self.calendar, self.box, self.airtable
        )
        self.reasoner = Reasoner()
        self.executor = ActionExecutor(self.gmail, self.monday)
        self.scanner = ProactiveScanner(self.monday)

        # Infrastructure
        self.queue = VendorWorkQueue()
        self.resolver = VendorResolver()
        self.telegram: TelegramListener | None = None
        self.teams: TeamsListener | None = None

        # State
        self._running = False
        self._processed_threads: set[str] = set()
        self._cycle_count = 0

    def initialize(self) -> None:
        """Initialize all systems and load vendor data."""
        logger.info("Initializing Command Center...")

        # Load vendor map from Monday.com
        self._load_vendor_map()

        # Start message listeners
        self._start_listeners()

        logger.info(
            f"Command Center initialized: "
            f"{self.resolver.vendor_count} vendors loaded"
        )

    def _load_vendor_map(self) -> None:
        """Load vendor names and slugs from Monday.com boards."""
        vendors = []

        for board_id in [settings.MONDAY_BUYERS_BOARD_ID, settings.MONDAY_AFFILIATES_BOARD_ID]:
            if not board_id:
                continue
            try:
                items = self.monday.get_all_vendors(board_id)
                for item in items:
                    slug = self.resolver.slug_for_name(item.name)
                    vendors.append({"name": item.name, "slug": slug})
            except Exception as e:
                logger.error(f"Failed to load vendors from board {board_id}: {e}")

        self.resolver.load_vendor_map(vendors)

    def _start_listeners(self) -> None:
        """Start Telegram and Teams listeners for ad-hoc items."""
        # Telegram
        self.telegram = TelegramListener(on_message=self._handle_adhoc_message)
        self.telegram.start()

        # Teams
        self.teams = TeamsListener(on_message=self._handle_adhoc_message)
        self.teams.start()

    def _handle_adhoc_message(self, vendor_name: str, message: str) -> None:
        """Callback for ad-hoc messages from Telegram/Teams."""
        resolved_name, slug = self.resolver.resolve_from_name(vendor_name)
        if not resolved_name:
            resolved_name = vendor_name
            slug = self.resolver.slug_for_name(vendor_name)

        self.queue.push_urgent(resolved_name, message, WorkItemSource.TELEGRAM)
        logger.info(f"Ad-hoc item queued: {resolved_name} - {message[:50]}")

    def run(self) -> None:
        """
        Main processing loop.

        1. Scan Gmail for live emails → fill queue
        2. Process queue items (highest priority first)
        3. When queue is empty → proactive scan
        4. Repeat
        """
        self._running = True
        logger.info("Command Center loop starting...")

        while self._running:
            self._cycle_count += 1
            logger.info(f"\n{'='*60}\nCycle {self._cycle_count} starting\n{'='*60}")

            try:
                # Phase 1: Fill queue from Gmail
                email_count = self._scan_gmail()
                logger.info(f"Found {email_count} new email work items")

                # Phase 2: Process the queue
                if not self.queue.is_empty:
                    self._process_queue()
                else:
                    logger.info("Queue empty - entering proactive mode")

                    # Phase 3: Proactive scan if enabled
                    if settings.PROACTIVE_SCAN_ENABLED:
                        self._run_proactive_scan()

                    if self.queue.is_empty:
                        # Nothing to do - wait before next cycle
                        logger.info("All clear. Waiting 60 seconds before next scan...")
                        self._sleep_interruptible(60)

            except KeyboardInterrupt:
                logger.info("Interrupted by user")
                break
            except Exception as e:
                logger.error(f"Error in main loop: {e}", exc_info=True)
                time.sleep(10)

        self.shutdown()

    def _scan_gmail(self) -> int:
        """Scan Gmail for live vendor emails and add them to the queue."""
        try:
            threads = self.gmail.get_live_vendor_emails()
        except Exception as e:
            logger.error(f"Failed to scan Gmail: {e}")
            return 0

        count = 0
        for thread in threads:
            thread_id = thread["thread_id"]

            # Skip already-processed threads
            if thread_id in self._processed_threads:
                continue

            # Resolve vendor from thread labels
            label_names = thread.get("labels", [])
            vendor_name, vendor_slug = self.resolver.resolve_from_labels(label_names)

            if not vendor_name:
                # Try resolving from sender email
                from_addr = thread.get("from", "")
                vendor_name, vendor_slug = self.resolver.resolve_from_email(from_addr)

            if not vendor_name:
                vendor_name = f"Unknown ({thread.get('from', 'no sender')})"
                vendor_slug = "unknown"

            # Determine if inbox or just received
            is_inbox = "INBOX" in label_names

            self.queue.push_email(
                vendor_name=vendor_name,
                vendor_slug=vendor_slug or "unknown",
                thread_id=thread_id,
                message_date=datetime.now(),  # Could parse from thread
                is_inbox=is_inbox,
            )
            count += 1

        return count

    def _process_queue(self) -> None:
        """Process items from the queue until it's empty or interrupted."""
        while not self.queue.is_empty and self._running:
            item = self.queue.pop()
            if not item:
                break

            logger.info(
                f"\n--- Processing: {item.vendor_name} "
                f"[{item.priority.name}] [{item.source.value}] ---"
            )

            try:
                result = self._process_work_item(item)
                self.queue.mark_processed(item.id)
                if item.gmail_thread_id:
                    self._processed_threads.add(item.gmail_thread_id)

                logger.info(f"Result: {result.summary()}")

                # Check for urgent interrupts between items
                if self.queue.has_urgent():
                    logger.info("Urgent item detected - processing next")

            except Exception as e:
                logger.error(f"Error processing {item.vendor_name}: {e}", exc_info=True)

    def _process_work_item(self, item: WorkItem) -> ExecutionResult:
        """
        Process a single work item through the full pipeline:
        1. Build vendor context (before reading the trigger)
        2. Read the trigger (email/message)
        3. Reason and plan
        4. Execute actions
        """
        vendor_name = item.vendor_name or "Unknown"
        vendor_slug = item.vendor_slug or "unknown"

        # Step 1: Build vendor context BEFORE reading the email
        logger.info(f"  Step 1: Building context for {vendor_name}...")
        vendor_context = self.context_builder.build(vendor_name, vendor_slug)

        # Step 2: Read the trigger
        trigger_email = None
        if item.gmail_thread_id:
            logger.info(f"  Step 2: Reading trigger email...")
            trigger_email = self.gmail.get_thread_full(item.gmail_thread_id)

        # Step 3: Reason and plan
        logger.info(f"  Step 3: Analyzing with Claude...")
        action_plan = self.reasoner.analyze_and_plan(item, vendor_context, trigger_email)

        logger.info(f"  Assessment: {action_plan.assessment.summary}")
        logger.info(f"  Urgency: {action_plan.assessment.urgency}")
        logger.info(f"  Actions: {len(action_plan.actions)}")

        # Step 4: Execute
        logger.info(f"  Step 4: Executing {len(action_plan.actions)} actions...")
        result = self.executor.execute(action_plan)

        return result

    def _run_proactive_scan(self) -> None:
        """Run proactive vendor scan and queue any findings."""
        logger.info("Running proactive vendor scan...")

        # Build context for all vendors (lightweight version)
        all_vendors = self.resolver.all_vendors()
        contexts = []

        for name, slug in all_vendors[:50]:  # Limit to avoid API rate limits
            try:
                ctx = self.context_builder.build(name, slug)
                contexts.append(ctx)
            except Exception as e:
                logger.debug(f"Skipping {name} in proactive scan: {e}")

        # Scan for proactive opportunities
        work_items = self.scanner.scan_all_vendors(contexts)

        for item in work_items:
            self.queue.push(item)

        logger.info(f"Proactive scan queued {len(work_items)} items")

    def _sleep_interruptible(self, seconds: int) -> None:
        """Sleep that can be interrupted by urgent queue items."""
        for _ in range(seconds):
            if not self._running or self.queue.has_urgent():
                return
            time.sleep(1)

    def shutdown(self) -> None:
        """Clean shutdown."""
        self._running = False
        if self.telegram:
            self.telegram.stop()
        if self.teams:
            self.teams.stop()
        logger.info("Command Center shut down")

    def process_single_vendor(self, vendor_name: str) -> ExecutionResult | None:
        """
        Process a single vendor on-demand (for testing or manual trigger).
        """
        resolved_name, slug = self.resolver.resolve_from_name(vendor_name)
        if not resolved_name:
            resolved_name = vendor_name
            slug = self.resolver.slug_for_name(vendor_name)

        item = WorkItem(
            source=WorkItemSource.TELEGRAM,
            vendor_name=resolved_name,
            vendor_slug=slug,
            message_text=f"Manual processing requested for {resolved_name}",
        )

        return self._process_work_item(item)
