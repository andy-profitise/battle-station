"""
Context Builder - gathers everything we know about a vendor BEFORE reading the trigger.

This is the "look before you read" step. We want Claude to understand the full
vendor picture so it can properly contextualize whatever email/message triggered
this work item.
"""
from __future__ import annotations

import logging
import os
from datetime import datetime
from pathlib import Path

from core.models import VendorContext
from integrations.gmail import GmailClient
from integrations.monday import MondayClient
from integrations.calendar_api import CalendarClient
from integrations.box_api import BoxClient
from integrations.airtable import AirtableClient
from config import settings

logger = logging.getLogger(__name__)

# Directory for per-vendor instruction files (corrections / learnings)
VENDOR_INSTRUCTIONS_DIR = Path(__file__).parent.parent / "vendor-instructions"


class ContextBuilder:
    """Builds a complete VendorContext by querying all data sources."""

    def __init__(
        self,
        gmail: GmailClient,
        monday: MondayClient,
        calendar: CalendarClient,
        box: BoxClient,
        airtable: AirtableClient,
    ):
        self.gmail = gmail
        self.monday = monday
        self.calendar = calendar
        self.box = box
        self.airtable = airtable

    def build(self, vendor_name: str, vendor_slug: str) -> VendorContext:
        """
        Build full vendor context from all sources.

        Order matters - we gather from most reliable/fast to least:
        1. Monday.com (status, tasks, notes, contacts)
        2. Gmail - open emails (00.received, last 90 days = unsolved problems)
        3. Gmail - resolved emails (vendor label, no 00.received = already handled)
        4. Calendar (upcoming meetings)
        5. Box.com (documents)
        6. Airtable (contracts)
        7. Vendor instructions (per-vendor corrections/learnings)
        """
        ctx = VendorContext(vendor_name=vendor_name, vendor_slug=vendor_slug)

        # 1. Monday.com - try both buyer and affiliate boards
        self._load_monday_data(ctx)

        # 2 & 3. Gmail - split into open vs resolved
        self._load_email_history(ctx)

        # 4. Calendar
        self._load_calendar(ctx)

        # 5. Box.com
        self._load_box_docs(ctx)

        # 6. Airtable contracts
        self._load_contracts(ctx)

        # 7. Vendor-specific instructions
        self._load_vendor_instructions(ctx)

        # Compute derived fields
        self._compute_derived(ctx)

        logger.info(f"Built context for {vendor_name}: {ctx.summary()}")
        return ctx

    def _load_monday_data(self, ctx: VendorContext) -> None:
        """Load vendor data from Monday.com boards."""
        try:
            # Try buyers board first
            item = self.monday.get_vendor_item(
                settings.MONDAY_BUYERS_BOARD_ID, ctx.vendor_name
            )
            if item:
                ctx.vendor_type = "buyer"
                ctx.monday_item = item
                ctx.monday_status = item.column_values.get("status")
            else:
                # Try affiliates board
                item = self.monday.get_vendor_item(
                    settings.MONDAY_AFFILIATES_BOARD_ID, ctx.vendor_name
                )
                if item:
                    ctx.vendor_type = "affiliate"
                    ctx.monday_item = item
                    ctx.monday_status = item.column_values.get("status")

            # Tasks
            if settings.MONDAY_TASKS_BOARD_ID:
                ctx.monday_tasks = self.monday.get_vendor_tasks(
                    settings.MONDAY_TASKS_BOARD_ID, ctx.vendor_name
                )

            # Contacts
            if settings.MONDAY_CONTACTS_BOARD_ID:
                ctx.monday_contacts = self.monday.get_contacts_for_vendor(
                    settings.MONDAY_CONTACTS_BOARD_ID, ctx.vendor_name
                )

        except Exception as e:
            logger.warning(f"Failed to load Monday.com data for {ctx.vendor_name}: {e}")

    def _load_email_history(self, ctx: VendorContext) -> None:
        """
        Load email threads for this vendor, split by resolution status.

        - open_emails: label:00.received (last 90 days) — unsolved problems
        - resolved_emails: vendor label but NOT 00.received — already handled
        - recent_emails: kept for backward compat (all vendor threads)
        """
        try:
            ctx.open_emails = self.gmail.get_vendor_open_emails(
                ctx.vendor_slug, max_results=50
            )
            logger.info(
                f"  {ctx.vendor_name}: {len(ctx.open_emails)} open emails (00.received)"
            )
        except Exception as e:
            logger.warning(f"Failed to load open emails for {ctx.vendor_name}: {e}")

        try:
            ctx.resolved_emails = self.gmail.get_vendor_resolved_emails(
                ctx.vendor_slug, max_results=20
            )
            logger.info(
                f"  {ctx.vendor_name}: {len(ctx.resolved_emails)} resolved emails"
            )
        except Exception as e:
            logger.warning(f"Failed to load resolved emails for {ctx.vendor_name}: {e}")

        # Backward compat: merge into recent_emails
        ctx.recent_emails = ctx.open_emails + ctx.resolved_emails

    def _load_calendar(self, ctx: VendorContext) -> None:
        """Load upcoming meetings for this vendor."""
        try:
            ctx.upcoming_meetings = self.calendar.get_upcoming_meetings(ctx.vendor_name)
        except Exception as e:
            logger.warning(f"Failed to load calendar for {ctx.vendor_name}: {e}")

    def _load_box_docs(self, ctx: VendorContext) -> None:
        """Load Box.com documents for this vendor."""
        try:
            ctx.box_documents = self.box.search_vendor_files(ctx.vendor_name)
        except Exception as e:
            logger.warning(f"Failed to load Box docs for {ctx.vendor_name}: {e}")

    def _load_contracts(self, ctx: VendorContext) -> None:
        """Load Airtable contract data for this vendor."""
        try:
            ctx.contracts = self.airtable.get_vendor_contracts(ctx.vendor_name)
        except Exception as e:
            logger.warning(f"Failed to load contracts for {ctx.vendor_name}: {e}")

    def _load_vendor_instructions(self, ctx: VendorContext) -> None:
        """
        Load per-vendor instructions from the vendor-instructions directory.

        These are plain text files named by vendor slug (e.g., vendor-instructions/acme-corp.md).
        They contain corrections and learnings from past mistakes so the system doesn't
        repeat them. Users can edit these files directly to guide future processing.
        """
        instructions_file = VENDOR_INSTRUCTIONS_DIR / f"{ctx.vendor_slug}.md"
        if instructions_file.exists():
            try:
                ctx.vendor_instructions = instructions_file.read_text().strip()
                if ctx.vendor_instructions:
                    logger.info(
                        f"  Loaded vendor instructions for {ctx.vendor_name} "
                        f"({len(ctx.vendor_instructions)} chars)"
                    )
            except Exception as e:
                logger.warning(f"Failed to read instructions for {ctx.vendor_name}: {e}")

    def _compute_derived(self, ctx: VendorContext) -> None:
        """Compute derived fields from the gathered data."""
        # Open and overdue tasks
        now = datetime.now()
        for task in ctx.monday_tasks:
            if task.status and task.status.lower() not in ("done", "completed", "closed"):
                ctx.open_task_count += 1
                if task.due_date:
                    try:
                        due = datetime.fromisoformat(task.due_date)
                        if due < now:
                            ctx.overdue_task_count += 1
                    except ValueError:
                        pass

        # Last contact date from emails
        if ctx.recent_emails:
            for email in ctx.recent_emails:
                if email.date and (ctx.last_contact_date is None or email.date > ctx.last_contact_date):
                    ctx.last_contact_date = email.date

            if ctx.last_contact_date:
                ctx.days_since_contact = (now - ctx.last_contact_date).days
