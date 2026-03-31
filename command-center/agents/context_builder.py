"""
Context Builder - gathers everything we know about a vendor BEFORE reading the trigger.

This is the "look before you read" step. We want Claude to understand the full
vendor picture so it can properly contextualize whatever email/message triggered
this work item.
"""
from __future__ import annotations

import logging
from datetime import datetime

from core.models import VendorContext
from integrations.gmail import GmailClient
from integrations.monday import MondayClient
from integrations.calendar_api import CalendarClient
from integrations.box_api import BoxClient
from integrations.airtable import AirtableClient
from config import settings

logger = logging.getLogger(__name__)


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
        2. Gmail (recent threads, excluding the trigger)
        3. Calendar (upcoming meetings)
        4. Box.com (documents)
        5. Airtable (contracts)
        """
        ctx = VendorContext(vendor_name=vendor_name, vendor_slug=vendor_slug)

        # 1. Monday.com - try both buyer and affiliate boards
        self._load_monday_data(ctx)

        # 2. Gmail - recent threads for this vendor
        self._load_email_history(ctx)

        # 3. Calendar
        self._load_calendar(ctx)

        # 4. Box.com
        self._load_box_docs(ctx)

        # 5. Airtable contracts
        self._load_contracts(ctx)

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
        """Load recent email threads for this vendor."""
        try:
            ctx.recent_emails = self.gmail.get_vendor_threads(ctx.vendor_slug, max_results=10)
        except Exception as e:
            logger.warning(f"Failed to load email history for {ctx.vendor_name}: {e}")

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
