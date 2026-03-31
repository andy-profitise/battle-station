"""
Proactive Scanner - scans all vendors for things that need attention.

This runs when the inbox is clear and looks for:
- Vendors we haven't contacted in too long
- Overdue tasks
- Approaching contract renewals
- Status changes that need follow-up
- Revenue anomalies
"""
from __future__ import annotations

import logging
from datetime import datetime, timedelta

from core.models import VendorContext, WorkItem, WorkItemSource, Priority
from integrations.monday import MondayClient
from config import settings

logger = logging.getLogger(__name__)


# Thresholds for proactive outreach
DAYS_NO_CONTACT_LIVE = 14       # Live vendor, no contact in 14 days
DAYS_NO_CONTACT_ONBOARDING = 7  # Onboarding vendor, no contact in 7 days
DAYS_CONTRACT_RENEWAL_WARNING = 30  # Warn 30 days before contract expires


class ProactiveScanner:
    """Scans vendors for proactive action opportunities."""

    def __init__(self, monday: MondayClient):
        self.monday = monday

    def scan_all_vendors(self, vendor_contexts: list[VendorContext]) -> list[WorkItem]:
        """
        Scan all vendors and return work items for anything that needs attention.
        """
        work_items = []

        for ctx in vendor_contexts:
            reasons = self._evaluate_vendor(ctx)
            if reasons:
                reason_text = "; ".join(reasons)
                work_items.append(WorkItem(
                    source=WorkItemSource.PROACTIVE_SCAN,
                    priority=Priority.LOW,
                    vendor_name=ctx.vendor_name,
                    vendor_slug=ctx.vendor_slug,
                    message_text=f"Proactive scan findings: {reason_text}",
                    message_date=datetime.now(),
                ))

        logger.info(f"Proactive scan found {len(work_items)} vendors needing attention")
        return work_items

    def _evaluate_vendor(self, ctx: VendorContext) -> list[str]:
        """Evaluate a single vendor for proactive opportunities. Returns list of reasons."""
        reasons = []
        now = datetime.now()

        # Check: no contact for too long based on status
        if ctx.days_since_contact is not None:
            status = (ctx.monday_status or "").lower()
            if status == "live" and ctx.days_since_contact > DAYS_NO_CONTACT_LIVE:
                reasons.append(
                    f"Live vendor with no contact in {ctx.days_since_contact} days"
                )
            elif status == "onboarding" and ctx.days_since_contact > DAYS_NO_CONTACT_ONBOARDING:
                reasons.append(
                    f"Onboarding vendor with no contact in {ctx.days_since_contact} days"
                )

        # Check: overdue tasks
        if ctx.overdue_task_count > 0:
            reasons.append(f"{ctx.overdue_task_count} overdue tasks")

        # Check: contract renewals approaching
        for contract in ctx.contracts:
            if contract.end_date:
                try:
                    end = datetime.fromisoformat(contract.end_date)
                    days_until = (end - now).days
                    if 0 < days_until <= DAYS_CONTRACT_RENEWAL_WARNING:
                        reasons.append(
                            f"Contract expires in {days_until} days ({contract.contract_type})"
                        )
                    elif days_until <= 0:
                        reasons.append(
                            f"Contract EXPIRED {abs(days_until)} days ago ({contract.contract_type})"
                        )
                except ValueError:
                    pass

        # Check: upcoming meeting with no recent prep
        if ctx.upcoming_meetings and ctx.days_since_contact and ctx.days_since_contact > 7:
            next_meeting = ctx.upcoming_meetings[0]
            if next_meeting.start:
                days_until_meeting = (next_meeting.start - now).days
                if 0 < days_until_meeting <= 3:
                    reasons.append(
                        f"Meeting in {days_until_meeting} days but no contact in "
                        f"{ctx.days_since_contact} days - prep needed"
                    )

        return reasons
