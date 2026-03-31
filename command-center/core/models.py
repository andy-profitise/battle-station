"""
Core data models for the Vendor Command Center.
"""
from __future__ import annotations

import enum
import uuid
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any


# ---------------------------------------------------------------------------
# Enums
# ---------------------------------------------------------------------------

class Priority(enum.IntEnum):
    URGENT = 1      # Ad-hoc from Telegram/Teams
    HIGH = 2        # Inbox emails
    MEDIUM = 3      # 00.received emails
    LOW = 4         # Proactive vendor sweep


class WorkItemSource(str, enum.Enum):
    GMAIL_INBOX = "gmail_inbox"
    GMAIL_RECEIVED = "gmail_received"
    TELEGRAM = "telegram"
    TEAMS = "teams"
    PROACTIVE_SCAN = "proactive_scan"


class ActionType(str, enum.Enum):
    SEND_EMAIL = "send_email"
    DRAFT_EMAIL = "draft_email"
    CREATE_TASK = "create_task"
    UPDATE_TASK = "update_task"
    UPDATE_STATUS = "update_status"
    UPDATE_NOTES = "update_notes"
    ARCHIVE_THREAD = "archive_thread"
    LABEL_THREAD = "label_thread"
    SCHEDULE_FOLLOWUP = "schedule_followup"
    FLAG_FOR_HUMAN = "flag_for_human"
    SNOOZE = "snooze"


class ApprovalMode(str, enum.Enum):
    AUTO = "auto"           # Execute immediately
    DRAFT = "draft"         # Queue for human review
    FLAG = "flag"           # Flag and wait for explicit approval


# ---------------------------------------------------------------------------
# Work Items (things in the queue)
# ---------------------------------------------------------------------------

@dataclass
class WorkItem:
    """A unit of work to process - an email, a Telegram message, a proactive scan hit."""
    id: str = field(default_factory=lambda: str(uuid.uuid4())[:8])
    source: WorkItemSource = WorkItemSource.GMAIL_INBOX
    priority: Priority = Priority.MEDIUM
    vendor_name: str | None = None
    vendor_slug: str | None = None

    # Source-specific payload
    gmail_thread_id: str | None = None
    gmail_message_id: str | None = None
    message_text: str | None = None
    message_date: datetime | None = None
    raw_data: dict[str, Any] = field(default_factory=dict)

    created_at: datetime = field(default_factory=datetime.now)

    def __lt__(self, other: WorkItem) -> bool:
        """For priority queue ordering. Lower priority value = higher urgency."""
        if self.priority != other.priority:
            return self.priority < other.priority
        # Within same priority, oldest first
        if self.message_date and other.message_date:
            return self.message_date < other.message_date
        return self.created_at < other.created_at


# ---------------------------------------------------------------------------
# Vendor Context (everything we know before reading the trigger)
# ---------------------------------------------------------------------------

@dataclass
class MondayItem:
    id: str
    name: str
    status: str | None = None
    group: str | None = None
    column_values: dict[str, Any] = field(default_factory=dict)


@dataclass
class MondayTask:
    id: str
    name: str
    status: str | None = None
    due_date: str | None = None
    assignee: str | None = None
    last_updated: str | None = None


@dataclass
class EmailThread:
    thread_id: str
    subject: str
    snippet: str
    date: datetime | None = None
    from_addr: str | None = None
    labels: list[str] = field(default_factory=list)
    messages: list[dict[str, Any]] = field(default_factory=list)


@dataclass
class CalendarEvent:
    event_id: str
    summary: str
    start: datetime | None = None
    end: datetime | None = None
    attendees: list[str] = field(default_factory=list)
    link: str | None = None


@dataclass
class BoxDocument:
    file_id: str
    name: str
    modified_at: str | None = None
    size: int = 0
    download_url: str | None = None


@dataclass
class AirtableContract:
    record_id: str
    vendor_name: str
    contract_type: str | None = None
    start_date: str | None = None
    end_date: str | None = None
    status: str | None = None
    values: dict[str, Any] = field(default_factory=dict)


@dataclass
class VendorContext:
    """Full context for a vendor, gathered from all sources before reading the trigger."""
    vendor_name: str
    vendor_slug: str
    vendor_type: str | None = None       # buyer, affiliate, both

    # Monday.com
    monday_item: MondayItem | None = None
    monday_status: str | None = None
    monday_notes: str | None = None
    monday_tasks: list[MondayTask] = field(default_factory=list)
    monday_contacts: list[dict[str, Any]] = field(default_factory=list)
    monday_helpful_links: list[dict[str, Any]] = field(default_factory=list)

    # Gmail - split by resolution status
    recent_emails: list[EmailThread] = field(default_factory=list)          # Legacy: all recent
    open_emails: list[EmailThread] = field(default_factory=list)            # label:00.received (unsolved, last 90 days)
    resolved_emails: list[EmailThread] = field(default_factory=list)        # vendor label but NOT 00.received (handled)

    # Calendar
    upcoming_meetings: list[CalendarEvent] = field(default_factory=list)

    # Box.com
    box_documents: list[BoxDocument] = field(default_factory=list)

    # Airtable
    contracts: list[AirtableContract] = field(default_factory=list)

    # Vendor-specific instructions (corrections / learnings from past mistakes)
    vendor_instructions: str | None = None

    # Computed
    last_contact_date: datetime | None = None
    days_since_contact: int | None = None
    open_task_count: int = 0
    overdue_task_count: int = 0

    def summary(self) -> str:
        """One-paragraph summary of this vendor's current state."""
        parts = [f"Vendor: {self.vendor_name} ({self.vendor_type or 'unknown type'})"]
        if self.monday_status:
            parts.append(f"Status: {self.monday_status}")
        if self.monday_notes:
            # Truncate notes to first 200 chars
            notes_preview = self.monday_notes[:200] + ("..." if len(self.monday_notes) > 200 else "")
            parts.append(f"Notes: {notes_preview}")
        if self.monday_tasks:
            parts.append(f"Tasks: {len(self.monday_tasks)} ({self.overdue_task_count} overdue)")
        if self.upcoming_meetings:
            next_meeting = self.upcoming_meetings[0]
            parts.append(f"Next meeting: {next_meeting.summary} on {next_meeting.start}")
        if self.open_emails:
            parts.append(f"Open emails (00.received): {len(self.open_emails)}")
        if self.resolved_emails:
            parts.append(f"Resolved emails: {len(self.resolved_emails)}")
        if self.recent_emails:
            parts.append(f"Recent emails: {len(self.recent_emails)} threads")
        if self.contracts:
            parts.append(f"Contracts: {len(self.contracts)}")
        if self.days_since_contact is not None:
            parts.append(f"Days since last contact: {self.days_since_contact}")
        return " | ".join(parts)


# ---------------------------------------------------------------------------
# Action Plan (what we decide to do)
# ---------------------------------------------------------------------------

@dataclass
class Action:
    """A single action to take."""
    action_type: ActionType
    approval_mode: ApprovalMode = ApprovalMode.DRAFT
    description: str = ""

    # Action-specific data
    email_to: str | None = None
    email_subject: str | None = None
    email_body: str | None = None
    email_thread_id: str | None = None      # For replies

    task_name: str | None = None
    task_board_id: str | None = None
    task_assignee: str | None = None
    task_due_date: str | None = None

    status_value: str | None = None
    notes_text: str | None = None
    label_name: str | None = None

    snooze_until: datetime | None = None


@dataclass
class SituationalAssessment:
    """Claude's analysis of what's happening with this vendor right now."""
    summary: str = ""
    vendor_needs: str = ""
    our_position: str = ""
    urgency: str = ""           # low, medium, high, critical
    sentiment: str = ""         # positive, neutral, negative, adversarial
    key_facts: list[str] = field(default_factory=list)
    questions: list[str] = field(default_factory=list)          # Questions arising from the new email
    resolution_plan: list[str] = field(default_factory=list)    # Steps to solve the issues
    can_resolve_now: bool = False                               # True = do it now, False = defer and notify client
    defer_reason: str = ""                                      # Why we're deferring (if applicable)


@dataclass
class ActionPlan:
    """The complete decision output for a work item."""
    work_item: WorkItem | None = None
    vendor_context: VendorContext | None = None
    assessment: SituationalAssessment = field(default_factory=SituationalAssessment)
    actions: list[Action] = field(default_factory=list)
    reasoning: str = ""         # Why we chose these actions
    created_at: datetime = field(default_factory=datetime.now)
