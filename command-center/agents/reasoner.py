"""
Reasoner - Claude-powered decision engine.

Given a vendor's full context and a trigger (email/message), determines:
1. What is happening / what does the vendor need?
2. What is our current position / relationship?
3. What actions should we take?
4. What is the priority and urgency?
"""
from __future__ import annotations

import json
import logging
from typing import Any

import anthropic

from config import settings
from core.models import (
    Action,
    ActionPlan,
    ActionType,
    ApprovalMode,
    EmailThread,
    SituationalAssessment,
    VendorContext,
    WorkItem,
)

logger = logging.getLogger(__name__)


SYSTEM_PROMPT = """You are A(I)DEN, an AI vendor relationship manager for Profitise.
You help manage vendor relationships by analyzing emails and vendor data, then
recommending specific actions.

## How Email Status Works

- **Open emails (label:00.received)**: These are UNSOLVED problems. They landed in
  the received queue and have not been handled yet. Each one represents something
  that needs attention — a question, a request, a complaint, etc.
- **Resolved emails (no 00.received label)**: These have already been dealt with.
  They show how we've handled past issues with this vendor and give you patterns to
  follow.

## Your Job

Given a vendor's full context (open problems, resolved history, tasks, meetings,
documents, contracts) and a trigger (new email or message):

1. **Present our stance**: What is our current position with this vendor? What do
   we want from this relationship? What leverage or obligations do we have?

2. **Identify questions**: What questions arise from the new email/trigger?
   What information is missing? What do we need to clarify?

3. **Analyze open problems**: Look at ALL open (00.received) emails. How do they
   relate to each other? Is there a pattern? What's the priority order?

4. **Create a resolution plan**: For each issue, how do we solve it? Be specific.
   Reference supporting data from Monday.com tasks, calendar meetings, Box docs,
   Airtable contracts, and Google Drive files.

5. **Decide: do it now or defer**: Can we resolve this immediately (draft the email,
   create the task, update the status)? Or do we need to tell the client what we're
   planning and then process it later?

6. **Draft actions**: Write the emails, create the tasks, update the notes.

## Vendor-Specific Instructions

If vendor instructions are provided, they represent corrections and learnings from
past mistakes. ALWAYS follow these instructions — they override your default behavior.
They exist because the system got something wrong before, and the human corrected it.

You should be professional, proactive, and thorough. When drafting emails, match the
tone of the existing relationship. Be concise but complete.

IMPORTANT: You must respond with valid JSON matching the schema provided."""


def _build_context_prompt(ctx: VendorContext, trigger_email: EmailThread | None,
                          work_item: WorkItem) -> str:
    """Build the full prompt with vendor context and trigger."""
    parts = []

    # Vendor overview
    parts.append(f"## Vendor: {ctx.vendor_name}")
    parts.append(f"Type: {ctx.vendor_type or 'Unknown'}")
    parts.append(f"Status: {ctx.monday_status or 'Unknown'}")

    # Vendor-specific instructions (corrections / learnings)
    if ctx.vendor_instructions:
        parts.append(f"\n### VENDOR-SPECIFIC INSTRUCTIONS (MUST FOLLOW)")
        parts.append(ctx.vendor_instructions)

    if ctx.monday_notes:
        parts.append(f"\n### Internal Notes\n{ctx.monday_notes}")

    # Tasks
    if ctx.monday_tasks:
        parts.append("\n### Open Tasks (Monday.com)")
        for task in ctx.monday_tasks:
            status = task.status or "no status"
            due = f" (due: {task.due_date})" if task.due_date else ""
            parts.append(f"- [{status}] {task.name}{due}")

    # Contacts
    if ctx.monday_contacts:
        parts.append("\n### Contacts (Monday.com)")
        for contact in ctx.monday_contacts:
            parts.append(f"- {contact.get('name', 'Unknown')}: {contact.get('email', '')}")

    # Upcoming meetings
    if ctx.upcoming_meetings:
        parts.append("\n### Upcoming Meetings (Google Calendar)")
        for meeting in ctx.upcoming_meetings:
            parts.append(f"- {meeting.summary} on {meeting.start}")
            if meeting.attendees:
                parts.append(f"  Attendees: {', '.join(meeting.attendees)}")

    # Open emails (00.received) — UNSOLVED problems
    if ctx.open_emails:
        parts.append(f"\n### OPEN EMAILS — {len(ctx.open_emails)} unsolved problems (label:00.received, last 90 days)")
        parts.append("Each of these is an unresolved issue that needs attention:")
        for i, email in enumerate(ctx.open_emails, 1):
            parts.append(f"\n**Open #{i}**: {email.subject}")
            parts.append(f"  From: {email.from_addr or 'unknown'} | Date: {email.date or 'unknown'}")
            if email.snippet:
                parts.append(f"  Preview: {email.snippet[:200]}")
    else:
        parts.append("\n### OPEN EMAILS: None (no 00.received emails in last 90 days)")

    # Resolved emails — already handled, for pattern reference
    if ctx.resolved_emails:
        parts.append(f"\n### RESOLVED EMAILS — {len(ctx.resolved_emails)} handled threads (last 90 days)")
        parts.append("These show how we've dealt with past issues:")
        for email in ctx.resolved_emails[:10]:
            parts.append(f"- {email.subject} (from: {email.from_addr or 'unknown'})")
            if email.snippet:
                parts.append(f"  {email.snippet[:150]}")

    # Box.com documents
    if ctx.box_documents:
        parts.append(f"\n### Documents (Box.com)")
        for doc in ctx.box_documents:
            parts.append(f"- {doc.name} (modified: {doc.modified_at or 'unknown'})")

    # Contracts
    if ctx.contracts:
        parts.append("\n### Contracts (Airtable)")
        for c in ctx.contracts:
            parts.append(f"- {c.contract_type}: {c.status} ({c.start_date} to {c.end_date})")

    # Derived metrics
    if ctx.days_since_contact is not None:
        parts.append(f"\nDays since last contact: {ctx.days_since_contact}")
    parts.append(f"Open tasks: {ctx.open_task_count} ({ctx.overdue_task_count} overdue)")

    # Now the trigger
    parts.append("\n---\n## INCOMING TRIGGER")

    if trigger_email:
        parts.append(f"Subject: {trigger_email.subject}")
        parts.append(f"From: {trigger_email.from_addr}")
        for msg in trigger_email.messages:
            parts.append(f"\n### Message from {msg.get('from', 'unknown')} ({msg.get('date', '')}):")
            body = msg.get('body', '')
            # Truncate very long emails
            if len(body) > 3000:
                body = body[:3000] + "\n... [truncated]"
            parts.append(body)
    elif work_item.message_text:
        parts.append(f"Ad-hoc message: {work_item.message_text}")
        parts.append(f"Source: {work_item.source.value}")

    return "\n".join(parts)


ACTION_SCHEMA = {
    "type": "object",
    "properties": {
        "assessment": {
            "type": "object",
            "properties": {
                "summary": {"type": "string", "description": "2-3 sentence summary of what's happening"},
                "vendor_needs": {"type": "string", "description": "What the vendor wants/needs from us"},
                "our_position": {
                    "type": "string",
                    "description": (
                        "Our current stance/position. What do we want from this relationship? "
                        "What leverage or obligations do we have? Be specific."
                    ),
                },
                "urgency": {"type": "string", "enum": ["low", "medium", "high", "critical"]},
                "sentiment": {"type": "string", "enum": ["positive", "neutral", "negative", "adversarial"]},
                "key_facts": {"type": "array", "items": {"type": "string"}},
                "questions": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": (
                        "Questions arising from the new email/trigger. What information is missing? "
                        "What do we need to clarify before we can fully resolve this?"
                    ),
                },
                "resolution_plan": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": (
                        "Step-by-step plan to solve the issues. Reference supporting data "
                        "(tasks, meetings, docs, contracts) where applicable."
                    ),
                },
                "can_resolve_now": {
                    "type": "boolean",
                    "description": (
                        "True if we can resolve this immediately (send email, create task, etc). "
                        "False if we need to inform the client of our plan and process later."
                    ),
                },
                "defer_reason": {
                    "type": "string",
                    "description": "If can_resolve_now is false, explain why we need to defer.",
                },
            },
            "required": [
                "summary", "vendor_needs", "our_position", "urgency",
                "sentiment", "questions", "resolution_plan", "can_resolve_now",
            ],
        },
        "actions": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "action_type": {
                        "type": "string",
                        "enum": [t.value for t in ActionType],
                    },
                    "description": {"type": "string"},
                    "email_to": {"type": "string"},
                    "email_subject": {"type": "string"},
                    "email_body": {"type": "string"},
                    "email_thread_id": {"type": "string"},
                    "task_name": {"type": "string"},
                    "notes_text": {"type": "string"},
                    "label_name": {"type": "string"},
                },
                "required": ["action_type", "description"],
            },
        },
        "reasoning": {"type": "string", "description": "Brief explanation of why these actions"},
    },
    "required": ["assessment", "actions", "reasoning"],
}


class Reasoner:
    """Claude-powered vendor situation analyzer and action planner."""

    def __init__(self):
        self.client = anthropic.Anthropic(api_key=settings.ANTHROPIC_API_KEY)

    def analyze_and_plan(
        self,
        work_item: WorkItem,
        vendor_context: VendorContext,
        trigger_email: EmailThread | None = None,
    ) -> ActionPlan:
        """
        Analyze the vendor situation and create an action plan.

        Steps:
        1. Build the full context prompt
        2. Send to Claude for analysis
        3. Parse the response into an ActionPlan
        """
        context_prompt = _build_context_prompt(vendor_context, trigger_email, work_item)

        response = self.client.messages.create(
            model=settings.CLAUDE_MODEL,
            max_tokens=4096,
            system=SYSTEM_PROMPT,
            messages=[
                {
                    "role": "user",
                    "content": (
                        f"{context_prompt}\n\n---\n\n"
                        "Analyze this situation following these steps:\n\n"
                        "1. **Our stance**: What is our position with this vendor? What do we want?\n"
                        "2. **Questions**: What questions arise from the new email? What's unclear?\n"
                        "3. **Open problems**: Look at ALL open (00.received) emails. How do they relate?\n"
                        "4. **Resolution plan**: For each issue, what specific steps solve it?\n"
                        "   Reference Monday.com tasks, calendar meetings, Box docs, contracts.\n"
                        "5. **Do now or defer**: Can we act immediately, or do we need to tell the\n"
                        "   client our plan and process later?\n"
                        "6. **Actions**: Draft emails, create tasks, update notes as needed.\n\n"
                        "Respond with JSON matching this schema:\n"
                        f"```json\n{json.dumps(ACTION_SCHEMA, indent=2)}\n```\n\n"
                        "If an email reply is needed, draft the full email body.\n"
                        "If tasks need to be created, provide clear task names.\n"
                        "If we're deferring, draft the 'here's what we're going to do' email.\n"
                        "Be specific and actionable."
                    ),
                }
            ],
        )

        # Parse Claude's response
        response_text = response.content[0].text
        return self._parse_response(response_text, work_item, vendor_context)

    def _parse_response(self, response_text: str, work_item: WorkItem,
                        vendor_context: VendorContext) -> ActionPlan:
        """Parse Claude's JSON response into an ActionPlan."""
        # Extract JSON from response (may be wrapped in markdown code blocks)
        json_text = response_text
        if "```json" in json_text:
            json_text = json_text.split("```json")[1].split("```")[0]
        elif "```" in json_text:
            json_text = json_text.split("```")[1].split("```")[0]

        try:
            data = json.loads(json_text.strip())
        except json.JSONDecodeError:
            logger.error(f"Failed to parse Claude response as JSON: {response_text[:500]}")
            return ActionPlan(
                work_item=work_item,
                vendor_context=vendor_context,
                assessment=SituationalAssessment(
                    summary="Failed to parse AI response",
                    urgency="medium",
                ),
                actions=[
                    Action(
                        action_type=ActionType.FLAG_FOR_HUMAN,
                        approval_mode=ApprovalMode.FLAG,
                        description="AI response parsing failed - needs human review",
                    )
                ],
                reasoning=f"Raw response: {response_text[:500]}",
            )

        # Build assessment
        assessment_data = data.get("assessment", {})
        assessment = SituationalAssessment(
            summary=assessment_data.get("summary", ""),
            vendor_needs=assessment_data.get("vendor_needs", ""),
            our_position=assessment_data.get("our_position", ""),
            urgency=assessment_data.get("urgency", "medium"),
            sentiment=assessment_data.get("sentiment", "neutral"),
            key_facts=assessment_data.get("key_facts", []),
            questions=assessment_data.get("questions", []),
            resolution_plan=assessment_data.get("resolution_plan", []),
            can_resolve_now=assessment_data.get("can_resolve_now", False),
            defer_reason=assessment_data.get("defer_reason", ""),
        )

        # Build actions
        actions = []
        for action_data in data.get("actions", []):
            action_type_str = action_data.get("action_type", "flag_for_human")
            try:
                action_type = ActionType(action_type_str)
            except ValueError:
                action_type = ActionType.FLAG_FOR_HUMAN

            # Determine approval mode based on action type
            if action_type in (ActionType.SEND_EMAIL, ActionType.DRAFT_EMAIL):
                approval = ApprovalMode.DRAFT
            elif action_type in (ActionType.FLAG_FOR_HUMAN,):
                approval = ApprovalMode.FLAG
            else:
                approval = ApprovalMode.AUTO

            actions.append(Action(
                action_type=action_type,
                approval_mode=approval,
                description=action_data.get("description", ""),
                email_to=action_data.get("email_to"),
                email_subject=action_data.get("email_subject"),
                email_body=action_data.get("email_body"),
                email_thread_id=action_data.get("email_thread_id"),
                task_name=action_data.get("task_name"),
                notes_text=action_data.get("notes_text"),
                label_name=action_data.get("label_name"),
            ))

        return ActionPlan(
            work_item=work_item,
            vendor_context=vendor_context,
            assessment=assessment,
            actions=actions,
            reasoning=data.get("reasoning", ""),
        )
