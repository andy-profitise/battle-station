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

Your job is to:
1. Assess the current situation with this vendor
2. Understand what the incoming email/message is about in context
3. Recommend specific, actionable next steps
4. Draft any emails that need to be sent

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

    if ctx.monday_notes:
        parts.append(f"\n### Internal Notes\n{ctx.monday_notes}")

    # Tasks
    if ctx.monday_tasks:
        parts.append("\n### Open Tasks")
        for task in ctx.monday_tasks:
            status = task.status or "no status"
            due = f" (due: {task.due_date})" if task.due_date else ""
            parts.append(f"- [{status}] {task.name}{due}")

    # Contacts
    if ctx.monday_contacts:
        parts.append("\n### Contacts")
        for contact in ctx.monday_contacts:
            parts.append(f"- {contact.get('name', 'Unknown')}: {contact.get('email', '')}")

    # Upcoming meetings
    if ctx.upcoming_meetings:
        parts.append("\n### Upcoming Meetings")
        for meeting in ctx.upcoming_meetings:
            parts.append(f"- {meeting.summary} on {meeting.start}")

    # Recent email history (summaries only, not the trigger)
    if ctx.recent_emails:
        parts.append("\n### Recent Email History")
        for email in ctx.recent_emails[:5]:
            parts.append(f"- [{email.date}] {email.subject}: {email.snippet[:100]}")

    # Contracts
    if ctx.contracts:
        parts.append("\n### Contracts")
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
                "our_position": {"type": "string", "description": "Our current stance/position"},
                "urgency": {"type": "string", "enum": ["low", "medium", "high", "critical"]},
                "sentiment": {"type": "string", "enum": ["positive", "neutral", "negative", "adversarial"]},
                "key_facts": {"type": "array", "items": {"type": "string"}},
            },
            "required": ["summary", "vendor_needs", "our_position", "urgency", "sentiment"],
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
                        "Analyze this situation and provide your assessment and recommended actions.\n"
                        "Respond with JSON matching this schema:\n"
                        f"```json\n{json.dumps(ACTION_SCHEMA, indent=2)}\n```\n\n"
                        "If an email reply is needed, draft the full email body.\n"
                        "If tasks need to be created, provide clear task names.\n"
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
