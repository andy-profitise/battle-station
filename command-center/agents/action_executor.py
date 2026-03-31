"""
Action Executor - carries out the actions from an ActionPlan.

Actions are executed based on their approval mode:
- AUTO: Execute immediately
- DRAFT: Create drafts/queue for human review
- FLAG: Log and notify, do not execute
"""
from __future__ import annotations

import logging
from dataclasses import dataclass, field

from core.models import Action, ActionPlan, ActionType, ApprovalMode
from integrations.gmail import GmailClient
from integrations.monday import MondayClient
from config import settings

logger = logging.getLogger(__name__)


@dataclass
class ExecutionResult:
    """Result of executing an action plan."""
    plan: ActionPlan
    executed: list[tuple[Action, str]] = field(default_factory=list)   # (action, result_msg)
    drafted: list[tuple[Action, str]] = field(default_factory=list)    # (action, draft_id/msg)
    flagged: list[tuple[Action, str]] = field(default_factory=list)    # (action, reason)
    errors: list[tuple[Action, str]] = field(default_factory=list)     # (action, error_msg)

    def summary(self) -> str:
        parts = []
        if self.executed:
            parts.append(f"Executed: {len(self.executed)} actions")
        if self.drafted:
            parts.append(f"Drafted: {len(self.drafted)} for review")
        if self.flagged:
            parts.append(f"Flagged: {len(self.flagged)} for human")
        if self.errors:
            parts.append(f"Errors: {len(self.errors)}")
        return " | ".join(parts) or "No actions taken"


class ActionExecutor:
    """Executes actions from ActionPlans."""

    def __init__(self, gmail: GmailClient, monday: MondayClient):
        self.gmail = gmail
        self.monday = monday

    def execute(self, plan: ActionPlan) -> ExecutionResult:
        """Execute all actions in an ActionPlan."""
        result = ExecutionResult(plan=plan)

        vendor = plan.vendor_context.vendor_name if plan.vendor_context else "Unknown"
        logger.info(f"Executing {len(plan.actions)} actions for {vendor}")

        for action in plan.actions:
            try:
                if action.approval_mode == ApprovalMode.FLAG:
                    result.flagged.append((action, action.description))
                    logger.info(f"FLAGGED: {action.description}")
                elif action.approval_mode == ApprovalMode.DRAFT:
                    draft_result = self._execute_as_draft(action)
                    result.drafted.append((action, draft_result))
                else:
                    exec_result = self._execute_action(action)
                    result.executed.append((action, exec_result))
            except Exception as e:
                logger.error(f"Error executing action {action.action_type}: {e}")
                result.errors.append((action, str(e)))

        logger.info(f"Execution complete for {vendor}: {result.summary()}")
        return result

    def _execute_action(self, action: Action) -> str:
        """Execute an auto-approved action."""
        match action.action_type:
            case ActionType.ARCHIVE_THREAD:
                if action.email_thread_id:
                    self.gmail.archive_thread(action.email_thread_id)
                    return f"Archived thread {action.email_thread_id}"

            case ActionType.LABEL_THREAD:
                if action.email_thread_id and action.label_name:
                    self.gmail.add_label(action.email_thread_id, action.label_name)
                    return f"Added label {action.label_name} to thread"

            case ActionType.CREATE_TASK:
                if action.task_name and settings.MONDAY_TASKS_BOARD_ID:
                    task_id = self.monday.create_task(
                        settings.MONDAY_TASKS_BOARD_ID,
                        "new_group",  # Default group - should be configurable
                        action.task_name,
                    )
                    return f"Created task {task_id}: {action.task_name}"

            case ActionType.UPDATE_NOTES:
                if action.notes_text:
                    # This would update Monday.com notes
                    logger.info(f"Would update notes: {action.notes_text[:100]}")
                    return "Notes update queued"

            case ActionType.SEND_EMAIL:
                # Even "auto" emails go through draft for safety
                return self._execute_as_draft(action)

            case _:
                return f"Action type {action.action_type} not implemented for auto-execute"

        return "No action taken (missing required fields)"

    def _execute_as_draft(self, action: Action) -> str:
        """Create a draft instead of executing directly."""
        match action.action_type:
            case ActionType.SEND_EMAIL | ActionType.DRAFT_EMAIL:
                if action.email_to and action.email_body:
                    subject = action.email_subject or "Re: Vendor Communication"
                    draft_id = self.gmail.create_draft(
                        to=action.email_to,
                        subject=subject,
                        body=action.email_body,
                        thread_id=action.email_thread_id,
                    )
                    return f"Created draft {draft_id}: {subject}"
                return "Draft skipped - missing email_to or email_body"

            case ActionType.CREATE_TASK:
                logger.info(f"DRAFT task: {action.task_name}")
                return f"Task queued for review: {action.task_name}"

            case _:
                logger.info(f"DRAFT action {action.action_type}: {action.description}")
                return f"Queued for review: {action.description}"
