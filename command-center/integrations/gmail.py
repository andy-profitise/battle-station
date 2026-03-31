"""
Gmail integration for the Vendor Command Center.

Handles searching, reading, labeling, archiving, and sending emails.
Uses the Gmail API via google-api-python-client.
"""
from __future__ import annotations

import base64
import logging
from datetime import datetime
from email.mime.text import MIMEText
from typing import Any

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

from config import settings
from core.models import EmailThread

logger = logging.getLogger(__name__)

SCOPES = [
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/gmail.send",
]


class GmailClient:
    """Gmail API client for the Command Center."""

    def __init__(self):
        self._service = None

    def _get_service(self):
        if self._service:
            return self._service

        creds = None
        import os
        if os.path.exists(settings.GMAIL_TOKEN_PATH):
            creds = Credentials.from_authorized_user_file(settings.GMAIL_TOKEN_PATH, SCOPES)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    settings.GMAIL_CREDENTIALS_PATH, SCOPES
                )
                creds = flow.run_local_server(port=0)

            with open(settings.GMAIL_TOKEN_PATH, "w") as token:
                token.write(creds.to_json())

        self._service = build("gmail", "v1", credentials=creds)
        return self._service

    def search_threads(self, query: str, max_results: int = 50) -> list[dict[str, Any]]:
        """
        Search Gmail threads matching a query.
        Returns list of thread metadata dicts sorted by date (oldest first).
        """
        service = self._get_service()
        results = service.users().threads().list(
            userId="me", q=query, maxResults=max_results
        ).execute()

        threads = results.get("threads", [])
        thread_data = []

        for t in threads:
            thread = service.users().threads().get(
                userId="me", id=t["id"], format="metadata",
                metadataHeaders=["Subject", "From", "Date"]
            ).execute()

            messages = thread.get("messages", [])
            if not messages:
                continue

            # Get metadata from the first message
            first_msg = messages[0]
            headers = {
                h["name"]: h["value"]
                for h in first_msg.get("payload", {}).get("headers", [])
            }

            # Get labels from all messages
            all_labels = set()
            for msg in messages:
                all_labels.update(msg.get("labelIds", []))

            thread_data.append({
                "thread_id": thread["id"],
                "subject": headers.get("Subject", "(no subject)"),
                "from": headers.get("From", ""),
                "date": headers.get("Date", ""),
                "snippet": thread.get("snippet", ""),
                "labels": list(all_labels),
                "message_count": len(messages),
            })

        # Sort by date, oldest first
        thread_data.sort(key=lambda t: t.get("date", ""))
        return thread_data

    def get_live_vendor_emails(self) -> list[dict[str, Any]]:
        """
        Get all "live" vendor emails using the standard Battle Station query:
        label:inbox OR label:00.received AND -is:snoozed
        """
        return self.search_threads(settings.GMAIL_INBOX_QUERY)

    def get_thread_full(self, thread_id: str) -> EmailThread:
        """Get full thread content including message bodies."""
        service = self._get_service()
        thread = service.users().threads().get(
            userId="me", id=thread_id, format="full"
        ).execute()

        messages = thread.get("messages", [])
        parsed_messages = []
        all_labels = set()

        for msg in messages:
            headers = {
                h["name"]: h["value"]
                for h in msg.get("payload", {}).get("headers", [])
            }
            body = self._extract_body(msg.get("payload", {}))
            all_labels.update(msg.get("labelIds", []))

            parsed_messages.append({
                "message_id": msg["id"],
                "from": headers.get("From", ""),
                "to": headers.get("To", ""),
                "date": headers.get("Date", ""),
                "subject": headers.get("Subject", ""),
                "body": body,
            })

        first_msg = parsed_messages[0] if parsed_messages else {}
        return EmailThread(
            thread_id=thread_id,
            subject=first_msg.get("subject", ""),
            snippet=thread.get("snippet", ""),
            from_addr=first_msg.get("from", ""),
            labels=list(all_labels),
            messages=parsed_messages,
        )

    def get_vendor_threads(self, vendor_slug: str, max_results: int = 10) -> list[EmailThread]:
        """Get recent email threads for a specific vendor by their label."""
        label_name = f"zzzvendors-{vendor_slug}"
        threads = self.search_threads(f"label:{label_name}", max_results=max_results)
        return [
            EmailThread(
                thread_id=t["thread_id"],
                subject=t["subject"],
                snippet=t["snippet"],
                from_addr=t.get("from", ""),
                labels=t.get("labels", []),
            )
            for t in threads
        ]

    def get_vendor_open_emails(self, vendor_slug: str, max_results: int = 50) -> list[EmailThread]:
        """
        Get open/unsolved emails for a vendor: label:00.received AND label:zzzvendors-{slug}
        within the last 90 days. These are problems that still need solving.
        """
        label_name = f"zzzvendors-{vendor_slug}"
        query = f"label:{label_name} label:00.received newer_than:90d"
        threads = self.search_threads(query, max_results=max_results)
        return [
            EmailThread(
                thread_id=t["thread_id"],
                subject=t["subject"],
                snippet=t["snippet"],
                from_addr=t.get("from", ""),
                labels=t.get("labels", []),
            )
            for t in threads
        ]

    def get_vendor_resolved_emails(self, vendor_slug: str, max_results: int = 20) -> list[EmailThread]:
        """
        Get resolved/handled emails for a vendor: label:zzzvendors-{slug} but NOT label:00.received.
        These have already been dealt with and provide context on how we've handled past issues.
        """
        label_name = f"zzzvendors-{vendor_slug}"
        query = f"label:{label_name} -label:00.received newer_than:90d"
        threads = self.search_threads(query, max_results=max_results)
        return [
            EmailThread(
                thread_id=t["thread_id"],
                subject=t["subject"],
                snippet=t["snippet"],
                from_addr=t.get("from", ""),
                labels=t.get("labels", []),
            )
            for t in threads
        ]

    def _extract_body(self, payload: dict) -> str:
        """Extract text body from a message payload."""
        if payload.get("mimeType") == "text/plain":
            data = payload.get("body", {}).get("data", "")
            if data:
                return base64.urlsafe_b64decode(data).decode("utf-8", errors="replace")

        # Check parts
        for part in payload.get("parts", []):
            if part.get("mimeType") == "text/plain":
                data = part.get("body", {}).get("data", "")
                if data:
                    return base64.urlsafe_b64decode(data).decode("utf-8", errors="replace")
            # Recurse into multipart
            if part.get("parts"):
                body = self._extract_body(part)
                if body:
                    return body

        return ""

    def get_label_names(self, label_ids: list[str]) -> list[str]:
        """Convert Gmail label IDs to human-readable names."""
        service = self._get_service()
        labels_response = service.users().labels().list(userId="me").execute()
        id_to_name = {l["id"]: l["name"] for l in labels_response.get("labels", [])}
        return [id_to_name.get(lid, lid) for lid in label_ids]

    def create_draft(self, to: str, subject: str, body: str,
                     thread_id: str | None = None) -> str:
        """Create a draft email. Returns the draft ID."""
        service = self._get_service()
        message = MIMEText(body)
        message["to"] = to
        message["subject"] = subject

        raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
        draft_body: dict[str, Any] = {"message": {"raw": raw}}
        if thread_id:
            draft_body["message"]["threadId"] = thread_id

        draft = service.users().drafts().create(
            userId="me", body=draft_body
        ).execute()

        logger.info(f"Created draft {draft['id']} to {to}: {subject}")
        return draft["id"]

    def send_email(self, to: str, subject: str, body: str,
                   thread_id: str | None = None) -> str:
        """Send an email. Returns the message ID."""
        service = self._get_service()
        message = MIMEText(body)
        message["to"] = to
        message["subject"] = subject

        raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
        send_body: dict[str, Any] = {"raw": raw}
        if thread_id:
            send_body["threadId"] = thread_id

        sent = service.users().messages().send(
            userId="me", body=send_body
        ).execute()

        logger.info(f"Sent email {sent['id']} to {to}: {subject}")
        return sent["id"]

    def archive_thread(self, thread_id: str) -> None:
        """Archive a thread (remove INBOX label)."""
        service = self._get_service()
        service.users().threads().modify(
            userId="me", id=thread_id,
            body={"removeLabelIds": ["INBOX"]}
        ).execute()
        logger.info(f"Archived thread {thread_id}")

    def add_label(self, thread_id: str, label_id: str) -> None:
        """Add a label to a thread."""
        service = self._get_service()
        service.users().threads().modify(
            userId="me", id=thread_id,
            body={"addLabelIds": [label_id]}
        ).execute()

    def remove_label(self, thread_id: str, label_id: str) -> None:
        """Remove a label from a thread."""
        service = self._get_service()
        service.users().threads().modify(
            userId="me", id=thread_id,
            body={"removeLabelIds": [label_id]}
        ).execute()
