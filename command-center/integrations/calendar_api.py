"""
Google Calendar integration for the Vendor Command Center.
"""
from __future__ import annotations

import logging
from datetime import datetime, timedelta

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

from config import settings
from core.models import CalendarEvent

logger = logging.getLogger(__name__)


class CalendarClient:
    """Google Calendar API client."""

    def __init__(self):
        self._service = None

    def _get_service(self):
        if self._service:
            return self._service

        import os
        creds = None
        if os.path.exists(settings.GMAIL_TOKEN_PATH):
            creds = Credentials.from_authorized_user_file(settings.GMAIL_TOKEN_PATH)

        if not creds or not creds.valid:
            raise RuntimeError("Calendar credentials not available. Run Gmail auth first.")

        self._service = build("calendar", "v3", credentials=creds)
        return self._service

    def get_upcoming_meetings(self, vendor_name: str,
                              days_ahead: int = 30) -> list[CalendarEvent]:
        """Find upcoming calendar events that mention a vendor."""
        service = self._get_service()
        now = datetime.utcnow().isoformat() + "Z"
        future = (datetime.utcnow() + timedelta(days=days_ahead)).isoformat() + "Z"

        events_result = service.events().list(
            calendarId="primary",
            timeMin=now,
            timeMax=future,
            maxResults=50,
            singleEvents=True,
            orderBy="startTime",
            q=vendor_name,
        ).execute()

        events = events_result.get("items", [])
        results = []

        for event in events:
            start = event.get("start", {}).get("dateTime", event.get("start", {}).get("date"))
            end = event.get("end", {}).get("dateTime", event.get("end", {}).get("date"))
            attendees = [a.get("email", "") for a in event.get("attendees", [])]

            results.append(CalendarEvent(
                event_id=event["id"],
                summary=event.get("summary", ""),
                start=datetime.fromisoformat(start) if start else None,
                end=datetime.fromisoformat(end) if end else None,
                attendees=attendees,
                link=event.get("htmlLink"),
            ))

        return results
