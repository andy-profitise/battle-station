"""
Box.com integration for the Vendor Command Center.
"""
from __future__ import annotations

import logging
from typing import Any

import requests

from config import settings
from core.models import BoxDocument

logger = logging.getLogger(__name__)


class BoxClient:
    """Box.com API client."""

    BASE_URL = "https://api.box.com/2.0"

    def __init__(self):
        self._headers = {
            "Authorization": f"Bearer {settings.BOX_ACCESS_TOKEN}",
            "Content-Type": "application/json",
        }

    def _get(self, endpoint: str, params: dict | None = None) -> dict[str, Any]:
        response = requests.get(
            f"{self.BASE_URL}/{endpoint}",
            headers=self._headers,
            params=params,
        )
        response.raise_for_status()
        return response.json()

    def search_vendor_files(self, vendor_name: str, limit: int = 20) -> list[BoxDocument]:
        """Search Box for files related to a vendor."""
        data = self._get("search", {
            "query": vendor_name,
            "type": "file",
            "limit": limit,
        })

        return [
            BoxDocument(
                file_id=entry["id"],
                name=entry.get("name", ""),
                modified_at=entry.get("modified_at"),
                size=entry.get("size", 0),
            )
            for entry in data.get("entries", [])
        ]

    def get_folder_files(self, folder_id: str, limit: int = 100) -> list[BoxDocument]:
        """List files in a specific Box folder."""
        data = self._get(f"folders/{folder_id}/items", {"limit": limit})

        return [
            BoxDocument(
                file_id=entry["id"],
                name=entry.get("name", ""),
                modified_at=entry.get("modified_at"),
                size=entry.get("size", 0),
            )
            for entry in data.get("entries", [])
            if entry.get("type") == "file"
        ]
