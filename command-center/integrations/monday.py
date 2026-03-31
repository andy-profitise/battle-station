"""
Monday.com integration for the Vendor Command Center.

Handles reading vendor status, tasks, notes, contacts, and helpful links.
Also supports creating/updating tasks and updating status.
"""
from __future__ import annotations

import logging
from typing import Any

import requests

from config import settings
from core.models import MondayItem, MondayTask

logger = logging.getLogger(__name__)


class MondayClient:
    """Monday.com GraphQL API client."""

    def __init__(self):
        self._url = settings.MONDAY_API_URL
        self._headers = {
            "Authorization": settings.MONDAY_API_TOKEN,
            "Content-Type": "application/json",
            "API-Version": "2024-10",
        }

    def _query(self, query: str, variables: dict | None = None) -> dict[str, Any]:
        """Execute a GraphQL query against Monday.com."""
        payload: dict[str, Any] = {"query": query}
        if variables:
            payload["variables"] = variables

        response = requests.post(self._url, json=payload, headers=self._headers)
        response.raise_for_status()
        data = response.json()

        if "errors" in data:
            logger.error(f"Monday.com API errors: {data['errors']}")
            raise RuntimeError(f"Monday.com API error: {data['errors'][0].get('message', 'Unknown error')}")

        return data.get("data", {})

    def get_all_vendors(self, board_id: str, limit: int = 500) -> list[MondayItem]:
        """Get all items from a board (vendors)."""
        query = """
        query($boardId: [ID!]!, $limit: Int!) {
            boards(ids: $boardId) {
                items_page(limit: $limit) {
                    items {
                        id
                        name
                        group { title }
                        column_values {
                            id
                            text
                            value
                        }
                    }
                }
            }
        }
        """
        data = self._query(query, {"boardId": [board_id], "limit": limit})
        boards = data.get("boards", [])
        if not boards:
            return []

        items = boards[0].get("items_page", {}).get("items", [])
        return [
            MondayItem(
                id=item["id"],
                name=item["name"],
                group=item.get("group", {}).get("title"),
                column_values={
                    cv["id"]: cv["text"]
                    for cv in item.get("column_values", [])
                },
            )
            for item in items
        ]

    def get_vendor_item(self, board_id: str, vendor_name: str) -> MondayItem | None:
        """Find a specific vendor by name on a board."""
        query = """
        query($boardId: [ID!]!, $name: CompareValue!) {
            boards(ids: $boardId) {
                items_page(query_params: {rules: [{column_id: "name", compare_value: $name}]}) {
                    items {
                        id
                        name
                        group { title }
                        column_values {
                            id
                            text
                            value
                        }
                    }
                }
            }
        }
        """
        data = self._query(query, {"boardId": [board_id], "name": vendor_name})
        boards = data.get("boards", [])
        if not boards:
            return None

        items = boards[0].get("items_page", {}).get("items", [])
        if not items:
            return None

        item = items[0]
        return MondayItem(
            id=item["id"],
            name=item["name"],
            group=item.get("group", {}).get("title"),
            column_values={
                cv["id"]: cv["text"]
                for cv in item.get("column_values", [])
            },
        )

    def get_vendor_tasks(self, board_id: str, vendor_name: str) -> list[MondayTask]:
        """Get tasks for a vendor from the tasks board."""
        query = """
        query($boardId: [ID!]!) {
            boards(ids: $boardId) {
                items_page(limit: 100) {
                    items {
                        id
                        name
                        column_values {
                            id
                            text
                            value
                        }
                    }
                }
            }
        }
        """
        data = self._query(query, {"boardId": [board_id]})
        boards = data.get("boards", [])
        if not boards:
            return []

        items = boards[0].get("items_page", {}).get("items", [])
        tasks = []
        vendor_lower = vendor_name.lower()

        for item in items:
            cols = {cv["id"]: cv["text"] for cv in item.get("column_values", [])}
            # Match tasks that reference this vendor (by name in task or vendor column)
            item_text = (item["name"] + " " + " ".join(cols.values())).lower()
            if vendor_lower in item_text:
                tasks.append(MondayTask(
                    id=item["id"],
                    name=item["name"],
                    status=cols.get("status"),
                    due_date=cols.get("date") or cols.get("due_date"),
                    assignee=cols.get("person"),
                    last_updated=cols.get("last_updated"),
                ))

        return tasks

    def get_vendor_notes(self, item_id: str, notes_column_id: str) -> str:
        """Get notes for a vendor item."""
        query = """
        query($itemId: [ID!]!) {
            items(ids: $itemId) {
                column_values {
                    id
                    text
                    value
                }
            }
        }
        """
        data = self._query(query, {"itemId": [item_id]})
        items = data.get("items", [])
        if not items:
            return ""

        for cv in items[0].get("column_values", []):
            if cv["id"] == notes_column_id:
                return cv.get("text", "")

        return ""

    def create_task(self, board_id: str, group_id: str, task_name: str,
                    column_values: dict | None = None) -> str:
        """Create a new task on a board. Returns the item ID."""
        import json
        query = """
        mutation($boardId: ID!, $groupId: String!, $itemName: String!, $columnValues: JSON) {
            create_item(
                board_id: $boardId,
                group_id: $groupId,
                item_name: $itemName,
                column_values: $columnValues
            ) {
                id
            }
        }
        """
        variables = {
            "boardId": board_id,
            "groupId": group_id,
            "itemName": task_name,
            "columnValues": json.dumps(column_values or {}),
        }
        data = self._query(query, variables)
        item_id = data.get("create_item", {}).get("id", "")
        logger.info(f"Created Monday.com task {item_id}: {task_name}")
        return item_id

    def update_column_value(self, board_id: str, item_id: str,
                            column_id: str, value: str) -> None:
        """Update a column value on an item."""
        import json
        query = """
        mutation($boardId: ID!, $itemId: ID!, $columnId: String!, $value: JSON!) {
            change_column_value(
                board_id: $boardId,
                item_id: $itemId,
                column_id: $columnId,
                value: $value
            ) {
                id
            }
        }
        """
        self._query(query, {
            "boardId": board_id,
            "itemId": item_id,
            "columnId": column_id,
            "value": json.dumps(value),
        })
        logger.info(f"Updated Monday.com item {item_id} column {column_id}")

    def get_contacts_for_vendor(self, board_id: str, vendor_name: str) -> list[dict[str, str]]:
        """Get contact records for a vendor from the contacts board."""
        query = """
        query($boardId: [ID!]!) {
            boards(ids: $boardId) {
                items_page(limit: 200) {
                    items {
                        id
                        name
                        column_values {
                            id
                            text
                        }
                    }
                }
            }
        }
        """
        data = self._query(query, {"boardId": [board_id]})
        boards = data.get("boards", [])
        if not boards:
            return []

        items = boards[0].get("items_page", {}).get("items", [])
        contacts = []
        vendor_lower = vendor_name.lower()

        for item in items:
            cols = {cv["id"]: cv["text"] for cv in item.get("column_values", [])}
            item_text = (item["name"] + " " + " ".join(cols.values())).lower()
            if vendor_lower in item_text:
                contacts.append({
                    "id": item["id"],
                    "name": item["name"],
                    **cols,
                })

        return contacts
