"""
Airtable integration for the Vendor Command Center.
"""
from __future__ import annotations

import logging
from typing import Any

import requests

from config import settings
from core.models import AirtableContract

logger = logging.getLogger(__name__)


class AirtableClient:
    """Airtable API client for contract data."""

    BASE_URL = "https://api.airtable.com/v0"

    def __init__(self):
        self._headers = {
            "Authorization": f"Bearer {settings.AIRTABLE_API_KEY}",
            "Content-Type": "application/json",
        }

    def _get(self, table_name: str, params: dict | None = None) -> list[dict[str, Any]]:
        url = f"{self.BASE_URL}/{settings.AIRTABLE_BASE_ID}/{table_name}"
        all_records = []
        offset = None

        while True:
            req_params = dict(params or {})
            if offset:
                req_params["offset"] = offset

            response = requests.get(url, headers=self._headers, params=req_params)
            response.raise_for_status()
            data = response.json()

            all_records.extend(data.get("records", []))
            offset = data.get("offset")
            if not offset:
                break

        return all_records

    def get_vendor_contracts(self, vendor_name: str,
                             table_name: str = "Contracts") -> list[AirtableContract]:
        """Get contract records for a vendor."""
        records = self._get(table_name, {
            "filterByFormula": f"FIND(LOWER(\"{vendor_name}\"), LOWER({{Vendor}}))",
        })

        return [
            AirtableContract(
                record_id=r["id"],
                vendor_name=r.get("fields", {}).get("Vendor", vendor_name),
                contract_type=r.get("fields", {}).get("Type"),
                start_date=r.get("fields", {}).get("Start Date"),
                end_date=r.get("fields", {}).get("End Date"),
                status=r.get("fields", {}).get("Status"),
                values=r.get("fields", {}),
            )
            for r in records
        ]
