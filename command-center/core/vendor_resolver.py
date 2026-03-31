"""
Vendor resolver - maps emails, messages, and labels to vendor identities.

Uses Gmail labels (zzzvendors-{slug}) and Monday.com board data to resolve
who a vendor is from an email address, thread labels, or message text.
"""
from __future__ import annotations

import logging
import re

logger = logging.getLogger(__name__)


class VendorResolver:
    """
    Resolves vendor identity from various signals.

    The primary mapping is Gmail label → vendor name, built from Monday.com boards.
    """

    def __init__(self):
        # slug → vendor name
        self._slug_to_name: dict[str, str] = {}
        # email domain → vendor slug
        self._domain_to_slug: dict[str, str] = {}
        # vendor name (lowercase) → slug
        self._name_to_slug: dict[str, str] = {}

    def load_vendor_map(self, vendors: list[dict[str, str]]) -> None:
        """
        Load vendor mapping from Monday.com board data.

        Each vendor dict should have at minimum: {name, slug}
        Optionally: {email_domains: ["domain1.com", "domain2.com"]}
        """
        for v in vendors:
            name = v.get("name", "")
            slug = v.get("slug", "")
            if not name or not slug:
                continue

            self._slug_to_name[slug] = name
            self._name_to_slug[name.lower()] = slug

            for domain in v.get("email_domains", []):
                self._domain_to_slug[domain.lower()] = slug

        logger.info(f"Loaded {len(self._slug_to_name)} vendor mappings")

    def resolve_from_labels(self, labels: list[str]) -> tuple[str | None, str | None]:
        """
        Resolve vendor from Gmail thread labels.
        Returns (vendor_name, vendor_slug) or (None, None).
        """
        for label in labels:
            label_lower = label.lower()
            if label_lower.startswith("zzzvendors-"):
                slug = label_lower.replace("zzzvendors-", "")
                # Also handle the "zzzvendors/" variant
                slug = slug.replace("/", "-")
                name = self._slug_to_name.get(slug)
                if name:
                    return name, slug
                # Slug exists in label but not in our map - return slug as name
                return slug.replace("-", " ").title(), slug
        return None, None

    def resolve_from_email(self, email_address: str) -> tuple[str | None, str | None]:
        """
        Resolve vendor from an email address using domain mapping.
        Returns (vendor_name, vendor_slug) or (None, None).
        """
        if not email_address or "@" not in email_address:
            return None, None
        domain = email_address.split("@")[1].lower()
        slug = self._domain_to_slug.get(domain)
        if slug:
            return self._slug_to_name.get(slug, slug), slug
        return None, None

    def resolve_from_name(self, vendor_name: str) -> tuple[str | None, str | None]:
        """
        Resolve vendor from a name string (e.g., from Telegram message).
        Uses fuzzy matching against known vendor names.
        Returns (vendor_name, vendor_slug) or (None, None).
        """
        name_lower = vendor_name.strip().lower()

        # Exact match
        slug = self._name_to_slug.get(name_lower)
        if slug:
            return self._slug_to_name[slug], slug

        # Partial match - vendor name contains the search term
        for known_name, slug in self._name_to_slug.items():
            if name_lower in known_name or known_name in name_lower:
                return self._slug_to_name[slug], slug

        return None, None

    def slug_for_name(self, vendor_name: str) -> str:
        """Generate a slug from a vendor name."""
        slug = vendor_name.lower().strip()
        slug = re.sub(r'[^a-z0-9]+', '-', slug)
        slug = slug.strip('-')
        return slug

    @property
    def vendor_count(self) -> int:
        return len(self._slug_to_name)

    def all_vendors(self) -> list[tuple[str, str]]:
        """Return all (name, slug) pairs."""
        return [(name, slug) for slug, name in self._slug_to_name.items()]
