"""
Priority queue with interrupt support for the Vendor Command Center.
Handles email work items and ad-hoc interrupts from Telegram/Teams.
"""
from __future__ import annotations

import heapq
import threading
from collections.abc import Iterator
from datetime import datetime

from core.models import Priority, WorkItem, WorkItemSource


class VendorWorkQueue:
    """
    Thread-safe priority queue for vendor work items.

    Priority order (lowest number = highest priority):
      1. URGENT  - Ad-hoc from Telegram/Teams
      2. HIGH    - Inbox emails
      3. MEDIUM  - 00.received emails
      4. LOW     - Proactive vendor sweep

    Within each priority level, items are ordered oldest-first.
    """

    def __init__(self):
        self._heap: list[WorkItem] = []
        self._lock = threading.Lock()
        self._processed: set[str] = set()   # Track processed item IDs

    def push(self, item: WorkItem) -> None:
        with self._lock:
            heapq.heappush(self._heap, item)

    def pop(self) -> WorkItem | None:
        with self._lock:
            while self._heap:
                item = heapq.heappop(self._heap)
                if item.id not in self._processed:
                    return item
            return None

    def peek(self) -> WorkItem | None:
        with self._lock:
            if self._heap:
                return self._heap[0]
            return None

    def mark_processed(self, item_id: str) -> None:
        with self._lock:
            self._processed.add(item_id)

    def push_urgent(self, vendor_name: str, message: str, source: WorkItemSource) -> WorkItem:
        """Push an ad-hoc urgent item from Telegram or Teams."""
        item = WorkItem(
            source=source,
            priority=Priority.URGENT,
            vendor_name=vendor_name,
            message_text=message,
            message_date=datetime.now(),
        )
        self.push(item)
        return item

    def push_email(self, vendor_name: str, vendor_slug: str,
                   thread_id: str, message_date: datetime,
                   is_inbox: bool = True) -> WorkItem:
        """Push an email work item."""
        item = WorkItem(
            source=WorkItemSource.GMAIL_INBOX if is_inbox else WorkItemSource.GMAIL_RECEIVED,
            priority=Priority.HIGH if is_inbox else Priority.MEDIUM,
            vendor_name=vendor_name,
            vendor_slug=vendor_slug,
            gmail_thread_id=thread_id,
            message_date=message_date,
        )
        self.push(item)
        return item

    def push_proactive(self, vendor_name: str, vendor_slug: str,
                       reason: str) -> WorkItem:
        """Push a proactive scan item."""
        item = WorkItem(
            source=WorkItemSource.PROACTIVE_SCAN,
            priority=Priority.LOW,
            vendor_name=vendor_name,
            vendor_slug=vendor_slug,
            message_text=reason,
        )
        self.push(item)
        return item

    @property
    def size(self) -> int:
        with self._lock:
            return len(self._heap)

    @property
    def is_empty(self) -> bool:
        return self.size == 0

    def has_urgent(self) -> bool:
        """Check if there are any urgent items waiting."""
        with self._lock:
            return any(
                item.priority == Priority.URGENT and item.id not in self._processed
                for item in self._heap
            )

    def drain(self) -> Iterator[WorkItem]:
        """Drain all items from the queue in priority order."""
        while True:
            item = self.pop()
            if item is None:
                break
            yield item

    def clear(self) -> None:
        with self._lock:
            self._heap.clear()
            self._processed.clear()

    def stats(self) -> dict[str, int]:
        """Return counts by priority level."""
        with self._lock:
            counts = {p.name: 0 for p in Priority}
            for item in self._heap:
                if item.id not in self._processed:
                    counts[item.priority.name] += 1
            return counts
