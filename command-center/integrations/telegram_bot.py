"""
Telegram bot integration for ad-hoc work items.

Listens for messages from authorized chats and injects urgent work items
into the priority queue.

Expected message format:
  @vendor_name: action description
  or
  vendor_name - action description

Example:
  @Acme Corp: they're asking about the new rates, need to respond ASAP
  MediaBuy Inc - check their monthly returns
"""
from __future__ import annotations

import logging
import re
import threading
from typing import Callable

from config import settings
from core.models import WorkItem, WorkItemSource

logger = logging.getLogger(__name__)


class TelegramListener:
    """
    Listens to Telegram for ad-hoc vendor work items.

    Uses polling (not webhooks) for simplicity. In production,
    switch to webhook mode for lower latency.
    """

    def __init__(self, on_message: Callable[[str, str], None]):
        """
        Args:
            on_message: Callback(vendor_name, message_text) when a valid message arrives.
        """
        self.on_message = on_message
        self._running = False
        self._thread: threading.Thread | None = None

    def start(self) -> None:
        """Start listening for Telegram messages in a background thread."""
        if not settings.TELEGRAM_BOT_TOKEN:
            logger.warning("Telegram bot token not configured, skipping")
            return

        self._running = True
        self._thread = threading.Thread(target=self._poll_loop, daemon=True)
        self._thread.start()
        logger.info("Telegram listener started")

    def stop(self) -> None:
        self._running = False
        if self._thread:
            self._thread.join(timeout=5)

    def _poll_loop(self) -> None:
        """Poll Telegram for new messages."""
        import requests
        import time

        base_url = f"https://api.telegram.org/bot{settings.TELEGRAM_BOT_TOKEN}"
        offset = 0

        while self._running:
            try:
                response = requests.get(
                    f"{base_url}/getUpdates",
                    params={"offset": offset, "timeout": 30},
                    timeout=35,
                )
                data = response.json()

                for update in data.get("result", []):
                    offset = update["update_id"] + 1
                    message = update.get("message", {})
                    chat_id = str(message.get("chat", {}).get("id", ""))
                    text = message.get("text", "")

                    # Check authorization
                    if chat_id not in settings.TELEGRAM_ALLOWED_CHAT_IDS:
                        logger.debug(f"Ignoring message from unauthorized chat {chat_id}")
                        continue

                    # Parse vendor and action from message
                    vendor, action = self._parse_message(text)
                    if vendor and action:
                        logger.info(f"Telegram ad-hoc: {vendor} -> {action}")
                        self.on_message(vendor, action)

            except Exception as e:
                logger.error(f"Telegram polling error: {e}")
                time.sleep(5)

    @staticmethod
    def _parse_message(text: str) -> tuple[str | None, str | None]:
        """
        Parse a message into (vendor_name, action_text).

        Supports formats:
          @Vendor Name: action text
          @Vendor Name - action text
          Vendor Name: action text
          Vendor Name - action text
        """
        if not text or text.startswith("/"):
            return None, None

        # Try @vendor: action format
        match = re.match(r'@?(.+?)\s*[-:]\s*(.+)', text, re.DOTALL)
        if match:
            vendor = match.group(1).strip()
            action = match.group(2).strip()
            if vendor and action:
                return vendor, action

        return None, None
