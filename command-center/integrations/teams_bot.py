"""
Microsoft Teams integration for ad-hoc work items.

Provides a simple webhook endpoint that receives messages from Teams
and injects them as urgent work items into the priority queue.

In Teams, set up an Outgoing Webhook or Power Automate flow that
POSTs to this endpoint when messages are received in a specific channel.
"""
from __future__ import annotations

import json
import logging
from http.server import HTTPServer, BaseHTTPRequestHandler
import threading
from typing import Callable

from integrations.telegram_bot import TelegramListener  # Reuse message parser

logger = logging.getLogger(__name__)


class TeamsWebhookHandler(BaseHTTPRequestHandler):
    """HTTP handler for Teams webhook messages."""

    callback: Callable[[str, str], None] | None = None

    def do_POST(self):
        content_length = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(content_length).decode("utf-8")

        try:
            data = json.loads(body)
            text = data.get("text", "")

            vendor, action = TelegramListener._parse_message(text)
            if vendor and action and self.callback:
                logger.info(f"Teams ad-hoc: {vendor} -> {action}")
                self.callback(vendor, action)
                self.send_response(200)
                self.end_headers()
                self.wfile.write(b'{"status": "ok"}')
            else:
                self.send_response(400)
                self.end_headers()
                self.wfile.write(b'{"error": "Could not parse vendor and action"}')

        except Exception as e:
            logger.error(f"Teams webhook error: {e}")
            self.send_response(500)
            self.end_headers()

    def log_message(self, format, *args):
        """Suppress default HTTP logging."""
        pass


class TeamsListener:
    """Runs a simple HTTP server to receive Teams webhook messages."""

    def __init__(self, on_message: Callable[[str, str], None], port: int = 8089):
        self.on_message = on_message
        self.port = port
        self._server: HTTPServer | None = None
        self._thread: threading.Thread | None = None

    def start(self) -> None:
        TeamsWebhookHandler.callback = self.on_message
        self._server = HTTPServer(("0.0.0.0", self.port), TeamsWebhookHandler)
        self._thread = threading.Thread(target=self._server.serve_forever, daemon=True)
        self._thread.start()
        logger.info(f"Teams webhook listener started on port {self.port}")

    def stop(self) -> None:
        if self._server:
            self._server.shutdown()
