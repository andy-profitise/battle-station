"""
Configuration for the Vendor Command Center.
All secrets come from environment variables.
"""
import os


class Settings:
    # Gmail
    GMAIL_CREDENTIALS_PATH = os.getenv("GMAIL_CREDENTIALS_PATH", "credentials/gmail.json")
    GMAIL_TOKEN_PATH = os.getenv("GMAIL_TOKEN_PATH", "credentials/gmail_token.json")
    GMAIL_INBOX_QUERY = "label:inbox OR label:00.received AND -is:snoozed"
    GMAIL_VENDOR_LABEL_PREFIX = "zzzvendors-"

    # Monday.com
    MONDAY_API_TOKEN = os.getenv("MONDAY_API_TOKEN", "")
    MONDAY_API_URL = "https://api.monday.com/v2"
    MONDAY_BUYERS_BOARD_ID = os.getenv("MONDAY_BUYERS_BOARD_ID", "")
    MONDAY_AFFILIATES_BOARD_ID = os.getenv("MONDAY_AFFILIATES_BOARD_ID", "")
    MONDAY_TASKS_BOARD_ID = os.getenv("MONDAY_TASKS_BOARD_ID", "")
    MONDAY_CONTACTS_BOARD_ID = os.getenv("MONDAY_CONTACTS_BOARD_ID", "")
    MONDAY_HELPFUL_LINKS_BOARD_ID = os.getenv("MONDAY_HELPFUL_LINKS_BOARD_ID", "")

    # Box.com
    BOX_CLIENT_ID = os.getenv("BOX_CLIENT_ID", "")
    BOX_CLIENT_SECRET = os.getenv("BOX_CLIENT_SECRET", "")
    BOX_ACCESS_TOKEN = os.getenv("BOX_ACCESS_TOKEN", "")

    # Airtable
    AIRTABLE_API_KEY = os.getenv("AIRTABLE_API_KEY", "")
    AIRTABLE_BASE_ID = os.getenv("AIRTABLE_BASE_ID", "")

    # Google Calendar
    CALENDAR_CREDENTIALS_PATH = os.getenv("CALENDAR_CREDENTIALS_PATH", "credentials/gmail.json")

    # Claude AI
    ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "")
    CLAUDE_MODEL = os.getenv("CLAUDE_MODEL", "claude-sonnet-4-6")

    # Telegram
    TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "")
    TELEGRAM_ALLOWED_CHAT_IDS = os.getenv("TELEGRAM_ALLOWED_CHAT_IDS", "").split(",")

    # Teams
    TEAMS_WEBHOOK_URL = os.getenv("TEAMS_WEBHOOK_URL", "")

    # Processing
    MAX_EMAILS_PER_CYCLE = int(os.getenv("MAX_EMAILS_PER_CYCLE", "20"))
    APPROVAL_MODE = os.getenv("APPROVAL_MODE", "draft")  # auto, draft, flag
    PROACTIVE_SCAN_ENABLED = os.getenv("PROACTIVE_SCAN_ENABLED", "true").lower() == "true"

    # Logging
    LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO")


settings = Settings()
