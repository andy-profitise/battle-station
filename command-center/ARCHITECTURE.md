# Vendor Command Center - Architecture

## Overview

The Vendor Command Center is an autonomous agent loop that proactively manages vendor
relationships. It replaces the reactive "navigate and review" workflow of the Battle
Station with an intelligent processing pipeline that picks up work, gathers context,
reasons about it, and takes action.

## Processing Loop

```
┌─────────────────────────────────────────────────────────┐
│                    PRIORITY QUEUE                        │
│                                                         │
│  1. Ad-hoc interrupts (Telegram/Teams)     [URGENT]     │
│  2. Inbox emails (oldest first)            [HIGH]       │
│  3. 00.received emails (oldest first)      [MEDIUM]     │
│  4. Proactive vendor sweep (Monday.com)    [LOW]        │
│                                                         │
└──────────────────────┬──────────────────────────────────┘
                       │
                       ▼
┌─────────────────────────────────────────────────────────┐
│              PICK NEXT WORK ITEM                        │
│                                                         │
│  Identify the vendor from the work item                 │
│  (email sender → vendor mapping, or direct reference)   │
└──────────────────────┬──────────────────────────────────┘
                       │
                       ▼
┌─────────────────────────────────────────────────────────┐
│           GATHER VENDOR CONTEXT (before reading email)  │
│                                                         │
│  ┌──────────┐ ┌──────────┐ ┌──────────┐ ┌──────────┐   │
│  │Monday.com│ │ Calendar │ │  Box.com │ │ Airtable │   │
│  │ Status   │ │ Meetings │ │   Docs   │ │Contracts │   │
│  │ Tasks    │ │          │ │          │ │          │   │
│  │ Notes    │ │          │ │          │ │          │   │
│  │ Links    │ │          │ │          │ │          │   │
│  └──────────┘ └──────────┘ └──────────┘ └──────────┘   │
│                                                         │
│  Output: VendorContext object with full picture          │
└──────────────────────┬──────────────────────────────────┘
                       │
                       ▼
┌─────────────────────────────────────────────────────────┐
│           READ & CONTEXTUALIZE THE TRIGGER              │
│                                                         │
│  Now read the email/message that triggered this work    │
│  Understand it in context of the vendor's full picture  │
│                                                         │
│  Output: Situational assessment                         │
└──────────────────────┬──────────────────────────────────┘
                       │
                       ▼
┌─────────────────────────────────────────────────────────┐
│           REASON & DECIDE (Claude)                      │
│                                                         │
│  Given the vendor context + trigger, determine:         │
│  - What is happening / what does the vendor need?       │
│  - What is our current relationship/status?             │
│  - What actions should we take?                         │
│  - What is the priority and urgency?                    │
│                                                         │
│  Output: ActionPlan with concrete steps                 │
└──────────────────────┬──────────────────────────────────┘
                       │
                       ▼
┌─────────────────────────────────────────────────────────┐
│           EXECUTE ACTIONS                               │
│                                                         │
│  ┌──────────────┐ ┌──────────────┐ ┌────────────────┐   │
│  │ Draft/Send   │ │ Create/Update│ │ Update Status  │   │
│  │ Email Reply  │ │ Monday Tasks │ │ / Notes        │   │
│  └──────────────┘ └──────────────┘ └────────────────┘   │
│  ┌──────────────┐ ┌──────────────┐ ┌────────────────┐   │
│  │ Archive/Label│ │ Schedule     │ │ Flag for Human │   │
│  │ Gmail Thread │ │ Follow-up    │ │ Review         │   │
│  └──────────────┘ └──────────────┘ └────────────────┘   │
│                                                         │
│  Some actions auto-execute, others queue for approval   │
└──────────────────────┬──────────────────────────────────┘
                       │
                       ▼
                  [Next item]
```

## Approval Modes

Not everything should auto-execute. Three modes:

- **Auto**: Archive, label, update notes, create tasks
- **Draft**: Email replies are drafted for human review before sending
- **Flag**: Unusual situations, large financial decisions, new vendor relationships

## Ad-Hoc Queue

Telegram and Teams messages inject work items into the priority queue:

```
Telegram Bot → Webhook → Parse vendor + action → Queue as URGENT
Teams Bot    → Webhook → Parse vendor + action → Queue as URGENT
```

These interrupt the normal email processing flow and get handled first.

## Proactive Vendor Sweep

When the inbox is clear (no emails to process), the system switches to proactive mode:

1. Pull all vendors from Monday.com "Vendors" tab
2. For each vendor, evaluate:
   - Last contact date (are we overdue for a check-in?)
   - Open tasks past due date
   - Contract renewal dates approaching
   - Status changes that need follow-up
   - Revenue trends that need attention
3. Generate proactive outreach or internal action items

## File Structure

```
command-center/
├── ARCHITECTURE.md          # This file
├── main.py                  # Entry point - the main loop
├── config/
│   └── settings.py          # Configuration and environment
├── core/
│   ├── loop.py              # Main processing loop
│   ├── models.py            # Data models (WorkItem, VendorContext, ActionPlan)
│   └── vendor_resolver.py   # Map emails/messages to vendors
├── integrations/
│   ├── gmail.py             # Gmail API - search, read, send, label, archive
│   ├── monday.py            # Monday.com API - status, tasks, notes, contacts
│   ├── box_api.py           # Box.com API - document listing/reading
│   ├── calendar_api.py      # Google Calendar API - meetings
│   ├── airtable.py          # Airtable API - contracts
│   ├── telegram_bot.py      # Telegram bot for ad-hoc items
│   └── teams_bot.py         # Teams bot for ad-hoc items
├── agents/
│   ├── context_builder.py   # Builds full vendor context from all sources
│   ├── reasoner.py          # Claude-powered reasoning about what to do
│   ├── email_composer.py    # Composes contextual email responses
│   └── proactive_scanner.py # Scans vendors for proactive opportunities
└── queue/
    ├── priority_queue.py    # Priority queue with interrupt support
    └── work_items.py        # Work item types and lifecycle
```

## Integration with Battle Station

The Command Center does NOT replace the Battle Station. It uses the same:
- Gmail labels (zzzvendors-*, 00.received, etc.)
- Monday.com boards and columns
- Box.com folders
- Google Calendar
- Airtable bases

The Battle Station remains available as a manual deep-dive tool. The Command Center
handles the flow and triage, and can link back to the Battle Station for complex
situations that need human judgment.
