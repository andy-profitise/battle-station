# Battle Station - Claude Code SOP

## Project Overview

**Battle Station (A(I)DEN)** is a vendor management system for Profitise. It aggregates data from Gmail, monday.com, Airtable, Google Sheets, and Google Calendar into a single workflow for reviewing vendor accounts.

This repo contains the Google Apps Script source (`src/battleStation.gs`, `src/buildList.gs`) that powers the Google Sheets dashboard. Claude Code sessions interact with the same data sources via MCP servers to run the **Account Analysis** workflow conversationally in the terminal.

## MCP Servers

Configured in `.claude/mcp.json`:

| Server | Package | Auth |
|--------|---------|------|
| **monday.com** | `@mondaydotcomorg/monday-api-mcp` | API token in env |
| **Airtable** | `airtable-mcp-server` | PAT in env |
| **Google Workspace** | `@googleworkspace/cli mcp` | OAuth (browser prompt on first use) |

Run `/mcp` at session start to verify all three are connected.

## Key Reference IDs

### Google Sheets
- **Spreadsheet ID**: `1dXWu3Uet1_659F6nYGF5rjpQbFgJlBD-sd9XTw-8hUU`
- **List tab**: `List` — the vendor list with columns: Vendor (A), TTL USD (B), Source (C), Status (D), Notes (E), Gmail Link (F), No Snooze (G), Processed (H)
- **A(I)DEN tab**: The dashboard view (not used directly in terminal workflow)

### monday.com Board IDs
| Board | ID |
|-------|-----|
| Buyers | `9007735194` |
| Affiliates | `9007716156` |
| Tasks | `9007661294` |
| Contacts | `9304296922` |
| Helpful Links | `18389463592` |

### monday.com Column IDs
| Column | Buyers | Affiliates |
|--------|--------|------------|
| Notes | `text_mkqnvsqh` | `text_mkrdahqz` |
| Contacts | `board_relation_mky0bt0z` | `board_relation_mky0n0rf` |
| Phonexa | `link_mksmwprd` | `link_mksmgnc0` |
| Live Verticals | `tag_mkskgt84` | `tag_mkskrddx` |
| Other Verticals | `tag_mkskewmq` | `tag_mkskfs70` |
| Live Modalities | `tag_mkskfmf3` | `tag_mksk7whx` |
| States (Buyers only) | `dropdown_mkyam4qw` | N/A |
| Dead States (Buyers only) | `dropdown_mkyazy2j` | N/A |
| Other Name | `text_mkvkr178` | `text_mksmcrpw` |

### Airtable
- **Base ID**: `appc6xu9qLlOP5G5m`
- **Contracts 2026**: table `tblYszexANBGGnyki`, view `viwfL5bDrEAvlmL7f`
- **Contracts 2025**: table `tblREBd6zFUUZV5eU`, view `viw8X7acqwTJEUi1R`
- **Contracts 2024**: table `tblYn8yBux9xe6sO0`, view `viwfGEvHlo8mT5FBX`
- **Key fields**: `Vendor Name`, `Status`, `Contract Type`, `Notes`, `Submitted By`, `Vertical`, `Created Date`
- **Filter by**: Submitted By in [`Andy Worford`, `Aden Ritz`], Vertical in [`Home Services`, `Solar`]

### Gmail
- **Vendor labels**: `zzzVendors/<vendor-name>` (e.g., `zzzVendors/acme-corp`)
- **Overdue labels**: `03.overdue/manual`, `02.waiting/customer`, `02.waiting/me`
- Pull **90 days** of email threads for analysis

## Account Analysis Workflow

This is the primary workflow. Run it conversationally, one vendor at a time.

### Step 1: Pick a Vendor
Read the **List** tab from the Google Sheet. Present the vendor list (or let the user pick). Each row has:
- **Vendor** name
- **TTL USD** (total spend)
- **Source** (`Buyer` or `Affiliate` — determines which monday.com board to use)
- **Status**, **Notes**, **Processed** flag

### Step 2: Gather Data (run in parallel where possible)

**Gmail** — Search for the vendor's label `zzzVendors/<slug>` and pull the last 90 days of email threads. Summarize: who's emailing, key topics, any unanswered threads, anything overdue.

**monday.com** — Search the appropriate board (Buyers `9007735194` or Affiliates `9007716156` based on Source column):
- Get the vendor's item: notes, status, contacts, live verticals, modalities, states
- Check the Tasks board (`9007661294`) for any open tasks mentioning this vendor
- Check Contacts board (`9304296922`) for contact details

**Airtable** — Search contracts tables (2026, 2025, 2024) for records matching the vendor name. Pull contract type, status, notes, vertical, created date.

### Step 3: Synthesize
Present a unified account briefing:
- **Status & Health**: What's the vendor's current status? Active? Issues?
- **Recent Communication**: Summary of email activity, who's talking, what about
- **Open Items**: Tasks, blockers, unanswered emails
- **Contracts**: Active contracts, any expiring or issues
- **Contacts**: Key people and their roles
- **Notes**: Current notes from monday.com

### Step 4: User Corrections
The user reviews and corrects your synthesis. Take notes on corrections.

**GENERAL notes**: If the user provides notes that apply to all vendors (process notes, team updates, policy changes), save these and carry them forward to subsequent vendors in the session.

### Step 5: Takeaways & Tasks
Work with the user to identify:
- Action items and who owns them
- Updated notes/blockers for the vendor
- Any tasks to create in monday.com

### Step 6: Update monday.com
Using the monday.com MCP tools:
- **Update Notes**: Use `change_item_column_values` on the vendor's item with the notes column ID (`text_mkqnvsqh` for Buyers, `text_mkrdahqz` for Affiliates)
- **Update Blocker/Status**: If the user identifies a blocker or status change, update the appropriate column
- Confirm the update with the user before writing

### Step 7: Next Vendor
Move to the next vendor. Carry forward any GENERAL notes from the session.

## Determining Buyer vs Affiliate

The **Source** column in the List tab tells you which board to use:
- Contains `buyer` (case-insensitive) → Buyers board `9007735194`
- Contains `affiliate` (case-insensitive) → Affiliates board `9007716156`
- Default/unclear → Buyers board

## Development SOP

See `SOP-BATTLE-STATION.md` for the full development workflow (branching, deployment, code structure). Key points:
- Branch naming: `claude/<description>-<sessionId>`
- Always update `CODE_VERSION` timestamp after code changes
- Config is in `BS_CFG` object at top of `src/battleStation.gs`
- Auto-deploys to Google Apps Script on merge to main
