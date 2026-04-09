# Claude Code Notes

## What This Project Is

**Battle Station (A(I)DEN)** is Andy Worford's vendor relationship management system at Profitise (lead generation company). It's a Google Apps Script project that runs inside Google Sheets and integrates Gmail, monday.com, Google Calendar, Airtable, Box.com, and Claude AI.

**The user is Andy Worford.** He manages vendor relationships (buyers and affiliates). The system helps him track emails, tasks, blockers, contacts, contracts, and notes for each vendor.

## Codebase Structure

All source is in `/src/`:
- **`battleStation.gs`** (~23K lines) — Main file. Everything: UI, menus, vendor loading, AI features, email analysis, navigation, workflows, Account Analysis
- **`buildList.gs`** (~1K lines) — Builds the vendor List from Gmail + monday.com data, sorts into priority zones
- **`boxDocuments.gs`** (~750 lines) — Box.com OAuth and document search
- **`vendorOcr.gs`** (~1.8K lines) — Claude Vision OCR for detecting vendors in chat screenshots
- **`appsscript.json`** — Manifest with API scopes

## Key Architecture

- **Google Apps Script V8 runtime** — 6-minute execution limit, HTML dialogs for UI
- **Config object `BS_CFG`** (top of battleStation.gs, ~lines 30-254) — ALL configuration: sheet names, column indices, API tokens, monday.com board/column IDs, color palette, Airtable config
- **Vendor data lives in multiple places:** monday.com (source of truth for notes/blockers/contacts/tasks), Gmail (emails via labels), Airtable (contracts), Google Sheets (List view, checksums, cache)

## Data Model Quick Reference

### Google Sheets Tabs
- **List** — Vendor list sorted by priority zone (Actionable > Chat > Monthly Returns > Hot > Normal)
- **A(I)DEN** — Main dashboard, dynamically populated per vendor
- **monday.com tasks** — Cached task data
- **BS_Checksums** — Change detection per vendor per module
- **BS_Cache** — Batch cache for Airtable/Box (24hr TTL)
- **Settings** — Gmail sublabel mappings, AI instructions, goals, email rules, Crystal Ball instructions
- **Vendor Review Log** — Inbox Review Q&A records
- **Action Plans** — Vendor Briefing and Account Analysis history
- **Account Analysis Log** — Deep review session records

### monday.com Boards
- **Buyers Board** (9007735194) — Buyer vendor records with notes, blockers, contacts, verticals, modalities, states
- **Affiliates Board** (9007716156) — Same structure for affiliates
- **Tasks Board** (9007661294) — Work items linked to vendors. Has a Blockers group for blocker tasks
- **Contacts Board** (9304296922) — Contact records linked to vendor boards
- **Helpful Links Board** (18389463592) — Resource links per vendor
- **Chat Links Board** (9333292060) — Preferred communication method per vendor

### Gmail Label Structure
- `00.received` — Active/inbox vendor emails (the important ones)
- `01.priority/1` — High priority flag
- `02.waiting/customer` — We're waiting on them
- `02.waiting/me` — They're waiting on us
- `03.overdue/manual` — Manually flagged overdue
- `zzzvendors-{slug}` — Per-vendor label (e.g., `zzzvendors-acme-corp`)

### Airtable
- Base ID: `appc6xu9qLlOP5G5m`
- Tables: Contracts 2024, 2025, 2026
- Filtered by Submitted By (Andy Worford, Aden Ritz) and Vertical (Home Services, Solar)

## Key Functions (Where to Find Things)

### Core Navigation
- `loadVendorData(index, options)` ~line 1380 — Main vendor load, fetches all data
- `getCurrentVendorIndex_()` ~line 1324 — Gets current vendor from nav bar
- `battleStationNext/Previous()` ~line 6764 — Navigation with auto-save

### Data Fetching
- `getVendorContacts_(vendor, listRow)` — monday.com contacts + notes + blockers + verticals
- `getEmailsForVendor_(vendor, listRow)` ~line 3923 — Gmail emails (active inbox)
- `getAllEmailsFromVendorLabel_(listRow, max)` ~line 4021 — Full email history from vendor label
- `getTasksForVendor_(vendor, listRow)` ~line 4215 — monday.com tasks
- `getVendorContracts_(vendor)` ~line 5027 — Airtable contracts
- `getCrystalBallData_(vendor, listRow, options)` ~line 5529 — Real-time email + blocker analysis

### AI Features
- `callClaudeAPI_(prompt, apiKey, options)` ~line 8845 — Central Claude API call
- `battleStationVendorBriefing()` ~line 19992 — Full intel briefing with revision
- `battleStationSmartBriefing()` ~line 12777 — Cross-vendor priority advisor
- `battleStationSummarizeToNotes()` ~line 13146 — Auto-summarize to monday.com notes
- `accountAnalysisStart()` ~line 22730+ — Deep 90-day account analysis (section cards UI)

### monday.com Updates
- `updateMondayComNotesForVendor_(vendor, notes, listRow)` ~line 7497 — Push notes
- `updateMondayBlockersForVendor_(vendor, blockers, listRow)` ~line 13824 — Push blockers
- `createBlockerForVendor()` ~line 13900 — Create blocker task from template

### Shared Utilities
- `getClaudeApiKey_()` ~line 12696 — Get API key (Script Properties or BS_CFG)
- `getAiInstructions_()` ~line 12763 — User's standing AI instructions
- `getGoalsContext_()` ~line 19651 — User's goals for prompt context
- `getGeneralNotes_()` ~line 22437 — Cross-vendor general notes from Action Plans
- `saveBriefingToSheet_(vendor, content)` ~line 22392 — Save to Action Plans sheet
- `escapeHtml_(text)` — HTML escape (defined multiple times due to GAS scoping)

## Account Analysis Feature (Recently Built)

A deep strategic review workflow in `battleStation.gs` (end of file, ~line 22630+):

### How It Works
1. **Gathers** 90 days of email history (full content via `acctAnalysisGetEmails90d_`), plus notes, blockers, tasks, contracts, contacts, previous analysis notes, general cross-vendor context
2. **Claude analyzes** everything into sections: Account Health, Relationship Timeline, What's Working, What's Lacking, What's Missing, Open Items Assessment, Strategic Direction, Recommended Actions, Proposed Notes, Proposed Blocker
3. **Dialog shows sections as collapsible cards** — each with its own note/correction field
4. **User corrects per-section**, hits "Revise With My Notes" — Claude regenerates and cards update in-place. Can revise unlimited times.
5. **Snapshot phase** — edit and save Notes + Blocker to monday.com
6. **Next vendor** — marks processed, loads next vendor in List

### Key Functions
- `acctAnalysisGetEmails90d_(vendor, listRow)` — 90-day email fetch with full content
- `accountAnalysisStart()` — Entry point, gathers all data, calls Claude
- `acctAnalysisShowDialog_(ctx)` — Builds the section-card dialog
- `acctAnalysisParseSections_(rawContent)` — Splits Claude output by ## headers
- `accountAnalysisRevise(feedback)` — Revision with per-section updates + GENERAL: notes
- `accountAnalysisSaveSnapshot(notes, blocker, andNext)` — Save to monday.com + advance
- `accountAnalysisNextUnprocessed()` — Jump to next unreviewed vendor

### Menu Location
A(I)DEN menu > "Account Analysis (Deep Review)" or "Account Analysis (Next Unprocessed)"

## Andy's Preferences

- Likes interactive review workflows: AI proposes, Andy corrects, AI re-runs. Not one-shot outputs.
- Uses `GENERAL:` prefix in corrections to save cross-vendor context (stored in Action Plans sheet under `_GENERAL` row)
- Wants things tight and iterative, not walls of text
- The "snapshot" pattern (propose Notes + Blocker, let Andy edit, save to monday.com) is the standard way to finish reviewing a vendor

## Git Workflow

- **Always fetch and merge main before starting work** on a feature branch: `git fetch origin main && git merge origin/main`. Other sessions may have merged changes concurrently.
- **Always fetch and merge main again before pushing**, to catch anything that landed while you were working. This avoids merge conflicts at PR time.
- This repo has multiple concurrent Claude Code sessions working on it. Never assume main is static.

## MCP Server Setup

Three MCP servers give Claude Code direct access to all Battle Station data sources:

### 1. Google Workspace MCP (Gmail + Sheets + Calendar + Drive)
- **Config:** `.claude/mcp.json` — `google-workspace` server
- **Requires:** Google Cloud OAuth 2.0 credentials (Desktop Application type)
- **Env vars needed:**
  - `GOOGLE_OAUTH_CLIENT_ID` — from GCP Console > APIs & Services > Credentials
  - `GOOGLE_OAUTH_CLIENT_SECRET` — same place
- **APIs to enable in GCP:** Gmail, Sheets, Drive, Calendar
- **First run:** Complete browser OAuth flow

### 2. Airtable MCP (Contracts)
- **Config:** `.claude/mcp.json` — `airtable` server
- **Requires:** Airtable Personal Access Token
- **Env vars needed:**
  - `AIRTABLE_API_KEY` — from Airtable > Account > Tokens
- **Scopes needed:** `schema.bases:read`, `data.records:read`

### 3. monday.com MCP (Tasks, Contacts, Notes, Blockers)
- Already configured globally (not in this project's config)

### Setup Steps
1. Set env vars in your shell profile or `.env` file
2. Restart Claude Code session
3. Run `/mcp` to verify servers are connected
4. Complete Google OAuth browser flow on first use
