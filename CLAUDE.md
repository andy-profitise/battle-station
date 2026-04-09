# Claude Code Notes

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
