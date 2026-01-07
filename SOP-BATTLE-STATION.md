# Battle Station Development SOP

## Project Overview

**Battle Station (A(I)DEN)** is a Google Apps Script vendor management dashboard that aggregates data from multiple sources into a Google Sheet interface.

### Data Sources
- **monday.com** - Contacts, tasks, notes, helpful links, status
- **Gmail** - Email threads with vendor labels (`zzzVendors/<vendor>`)
- **Google Calendar** - Upcoming meetings
- **Google Drive** - Vendor folder files
- **Box.com** - Document storage
- **Airtable** - Contracts data
- **Claude AI** - Crystal Ball email analysis

### Key Files
```
src/
├── battleStation.gs   # Main dashboard (~12K lines)
│   ├── loadVendorData()        # Load single vendor into portal
│   ├── turboTraverseAll()      # Batch process vendors for checksums
│   ├── skipToNextChanged()     # Navigate to next changed vendor
│   ├── battleStationNext/Prev  # Navigation functions
│   └── generateModuleChecksums_() # Change detection
│
└── buildList.gs       # Vendor list builder (~880 lines)
    └── buildListWithGmailAndNotes() # Build/refresh vendor list
```

### Config Location
```javascript
const BS_CFG = {
  CODE_VERSION: '2025-12-31 10:22AM PST',  // Update this for deployments
  // ... board IDs, sheet names, colors, etc.
}
```

---

## Development Workflow

### 1. Branch Naming
Always develop on a branch starting with `claude/` and ending with the session ID:
```
claude/battle-station-<description>-<sessionId>
```

### 2. Making Changes
1. Read the file first before editing
2. Make targeted edits (prefer Edit over Write for existing files)
3. Update `CODE_VERSION` timestamp after changes
4. Commit with descriptive message
5. Push to branch

### 3. Deployment (Auto via GitHub Actions)
```bash
# After pushing, create PR:
# https://github.com/andy-profitise/battle-station/compare/main...<branch-name>

# Merge PR to main -> Auto-deploys to Google Apps Script
```

### 4. Verify Deployment
- Check version number in the portal UI (top nav bar shows `v<timestamp>`)
- Run the function and check execution logs

---

## Key Function Hierarchies

### loadVendorData() - Single Vendor Load
```
loadVendorData(vendorIndex, options)
├── getVendorContacts_()          → monday.com contacts + Airtable contracts
├── getUpcomingMeetingsForVendor_() → Google Calendar
├── getHelpfulLinksForVendor_()   → monday.com helpful links  [SKIP in turbo]
├── loadBoxDocsForVendor_()       → Box.com files             [SKIP in turbo]
├── getGDriveFilesForVendor_()    → Google Drive files        [SKIP in turbo]
├── getEmailsForVendor_()         → Gmail threads
├── getTasksForVendor_()          → monday.com tasks
├── getCrystalBallData_()         → Claude AI analysis        [SKIP in turbo]
├── generateModuleChecksums_()    → Change detection
└── [Screen writes to spreadsheet]
```

### turboTraverseAll() - Batch Processing
- Processes vendors sequentially to update checksums
- 5-minute max runtime (GAS limit is 6 min)
- Saves progress in ScriptProperties for resume
- Skips expensive operations (Box, GDrive, Helpful Links, Crystal Ball)

### buildListWithGmailAndNotes() - Build Vendor List
```
buildListWithGmailAndNotes()
├── Read: Buyers L1M/L6M, Affiliates L1M/L6M, monday.com sheets
├── readBlacklist_() → Settings sheet
├── buildStatusMaps_() / buildNotesMaps_()
├── getHotVendorsFromGmail_() → Priority zone detection
└── Write to List sheet (sorted by priority zones)
```

---

## Turbo Mode Optimizations

When `turboMode: true` is passed to loadVendorData, these are skipped for speed:

| Module | Skip Pattern | Reason |
|--------|--------------|--------|
| Box.com | `if (turboMode) { Logger.log('Skipping...'); }` | 2+ min pre-fetch |
| Google Drive | Same pattern | API calls |
| Helpful Links | Same pattern | API calls |
| Crystal Ball | Same pattern | Redundant with Gmail |
| Calendar | Same pattern | Not essential for checksums |

### Adding New Turbo Skips
```javascript
// Pattern to follow:
let dataVar = [];  // or default value
if (turboMode) {
  Logger.log('Skipping <Module> in turbo mode');
} else {
  ss.toast('Loading...', 'Icon', 2);
  dataVar = getDataFunction_();
}
```

---

## Navigation Features

### Wrap-Around Navigation
- **Next Vendor**: At last vendor (#311), wraps to #1
- **Skip Unchanged**: At end of list, wraps to #1 and continues searching
- Tracks start position to prevent infinite loops

### Skip 5 & Return
- Finds 5 changed vendors, then returns to origin
- Does NOT wrap around (returns to origin at end of list)

---

## Change Detection (Checksums)

Each vendor has module checksums stored in the "Checksums" sheet:
- `emails`, `tasks`, `notes`, `status`, `contracts`, `helpfulLinks`, `meetings`, `boxDocs`, `gDriveFiles`, `contacts`

When checksums differ from stored values → vendor has changes.

---

## Common Tasks

### Skip a module in Turbo mode
1. Find the module's data fetch in `loadVendorData()`
2. Wrap in `if (turboMode) { skip } else { fetch }` pattern
3. Ensure variable is initialized with default value before the check

### Add a new data source
1. Create `get<Source>ForVendor_()` function
2. Add call in `loadVendorData()`
3. Add to checksum generation if tracking changes
4. Consider turbo mode skip if expensive

### Debug slow performance
1. Check GAS execution logs for timing
2. Look for API calls that could be cached
3. Consider adding turbo mode skip for expensive operations

---

## Timestamps

Always update `CODE_VERSION` when making changes:
```javascript
CODE_VERSION: 'YYYY-MM-DD HH:MMAM/PM PST',
```

This displays in the portal nav bar to confirm deployments.

---

## Useful Commands

```bash
# View current branch
git branch

# Check what's changed
git diff --stat

# Commit pattern
git add src/battleStation.gs && git commit -m "$(cat <<'EOF'
Short description

- Bullet point details
- Updated version timestamp
EOF
)"

# Push
git push -u origin <branch-name>
```

---

## PR URL Pattern
```
https://github.com/andy-profitise/battle-station/compare/main...<branch-name>
```
