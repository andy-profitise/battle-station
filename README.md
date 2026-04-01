# A(I)DEN - Battle Station

A Google Apps Script-based vendor review dashboard that integrates Gmail, monday.com, and Claude AI to streamline vendor management workflows.

## Overview

The Battle Station (A(I)DEN) lets you:
- Navigate through vendors one-by-one in a structured dashboard
- View vendor details, notes, status, and contacts from monday.com
- See all related emails with priority sorting
- View and manage monday.com tasks
- Generate AI-powered email responses with Claude
- Track which vendors you've reviewed
- Snooze vendors until a specific date

## File Structure

```
src/
├── battleStation.gs    # Main dashboard, menus, email responses, navigation
├── buildList.gs        # Builds the vendor list with priority sorting
├── boxDocuments.gs     # Box.com document integration
└── vendorOcr.gs        # OCR for detecting vendors from chat screenshots
```

## Key Configuration

### 1. BS_CFG in `battleStation.gs`

This is the main configuration object. You'll need to update:

```javascript
const BS_CFG = {
  // Sheet names - update if you use different names
  LIST_SHEET: 'List',
  BATTLE_SHEET: 'A(I)DEN',
  TASKS_SHEET: 'monday.com tasks',

  // API Keys - GET YOUR OWN
  MONDAY_API_TOKEN: 'your-monday-api-token',
  CLAUDE_API_KEY: 'your-anthropic-api-key',

  // monday.com Board IDs - find these in your monday.com URLs
  BUYERS_BOARD_ID: 'your-buyers-board-id',
  AFFILIATES_BOARD_ID: 'your-affiliates-board-id',
  TASKS_BOARD_ID: 'your-tasks-board-id',
  CONTACTS_BOARD_ID: 'your-contacts-board-id',
  HELPFUL_LINKS_BOARD_ID: 'your-helpful-links-board-id',

  // monday.com Column IDs - find these via monday.com API or dev tools
  BUYERS_NOTES_COLUMN: 'your-column-id',
  AFFILIATES_NOTES_COLUMN: 'your-column-id',
  // ... etc
};
```

### 2. Build List Configuration in `buildList.gs`

Update the source sheet names to match your data:

```javascript
const SHEET_BUYERS_L1M       = 'Buyers L1M';      // Buyers last 1 month
const SHEET_BUYERS_L6M       = 'Buyers L6M';      // Buyers last 6 months
const SHEET_AFFILIATES_L1M   = 'Affiliates L1M';  // Affiliates last 1 month
const SHEET_AFFILIATES_L6M   = 'Affiliates L6M';  // Affiliates last 6 months
const SHEET_MON_BUYERS       = 'buyers monday.com';
const SHEET_MON_AFFILIATES   = 'affiliates monday.com';
```

Update status priority for sorting vendors with $0 TTL:

```javascript
const STATUS_PRIORITY = [
  'live',
  'onboarding',
  'paused',
  'preonboarding',
  'early talks',
  'other',
  'dead'
];
```

## Gmail Label Structure

The system expects specific Gmail labels for email categorization:

```
00.received          # All received vendor emails
01.priority/
  └── 1              # High priority emails
02.waiting/
  ├── customer       # Waiting on customer response
  └── me             # Waiting on your response
03.noInbox           # Excluded from inbox priority
03.overdue/
  └── manual         # Manually marked as overdue
zzzvendors-{slug}    # Per-vendor labels (auto-generated from vendor name)
```

### How Labels Work

1. **00.received** - Apply to all inbound vendor emails (via Gmail filter)
2. **zzzvendors-{slug}** - Create a label for each vendor (e.g., `zzzvendors-acme-corp`)
3. **02.waiting/customer** - Applied automatically after you send a response
4. **01.priority/1** - High priority emails get a red dot in Crystal Ball

## Priority Zones (Build List)

When you run "Build List", vendors are sorted into priority zones:

| Zone | Description | Detection |
|------|-------------|-----------|
| Inbox | Emails currently in inbox | Gmail inbox search |
| Chat | Detected via OCR from screenshots | Chat OCR feature |
| Monthly Returns | Open Monthly Returns tasks | monday.com tasks |
| Hot | Recent emails (00.received last 7 days) | Gmail label search |
| Normal | Everyone else | Default |

The "Tranche" column (I) shows which zone each vendor is in.

## Required Sheets

Your Google Sheet needs these tabs:

1. **List** - Generated vendor list (created by Build List)
2. **A(I)DEN** - Main dashboard (created by Setup)
3. **Settings** - Configuration options:
   - Gmail Sublabel Mappings (vendor → label slug)
   - Blacklist (vendors to skip)
   - Email Response Instructions
   - Box Blacklist

4. **Source Data Sheets** (your naming may vary):
   - Buyers L1M / L6M
   - Affiliates L1M / L6M
   - buyers monday.com
   - affiliates monday.com

## Setup Steps

1. **Create a new Google Sheet**

2. **Open Apps Script**: Extensions → Apps Script

3. **Copy the code files** into your Apps Script project:
   - battleStation.gs
   - buildList.gs
   - boxDocuments.gs (optional)
   - vendorOcr.gs (optional)

4. **Update configuration**:
   - Set your API keys in `BS_CFG`
   - Update monday.com board/column IDs
   - Update sheet names to match your data

5. **Set up Gmail labels**:
   - Create the label structure above
   - Set up Gmail filters to auto-apply `00.received` and `zzzvendors-{slug}` labels

6. **Run Setup**: Refresh the sheet, then use menu: `A(I)DEN → Setup A(I)DEN`

7. **Build your list**: `A(I)DEN → Build List`

8. **Start reviewing**: Use Navigation menu to go through vendors

## Key Features

### Email Responses
Pre-built templates that use Claude AI to generate contextual responses:
- Cold Outreach Follow Up
- Schedule a Call
- Payment/Invoice Follow Up
- Check Affiliate
- Custom Response

### Vendor Deep Links
Share URLs that open directly to a specific vendor:
1. Deploy as web app (Extensions → Apps Script → Deploy)
2. Set the base URL via Navigation → Set Deep Link URL
3. Use Navigation → Copy Vendor Deep Link

### Crystal Ball
AI-powered email analysis showing:
- Summary of email threads
- Priority indicators
- What's being discussed

### Auto-Archive
After sending an email response:
- Thread is tracked for auto-archive
- On next refresh, labels swap from `02.waiting/me` → `02.waiting/customer`
- Thread is archived automatically

## Finding monday.com IDs

### Board IDs
Look at the URL when viewing a board:
```
https://profitise.monday.com/boards/9007735194
                                    ^^^^^^^^^^
                                    Board ID
```

### Column IDs
Use monday.com's API or browser dev tools:
1. Open dev tools (F12)
2. Go to Network tab
3. Perform an action on the board
4. Look for GraphQL requests
5. Find column IDs in the request/response

Or use the monday.com API Explorer:
```graphql
query {
  boards(ids: YOUR_BOARD_ID) {
    columns {
      id
      title
    }
  }
}
```

## Customization Tips

### Add New Email Response Types
1. Add menu item in `onOpen()` function
2. Create function that calls `generateEmailResponse_('Your Type Name')`
3. Optionally add built-in instructions in `generateEmailWithClaude_()`

### Modify Priority Zones
Edit the zone detection logic in `buildListWithGmailAndNotes()`:
```javascript
for (const r of all) {
  const nameLower = r.name.toLowerCase();
  if (inboxSet.has(nameLower)) {
    r.tranche = 'Inbox';
    inboxZone.push(r);
  } else if (/* your custom condition */) {
    r.tranche = 'Custom Zone';
    customZone.push(r);
  }
  // ... etc
}
```

### Change Color Scheme
Update the `COLOR_*` constants in `BS_CFG`.

## Troubleshooting

### "Permission denied" errors
- Re-authorize the script: Run any function and approve permissions
- Make sure Gmail API is enabled in your GCP project

### Emails not showing
- Check Gmail labels are set up correctly
- Verify the vendor has a matching `zzzvendors-{slug}` label
- Check Settings sheet for Gmail Sublabel Mappings

### monday.com data not loading
- Verify API token is valid
- Check board/column IDs are correct
- Look at Execution Log for detailed errors

## Support

This is an internal tool. For questions, ask Andy or check the code comments.
