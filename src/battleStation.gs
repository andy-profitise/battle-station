/************************************************************
 * BATTLE STATION - One-by-one vendor review dashboard
 *
 * Features:
 * - Navigate through vendors sequentially via menu
 * - View vendor details, notes, status, contacts
 * - See all related emails and monday.com tasks (live search)
 * - View helpful links from monday.com
 * - Update monday.com notes directly
 * - Mark vendors as reviewed/complete
 * - Email contacts directly from Battle Station
 * - Analyze emails with Claude AI (inline links)
 *
 * UPDATED: Snoozed=highest priority, Helpful Links, inline Claude links
 * FIXED: Removed duplicate functions, fixed checksum functionality
 ************************************************************/

const BS_CFG = {
  // Sheet names
  LIST_SHEET: 'List',
  BATTLE_SHEET: 'Battle Station',
  GMAIL_OUTPUT_SHEET: 'Gmail Review Output',
  TASKS_SHEET: 'monday.com tasks',

  // List sheet columns (0-based)
  L_VENDOR: 0,
  L_TTL_USD: 1,
  L_SOURCE: 2,
  L_STATUS: 3,
  L_NOTES: 4,
  L_GMAIL_LINK: 5,
  L_NO_SNOOZE: 6,
  L_PROCESSED: 7,

  // Battle Station layout
  HEADER_ROWS: 3,
  DATA_START_ROW: 5,

  // Colors
  COLOR_HEADER: '#4a86e8',
  COLOR_SUBHEADER: '#6d9eeb',
  COLOR_EMAIL: '#fff2cc',
  COLOR_TASK: '#d9ead3',
  COLOR_LINKS: '#e1d5e7',
  COLOR_BUTTON: '#e8f0fe',
  COLOR_WARNING: '#f4cccc',
  COLOR_SUCCESS: '#d9ead3',
  COLOR_SNOOZED: '#d0e8f2',
  COLOR_WAITING: '#fff44f',
  COLOR_MISSING: '#fce5cd',  // Light orange/red for missing data
  COLOR_PHONEXA: '#f4a460',  // Sandy brown/coral for Phonexa waiting (highest priority)
  COLOR_OVERDUE: '#f4cccc',  // Light red for overdue waiting/customer emails

  // Row highlight colors for skip/traverse
  COLOR_ROW_CHANGED: '#d9ead3',   // Green - vendor has changes, stopped here
  COLOR_ROW_SKIPPED: '#fff2cc',   // Yellow - vendor unchanged, passed over

  // Overdue threshold: emails with "02.waiting/customer" or "02.waiting/me" older than this many business hours
  OVERDUE_BUSINESS_HOURS: 16,


  // API Keys
  MONDAY_API_TOKEN: 'eyJhbGciOiJIUzI1NiJ9.eyJ0aWQiOjQ1NDc2OTk0NywiYWFpIjoxMSwidWlkIjo1MzkyOTA3OCwiaWFkIjoiMjAyNS0wMS0wN1QxODoyNzo1My4wMDBaIiwicGVyIjoibWU6d3JpdGUiLCJhY3RpZCI6MjA1NzI0NjIsInJnbiI6InVzZTEifQ.h8_RIEP9thz-UwIT5SSbkf73n4mzRwyu7ALSZkkTDE8',
  CLAUDE_API_KEY: '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',

  // Search terms to skip (too generic, cause false positives)
  SKIP_SEARCH_TERMS: ['LLC', 'Inc', 'Inc.', 'Corp', 'Corp.', 'Co', 'Co.', 'Ltd', 'Ltd.', 'LP', 'LLP', 'PC', 'PLLC', 'NA', 'N/A'],

  // monday.com Board IDs
  BUYERS_BOARD_ID: '9007735194',
  AFFILIATES_BOARD_ID: '9007716156',
  TASKS_BOARD_ID: '9007661294',
  CONTACTS_BOARD_ID: '9304296922',
  HELPFUL_LINKS_BOARD_ID: '18389463592',

  // monday.com Column IDs
  BUYERS_NOTES_COLUMN: 'text_mkqnvsqh',
  AFFILIATES_NOTES_COLUMN: 'text_mkrdahqz',
  BUYERS_CONTACTS_COLUMN: 'board_relation_mky0bt0z',
  AFFILIATES_CONTACTS_COLUMN: 'board_relation_mky0n0rf',

  // Helpful Links Column IDs
  HELPFUL_LINKS_LINK_COLUMN: 'link_mky0anm4',
  HELPFUL_LINKS_BUYERS_COLUMN: 'board_relation_mky03dt4',
  HELPFUL_LINKS_AFFILIATES_COLUMN: 'board_relation_mky0gxak',
  HELPFUL_LINKS_NOTES_COLUMN: 'text_mky08ybx',

  // Phonexa Link Column IDs
  BUYERS_PHONEXA_COLUMN: 'link_mksmwprd',
  AFFILIATES_PHONEXA_COLUMN: 'link_mksmgnc0',

  // Live Verticals Column IDs
  BUYERS_LIVE_VERTICALS_COLUMN: 'tag_mkskgt84',
  AFFILIATES_LIVE_VERTICALS_COLUMN: 'tag_mkskrddx',

  // Other Verticals Column IDs
  BUYERS_OTHER_VERTICALS_COLUMN: 'tag_mkskewmq',
  AFFILIATES_OTHER_VERTICALS_COLUMN: 'tag_mkskfs70',

  // Live Modalities Column IDs
  BUYERS_LIVE_MODALITIES_COLUMN: 'tag_mkskfmf3',
  AFFILIATES_LIVE_MODALITIES_COLUMN: 'tag_mksk7whx',

  // States Column IDs (Buyers only)
  BUYERS_STATES_COLUMN: 'dropdown_mkyam4qw',
  BUYERS_DEAD_STATES_COLUMN: 'dropdown_mkyazy2j',

  // Other Name Column ID (for alternate vendor names)
  BUYERS_OTHER_NAME_COLUMN: 'text_mkvkr178',
  AFFILIATES_OTHER_NAME_COLUMN: 'text_mksmcrpw',

  // Add to BS_CFG:
  TASKS_PROJECT_COLUMN: 'board_relation_mkqbg3mb',

    // Large-Scale Projects ID to Name mapping
    PROJECT_MAP: {
      '9520665110': 'Home Services',
      '9520665261': 'ACA',
      '9618492546': 'Vertical Activation',
      '9071022704': 'Monthly Returns',
      '9268820620': 'CPL/Zip Optimizations',
      '9520671333': 'Accounting/Invoices',
      '9521113689': 'System Admin',
      '9754457415': 'URL Whitelist',
      '9007621458': 'Outbound Communication',
      '9520669726': 'Pre-Onboarding',
      '9007619323': 'Appointments',
      '9080883844': 'Onboarding - Buyer',
      '9323973905': 'Onboarding - Affiliate',
      '9587318546': 'Onboarding - Vertical',
      '9080886761': 'Templates',
      '9549663466': 'Morning Meeting',
      '9681907462': 'Week of 07/28/25',
    },

  // Last Updated Column IDs
  BUYERS_LAST_UPDATED_INDEX: 15,      // Column P (1-based 16) -> 0-based 15
  AFFILIATES_LAST_UPDATED_INDEX: 16,  // Column Q (1-based 17) -> 0-based 16

  // Checksums
  CHECKSUMS_SHEET: 'BS_Checksums',

  // Cache sheet for Airtable/Box data
  CACHE_SHEET: 'BS_Cache',
  CACHE_MAX_AGE_HOURS: 24,  // Refresh cache if older than this

  // Airtable Contracts Configuration
  AIRTABLE_API_TOKEN: 'pat9P76pQ7lJ8Cwoa.2a614804ad532a6e957e931e65f0e4f228928e9f2ce3d57b5d918dde832842db',
  AIRTABLE_BASE_ID: 'appc6xu9qLlOP5G5m',

  // Contracts 2025
  AIRTABLE_CONTRACTS_TABLE_2025: 'Contracts 2025',
  AIRTABLE_CONTRACTS_TABLE_ID_2025: 'tblREBd6zFUUZV5eU',
  AIRTABLE_CONTRACTS_VIEW_ID_2025: 'viw8X7acqwTJEUi1R',

  // Contracts 2024
  AIRTABLE_CONTRACTS_TABLE_2024: 'Contracts 2024',
  AIRTABLE_CONTRACTS_TABLE_ID_2024: 'tblYn8yBux9xe6sO0',
  AIRTABLE_CONTRACTS_VIEW_ID_2024: 'viwfGEvHlo8mT5FBX',

  AIRTABLE_VENDOR_FIELD: 'Vendor Name',
  AIRTABLE_STATUS_FIELD: 'Status',
  AIRTABLE_CONTRACT_TYPE_FIELD: 'Contract Type',
  AIRTABLE_NOTES_FIELD: 'Notes',
  AIRTABLE_SUBMITTED_BY_FIELD: 'Submitted By',
  AIRTABLE_VERTICAL_FIELD: 'Vertical',
  AIRTABLE_CREATED_DATE_FIELD: 'Created Date',

  // Filter values for contracts
  AIRTABLE_ALLOWED_SUBMITTERS: ['Andy Worford', 'Aden Ritz'],
  AIRTABLE_ALLOWED_VERTICALS: ['Home Services', 'Solar'],

  AIRTABLE_API_BASE_URL: 'https://api.airtable.com/v0',

  // Google Drive Vendors Folder
  GDRIVE_VENDORS_FOLDER_ID: '1fZzQZ_srKJFZab73zE_6hqDLZrn7ud_C',

  // Max characters for notes display (truncate with "..." if longer)
  MAX_NOTES_LENGTH: 400
};

/**
 * Add menu to Google Sheets
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚡ Battle Station')
    .addItem('🔧 Setup Battle Station', 'setupBattleStation')
    .addItem('🔧 Build List', 'buildListWithGmailAndNotes')
    .addItem('🔄 Sync monday.com Data', 'syncMondayComBoards')
    .addItem('🔍 Check Duplicate Vendors', 'checkDuplicateVendors')
    .addSeparator()
    .addItem('⏭️ Skip Unchanged', 'skipToNextChanged')
    .addItem('🔄 Skip 5 & Return (Start/Continue)', 'skip5AndReturn')
    .addItem('↩️ Return to Origin (Skip 5)', 'continueSkip5AndReturn')
    .addItem('❌ Cancel Skip 5 Session', 'cancelSkip5Session')
    .addItem('🔁 Auto-Traverse All', 'autoTraverseVendors')
    .addItem('▶ Next Vendor', 'battleStationNext')
    .addItem('◀ Previous Vendor', 'battleStationPrevious')
    .addSeparator()
    .addItem('⚡ Quick Refresh (Email Only)', 'battleStationQuickRefresh')
    .addItem('🔄 Refresh', 'battleStationRefresh')
    .addItem('🔄 Hard Refresh (Clear Cache)', 'battleStationHardRefresh')
    .addSeparator()
    .addItem('💾 Update monday.com Notes', 'battleStationUpdateMondayNotes')
    .addItem('✓ Mark as Reviewed', 'battleStationMarkReviewed')
    .addItem('📧 Open Gmail Search', 'battleStationOpenGmail')
    .addItem('✉️ Email Contacts', 'battleStationEmailContacts')
    .addItem('🤖 Analyze Emails (Claude)', 'battleStationAnalyzeEmails')
    .addSeparator()
    .addItem('🔍 Go to Specific Vendor...', 'battleStationGoTo')
    .addToUi();
}

/**
 * Helper function: Get current vendor index from the display row
 */
function getCurrentVendorIndex_() {
  const ss = SpreadsheetApp.getActive();
  const bsSh = ss.getSheetByName(BS_CFG.BATTLE_SHEET);

  if (!bsSh) return null;

  // Navigation bar is in row 2, format: "◀ PREV  |  X of Y  |  NEXT ▶"
  const cellValue = String(bsSh.getRange(2, 1).getValue() || '');
  const match = cellValue.match(/\|\s*(\d+)\s*of\s*\d+\s*\|/);

  if (!match) {
    Logger.log(`Could not parse index from navigation: "${cellValue}"`);
    return null;
  }

  return parseInt(match[1]);
}

/**
 * Create or reset the Battle Station sheet
 */
function setupBattleStation() {
  const ss = SpreadsheetApp.getActive();
  let bsSh = ss.getSheetByName(BS_CFG.BATTLE_SHEET);

  if (!bsSh) {
    bsSh = ss.insertSheet(BS_CFG.BATTLE_SHEET);
  } else {
    bsSh.clear();
    bsSh.clearConditionalFormatRules();
  }

  // Narrower column widths for better screen fit at higher zoom
  // Left side columns (1-4)
  bsSh.setColumnWidth(1, 130);  // Labels (was 200)
  bsSh.setColumnWidth(2, 180);  // Values (was 250)
  bsSh.setColumnWidth(3, 90);   // Secondary labels (was 150)
  bsSh.setColumnWidth(4, 180);  // Secondary values (was 300)

  // Divider column (5) - thin black separator
  bsSh.setColumnWidth(5, 3);

  // Right side columns (6-9) for Contracts, Links, Documents
  bsSh.setColumnWidth(6, 130);  // Name (was 180)
  bsSh.setColumnWidth(7, 80);   // Type (was 120)
  bsSh.setColumnWidth(8, 90);   // Status (was 150)
  bsSh.setColumnWidth(9, 180);  // Notes/Folder (was 300)

  loadVendorData(1);

  SpreadsheetApp.getUi().alert('Battle Station initialized!\n\nUse the ⚡ Battle Station menu to navigate:\n- ▶ Next Vendor\n- ◀ Previous Vendor\n- 💾 Update monday.com Notes\n- ✓ Mark as Reviewed\n- ✉️ Email Contacts\n- 🤖 Analyze Emails (Claude)');
}

// NOTE: The rest of this file is very large (4000+ lines)
// It contains the full implementation of:
// - loadVendorData() - Main function to display vendor dashboard
// - getHelpfulLinksForVendor_() - Fetch helpful links from monday.com
// - getVendorContacts_() - Fetch contacts, notes, status from monday.com
// - getEmailsForVendor_() - Search Gmail for vendor emails
// - searchGmailFromLink_() - Helper for Gmail search
// - getTasksForVendor_() - Fetch tasks from monday.com
// - getUpcomingMeetingsForVendor_() - Search Google Calendar
// - Airtable contracts functions
// - Google Drive folder functions
// - Navigation functions (next, previous, refresh, etc.)
// - monday.com sync functions
// - Duplicate vendor checking
// - Checksum and change detection functions
// - Skip and traverse functions
// - And many more helper functions
//
// The full code was provided and should be pulled from Google Apps Script
// using: clasp pull
