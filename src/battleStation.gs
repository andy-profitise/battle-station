/************************************************************
 * A(I)DEN - One-by-one vendor review dashboard
 *
 * Last Updated: 2025-12-30 08:48PM PST
 *
 * Features:
 * - Navigate through vendors sequentially via menu
 * - View vendor details, notes, status, contacts
 * - See all related emails and monday.com tasks (live search)
 * - View helpful links from monday.com
 * - Update monday.com notes directly
 * - Mark vendors as reviewed/complete
 * - Email contacts directly from A(I)DEN
 * - Analyze emails with Claude AI (inline links)
 * - Snooze vendors until a specific date (skipped unless checksum changes)
 *
 * UPDATED: Added vendor snooze feature with date display in banner
 * FIXED: Preonboarding tasks warning, past meetings checksum, email diff logging
 ************************************************************/

const BS_CFG = {
  // Sheet names
  LIST_SHEET: 'List',
  BATTLE_SHEET: 'A(I)DEN',
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
  
  // A(I)DEN layout
  HEADER_ROWS: 3,
  DATA_START_ROW: 5,
  
  // Modern Color Palette - sleeker, more professional look
  COLOR_HEADER: '#1a73e8',        // Google Blue - main header
  COLOR_SUBHEADER: '#e8f0fe',     // Light blue - section headers
  COLOR_EMAIL: '#fef7e0',         // Warm cream - email section
  COLOR_TASK: '#e6f4ea',          // Fresh mint - tasks section
  COLOR_LINKS: '#f3e8fd',         // Soft lavender - helpful links
  COLOR_BUTTON: '#e8f0fe',        // Light blue - buttons
  COLOR_WARNING: '#fce8e6',       // Soft coral - warnings
  COLOR_SUCCESS: '#ceead6',       // Success green
  COLOR_SNOOZED: '#e1f5fe',       // Ice blue - snoozed emails
  COLOR_WAITING: '#fff8e1',       // Warm yellow - waiting
  COLOR_MISSING: '#fff3e0',       // Light amber - missing data
  COLOR_PHONEXA: '#ffe0b2',       // Peach - Phonexa waiting
  COLOR_OVERDUE: '#ffcdd2',       // Light red - overdue

  // Row highlight colors for skip/traverse
  COLOR_ROW_CHANGED: '#c8e6c9',   // Green - vendor has changes
  COLOR_ROW_SKIPPED: '#fff9c4',   // Yellow - vendor unchanged

  // Section styling
  COLOR_SECTION_BG: '#fafafa',    // Light gray for section backgrounds
  COLOR_TABLE_HEADER: '#f5f5f5',  // Table header background
  COLOR_TABLE_ALT: '#fafafa',     // Alternating row color
  COLOR_BORDER: '#e0e0e0',        // Border color
  COLOR_TEXT_MUTED: '#757575',    // Muted text
  COLOR_TEXT_LINK: '#1a73e8',     // Link color
  
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
  
  // Contracts 2026
  AIRTABLE_CONTRACTS_TABLE_2026: 'Contracts 2026',
  AIRTABLE_CONTRACTS_TABLE_ID_2026: 'tblYszexANBGGnyki',
  AIRTABLE_CONTRACTS_VIEW_ID_2026: 'viwfL5bDrEAvlmL7f',

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
  MAX_NOTES_LENGTH: 400,

  // Canned Response Attachments (Google Drive file IDs)
  REFERRAL_CONTRACT_FILE_ID: '1ON0EZmKvDyvwUDYBFyVkmj8TzkCBl2UX',
  INITIAL_CALL_FOLLOWUP_FILE_ID: '1T4N0Mia9icYrKH1B4rawu9qRBiqTgfrL',

  // Canned Response Templates - Google Doc IDs
  // Use <CONTACT_NAME> and <VENDOR_NAME> as placeholders in the docs
  // To get the ID: https://docs.google.com/document/d/[THIS_IS_THE_ID]/edit
  CANNED_RESPONSE_DOCS: {
    REFERRAL_PROGRAM: '1tn3uQMvVR6ZItk1p0k-0-BclMJxbNBDzuh4jSt8PT9c',
    INITIAL_CALL_FOLLOWUP: '1z2lUA4aDD1_zSq1gxX13pcrqlHMh80CmUvorfwLw5ws'
  }
};

/**
 * Add menu to Google Sheets
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // Main A(I)DEN menu - setup, sync, actions
  ui.createMenu('⚡ A(I)DEN')
    .addItem('🔧 Setup A(I)DEN', 'setupBattleStation')
    .addItem('🔧 Build List', 'buildListWithGmailAndNotes')
    .addItem('🔄 Sync monday.com Data', 'syncMondayComBoards')
    .addItem('🔍 Check Duplicate Vendors', 'checkDuplicateVendors')
    .addSeparator()
    .addItem('💾 Update monday.com Notes', 'battleStationUpdateMondayNotes')
    .addItem('✓ Mark as Reviewed', 'battleStationMarkReviewed')
    .addItem('⚑ Flag/Unflag Vendor', 'battleStationToggleFlag')
    .addItem('💤 Snooze Vendor...', 'battleStationSnoozeVendor')
    .addItem('📧 Open Gmail Search', 'battleStationOpenGmail')
    .addItem('✉️ Email Contacts', 'battleStationEmailContacts')
    .addItem('📇 Discover Contacts from Gmail', 'discoverContactsFromGmail')
    .addItem('🤖 Analyze Emails (Claude)', 'battleStationAnalyzeEmails')
    .addItem('🤖 Analyze Tasks (Claude)', 'analyzeTasksFromEmails')
    .addItem('❓ Ask About Vendor (Claude)', 'askAboutVendor')
    .addToUi();

  // Refresh menu - refresh current vendor view
  ui.createMenu('🔄 Refresh')
    .addItem('⚡ Quick Refresh (Email Only)', 'battleStationQuickRefresh')
    .addItem('🔁 Quick Refresh Until Changed', 'battleStationQuickRefreshUntilChanged')
    .addItem('🔄 Hard Refresh (Clear Cache)', 'battleStationHardRefresh')
    .addSeparator()
    .addItem('🗑️ Reset Module Checksums (Fix False Positives)', 'resetAllModuleChecksums')
    .addToUi();

  // Navigation menu - movement and traversal
  ui.createMenu('🧭 Navigation')
    .addItem('⏭️ Skip Unchanged', 'skipToNextChanged')
    .addItem('🔁 Auto-Traverse All', 'autoTraverseVendors')
    .addSeparator()
    .addItem('🔄 Skip 5 & Return (Start/Continue)', 'skip5AndReturn')
    .addItem('↩️ Return to Origin (Skip 5)', 'continueSkip5AndReturn')
    .addItem('❌ Cancel Skip 5 Session', 'cancelSkip5Session')
    .addSeparator()
    .addItem('▶ Next Vendor', 'battleStationNext')
    .addItem('◀ Previous Vendor', 'battleStationPrevious')
    .addItem('🔍 Go to Specific Vendor...', 'battleStationGoTo')
    .addSeparator()
    .addItem('⚑ Flag/Unflag Vendor', 'battleStationToggleFlag')
    .addItem('💤 Snooze Vendor...', 'battleStationSnoozeVendor')
    .addToUi();

  // Email Response Templates menu
  ui.createMenu('📧 Email Responses')
    .addItem('🔄 Cold Outreach - Follow Up', 'emailResponseColdFollowUp')
    .addItem('📅 Schedule a Call', 'emailResponseScheduleCall')
    .addItem('💰 Payment/Invoice Follow Up', 'emailResponsePaymentFollowUp')
    .addItem('📋 General Follow Up', 'emailResponseGeneralFollowUp')
    .addItem('🚫 Missed Meeting', 'emailResponseMissedMeeting')
    .addItem('✍️ Custom Response...', 'emailResponseCustom')
    .addSeparator()
    .addItem('📨 Referral Program - Canned', 'cannedResponseReferralProgram')
    .addItem('📞 Initial Call Follow-up - Canned', 'cannedResponseInitialCallFollowup')
    .addToUi();
}

/************************************************************
 * STYLING HELPER FUNCTIONS
 * Reduce repetitive styling code and ensure visual consistency
 ************************************************************/

/**
 * Apply section header styling (main sections like VENDOR INFO, EMAILS)
 * @param {Range} range - The range to style
 * @param {string} text - Header text
 * @param {string} [bgColor] - Optional background color (defaults to SUBHEADER)
 */
function styleHeader_(range, text, bgColor) {
  range.setValue(text)
    .setBackground(bgColor || BS_CFG.COLOR_SUBHEADER)
    .setFontWeight('bold')
    .setFontSize(11)
    .setFontColor('#1a73e8')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  return range;
}

/**
 * Apply sub-section header styling (smaller headers within sections)
 * @param {Range} range - The range to style
 * @param {string} text - Header text
 */
function styleSubHeader_(range, text) {
  range.setValue(text)
    .setBackground(BS_CFG.COLOR_SECTION_BG)
    .setFontWeight('bold')
    .setFontSize(10)
    .setFontColor(BS_CFG.COLOR_TEXT_LINK)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  return range;
}

/**
 * Apply table header styling (column headers in tables)
 * @param {Range} range - The range to style
 * @param {string} text - Header text
 */
function styleTableHeader_(range, text) {
  range.setValue(text)
    .setFontWeight('bold')
    .setFontSize(9)
    .setBackground(BS_CFG.COLOR_TABLE_HEADER)
    .setHorizontalAlignment('left');
  return range;
}

/**
 * Apply link cell styling
 * @param {Range} range - The range to style
 * @param {string} url - URL for the hyperlink
 * @param {string} displayText - Text to display
 */
function styleLink_(range, url, displayText) {
  range.setFormula(`=HYPERLINK("${url}", "${displayText.replace(/"/g, '""')}")`)
    .setFontColor(BS_CFG.COLOR_TEXT_LINK);
  return range;
}

/**
 * Apply empty/no data styling
 * @param {Range} range - The range to style
 * @param {string} text - Text to display (e.g., "No data found")
 */
function styleEmpty_(range, text) {
  range.setValue(text)
    .setFontStyle('italic')
    .setFontColor(BS_CFG.COLOR_TEXT_MUTED)
    .setBackground(BS_CFG.COLOR_SECTION_BG)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  return range;
}

/**
 * Apply warning/missing data styling
 * @param {Range} range - The range to style
 * @param {string} text - Warning text
 * @param {string} [linkUrl] - Optional link to fix the issue
 */
function styleWarning_(range, text, linkUrl) {
  if (linkUrl) {
    range.setFormula(`=HYPERLINK("${linkUrl}", "${text}")`)
      .setBackground(BS_CFG.COLOR_MISSING)
      .setFontColor(BS_CFG.COLOR_TEXT_LINK);
  } else {
    range.setValue(text)
      .setBackground(BS_CFG.COLOR_WARNING)
      .setFontColor('#c62828');
  }
  return range;
}

/**
 * Apply label styling (left column labels like "Vendor:", "Status:")
 * @param {Range} range - The range to style
 * @param {string} text - Label text
 */
function styleLabel_(range, text) {
  range.setValue(text)
    .setFontWeight('bold')
    .setFontColor('#424242');
  return range;
}

/**
 * Set column divider styling (thin black separator)
 * @param {Sheet} sheet - The sheet to style
 * @param {number} col - Column number for the divider
 * @param {number} startRow - Starting row
 * @param {number} numRows - Number of rows
 */
function styleColumnDivider_(sheet, col, startRow, numRows) {
  sheet.getRange(startRow, col, numRows, 1)
    .setBackground('#424242');
}

/**
 * Batch set multiple cell values and styles efficiently
 * @param {Sheet} sheet - The sheet
 * @param {Array} cells - Array of {row, col, value, styles} objects
 *   styles can include: bg, fontWeight, fontSize, fontColor, align, wrap
 */
function batchStyleCells_(sheet, cells) {
  for (const cell of cells) {
    const range = sheet.getRange(cell.row, cell.col);

    if (cell.value !== undefined) {
      if (cell.formula) {
        range.setFormula(cell.value);
      } else {
        range.setValue(cell.value);
      }
    }

    const s = cell.styles || {};
    if (s.bg) range.setBackground(s.bg);
    if (s.fontWeight) range.setFontWeight(s.fontWeight);
    if (s.fontSize) range.setFontSize(s.fontSize);
    if (s.fontColor) range.setFontColor(s.fontColor);
    if (s.align) range.setHorizontalAlignment(s.align);
    if (s.vAlign) range.setVerticalAlignment(s.vAlign);
    if (s.wrap) range.setWrap(s.wrap);
    if (s.fontStyle) range.setFontStyle(s.fontStyle);
    if (s.numberFormat) range.setNumberFormat(s.numberFormat);
  }
}

/**
 * Helper function: Get current vendor index from the display row
 */
function getCurrentVendorIndex_() {
  const ss = SpreadsheetApp.getActive();
  const bsSh = ss.getSheetByName(BS_CFG.BATTLE_SHEET);

  if (!bsSh) return null;

  // Navigation bar is in row 3, format: "◀  X / Y  ▶"
  const cellValue = String(bsSh.getRange(3, 1).getValue() || '');
  const match = cellValue.match(/(\d+)\s*\/\s*\d+/);

  if (!match) {
    Logger.log(`Could not parse index from navigation: "${cellValue}"`);
    return null;
  }

  return parseInt(match[1]);
}

/**
 * Create or reset the A(I)DEN sheet
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
  
  SpreadsheetApp.getUi().alert('A(I)DEN initialized!\n\nUse the ⚡ A(I)DEN menu to navigate:\n- ▶ Next Vendor\n- ◀ Previous Vendor\n- 💾 Update monday.com Notes\n- ✓ Mark as Reviewed\n- ✉️ Email Contacts\n- 🤖 Analyze Emails (Claude)');
}

/**
 * Load and display data for a specific vendor by index
 */
function loadVendorData(vendorIndex, options) {
  // Default options
  options = options || {};
  const useCache = options.useCache !== undefined ? options.useCache : false;
  const forceChanged = options.forceChanged || false;  // If true, skip the ✅ indicator (used when skipToNextChanged detected a change)
  const changeType = options.changeType || null;  // The type of change detected (e.g., 'overdue emails')
  
  const ss = SpreadsheetApp.getActive();
  const bsSh = ss.getSheetByName(BS_CFG.BATTLE_SHEET);
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
  
  if (!bsSh || !listSh) {
    throw new Error('Required sheets not found');
  }
  
  const totalVendors = listSh.getLastRow() - 1;
  
  if (vendorIndex < 1) vendorIndex = 1;
  if (vendorIndex > totalVendors) vendorIndex = totalVendors;
  
  const listRow = vendorIndex + 1;
  const vendorData = listSh.getRange(listRow, 1, 1, 8).getValues()[0];
  
  const vendor = vendorData[BS_CFG.L_VENDOR] || '';
  const ttlUsd = vendorData[BS_CFG.L_TTL_USD] || 0;
  const source = vendorData[BS_CFG.L_SOURCE] || '';
  const status = vendorData[BS_CFG.L_STATUS] || '';
  const notes = vendorData[BS_CFG.L_NOTES] || '';
  const processed = vendorData[BS_CFG.L_PROCESSED] || false;
  
  const mondayBoardId = source.toLowerCase().includes('buyer') ? `${BS_CFG.BUYERS_BOARD_ID} (Buyers)` : 
                        source.toLowerCase().includes('affiliate') ? `${BS_CFG.AFFILIATES_BOARD_ID} (Affiliates)` : 
                        `${BS_CFG.BUYERS_BOARD_ID} (Buyers - default)`;
  
  const processedDisplay = processed ? '✅ Yes (Reviewed)' : '⚠️ No (Needs Review)';
  
  // Clear entire sheet (9 columns now - includes divider)
  const lastRow = bsSh.getMaxRows();
  if (lastRow > 0) {
    bsSh.getRange(1, 1, lastRow, 9).clearContent().clearFormat().clearDataValidations();
  }
  
  let currentRow = 1;

  // Title - full width, modern blue header with subtle shadow effect
  bsSh.getRange(currentRow, 1, 1, 9).merge()
    .setValue(`⚡ A(I)DEN`)
    .setFontSize(16).setFontWeight('bold')
    .setBackground(BS_CFG.COLOR_HEADER)
    .setFontColor('white')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  bsSh.setRowHeight(currentRow, 40);
  currentRow++;

  // Vendor name banner - prominent display (with flag/snooze indicators)
  let vendorDisplay = vendor;
  if (isVendorFlagged_(vendor)) {
    vendorDisplay += ' ⚑';
  }
  const snoozeDate = getVendorSnoozeDate_(vendor);
  if (snoozeDate && snoozeDate > new Date()) {
    const dateStr = Utilities.formatDate(snoozeDate, Session.getScriptTimeZone(), 'M/d');
    vendorDisplay += ` 💤${dateStr}`;
  }
  bsSh.getRange(currentRow, 1, 1, 9).merge()
    .setValue(vendorDisplay)
    .setFontSize(13).setFontWeight('bold')
    .setBackground('#e3f2fd')
    .setFontColor('#1565c0')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  bsSh.setRowHeight(currentRow, 32);
  currentRow++;

  // Navigation bar - cleaner, more modern
  const navText = `◀  ${vendorIndex} / ${totalVendors}  ▶`;
  bsSh.getRange(currentRow, 1, 1, 9).merge()
    .setValue(navText)
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground('#fafafa')
    .setFontColor(BS_CFG.COLOR_TEXT_MUTED);
  bsSh.setRowHeight(currentRow, 22);
  currentRow++;

  bsSh.setFrozenRows(currentRow - 1);

  // Spacer row
  bsSh.setRowHeight(currentRow, 6);
  currentRow++;

  // VENDOR INFO SECTION - using helper
  bsSh.getRange(currentRow, 1, 1, 4).merge();
  styleHeader_(bsSh.getRange(currentRow, 1), `📊 VENDOR INFO`)
    .setFontSize(11)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  bsSh.setRowHeight(currentRow, 26);
  
  // Track right column starting row - starts at same row as VENDOR header
  const rightColumnStartRow = currentRow;
  let rightColumnRow = rightColumnStartRow;
  
  currentRow++;
  
  // Get contacts and notes from monday.com
  ss.toast('Loading vendor details...', '📊 Loading', 2);
  const contactData = getVendorContacts_(vendor, listRow);
  const mondayNotes = contactData.notes || notes;
  const contacts = contactData.contacts;
  
  // Vendor details - build Phonexa link display
  let phonexaDisplay = '';
  let phonexaFormula = '';
  let phonexaMissing = false;
  
  if (contactData.phonexaLink) {
    phonexaDisplay = contactData.phonexaLink;
    phonexaFormula = `=HYPERLINK("${contactData.phonexaLink}", "Open in Phonexa")`;
  } else {
    // Generate monday.com link filtered by vendor name
    const encodedVendor = encodeURIComponent(vendor);
    const mondayFilterLink = `https://profitise-company.monday.com/boards/${contactData.boardId || (source.toLowerCase().includes('affiliate') ? BS_CFG.AFFILIATES_BOARD_ID : BS_CFG.BUYERS_BOARD_ID)}?term=${encodedVendor}`;
    phonexaDisplay = '(not set)';
    phonexaFormula = `=HYPERLINK("${mondayFilterLink}", "⚠️ Add in monday.com →")`;
    phonexaMissing = true;
  }
  
  // Use live status from monday.com, fallback to List sheet
  const liveStatus = contactData.liveStatus || vendorData[BS_CFG.L_STATUS] || '';
  
  // Build Status link to appropriate board
  const vendorBoardId = source.toLowerCase().includes('affiliate') ? BS_CFG.AFFILIATES_BOARD_ID : BS_CFG.BUYERS_BOARD_ID;
  const encodedVendorForStatus = encodeURIComponent(vendor);
  const statusLink = `https://profitise-company.monday.com/boards/${vendorBoardId}?term=${encodedVendorForStatus}`;
  const statusFormula = `=HYPERLINK("${statusLink}", "${liveStatus}${status === 'Dead' ? ' ⚠️' : ' ✅'}")`;
  
  // VENDOR INFO - 2 COLUMN LAYOUT
  // Left column: Label (col 1) + Value (col 2)
  // Right column: Label (col 3) + Value (col 4)
  
  // Row 1: Vendor | Status
  bsSh.getRange(currentRow, 1).setValue('Vendor:').setFontWeight('bold');
  bsSh.getRange(currentRow, 2).setValue(vendor);
  bsSh.getRange(currentRow, 3).setValue('Status:').setFontWeight('bold');
  const statusCell = bsSh.getRange(currentRow, 4);
  statusCell.setFormula(statusFormula).setFontColor('#1a73e8');
  if (status === 'Dead') {
    statusCell.setBackground(BS_CFG.COLOR_WARNING);
  }
  currentRow++;
  
  // Row 2: Source | Total USD
  bsSh.getRange(currentRow, 1).setValue('Source:').setFontWeight('bold');
  bsSh.getRange(currentRow, 2).setValue(source);
  bsSh.getRange(currentRow, 3).setValue('Total USD:').setFontWeight('bold');
  bsSh.getRange(currentRow, 4).setValue(`$${Number(ttlUsd).toLocaleString()}`).setHorizontalAlignment('left');
  currentRow++;
  
  // Row 3: Live Verticals | Live Modalities
  bsSh.getRange(currentRow, 1).setValue('Live Verticals:').setFontWeight('bold');
  const liveVertCell = bsSh.getRange(currentRow, 2).setValue(contactData.liveVerticals || '(none)');
  if (!contactData.liveVerticals) liveVertCell.setBackground(BS_CFG.COLOR_MISSING);
  bsSh.getRange(currentRow, 3).setValue('Live Modalities:').setFontWeight('bold');
  const liveModCell = bsSh.getRange(currentRow, 4).setValue(contactData.liveModalities || '(none)');
  if (!contactData.liveModalities) liveModCell.setBackground(BS_CFG.COLOR_MISSING);
  currentRow++;
  
  // Row 4: Other Verticals | Phonexa Link
  bsSh.getRange(currentRow, 1).setValue('Other Verticals:').setFontWeight('bold');
  bsSh.getRange(currentRow, 2).setValue(contactData.otherVerticals || '(none)');
  bsSh.getRange(currentRow, 3).setValue('Phonexa Link:').setFontWeight('bold');
  const phonexaCell = bsSh.getRange(currentRow, 4);
  phonexaCell.setFormula(phonexaFormula).setFontColor('#1a73e8');
  if (phonexaMissing) phonexaCell.setBackground(BS_CFG.COLOR_MISSING);
  currentRow++;
  
  // Row 5: State(s) - full width (all 4 columns)
  bsSh.getRange(currentRow, 1).setValue('State(s):').setFontWeight('bold');
  const statesCell = bsSh.getRange(currentRow, 2, 1, 3).merge();
  if (contactData.states) {
    statesCell.setValue(contactData.states);
  } else if (!source.toLowerCase().includes('affiliate')) {
    // Only show as missing for Buyers (Affiliates don't have states)
    const encodedVendor = encodeURIComponent(vendor);
    const mondayStatesLink = `https://profitise-company.monday.com/boards/${contactData.boardId || BS_CFG.BUYERS_BOARD_ID}?term=${encodedVendor}`;
    statesCell.setFormula(`=HYPERLINK("${mondayStatesLink}", "⚠️ Add in monday.com")`);
    statesCell.setBackground(BS_CFG.COLOR_WARNING).setFontColor('#1a73e8');
  } else {
    statesCell.setValue('N/A');
  }
  currentRow++;
  
  // Row 6: Dead State(s) - full width (Buyers only) - strikethrough to indicate dead
  bsSh.getRange(currentRow, 1).setValue('Dead State(s):').setFontWeight('bold').setFontLine('line-through');
  const deadStatesCell = bsSh.getRange(currentRow, 2, 1, 3).merge().setFontLine('line-through');
  if (contactData.deadStates) {
    deadStatesCell.setValue(contactData.deadStates);
  } else if (!source.toLowerCase().includes('affiliate')) {
    deadStatesCell.setValue('(none)').setFontStyle('italic').setFontColor('#999999');
  } else {
    deadStatesCell.setValue('N/A');
  }
  currentRow++;
  
  // Row 7: Last Updated | Processed
  bsSh.getRange(currentRow, 1).setValue('Last Updated:').setFontWeight('bold');
  
  // lastUpdated now comes formatted from API as "Dec 3, 2025 10:25 PM"
  const lastUpdDisplay = contactData.lastUpdated || '(not available)';
  
  const lastUpdCell = bsSh.getRange(currentRow, 2).setValue(lastUpdDisplay).setHorizontalAlignment('left');
  if (!contactData.lastUpdated) lastUpdCell.setBackground(BS_CFG.COLOR_MISSING);
  bsSh.getRange(currentRow, 3).setValue('Processed:').setFontWeight('bold');
  const processedCell = bsSh.getRange(currentRow, 4).setValue(processedDisplay);
  if (processed) {
    processedCell.setBackground(BS_CFG.COLOR_SUCCESS);
  } else {
    processedCell.setBackground('#fff4e5');
  }
  currentRow++;
  
  currentRow++;
  
  // ========== LEFT SIDE (Columns 1-4) ==========
  
  // Track row where Contacts starts (for Helpful Links alignment)
  let helpfulLinksStartRow = currentRow;
  
  // CONTACTS SECTION (moved above calendar)
  if (contacts.length > 0) {
    // Define contact type priority order
    const contactTypePriority = {
      'Primary': 1,
      'Technical': 2,
      'Contracts': 3,
      'Accounting': 4,
      'Management': 5,
      'Other/Unknown': 6
    };

    // Sort contacts by: Status (Active first) -> Type priority -> Name
    contacts.sort((a, b) => {
      // First: Sort by status (Active before Not Active)
      const aActive = (a.status && a.status.toLowerCase() !== 'not active') ? 0 : 1;
      const bActive = (b.status && b.status.toLowerCase() !== 'not active') ? 0 : 1;
      
      if (aActive !== bActive) {
        return aActive - bActive;
      }
      
      // Second: Sort by contact type priority
      const aPriority = contactTypePriority[a.contactType] || 999;
      const bPriority = contactTypePriority[b.contactType] || 999;
      
      if (aPriority !== bPriority) {
        return aPriority - bPriority;
      }
      
      // Third: Sort alphabetically by name
      return a.name.localeCompare(b.name);
    });
    
    // Get last Contact Discovery run date for this vendor
    const contactDiscoveryProps = PropertiesService.getScriptProperties();
    const contactDiscoveryKey = `BS_CONTACT_DISCOVERY_${vendor.replace(/[^a-zA-Z0-9]/g, '_')}`;
    const lastDiscoveryDate = contactDiscoveryProps.getProperty(contactDiscoveryKey);
    const discoveryDisplay = lastDiscoveryDate ? `  📅 Last scanned: ${lastDiscoveryDate}` : '';

    bsSh.getRange(currentRow, 1, 1, 4).merge()
      .setValue(`👤 CONTACTS (${contacts.length})${discoveryDisplay}`)
      .setBackground('#f8f9fa')
      .setFontWeight('bold')
      .setFontSize(10)
      .setFontColor('#1a73e8')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle');
    bsSh.setRowHeight(currentRow, 24);
    currentRow++;
    
    // Header row for contacts
    bsSh.getRange(currentRow, 1).setValue('Name').setFontWeight('bold').setBackground('#f3f3f3').setFontSize(9);
    bsSh.getRange(currentRow, 2).setValue('Email / Phone').setFontWeight('bold').setBackground('#f3f3f3').setFontSize(9);
    bsSh.getRange(currentRow, 3).setValue('Status').setFontWeight('bold').setBackground('#f3f3f3').setFontSize(9);
    bsSh.getRange(currentRow, 4).setValue('Type').setFontWeight('bold').setBackground('#f3f3f3').setFontSize(9);
    currentRow++;
    
    for (const contact of contacts) {
      // Name column - ALWAYS clickable and links to Contacts board
      const encodedContact = encodeURIComponent(contact.name);
      const contactsFilterLink = `https://profitise-company.monday.com/boards/${BS_CFG.CONTACTS_BOARD_ID}?term=${encodedContact}`;
      const nameCell = bsSh.getRange(currentRow, 1)
        .setFormula(`=HYPERLINK("${contactsFilterLink}", "${contact.name}")`)
        .setFontSize(10)
        .setBackground('#f0f8ff')
        .setFontColor('#1a73e8');
      
      // Email / Phone column - highlight if both missing
      // Format phone: normalize to (XXX) XXX-XXXX
      let phone = contact.phone || '';
      if (phone) {
        // Strip all non-digits
        let digits = phone.replace(/\D/g, '');
        // Remove leading "1" if 11 digits (country code)
        if (digits.length === 11 && digits.startsWith('1')) {
          digits = digits.substring(1);
        }
        // Format as (XXX) XXX-XXXX if we have 10 digits
        if (digits.length === 10) {
          phone = digits.replace(/(\d{3})(\d{3})(\d{4})/, '($1) $2-$3');
        }
      }
      
      const emailPhone = [contact.email, phone].filter(x => x).join(' / ');
      const emailPhoneCell = bsSh.getRange(currentRow, 2).setValue(emailPhone || '(missing)').setFontSize(10);
      
      if (!emailPhone) {
        emailPhoneCell.setFormula(`=HYPERLINK("${contactsFilterLink}", "⚠️ Add in monday.com")`);
        emailPhoneCell.setBackground(BS_CFG.COLOR_MISSING).setFontColor('#1a73e8');
      } else {
        emailPhoneCell.setBackground('#f0f8ff');
      }
      
      // Status column - highlight if missing
      const statusCell = bsSh.getRange(currentRow, 3).setValue(contact.status || '(missing)').setFontSize(10);
      if (!contact.status) {
        statusCell.setFormula(`=HYPERLINK("${contactsFilterLink}", "⚠️ Add")`);
        statusCell.setBackground(BS_CFG.COLOR_MISSING).setFontColor('#1a73e8');
      } else {
        statusCell.setBackground('#f0f8ff');
      }
      
      // Contact Type column - highlight if missing
      const typeCell = bsSh.getRange(currentRow, 4).setValue(contact.contactType || '(missing)').setFontSize(10);
      if (!contact.contactType) {
        typeCell.setFormula(`=HYPERLINK("${contactsFilterLink}", "⚠️ Add")`);
        typeCell.setBackground(BS_CFG.COLOR_MISSING).setFontColor('#1a73e8');
      } else {
        typeCell.setBackground('#f0f8ff');
      }
      
      // If Not Active, strikethrough the entire row
      if (contact.status && contact.status.toLowerCase() === 'not active') {
        bsSh.getRange(currentRow, 1, 1, 4)
          .setFontLine('line-through')
          .setFontColor('#999999');
      }
      
      currentRow++;
    }
    currentRow++;
  }
  
  // Track row where Upcoming Meetings starts (for Box Documents alignment)
  const upcomingMeetingsStartRow = currentRow;
  
  // CALENDAR MEETINGS SECTION
  ss.toast('Checking calendar...', '📅 Loading', 2);
  // Extract contact emails to search for in calendar events
  const contactEmails = (contacts || []).map(c => c.email).filter(e => e && e.includes('@'));
  const meetingsResult = getUpcomingMeetingsForVendor_(vendor, contactEmails);
  const meetings = meetingsResult.meetings || [];
  const totalMeetingCount = meetingsResult.totalCount || 0;
  
  bsSh.getRange(currentRow, 1, 1, 4).merge()
    .setValue(`📅 UPCOMING MEETINGS (${meetings.length})`)
    .setBackground('#f8f9fa')
    .setFontWeight('bold')
    .setFontSize(10)
    .setFontColor('#1a73e8')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  bsSh.setRowHeight(currentRow, 24);
  currentRow++;
  
  if (meetings.length === 0) {
    bsSh.getRange(currentRow, 1, 1, 4).merge()
      .setValue('No upcoming meetings found')
      .setFontStyle('italic')
      .setBackground('#fafafa')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle');
    bsSh.setRowHeight(currentRow, 25);
    currentRow++;
  } else {
    // Meeting headers
    bsSh.getRange(currentRow, 1).setValue('Event').setFontWeight('bold').setBackground('#f3f3f3').setHorizontalAlignment('left');
    bsSh.getRange(currentRow, 2).setValue('Date').setFontWeight('bold').setBackground('#f3f3f3').setHorizontalAlignment('left');
    bsSh.getRange(currentRow, 3).setValue('Time').setFontWeight('bold').setBackground('#f3f3f3').setHorizontalAlignment('left');
    bsSh.getRange(currentRow, 4).setValue('Status').setFontWeight('bold').setBackground('#f3f3f3').setHorizontalAlignment('left');
    currentRow++;
    
    for (const meeting of meetings.slice(0, 10)) {
      // Event title - clickable link
      if (meeting.link) {
        bsSh.getRange(currentRow, 1)
          .setFormula(`=HYPERLINK("${meeting.link}", "${meeting.title.replace(/"/g, '""')}")`)
          .setFontColor('#1a73e8');
      } else {
        bsSh.getRange(currentRow, 1).setValue(meeting.title);
      }
      
      bsSh.getRange(currentRow, 2).setValue(meeting.date).setNumberFormat('@').setHorizontalAlignment('left');
      bsSh.getRange(currentRow, 3).setValue(meeting.time).setHorizontalAlignment('left');
      bsSh.getRange(currentRow, 4).setValue(meeting.status).setHorizontalAlignment('left');
      
      // Color code by timing
      if (meeting.isToday) {
        bsSh.getRange(currentRow, 1, 1, 4).setBackground('#fff2cc'); // Yellow for today
      } else if (meeting.isPast) {
        bsSh.getRange(currentRow, 1, 1, 4).setBackground('#f3f3f3'); // Gray for past
      } else {
        bsSh.getRange(currentRow, 1, 1, 4).setBackground('#d9ead3'); // Green for upcoming
      }
      
      currentRow++;
    }
    
    // Show "more meetings" link with Google Calendar search
    if (totalMeetingCount > meetings.length || meetings.length > 10) {
      const moreCount = totalMeetingCount > meetings.length ? totalMeetingCount - meetings.length : meetings.length - 10;
      const calSearchUrl = `https://calendar.google.com/calendar/r/search?q=${encodeURIComponent(vendor)}`;
      bsSh.getRange(currentRow, 1, 1, 4).merge()
        .setFormula(`=HYPERLINK("${calSearchUrl}", "🔍 ${moreCount}+ more meetings - Search in Google Calendar")`)
        .setFontStyle('italic')
        .setFontColor('#1a73e8')
        .setHorizontalAlignment('left');
      currentRow++;
    }
  }
  
  currentRow++;
  
  // ========== RIGHT SIDE (Columns 5-8) ==========
  
  // HELPFUL LINKS SECTION (right side - aligned with VENDOR INFO at top)
  ss.toast('Loading helpful links...', '🔗 Loading', 2);
  const helpfulLinks = getHelpfulLinksForVendor_(vendor, listRow);

  // Generate L2M Reporting link if we have a Phonexa link
  const l2mLink = getL2MReportingLink_(contactData.phonexaLink, source);
  const totalLinksCount = helpfulLinks.length + 1; // +1 for L2M row (always shown)

  const helpfulLinksUrl = `https://profitise-company.monday.com/boards/${BS_CFG.HELPFUL_LINKS_BOARD_ID}`;
  bsSh.getRange(rightColumnRow, 6, 1, 4).merge()
    .setFormula(`=HYPERLINK("${helpfulLinksUrl}", "🔗 HELPFUL LINKS (${totalLinksCount})")`)
    .setBackground('#f8f9fa')
    .setFontWeight('bold')
    .setFontSize(10)
    .setFontColor('#1a73e8')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('top');
  bsSh.setRowHeight(rightColumnRow, 24);
  rightColumnRow++;

  // Header row for links table
  bsSh.getRange(rightColumnRow, 6).setValue('Description').setFontWeight('bold').setBackground('#f3f3f3').setHorizontalAlignment('left');
  bsSh.getRange(rightColumnRow, 7, 1, 3).merge().setValue('Link').setFontWeight('bold').setBackground('#f3f3f3').setHorizontalAlignment('left');
  rightColumnRow++;

  // L2M Reporting link first (light grey background to distinguish from monday.com links)
  const l2mBgColor = '#e8e8e8'; // Light grey
  if (l2mLink) {
    bsSh.getRange(rightColumnRow, 6).setValue(l2mLink.label).setWrap(true).setHorizontalAlignment('left').setVerticalAlignment('top').setBackground(l2mBgColor);
    bsSh.getRange(rightColumnRow, 7, 1, 3).merge()
      .setFormula(`=HYPERLINK("${l2mLink.url}", "https://cp.profitise.com/p2/report/...")`)
      .setFontColor('#1a73e8')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('top')
      .setBackground(l2mBgColor);
    rightColumnRow++;
  } else {
    // No Phonexa link - show warning with link to monday.com board
    const encodedVendor = encodeURIComponent(vendor);
    const vendorBoardId = source.toLowerCase().includes('affiliate') ? BS_CFG.AFFILIATES_BOARD_ID : BS_CFG.BUYERS_BOARD_ID;
    const mondayLink = `https://profitise-company.monday.com/boards/${vendorBoardId}?term=${encodedVendor}`;
    bsSh.getRange(rightColumnRow, 6).setValue('⚠️ NO PHONEXA LINK FOUND').setWrap(true).setHorizontalAlignment('left').setVerticalAlignment('top').setBackground(l2mBgColor).setFontColor('#b71c1c');
    bsSh.getRange(rightColumnRow, 7, 1, 3).merge()
      .setFormula(`=HYPERLINK("${mondayLink}", "Add in monday.com →")`)
      .setFontColor('#1a73e8')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('top')
      .setBackground(l2mBgColor);
    rightColumnRow++;
  }

  if (helpfulLinks.length === 0) {
    bsSh.getRange(rightColumnRow, 6, 1, 4).merge()
      .setValue('No helpful links found')
      .setFontStyle('italic')
      .setBackground('#fafafa')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('top');
    bsSh.setRowHeight(rightColumnRow, 25);
    rightColumnRow++;
  } else if (helpfulLinks.length > 0) {
    // monday.com links (no grey background - default white)
    for (const link of helpfulLinks.slice(0, 8)) {
      bsSh.getRange(rightColumnRow, 6).setValue(link.notes || '(no description)').setWrap(true).setHorizontalAlignment('left').setVerticalAlignment('top');

      if (link.url) {
        bsSh.getRange(rightColumnRow, 7, 1, 3).merge()
          .setFormula(`=HYPERLINK("${link.url}", "${link.url.substring(0, 50)}${link.url.length > 50 ? '...' : ''}")`)
          .setFontColor('#1a73e8')
          .setHorizontalAlignment('left')
          .setVerticalAlignment('top');
      } else {
        bsSh.getRange(rightColumnRow, 7, 1, 3).merge().setValue('(no URL)').setHorizontalAlignment('left').setVerticalAlignment('top');
      }

      rightColumnRow++;
    }

    if (helpfulLinks.length > 8) {
      bsSh.getRange(rightColumnRow, 6, 1, 4).merge()
        .setValue(`... and ${helpfulLinks.length - 8} more links`)
        .setFontStyle('italic')
        .setHorizontalAlignment('left');
      rightColumnRow++;
    }
  }
  
  rightColumnRow++;
  
  // CONTRACTS SECTION (right side - aligned with CONTACTS)
  ss.toast('Checking contracts...', '📋 Loading', 2);
  let contractsData = getVendorContracts_(vendor);
  let contractsMatchedOn = contractsData.hasContracts ? vendor : '';
  
  // If no contracts found, try Other Name(s)
  if (!contractsData.hasContracts && contactData.otherName) {
    // First try the full Other Name value (in case it's like "Profitise, LLC")
    Logger.log(`No contracts for "${vendor}", trying full Other Name: "${contactData.otherName}"`);
    const fullResult = getVendorContracts_(contactData.otherName);
    
    if (fullResult.hasContracts) {
      contractsData = fullResult;
      contractsMatchedOn = contactData.otherName;
    } else if (contactData.otherName.includes(',')) {
      // If full name found nothing, try splitting by comma for multiple values
      const otherNames = contactData.otherName.split(',').map(n => n.trim()).filter(n => n.length > 0);
      Logger.log(`Full name found nothing, trying individual values: ${otherNames.join(', ')}`);
      
      for (const altName of otherNames) {
        // Skip generic terms that cause false positives
        if (BS_CFG.SKIP_SEARCH_TERMS.some(term => term.toLowerCase() === altName.toLowerCase())) {
          Logger.log(`Skipping generic term: "${altName}"`);
          continue;
        }
        Logger.log(`Searching Airtable contracts for: "${altName}"`);
        const altResult = getVendorContracts_(altName);
        if (altResult.hasContracts) {
          contractsData = altResult;
          contractsMatchedOn = altName;
          break; // Found contracts, stop searching
        }
      }
    }
  }
  
  // Use helpfulLinksStartRow to align with Contacts section
  let contractsRow = helpfulLinksStartRow;
  
  const airtableContractsUrl = 'https://airtable.com/appc6xu9qLlOP5G5m/tblYszexANBGGnyki/viwfL5bDrEAvlmL7f?blocks=hide';
  // Escape quotes in matchedOn for use in formula
  const escapedMatchedOn = contractsMatchedOn ? contractsMatchedOn.replace(/"/g, '""') : '';
  const matchedDisplay = escapedMatchedOn && contractsMatchedOn !== vendor ? ` (matched ""${escapedMatchedOn}"")` : '';
  bsSh.getRange(contractsRow, 6, 1, 4).merge()
    .setFormula(`=HYPERLINK("${airtableContractsUrl}", "📋 AIRTABLE CONTRACTS (${contractsData.contractCount})${matchedDisplay}")`)
    .setBackground('#f8f9fa')
    .setFontWeight('bold')
    .setFontSize(10)
    .setFontColor('#1a73e8')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('top');
  bsSh.setRowHeight(contractsRow, 24);
  contractsRow++;
  
  if (!contractsData.hasContracts) {
    bsSh.getRange(contractsRow, 6, 1, 4).merge()
      .setValue('No contracts found')
      .setFontStyle('italic')
      .setBackground('#fafafa')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('top');
    bsSh.setRowHeight(contractsRow, 25);
    contractsRow++;
  } else {
    // Contract headers
    bsSh.getRange(contractsRow, 6).setValue('Contract').setFontWeight('bold').setBackground('#f3f3f3').setHorizontalAlignment('left');
    bsSh.getRange(contractsRow, 7).setValue('Type').setFontWeight('bold').setBackground('#f3f3f3').setHorizontalAlignment('left');
    bsSh.getRange(contractsRow, 8).setValue('Status').setFontWeight('bold').setBackground('#f3f3f3').setHorizontalAlignment('left');
    bsSh.getRange(contractsRow, 9).setValue('Notes').setFontWeight('bold').setBackground('#f3f3f3').setHorizontalAlignment('left');
    contractsRow++;
    
    for (const contract of contractsData.contracts.slice(0, 10)) {
      // Contract name - clickable link to Airtable
      const contractTitle = (contract.vendorName || 'View Contract')
        .replace(/"/g, '""')  // Escape double quotes for formula
        .replace(/\n/g, ' ')  // Replace newlines with space
        .replace(/\r/g, '');  // Remove carriage returns
      
      // Clean the URL - remove any problematic characters
      const cleanUrl = (contract.airtableUrl || '')
        .replace(/"/g, '%22')  // URL encode any quotes in URL
        .replace(/\s/g, '%20'); // URL encode spaces
      
      if (cleanUrl) {
        const formula = `=HYPERLINK("${cleanUrl}", "${contractTitle}")`;
        Logger.log(`Contract formula: ${formula}`);
        bsSh.getRange(contractsRow, 6)
          .setFormula(formula)
          .setFontColor('#1a73e8')
          .setHorizontalAlignment('left')
          .setVerticalAlignment('top');
      } else {
        bsSh.getRange(contractsRow, 6).setValue(contract.vendorName || 'Contract').setHorizontalAlignment('left').setVerticalAlignment('top');
      }
      
      bsSh.getRange(contractsRow, 7).setValue(contract.contractType || '').setHorizontalAlignment('left').setVerticalAlignment('top');

      // Default blank status to "Waiting on Legal"
      const displayStatus = contract.status || 'Waiting on Legal';
      bsSh.getRange(contractsRow, 8).setValue(displayStatus).setWrap(true).setHorizontalAlignment('left').setVerticalAlignment('top');
      bsSh.getRange(contractsRow, 9).setValue(contract.notes || '').setWrap(true).setHorizontalAlignment('left').setVerticalAlignment('top');
      
      // Color code by status
      const status = displayStatus.toLowerCase();
      if (status.includes('active') || status.includes('signed') || status.includes('executed')) {
        bsSh.getRange(contractsRow, 6, 1, 4).setBackground('#d9ead3'); // Green for active
      } else if (status.includes('pending') || status.includes('draft')) {
        bsSh.getRange(contractsRow, 6, 1, 4).setBackground('#fff2cc'); // Yellow for pending
      } else if (status.includes('expired') || status.includes('terminated')) {
        bsSh.getRange(contractsRow, 6, 1, 4).setBackground('#f3f3f3'); // Gray for expired
      } else if (status.includes('waiting')) {
        bsSh.getRange(contractsRow, 6, 1, 4).setBackground('#fce5cd'); // Light orange for waiting
      }
      
      contractsRow++;
    }
    
    if (contractsData.contractCount > 10) {
      bsSh.getRange(contractsRow, 6, 1, 4).merge()
        .setValue(`... and ${contractsData.contractCount - 10} more contracts`)
        .setFontStyle('italic')
        .setHorizontalAlignment('left');
      contractsRow++;
    }
  }
  
  // Update rightColumnRow to track furthest row used on right side
  rightColumnRow = Math.max(rightColumnRow, contractsRow);
  
  contractsRow++; // Add spacing
  
  // BOX DOCUMENTS SECTION (right side - aligned with Upcoming Meetings)
  ss.toast('Searching Box...', '📦 Loading', 2);
  let boxDocs = [];
  let boxRow = upcomingMeetingsStartRow;
  
  // Get blacklist from Settings sheet
  const boxBlacklist = getBoxBlacklist_();
  
  // Check cache for Box docs if useCache is true
  let boxDocsFromCache = null;
  if (useCache) {
    boxDocsFromCache = getCachedData_('box', vendor);
    if (boxDocsFromCache) {
      boxDocs = boxDocsFromCache;
      Logger.log(`Box docs loaded from cache: ${boxDocs.length} documents`);
    }
  }
  
  // Only search Box if not loaded from cache
  if (!boxDocsFromCache) {
    try {
      // Check if Box is authorized before searching
      const boxService = getBoxService_();
      if (boxService.hasAccess()) {
        // Search with primary vendor name
        const primaryDocs = searchBoxForVendor(vendor);
        Logger.log(`Box search for "${vendor}" found ${primaryDocs.length} results`);
      
      // Tag each result with the search term that found it
      for (const doc of primaryDocs) {
        doc.matchedOn = vendor;
        boxDocs.push(doc);
      }
      
      // Also try Other Name(s)
      if (contactData.otherName) {
        const existingIds = new Set(boxDocs.map(d => d.id));
        
        // First try the full Other Name value (in case it's like "Profitise, LLC")
        Logger.log(`Searching Box for full Other Name: "${contactData.otherName}"`);
        const fullNameDocs = searchBoxForVendor(contactData.otherName);
        Logger.log(`Box search for "${contactData.otherName}" found ${fullNameDocs.length} results`);
        
        for (const doc of fullNameDocs) {
          if (!existingIds.has(doc.id)) {
            doc.matchedOn = contactData.otherName;
            boxDocs.push(doc);
            existingIds.add(doc.id);
          }
        }
        
        // If full name found nothing, also try splitting by comma for multiple values
        if (fullNameDocs.length === 0 && contactData.otherName.includes(',')) {
          const otherNames = contactData.otherName.split(',').map(n => n.trim()).filter(n => n.length > 0);
          Logger.log(`Full name found nothing, trying individual values: ${otherNames.join(', ')}`);
          
          for (const altName of otherNames) {
            // Skip generic terms that cause false positives
            if (BS_CFG.SKIP_SEARCH_TERMS.some(term => term.toLowerCase() === altName.toLowerCase())) {
              Logger.log(`Skipping generic term: "${altName}"`);
              continue;
            }
            
            Logger.log(`Searching Box for: "${altName}"`);
            const otherNameDocs = searchBoxForVendor(altName);
            Logger.log(`Box search for "${altName}" found ${otherNameDocs.length} results`);
            
            // Add any new results (not already in boxDocs)
            for (const doc of otherNameDocs) {
              if (!existingIds.has(doc.id)) {
                doc.matchedOn = altName;
                boxDocs.push(doc);
                existingIds.add(doc.id);
              }
            }
          }
        }
        Logger.log(`Combined Box results: ${boxDocs.length} unique documents`);
      }

      // Filter out "Signing Log" documents (Box Sign creates these alongside the actual document)
      const beforeSigningLogCount = boxDocs.length;
      boxDocs = boxDocs.filter(doc => !(doc.name || '').toLowerCase().includes('signing log'));
      if (beforeSigningLogCount !== boxDocs.length) {
        Logger.log(`Removed ${beforeSigningLogCount - boxDocs.length} Signing Log documents`);
      }

      // Dedupe documents with same name + folder + modified + matchedOn
      // (Box API sometimes returns duplicates or same document with different IDs)
      const seen = new Set();
      const beforeDedupeCount = boxDocs.length;
      boxDocs = boxDocs.filter(doc => {
        const key = `${doc.name || ''}|${doc.folderName || ''}|${doc.modifiedAt || ''}|${doc.matchedOn || ''}`;
        if (seen.has(key)) {
          return false;
        }
        seen.add(key);
        return true;
      });
      if (beforeDedupeCount !== boxDocs.length) {
        Logger.log(`Deduped Box results: ${beforeDedupeCount} -> ${boxDocs.length}`);
      }

      // Helper to strip file extension for comparison
      const stripExtension = (name) => {
        return (name || '').toLowerCase().replace(/\.(pdf|doc|docx|xlsx|xls|ppt|pptx)$/i, '');
      };

      // Helper to check if file is a PDF
      const isPdf = (name) => (name || '').toLowerCase().endsWith('.pdf');

      // Remove "My Sign Requests" documents if same document exists in another folder
      // (My Sign Requests is a draft/pending folder, prefer the final version)
      // Compare without extensions so "file.pdf" and "file.docx" are considered duplicates
      const docNamesInOtherFolders = new Set();
      for (const doc of boxDocs) {
        const folderPath = (doc.folderPath || doc.parentFolder || '').toLowerCase();
        if (!folderPath.includes('my sign requests')) {
          docNamesInOtherFolders.add(stripExtension(doc.name));
        }
      }
      const beforeSignReqCount = boxDocs.length;
      boxDocs = boxDocs.filter(doc => {
        const folderPath = (doc.folderPath || doc.parentFolder || '').toLowerCase();
        const docNameNoExt = stripExtension(doc.name);
        // Keep if not in My Sign Requests, OR if no duplicate exists elsewhere
        return !folderPath.includes('my sign requests') || !docNamesInOtherFolders.has(docNameNoExt);
      });
      if (beforeSignReqCount !== boxDocs.length) {
        Logger.log(`Removed ${beforeSignReqCount - boxDocs.length} My Sign Requests duplicates`);
      }

      // Dedupe by base name (without extension), preferring PDF over DOC/DOCX
      const seenBaseNames = new Map(); // baseName -> doc (prefer PDF)
      for (const doc of boxDocs) {
        const baseName = stripExtension(doc.name);
        const existing = seenBaseNames.get(baseName);
        if (!existing) {
          seenBaseNames.set(baseName, doc);
        } else {
          // Prefer PDF over non-PDF
          if (isPdf(doc.name) && !isPdf(existing.name)) {
            seenBaseNames.set(baseName, doc);
          }
          // If both same type, keep the one in a better folder (Profitise > others)
          else if (isPdf(doc.name) === isPdf(existing.name)) {
            const docFolder = (doc.folderPath || '').toLowerCase();
            const existingFolder = (existing.folderPath || '').toLowerCase();
            if (docFolder.includes('profitise') && !existingFolder.includes('profitise')) {
              seenBaseNames.set(baseName, doc);
            }
          }
        }
      }
      const beforeExtDedupeCount = boxDocs.length;
      boxDocs = [...seenBaseNames.values()];
      if (beforeExtDedupeCount !== boxDocs.length) {
        Logger.log(`Removed ${beforeExtDedupeCount - boxDocs.length} extension duplicates (preferring PDF)`);
      }

      // Apply blacklist - remove files blacklisted for this vendor
      if (boxBlacklist[vendor]) {
        const blacklistedIds = boxBlacklist[vendor];
        const beforeCount = boxDocs.length;
        boxDocs = boxDocs.filter(doc => !blacklistedIds.includes(doc.id));
        if (beforeCount !== boxDocs.length) {
          Logger.log(`Blacklist removed ${beforeCount - boxDocs.length} files for ${vendor}`);
        }
      }
      
      // Sort Box results: 
      // 1. Vendor name matches first, then Other Name matches (in order they were searched)
      // 2. Then by Modified date DESC
      // 3. Then by Folder name DESC
      if (boxDocs.length > 0) {
        // Build priority map: vendor name = 0, then each other name in order
        const matchPriority = { [vendor]: 0 };
        if (contactData.otherName) {
          // Full other name gets priority 1
          matchPriority[contactData.otherName] = 1;
          // Individual values get subsequent priorities
          if (contactData.otherName.includes(',')) {
            const otherNames = contactData.otherName.split(',').map(n => n.trim()).filter(n => n.length > 0);
            otherNames.forEach((name, idx) => {
              if (!(name in matchPriority)) {
                matchPriority[name] = idx + 2;
              }
            });
          }
        }
        
        boxDocs.sort((a, b) => {
          // First: sort by modified date DESC (newest first)
          const dateA = a.modifiedAt || '';
          const dateB = b.modifiedAt || '';
          if (dateA !== dateB) return dateB.localeCompare(dateA);
          
          // Second: sort by match priority ASC (vendor name first, then other names in order)
          const priorityA = matchPriority[a.matchedOn] ?? 999;
          const priorityB = matchPriority[b.matchedOn] ?? 999;
          if (priorityA !== priorityB) return priorityA - priorityB;
          
          // Third: sort by document name ASC
          const nameA = a.name || '';
          const nameB = b.name || '';
          return nameA.localeCompare(nameB);
        });
        
        Logger.log(`Sorted Box results by modified DESC, then matched ASC, then document ASC`);
      }
      
      // Cache the Box results
      setCachedData_('box', vendor, boxDocs);
      
    } else {
      Logger.log('Box not authorized - skipping Box search');
    }
  } catch (e) {
    Logger.log(`Box search error: ${e.message}`);
  }
  } // End of !boxDocsFromCache block
  
  bsSh.getRange(boxRow, 6, 1, 4).merge()
    .setValue(`📦 BOX DOCUMENTS (${boxDocs.length})`)
    .setBackground('#f8f9fa')
    .setFontWeight('bold')
    .setFontSize(10)
    .setFontColor('#1a73e8')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('top');
  bsSh.setRowHeight(boxRow, 24);
  boxRow++;
  
  if (boxDocs.length === 0) {
    bsSh.getRange(boxRow, 6, 1, 4).merge()
      .setValue('No Box documents found')
      .setFontStyle('italic')
      .setBackground('#fafafa')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('top');
    bsSh.setRowHeight(boxRow, 25);
    boxRow++;
  } else {
    // Box document headers
    bsSh.getRange(boxRow, 6).setValue('Document').setFontWeight('bold').setBackground('#f3f3f3').setHorizontalAlignment('left');
    bsSh.getRange(boxRow, 7).setValue('Folder').setFontWeight('bold').setBackground('#f3f3f3').setHorizontalAlignment('left');
    bsSh.getRange(boxRow, 8).setValue('Modified').setFontWeight('bold').setBackground('#f3f3f3').setHorizontalAlignment('left');
    bsSh.getRange(boxRow, 9).setValue('Matched').setFontWeight('bold').setBackground('#f3f3f3').setHorizontalAlignment('left');
    boxRow++;

    // Check if there's a document in "Profitise > VENDOR_NAME" folder (the primary/final location)
    // System folders to exclude from being considered "vendor folders"
    const systemFolders = ['w-9', 'w9', 'new publishers', 'my sign', 'my signed', 'templates'];

    // Helper to check if a folder path is a vendor-specific Profitise folder
    const isProfitiseVendorFolder = (path) => {
      const pathLower = (path || '').toLowerCase();
      if (!pathLower.includes('profitise/')) return false;
      // Check it's not a system folder
      return !systemFolders.some(sys => pathLower.includes(sys));
    };

    // Find if any doc is in a Profitise vendor folder
    const hasProfitiseVendorFolder = boxDocs.some(doc => isProfitiseVendorFolder(doc.folderPath));

    for (const doc of boxDocs.slice(0, 10)) {
      const folderPath = (doc.folderPath || '').toLowerCase();
      const isInProfitiseVendorFolder = isProfitiseVendorFolder(doc.folderPath);
      const isInW9Folder = folderPath.includes('w-9') || folderPath.includes('w9');
      // Gray out if: there's a vendor folder doc AND this isn't in vendor folder AND this isn't W-9
      const shouldGrayOut = hasProfitiseVendorFolder && !isInProfitiseVendorFolder && !isInW9Folder;

      // Document name - clickable link to Box (no truncation - user controls column width)
      const docCell = bsSh.getRange(boxRow, 6)
        .setFormula(`=HYPERLINK("${doc.webUrl}", "${doc.name.replace(/"/g, '""')}")`)
        .setHorizontalAlignment('left')
        .setVerticalAlignment('top');

      if (shouldGrayOut) {
        docCell.setFontColor('#999999');
      } else {
        docCell.setFontColor('#1a73e8');
      }
      
      // Folder - show full path with underscores, clickable link to parent folder
      // folderPath is like "All Files/Profitise/Company Name" 
      // Skip "All Files" and join with " > "
      let folderDisplayName = 'Root';
      if (doc.folderPath) {
        const pathParts = doc.folderPath.split('/').filter(p => p && p !== 'All Files');
        folderDisplayName = pathParts.join(' > ') || 'Root';
      } else if (doc.parentFolder) {
        folderDisplayName = doc.parentFolder;
      }
      
      // No truncation - user controls column width
      const folderUrl = doc.parentFolderUrl || '';
      const folderCell = bsSh.getRange(boxRow, 7);
      if (folderUrl) {
        folderCell
          .setFormula(`=HYPERLINK("${folderUrl}", "${folderDisplayName.replace(/"/g, '""')}")`)
          .setHorizontalAlignment('left')
          .setVerticalAlignment('top');
      } else {
        folderCell.setValue(folderDisplayName).setHorizontalAlignment('left').setVerticalAlignment('top');
      }
      folderCell.setFontColor(shouldGrayOut ? '#999999' : '#1a73e8');

      // Modified date
      const modDate = doc.modifiedAt ? Utilities.formatDate(new Date(doc.modifiedAt), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
      const modCell = bsSh.getRange(boxRow, 8).setValue(modDate).setHorizontalAlignment('left').setVerticalAlignment('top');
      if (shouldGrayOut) modCell.setFontColor('#999999');

      // Matched search term (in quotes)
      const matchedTerm = doc.matchedOn ? `"${doc.matchedOn}"` : '';
      const matchCell = bsSh.getRange(boxRow, 9).setValue(matchedTerm).setFontStyle('italic').setHorizontalAlignment('left').setVerticalAlignment('top');
      matchCell.setFontColor(shouldGrayOut ? '#999999' : '#666666');
      
      boxRow++;
    }
    
    if (boxDocs.length > 10) {
      bsSh.getRange(boxRow, 6, 1, 4).merge()
        .setValue(`... and ${boxDocs.length - 10} more documents`)
        .setFontStyle('italic')
        .setHorizontalAlignment('left');
      boxRow++;
    }
  }
  
  // Update rightColumnRow to track furthest row used on right side
  rightColumnRow = Math.max(rightColumnRow, boxRow);
  
  // Fetch Google Drive files now, but display section later (aligned with EMAILS)
  ss.toast('Searching Google Drive...', '📁 Loading', 2);
  let gDriveFiles = [];
  let gDriveFolderFound = false;
  let gDriveFolderUrl = null;
  let gDriveMatchedOn = '';
  
  // Check cache for GDrive files if useCache is true
  let gDriveFromCache = null;
  if (useCache) {
    gDriveFromCache = getCachedData_('gdrive', vendor);
    if (gDriveFromCache) {
      gDriveFiles = gDriveFromCache.files || [];
      gDriveFolderFound = gDriveFromCache.folderFound || false;
      gDriveFolderUrl = gDriveFromCache.folderUrl || null;
      gDriveMatchedOn = gDriveFromCache.matchedOn || '';
      Logger.log(`GDrive files loaded from cache: ${gDriveFiles.length} files`);
    }
  }
  
  // Only search GDrive if not loaded from cache
  if (!gDriveFromCache) {
    try {
      const result = getGDriveFilesForVendor_(vendor);
      gDriveFiles = result.files || [];
      gDriveFolderFound = result.folderFound || false;
      gDriveFolderUrl = result.folderUrl || null;
      if (gDriveFolderFound) gDriveMatchedOn = vendor;
      
      // Only try Other Name(s) if NO FOLDER was found (not just empty folder)
      if (!gDriveFolderFound && contactData.otherName) {
        // First try the full Other Name value (in case it's like "Profitise, LLC")
        Logger.log(`No GDrive folder for "${vendor}", trying full Other Name: "${contactData.otherName}"`);
        const fullResult = getGDriveFilesForVendor_(contactData.otherName);
        
        if (fullResult.folderFound) {
          gDriveFiles = fullResult.files || [];
          gDriveFolderFound = true;
          gDriveFolderUrl = fullResult.folderUrl || null;
          gDriveMatchedOn = contactData.otherName;
        } else if (contactData.otherName.includes(',')) {
          // If full name found nothing, try splitting by comma for multiple values
          const otherNames = contactData.otherName.split(',').map(n => n.trim()).filter(n => n.length > 0);
          Logger.log(`Full name found nothing, trying individual values: ${otherNames.join(', ')}`);
          
          for (const altName of otherNames) {
            // Skip generic terms that cause false positives
            if (BS_CFG.SKIP_SEARCH_TERMS.some(term => term.toLowerCase() === altName.toLowerCase())) {
              Logger.log(`Skipping generic term: "${altName}"`);
              continue;
            }
            Logger.log(`Searching GDrive for: "${altName}"`);
            const altResult = getGDriveFilesForVendor_(altName);
            if (altResult.folderFound) {
              gDriveFiles = altResult.files || [];
              gDriveFolderFound = true;
              gDriveFolderUrl = altResult.folderUrl || null;
              gDriveMatchedOn = altName;
              break; // Found a folder, stop searching
            }
          }
        }
      }
      
      // Cache the GDrive results
      setCachedData_('gdrive', vendor, {
        files: gDriveFiles,
        folderFound: gDriveFolderFound,
        folderUrl: gDriveFolderUrl,
        matchedOn: gDriveMatchedOn
      });
      
    } catch (e) {
      Logger.log(`Google Drive search error: ${e.message}`);
    }
  }
  
  // ========== CONTINUE LEFT SIDE ==========
  
  // Notes section
  bsSh.getRange(currentRow, 1, 1, 4).merge()
    .setValue('📝 NOTES')
    .setBackground('#f8f9fa')
    .setFontWeight('bold')
    .setFontSize(10)
    .setFontColor('#1a73e8')
    .setHorizontalAlignment('left');
  bsSh.setRowHeight(currentRow, 24);
  currentRow++;
  
  const notesCell = bsSh.getRange(currentRow, 1, 2, 4).merge()
    .setValue(mondayNotes || '(no notes)')
    .setWrap(true)
    .setVerticalAlignment('top');
  
  if (!mondayNotes || mondayNotes === '(no notes)') {
    // Link to monday.com board filtered by vendor
    const encodedVendor = encodeURIComponent(vendor);
    const mondayFilterLink = `https://profitise-company.monday.com/boards/${contactData.boardId || (source.toLowerCase().includes('affiliate') ? BS_CFG.AFFILIATES_BOARD_ID : BS_CFG.BUYERS_BOARD_ID)}?term=${encodedVendor}`;
    notesCell.setFormula(`=HYPERLINK("${mondayFilterLink}", "⚠️ Add notes in monday.com")`);
    notesCell.setBackground(BS_CFG.COLOR_MISSING).setFontColor('#1a73e8');
  } else {
    notesCell.setBackground('#fafafa');
  }
  
  currentRow += 2;
  currentRow++;
  
  // Track row where Emails starts (for Google Drive alignment)
  const emailsStartRow = currentRow;
  
  // GOOGLE DRIVE FOLDER SECTION (right side - starts after Box section OR aligned with Emails, whichever is later)
  let gDriveRow = Math.max(emailsStartRow, rightColumnRow);
  
  // Get folder URL - use tracked URL or default to Vendors folder
  const displayFolderUrl = gDriveFolderUrl || `https://drive.google.com/drive/folders/${BS_CFG.GDRIVE_VENDORS_FOLDER_ID}`;
  
  bsSh.getRange(gDriveRow, 6, 1, 4).merge()
    .setFormula(`=HYPERLINK("${displayFolderUrl}", "📁 GOOGLE DRIVE (${gDriveFiles.length})")`)
    .setBackground('#f8f9fa')
    .setFontWeight('bold')
    .setFontSize(10)
    .setFontColor('#1a73e8')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('top');
  bsSh.setRowHeight(gDriveRow, 24);
  gDriveRow++;
  
  if (!gDriveFolderFound) {
    // No folder found - show link to create
    const vendorsFolderUrl = `https://drive.google.com/drive/folders/${BS_CFG.GDRIVE_VENDORS_FOLDER_ID}`;
    bsSh.getRange(gDriveRow, 6, 1, 4).merge()
      .setFormula(`=HYPERLINK("${vendorsFolderUrl}", "📂 No folder found - Click to create in Vendors")`)
      .setFontStyle('italic')
      .setFontColor('#1a73e8')
      .setBackground('#fafafa')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('top');
    bsSh.setRowHeight(gDriveRow, 25);
    gDriveRow++;
  } else if (gDriveFiles.length === 0) {
    // Folder found but empty
    bsSh.getRange(gDriveRow, 6, 1, 4).merge()
      .setFormula(`=HYPERLINK("${displayFolderUrl}", "📂 Folder empty - Click to open")`)
      .setFontStyle('italic')
      .setFontColor('#1a73e8')
      .setBackground('#fafafa')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('top');
    bsSh.setRowHeight(gDriveRow, 25);
    gDriveRow++;
  } else {
    // Google Drive file headers - show matched term in header if available
    const matchedDisplay = gDriveMatchedOn ? ` (matched "${gDriveMatchedOn}")` : '';
    bsSh.getRange(gDriveRow, 6).setValue('File').setFontWeight('bold').setBackground('#f3f3f3').setHorizontalAlignment('left');
    bsSh.getRange(gDriveRow, 7).setValue('Type').setFontWeight('bold').setBackground('#f3f3f3').setHorizontalAlignment('left');
    bsSh.getRange(gDriveRow, 8).setValue('Modified').setFontWeight('bold').setBackground('#f3f3f3').setHorizontalAlignment('left');
    bsSh.getRange(gDriveRow, 9).setValue(matchedDisplay).setFontStyle('italic').setFontColor('#666666').setBackground('#f3f3f3').setHorizontalAlignment('left');
    gDriveRow++;
    
    for (const file of gDriveFiles.slice(0, 10)) {
      // File name - clickable link (no truncation - user controls column width)
      bsSh.getRange(gDriveRow, 6)
        .setFormula(`=HYPERLINK("${file.url}", "${file.name.replace(/"/g, '""')}")`)
        .setFontColor('#1a73e8')
        .setHorizontalAlignment('left')
        .setVerticalAlignment('top');
      
      // File type
      bsSh.getRange(gDriveRow, 7).setValue(file.type).setHorizontalAlignment('left').setVerticalAlignment('top');
      
      // Modified date
      bsSh.getRange(gDriveRow, 8).setValue(file.modified).setHorizontalAlignment('left').setVerticalAlignment('top');
      
      // Clear column 9 (no per-file matched term for Google Drive - it's at folder level)
      bsSh.getRange(gDriveRow, 9).setValue('').setBackground(null);
      
      gDriveRow++;
    }
    
    if (gDriveFiles.length > 10) {
      // Link to the vendor's folder in Google Drive
      const folderUrl = gDriveFiles[0].folderUrl || `https://drive.google.com/drive/folders/${BS_CFG.GDRIVE_VENDORS_FOLDER_ID}`;
      bsSh.getRange(gDriveRow, 6, 1, 4).merge()
        .setFormula(`=HYPERLINK("${folderUrl}", "... and ${gDriveFiles.length - 10} more files - Open Folder")`)
        .setFontStyle('italic')
        .setFontColor('#1a73e8')
        .setHorizontalAlignment('left');
      gDriveRow++;
    }
  }
  
  // Update rightColumnRow to track furthest row used on right side
  rightColumnRow = Math.max(rightColumnRow, gDriveRow);

  // CRYSTAL BALL SECTION (right side - below Google Drive)
  ss.toast('Analyzing emails...', '🔮 Crystal Ball', 2);
  const crystalBall = getCrystalBallData_(vendor, listRow);

  let crystalRow = rightColumnRow + 1;

  // Crystal Ball header
  const crystalCount = crystalBall.items.length + crystalBall.snoozed.length;
  bsSh.getRange(crystalRow, 6, 1, 4).merge()
    .setValue(`🔮 CRYSTAL BALL (${crystalCount} threads)`)
    .setBackground('#e8f5e9')  // Light green
    .setFontWeight('bold')
    .setFontSize(10)
    .setFontColor('#2e7d32')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('top');
  bsSh.setRowHeight(crystalRow, 24);
  crystalRow++;

  if (crystalBall.error) {
    bsSh.getRange(crystalRow, 6, 1, 4).merge()
      .setValue(`Error: ${crystalBall.error}`)
      .setFontStyle('italic')
      .setFontColor('#d32f2f')
      .setBackground('#fafafa')
      .setHorizontalAlignment('left');
    crystalRow++;
  } else if (crystalBall.summary) {
    // Display the AI-generated summary
    const summaryLines = crystalBall.summary.split('\n').filter(l => l.trim());

    for (const line of summaryLines.slice(0, 8)) {
      bsSh.getRange(crystalRow, 6, 1, 4).merge()
        .setValue(line.trim())
        .setFontSize(9)
        .setBackground('#fafafa')
        .setWrap(true)
        .setHorizontalAlignment('left')
        .setVerticalAlignment('top');
      bsSh.setRowHeight(crystalRow, 22);
      crystalRow++;
    }

    if (summaryLines.length > 8) {
      bsSh.getRange(crystalRow, 6, 1, 4).merge()
        .setValue(`... and ${summaryLines.length - 8} more items`)
        .setFontStyle('italic')
        .setFontColor('#666666')
        .setBackground('#fafafa')
        .setHorizontalAlignment('left');
      crystalRow++;
    }
  } else {
    bsSh.getRange(crystalRow, 6, 1, 4).merge()
      .setValue('No outstanding items found')
      .setFontStyle('italic')
      .setFontColor('#666666')
      .setBackground('#fafafa')
      .setHorizontalAlignment('left');
    crystalRow++;
  }

  // Update rightColumnRow
  rightColumnRow = Math.max(rightColumnRow, crystalRow);

  // EMAILS SECTION
  ss.toast('Searching Gmail...', '📧 Loading', 2);
  const emails = getEmailsForVendor_(vendor, listRow);
  
  bsSh.getRange(currentRow, 1, 1, 4).merge()
    .setValue(`📧 EMAILS (${emails.length})  |  🔵 Snoozed  🔴 Overdue  🟠 Phonexa  🟢 Accounting  🟡 Waiting`)
    .setBackground('#f8f9fa')
    .setFontWeight('bold')
    .setFontSize(10)
    .setFontColor('#1a73e8')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  bsSh.setRowHeight(currentRow, 24);
  currentRow++;
  
  if (emails.length === 0) {
    bsSh.getRange(currentRow, 1, 1, 4).merge()
      .setValue('No emails found')
      .setFontStyle('italic')
      .setBackground('#fafafa')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    bsSh.setRowHeight(currentRow, 25);
    currentRow++;
  } else {
    // Email headers
    bsSh.getRange(currentRow, 1).setValue('Subject').setFontWeight('bold').setBackground('#f3f3f3');
    bsSh.getRange(currentRow, 2).setValue('Date').setFontWeight('bold').setBackground('#f3f3f3');
    bsSh.getRange(currentRow, 3).setValue('Last').setFontWeight('bold').setBackground('#f3f3f3');
    bsSh.getRange(currentRow, 4).setValue('Labels').setFontWeight('bold').setBackground('#f3f3f3');
    currentRow++;

    for (const email of emails.slice(0, 20)) {
      bsSh.getRange(currentRow, 1).setValue(email.subject);
      const emailDateCell = bsSh.getRange(currentRow, 2);
      emailDateCell.setNumberFormat('@'); // Set format BEFORE value to prevent auto-parsing
      emailDateCell.setValue(email.date);
      bsSh.getRange(currentRow, 3).setValue(email.lastFrom);
      bsSh.getRange(currentRow, 4).setValue(email.labels);
      
      if (email.link) {
        bsSh.getRange(currentRow, 1)
          .setFormula(`=HYPERLINK("${email.link}", "${email.subject.replace(/"/g, '""')}")`);
      }
      
      // Check if this email is overdue (waiting/customer + >16 business hours)
      const isOverdue = isEmailOverdue_(email);
      
      // Check if email has priority label
      const hasPriority = email.labels.includes('01.priority/1');
      
      // Color priority: Snoozed (blue) > OVERDUE (red) > Phonexa (coral) > Accounting (green) > Waiting/Customer (yellow) > Active (white)
      if (email.isSnoozed) {
        bsSh.getRange(currentRow, 1, 1, 4).setBackground(BS_CFG.COLOR_SNOOZED);
      } else if (isOverdue) {
        bsSh.getRange(currentRow, 1, 1, 4).setBackground(BS_CFG.COLOR_OVERDUE);
        bsSh.getRange(currentRow, 1, 1, 4).setFontWeight('bold');
      } else if (email.labels.includes('02.waiting/phonexa')) {
        bsSh.getRange(currentRow, 1, 1, 4).setBackground(BS_CFG.COLOR_PHONEXA);
      } else if (email.labels.includes('04.accounting-invoices')) {
        bsSh.getRange(currentRow, 1, 1, 4).setBackground('#d9ead3'); // Green
      } else if (email.labels.includes('02.waiting/customer')) {
        bsSh.getRange(currentRow, 1, 1, 4).setBackground(BS_CFG.COLOR_WAITING);
      } else {
        bsSh.getRange(currentRow, 1, 1, 4).setBackground('#ffffff');
      }
      
      // If missing 01.priority/1, make text grey to indicate lower importance
      if (!hasPriority) {
        bsSh.getRange(currentRow, 1, 1, 4).setFontColor('#999999');
      }
      
      currentRow++;
    }
    
    if (emails.length > 20) {
      bsSh.getRange(currentRow, 1, 1, 4).merge()
        .setValue(`... and ${emails.length - 20} more emails (showing first 20)`)
        .setFontStyle('italic')
        .setHorizontalAlignment('center');
      currentRow++;
    }
  }
  
  bsSh.setRowHeight(currentRow, 10);
  currentRow++;
  
  // TASKS SECTION
  let tasks = getTasksForVendor_(vendor, listRow);
  
  // Filter out inappropriate onboarding tasks based on source
  // If source is Affiliates, don't show "Onboarding - Buyer" tasks
  // If source is Buyers, don't show "Onboarding - Affiliate" tasks
  const isAffiliate = source.toLowerCase().includes('affiliate');
  const isBuyer = source.toLowerCase().includes('buyer');
  
  tasks = tasks.filter(task => {
    const project = (task.project || '').toLowerCase();
    if (isAffiliate && project.includes('onboarding - buyer')) {
      return false;
    }
    if (isBuyer && project.includes('onboarding - affiliate')) {
      return false;
    }
    return true;
  });
  
  const nonDoneTasks = tasks.filter(t => !t.isDone);
  
  bsSh.getRange(currentRow, 1, 1, 4).merge()
    .setValue(`📋 MONDAY.COM TASKS (${nonDoneTasks.length})`)
    .setBackground('#f8f9fa')
    .setFontWeight('bold')
    .setFontSize(10)
    .setFontColor('#1a73e8')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  bsSh.setRowHeight(currentRow, 24);
  currentRow++;
  
  if (tasks.length === 0) {
    // Check if vendor is Live/Paused/Onboarding - should have tasks
    // Note: "preonboarding" contains "onboarding" so we must exclude it explicitly
    const statusLower = (liveStatus || '').toLowerCase();
    const isPreonboarding = statusLower.includes('pre');
    const needsTasksWarning = !isPreonboarding && (
                               statusLower.includes('live') ||
                               statusLower.includes('onboarding') ||
                               statusLower.includes('paused'));

    if (needsTasksWarning) {
      // Show warning with link to Claude task generator
      const vendorType = source.toLowerCase().includes('affiliate') ? 'Affiliate' : 'Buyer';
      const claudeChatUrl = 'https://claude.ai/chat/33d0e36c-23ad-4e7d-b354-bd6cf3692f3f';
      bsSh.getRange(currentRow, 1, 1, 4).merge()
        .setFormula(`=HYPERLINK("${claudeChatUrl}", "⚠️ No tasks - Click to generate tasks")`)
        .setFontColor('#d32f2f')
        .setBackground('#ffebee')
        .setHorizontalAlignment('left')
        .setVerticalAlignment('middle');
      bsSh.setRowHeight(currentRow, 25);
      currentRow++;

      // Add copy/paste line for vendor name and type
      bsSh.getRange(currentRow, 1, 1, 4).merge()
        .setValue(`${vendor} (${vendorType})`)
        .setFontStyle('italic')
        .setFontColor('#666666')
        .setBackground('#fafafa')
        .setHorizontalAlignment('left')
        .setVerticalAlignment('middle');
      bsSh.setRowHeight(currentRow, 22);
      currentRow++;
    } else {
      bsSh.getRange(currentRow, 1, 1, 4).merge()
        .setValue('No tasks found')
        .setFontStyle('italic')
        .setBackground('#fafafa')
        .setHorizontalAlignment('left')
        .setVerticalAlignment('middle');
      bsSh.setRowHeight(currentRow, 25);
      currentRow++;
    }
  } else {
    bsSh.getRange(currentRow, 1).setValue('Task').setFontWeight('bold').setBackground('#f3f3f3');
    bsSh.getRange(currentRow, 2).setValue('Status').setFontWeight('bold').setBackground('#f3f3f3');
    bsSh.getRange(currentRow, 3).setValue('Created').setFontWeight('bold').setBackground('#f3f3f3');
    bsSh.getRange(currentRow, 4).setValue('Project').setFontWeight('bold').setBackground('#f3f3f3');
    currentRow++;
    
    for (const task of tasks) {
      // Task name - clickable link to Tasks board filtered by task name
      const encodedTask = encodeURIComponent(task.subject);
      const taskFilterLink = `https://profitise-company.monday.com/boards/${BS_CFG.TASKS_BOARD_ID}?term=${encodedTask}`;
      bsSh.getRange(currentRow, 1)
        .setFormula(`=HYPERLINK("${taskFilterLink}", "${task.subject.replace(/"/g, '""')}")`)
        .setWrap(true)
        .setFontColor('#1a73e8');
      
      // Status display:
      // - For Done tasks: show "Done - YYYY-MM-DD" (lastUpdated date)
      // - For non-Done tasks with taskDate: show "Status - taskDate"
      // - Otherwise: just show status
      let statusDisplay = task.status;
      if (task.isDone && task.lastUpdated) {
        statusDisplay = `Done - ${task.lastUpdated}`;
      } else if (task.taskDate && !task.isDone) {
        statusDisplay = `${task.status} - ${task.taskDate}`;
      }
      bsSh.getRange(currentRow, 2).setValue(statusDisplay).setWrap(true);
      const taskDateCell = bsSh.getRange(currentRow, 3);
      taskDateCell.setNumberFormat('@'); // Set format BEFORE value to prevent auto-parsing
      taskDateCell.setValue(task.created).setWrap(true);
      bsSh.getRange(currentRow, 4).setValue(task.project).setWrap(true);
      
      // Color coding for task status
      if (task.isDone) {
        bsSh.getRange(currentRow, 1, 1, 4)
          .setFontLine('line-through')
          .setFontColor('#999999');
      } else if (task.status && task.status.toLowerCase().includes('waiting on phonexa')) {
        bsSh.getRange(currentRow, 1, 1, 4)
          .setBackground('#ffcdd2');  // Red for waiting on phonexa
      } else if (task.status && task.status.toLowerCase().includes('waiting on client')) {
        bsSh.getRange(currentRow, 1, 1, 4)
          .setBackground('#fff2cc');  // Yellow for waiting on client
      }
      
      currentRow++;
    }
  }
  
  
  // Use the greater of currentRow or rightColumnRow for final row count
  const finalRow = Math.max(currentRow, rightColumnRow);
  
  // Style the divider column (column 5) - black background from row 3 to end
  if (finalRow > 2) {
    bsSh.getRange(3, 5, finalRow - 2, 1)
      .setBackground('#000000')
      .setValue('');
  }
  
  if (finalRow > 5) {
    bsSh.autoResizeRows(5, finalRow - 5);
  }
  
  // Generate and compare checksum to detect changes
  try {
    // Generate module sub-checksums
    const newModuleChecksums = generateModuleChecksums_(
      vendor, emails, tasks, contactData.notes || '', contactData.liveStatus || '',
      contactData.states || '', contractsData.contracts || [], helpfulLinks || [],
      meetings || [], boxDocs || [], gDriveFiles || [], contacts || []
    );
    
    // Generate email sub-checksum (for backward compatibility and early exit)
    const newEmailChecksum = newModuleChecksums.emails;
    
    // Get previously stored checksums
    const storedData = getStoredChecksum_(vendor);
    Logger.log(`Stored data for ${vendor}: ${storedData ? JSON.stringify({checksum: storedData.checksum, hasModules: !!storedData.moduleChecksums}) : 'null'}`);
    
    // Generate full checksum
    const newChecksum = generateVendorChecksum_(
      vendor, emails, tasks, contactData.notes || '', contactData.liveStatus || '',
      contactData.states || '', contractsData.contracts || [], helpfulLinks || [],
      meetings || [], boxDocs || [], gDriveFiles || []
    );
    
    Logger.log(`Checksum comparison for ${vendor}: stored=${storedData?.checksum} (type: ${typeof storedData?.checksum}), new=${newChecksum} (type: ${typeof newChecksum})`);
    const isUnchanged = !forceChanged && storedData && String(storedData.checksum) === String(newChecksum);
    Logger.log(`Is unchanged: ${isUnchanged}${forceChanged ? ' (forceChanged=true)' : ''}`);
    
    // Determine which modules changed (only if stored version matches to avoid false positives)
    const changedModules = [];
    const storedModuleVersion = storedData?.moduleChecksums?._version || 0;
    if (storedData && storedData.moduleChecksums && storedModuleVersion === MODULE_CHECKSUMS_VERSION) {
      const stored = storedData.moduleChecksums;
      if (stored.emails !== newModuleChecksums.emails) changedModules.push('emails');
      if (stored.tasks !== newModuleChecksums.tasks) changedModules.push('tasks');
      if (stored.notes !== newModuleChecksums.notes) changedModules.push('notes');
      if (stored.status !== newModuleChecksums.status) changedModules.push('status');
      if (stored.states !== newModuleChecksums.states) changedModules.push('states');
      if (stored.contracts !== newModuleChecksums.contracts) changedModules.push('contracts');
      if (stored.helpfulLinks !== newModuleChecksums.helpfulLinks) changedModules.push('helpfulLinks');
      if (stored.meetings !== newModuleChecksums.meetings) changedModules.push('meetings');
      if (stored.boxDocs !== newModuleChecksums.boxDocs) changedModules.push('boxDocs');
      if (stored.gDriveFiles !== newModuleChecksums.gDriveFiles) changedModules.push('gDriveFiles');
      if (stored.contacts !== newModuleChecksums.contacts) changedModules.push('contacts');
    }
    
    // If we stopped due to overdue emails, make sure emails is in changedModules
    if (changeType === 'overdue emails' && !changedModules.includes('emails')) {
      changedModules.push('emails');
      Logger.log(`Added emails to changedModules due to overdue emails`);
    }

    if (isUnchanged && !changedModules.length) {
      // Add ✅ to title row
      const currentTitle = bsSh.getRange(1, 1).getValue();
      bsSh.getRange(1, 1).setValue(`${currentTitle} ✅`);
      Logger.log(`Added ✅ indicator for ${vendor} - no changes`);
    } else if (storedData && changedModules.length > 0) {
      // Highlight changed section headers with 🔄
      Logger.log(`Changed modules for ${vendor}: ${changedModules.join(', ')}`);

      // Map module names to their header row search patterns
      const moduleHeaderMap = {
        'emails': '📧 EMAILS',
        'tasks': '📋 MONDAY.COM TASKS',
        'meetings': '📅 UPCOMING MEETINGS',
        'contracts': '📋 AIRTABLE CONTRACTS',
        'boxDocs': '📦 BOX DOCUMENTS',
        'gDriveFiles': '📁 GOOGLE DRIVE',
        'contacts': '👤 CONTACTS',
        'notes': '📝 NOTES',
        'helpfulLinks': '🔗 HELPFUL LINKS',
        'status': '📊 VENDOR INFO',
        'states': '📊 VENDOR INFO'
      };
      
      // Find and update headers for changed modules
      const dataRange = bsSh.getDataRange();
      const values = dataRange.getValues();
      
      for (const moduleName of changedModules) {
        const searchPattern = moduleHeaderMap[moduleName];
        if (!searchPattern) continue;
        
        for (let row = 0; row < values.length; row++) {
          const cellValue = String(values[row][0] || '');
          if (cellValue.includes(searchPattern) && !cellValue.includes('🔄')) {
            // Add 🔄 indicator to show this section changed
            bsSh.getRange(row + 1, 1).setValue(cellValue + ' 🔄');
            Logger.log(`Marked ${searchPattern} as changed (row ${row + 1})`);
            break;
          }
        }
      }
    } else if (!storedData) {
      Logger.log(`First view for ${vendor} - no previous checksums`);
    }
    
    // Create email summary data for change tracking
    const emailData = (emails || []).map(e => ({
      subject: (e.subject || '').substring(0, 60),
      date: e.date
    }));

    // Log email changes and get added/removed details
    let emailChanges = { added: [], removed: [] };
    if (changedModules.includes('emails') && storedData && storedData.emailData) {
      emailChanges = logEmailChanges_(storedData.emailData, emailData, vendor);
    }

    // ========== WHAT CHANGED SECTION (right side - below Google Drive) ==========
    // Show if:
    // 1. There's a legitimate changeType (Overdue emails, Flagged, etc.) - always show
    // 2. There are changedModules AND stored version matches (avoids false positives from format changes)
    const hasLegitimateChangeType = changeType && changeType !== 'First view' && changeType !== 'unchanged';
    const hasCompatibleModuleChanges = changedModules.length > 0;  // Already filtered by version above
    const shouldShowWhatChanged = forceChanged || hasLegitimateChangeType || hasCompatibleModuleChanges;
    if (shouldShowWhatChanged) {
      // Find where to render - use rightColumnRow which tracks furthest row on right side
      let changeRow = rightColumnRow + 1;

      bsSh.getRange(changeRow, 6, 1, 4).merge()
        .setValue('🔄 WHAT CHANGED')
        .setBackground('#fff3cd')  // Light yellow/amber
        .setFontWeight('bold')
        .setFontSize(10)
        .setFontColor('#856404')
        .setHorizontalAlignment('left')
        .setVerticalAlignment('top');
      bsSh.setRowHeight(changeRow, 24);
      changeRow++;

      // Build change descriptions
      const changeDescriptions = [];

      // Add the main changeType if provided (from skip detection)
      if (changeType && changeType !== 'First view') {
        changeDescriptions.push({ text: `• ${changeType}` });
      }

      // Add detailed module changes
      if (changedModules.length > 0) {
        // Email changes with details
        if (changedModules.includes('emails')) {
          const emailDesc = getEmailChangeDescription_(
            storedData?.emailChecksum,
            newEmailChecksum,
            emailChanges.added.length,
            emailChanges.removed.length
          );
          if (!changeType || !changeType.includes('emails')) {
            changeDescriptions.push({ text: `• Emails: ${emailDesc}` });
          }
          // Show specific added emails with clickable links
          for (const e of emailChanges.added.slice(0, 3)) {
            const subjectDisplay = e.subject.substring(0, 50) + (e.subject.length > 50 ? '...' : '');
            const gmailLink = `https://mail.google.com/mail/u/0/#inbox/${e.threadId}`;
            changeDescriptions.push({
              text: `  ➕ "${subjectDisplay}"`,
              link: gmailLink,
              linkText: subjectDisplay
            });
          }
          if (emailChanges.added.length > 3) {
            changeDescriptions.push({ text: `  ... and ${emailChanges.added.length - 3} more added` });
          }
          // Show specific removed emails (no link - they're gone)
          for (const e of emailChanges.removed.slice(0, 3)) {
            const subjectDisplay = e.subject.substring(0, 50) + (e.subject.length > 50 ? '...' : '');
            changeDescriptions.push({ text: `  ➖ "${subjectDisplay}"` });
          }
          if (emailChanges.removed.length > 3) {
            changeDescriptions.push({ text: `  ... and ${emailChanges.removed.length - 3} more removed` });
          }

          // Show specific emails that are no longer overdue
          const oldOverdueEmails = storedData?.moduleChecksums?.overdueEmails || [];
          const newOverdueEmails = newModuleChecksums?.overdueEmails || [];
          const newOverdueIds = new Set(newOverdueEmails.map(e => e.threadId));
          const noLongerOverdue = oldOverdueEmails.filter(e => !newOverdueIds.has(e.threadId));

          for (const e of noLongerOverdue.slice(0, 3)) {
            const subjectDisplay = e.subject.substring(0, 50) + (e.subject.length > 50 ? '...' : '');
            const gmailLink = `https://mail.google.com/mail/u/0/#inbox/${e.threadId}`;
            changeDescriptions.push({
              text: `  ✅ No longer overdue: "${subjectDisplay}"`,
              link: gmailLink,
              linkText: subjectDisplay
            });
          }
          if (noLongerOverdue.length > 3) {
            changeDescriptions.push({ text: `  ... and ${noLongerOverdue.length - 3} more no longer overdue` });
          }
        }

        // Other module changes (skip if already mentioned in changeType to avoid duplicates)
        const ct = (changeType || '').toLowerCase();
        if (changedModules.includes('tasks') && !ct.includes('task')) changeDescriptions.push({ text: '• Tasks updated' });
        if (changedModules.includes('notes') && !ct.includes('note')) changeDescriptions.push({ text: '• Notes changed' });
        if (changedModules.includes('status') && !ct.includes('status')) changeDescriptions.push({ text: '• Status changed' });
        if (changedModules.includes('states') && !ct.includes('state')) changeDescriptions.push({ text: '• States changed' });
        if (changedModules.includes('contacts') && !ct.includes('contact')) changeDescriptions.push({ text: '• Contacts updated' });
        if (changedModules.includes('meetings') && !ct.includes('meeting')) changeDescriptions.push({ text: '• Meetings changed' });
        if (changedModules.includes('contracts') && !ct.includes('contract')) changeDescriptions.push({ text: '• Contracts updated' });
        if (changedModules.includes('helpfulLinks') && !ct.includes('link')) changeDescriptions.push({ text: '• Helpful links changed' });
        if (changedModules.includes('boxDocs') && !ct.includes('box')) changeDescriptions.push({ text: '• Box documents changed' });
        if (changedModules.includes('gDriveFiles') && !ct.includes('drive')) changeDescriptions.push({ text: '• Google Drive files changed' });
      }

      // First view message
      if (changeType === 'First view' || !storedData) {
        changeDescriptions.push({ text: '• First time viewing this vendor' });
      }

      // Render each change description
      if (changeDescriptions.length === 0) {
        changeDescriptions.push({ text: '• Changes detected (details unavailable)' });
      }

      for (const desc of changeDescriptions) {
        const range = bsSh.getRange(changeRow, 6, 1, 4).merge();

        if (desc.link && desc.linkText) {
          // Create rich text with clickable link
          const prefix = desc.text.substring(0, desc.text.indexOf(desc.linkText));
          const suffix = desc.text.substring(desc.text.indexOf(desc.linkText) + desc.linkText.length);
          const richText = SpreadsheetApp.newRichTextValue()
            .setText(desc.text)
            .setLinkUrl(prefix.length, prefix.length + desc.linkText.length, desc.link)
            .build();
          range.setRichTextValue(richText);
        } else {
          range.setValue(desc.text);
        }

        range.setBackground('#fff9e6')  // Lighter yellow
          .setFontColor('#664d03')
          .setHorizontalAlignment('left')
          .setVerticalAlignment('top')
          .setWrap(true);
        changeRow++;
      }

      // Update rightColumnRow
      rightColumnRow = changeRow;
    }

    // Generate vendor-label-only checksum (secondary checksum for catching unlabeled emails)
    const gmailLinkForChecksum = listSh.getRange(listRow, BS_CFG.L_GMAIL_LINK + 1).getValue();
    const vendorLabelChecksum = generateVendorLabelChecksum_(gmailLinkForChecksum);

    // Store the new checksums including module checksums, email data, and vendor label checksum
    storeChecksum_(vendor, newChecksum, newEmailChecksum, newModuleChecksums, emailData, vendorLabelChecksum);
    Logger.log(`Stored checksums for ${vendor}: full=${newChecksum}, vendorLabel=${vendorLabelChecksum || 'null'}`);
  } catch (e) {
    Logger.log(`Error with checksum: ${e.message}`);
  }
  
  ss.toast(`Loaded vendor ${vendorIndex} of ${totalVendors}`, '✅ Ready', 2);
}

/**
 * Generate L2M (Last 2 Months) Reporting permalink for a vendor
 * @param {string} phonexaLink - The Phonexa link from monday.com
 * @param {string} source - The source (Buyers or Affiliates)
 * @returns {object|null} - { url, label } or null if can't generate
 */
function getL2MReportingLink_(phonexaLink, source) {
  if (!phonexaLink) return null;

  const isBuyer = source.toLowerCase().includes('buyer');
  const isAffiliate = source.toLowerCase().includes('affiliate');

  if (!isBuyer && !isAffiliate) return null;

  // Extract the ID from the Phonexa link
  let vendorId = null;

  if (isBuyer) {
    // Buyers: https://cp.profitise.com/p2/buyer/partner/view?id=65
    const match = phonexaLink.match(/[?&]id=(\d+)/);
    if (match) vendorId = match[1];
  } else {
    // Affiliates: https://cp.profitise.com/p2/user/webmaster/view?userId=2160
    const match = phonexaLink.match(/[?&]userId=(\d+)/);
    if (match) vendorId = match[1];
  }

  if (!vendorId) return null;

  // Generate date range: 2 months ago to today
  const today = new Date();
  const twoMonthsAgo = new Date(today);
  twoMonthsAgo.setMonth(twoMonthsAgo.getMonth() - 2);

  const formatDate = (d) => {
    const month = d.getMonth() + 1;
    const day = d.getDate();
    const year = d.getFullYear();
    return `${month}/${day}/${year}`;
  };

  const dateRange = `${formatDate(twoMonthsAgo)}%20-%20${formatDate(today)}`;

  let url;
  if (isBuyer) {
    url = `https://cp.profitise.com/p2/report/summarypartnerbypartner/index?searchForm[pr.productType]=1&searchForm[productId]=&searchForm[date]=${dateRange}&searchForm[partnerId]=${vendorId}`;
  } else {
    url = `https://cp.profitise.com/p2/report/summarypublisherbywm/index?searchForm[pr.productType]=1&searchForm[r.date]=${dateRange}&searchForm[productId]=&searchForm[webmasterId]=${vendorId}`;
  }

  return { url: url, label: 'Link to L2M Reporting' };
}

/**
 * Get helpful links for a specific vendor from monday.com
 */
function getHelpfulLinksForVendor_(vendor, listRow) {
  const ss = SpreadsheetApp.getActive();
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
  
  if (!listSh) return [];
  
  const apiToken = BS_CFG.MONDAY_API_TOKEN;
  const source = listSh.getRange(listRow, BS_CFG.L_SOURCE + 1).getValue() || '';
  
  // Determine if buyer or affiliate
  const isBuyer = source.toLowerCase().includes('buyer');
  const isAffiliate = source.toLowerCase().includes('affiliate');
  
  Logger.log(`=== HELPFUL LINKS SEARCH ===`);
  Logger.log(`Vendor: ${vendor}`);
  Logger.log(`Source: ${source} (isBuyer: ${isBuyer}, isAffiliate: ${isAffiliate})`);
  
  // Query all helpful links with linked_items
  const query = `
    query {
      boards (ids: [${BS_CFG.HELPFUL_LINKS_BOARD_ID}]) {
        items_page (limit: 200) {
          items {
            id
            name
            column_values {
              id
              type
              text
              value
              ... on BoardRelationValue {
                linked_items {
                  id
                  name
                }
              }
            }
          }
        }
      }
    }
  `;
  
  try {
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': apiToken },
      payload: JSON.stringify({ query: query }),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch('https://api.monday.com/v2', options);
    const result = JSON.parse(response.getContentText());
    
    if (!result.data?.boards?.[0]?.items_page?.items) {
      Logger.log('No helpful links found');
      return [];
    }
    
    const allLinks = result.data.boards[0].items_page.items;
    const matchingLinks = [];
    
    for (const item of allLinks) {
      let isMatch = false;
      let linkUrl = '';
      let linkNotes = '';
      
      for (const col of item.column_values) {
        // Check Buyers relation
        if (col.id === BS_CFG.HELPFUL_LINKS_BUYERS_COLUMN && col.linked_items) {
          for (const linkedItem of col.linked_items) {
            if (linkedItem.name.toLowerCase() === vendor.toLowerCase()) {
              isMatch = true;
              break;
            }
          }
        }
        
        // Check Affiliates relation
        if (col.id === BS_CFG.HELPFUL_LINKS_AFFILIATES_COLUMN && col.linked_items) {
          for (const linkedItem of col.linked_items) {
            if (linkedItem.name.toLowerCase() === vendor.toLowerCase()) {
              isMatch = true;
              break;
            }
          }
        }
        
        // Get link URL
        if (col.id === BS_CFG.HELPFUL_LINKS_LINK_COLUMN && col.text) {
          linkUrl = col.text;
        }
        
        // Get notes
        if (col.id === BS_CFG.HELPFUL_LINKS_NOTES_COLUMN && col.text) {
          linkNotes = col.text;
        }
      }
      
      if (isMatch) {
        matchingLinks.push({
          name: item.name,
          url: linkUrl,
          notes: linkNotes
        });
        Logger.log(`Found matching link: ${linkNotes || item.name} -> ${linkUrl}`);
      }
    }
    
    Logger.log(`Returning ${matchingLinks.length} helpful links`);
    return matchingLinks;
    
  } catch (e) {
    Logger.log(`Error fetching helpful links: ${e.message}`);
    return [];
  }
}

/**
 * Get contacts, notes, Phonexa link, LIVE STATUS (from group), and last updated for a vendor from monday.com API
 */
function getVendorContacts_(vendor, listRow) {
  const ss = SpreadsheetApp.getActive();
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
  
  if (!listSh) return { contacts: [], notes: '', phonexaLink: '', lastUpdated: '', liveStatus: '', liveVerticals: '', otherVerticals: '', liveModalities: '', states: '', deadStates: '', otherName: '', mondayItemId: null, boardId: null };
  
  const apiToken = BS_CFG.MONDAY_API_TOKEN;
  const source = listSh.getRange(listRow, BS_CFG.L_SOURCE + 1).getValue() || '';
  
  let boardId, notesColumnId, contactsColumnId, phonexaColumnId, liveVerticalsColumnId, otherVerticalsColumnId, liveModalitiesColumnId, statesColumnId, deadStatesColumnId, otherNameColumnId, isAffiliates;
  
  if (source.toLowerCase().includes('buyer')) {
    boardId = BS_CFG.BUYERS_BOARD_ID;
    notesColumnId = BS_CFG.BUYERS_NOTES_COLUMN;
    contactsColumnId = BS_CFG.BUYERS_CONTACTS_COLUMN;
    phonexaColumnId = BS_CFG.BUYERS_PHONEXA_COLUMN;
    liveVerticalsColumnId = BS_CFG.BUYERS_LIVE_VERTICALS_COLUMN;
    otherVerticalsColumnId = BS_CFG.BUYERS_OTHER_VERTICALS_COLUMN;
    liveModalitiesColumnId = BS_CFG.BUYERS_LIVE_MODALITIES_COLUMN;
    statesColumnId = BS_CFG.BUYERS_STATES_COLUMN;
    deadStatesColumnId = BS_CFG.BUYERS_DEAD_STATES_COLUMN;
    otherNameColumnId = BS_CFG.BUYERS_OTHER_NAME_COLUMN;
    isAffiliates = false;
  } else if (source.toLowerCase().includes('affiliate')) {
    boardId = BS_CFG.AFFILIATES_BOARD_ID;
    notesColumnId = BS_CFG.AFFILIATES_NOTES_COLUMN;
    contactsColumnId = BS_CFG.AFFILIATES_CONTACTS_COLUMN;
    phonexaColumnId = BS_CFG.AFFILIATES_PHONEXA_COLUMN;
    liveVerticalsColumnId = BS_CFG.AFFILIATES_LIVE_VERTICALS_COLUMN;
    otherVerticalsColumnId = BS_CFG.AFFILIATES_OTHER_VERTICALS_COLUMN;
    liveModalitiesColumnId = BS_CFG.AFFILIATES_LIVE_MODALITIES_COLUMN;
    statesColumnId = null; // Affiliates don't have states
    deadStatesColumnId = null; // Affiliates don't have dead states
    otherNameColumnId = BS_CFG.AFFILIATES_OTHER_NAME_COLUMN;
    isAffiliates = true;
  } else {
    boardId = BS_CFG.BUYERS_BOARD_ID;
    notesColumnId = BS_CFG.BUYERS_NOTES_COLUMN;
    contactsColumnId = BS_CFG.BUYERS_CONTACTS_COLUMN;
    phonexaColumnId = BS_CFG.BUYERS_PHONEXA_COLUMN;
    liveVerticalsColumnId = BS_CFG.BUYERS_LIVE_VERTICALS_COLUMN;
    otherVerticalsColumnId = BS_CFG.BUYERS_OTHER_VERTICALS_COLUMN;
    liveModalitiesColumnId = BS_CFG.BUYERS_LIVE_MODALITIES_COLUMN;
    statesColumnId = BS_CFG.BUYERS_STATES_COLUMN;
    deadStatesColumnId = BS_CFG.BUYERS_DEAD_STATES_COLUMN;
    otherNameColumnId = BS_CFG.BUYERS_OTHER_NAME_COLUMN;
    isAffiliates = false;
  }
  
  Logger.log(`=== CONTACTS SEARCH ===`);
  Logger.log(`Vendor: ${vendor}`);
  Logger.log(`Board ID: ${boardId}`);
  Logger.log(`Is Affiliates: ${isAffiliates}`);
  
  const itemId = findMondayItemIdByVendor_(vendor, boardId, apiToken);
  
  if (!itemId) {
    Logger.log('Could not find monday.com item for contacts');
    return { contacts: [], notes: '', phonexaLink: '', lastUpdated: '', liveStatus: '', liveVerticals: '', otherVerticals: '', liveModalities: '', states: '', deadStates: '', otherName: '', mondayItemId: null, boardId: boardId };
  }
  
  let lastUpdated = '';
  
  // Fetch item data including GROUP, updated_at, and column values
  const query = `
    query {
      items (ids: [${itemId}]) {
        updated_at
        group {
          id
          title
        }
        column_values {
          id
          type
          text
          value
          ... on BoardRelationValue {
            linked_items {
              id
              name
            }
          }
        }
      }
    }
  `;
  
  try {
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': apiToken },
      payload: JSON.stringify({ query: query }),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch('https://api.monday.com/v2', options);
    const result = JSON.parse(response.getContentText());
    
    if (!result.data?.items?.[0]) {
      Logger.log('No item data found');
      return { contacts: [], notes: '', phonexaLink: '', lastUpdated: lastUpdated, liveStatus: '', liveVerticals: '', otherVerticals: '', liveModalities: '', states: '', deadStates: '', otherName: '', mondayItemId: itemId, boardId: boardId };
    }
    
    // Get updated_at from the item (ISO 8601 format)
    const updatedAtRaw = result.data.items[0].updated_at || '';
    if (updatedAtRaw) {
      const updatedDate = new Date(updatedAtRaw);
      const tz = 'America/Los_Angeles';
      lastUpdated = Utilities.formatDate(updatedDate, tz, 'MMM d, yyyy h:mm a');
      Logger.log(`Last Updated (from API): ${lastUpdated}`);
    }
    
    // Get status from GROUP title (not a column!)
    const liveStatus = result.data.items[0].group?.title || '';
    Logger.log(`Live Status from Group: ${liveStatus}`);
    
    const columnValues = result.data.items[0].column_values;
    const contactIds = [];
    let notes = '';
    let phonexaLink = '';
    let liveVerticals = '';
    let otherVerticals = '';
    let liveModalities = '';
    let states = '';
    let deadStates = '';
    let otherName = '';
    
    for (const col of columnValues) {
      // Get contacts from linked_items (for 2-way board relations)
      if (col.id === contactsColumnId && col.linked_items && col.linked_items.length > 0) {
        Logger.log(`Found ${col.linked_items.length} linked contacts`);
        for (const linkedItem of col.linked_items) {
          contactIds.push(linkedItem.id);
          Logger.log(`  Contact ID: ${linkedItem.id}, Name: ${linkedItem.name}`);
        }
      }
      // Notes column
      else if (col.id === notesColumnId && col.text) {
        notes = col.text;
      }
      // Phonexa Link column
      else if (col.id === phonexaColumnId && col.text) {
        phonexaLink = col.text;
        Logger.log(`Phonexa Link: ${phonexaLink}`);
      }
      // Live Verticals column (tags)
      else if (col.id === liveVerticalsColumnId && col.text) {
        liveVerticals = col.text;
        Logger.log(`Live Verticals: ${liveVerticals}`);
      }
      // Other Verticals column (tags)
      else if (col.id === otherVerticalsColumnId && col.text) {
        otherVerticals = col.text;
        Logger.log(`Other Verticals: ${otherVerticals}`);
      }
      // Live Modalities column (tags)
      else if (col.id === liveModalitiesColumnId && col.text) {
        liveModalities = col.text;
        Logger.log(`Live Modalities: ${liveModalities}`);
      }
      // States column (dropdown - Buyers only)
      else if (statesColumnId && col.id === statesColumnId && col.text) {
        states = col.text;
        Logger.log(`States: ${states}`);
      }
      // Dead States column (dropdown - Buyers only)
      else if (deadStatesColumnId && col.id === deadStatesColumnId && col.text) {
        deadStates = col.text;
        Logger.log(`Dead States: ${deadStates}`);
      }
      // Other Name column (text - Buyers only)
      else if (otherNameColumnId && col.id === otherNameColumnId && col.text) {
        otherName = col.text;
        Logger.log(`Other Name: ${otherName}`);
      }
    }
    
    Logger.log(`Found ${contactIds.length} contact IDs`);
    Logger.log(`Notes: ${notes}`);
    
    const contacts = [];
    
    if (contactIds.length > 0) {
      const contactsQuery = `
        query {
          items (ids: [${contactIds.join(', ')}]) {
            id
            name
            column_values {
              id
              text
              type
            }
          }
        }
      `;
      
      const contactsOptions = {
        method: 'post',
        contentType: 'application/json',
        headers: { 'Authorization': apiToken },
        payload: JSON.stringify({ query: contactsQuery }),
        muteHttpExceptions: true
      };
      
      const contactsResponse = UrlFetchApp.fetch('https://api.monday.com/v2', contactsOptions);
      const contactsResult = JSON.parse(contactsResponse.getContentText());
      
      if (contactsResult.data?.items) {
        for (const item of contactsResult.data.items) {
          const contact = {
            name: item.name,
            email: '',
            phone: '',
            status: '',
            contactType: ''
          };
          
          for (const col of item.column_values) {
            if (col.id === 'email_mkrk53z4') {
              contact.email = col.text || '';
            } else if (col.id === 'phone_mkrkzxq2') {
              contact.phone = col.text || '';
            } else if (col.id === 'status') {
              contact.status = col.text || '';
            } else if (col.id === 'color_mkrkh4bk') {
              contact.contactType = col.text || '';
            }
          }
          
          contacts.push(contact);
          Logger.log(`Contact found: ${contact.name} <${contact.email}> | ${contact.phone} | ${contact.status} | ${contact.contactType}`);
        }
      }
    }
    
    Logger.log(`Returning ${contacts.length} contacts`);
    return { contacts: contacts, notes: notes, phonexaLink: phonexaLink, lastUpdated: lastUpdated, liveStatus: liveStatus, liveVerticals: liveVerticals, otherVerticals: otherVerticals, liveModalities: liveModalities, states: states, deadStates: deadStates, otherName: otherName, mondayItemId: itemId, boardId: boardId };
    
  } catch (e) {
    Logger.log(`Error fetching contacts: ${e.message}`);
    return { contacts: [], notes: '', phonexaLink: '', lastUpdated: lastUpdated, liveStatus: '', liveVerticals: '', otherVerticals: '', liveModalities: '', states: '', deadStates: '', otherName: '', mondayItemId: itemId, boardId: boardId };
  }
}

/**
 * Get emails for a specific vendor by searching Gmail live
 */
function getEmailsForVendor_(vendor, listRow) {
  const ss = SpreadsheetApp.getActive();
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
  
  if (!listSh) {
    Logger.log('List sheet not found');
    return [];
  }
  
  try {
    const gmailLinkAll = listSh.getRange(listRow, BS_CFG.L_GMAIL_LINK + 1).getValue();
    const gmailLinkNoSnooze = listSh.getRange(listRow, BS_CFG.L_NO_SNOOZE + 1).getValue();
    
    Logger.log(`=== GMAIL SEARCH ===`);
    Logger.log(`Vendor: ${vendor}`);
    Logger.log(`Gmail Link (All): ${gmailLinkAll}`);
    Logger.log(`Gmail Link (No Snooze): ${gmailLinkNoSnooze}`);
    
    const allEmails = [];
    
    if (gmailLinkAll && gmailLinkAll.toString().includes('#search')) {
      const threadsAll = searchGmailFromLink_(gmailLinkAll, 'All');
      Logger.log(`Found ${threadsAll.length} threads in "All" query`);
      allEmails.push(...threadsAll);
    }
    
    const noSnoozeThreadIds = new Set();
    if (gmailLinkNoSnooze && gmailLinkNoSnooze.toString().includes('#search')) {
      const threadsNoSnooze = searchGmailFromLink_(gmailLinkNoSnooze, 'No Snooze');
      Logger.log(`Found ${threadsNoSnooze.length} threads in "No Snooze" query`);
      
      for (const email of threadsNoSnooze) {
        noSnoozeThreadIds.add(email.threadId);
      }
    }
    
    const uniqueEmails = [];
    const seenThreadIds = new Set();
    
    for (const email of allEmails) {
      if (!seenThreadIds.has(email.threadId)) {
        seenThreadIds.add(email.threadId);
        email.isSnoozed = !noSnoozeThreadIds.has(email.threadId);
        uniqueEmails.push(email);
      }
    }
    
    uniqueEmails.sort((a, b) => new Date(b.date) - new Date(a.date));
    
    Logger.log(`Returning ${uniqueEmails.length} total emails (${uniqueEmails.filter(e => e.isSnoozed).length} snoozed)`);
    return uniqueEmails;
    
  } catch (e) {
    Logger.log(`ERROR in getEmailsForVendor_: ${e.message}`);
    return [];
  }
}

/**
 * FAST: Get just email thread count without fetching message content
 * Used for quick change detection in Skip Unchanged
 * @param {number} listRow - Row number in List sheet
 * @returns {number} Thread count, or -1 on error
 */
function getEmailThreadCountFast_(listRow) {
  const ss = SpreadsheetApp.getActive();
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);

  if (!listSh) return -1;

  try {
    const gmailLinkAll = listSh.getRange(listRow, BS_CFG.L_GMAIL_LINK + 1).getValue();
    if (!gmailLinkAll || !gmailLinkAll.toString().includes('#search')) return 0;

    const gmailLinkStr = gmailLinkAll.toString();
    const urlParts = gmailLinkStr.split('#search/');
    if (urlParts.length < 2) return 0;

    let searchQuery = decodeURIComponent(urlParts[1]);
    if (searchQuery.includes('?')) searchQuery = searchQuery.split('?')[0];
    searchQuery = searchQuery.replace(/\+/g, ' ').replace(/-is:snoozed/gi, '-label:snoozed');

    // Just get thread count - don't fetch messages (fast!)
    const threads = GmailApp.search(searchQuery, 0, 50);
    return threads.length;
  } catch (e) {
    Logger.log(`ERROR in getEmailThreadCountFast_: ${e.message}`);
    return -1;
  }
}

/**
 * Get ALL emails from vendor's Gmail sublabel for task analysis context
 * Uses label search (not filtered search) to get complete email history
 * @param {number} listRow - The row number in the list sheet
 * @param {number} maxThreads - Maximum threads to fetch (default 50)
 * @returns {Array} - Array of email objects with full context
 */
function getAllEmailsFromVendorLabel_(listRow, maxThreads = 50) {
  const ss = SpreadsheetApp.getActive();
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);

  if (!listSh) return [];

  try {
    // Get the Gmail link to extract the vendor label
    const gmailLinkAll = listSh.getRange(listRow, BS_CFG.L_GMAIL_LINK + 1).getValue();
    if (!gmailLinkAll) return [];

    const gmailLinkStr = gmailLinkAll.toString();

    // Extract vendor label (zzzvendors-*) from URL
    // Format: label%3Azzzvendors-vendorname or label:zzzvendors-vendorname
    // Include periods, underscores, hyphens in vendor name (e.g., "inc." in "american-remodeling-enterprises-inc.")
    const vendorLabelMatch = gmailLinkStr.match(/label[:%]3A(zzzvendors-[a-z0-9_.\-]+)/i);
    if (!vendorLabelMatch) {
      Logger.log('Could not extract vendor label from Gmail link');
      return [];
    }

    const vendorLabel = vendorLabelMatch[1];
    // Convert dash format to slash format for GmailApp: zzzvendors-name -> zzzVendors/Name
    const labelPath = vendorLabel.replace('zzzvendors-', 'zzzVendors/').replace(/-/g, ' ');

    Logger.log(`Searching vendor label: ${labelPath}`);

    // Search by label only - no filters, get ALL emails
    const searchQuery = `label:${vendorLabel}`;
    const threads = GmailApp.search(searchQuery, 0, maxThreads);

    Logger.log(`Found ${threads.length} threads in vendor label`);

    const emails = [];
    for (const thread of threads) {
      const messages = thread.getMessages();
      const lastMessage = messages[messages.length - 1];
      const firstMessage = messages[0];

      // Get labels for this thread
      const labels = thread.getLabels().map(l => l.getName());

      emails.push({
        threadId: thread.getId(),
        subject: thread.getFirstMessageSubject(),
        date: lastMessage.getDate().toISOString().split('T')[0],
        from: lastMessage.getFrom(),
        to: lastMessage.getTo(),
        snippet: lastMessage.getPlainBody().substring(0, 1500),
        messageCount: messages.length,
        labels: labels,
        isUnread: thread.isUnread(),
        lastMessageDate: lastMessage.getDate()
      });
    }

    // Sort by date descending
    emails.sort((a, b) => b.lastMessageDate - a.lastMessageDate);

    return emails;

  } catch (e) {
    Logger.log(`ERROR in getAllEmailsFromVendorLabel_: ${e.message}`);
    return [];
  }
}

/**
 * Helper function to search Gmail from a link
 */
function searchGmailFromLink_(gmailLink, querySetName) {
  try {
    const gmailLinkStr = gmailLink.toString();
    const urlParts = gmailLinkStr.split('#search/');
    
    if (urlParts.length < 2) {
      Logger.log(`Could not parse Gmail search URL for ${querySetName}`);
      return [];
    }
    
    let searchQuery = decodeURIComponent(urlParts[1]);
    
    if (searchQuery.includes('?')) {
      searchQuery = searchQuery.split('?')[0];
    }
    
    searchQuery = searchQuery.replace(/\+/g, ' ');
    searchQuery = searchQuery.replace(/-is:snoozed/gi, '-label:snoozed');
    
    Logger.log(`${querySetName} search query: ${searchQuery}`);
    
    const threads = GmailApp.search(searchQuery, 0, 50);
    Logger.log(`${querySetName} found ${threads.length} threads`);
    
    const emails = [];
    
    for (const thread of threads) {
      const messages = thread.getMessages();
      if (messages.length === 0) continue;

      // Find the last message from a real person (skip bounces/system messages)
      let lastMessage = messages[messages.length - 1];
      let lastSender = lastMessage.getFrom().toLowerCase();

      // Check if last message is a bounce/system message, if so find the previous real message
      if (isSystemOrBounceEmail_(lastSender)) {
        for (let i = messages.length - 2; i >= 0; i--) {
          const prevSender = messages[i].getFrom().toLowerCase();
          if (!isSystemOrBounceEmail_(prevSender)) {
            lastMessage = messages[i];
            lastSender = prevSender;
            break;
          }
        }
      }

      const subject = thread.getFirstMessageSubject();
      const date = lastMessage.getDate(); // Use last real message date

      // Determine who sent the last message
      const myEmail = Session.getActiveUser().getEmail().toLowerCase();
      let lastFrom = 'VENDOR';
      if (lastSender.includes(myEmail)) {
        lastFrom = 'ME';
      } else if (lastSender.includes('aden')) {
        lastFrom = 'ADEN';
      } else if (lastSender.includes('tina@zeroparallel.com')) {
        lastFrom = 'LEGAL';
      } else if (lastSender.includes('accounting')) {
        lastFrom = 'ACCOUNTING';
      } else if (lastSender.includes('@zeroparallel.com') || lastSender.includes('@profitise.com') || lastSender.includes('@phonexa.com')) {
        lastFrom = 'INTERNAL';
      }

      // Get labels - include INBOX if thread is in inbox (system labels not returned by getLabels)
      const userLabels = thread.getLabels().map(label => label.getName());
      if (thread.isInInbox()) {
        userLabels.push('INBOX');
      }
      const labels = userLabels.sort().join(', ');
      const threadId = thread.getId();
      const threadLink = `https://mail.google.com/mail/u/0/#inbox/${threadId}`;
      
      let snippet = '';
      try {
        snippet = lastMessage.getPlainBody().substring(0, 200) + '...';
      } catch (e) {
        snippet = '(unable to load snippet)';
      }
      
      // Format date in Pacific timezone with leading zeros
      const tz = 'America/Los_Angeles';
      const dateFormatted = Utilities.formatDate(date, tz, 'yyyy-MM-dd HH:mm');
      
      emails.push({
        threadId: threadId,
        subject: subject || '(no subject)',
        date: dateFormatted,
        lastFrom: lastFrom,
        count: messages.length,  // Keep for Claude analysis context
        labels: labels,
        link: threadLink,
        querySet: querySetName,
        snippet: snippet,
        isSnoozed: false
      });
    }
    
    return emails;
    
  } catch (e) {
    Logger.log(`Error searching Gmail for ${querySetName}: ${e.message}`);
    return [];
  }
}

/**
 * Get tasks for a specific vendor from monday.com Tasks board
 */
function getTasksForVendor_(vendor, listRow) {
  const apiToken = BS_CFG.MONDAY_API_TOKEN;
  const tasksBoardId = BS_CFG.TASKS_BOARD_ID;
  
  Logger.log(`=== MONDAY.COM TASKS SEARCH ===`);
  Logger.log(`Vendor: ${vendor}`);
  
  const escapedVendor = vendor.replace(/"/g, '\\"');
  
  const query = `
    query {
      boards (ids: [${tasksBoardId}]) {
        items_page (
          limit: 100
          query_params: {
            rules: [
              {
                column_id: "name"
                compare_value: ["${escapedVendor}"]
                operator: contains_text
              }
            ]
          }
        ) {
          items {
            id
            name
            group {
              id
              title
            }
            column_values {
              id
              text
              type
              ... on BoardRelationValue {
                linked_items {
                  id
                  name
                }
              }
            }
            created_at
            updated_at
          }
        }
      }
    }
  `;
  
  try {
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': apiToken },
      payload: JSON.stringify({ query: query }),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch('https://api.monday.com/v2', options);
    const result = JSON.parse(response.getContentText());
    
    if (result.errors && result.errors.length > 0) {
      Logger.log(`API Error: ${result.errors[0].message}`);
      return [];
    }
    
    if (!result.data?.boards?.[0]?.items_page?.items) {
      Logger.log('No items found in Tasks board');
      return [];
    }
    
    const items = result.data.boards[0].items_page.items;
    Logger.log(`Found ${items.length} tasks matching vendor`);
    
    const tasks = [];
    
    // Group priority for sorting (DESC = higher number first)
    const groupPriority = {
      'topics': 3,                    // Ongoing Projects
      'group_mkqb5pzw': 2,           // Upcoming/Paused Projects
      'group_title': 1,              // Completed Projects
      'group_mkqf4yzy': 0            // Task Templates
    };
    
    // Explicit project order (from Large-Scale Projects board structure)
    const projectPriority = {
      'Home Services': 1,
      'ACA': 2,
      'Vertical Activation': 3,
      'Monthly Returns': 4,
      'CPL/Zip Optimizations': 5,
      'Accounting/Invoices': 6,
      'System Admin': 7,
      'URL Whitelist': 8,
      'Outbound Communication': 9,
      'Pre-Onboarding': 10,
      'Appointments': 11,
      'Onboarding - Buyer': 12,
      'Onboarding - Affiliate': 13,
      'Onboarding - Vertical': 14,
      'Templates': 15,
      'Morning Meeting': 16,
      'Week of 07/28/25': 17
    };
    
    for (const item of items) {
      const itemName = String(item.name || '');
      const groupId = item.group?.id || '';
      const groupTitle = item.group?.title || '';
      
      let status = '';
      let projectName = '';
      let tempInd = null;
      let taskDate = '';  // Date from timeline column
      let statusColumnId = 'status';  // Default, will be updated if found
      let lastUpdated = '';  // Last Updated column

      for (const col of item.column_values) {
        // Status column
        if (col.id === 'status' || col.id === 'status4' || col.id === 'status_1') {
          status = col.text || '';
          statusColumnId = col.id;  // Remember which column has the status
        }
        // Project column (board relation)
        else if (col.id === BS_CFG.TASKS_PROJECT_COLUMN) {
          if (col.linked_items && col.linked_items.length > 0) {
            projectName = col.linked_items.map(li => li.name).join(', ');
          } else if (col.text) {
            projectName = col.text;
          }
        }
        // temp_ind column
        else if (col.id === 'numeric_mkt9pbj0') {
          if (col.text && col.text.trim()) {
            tempInd = parseFloat(col.text);
          }
        }
        // Date column (timeline type - shows as "YYYY-MM-DD - YYYY-MM-DD" or just end date)
        else if (col.id === 'timerange_mkqfmzyf') {
          if (col.text && col.text.trim()) {
            // Timeline format is "Start - End", we want the end date
            const dateParts = col.text.split(' - ');
            taskDate = dateParts.length > 1 ? dateParts[1].trim() : dateParts[0].trim();
          }
        }
        // Last Updated column
        else if (col.id === 'pulse_updated_mkyyesw') {
          if (col.text && col.text.trim()) {
            // Format is typically "YYYY-MM-DD HH:MM:SS" or similar, extract just the date
            const datePart = col.text.split(' ')[0];
            lastUpdated = datePart;
          }
        }
      }
      
      const createdDate = new Date(item.created_at);
      const tz = 'America/Los_Angeles';
      const createdFormatted = Utilities.formatDate(createdDate, tz, 'yyyy-MM-dd HH:mm');
      
      tasks.push({
        itemId: item.id,  // Monday.com item ID for updates
        statusColumnId: statusColumnId,  // Status column ID for updates
        subject: itemName,
        status: status || 'No Status',
        taskDate: taskDate,  // Date from timeline column
        lastUpdated: lastUpdated,  // Last Updated column
        created: createdFormatted,
        project: projectName || `Group: ${groupTitle}`,
        isDone: (status.toLowerCase() === 'done'),

        // Sorting fields
        groupId: groupId,
        groupPriority: groupPriority[groupId] || -1,
        projectName: projectName,
        projectPriority: projectPriority[projectName] || 999,
        tempInd: tempInd,
        tempIndSort: (tempInd === null) ? 999999 : tempInd  // Blanks at end
      });
      
      Logger.log(`Task: ${itemName} | Group: ${groupTitle} | Project: ${projectName} | temp_ind: ${tempInd}`);
    }
    
    // Sort by: Done status (not done first)
    // For NOT DONE: Group (DESC) -> Project (explicit order) -> temp_ind (ASC, blanks last) -> Created Date (DESC)
    // For DONE: Created Date (DESC) only
    tasks.sort((a, b) => {
      // 0. Done status (not done before done)
      if (a.isDone !== b.isDone) {
        return a.isDone ? 1 : -1;
      }
      
      // If BOTH are done, sort only by Created Date DESC
      if (a.isDone && b.isDone) {
        const dateA = a.created ? new Date(a.created.replace(' ', 'T')) : new Date(0);
        const dateB = b.created ? new Date(b.created.replace(' ', 'T')) : new Date(0);
        return dateB - dateA;
      }
      
      // For not-done tasks, use full sorting logic:
      // 1. Group priority (DESC - higher numbers first)
      if (a.groupPriority !== b.groupPriority) {
        return b.groupPriority - a.groupPriority;
      }
      
      // 2. Project (explicit order from Large-Scale Projects)
      if (a.projectPriority !== b.projectPriority) {
        return a.projectPriority - b.projectPriority;
      }
      
      // 3. temp_ind (ASC - low to high, blanks at end)
      if (a.tempIndSort !== b.tempIndSort) {
        return a.tempIndSort - b.tempIndSort;
      }
      
      // 4. Created Date (DESC - newest first)
      const dateA = a.created ? new Date(a.created.replace(' ', 'T')) : new Date(0);
      const dateB = b.created ? new Date(b.created.replace(' ', 'T')) : new Date(0);
      return dateB - dateA;
    });
    
    return tasks.slice(0, 30);
    
  } catch (e) {
    Logger.log(`Error fetching monday.com tasks: ${e.message}`);
    return [];
  }
}

/**
 * Get upcoming calendar meetings for a vendor
 * Searches Google Calendar for events containing the vendor name
 * Searches in: title, description/notes, location, and attendee emails
 * Returns meetings from 30 days ago to 60 days in the future
 */
function getUpcomingMeetingsForVendor_(vendor, contactEmails) {
  Logger.log(`=== CALENDAR SEARCH ===`);
  Logger.log(`Vendor: ${vendor}`);
  if (contactEmails && contactEmails.length > 0) {
    Logger.log(`Contact emails: ${contactEmails.join(', ')}`);
  }
  
  try {
    const now = new Date();
    // Force Pacific timezone for consistent date comparison
    const tz = 'America/Los_Angeles';
    const todayStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    Logger.log(`Today's date (${tz}): ${todayStr}`);
    Logger.log(`Current time UTC: ${now.toISOString()}`);
    
    const pastDate = new Date(now.getTime() - (30 * 24 * 60 * 60 * 1000)); // 30 days ago
    const futureDate = new Date(now.getTime() + (60 * 24 * 60 * 60 * 1000)); // 60 days ahead
    
    const calendars = CalendarApp.getAllCalendars();
    const meetings = [];
    
    // Only search calendars owned by the user (not shared calendars)
    const ownedCalendars = calendars.filter(cal => {
      try {
        // Primary calendar or calendars where user is owner
        return cal.isOwnedByMe();
      } catch (e) {
        return false;
      }
    });
    
    Logger.log(`Searching ${ownedCalendars.length} owned calendars (filtered from ${calendars.length} total)`);
    
    // Search terms - vendor name and variations
    const searchTerms = [vendor.toLowerCase()];
    
    // Also try without parentheses
    const withoutParens = vendor.replace(/\s*\([^)]*\)/g, '').trim().toLowerCase();
    if (withoutParens !== vendor.toLowerCase()) {
      searchTerms.push(withoutParens);
    }
    
    // Also try first word only (company name) - but only if vendor is a single word and 4+ chars
    const words = vendor.split(/[\s\-\(\)]+/).filter(w => w.length > 0);
    if (words.length === 1) {
      const firstWord = words[0].toLowerCase();
      if (firstWord.length >= 4 && !searchTerms.includes(firstWord)) {
        searchTerms.push(firstWord);
      }
    }
    
    // Also try without common suffixes like LLC, Inc, Corp
    const withoutSuffix = vendor.replace(/,?\s*(LLC|Inc\.?|Corp\.?|Corporation|Company|Co\.?)$/i, '').trim().toLowerCase();
    if (withoutSuffix !== vendor.toLowerCase() && !searchTerms.includes(withoutSuffix)) {
      searchTerms.push(withoutSuffix);
    }
    
    // Add contact emails as search terms (for matching attendees)
    const emailSearchTerms = [];
    if (contactEmails && contactEmails.length > 0) {
      for (const email of contactEmails) {
        if (email && email.includes('@')) {
          const emailLower = email.toLowerCase().trim();
          if (!emailSearchTerms.includes(emailLower)) {
            emailSearchTerms.push(emailLower);
          }
        }
      }
    }
    
    // For multi-word vendor names, require at least 2 consecutive words to match
    // This prevents "Solar Energy World" from matching just "Solar"
    const isMultiWord = words.length >= 2;
    
    Logger.log(`Search terms: ${searchTerms.join(', ')} (multi-word: ${isMultiWord})`);
    if (emailSearchTerms.length > 0) {
      Logger.log(`Email search terms: ${emailSearchTerms.join(', ')}`);
    }
    
    for (const calendar of ownedCalendars) {
      try {
        const events = calendar.getEvents(pastDate, futureDate);
        
        for (const event of events) {
          const title = event.getTitle() || '';
          const description = event.getDescription() || '';
          const location = event.getLocation() || '';
          
          // Get attendee emails to search
          let attendeeEmails = '';
          try {
            const guests = event.getGuestList();
            attendeeEmails = guests.map(g => g.getEmail()).join(' ');
          } catch (e) {
            // Some events may not allow guest list access
          }
          
          // Combine all searchable text: title, description/notes, location, attendee emails
          const searchText = (title + ' ' + description + ' ' + location + ' ' + attendeeEmails).toLowerCase();
          
          // Check if any search term matches
          let isMatch = false;
          let matchedTerm = '';
          let matchedIn = '';
          
          // First check vendor name search terms against all searchable text
          for (const term of searchTerms) {
            if (searchText.includes(term)) {
              // For multi-word vendors, require at least 2 words from vendor name to match
              // Skip single-word matches like just "solar" for "Solar Energy World"
              if (isMultiWord) {
                const termWords = term.split(/\s+/);
                if (termWords.length < 2) {
                  // Single word term - skip unless it's a unique identifier (not common words)
                  const commonWords = ['solar', 'energy', 'home', 'power', 'green', 'sun', 'electric', 'services', 'solutions', 'group', 'pro', 'usa', 'national'];
                  if (commonWords.includes(term)) {
                    continue; // Skip common single words for multi-word vendors
                  }
                }
              }
              isMatch = true;
              matchedTerm = term;
              matchedIn = title.toLowerCase().includes(term) ? 'title' : 
                         description.toLowerCase().includes(term) ? 'notes' :
                         location.toLowerCase().includes(term) ? 'location' : 'attendees';
              break;
            }
          }
          
          // If no match yet, check contact emails against attendee emails
          if (!isMatch && emailSearchTerms.length > 0 && attendeeEmails) {
            const attendeeEmailsLower = attendeeEmails.toLowerCase();
            for (const email of emailSearchTerms) {
              if (attendeeEmailsLower.includes(email)) {
                isMatch = true;
                matchedTerm = email;
                matchedIn = 'contact email';
                break;
              }
            }
          }
          
          if (isMatch) {
            const startTime = event.getStartTime();
            const endTime = event.getEndTime();
            const isAllDay = event.isAllDayEvent();
            
            // Determine if past, today, or future using script timezone
            const eventStr = Utilities.formatDate(startTime, tz, 'yyyy-MM-dd');
            
            const isPast = startTime < now;
            const isToday = (eventStr === todayStr);
            
            Logger.log(`Event: ${event.getTitle()} | Date: ${eventStr} | Today: ${todayStr} | isToday: ${isToday}`);
            
            // Determine status
            let status = '';
            if (isPast) {
              status = '✓ Past';
            } else if (isToday) {
              status = '🔴 Today';
            } else {
              // Calculate days until using date strings to avoid timezone issues
              const eventDate = new Date(eventStr + 'T00:00:00');
              const todayDate = new Date(todayStr + 'T00:00:00');
              const daysUntil = Math.round((eventDate - todayDate) / (24 * 60 * 60 * 1000));
              status = `📅 In ${daysUntil} day${daysUntil !== 1 ? 's' : ''}`;
            }
            
            // Note where the match was found (use matchedIn from earlier detection)
            const matchSource = matchedIn || 'unknown';
            
            // Format times in Pacific timezone (use tz already defined above)
            const startHH = Utilities.formatDate(startTime, tz, 'HH');
            const startMM = Utilities.formatDate(startTime, tz, 'mm');
            const endHH = Utilities.formatDate(endTime, tz, 'HH');
            const endMM = Utilities.formatDate(endTime, tz, 'mm');
            const timeFormatted = isAllDay ? 'All Day' : `${startHH}:${startMM} - ${endHH}:${endMM}`;
            
            meetings.push({
              title: title,
              date: Utilities.formatDate(startTime, tz, 'yyyy-MM-dd'),
              time: timeFormatted,
              status: status,
              isPast: isPast,
              isToday: isToday,
              startTime: startTime,
              matchSource: matchSource,
              link: `https://calendar.google.com/calendar/event?eid=${Utilities.base64Encode(event.getId().split('@')[0] + ' ' + calendar.getId())}`
            });
            
            Logger.log(`Found meeting: ${title} on ${Utilities.formatDate(startTime, tz, 'yyyy-MM-dd HH:mm')} (matched in ${matchSource})`);
          }
        }
      } catch (calError) {
        Logger.log(`Error searching calendar ${calendar.getName()}: ${calError.message}`);
      }
    }
    
    // Sort by date (upcoming first, then past)
    meetings.sort((a, b) => {
      // Today first
      if (a.isToday && !b.isToday) return -1;
      if (!a.isToday && b.isToday) return 1;
      
      // Future before past
      if (!a.isPast && b.isPast) return -1;
      if (a.isPast && !b.isPast) return 1;
      
      // Within same category, sort by date
      return a.startTime - b.startTime;
    });
    
    // Remove duplicates - only keep the FIRST instance of each meeting title
    // Also filter out past "checkup" meetings
    // For "Day ##" pattern meetings, dedupe based on title without the day number
    const seenTitles = new Set();
    const uniqueMeetings = meetings.filter(m => {
      // Skip past events with "checkup" in the name
      if (m.isPast && m.title.toLowerCase().includes('checkup')) {
        return false;
      }
      
      // Normalize title for deduplication:
      // Remove "Day ##" or "Day #" patterns (e.g., "Solar Energy World - FL (zips) - $51 Checkup - Day 3" -> "Solar Energy World - FL (zips) - $51 Checkup")
      const normalizedTitle = m.title
        .replace(/\s*-?\s*Day\s*\d+\s*$/i, '')  // Remove " - Day 3" or " Day 14" at end
        .replace(/\s*-?\s*Day\s*\d+\s*-/i, ' - ')  // Remove "Day 3 -" in middle
        .trim();
      
      if (seenTitles.has(normalizedTitle)) return false;
      seenTitles.add(normalizedTitle);
      return true;
    });
    
    // Return both unique meetings and total count for "more" link
    Logger.log(`Found ${uniqueMeetings.length} unique meetings for ${vendor} (${meetings.length} total)`);
    return { meetings: uniqueMeetings, totalCount: meetings.length };
    
  } catch (e) {
    Logger.log(`Error fetching calendar meetings: ${e.message}`);
    return { meetings: [], totalCount: 0 };
  }
}

/************************************************************
 * AIRTABLE CONTRACTS FUNCTIONS
 * Polls Airtable API to get Contract records linked to vendors
 ************************************************************/

/**
 * Fuzzy match vendor names (handles variations)
 */
function isFuzzyMatch_(name1, name2) {
  if (!name1 || !name2) return false;
  
  // Normalize both names
  const n1 = name1.toLowerCase().replace(/[^a-z0-9]/g, '');
  const n2 = name2.toLowerCase().replace(/[^a-z0-9]/g, '');
  
  // Exact match after normalization
  if (n1 === n2) return true;
  
  // One contains the other
  if (n1.includes(n2) || n2.includes(n1)) return true;
  
  // Check if significant portion matches (at least 50% of shorter string in longer)
  const shorter = n1.length < n2.length ? n1 : n2;
  const longer = n1.length < n2.length ? n2 : n1;
  
  if (longer.includes(shorter) && shorter.length >= longer.length * 0.5) {
    return true;
  }
  
  return false;
}

/**
 * Get all contracts from Airtable (with pagination)
 * Fetches from both Contracts 2025 and Contracts 2024 tables
 */
function getAllContracts_(useCache) {
  // Check cache first if requested
  if (useCache !== false) {
    const cached = getCachedData_('airtable', 'all_contracts');
    if (cached) {
      return cached;
    }
  }
  
  const allRecords = [];
  
  // Fetch from all contract tables
  const tables = [
    { name: 'Contracts 2026', tableId: BS_CFG.AIRTABLE_CONTRACTS_TABLE_ID_2026, viewId: BS_CFG.AIRTABLE_CONTRACTS_VIEW_ID_2026 },
    { name: 'Contracts 2025', tableId: BS_CFG.AIRTABLE_CONTRACTS_TABLE_ID_2025, viewId: BS_CFG.AIRTABLE_CONTRACTS_VIEW_ID_2025 },
    { name: 'Contracts 2024', tableId: BS_CFG.AIRTABLE_CONTRACTS_TABLE_ID_2024, viewId: BS_CFG.AIRTABLE_CONTRACTS_VIEW_ID_2024 }
  ];
  
  for (const table of tables) {
    let offset = null;
    
    do {
      let url = `${BS_CFG.AIRTABLE_API_BASE_URL}/${BS_CFG.AIRTABLE_BASE_ID}/${table.tableId}?maxRecords=100`;
      
      if (offset) {
        url += `&offset=${offset}`;
      }
      
      const options = {
        method: 'get',
        headers: {
          'Authorization': `Bearer ${BS_CFG.AIRTABLE_API_TOKEN}`,
          'Content-Type': 'application/json'
        },
        muteHttpExceptions: true
      };
      
      try {
        const response = UrlFetchApp.fetch(url, options);
        const responseCode = response.getResponseCode();
        
        if (responseCode !== 200) {
          Logger.log(`Airtable API error for ${table.name}: ${response.getContentText()}`);
          break;
        }
        
        const data = JSON.parse(response.getContentText());
        
        if (data.records) {
          // Add table info to each record for URL construction
          data.records.forEach(record => {
            record._tableId = table.tableId;
            record._viewId = table.viewId;
            record._tableName = table.name;
          });
          allRecords.push(...data.records);
        }
        
        offset = data.offset || null;
        
        // Respect rate limits (5 requests/second)
        if (offset) {
          Utilities.sleep(250);
        }
      } catch (e) {
        Logger.log(`Error fetching Airtable contracts from ${table.name}: ${e.message}`);
        break;
      }
      
    } while (offset);
  }
  
  Logger.log(`📄 Retrieved ${allRecords.length} total contracts from Airtable (2024 + 2025)`);
  
  // Cache the results
  setCachedData_('airtable', 'all_contracts', allRecords);
  
  return allRecords;
}

/**
 * Get contracts for a specific vendor using fuzzy matching
 * Searches both Vendor Name field AND Notes field for matches
 */
function getContractsForVendor_(vendorName) {
  if (!vendorName) {
    Logger.log('⚠️ No vendor name provided');
    return [];
  }
  
  // Get all contracts and filter locally for fuzzy matching
  const allContracts = getAllContracts_();
  
  Logger.log(`📋 Filtering ${allContracts.length} total contracts for vendor: ${vendorName}`);
  
  const matches = allContracts.filter(record => {
    const contractVendor = record.fields[BS_CFG.AIRTABLE_VENDOR_FIELD] || '';
    const contractNotes = record.fields[BS_CFG.AIRTABLE_NOTES_FIELD] || '';
    
    // Handle Submitted By - could be a user/collaborator field (object with email/name) or linked record
    let submittedBy = '';
    const submittedByRaw = record.fields[BS_CFG.AIRTABLE_SUBMITTED_BY_FIELD];
    if (submittedByRaw) {
      if (typeof submittedByRaw === 'string') {
        submittedBy = submittedByRaw;
      } else if (submittedByRaw.name) {
        // Collaborator field format: {id, email, name}
        submittedBy = submittedByRaw.name;
      } else if (submittedByRaw.email) {
        submittedBy = submittedByRaw.email;
      } else if (Array.isArray(submittedByRaw)) {
        // Could be array of collaborators
        submittedBy = submittedByRaw.map(s => s.name || s.email || String(s)).join(', ');
      } else {
        submittedBy = JSON.stringify(submittedByRaw);
        Logger.log(`📋 Unexpected Submitted By format: ${submittedBy}`);
      }
    }
    
    const vertical = record.fields[BS_CFG.AIRTABLE_VERTICAL_FIELD] || '';
    
    // First check if this contract matches the vendor name at all
    let vendorMatches = false;
    
    if (isFuzzyMatch_(vendorName, contractVendor)) {
      vendorMatches = true;
    } else if (contractNotes && contractNotes.toLowerCase().includes(vendorName.toLowerCase())) {
      vendorMatches = true;
    } else {
      // Try simplified name matching
      const vendorSimple = vendorName.toLowerCase()
        .replace(/\s*\([^)]*\)/g, '')
        .replace(/,?\s*(LLC|Inc\.?|Corp\.?|Corporation|Company|Co\.?)$/i, '')
        .trim();
      if (vendorSimple.length >= 4 && contractNotes.toLowerCase().includes(vendorSimple)) {
        vendorMatches = true;
      }
    }
    
    // If vendor doesn't match, skip early
    if (!vendorMatches) {
      return false;
    }
    
    // Log potential match for debugging
    Logger.log(`📋 Potential match: "${contractVendor}" | Submitted By: "${submittedBy}" | Vertical: "${vertical}"`);
    
    // Filter by Submitted By (Andy Worford OR Aden Ritz)
    const submitterMatch = BS_CFG.AIRTABLE_ALLOWED_SUBMITTERS.some(
      allowed => submittedBy.toLowerCase().includes(allowed.toLowerCase())
    );
    if (!submitterMatch) {
      Logger.log(`   ❌ Filtered out: Submitted By "${submittedBy}" not in allowed list`);
      return false;
    }
    
    // Filter by Vertical (Home Services OR Solar)
    const verticalMatch = BS_CFG.AIRTABLE_ALLOWED_VERTICALS.some(
      allowed => vertical.toLowerCase().includes(allowed.toLowerCase())
    );
    if (!verticalMatch) {
      Logger.log(`   ❌ Filtered out: Vertical "${vertical}" not in allowed list`);
      return false;
    }
    
    Logger.log(`   ✅ Match included!`);
    return true;
  });
  
  Logger.log(`📋 Found ${matches.length} contract(s) for vendor: ${vendorName}`);
  return matches;
}

/**
 * Format contracts for display in A(I)DEN
 */
function formatContractsForDisplay_(contracts) {
  if (!contracts || contracts.length === 0) {
    return [];
  }
  
  const formatted = contracts.map(record => {
    const fields = record.fields;
    // Use table/view IDs stored on the record, or fall back to 2025 defaults
    const tableId = record._tableId || BS_CFG.AIRTABLE_CONTRACTS_TABLE_ID_2025;
    const viewId = record._viewId || BS_CFG.AIRTABLE_CONTRACTS_VIEW_ID_2025;
    const tableName = record._tableName || 'Contracts 2025';
    
    // Trim notes to remove trailing newlines/whitespace, and cap length
    const rawNotes = fields[BS_CFG.AIRTABLE_NOTES_FIELD] || '';
    let cleanNotes = rawNotes.trim();
    if (cleanNotes.length > BS_CFG.MAX_NOTES_LENGTH) {
      cleanNotes = cleanNotes.substring(0, BS_CFG.MAX_NOTES_LENGTH) + '...';
    }
    
    // Get created date and format it
    // Try both "Created Date" and "Created" as field names
    const createdDateRaw = fields[BS_CFG.AIRTABLE_CREATED_DATE_FIELD] || fields['Created'] || fields['Created Time'] || '';
    let createdDateFormatted = '';
    if (createdDateRaw) {
      try {
        const date = new Date(createdDateRaw);
        if (!isNaN(date.getTime())) {
          const tz = 'America/Los_Angeles';
          createdDateFormatted = Utilities.formatDate(date, tz, 'yyyy-MM-dd HH:mm');
        }
      } catch (e) {
        createdDateFormatted = createdDateRaw;
      }
    }
    
    // Combine status with created date
    const baseStatus = fields[BS_CFG.AIRTABLE_STATUS_FIELD] || '';
    const statusWithDate = createdDateFormatted ? `${baseStatus} - ${createdDateFormatted}` : baseStatus;
    
    // Log first contract's available fields for debugging
    if (contracts.indexOf(record) === 0) {
      Logger.log(`📋 Contract fields available: ${Object.keys(fields).join(', ')}`);
      Logger.log(`📋 Created date raw value: "${createdDateRaw}" (field: ${BS_CFG.AIRTABLE_CREATED_DATE_FIELD})`);
      Logger.log(`📋 Status with date: "${statusWithDate}"`);
    }
    
    return {
      id: record.id,
      vendorName: fields[BS_CFG.AIRTABLE_VENDOR_FIELD] || '',
      status: statusWithDate,
      contractType: fields[BS_CFG.AIRTABLE_CONTRACT_TYPE_FIELD] || '',
      notes: cleanNotes,
      tableName: tableName,
      createdDate: createdDateRaw,  // Keep raw for sorting
      // Direct link to this record in Airtable using table ID and view ID
      airtableUrl: `https://airtable.com/${BS_CFG.AIRTABLE_BASE_ID}/${tableId}/${viewId}/${record.id}`
    };
  });
  
  // Sort: "Other" type goes last, then by created date DESC
  formatted.sort((a, b) => {
    const typeA = (a.contractType || '').toLowerCase();
    const typeB = (b.contractType || '').toLowerCase();
    
    // "Other" type always goes last
    if (typeA === 'other' && typeB !== 'other') return 1;
    if (typeA !== 'other' && typeB === 'other') return -1;
    
    // Same type priority, sort by created date DESC (newest first)
    const dateA = a.createdDate || '';
    const dateB = b.createdDate || '';
    return dateB.localeCompare(dateA);
  });
  
  return formatted;
}

/**
 * Get contracts for vendor and format for A(I)DEN display
 * This is the main function to call from A(I)DEN
 */
function getVendorContracts_(vendorName) {
  try {
    const contracts = getContractsForVendor_(vendorName);
    const formatted = formatContractsForDisplay_(contracts);
    
    return {
      vendorName: vendorName,
      contractCount: formatted.length,
      contracts: formatted,
      hasContracts: formatted.length > 0
    };
  } catch (e) {
    Logger.log(`Error fetching vendor contracts: ${e.message}`);
    return {
      vendorName: vendorName,
      contractCount: 0,
      contracts: [],
      hasContracts: false
    };
  }
}

/**
 * Test Airtable connection - run this to verify credentials
 */
function testAirtableConnection() {
  const tables = [
    { name: 'Contracts 2026', tableId: BS_CFG.AIRTABLE_CONTRACTS_TABLE_ID_2026 },
    { name: 'Contracts 2025', tableId: BS_CFG.AIRTABLE_CONTRACTS_TABLE_ID_2025 },
    { name: 'Contracts 2024', tableId: BS_CFG.AIRTABLE_CONTRACTS_TABLE_ID_2024 }
  ];

  let allSuccess = true;
  
  for (const table of tables) {
    try {
      const url = `${BS_CFG.AIRTABLE_API_BASE_URL}/${BS_CFG.AIRTABLE_BASE_ID}/${table.tableId}?maxRecords=1`;
      
      const options = {
        method: 'get',
        headers: {
          'Authorization': `Bearer ${BS_CFG.AIRTABLE_API_TOKEN}`,
          'Content-Type': 'application/json'
        },
        muteHttpExceptions: true
      };
      
      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      if (responseCode === 200) {
        const data = JSON.parse(responseText);
        Logger.log(`✅ ${table.name} connection successful!`);
        Logger.log(`   Found ${data.records.length} record(s) in test query`);
      } else {
        Logger.log(`❌ ${table.name} connection failed with status ${responseCode}`);
        Logger.log(`   Response: ${responseText}`);
        allSuccess = false;
      }
    } catch (error) {
      Logger.log(`❌ Error connecting to ${table.name}: ${error.message}`);
      allSuccess = false;
    }
  }
  
  return allSuccess;
}

/************************************************************
 * GOOGLE DRIVE FOLDER FUNCTIONS
 * Searches for vendor folder in Google Drive Vendors folder
 ************************************************************/

/**
 * Get files from vendor's Google Drive folder
 * Searches for a folder matching the vendor name (ignoring YYMMDD prefix)
 * Only returns files at the first level (not in subfolders)
 * 
 * @param {string} vendorName - The vendor name to search for
 * @returns {array} Array of file objects with name, type, modified, url
 */
function getGDriveFilesForVendor_(vendorName) {
  if (!vendorName || vendorName.trim() === '') {
    Logger.log('No vendor name provided for Google Drive search');
    return [];
  }
  
  Logger.log(`=== GOOGLE DRIVE SEARCH ===`);
  Logger.log(`Vendor: ${vendorName}`);
  
  try {
    const parentFolderId = BS_CFG.GDRIVE_VENDORS_FOLDER_ID;
    const cleanVendorName = vendorName.trim().toLowerCase();
    
    Logger.log(`Searching for: "${cleanVendorName}"`);
    
    // Remove common suffixes for alternate search
    const nameWithoutSuffix = cleanVendorName
      .replace(/,?\s*(llc|inc\.?|corp\.?|corporation|company|co\.|l\.?l\.?c\.?)$/i, '')
      .trim();
    
    let vendorFolder = null;
    
    // Strategy 1: Search for exact vendor name phrase
    let searchQuery = `title contains '${vendorName.replace(/'/g, "\\'")}' and '${parentFolderId}' in parents`;
    Logger.log(`Search query: ${searchQuery}`);
    
    let folderIterator = DriveApp.searchFolders(searchQuery);
    
    while (folderIterator.hasNext()) {
      const folder = folderIterator.next();
      const folderName = folder.getName();
      const folderNameLower = folderName.toLowerCase();
      Logger.log(`  Found: "${folderName}"`);
      
      // Check if folder contains the exact vendor name (case-insensitive)
      if (folderNameLower.includes(cleanVendorName)) {
        vendorFolder = folder;
        Logger.log(`Exact match! Using vendor folder: ${folderName}`);
        break;
      }
      
      // Also accept if it contains the name without suffix
      if (nameWithoutSuffix !== cleanVendorName && folderNameLower.includes(nameWithoutSuffix)) {
        vendorFolder = folder;
        Logger.log(`Match without suffix! Using vendor folder: ${folderName}`);
        break;
      }
    }
    
    // Strategy 2: If not found, try searching without suffix
    if (!vendorFolder && nameWithoutSuffix !== cleanVendorName) {
      searchQuery = `title contains '${nameWithoutSuffix.replace(/'/g, "\\'")}' and '${parentFolderId}' in parents`;
      Logger.log(`Trying without suffix: ${searchQuery}`);
      
      folderIterator = DriveApp.searchFolders(searchQuery);
      
      while (folderIterator.hasNext()) {
        const folder = folderIterator.next();
        const folderName = folder.getName();
        const folderNameLower = folderName.toLowerCase();
        Logger.log(`  Found: "${folderName}"`);
        
        if (folderNameLower.includes(nameWithoutSuffix)) {
          vendorFolder = folder;
          Logger.log(`Match! Using vendor folder: ${folderName}`);
          break;
        }
      }
    }
    
    // Strategy 3: Fallback - iterate through parent folder directly (for newly created folders)
    if (!vendorFolder) {
      Logger.log('Search API found nothing, trying direct folder iteration...');
      const parentFolder = DriveApp.getFolderById(parentFolderId);
      folderIterator = parentFolder.getFolders();
      
      while (folderIterator.hasNext()) {
        const folder = folderIterator.next();
        const folderName = folder.getName();
        const folderNameLower = folderName.toLowerCase();
        
        // Check for exact phrase match
        if (folderNameLower.includes(cleanVendorName) || 
            (nameWithoutSuffix !== cleanVendorName && folderNameLower.includes(nameWithoutSuffix))) {
          vendorFolder = folder;
          Logger.log(`Found via direct iteration: ${folderName}`);
          break;
        }
      }
    }
    
    if (!vendorFolder) {
      Logger.log('No matching vendor folder found in Google Drive');
      return { files: [], folderFound: false, folderUrl: null };
    }
    
    // Get files at the first level only (no subfolders)
    const files = [];
    const folderUrl = vendorFolder.getUrl();
    const fileIterator = vendorFolder.getFiles();
    
    while (fileIterator.hasNext()) {
      const file = fileIterator.next();
      
      try {
        const mimeType = file.getMimeType() || '';
        
        // Skip shortcuts (they appear as application/vnd.google-apps.shortcut)
        if (mimeType.includes('shortcut')) {
          Logger.log(`Skipping shortcut: ${file.getName()}`);
          continue;
        }
        
        // Determine file type from mime type
        let fileType = 'File';
        if (mimeType.includes('pdf')) fileType = 'PDF';
        else if (mimeType.includes('spreadsheet') || mimeType.includes('excel')) fileType = 'Sheet';
        else if (mimeType.includes('document') || mimeType.includes('word')) fileType = 'Doc';
        else if (mimeType.includes('presentation') || mimeType.includes('powerpoint')) fileType = 'Slides';
        else if (mimeType.includes('image')) fileType = 'Image';
        else if (mimeType.includes('folder')) continue; // Skip folders
        
        const modDate = file.getLastUpdated();
        const tz = Session.getScriptTimeZone();
        
        files.push({
          name: file.getName(),
          type: fileType,
          modified: modDate ? Utilities.formatDate(modDate, tz, 'yyyy-MM-dd') : '',
          url: file.getUrl(),
          folderUrl: folderUrl
        });
      } catch (e) {
        // Log error but continue with other files
        Logger.log(`Error processing file ${file.getName()}: ${e.message}`);
        files.push({
          name: file.getName(),
          type: '',
          modified: '',
          url: file.getUrl() || '',
          folderUrl: folderUrl
        });
      }
    }
    
    // Sort by modified date (newest first)
    files.sort((a, b) => b.modified.localeCompare(a.modified));
    
    Logger.log(`Found ${files.length} files in vendor folder`);
    return { files: files, folderFound: true, folderUrl: folderUrl };
    
  } catch (e) {
    Logger.log(`Error searching Google Drive: ${e.message}`);
    return { files: [], folderFound: false, folderUrl: null };
  }
}

/************************************************************
 * CRYSTAL BALL - Vendor Email Analysis
 * Analyzes outstanding items and snoozed emails for a vendor
 ************************************************************/

/**
 * Get Crystal Ball analysis for a vendor
 * Analyzes emails in 00.received and snoozed emails to determine what's outstanding
 * Uses caching - only re-runs Claude analysis if email threads have changed
 *
 * @param {string} vendor - The vendor name
 * @param {number} listRow - The row number in the list sheet
 * @returns {object} - { items: [], snoozed: [], summary: string, error: string }
 */
function getCrystalBallData_(vendor, listRow) {
  const ss = SpreadsheetApp.getActive();
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);

  if (!listSh) {
    return { items: [], snoozed: [], summary: '', error: 'List sheet not found' };
  }

  try {
    // Get the Gmail link to extract the vendor label
    const gmailLinkAll = listSh.getRange(listRow, BS_CFG.L_GMAIL_LINK + 1).getValue();
    if (!gmailLinkAll) {
      return { items: [], snoozed: [], summary: '', error: 'No Gmail link found' };
    }

    const gmailLinkStr = gmailLinkAll.toString();

    // Extract vendor label (zzzvendors-*) from URL
    const vendorLabelMatch = gmailLinkStr.match(/label[:%]3A(zzzvendors-[a-z0-9_.\-]+)/i);
    if (!vendorLabelMatch) {
      return { items: [], snoozed: [], summary: '', error: 'Could not extract vendor label' };
    }

    const vendorLabel = vendorLabelMatch[1];
    Logger.log(`[Crystal Ball] Analyzing vendor: ${vendor}, label: ${vendorLabel}`);

    // Search 1: Threads in 00.received with this vendor's label
    const receivedQuery = `label:00.received label:${vendorLabel}`;
    const receivedThreads = GmailApp.search(receivedQuery, 0, 30);
    Logger.log(`[Crystal Ball] Found ${receivedThreads.length} threads in 00.received`);

    // Search 2: Snoozed threads for this vendor
    const snoozedQuery = `is:snoozed label:${vendorLabel}`;
    const snoozedThreads = GmailApp.search(snoozedQuery, 0, 20);
    Logger.log(`[Crystal Ball] Found ${snoozedThreads.length} snoozed threads`);

    // Use existing email checksum (already computed by loadVendorData) + snoozed count
    // This avoids redundant fingerprint computation
    const storedData = getStoredChecksum_(vendor);
    const emailChecksum = storedData?.emailChecksum || 'none';
    const snoozedIds = snoozedThreads.map(t => t.getId()).sort().join(',');
    const currentChecksum = `${emailChecksum}|snoozed:${snoozedThreads.length}:${hashString_(snoozedIds)}`;

    // Check cache - if checksum matches, return cached result
    const cacheKey = `crystal_${vendor.replace(/[^a-zA-Z0-9]/g, '_')}`;
    const cached = getCachedData_('crystalball', cacheKey);
    if (cached && cached.checksum === currentChecksum) {
      Logger.log(`[Crystal Ball] Cache HIT - emails unchanged, using cached analysis`);
      return {
        items: cached.items || [],
        snoozed: cached.snoozed || [],
        summary: cached.summary || '',
        error: null
      };
    }

    Logger.log(`[Crystal Ball] Cache MISS or emails changed - running full analysis`);

    // Process received threads
    const receivedItems = [];
    for (const thread of receivedThreads) {
      const messages = thread.getMessages();
      const lastMessage = messages[messages.length - 1];
      const labels = thread.getLabels().map(l => l.getName());

      // Determine status based on labels
      let status = 'inbox';
      if (labels.some(l => l.includes('02.waiting/customer'))) status = 'waiting-customer';
      else if (labels.some(l => l.includes('02.waiting/me'))) status = 'waiting-me';
      else if (labels.some(l => l.includes('02.waiting/phonexa'))) status = 'waiting-phonexa';
      else if (labels.some(l => l.includes('01.accounting'))) status = 'accounting';

      // Get last sender (skip system emails)
      let lastFrom = lastMessage.getFrom();
      let lastSenderEmail = '';
      for (let i = messages.length - 1; i >= 0; i--) {
        const sender = messages[i].getFrom();
        if (!isSystemOrBounceEmail_(sender)) {
          lastFrom = sender;
          // Extract email from "Name <email>" format
          const emailMatch = sender.match(/<([^>]+)>/) || [null, sender];
          lastSenderEmail = (emailMatch[1] || sender).toLowerCase();
          break;
        }
      }

      // Determine if the last sender is from our team
      const ourDomains = ['profitise.com', 'zeroparallel.com', 'phonexa.com'];
      const lastSenderIsUs = ourDomains.some(d => lastSenderEmail.includes(d));

      receivedItems.push({
        threadId: thread.getId(),
        subject: thread.getFirstMessageSubject(),
        date: lastMessage.getDate().toISOString().split('T')[0],
        lastFrom: lastFrom,
        lastSenderIsUs: lastSenderIsUs,
        status: status,
        labels: labels.filter(l => l.startsWith('02.') || l.startsWith('01.')).join(', '),
        snippet: lastMessage.getPlainBody().substring(0, 200).replace(/\n/g, ' ')
      });
    }

    // Process snoozed threads
    const snoozedItems = [];
    for (const thread of snoozedThreads) {
      const messages = thread.getMessages();
      const lastMessage = messages[messages.length - 1];

      snoozedItems.push({
        threadId: thread.getId(),
        subject: thread.getFirstMessageSubject(),
        date: lastMessage.getDate().toISOString().split('T')[0],
        snippet: lastMessage.getPlainBody().substring(0, 200).replace(/\n/g, ' ')
      });
    }

    // If no emails, return early (and cache this state)
    if (receivedItems.length === 0 && snoozedItems.length === 0) {
      const result = {
        items: [],
        snoozed: [],
        summary: 'No outstanding emails found.',
        error: null
      };
      setCachedData_('crystalball', cacheKey, { ...result, checksum: currentChecksum });
      return result;
    }

    // Build context for Claude analysis - clearly indicate who sent the last message
    const emailContext = receivedItems.map((e, i) => {
      const ballInCourt = e.lastSenderIsUs ? '⚪ BALL IN THEIR COURT (we sent last)' : '🔴 BALL IN OUR COURT (they sent last)';
      return `${i+1}. [${e.status.toUpperCase()}] "${e.subject}" (${e.date})\n   Last message from: ${e.lastFrom}\n   ${ballInCourt}\n   Labels: ${e.labels || 'none'}\n   Preview: ${e.snippet}...`;
    }).join('\n\n');

    const snoozedContext = snoozedItems.map((e, i) =>
      `${i+1}. "${e.subject}" (${e.date})\n   Preview: ${e.snippet}...`
    ).join('\n\n');

    // Call Claude for analysis
    const prompt = `You are analyzing emails for a vendor named "${vendor}" to determine what is OUTSTANDING and needs attention.

CRITICAL CONTEXT:
- "OUR TEAM" = Profitise/ZeroParallel/Phonexa (emails from @profitise.com, @zeroparallel.com, @phonexa.com)
- "VENDOR" = ${vendor} (the external party)
- If WE sent the last message → we are WAITING on the vendor to respond (🟡)
- If THEY sent the last message → we need to take action/respond (🔴)

ACTIVE EMAILS IN INBOX (label:00.received):
${emailContext || 'None'}

SNOOZED EMAILS (deferred for later):
${snoozedContext || 'None'}

Analyze these emails and provide a BRIEF, ACTIONABLE summary. Pay close attention to WHO sent the last message:
- If our team sent the last email, we're WAITING on the vendor - use 🟡
- If the vendor sent the last email, we need to respond - use 🔴

Format your response as a SHORT bulleted list (3-6 bullets max). Be concise - each bullet should be 10 words or less.
Start each bullet with an emoji:
- 🔴 = We need to respond/take action (they sent last)
- 🟡 = Waiting on vendor to respond (we sent last)
- 🔵 = Snoozed/deferred for later
- 🟢 = FYI/informational

Example format:
• 🟡 Awaiting vendor response on Invoice #123
• 🔴 Need to reply to their pricing question
• 🔵 Follow-up snoozed until Jan 15`;

    const apiKey = BS_CFG.CLAUDE_API_KEY;
    if (!apiKey) {
      // No API key - return raw data without analysis
      const result = {
        items: receivedItems,
        snoozed: snoozedItems,
        summary: `${receivedItems.length} active emails, ${snoozedItems.length} snoozed`,
        error: null
      };
      setCachedData_('crystalball', cacheKey, { ...result, checksum: currentChecksum });
      return result;
    }

    const claudeResult = callClaudeAPI_(prompt, apiKey);

    if (claudeResult.error) {
      Logger.log(`[Crystal Ball] Claude API error: ${claudeResult.error}`);
      return {
        items: receivedItems,
        snoozed: snoozedItems,
        summary: `${receivedItems.length} active, ${snoozedItems.length} snoozed (analysis failed)`,
        error: claudeResult.error
      };
    }

    // Cache the successful result with fingerprint
    const result = {
      items: receivedItems,
      snoozed: snoozedItems,
      summary: claudeResult.content,
      error: null
    };
    setCachedData_('crystalball', cacheKey, { ...result, checksum: currentChecksum });
    Logger.log(`[Crystal Ball] Analysis complete and cached`);

    return result;

  } catch (e) {
    Logger.log(`[Crystal Ball] Error: ${e.message}`);
    return { items: [], snoozed: [], summary: '', error: e.message };
  }
}

/**
 * Ask a question about a vendor and search their emails for the answer
 * Uses Claude to analyze all emails with the vendor's zzzVendors label
 */
function askAboutVendor() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);

  if (!listSh) {
    ui.alert('Vendor list not found.');
    return;
  }

  const currentIndex = getCurrentVendorIndex_();
  const listRow = currentIndex + 1;
  const vendor = listSh.getRange(listRow, 1).getValue();

  if (!vendor) {
    ui.alert('No vendor currently loaded. Please navigate to a vendor first.');
    return;
  }

  // Prompt for the question
  const response = ui.prompt(
    `❓ Ask About ${vendor}`,
    'Enter your question (e.g., "What verticals do they want?", "What was their last pricing request?"):',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const question = response.getResponseText().trim();
  if (!question) {
    ui.alert('Please enter a question.');
    return;
  }

  ss.toast('Searching emails and analyzing...', '🔍 Searching', 10);

  // Get all emails from vendor label
  const emails = getAllEmailsFromVendorLabel_(listRow, 100); // Get more emails for context

  if (emails.length === 0) {
    ui.alert('No emails found for this vendor.');
    return;
  }

  Logger.log(`[Ask About Vendor] Found ${emails.length} emails for ${vendor}`);

  // Build email context for Claude - include more detail than Crystal Ball
  const emailContext = emails.slice(0, 50).map((e, i) => {
    // Format date nicely
    const dateStr = e.date || 'Unknown date';

    // Get the content snippet
    const snippet = (e.snippet || '').replace(/\n+/g, ' ').trim();

    // Format labels nicely
    const labelsStr = Array.isArray(e.labels) ? e.labels.join(', ') : (e.labels || 'none');

    return `--- Email ${i+1} ---
Subject: ${e.subject}
Date: ${dateStr}
From: ${e.from || 'Unknown'}
To: ${e.to || 'Unknown'}
Labels: ${labelsStr}
Content: ${snippet}`;
  }).join('\n\n');

  // Call Claude with the question
  const prompt = `You are helping analyze emails for a vendor named "${vendor}".

USER'S QUESTION: ${question}

Here are the emails from this vendor (most recent first):

${emailContext}

Based on these emails, please answer the user's question. Be specific and cite relevant emails when possible (mention the subject line or date). If you can't find a clear answer in the emails, say so and explain what related information you did find.

Keep your answer concise but complete.`;

  const apiKey = BS_CFG.CLAUDE_API_KEY;
  if (!apiKey) {
    ui.alert('Claude API key not configured.');
    return;
  }

  const result = callClaudeAPI_(prompt, apiKey);

  if (result.error) {
    ui.alert(`Error: ${result.error}`);
    return;
  }

  // Store the question for "Search More" functionality
  const props = PropertiesService.getScriptProperties();
  props.setProperty('BS_ASK_VENDOR_QUESTION', question);
  props.setProperty('BS_ASK_VENDOR_OFFSET', '100'); // Next search starts at 100

  // Show answer in a dialog with "Search Older Emails" button
  const html = `
    <style>
      body { font-family: Arial, sans-serif; font-size: 14px; padding: 15px; line-height: 1.5; }
      h2 { color: #1a73e8; margin-bottom: 10px; }
      .question { background: #e8f0fe; padding: 10px; border-radius: 5px; margin-bottom: 15px; font-weight: bold; }
      .answer { background: #f8f9fa; padding: 15px; border-radius: 5px; white-space: pre-wrap; }
      .meta { color: #666; font-size: 12px; margin-top: 15px; }
      .btn { padding: 10px 20px; margin-top: 15px; cursor: pointer; border: none; border-radius: 4px; font-size: 14px; }
      .btn-secondary { background: #6c757d; color: white; }
      .btn-secondary:hover { background: #5a6268; }
      .btn-secondary:disabled { background: #ccc; cursor: not-allowed; }
      #loading { display: none; color: #666; margin-top: 10px; }
    </style>
    <h2>❓ Answer for ${escapeHtml_(vendor)}</h2>
    <div class="question">Q: ${escapeHtml_(question)}</div>
    <div class="answer" id="answerBox">${escapeHtml_(result.content)}</div>
    <div class="meta" id="metaBox">Searched emails 1-${Math.min(emails.length, 50)} of ${emails.length} retrieved</div>
    <button class="btn btn-secondary" id="searchMoreBtn" onclick="searchMore()">🔍 Search Older Emails (101-200)</button>
    <div id="loading">⏳ Searching older emails...</div>

    <script>
      function searchMore() {
        document.getElementById('searchMoreBtn').disabled = true;
        document.getElementById('loading').style.display = 'block';

        google.script.run
          .withSuccessHandler(function(result) {
            document.getElementById('loading').style.display = 'none';
            if (result.error) {
              alert('Error: ' + result.error);
              document.getElementById('searchMoreBtn').disabled = false;
            } else {
              document.getElementById('answerBox').textContent = result.answer;
              document.getElementById('metaBox').textContent = result.meta;
              if (result.hasMore) {
                document.getElementById('searchMoreBtn').textContent = '🔍 Search Even Older Emails (' + result.nextRange + ')';
                document.getElementById('searchMoreBtn').disabled = false;
              } else {
                document.getElementById('searchMoreBtn').textContent = '✓ No more emails to search';
                document.getElementById('searchMoreBtn').disabled = true;
              }
            }
          })
          .withFailureHandler(function(error) {
            document.getElementById('loading').style.display = 'none';
            alert('Error: ' + error.message);
            document.getElementById('searchMoreBtn').disabled = false;
          })
          .askAboutVendorContinue();
      }
    </script>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(700)
    .setHeight(550);

  ui.showModalDialog(htmlOutput, `❓ Ask About ${vendor}`);

  Logger.log(`[Ask About Vendor] Answered question for ${vendor}: "${question}"`);
}

/**
 * Continue searching older emails for Ask About Vendor
 * Called from the dialog when user clicks "Search Older Emails"
 */
function askAboutVendorContinue() {
  const ss = SpreadsheetApp.getActive();
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);

  const props = PropertiesService.getScriptProperties();
  const question = props.getProperty('BS_ASK_VENDOR_QUESTION');
  const offset = parseInt(props.getProperty('BS_ASK_VENDOR_OFFSET') || '100', 10);

  if (!question) {
    return { error: 'No active question. Please start a new search.' };
  }

  const currentIndex = getCurrentVendorIndex_();
  const listRow = currentIndex + 1;
  const vendor = listSh.getRange(listRow, 1).getValue();

  // Get emails with offset
  const emails = getAllEmailsFromVendorLabelWithOffset_(listRow, 100, offset);

  if (emails.length === 0) {
    return {
      error: null,
      answer: 'No more emails found in this range.',
      meta: `Searched emails ${offset + 1}-${offset + 100} (none found)`,
      hasMore: false,
      nextRange: ''
    };
  }

  Logger.log(`[Ask About Vendor Continue] Found ${emails.length} emails at offset ${offset} for ${vendor}`);

  // Build email context
  const emailContext = emails.slice(0, 50).map((e, i) => {
    const dateStr = e.date || 'Unknown date';
    const snippet = (e.snippet || '').replace(/\n+/g, ' ').trim();
    const labelsStr = Array.isArray(e.labels) ? e.labels.join(', ') : (e.labels || 'none');

    return `--- Email ${offset + i + 1} ---
Subject: ${e.subject}
Date: ${dateStr}
From: ${e.from || 'Unknown'}
To: ${e.to || 'Unknown'}
Labels: ${labelsStr}
Content: ${snippet}`;
  }).join('\n\n');

  const prompt = `You are helping analyze emails for a vendor named "${vendor}".

USER'S QUESTION: ${question}

Here are OLDER emails from this vendor (emails ${offset + 1}-${offset + emails.length}):

${emailContext}

Based on these emails, please answer the user's question. Be specific and cite relevant emails when possible (mention the subject line or date). If you can't find a clear answer in the emails, say so and explain what related information you did find.

Keep your answer concise but complete.`;

  const apiKey = BS_CFG.CLAUDE_API_KEY;
  const result = callClaudeAPI_(prompt, apiKey);

  if (result.error) {
    return { error: result.error };
  }

  // Update offset for next search
  const newOffset = offset + 100;
  props.setProperty('BS_ASK_VENDOR_OFFSET', newOffset.toString());

  const hasMore = emails.length >= 50; // If we got a full batch, there might be more
  const nextRange = `${newOffset + 1}-${newOffset + 100}`;

  return {
    error: null,
    answer: result.content,
    meta: `Searched emails ${offset + 1}-${offset + Math.min(emails.length, 50)}`,
    hasMore: hasMore,
    nextRange: nextRange
  };
}

/**
 * Get emails from vendor label with offset for pagination
 */
function getAllEmailsFromVendorLabelWithOffset_(listRow, maxThreads, offset) {
  const ss = SpreadsheetApp.getActive();
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);

  if (!listSh) return [];

  try {
    const gmailLinkAll = listSh.getRange(listRow, BS_CFG.L_GMAIL_LINK + 1).getValue();
    if (!gmailLinkAll) return [];

    const gmailLinkStr = gmailLinkAll.toString();
    const vendorLabelMatch = gmailLinkStr.match(/label[:%]3A(zzzvendors-[a-z0-9_.\-]+)/i);
    if (!vendorLabelMatch) return [];

    const vendorLabel = vendorLabelMatch[1];
    const searchQuery = `label:${vendorLabel}`;

    // Search with offset
    const threads = GmailApp.search(searchQuery, offset, maxThreads);

    Logger.log(`Found ${threads.length} threads at offset ${offset}`);

    const emails = [];
    for (const thread of threads) {
      const messages = thread.getMessages();
      const lastMessage = messages[messages.length - 1];
      const labels = thread.getLabels().map(l => l.getName());

      emails.push({
        threadId: thread.getId(),
        subject: thread.getFirstMessageSubject(),
        date: lastMessage.getDate().toISOString().split('T')[0],
        from: lastMessage.getFrom(),
        to: lastMessage.getTo(),
        snippet: lastMessage.getPlainBody().substring(0, 1500),
        messageCount: messages.length,
        labels: labels,
        lastMessageDate: lastMessage.getDate()
      });
    }

    emails.sort((a, b) => b.lastMessageDate - a.lastMessageDate);
    return emails;

  } catch (e) {
    Logger.log(`ERROR in getAllEmailsFromVendorLabelWithOffset_: ${e.message}`);
    return [];
  }
}

/**
 * Find a monday.com item ID by searching for a vendor name
 */
function findMondayItemIdByVendor_(vendor, boardId, apiToken) {
  Logger.log(`=== SEARCHING FOR VENDOR ===`);
  Logger.log(`Search term: "${vendor}"`);
  Logger.log(`Board ID: ${boardId}`);
  
  let query = `
    query {
      boards (ids: [${boardId}]) {
        items_page (limit: 100, query_params: {rules: [{column_id: "name", compare_value: ["${vendor.replace(/"/g, '\\"')}"]}]}) {
          items { id name }
        }
      }
    }
  `;
  
  let itemId = tryFindItem_(query, apiToken, 'Exact match');
  if (itemId) return itemId;
  
  const withoutParens = vendor.replace(/\s*\([^)]*\)/g, '').trim();
  if (withoutParens !== vendor) {
    Logger.log(`Trying without parentheses: "${withoutParens}"`);
    query = `
      query {
        boards (ids: [${boardId}]) {
          items_page (limit: 100, query_params: {rules: [{column_id: "name", compare_value: ["${withoutParens.replace(/"/g, '\\"')}"]}]}) {
            items { id name }
          }
        }
      }
    `;
    
    itemId = tryFindItem_(query, apiToken, 'Without parentheses');
    if (itemId) return itemId;
  }
  
  Logger.log(`Trying contains search...`);
  query = `
    query {
      boards (ids: [${boardId}]) {
        items_page (limit: 500) {
          items { id name }
        }
      }
    }
  `;
  
  try {
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': apiToken },
      payload: JSON.stringify({ query: query }),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch('https://api.monday.com/v2', options);
    const result = JSON.parse(response.getContentText());
    
    if (result.data?.boards?.[0]?.items_page?.items) {
      const items = result.data.boards[0].items_page.items;
      
      for (const item of items) {
        const itemName = String(item.name || '').toLowerCase();
        const searchTerm = vendor.toLowerCase();
        const searchTermNoParens = withoutParens.toLowerCase();
        
        if (itemName.includes(searchTerm) || searchTerm.includes(itemName) ||
            itemName.includes(searchTermNoParens) || searchTermNoParens.includes(itemName)) {
          Logger.log(`✓ FOUND MATCH: "${item.name}" (ID: ${item.id})`);
          return item.id;
        }
      }
    }
    
    return null;
  } catch (e) {
    Logger.log(`Error in contains search: ${e}`);
    return null;
  }
}

/**
 * Helper to try a specific query
 */
function tryFindItem_(query, apiToken, attemptName) {
  try {
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': apiToken },
      payload: JSON.stringify({ query: query }),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch('https://api.monday.com/v2', options);
    const result = JSON.parse(response.getContentText());
    
    if (result.data?.boards?.[0]?.items_page?.items?.length > 0) {
      const item = result.data.boards[0].items_page.items[0];
      Logger.log(`✓ ${attemptName} SUCCESS: Found "${item.name}" (ID: ${item.id})`);
      return item.id;
    } else {
      Logger.log(`✗ ${attemptName} failed: No items found`);
    }
    
    return null;
  } catch (e) {
    Logger.log(`✗ ${attemptName} error: ${e}`);
    return null;
  }
}

/**
 * Navigation: Go to next vendor
 */
function battleStationNext() {
  const ss = SpreadsheetApp.getActive();
  const bsSh = ss.getSheetByName(BS_CFG.BATTLE_SHEET);
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
  
  if (!bsSh || !listSh) {
    SpreadsheetApp.getUi().alert('A(I)DEN not found. Run setupBattleStation() first.');
    return;
  }
  
  const currentIndex = getCurrentVendorIndex_();
  
  if (!currentIndex) {
    loadVendorData(1);
    return;
  }
  
  const totalVendors = listSh.getLastRow() - 1;
  
  if (currentIndex >= totalVendors) {
    ss.toast('Already at the last vendor!', '⚠️ End of List', 3);
    return;
  }
  
  const listRow = currentIndex + 1;
  const vendor = listSh.getRange(listRow, BS_CFG.L_VENDOR + 1).getValue();
  listSh.getRange(listRow, BS_CFG.L_PROCESSED + 1).setValue(true);
  
  ss.toast(`Marked "${vendor}" as reviewed`, '▶️ Next', 2);
  loadVendorData(currentIndex + 1);
}

/**
 * Navigation: Go to previous vendor
 */
function battleStationPrevious() {
  const ss = SpreadsheetApp.getActive();
  const bsSh = ss.getSheetByName(BS_CFG.BATTLE_SHEET);
  
  if (!bsSh) {
    SpreadsheetApp.getUi().alert('A(I)DEN not found. Run setupBattleStation() first.');
    return;
  }
  
  const currentIndex = getCurrentVendorIndex_();
  
  if (!currentIndex) {
    loadVendorData(1);
    return;
  }
  
  if (currentIndex <= 1) {
    SpreadsheetApp.getActive().toast('Already at the first vendor!', '⚠️ Start of List', 3);
    return;
  }
  
  loadVendorData(currentIndex - 1);
}

/**
 * Refresh current vendor data
 */
function battleStationRefresh() {
  const ss = SpreadsheetApp.getActive();
  const bsSh = ss.getSheetByName(BS_CFG.BATTLE_SHEET);
  
  if (!bsSh) {
    SpreadsheetApp.getUi().alert('A(I)DEN not found. Run setupBattleStation() first.');
    return;
  }
  
  const currentIndex = getCurrentVendorIndex_();
  loadVendorData(currentIndex || 1);
}

/**
 * Quick Refresh - Email only, uses cached data for Airtable/Box
 * Much faster than full refresh when you've only changed emails
 */
/**
 * Quick Refresh - Email only, just refreshes the EMAILS section without redrawing everything
 * Much faster than full refresh
 */
function battleStationQuickRefresh() {
  const ss = SpreadsheetApp.getActive();
  const bsSh = ss.getSheetByName(BS_CFG.BATTLE_SHEET);
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);

  if (!bsSh || !listSh) {
    SpreadsheetApp.getUi().alert('A(I)DEN not found. Run setupBattleStation() first.');
    return;
  }

  // Check for pending archive threads (from Email Responses) and archive if I sent the reply
  const didArchive = checkAndArchivePendingThreads_();
  if (didArchive) {
    // Threads were archived - need to wait for Gmail to update, then refresh
    ss.toast('Archived sent emails, waiting for Gmail to update...', '📬 Auto-Archive', 2);
    Utilities.sleep(2000);
    // Do a Quick Refresh Until Changed to catch the archive
    battleStationQuickRefreshUntilChanged();
    return;
  }

  ss.toast('Refreshing emails only...', '⚡ Quick Refresh', 2);
  
  // Get current vendor info
  const currentIndex = getCurrentVendorIndex_();
  const listRow = currentIndex + 1;
  const vendor = listSh.getRange(listRow, 1).getValue();
  
  // Find the EMAILS section in the sheet
  const dataRange = bsSh.getDataRange();
  const values = dataRange.getValues();
  
  let emailsHeaderRow = -1;
  let emailsEndRow = -1;
  
  // Find emails section start
  for (let row = 0; row < values.length; row++) {
    const cellValue = String(values[row][0] || '');
    if (cellValue.includes('📧 EMAILS')) {
      emailsHeaderRow = row + 1; // 1-indexed
    } else if (emailsHeaderRow > 0 && (cellValue.includes('📋 MONDAY.COM TASKS') || cellValue.includes('📋 MONDAY'))) {
      // Found the next section (TASKS comes after EMAILS)
      emailsEndRow = row; // 0-indexed, so this is the row before TASKS header (1-indexed)
      break;
    }
  }
  
  if (emailsHeaderRow < 0) {
    ss.toast('Could not find EMAILS section', '❌ Error', 3);
    return;
  }
  
  // If we didn't find the end, assume it goes to current row count
  if (emailsEndRow < 0) {
    emailsEndRow = values.length;
  }
  
  // Fetch fresh emails
  const emails = getEmailsForVendor_(vendor, listRow);
  
  // Calculate how many rows we need vs how many we have
  const emailRowsNeeded = emails.length > 0 ? Math.min(emails.length, 20) + 2 : 2; // +2 for header row and column headers (or "no emails" row)
  if (emails.length > 20) emailRowsNeeded + 1; // +1 for "and X more" row
  
  const availableRows = emailsEndRow - emailsHeaderRow;
  
  // Clear existing email content (keep header row formatting)
  const clearStartRow = emailsHeaderRow + 1; // Row after "📧 EMAILS" header
  const clearEndRow = emailsEndRow;
  
  if (clearEndRow > clearStartRow) {
    bsSh.getRange(clearStartRow, 1, clearEndRow - clearStartRow, 4)
      .clearContent()
      .clearFormat()
      .setBackground('#ffffff')
      .setFontWeight('normal')
      .setFontStyle('normal');
    
    // Unmerge any merged cells in the range
    for (let r = clearStartRow; r < clearEndRow; r++) {
      try {
        bsSh.getRange(r, 1, 1, 4).breakApart();
      } catch (e) {
        // Ignore if not merged
      }
    }
  }
  
  // Update the header with new count
  bsSh.getRange(emailsHeaderRow, 1, 1, 4).breakApart();
  bsSh.getRange(emailsHeaderRow, 1, 1, 4).merge()
    .setValue(`📧 EMAILS (${emails.length})  |  🔵 Snoozed  🔴 Overdue  🟠 Phonexa  🟢 Accounting  🟡 Waiting`)
    .setBackground('#f8f9fa')
    .setFontWeight('bold')
    .setFontSize(10)
    .setFontColor('#1a73e8')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  
  let currentRow = emailsHeaderRow + 1;
  
  if (emails.length === 0) {
    bsSh.getRange(currentRow, 1, 1, 4).merge()
      .setValue('No emails found')
      .setFontStyle('italic')
      .setBackground('#fafafa')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    currentRow++;
  } else {
    // Email headers
    bsSh.getRange(currentRow, 1).setValue('Subject').setFontWeight('bold').setBackground('#f3f3f3');
    bsSh.getRange(currentRow, 2).setValue('Date').setFontWeight('bold').setBackground('#f3f3f3');
    bsSh.getRange(currentRow, 3).setValue('Last').setFontWeight('bold').setBackground('#f3f3f3');
    bsSh.getRange(currentRow, 4).setValue('Labels').setFontWeight('bold').setBackground('#f3f3f3');
    currentRow++;

    for (const email of emails.slice(0, 20)) {
      bsSh.getRange(currentRow, 1).setValue(email.subject);
      const emailDateCell2 = bsSh.getRange(currentRow, 2);
      emailDateCell2.setNumberFormat('@'); // Set format BEFORE value to prevent auto-parsing
      emailDateCell2.setValue(email.date);
      bsSh.getRange(currentRow, 3).setValue(email.lastFrom);
      bsSh.getRange(currentRow, 4).setValue(email.labels);

      if (email.link) {
        bsSh.getRange(currentRow, 1)
          .setFormula(`=HYPERLINK("${email.link}", "${email.subject.replace(/"/g, '""')}")`);
      }
      
      // Check if this email is overdue
      const isOverdue = isEmailOverdue_(email);
      
      // Check if email has priority label
      const hasPriority = email.labels.includes('01.priority/1');
      
      // Color priority: Snoozed > OVERDUE > Phonexa > Accounting > Waiting/Customer > Active
      if (email.isSnoozed) {
        bsSh.getRange(currentRow, 1, 1, 4).setBackground(BS_CFG.COLOR_SNOOZED);
      } else if (isOverdue) {
        bsSh.getRange(currentRow, 1, 1, 4).setBackground(BS_CFG.COLOR_OVERDUE);
        bsSh.getRange(currentRow, 1, 1, 4).setFontWeight('bold');
      } else if (email.labels.includes('02.waiting/phonexa')) {
        bsSh.getRange(currentRow, 1, 1, 4).setBackground(BS_CFG.COLOR_PHONEXA);
      } else if (email.labels.includes('04.accounting-invoices')) {
        bsSh.getRange(currentRow, 1, 1, 4).setBackground('#d9ead3');
      } else if (email.labels.includes('02.waiting/customer')) {
        bsSh.getRange(currentRow, 1, 1, 4).setBackground(BS_CFG.COLOR_WAITING);
      } else {
        bsSh.getRange(currentRow, 1, 1, 4).setBackground('#ffffff');
      }
      
      // If missing 01.priority/1, make text grey to indicate lower importance
      if (!hasPriority) {
        bsSh.getRange(currentRow, 1, 1, 4).setFontColor('#999999');
      }
      
      currentRow++;
    }
    
    if (emails.length > 20) {
      bsSh.getRange(currentRow, 1, 1, 4).merge()
        .setValue(`... and ${emails.length - 20} more emails (showing first 20)`)
        .setFontStyle('italic')
        .setHorizontalAlignment('center');
      currentRow++;
    }
  }
  
  // Update checksum for emails
  const newEmailChecksum = generateEmailChecksum_(emails);
  updateEmailChecksum_(vendor, newEmailChecksum);

  ss.toast('Emails refreshed!', '⚡ Done', 2);
}

/**
 * Quick Refresh Until Changed - repeatedly refreshes emails until a change is detected
 * Useful when you've made email changes and are waiting for them to take effect
 */
function battleStationQuickRefreshUntilChanged() {
  const ss = SpreadsheetApp.getActive();
  const bsSh = ss.getSheetByName(BS_CFG.BATTLE_SHEET);
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);

  if (!bsSh || !listSh) {
    SpreadsheetApp.getUi().alert('A(I)DEN not found. Run setupBattleStation() first.');
    return;
  }

  const currentIndex = getCurrentVendorIndex_();
  const listRow = currentIndex + 1;
  const vendor = listSh.getRange(listRow, 1).getValue();

  // Get current email checksum before refreshing
  const storedData = getStoredChecksum_(vendor);
  // Convert to string for consistent comparison (spreadsheet reads numbers as strings)
  const oldChecksum = storedData && storedData.emailChecksum != null ? String(storedData.emailChecksum) : null;

  const maxAttempts = 20;  // Max attempts to prevent infinite loop
  const delaySeconds = 3;  // Seconds between attempts

  Logger.log(`Quick Refresh Until Changed - Starting for ${vendor}`);
  Logger.log(`Old checksum: ${oldChecksum} (type: ${typeof oldChecksum})`);

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    ss.toast(`Checking for email changes... (attempt ${attempt}/${maxAttempts})`, '🔁 Waiting', delaySeconds + 1);

    // Fetch fresh emails
    const emails = getEmailsForVendor_(vendor, listRow);
    const newChecksum = String(generateEmailChecksum_(emails));

    Logger.log(`Attempt ${attempt}: newChecksum = ${newChecksum} (type: ${typeof newChecksum})`);

    if (newChecksum !== oldChecksum) {
      // Change detected! Do a full quick refresh to update the display
      Logger.log(`Change detected! Old: ${oldChecksum}, New: ${newChecksum}`);
      ss.toast('Change detected! Refreshing display...', '✅ Found', 2);
      battleStationQuickRefresh();
      return;
    }

    // No change yet, wait before next attempt
    if (attempt < maxAttempts) {
      Utilities.sleep(delaySeconds * 1000);
    }
  }

  ss.toast(`No changes detected after ${maxAttempts} attempts`, '⏱️ Timeout', 3);
}

/**
 * Update just the email checksum for a vendor
 */
function updateEmailChecksum_(vendor, newEmailChecksum) {
  const sh = getChecksumsSheet_();
  const data = sh.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === vendor) {
      sh.getRange(i + 1, 3).setValue(newEmailChecksum); // Column C = EmailChecksum
      sh.getRange(i + 1, 5).setValue(new Date()); // Column E = Last Viewed
      Logger.log(`Updated email checksum for ${vendor}: ${newEmailChecksum}`);
      return;
    }
  }
  
  // Vendor not found - this shouldn't happen but handle it
  Logger.log(`Warning: Could not find ${vendor} in checksums sheet to update email checksum`);
}

/**
 * Hard Refresh - Clear cache and reload everything fresh
 */
function battleStationHardRefresh() {
  const ss = SpreadsheetApp.getActive();
  const bsSh = ss.getSheetByName(BS_CFG.BATTLE_SHEET);
  
  if (!bsSh) {
    SpreadsheetApp.getUi().alert('A(I)DEN not found. Run setupBattleStation() first.');
    return;
  }
  
  // Clear the cache
  clearBSCache_();
  ss.toast('Cache cleared, refreshing...', '🔄 Hard Refresh', 2);
  
  const currentIndex = getCurrentVendorIndex_();
  loadVendorData(currentIndex || 1, { useCache: false });
}

/**
 * Get or create the cache sheet
 */
function getBSCacheSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(BS_CFG.CACHE_SHEET);
  
  if (!sh) {
    sh = ss.insertSheet(BS_CFG.CACHE_SHEET);
    // Set up headers: Type, Key, Data (JSON), LastUpdated
    sh.getRange(1, 1, 1, 4).setValues([['Type', 'Key', 'Data', 'LastUpdated']]);
    sh.hideSheet();
  }
  
  return sh;
}

/**
 * Clear the BS cache
 */
function clearBSCache_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(BS_CFG.CACHE_SHEET);
  
  if (sh) {
    sh.clear();
    sh.getRange(1, 1, 1, 4).setValues([['Type', 'Key', 'Data', 'LastUpdated']]);
  }
  
  Logger.log('BS Cache cleared');
}

/**
 * Get cached data if fresh enough
 * Returns null if cache is stale or doesn't exist
 */
function getCachedData_(type, key) {
  const sh = getBSCacheSheet_();
  const data = sh.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === type && data[i][1] === key) {
      const lastUpdated = new Date(data[i][3]);
      const ageHours = (new Date() - lastUpdated) / (1000 * 60 * 60);
      
      if (ageHours < BS_CFG.CACHE_MAX_AGE_HOURS) {
        try {
          Logger.log(`Cache HIT: ${type}/${key} (${Math.round(ageHours * 10) / 10}h old)`);
          return JSON.parse(data[i][2]);
        } catch (e) {
          Logger.log(`Cache parse error: ${e.message}`);
          return null;
        }
      } else {
        Logger.log(`Cache STALE: ${type}/${key} (${Math.round(ageHours * 10) / 10}h old)`);
        return null;
      }
    }
  }
  
  Logger.log(`Cache MISS: ${type}/${key}`);
  return null;
}

/**
 * Store data in cache
 */
function setCachedData_(type, key, data) {
  const sh = getBSCacheSheet_();
  const existingData = sh.getDataRange().getValues();
  
  // Look for existing row to update
  for (let i = 1; i < existingData.length; i++) {
    if (existingData[i][0] === type && existingData[i][1] === key) {
      sh.getRange(i + 1, 3, 1, 2).setValues([[JSON.stringify(data), new Date()]]);
      Logger.log(`Cache UPDATE: ${type}/${key}`);
      return;
    }
  }
  
  // Add new row
  sh.appendRow([type, key, JSON.stringify(data), new Date()]);
  Logger.log(`Cache SET: ${type}/${key}`);
}

/**
 * Update monday.com notes for current vendor - NO DIALOGS
 */
function battleStationUpdateMondayNotes() {
  const ss = SpreadsheetApp.getActive();
  const bsSh = ss.getSheetByName(BS_CFG.BATTLE_SHEET);
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
  
  if (!bsSh || !listSh) {
    ss.toast('Required sheets not found', '❌ Error', 3);
    return;
  }
  
  const currentIndex = getCurrentVendorIndex_();
  
  if (!currentIndex || isNaN(currentIndex)) {
    ss.toast('Could not determine current vendor', '❌ Error', 3);
    return;
  }
  
  const listRow = currentIndex + 1;
  const vendor = String(listSh.getRange(listRow, BS_CFG.L_VENDOR + 1).getValue() || '').trim();
  
  if (!vendor) {
    ss.toast('Could not determine vendor name', '❌ Error', 3);
    return;
  }
  
  // Find the notes row - look for 📝 NOTES header
  let notesRow = -1;
  for (let i = 5; i < 50; i++) {
    const label = String(bsSh.getRange(i, 1).getValue() || '');
    if (label.indexOf('📝 NOTES') !== -1) {
      notesRow = i + 1;  // Notes content is in the row after the header
      break;
    }
  }
  
  if (notesRow === -1) {
    ss.toast('Could not find notes field', '❌ Error', 3);
    return;
  }
  
  const notes = String(bsSh.getRange(notesRow, 1).getValue() || '').trim();
  
  if (!notes || notes === '(no notes)') {
    ss.toast('No notes to sync - edit the notes field first', '⚠️ Empty', 3);
    return;
  }
  
  ss.toast(`Updating ${vendor}...`, '⚡ Syncing', 3);
  
  try {
    listSh.getRange(listRow, BS_CFG.L_NOTES + 1).setValue(notes);
    
    const result = updateMondayComNotesForVendor_(vendor, notes, listRow);
    
    if (result.success) {
      ss.toast(`Notes updated for ${vendor}`, '✅ Success', 3);
      battleStationRefresh();
    } else {
      ss.toast(`Failed: ${result.error}`, '❌ Error', 5);
    }
  } catch (e) {
    ss.toast(`Error: ${e.message}`, '❌ Error', 5);
  }
}

/**
 * Helper function to update monday.com notes via API
 */
function updateMondayComNotesForVendor_(vendor, notes, listRow) {
  const apiToken = BS_CFG.MONDAY_API_TOKEN;
  const ss = SpreadsheetApp.getActive();
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
  
  if (!listRow) {
    const currentIndex = getCurrentVendorIndex_();
    if (!currentIndex) return { success: false, error: 'Could not determine vendor index' };
    listRow = currentIndex + 1;
  }
  
  const source = String(listSh.getRange(listRow, BS_CFG.L_SOURCE + 1).getValue() || '');
  
  let boardId, notesColumnId;
  if (source.toLowerCase().includes('buyer')) {
    boardId = BS_CFG.BUYERS_BOARD_ID;
    notesColumnId = BS_CFG.BUYERS_NOTES_COLUMN;
  } else if (source.toLowerCase().includes('affiliate')) {
    boardId = BS_CFG.AFFILIATES_BOARD_ID;
    notesColumnId = BS_CFG.AFFILIATES_NOTES_COLUMN;
  } else {
    boardId = BS_CFG.BUYERS_BOARD_ID;
    notesColumnId = BS_CFG.BUYERS_NOTES_COLUMN;
  }
  
  const itemId = findMondayItemIdByVendor_(vendor, boardId, apiToken);
  
  if (!itemId) {
    return { success: false, error: `Could not find monday.com item for vendor: ${vendor}` };
  }
  
  const valueJson = JSON.stringify(notes);
  const escapedValue = valueJson.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
  
  const mutation = `
    mutation {
      change_column_value (
        board_id: ${boardId},
        item_id: ${itemId},
        column_id: "${notesColumnId}",
        value: "${escapedValue}"
      ) { id }
    }
  `;
  
  try {
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': apiToken },
      payload: JSON.stringify({ query: mutation }),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch('https://api.monday.com/v2', options);
    const result = JSON.parse(response.getContentText());
    
    if (result.errors && result.errors.length > 0) {
      return { success: false, error: result.errors[0].message };
    }
    
    if (result.data?.change_column_value?.id) {
      return { success: true, itemId: result.data.change_column_value.id };
    }
    
    return { success: false, error: 'Unexpected API response' };
    
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Mark current vendor as reviewed
 */
function battleStationMarkReviewed() {
  const ss = SpreadsheetApp.getActive();
  const bsSh = ss.getSheetByName(BS_CFG.BATTLE_SHEET);
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
  
  if (!bsSh || !listSh) {
    SpreadsheetApp.getUi().alert('Required sheets not found.');
    return;
  }
  
  const currentIndex = getCurrentVendorIndex_();
  
  if (!currentIndex) {
    SpreadsheetApp.getUi().alert('Error: Could not determine current vendor index.');
    return;
  }
  
  const listRow = currentIndex + 1;
  const vendor = listSh.getRange(listRow, BS_CFG.L_VENDOR + 1).getValue();
  
  listSh.getRange(listRow, BS_CFG.L_PROCESSED + 1).setValue(true);
  
  battleStationRefresh();
  ss.toast('Marked as reviewed!', '✅ ' + vendor, 3);
}

/**
 * Open Gmail search for current vendor
 */
function battleStationOpenGmail() {
  const ss = SpreadsheetApp.getActive();
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
  
  if (!listSh) {
    SpreadsheetApp.getUi().alert('Required sheets not found.');
    return;
  }
  
  const currentIndex = getCurrentVendorIndex_();
  
  if (!currentIndex) {
    SpreadsheetApp.getUi().alert('Error: Could not determine current vendor index.');
    return;
  }
  
  const listRow = currentIndex + 1;
  const gmailLink = listSh.getRange(listRow, BS_CFG.L_GMAIL_LINK + 1).getValue();
  
  if (!gmailLink || gmailLink.toString().indexOf('#search') === -1) {
    SpreadsheetApp.getUi().alert('No valid Gmail search link found.');
    return;
  }
  
  const html = `<html><body><script>window.open('${gmailLink}', '_blank');google.script.host.close();</script></body></html>`;
  const ui = HtmlService.createHtmlOutput(html).setWidth(200).setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(ui, 'Opening Gmail...');
}

/**
 * Create Gmail draft to vendor contacts
 */
function battleStationEmailContacts() {
  const ss = SpreadsheetApp.getActive();
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
  
  if (!listSh) {
    SpreadsheetApp.getUi().alert('Required sheets not found.');
    return;
  }
  
  const currentIndex = getCurrentVendorIndex_();
  
  if (!currentIndex || isNaN(currentIndex)) {
    SpreadsheetApp.getUi().alert('Error: Could not determine current vendor index.');
    return;
  }
  
  const listRow = currentIndex + 1;
  const vendor = String(listSh.getRange(listRow, BS_CFG.L_VENDOR + 1).getValue() || '').trim();
  
  ss.toast('Finding contacts...', '📧 Creating Draft', 2);
  
  const contactData = getVendorContacts_(vendor, listRow);
  
  if (contactData.contacts.length === 0) {
    SpreadsheetApp.getUi().alert('No contacts found for this vendor.');
    return;
  }
  
  const recipients = contactData.contacts.map(c => c.email).filter(e => e).join(', ');
  
  if (!recipients) {
    SpreadsheetApp.getUi().alert('No email addresses found for contacts.');
    return;
  }
  
  try {
    const subject = `Re: ${vendor}`;
    const body = `Hi,\n\n\n\nBest regards,\nAndy Worford\nProfitise`;
    
    GmailApp.createDraft(recipients, subject, body);
    
    ss.toast('Draft created!', '✅ Success', 3);
    SpreadsheetApp.getUi().alert(`✓ Draft created!\n\nTo: ${recipients}\nSubject: ${subject}\n\nCheck your Gmail drafts.`);
    
  } catch (e) {
    SpreadsheetApp.getUi().alert(`Error creating draft: ${e.message}`);
  }
}

/**
 * Analyze unsnoozed emails with Claude AI - Individual breakdown with inline links
 */
function battleStationAnalyzeEmails() {
  const ss = SpreadsheetApp.getActive();
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
  
  const claudeApiKey = BS_CFG.CLAUDE_API_KEY;
  
  if (!claudeApiKey || claudeApiKey === 'YOUR_ANTHROPIC_API_KEY_HERE') {
    SpreadsheetApp.getUi().alert('Please set your Anthropic API key in BS_CFG.CLAUDE_API_KEY');
    return;
  }
  
  if (!listSh) {
    SpreadsheetApp.getUi().alert('Required sheets not found.');
    return;
  }
  
  const currentIndex = getCurrentVendorIndex_();
  
  if (!currentIndex || isNaN(currentIndex)) {
    SpreadsheetApp.getUi().alert('Error: Could not determine current vendor index.');
    return;
  }
  
  const listRow = currentIndex + 1;
  const vendor = String(listSh.getRange(listRow, BS_CFG.L_VENDOR + 1).getValue() || '').trim();
  
  ss.toast('Fetching emails for analysis...', '🤖 Claude Analysis', 3);
  
  const emails = getEmailsForVendor_(vendor, listRow);
  const unsnoozedEmails = emails.filter(e => !e.isSnoozed);
  
  if (unsnoozedEmails.length === 0) {
    SpreadsheetApp.getUi().alert('No unsnoozed emails found for this vendor.');
    return;
  }
  
  ss.toast(`Analyzing ${unsnoozedEmails.length} emails with Claude...`, '🤖 Processing', 10);
  
  // Gather email content with metadata
  const emailData = [];
  
  for (const email of unsnoozedEmails.slice(0, 10)) {
    try {
      const thread = GmailApp.getThreadById(email.threadId);
      if (thread) {
        const messages = thread.getMessages();
        let fullContent = '';
        
        for (const msg of messages.slice(-5)) {
          fullContent += `\n--- From: ${msg.getFrom()} | Date: ${msg.getDate()} ---\n`;
          fullContent += msg.getPlainBody().substring(0, 1500);
        }
        
        emailData.push({
          subject: email.subject,
          date: email.date,
          labels: email.labels,
          link: email.link,
          threadId: email.threadId,
          messageCount: email.count,
          content: fullContent
        });
      }
    } catch (e) {
      emailData.push({
        subject: email.subject,
        date: email.date,
        labels: email.labels,
        link: email.link,
        threadId: email.threadId,
        messageCount: email.count,
        content: email.snippet || '(could not fetch content)'
      });
    }
  }
  
  if (emailData.length === 0) {
    SpreadsheetApp.getUi().alert('Could not fetch email content for analysis.');
    return;
  }
  
  // Build prompt - use numbered emails that we can match to links
  const emailsText = emailData.map((e, i) => 
    `\n\n=== EMAIL_${i + 1} ===
Subject: ${e.subject}
Date: ${e.date}
Labels: ${e.labels}
Messages in thread: ${e.messageCount}

Content:
${e.content}`
  ).join('\n');
  
  const prompt = `You are analyzing email communications for a vendor relationship manager at a lead generation company called Profitise.

The vendor being analyzed is: ${vendor}

Here are the recent unsnoozed emails (${emailData.length} threads):
${emailsText}

Please provide your analysis in this EXACT format (keep EMAIL_1, EMAIL_2 etc as markers - they will be replaced with links):

## OVERALL SUMMARY
[2-3 sentences about the overall relationship status]

## EMAIL BREAKDOWN

EMAIL_1
**Status**: [Active / Waiting on them / Waiting on us / FYI / Urgent]
**Summary**: [1-2 sentence summary]
**Action**: [What to do, or "None"]

EMAIL_2
**Status**: [Status]
**Summary**: [Summary]
**Action**: [Action]

[Continue for each email...]

## PRIORITY ACTIONS
[Numbered list of most important actions]

Be concise. Use the exact EMAIL_1, EMAIL_2 markers so they can be linked.`;

  try {
    const response = callClaudeAPI_(prompt, claudeApiKey);
    
    if (response.error) {
      SpreadsheetApp.getUi().alert(`Claude API Error: ${response.error}`);
      return;
    }
    
    // Format content and replace EMAIL_X markers with clickable links
    let formattedContent = response.content
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
    
    // Replace EMAIL_X markers with linked subject lines
    emailData.forEach((e, i) => {
      const marker = `EMAIL_${i + 1}`;
      const shortSubject = e.subject.length > 60 ? e.subject.substring(0, 60) + '...' : e.subject;
      const linkedSubject = `<a href="${e.link}" target="_blank" class="email-link">📧 ${shortSubject}</a> <span class="date">(${e.date})</span>`;
      formattedContent = formattedContent.replace(new RegExp(marker, 'g'), linkedSubject);
    });
    
    // Apply formatting
    formattedContent = formattedContent
      .replace(/\n/g, '<br>')
      .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
      .replace(/## (.*?)<br>/g, '<h3>$1</h3>')
      .replace(/### (.*?)<br>/g, '<h4>$1</h4>');
    
    const htmlContent = `
      <style>
        body { font-family: Arial, sans-serif; padding: 15px; line-height: 1.6; font-size: 13px; }
        h2 { color: #4a86e8; margin-top: 0; margin-bottom: 10px; }
        h3 { color: #4a86e8; margin-top: 20px; margin-bottom: 10px; border-bottom: 2px solid #4a86e8; padding-bottom: 5px; }
        h4 { color: #666; margin-top: 12px; margin-bottom: 5px; }
        .email-link { 
          color: #1a73e8; 
          text-decoration: none; 
          font-weight: bold;
          font-size: 14px;
          display: inline-block;
          margin-top: 10px;
          padding: 5px 10px;
          background: #e8f0fe;
          border-radius: 4px;
        }
        .email-link:hover { background: #d0e1fd; text-decoration: none; }
        .date { color: #888; font-size: 11px; }
        .content { background: #fafafa; padding: 15px; border-radius: 5px; }
        strong { color: #333; }
      </style>
      <h2>🤖 Claude Analysis: ${vendor}</h2>
      <p><em>Analyzed ${emailData.length} unsnoozed email threads</em></p>
      <div class="content">${formattedContent}</div>
    `;
    
    const html = HtmlService.createHtmlOutput(htmlContent).setWidth(750).setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(html, `Claude Analysis: ${vendor}`);
    
    ss.toast('Analysis complete!', '✅ Done', 3);
    
  } catch (e) {
    SpreadsheetApp.getUi().alert(`Error: ${e.message}`);
  }
}

/**
 * Call Claude API
 */
function callClaudeAPI_(prompt, apiKey) {
  const url = 'https://api.anthropic.com/v1/messages';
  
  const payload = {
    model: 'claude-sonnet-4-20250514',
    max_tokens: 2000,
    messages: [{ role: 'user', content: prompt }]
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    if (responseCode !== 200) {
      return { error: `API returned ${responseCode}: ${responseText}` };
    }
    
    const result = JSON.parse(responseText);
    
    if (result.content && result.content.length > 0) {
      return { content: result.content[0].text };
    }
    
    return { error: 'No content in response' };
    
  } catch (e) {
    return { error: e.message };
  }
}

/**
 * Go to a specific vendor by index or name
 */
function battleStationGoTo() {
  const ss = SpreadsheetApp.getActive();
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
  
  if (!listSh) return;
  
  const totalVendors = listSh.getLastRow() - 1;
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Go to Vendor',
    `Enter vendor index (1-${totalVendors}) or vendor name:`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const input = response.getResponseText().trim();
    const index = parseInt(input);
    
    // If it's a number, go directly to that index
    if (!isNaN(index) && index >= 1 && index <= totalVendors) {
      loadVendorData(index);
      return;
    }
    
    // Otherwise, search by name
    const searchTerm = input.toLowerCase();
    const data = listSh.getDataRange().getValues();
    const matches = [];
    
    for (let i = 1; i < data.length; i++) {
      const vendorName = String(data[i][0] || '').toLowerCase();
      if (vendorName.includes(searchTerm)) {
        matches.push({ index: i, name: data[i][0] });
      }
    }
    
    if (matches.length === 0) {
      ui.alert(`No vendors found matching "${input}"`);
    } else if (matches.length === 1) {
      loadVendorData(matches[0].index);
    } else {
      // Multiple matches - show list
      let matchList = matches.slice(0, 15).map(m => `${m.index}: ${m.name}`).join('\n');
      if (matches.length > 15) {
        matchList += `\n... and ${matches.length - 15} more`;
      }
      
      const pickResponse = ui.prompt(
        `Found ${matches.length} matches`,
        `${matchList}\n\nEnter the index number to go to:`,
        ui.ButtonSet.OK_CANCEL
      );
      
      if (pickResponse.getSelectedButton() === ui.Button.OK) {
        const pickIndex = parseInt(pickResponse.getResponseText());
        if (!isNaN(pickIndex) && pickIndex >= 1 && pickIndex <= totalVendors) {
          loadVendorData(pickIndex);
        }
      }
    }
  }
}

/**
 * Generate a checksum for vendor data
 * Based on: emails, tasks, notes, status, states, contracts, helpful links, meetings, box documents, gdrive files
 */
function generateVendorChecksum_(vendor, emails, tasks, notes, status, states, contracts, helpfulLinks, meetings, boxDocs, gDriveFiles) {
  const data = {
    vendor: vendor,
    status: status || '',
    notes: notes || '',
    states: states || '',
    emails: (emails || []).map(e => ({
      subject: e.subject,
      date: e.date,
      labels: e.labels
    })),
    // Include overdue count so full checksum changes when emails become overdue
    overdueEmailCount: (emails || []).filter(e => isEmailOverdue_(e)).length,
    tasks: (tasks || []).map(t => ({
      name: t.subject,
      status: t.status,
      created: t.created
    })),
    contracts: (contracts || []).map(c => ({
      vendor: c.vendorName,
      status: c.status,
      type: c.contractType,
      notes: c.notes
    })),
    helpfulLinks: (helpfulLinks || []).map(l => ({
      name: l.name,
      url: l.url,
      notes: l.notes
    })),
    meetings: (meetings || []).map(m => ({
      title: m.title,
      date: m.date,
      time: m.time,
      isPast: m.isPast  // Track if meeting moved from future to past
    })),
    boxDocs: (boxDocs || []).map(d => ({
      name: d.name,
      modified: d.modifiedAt
    })),
    gDriveFiles: (gDriveFiles || []).map(f => ({
      name: f.name,
      modified: f.modified
    }))
  };
  
  const jsonStr = JSON.stringify(data);
  return hashString_(jsonStr);
}

/**
 * Simple hash function used by all checksum generators
 */
function hashString_(str) {
  let hash = 0;
  for (let i = 0; i < str.length; i++) {
    const char = str.charCodeAt(i);
    hash = ((hash << 5) - hash) + char;
    hash = hash & hash;
  }
  return hash.toString(16);
}

// Increment this when checksum format changes to avoid false positives
const MODULE_CHECKSUMS_VERSION = 3;

/**
 * Generate sub-checksums for each module
 * Returns an object with checksums for each data section
 */
function generateModuleChecksums_(vendor, emails, tasks, notes, status, states, contracts, helpfulLinks, meetings, boxDocs, gDriveFiles, contacts) {
  // Track which specific emails are overdue (for showing "no longer overdue" details)
  const overdueEmails = (emails || [])
    .filter(e => isEmailOverdue_(e))
    .map(e => ({ threadId: e.threadId, subject: e.subject }));

  return {
    _version: MODULE_CHECKSUMS_VERSION,  // Version for compatibility checking
    emails: generateEmailChecksum_(emails),  // Use the overdue-aware email checksum
    tasks: generateTasksChecksum_(tasks),
    notes: hashString_(JSON.stringify(notes || '')),
    status: hashString_(JSON.stringify(status || '')),
    states: hashString_(JSON.stringify(states || '')),
    contracts: hashString_(JSON.stringify((contracts || []).map(c => ({ vendor: c.vendorName, status: c.status, type: c.contractType })))),
    helpfulLinks: hashString_(JSON.stringify((helpfulLinks || []).map(l => ({ name: l.name, url: l.url })))),
    meetings: generateMeetingsChecksum_(meetings),
    boxDocs: hashString_(JSON.stringify((boxDocs || []).map(d => ({ name: d.name, modified: d.modifiedAt })))),
    gDriveFiles: hashString_(JSON.stringify((gDriveFiles || []).map(f => ({ name: f.name, modified: f.modified })))),
    contacts: generateContactsChecksum_(contacts),
    overdueEmails: overdueEmails  // Store specific overdue emails for comparison
  };
}

/**
 * Generate a sub-checksum for just emails (most volatile data)
 * Prefixed with thread count for quick change detection
 */
function generateEmailChecksum_(emails) {
  const count = (emails || []).length;

  // Base checksum - same as before so existing checksums stay valid
  const data = (emails || []).map(e => ({
    subject: e.subject,
    date: e.date,
    labels: e.labels
  }));
  const baseChecksum = hashString_(JSON.stringify(data));

  // Count overdue emails - append to checksum so it changes when overdue status changes
  const overdueCount = (emails || []).filter(e => isEmailOverdue_(e)).length;

  // Format: count:hash or count:hash_ODn
  // Prefix with count allows fast count-based change detection
  const hashPart = overdueCount > 0 ? `${baseChecksum}_OD${overdueCount}` : baseChecksum;
  return `${count}:${hashPart}`;
}

/**
 * Generate a secondary checksum for vendor-label-only Gmail search
 * This catches emails that might not have label:00.received yet
 * @param {string} gmailLink - The full Gmail search link (with 00.received)
 * @returns {string|null} - Checksum or null if can't extract vendor label
 */
function generateVendorLabelChecksum_(gmailLink) {
  if (!gmailLink) return null;

  try {
    const gmailLinkStr = gmailLink.toString();

    // Extract the vendor label (zzzvendors-*) from the URL
    // Include periods in vendor name (e.g., "inc." in "american-remodeling-enterprises-inc.")
    const vendorLabelMatch = gmailLinkStr.match(/label[:%]3A(zzzvendors-[a-z0-9_.\-]+)/i);
    if (!vendorLabelMatch) {
      Logger.log('Could not extract vendor label from Gmail link');
      return null;
    }

    const vendorLabel = vendorLabelMatch[1];

    // Build simplified search query: just vendor label, no snoozed, no noInbox
    const searchQuery = `label:${vendorLabel} -is:snoozed -label:03.noInbox`;
    Logger.log(`Vendor label search query: ${searchQuery}`);

    // Search Gmail
    const threads = GmailApp.search(searchQuery, 0, 100);
    Logger.log(`Vendor label search found ${threads.length} threads`);

    // Generate simple checksum based on thread count and thread IDs
    const threadData = threads.map(t => ({
      id: t.getId(),
      lastDate: t.getLastMessageDate().getTime()
    }));

    return hashString_(JSON.stringify(threadData));
  } catch (e) {
    Logger.log(`Error in generateVendorLabelChecksum_: ${e.message}`);
    return null;
  }
}

/**
 * Check if an email is overdue:
 * - REQUIRES 01.priority/1 label (non-priority emails are never overdue)
 * - Priority + any waiting label (customer, me, phonexa) + >16 business hours
 */
function isEmailOverdue_(email) {
  if (!email || !email.labels) return false;

  // Snoozed emails are never overdue
  if (email.isSnoozed) return false;

  // MUST have priority label - non-priority emails are never overdue
  if (!email.labels.includes('01.priority/1')) return false;

  // Parse the email date
  const emailDate = parseEmailDate_(email.date);
  if (!emailDate) return false;

  // Check for any waiting label - 16 business hours
  const isWaiting = email.labels.includes('02.waiting/customer') ||
                    email.labels.includes('02.waiting/me') ||
                    email.labels.includes('02.waiting/phonexa');
  if (isWaiting) {
    const businessHours = getBusinessHoursElapsed_(emailDate);
    if (businessHours > BS_CFG.OVERDUE_BUSINESS_HOURS) return true;
  }

  return false;
}

/**
 * Parse email date string to Date object
 * Handles formats like "Dec 5, 2024 3:45 PM" or "Dec 5, 3:45 PM"
 */
function parseEmailDate_(dateStr) {
  if (!dateStr) return null;
  
  try {
    // Try direct parse first
    let date = new Date(dateStr);
    if (!isNaN(date.getTime())) return date;
    
    // If no year, assume current year
    const currentYear = new Date().getFullYear();
    date = new Date(dateStr + ' ' + currentYear);
    if (!isNaN(date.getTime())) return date;
    
    return null;
  } catch (e) {
    return null;
  }
}

/**
 * Calculate business hours elapsed since a given date
 * Business hours: Monday-Friday, 9 AM - 5 PM Pacific
 */
function getBusinessHoursElapsed_(startDate) {
  if (!startDate) return 0;
  
  const now = new Date();
  const start = new Date(startDate);
  
  if (start >= now) return 0;
  
  let businessHours = 0;
  let current = new Date(start);
  
  // Iterate hour by hour (simplified approach)
  while (current < now) {
    const day = current.getDay(); // 0=Sun, 1=Mon, ..., 6=Sat
    const hour = current.getHours();
    
    // Business hours: Monday-Friday (1-5), 9 AM - 5 PM (9-16)
    if (day >= 1 && day <= 5 && hour >= 9 && hour < 17) {
      businessHours++;
    }
    
    current.setHours(current.getHours() + 1);
    
    // Safety limit - don't count more than 1000 hours
    if (businessHours > 1000) break;
  }
  
  return businessHours;
}

/**
 * Get or create the Checksums sheet
 */
function getChecksumsSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(BS_CFG.CHECKSUMS_SHEET);

  if (!sh) {
    sh = ss.insertSheet(BS_CFG.CHECKSUMS_SHEET);
    sh.getRange(1, 1, 1, 9).setValues([['Vendor', 'Checksum', 'EmailChecksum', 'ModuleChecksums', 'Last Viewed', 'Flagged', 'EmailData', 'SnoozedUntil', 'VendorLabelChecksum']]);
    sh.getRange(1, 1, 1, 9).setFontWeight('bold');
    sh.hideSheet();
  } else {
    // Check if we need to add columns (migration)
    const headers = sh.getRange(1, 1, 1, 8).getValues()[0];
    if (headers[3] !== 'ModuleChecksums') {
      const numCols = sh.getLastColumn();
      if (numCols < 5) {
        sh.insertColumnAfter(3);
        sh.getRange(1, 4).setValue('ModuleChecksums').setFontWeight('bold');
        sh.getRange(1, 5).setValue('Last Viewed').setFontWeight('bold');
      }
    }
    // Add Flagged column if missing
    if (headers[5] !== 'Flagged') {
      const numCols = sh.getLastColumn();
      if (numCols < 6) {
        sh.getRange(1, 6).setValue('Flagged').setFontWeight('bold');
      }
    }
    // Add EmailData column if missing
    if (headers[6] !== 'EmailData') {
      const numCols = sh.getLastColumn();
      if (numCols < 7) {
        sh.getRange(1, 7).setValue('EmailData').setFontWeight('bold');
      }
    }
    // Add SnoozedUntil column if missing
    if (headers[7] !== 'SnoozedUntil') {
      const numCols = sh.getLastColumn();
      if (numCols < 8) {
        sh.getRange(1, 8).setValue('SnoozedUntil').setFontWeight('bold');
      }
    }
    // Add VendorLabelChecksum column if missing (column 9)
    const numCols = sh.getLastColumn();
    if (numCols < 9) {
      sh.getRange(1, 9).setValue('VendorLabelChecksum').setFontWeight('bold');
    }
  }

  return sh;
}

/**
 * Reset all module checksums to force a fresh baseline
 * This clears only the ModuleChecksums column (4), preserving flags, snooze dates, etc.
 * After running this, the first traversal will establish new baselines without false positives
 */
function resetAllModuleChecksums() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    'Reset Module Checksums',
    'This will clear all stored module checksums. The next traversal will establish fresh baselines.\n\nFlags, snooze dates, and email data will be preserved.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );

  if (result !== ui.Button.YES) {
    ui.alert('Cancelled');
    return;
  }

  const sh = getChecksumsSheet_();
  const lastRow = sh.getLastRow();

  if (lastRow > 1) {
    // Clear column 4 (ModuleChecksums) for all data rows
    sh.getRange(2, 4, lastRow - 1, 1).clearContent();
  }

  ui.alert('Done', `Cleared module checksums for ${lastRow - 1} vendors. Next traversal will establish fresh baselines.`, ui.ButtonSet.OK);
}

/**
 * Check if a vendor is flagged for review
 */
function isVendorFlagged_(vendor) {
  const sh = getChecksumsSheet_();
  const data = sh.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === vendor.toLowerCase()) {
      return data[i][5] === true || data[i][5] === 'TRUE' || data[i][5] === 'true';
    }
  }
  return false;
}

/**
 * Set or clear the flag for a vendor
 */
function setVendorFlag_(vendor, flagged) {
  const sh = getChecksumsSheet_();
  const data = sh.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === vendor.toLowerCase()) {
      sh.getRange(i + 1, 6).setValue(flagged);
      return true;
    }
  }

  // Vendor not in checksums yet - add a row
  if (flagged) {
    const lastRow = sh.getLastRow();
    sh.getRange(lastRow + 1, 1).setValue(vendor);
    sh.getRange(lastRow + 1, 6).setValue(true);
  }
  return true;
}

/**
 * Toggle flag for the currently displayed vendor
 */
function battleStationToggleFlag() {
  const ss = SpreadsheetApp.getActive();
  const bsSh = ss.getSheetByName(BS_CFG.BATTLE_SHEET);

  if (!bsSh) {
    SpreadsheetApp.getUi().alert('A(I)DEN sheet not found.');
    return;
  }

  // Get vendor name from row 2 (vendor name banner) - remove any icons (flag/snooze)
  const rawValue = String(bsSh.getRange(2, 1).getValue() || '').trim();
  const vendor = rawValue.replace(/\s*⚑\s*/, '').replace(/\s*💤\d+\/\d+\s*$/, '').trim();
  if (!vendor) {
    SpreadsheetApp.getUi().alert('No vendor currently displayed.');
    return;
  }

  const currentlyFlagged = isVendorFlagged_(vendor);
  setVendorFlag_(vendor, !currentlyFlagged);

  // Update the display with or without flag icon
  if (!currentlyFlagged) {
    bsSh.getRange(2, 1).setValue(`${vendor} ⚑`);
    ss.toast(`⚑ Flagged "${vendor}" - will stop here on next skip`, '⚑ Flagged', 3);
  } else {
    bsSh.getRange(2, 1).setValue(vendor);
    ss.toast(`Unflagged "${vendor}"`, '⚑ Unflagged', 3);
  }
}

/**
 * Check if a vendor is snoozed (and snooze date hasn't passed)
 * Returns true if snoozed and date is in the future
 */
function isVendorSnoozed_(vendor) {
  const snoozeDate = getVendorSnoozeDate_(vendor);
  if (!snoozeDate) return false;

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  return snoozeDate > today;
}

/**
 * Get the snooze date for a vendor
 * Returns Date object or null if not snoozed
 */
function getVendorSnoozeDate_(vendor) {
  const sh = getChecksumsSheet_();
  const data = sh.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === vendor.toLowerCase()) {
      const snoozeVal = data[i][7];
      if (snoozeVal instanceof Date) {
        return snoozeVal;
      }
      if (snoozeVal) {
        const parsed = new Date(snoozeVal);
        if (!isNaN(parsed.getTime())) {
          return parsed;
        }
      }
      return null;
    }
  }
  return null;
}

/**
 * Set or clear the snooze date for a vendor
 */
function setVendorSnooze_(vendor, snoozeDate) {
  const sh = getChecksumsSheet_();
  const data = sh.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === vendor.toLowerCase()) {
      sh.getRange(i + 1, 8).setValue(snoozeDate || '');
      return true;
    }
  }

  // Vendor not in checksums yet - add a row
  if (snoozeDate) {
    const lastRow = sh.getLastRow();
    sh.getRange(lastRow + 1, 1).setValue(vendor);
    sh.getRange(lastRow + 1, 8).setValue(snoozeDate);
  }
  return true;
}

/**
 * Snooze the currently displayed vendor until a specified date
 */
function battleStationSnoozeVendor() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const bsSh = ss.getSheetByName(BS_CFG.BATTLE_SHEET);

  if (!bsSh) {
    ui.alert('A(I)DEN sheet not found.');
    return;
  }

  // Get vendor name from row 2 (vendor name banner) - remove any icons
  const rawValue = String(bsSh.getRange(2, 1).getValue() || '').trim();
  const vendor = rawValue.replace(/\s*⚑\s*$/, '').replace(/\s*💤.*$/, '').trim();
  if (!vendor) {
    ui.alert('No vendor currently displayed.');
    return;
  }

  // Check if already snoozed
  const currentSnooze = getVendorSnoozeDate_(vendor);
  if (currentSnooze && currentSnooze > new Date()) {
    const response = ui.alert(
      'Already Snoozed',
      `"${vendor}" is snoozed until ${currentSnooze.toLocaleDateString()}.\n\nDo you want to clear the snooze?`,
      ui.ButtonSet.YES_NO
    );
    if (response === ui.Button.YES) {
      setVendorSnooze_(vendor, null);
      // Update display
      const newDisplay = rawValue.replace(/\s*💤.*$/, '').trim();
      bsSh.getRange(2, 1).setValue(newDisplay);
      ss.toast(`Snooze cleared for "${vendor}"`, '⏰ Unsnooze', 3);
    }
    return;
  }

  // Prompt for snooze date
  const response = ui.prompt(
    'Snooze Vendor',
    `Enter snooze date for "${vendor}" (YYYY-MM-DD):`,
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const dateStr = response.getResponseText().trim();

  // Parse date parts to avoid timezone issues (YYYY-MM-DD interpreted as UTC)
  const dateParts = dateStr.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (!dateParts) {
    ui.alert('Invalid date format. Please use YYYY-MM-DD.');
    return;
  }

  const snoozeDate = new Date(
    parseInt(dateParts[1], 10),      // year
    parseInt(dateParts[2], 10) - 1,  // month (0-indexed)
    parseInt(dateParts[3], 10)       // day
  );

  if (isNaN(snoozeDate.getTime())) {
    ui.alert('Invalid date format. Please use YYYY-MM-DD.');
    return;
  }

  // Set snooze to end of day
  snoozeDate.setHours(23, 59, 59, 999);

  if (snoozeDate <= new Date()) {
    ui.alert('Snooze date must be in the future.');
    return;
  }

  setVendorSnooze_(vendor, snoozeDate);

  // Update display with snooze indicator
  const displayDate = snoozeDate.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
  const baseDisplay = rawValue.replace(/\s*💤.*$/, '').trim();
  bsSh.getRange(2, 1).setValue(`${baseDisplay} 💤${displayDate}`);

  ss.toast(`Snoozed "${vendor}" until ${displayDate}`, '💤 Snoozed', 3);
}

/**
 * Get Box file blacklist from Settings sheet
 * Returns object like { "ANGI": ["fileId1", "fileId2"], "Vendor2": ["fileId3"] }
 * 
 * Settings sheet format (Box Blacklist section - can start in any column):
 * Row: "Box Blacklist" | (empty) | (empty) | (empty)
 * Row: "Vendor Name" | "File ID" | "File Name (for reference)" | "Reason"
 * Row: "ANGI" | "123456789" | "inbound1_IO.docx" | "False positive - ANGI in content"
 */
function getBoxBlacklist_() {
  const ss = SpreadsheetApp.getActive();
  const settingsSh = ss.getSheetByName('Settings');
  
  if (!settingsSh) {
    return {};
  }
  
  const data = settingsSh.getDataRange().getValues();
  const blacklist = {};
  let inBlacklistSection = false;
  let startCol = -1;  // Track which column the blacklist starts in
  
  for (let i = 0; i < data.length; i++) {
    // Search for "Box Blacklist" header in any column
    if (!inBlacklistSection) {
      for (let col = 0; col < data[i].length; col++) {
        if (String(data[i][col] || '').trim().toLowerCase() === 'box blacklist') {
          inBlacklistSection = true;
          startCol = col;
          break;
        }
      }
      continue;
    }
    
    // Now we're in the blacklist section, use startCol for data
    const vendorCell = String(data[i][startCol] || '').trim();
    const fileIdCell = String(data[i][startCol + 1] || '').trim();
    
    // Skip header row
    if (vendorCell.toLowerCase() === 'vendor name') {
      continue;
    }
    
    // Exit if we hit an empty row (both vendor and file ID empty)
    if (vendorCell === '' && fileIdCell === '') {
      break;
    }
    
    // Parse blacklist entries
    if (vendorCell !== '' && fileIdCell !== '') {
      if (!blacklist[vendorCell]) {
        blacklist[vendorCell] = [];
      }
      blacklist[vendorCell].push(fileIdCell);
    }
  }
  
  Logger.log(`Box blacklist loaded: ${JSON.stringify(blacklist)}`);
  return blacklist;
}

/**
 * Get stored checksum for a vendor
 * Returns object with { checksum, emailChecksum, moduleChecksums, emailData }
 */
function getStoredChecksum_(vendor) {
  const sh = getChecksumsSheet_();
  const data = sh.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === vendor.toLowerCase()) {
      let moduleChecksums = null;
      let emailData = null;
      try {
        if (data[i][3]) {
          moduleChecksums = JSON.parse(data[i][3]);
        }
      } catch (e) {
        // Invalid JSON, ignore
      }
      try {
        if (data[i][6]) {
          emailData = JSON.parse(data[i][6]);
        }
      } catch (e) {
        // Invalid JSON, ignore
      }
      return {
        checksum: data[i][1],
        emailChecksum: data[i][2] || null,
        moduleChecksums: moduleChecksums,
        emailData: emailData,
        vendorLabelChecksum: data[i][8] || null  // Column 9 (0-indexed = 8)
      };
    }
  }

  return null;
}

/**
 * Store checksum for a vendor (including module sub-checksums and email data)
 * @param {string} vendor
 * @param {string} checksum
 * @param {string} emailChecksum
 * @param {object} moduleChecksums
 * @param {object} emailData
 * @param {string} vendorLabelChecksum - Optional secondary checksum for vendor-label-only search
 */
function storeChecksum_(vendor, checksum, emailChecksum, moduleChecksums, emailData, vendorLabelChecksum) {
  const sh = getChecksumsSheet_();
  const data = sh.getDataRange().getValues();
  const now = new Date();
  const moduleJson = moduleChecksums ? JSON.stringify(moduleChecksums) : '';
  const emailDataJson = emailData ? JSON.stringify(emailData) : '';

  // Look for existing row
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === vendor.toLowerCase()) {
      // Update columns 2-5 (Checksum, EmailChecksum, ModuleChecksums, Last Viewed) and column 7 (EmailData), column 9 (VendorLabelChecksum)
      sh.getRange(i + 1, 2, 1, 4).setValues([[checksum, emailChecksum || '', moduleJson, now]]);
      sh.getRange(i + 1, 7).setValue(emailDataJson);
      if (vendorLabelChecksum) {
        sh.getRange(i + 1, 9).setValue(vendorLabelChecksum);
      }
      return;
    }
  }

  // Add new row (9 columns now)
  sh.appendRow([vendor, checksum, emailChecksum || '', moduleJson, now, '', emailDataJson, '', vendorLabelChecksum || '']);
}

/**
 * Log what changed in emails by comparing old and new email data
 * @returns {object} { added: [], removed: [] }
 */
function logEmailChanges_(oldEmails, newEmails, vendor) {
  // Create maps for comparison using subject+date as key
  const oldMap = new Map((oldEmails || []).map(e => [`${e.subject}|${e.date}`, e]));
  const newMap = new Map((newEmails || []).map(e => [`${e.subject}|${e.date}`, e]));

  const added = [];
  const removed = [];

  // Find added emails (in new but not in old)
  for (const [key, email] of newMap) {
    if (!oldMap.has(key)) {
      added.push(email);
    }
  }

  // Find removed emails (in old but not in new)
  for (const [key, email] of oldMap) {
    if (!newMap.has(key)) {
      removed.push(email);
    }
  }

  // Log the changes
  if (added.length > 0 || removed.length > 0) {
    Logger.log(`📧 Email changes for ${vendor}:`);
    if (added.length > 0) {
      Logger.log(`  ➕ ADDED (${added.length}):`);
      for (const e of added.slice(0, 5)) {
        Logger.log(`     - "${e.subject}" (${e.date})`);
      }
      if (added.length > 5) Logger.log(`     ... and ${added.length - 5} more`);
    }
    if (removed.length > 0) {
      Logger.log(`  ➖ REMOVED (${removed.length}):`);
      for (const e of removed.slice(0, 5)) {
        Logger.log(`     - "${e.subject}" (${e.date})`);
      }
      if (removed.length > 5) Logger.log(`     ... and ${removed.length - 5} more`);
    }
  } else {
    // If no added/removed but checksum changed, might be label changes
    Logger.log(`📧 Email changes for ${vendor}: Labels or other metadata changed (same threads)`);
  }

  return { added, removed };
}

/**
 * Parse the overdue count from an email checksum (e.g., "-17ff163c_OD1" -> 1)
 */
function parseOverdueCount_(checksum) {
  if (!checksum) return 0;
  const match = String(checksum).match(/_OD(\d+)$/);
  return match ? parseInt(match[1], 10) : 0;
}

/**
 * Get human-readable description of email checksum change
 */
function getEmailChangeDescription_(oldChecksum, newChecksum, addedCount, removedCount) {
  const oldOverdue = parseOverdueCount_(oldChecksum);
  const newOverdue = parseOverdueCount_(newChecksum);

  const changes = [];

  if (addedCount > 0) {
    changes.push(`${addedCount} new email${addedCount > 1 ? 's' : ''}`);
  }
  if (removedCount > 0) {
    changes.push(`${removedCount} removed`);
  }
  if (oldOverdue !== newOverdue) {
    if (newOverdue > oldOverdue) {
      changes.push(`${newOverdue - oldOverdue} became overdue`);
    } else {
      changes.push(`${oldOverdue - newOverdue} no longer overdue`);
    }
  }

  if (changes.length === 0) {
    changes.push('labels or metadata changed');
  }

  return changes.join(', ');
}

/************************************************************
 * HELPER FUNCTIONS - Shared utilities for change detection
 ************************************************************/

/**
 * Make a monday.com API request
 * Centralizes all monday API calls for consistency
 * @param {string} query - GraphQL query
 * @param {string} apiToken - API token (optional, defaults to BS_CFG.MONDAY_API_TOKEN)
 * @returns {object} Parsed JSON response
 */
function mondayApiRequest_(query, apiToken) {
  const token = apiToken || BS_CFG.MONDAY_API_TOKEN;
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': token },
    payload: JSON.stringify({ query: query }),
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch('https://api.monday.com/v2', options);
  return JSON.parse(response.getContentText());
}

/**
 * Generate checksum for tasks data
 * @param {array} tasks - Array of task objects
 * @returns {string} Hash string
 */
function generateTasksChecksum_(tasks) {
  return hashString_(JSON.stringify((tasks || []).map(t => ({ 
    name: t.subject, 
    status: t.status, 
    created: t.created 
  }))));
}

/**
 * Generate checksum for contacts data (sorted for consistency)
 * @param {array} contacts - Array of contact objects
 * @returns {string} Hash string
 */
function generateContactsChecksum_(contacts) {
  return hashString_(JSON.stringify(
    (contacts || [])
      .map(c => ({ name: c.name, email: c.email, status: c.status }))
      .sort((a, b) => (a.name || '').localeCompare(b.name || ''))
  ));
}

/**
 * Generate checksum for meetings data
 * Only includes future meetings - past meetings falling off shouldn't trigger changes
 * @param {array} meetings - Array of meeting objects
 * @returns {string} Hash string
 */
function generateMeetingsChecksum_(meetings) {
  // Filter to only future meetings for checksum
  const futureMeetings = (meetings || []).filter(m => !m.isPast);
  return hashString_(JSON.stringify(futureMeetings.map(m => ({
    title: m.title,
    date: m.date,
    time: m.time
  }))));
}

/**
 * Filter tasks based on vendor source (affiliate vs buyer)
 * Removes inappropriate onboarding tasks
 * @param {array} tasks - Array of task objects
 * @param {string} source - Vendor source (e.g., "Buyers L1M", "affiliates monday.com")
 * @returns {array} Filtered tasks
 */
function filterTasksBySource_(tasks, source) {
  const sourceLower = (source || '').toLowerCase();
  const isAffiliate = sourceLower.includes('affiliate');
  const isBuyer = sourceLower.includes('buyer');
  
  return (tasks || []).filter(task => {
    const project = (task.project || '').toLowerCase();
    if (isAffiliate && project.includes('onboarding - buyer')) return false;
    if (isBuyer && project.includes('onboarding - affiliate')) return false;
    return true;
  });
}

/**
 * Check if a vendor has changes compared to stored checksums
 * Returns detailed change info for use in skip functions
 * 
 * @param {string} vendor - Vendor name
 * @param {number} listRow - Row number in List sheet
 * @param {string} source - Vendor source for task filtering
 * @returns {object} { hasChanges, changeType, data }
 *   - hasChanges: boolean
 *   - changeType: string describing what changed (or 'first_view' or 'unchanged')
 *   - data: object with fetched data for reuse { emails, tasks, contactData, meetings }
 */
function checkVendorForChanges_(vendor, listRow, source) {
  // Check if vendor is flagged - always stop on flagged vendors (fast - no API)
  if (isVendorFlagged_(vendor)) {
    Logger.log(`${vendor}: flagged for review`);
    // Clear the flag since we're stopping here
    setVendorFlag_(vendor, false);
    return {
      hasChanges: true,
      changeType: 'Flagged',
      data: null
    };
  }

  // Get stored checksum data first (fast - just reads from sheet)
  const storedData = getStoredChecksum_(vendor);

  // If no stored data, this is a first view - definitely has changes
  if (!storedData) {
    Logger.log(`${vendor}: no stored checksums - first view`);
    return {
      hasChanges: true,
      changeType: 'First view',
      data: null
    };
  }

  // FAST PATH: Quick thread count check before fetching full email data
  // If count differs from stored, we know there's a change without expensive processing
  const storedEmailChecksum = storedData.emailChecksum || '';
  const storedCountMatch = storedEmailChecksum.match(/^(\d+):/);
  if (storedCountMatch) {
    const storedCount = parseInt(storedCountMatch[1], 10);
    const currentCount = getEmailThreadCountFast_(listRow);

    if (currentCount >= 0 && currentCount !== storedCount) {
      Logger.log(`${vendor}: FAST DETECT - email count changed (${storedCount} → ${currentCount})`);
      return {
        hasChanges: true,
        changeType: `Emails: ${currentCount > storedCount ? 'new emails' : 'emails removed'}`,
        data: null  // Don't fetch full data - loadVendorData will do that
      };
    }
    // Count matches - need to fetch full data for detailed comparison
  }

  // Full email fetch needed for overdue check and detailed checksum
  const emails = getEmailsForVendor_(vendor, listRow) || [];
  const overdueEmails = emails.filter(e => isEmailOverdue_(e));
  if (overdueEmails.length > 0) {
    Logger.log(`${vendor}: has ${overdueEmails.length} overdue email(s)`);
    return {
      hasChanges: true,
      changeType: 'Overdue emails',
      data: { emails }
    };
  }

  // Check emails (most volatile) - emails already fetched above for overdue check
  const newEmailChecksum = generateEmailChecksum_(emails);

  if (storedData.emailChecksum !== newEmailChecksum) {
    // Get human-readable description of what changed
    const changeDesc = getEmailChangeDescription_(storedData.emailChecksum, newEmailChecksum, 0, 0);
    Logger.log(`${vendor}: emails changed (stored=${storedData.emailChecksum}, new=${newEmailChecksum}) - ${changeDesc}`);
    return {
      hasChanges: true,
      changeType: `Emails: ${changeDesc}`,
      data: { emails }
    };
  }

  // Check vendor-label-only checksum (catches emails without 00.received label)
  // Only check if vendor has a stored vendorLabelChecksum - if not, skip this check
  if (storedData.vendorLabelChecksum) {
    const ss = SpreadsheetApp.getActive();
    const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
    const gmailLink = listSh.getRange(listRow, BS_CFG.L_GMAIL_LINK + 1).getValue();
    const newVendorLabelChecksum = generateVendorLabelChecksum_(gmailLink);

    if (newVendorLabelChecksum && storedData.vendorLabelChecksum !== newVendorLabelChecksum) {
      Logger.log(`${vendor}: vendor label emails changed (stored=${storedData.vendorLabelChecksum}, new=${newVendorLabelChecksum})`);
      return {
        hasChanges: true,
        changeType: 'New vendor emails (unlabeled)',
        data: { emails }
      };
    }
  }

  // Check tasks (second most volatile)
  let tasks = getTasksForVendor_(vendor, listRow) || [];
  tasks = filterTasksBySource_(tasks, source);
  const newTasksChecksum = generateTasksChecksum_(tasks);
  
  if (storedData.moduleChecksums && storedData.moduleChecksums.tasks !== newTasksChecksum) {
    Logger.log(`${vendor}: tasks changed`);
    return { 
      hasChanges: true, 
      changeType: 'Tasks changed',
      data: { emails, tasks } 
    };
  }
  
  // Check contacts/notes/status
  const contactData = getVendorContacts_(vendor, listRow);
  const newNotesChecksum = hashString_(JSON.stringify(contactData.notes || ''));
  const newStatusChecksum = hashString_(JSON.stringify(contactData.liveStatus || ''));
  const newContactsChecksum = generateContactsChecksum_(contactData.contacts);
  
  if (storedData.moduleChecksums) {
    if (storedData.moduleChecksums.notes !== newNotesChecksum) {
      Logger.log(`${vendor}: notes changed`);
      return { 
        hasChanges: true, 
        changeType: 'Notes changed',
        data: { emails, tasks, contactData } 
      };
    }
    if (storedData.moduleChecksums.status !== newStatusChecksum) {
      Logger.log(`${vendor}: status changed`);
      return { 
        hasChanges: true, 
        changeType: 'Status changed',
        data: { emails, tasks, contactData } 
      };
    }
    if (storedData.moduleChecksums.contacts !== newContactsChecksum) {
      Logger.log(`${vendor}: contacts changed`);
      return { 
        hasChanges: true, 
        changeType: 'Contacts changed',
        data: { emails, tasks, contactData } 
      };
    }
  }
  
  // Check meetings
  const contactEmails = (contactData.contacts || []).map(c => c.email).filter(e => e && e.includes('@'));
  const meetingsResult = getUpcomingMeetingsForVendor_(vendor, contactEmails);
  const meetings = meetingsResult.meetings || [];
  const newMeetingsChecksum = generateMeetingsChecksum_(meetings);
  
  if (storedData.moduleChecksums && storedData.moduleChecksums.meetings !== newMeetingsChecksum) {
    Logger.log(`${vendor}: meetings changed`);
    return { 
      hasChanges: true, 
      changeType: 'Meetings changed',
      data: { emails, tasks, contactData, meetings } 
    };
  }
  
  // No changes detected
  Logger.log(`${vendor}: unchanged (emails, tasks, notes, status, contacts, meetings all same)`);
  return { 
    hasChanges: false, 
    changeType: 'unchanged',
    data: { emails, tasks, contactData, meetings } 
  };
}

/**
 * Format change type for display in toast notifications
 * @param {string} changeType - Raw change type from checkVendorForChanges_
 * @return {string} Formatted change type with emoji
 */
function formatChangeType_(changeType) {
  const typeMap = {
    'Flagged': '⚑ Flagged for review',
    'Overdue emails': '🔴 Overdue emails need attention',
    'First view': '🆕 First time viewing',
    'Emails changed': '📧 New or updated emails',
    'Tasks changed': '📋 Tasks changed on monday.com',
    'Notes changed': '📝 Notes updated',
    'Status changed': '🔄 Status changed',
    'Contacts changed': '👤 Contacts updated',
    'Meetings changed': '📅 Meetings changed'
  };
  return typeMap[changeType] || changeType;
}

/**
 * Set row background color in List sheet
 * @param {Sheet} listSh - List sheet
 * @param {number} listRow - Row number (1-based)
 * @param {string} color - Background color
 */
function setListRowColor_(listSh, listRow, color) {
  const numCols = listSh.getLastColumn();
  listSh.getRange(listRow, 1, 1, numCols).setBackground(color);
}

/**
 * Skip to next vendor with changes (different checksum)
 */
function skipToNextChanged(trackComeback) {
  const ss = SpreadsheetApp.getActive();
  const bsSh = ss.getSheetByName(BS_CFG.BATTLE_SHEET);
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
  const props = PropertiesService.getScriptProperties();

  if (!bsSh || !listSh) {
    SpreadsheetApp.getUi().alert('Required sheets not found.');
    return;
  }

  // Check if there's an active Skip 5 & Return session
  let skip5Session = null;
  const skip5Str = props.getProperty('BS_SKIP5_SESSION');
  if (skip5Str) {
    try {
      skip5Session = JSON.parse(skip5Str);

      // Check if session is marked complete (ready to return)
      if (skip5Session.complete) {
        ss.toast(`Returning to: ${skip5Session.originVendor}`, '🔄 Skip 5 Complete', 3);
        Utilities.sleep(500);
        loadVendorData(skip5Session.originIdx);
        props.deleteProperty('BS_SKIP5_SESSION');

        SpreadsheetApp.getUi().alert(
          'Skip 5 & Return Complete',
          `Found ${skip5Session.changedFound} changed vendor(s).\n` +
          `Skipped ${skip5Session.skippedCount || 0} unchanged vendor(s).\n\n` +
          `Returned to: ${skip5Session.originVendor}`,
          SpreadsheetApp.getUi().ButtonSet.OK
        );
        return;
      }
    } catch (e) {
      props.deleteProperty('BS_SKIP5_SESSION');
      skip5Session = null;
    }
  }

  // If trackComeback wasn't explicitly passed, check if there's an existing comeback pending
  if (trackComeback === undefined) {
    const comebackStr = props.getProperty('BS_COMEBACK');
    if (comebackStr) {
      trackComeback = true;  // Auto-enable if comeback is pending
    }
  }

  const listData = listSh.getDataRange().getValues();
  const totalVendors = listData.length - 1;

  // Get current index using the same function as other navigation
  let currentIdx = getCurrentVendorIndex_() || 1;

  let skippedCount = 0;

  // Show progress message based on session state
  if (skip5Session) {
    ss.toast(`Skip 5 & Return: ${skip5Session.changedFound}/${skip5Session.changedTarget} found. Searching...`, '🔄 Skip 5', -1);
  } else {
    ss.toast('Searching for changed vendors...', '⏭️ Skipping', -1);
  }

  // Loop through vendors looking for one with changes - start from NEXT vendor
  while (true) {
    currentIdx++;  // Move to next vendor FIRST

    // Stop at end
    if (currentIdx > totalVendors) {
      ss.toast('');

      // If Skip 5 session is active, return to origin
      if (skip5Session) {
        ss.toast(`Returning to: ${skip5Session.originVendor}`, '🔄 End of List', 3);
        Utilities.sleep(500);
        loadVendorData(skip5Session.originIdx);
        props.deleteProperty('BS_SKIP5_SESSION');

        SpreadsheetApp.getUi().alert(
          'Skip 5 & Return Complete',
          `Reached end of list.\n` +
          `Found ${skip5Session.changedFound} changed vendor(s).\n` +
          `Skipped ${skippedCount + (skip5Session.skippedCount || 0)} unchanged vendor(s).\n\n` +
          `Returned to: ${skip5Session.originVendor}`,
          SpreadsheetApp.getUi().ButtonSet.OK
        );
        return;
      }

      SpreadsheetApp.getUi().alert(`Checked all remaining vendors.\nSkipped ${skippedCount} unchanged vendor(s).\nNo more vendors with changes found.`);
      return;
    }

    const vendor = listData[currentIdx][BS_CFG.L_VENDOR];
    const source = listData[currentIdx][BS_CFG.L_SOURCE] || '';
    const listRow = currentIdx + 1;

    if (skip5Session) {
      ss.toast(`Checking ${vendor}... (${skip5Session.changedFound}/${skip5Session.changedTarget}, ${skippedCount} skipped)`, '🔄 Skip 5', -1);
    } else {
      ss.toast(`Checking ${vendor}... (${skippedCount} skipped so far)`, '⏭️ Skipping', -1);
    }

    // Safeguard: After 30 skips, stop on this vendor to avoid 6-minute timeout
    const MAX_SKIPS = 30;
    if (skippedCount >= MAX_SKIPS) {
      ss.toast(`Stopped after ${skippedCount} skips to avoid timeout`, '⚠️ Safeguard', 5);

      // Update Skip 5 session if active
      if (skip5Session) {
        skip5Session.skippedCount = (skip5Session.skippedCount || 0) + skippedCount;
        props.setProperty('BS_SKIP5_SESSION', JSON.stringify(skip5Session));
      }

      // Load this vendor with a note about the safeguard
      loadVendorData(currentIdx, { forceChanged: false, changeType: `No changes - stopped after ${skippedCount} skips (timeout safeguard)` });
      setListRowColor_(listSh, listRow, BS_CFG.COLOR_ROW_SKIPPED);
      if (trackComeback) checkComeback_();
      return;
    }

    // Use the centralized change detection helper
    const changeResult = checkVendorForChanges_(vendor, listRow, source);

    if (changeResult.hasChanges) {
      // If Skip 5 session is active, update it
      if (skip5Session) {
        skip5Session.changedFound++;
        skip5Session.currentIdx = currentIdx;
        skip5Session.skippedCount = (skip5Session.skippedCount || 0) + skippedCount;

        // Check if we've found all 5
        if (skip5Session.changedFound >= skip5Session.changedTarget) {
          skip5Session.complete = true;
          ss.toast(`Found ${skip5Session.changedFound}/${skip5Session.changedTarget}! Next skip will return to ${skip5Session.originVendor}`, '🔄 5/5 Found', 5);
        } else {
          ss.toast(`Skip 5: ${skip5Session.changedFound}/${skip5Session.changedTarget} - ${vendor}`, '🔄 Skip 5', 3);
        }

        props.setProperty('BS_SKIP5_SESSION', JSON.stringify(skip5Session));
      }

      // Show what changed in a toast BEFORE loading modules
      const changeLabel = formatChangeType_(changeResult.changeType);
      ss.toast(`${vendor}\n${changeLabel}`, '🔔 Change Detected', 5);

      loadVendorData(currentIdx, { forceChanged: true, changeType: changeResult.changeType });
      setListRowColor_(listSh, listRow, BS_CFG.COLOR_ROW_CHANGED);
      // WHAT CHANGED section now shows the details, no need for dialog
      if (trackComeback) checkComeback_();
      return;
    }

    // No changes - mark as skipped (yellow)
    setListRowColor_(listSh, listRow, BS_CFG.COLOR_ROW_SKIPPED);
    skippedCount++;
  }
}

/**
 * Skip with Comeback - Skip to next changed vendor but mark current vendor to revisit
 * After N vendors are viewed, automatically returns to the marked vendor
 */
function skipWithComeback() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
  const checkSh = ss.getSheetByName(BS_CFG.CHECKSUMS_SHEET);
  
  if (!listSh || !checkSh) {
    ui.alert('Error', 'Required sheets not found.', ui.ButtonSet.OK);
    return;
  }
  
  const currentIdx = getCurrentVendorIndex_();
  if (!currentIdx || currentIdx < 1) {
    ui.alert('Error', 'No vendor currently loaded.', ui.ButtonSet.OK);
    return;
  }
  
  const listRow = currentIdx + 1;
  const vendor = listSh.getRange(listRow, 1).getValue();
  
  // Ask how many vendors to see before coming back
  const response = ui.prompt(
    '⏰ Skip with Comeback',
    `Current vendor: ${vendor}\n\nHow many vendors do you want to review before coming back to this one?`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const countStr = response.getResponseText().trim();
  const comebackAfter = parseInt(countStr, 10);
  
  if (isNaN(comebackAfter) || comebackAfter < 1 || comebackAfter > 50) {
    ui.alert('Error', 'Please enter a number between 1 and 50.', ui.ButtonSet.OK);
    return;
  }
  
  // Store comeback info in script properties
  const props = PropertiesService.getScriptProperties();
  const comebackData = {
    vendorIndex: currentIdx,
    vendorName: vendor,
    comebackAfter: comebackAfter,
    vendorsSeen: 0,
    setAt: new Date().toISOString()
  };
  props.setProperty('BS_COMEBACK', JSON.stringify(comebackData));
  
  // Color the comeback vendor row with a distinct color (light purple)
  const numCols = listSh.getLastColumn();
  listSh.getRange(listRow, 1, 1, numCols).setBackground('#e1d5e7');
  
  Logger.log(`Comeback set for ${vendor} after ${comebackAfter} vendors`);
  ss.toast(`Will return to ${vendor} after ${comebackAfter} vendors`, '⏰ Comeback Set', 3);
  
  // Now run skipToNextChanged with comeback tracking enabled
  skipToNextChanged(true);
}

/**
 * Check and handle comeback when using regular Skip Unchanged
 * Call this at the end of skipToNextChanged or wrap it
 */
function checkComeback_() {
  const props = PropertiesService.getScriptProperties();
  const comebackStr = props.getProperty('BS_COMEBACK');
  
  if (!comebackStr) return false;
  
  try {
    const comebackData = JSON.parse(comebackStr);
    comebackData.vendorsSeen++;
    
    const ss = SpreadsheetApp.getActive();
    const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
    
    Logger.log(`Comeback tracker: ${comebackData.vendorsSeen}/${comebackData.comebackAfter} vendors seen`);
    
    if (comebackData.vendorsSeen >= comebackData.comebackAfter) {
      // Time to go back!
      props.deleteProperty('BS_COMEBACK');
      
      // Clear the purple highlight
      const numCols = listSh.getLastColumn();
      listSh.getRange(comebackData.vendorIndex + 1, 1, 1, numCols).setBackground(null);
      
      // Load the comeback vendor
      ss.toast(`Returning to ${comebackData.vendorName}...`, '⏰ Comeback!', 2);
      loadVendorData(comebackData.vendorIndex);
      
      SpreadsheetApp.getUi().alert('⏰ Comeback!', `Reviewed ${comebackData.vendorsSeen} vendors.\nNow returning to: ${comebackData.vendorName}`, SpreadsheetApp.getUi().ButtonSet.OK);
      return true; // Indicates comeback was triggered
    } else {
      // Update counter
      props.setProperty('BS_COMEBACK', JSON.stringify(comebackData));
      const remaining = comebackData.comebackAfter - comebackData.vendorsSeen;
      ss.toast(`${remaining} more vendor(s) until returning to ${comebackData.vendorName}`, '⏰ Comeback', 2);
      return false;
    }
  } catch (e) {
    props.deleteProperty('BS_COMEBACK');
    return false;
  }
}

/**
 * Cancel any pending comeback
 */
function cancelComeback() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
  const props = PropertiesService.getScriptProperties();
  
  const comebackStr = props.getProperty('BS_COMEBACK');
  
  if (!comebackStr) {
    ui.alert('No Comeback', 'There is no pending comeback to cancel.', ui.ButtonSet.OK);
    return;
  }
  
  try {
    const comebackData = JSON.parse(comebackStr);
    
    // Clear the purple highlight
    const numCols = listSh.getLastColumn();
    listSh.getRange(comebackData.vendorIndex + 1, 1, 1, numCols).setBackground(null);
    
    props.deleteProperty('BS_COMEBACK');
    
    ui.alert('Comeback Cancelled', `Cancelled comeback to: ${comebackData.vendorName}`, ui.ButtonSet.OK);
  } catch (e) {
    props.deleteProperty('BS_COMEBACK');
    ui.alert('Comeback Cancelled', 'Pending comeback has been cleared.', ui.ButtonSet.OK);
  }
}

/**
 * View current comeback status
 */
function viewComebackStatus() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  
  const comebackStr = props.getProperty('BS_COMEBACK');
  
  if (!comebackStr) {
    ui.alert('Comeback Status', 'No comeback currently scheduled.', ui.ButtonSet.OK);
    return;
  }
  
  try {
    const comebackData = JSON.parse(comebackStr);
    const remaining = comebackData.comebackAfter - comebackData.vendorsSeen;
    
    ui.alert('⏰ Comeback Status', 
      `Vendor: ${comebackData.vendorName}\n` +
      `Vendors seen: ${comebackData.vendorsSeen}/${comebackData.comebackAfter}\n` +
      `Remaining: ${remaining} vendor(s)\n` +
      `Set at: ${comebackData.setAt}`,
      ui.ButtonSet.OK);
  } catch (e) {
    props.deleteProperty('BS_COMEBACK');
    ui.alert('Comeback Status', 'No valid comeback scheduled.', ui.ButtonSet.OK);
  }
}

/**
 * Skip to next changed vendor 5 times, then return to the original
 * Useful for putting things in motion and letting them process while reviewing others
 * This function initializes a session and delegates to skipToNextChanged()
 */
function skip5AndReturn() {
  const ss = SpreadsheetApp.getActive();
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
  const props = PropertiesService.getScriptProperties();

  if (!listSh) {
    SpreadsheetApp.getUi().alert('Required sheets not found.');
    return;
  }

  // Check if we're already in a "Skip 5 & Return" session
  const sessionStr = props.getProperty('BS_SKIP5_SESSION');

  if (!sessionStr) {
    // Start a new session
    const originIdx = getCurrentVendorIndex_();

    if (!originIdx) {
      SpreadsheetApp.getUi().alert('Could not determine current vendor index.');
      return;
    }

    const originVendor = listSh.getRange(originIdx + 1, 1).getValue();

    const session = {
      originIdx: originIdx,
      originVendor: originVendor,
      currentIdx: originIdx,
      changedFound: 0,
      changedTarget: 5,
      skippedCount: 0,
      startedAt: new Date().toISOString()
    };

    props.setProperty('BS_SKIP5_SESSION', JSON.stringify(session));
    ss.toast(`Starting Skip 5 & Return from: ${originVendor}`, '🔄 Skip 5 & Return', 3);
  }

  // Delegate to skipToNextChanged which now handles Skip 5 sessions
  skipToNextChanged();
}

/**
 * Continue or complete a Skip 5 & Return session
 * Called automatically or can be called manually
 */
function continueSkip5AndReturn() {
  const props = PropertiesService.getScriptProperties();
  const sessionStr = props.getProperty('BS_SKIP5_SESSION');
  
  if (!sessionStr) {
    SpreadsheetApp.getUi().alert('No Skip 5 & Return session active.\n\nUse "Skip 5 & Return" to start a new session.');
    return;
  }
  
  let session;
  try {
    session = JSON.parse(sessionStr);
  } catch (e) {
    props.deleteProperty('BS_SKIP5_SESSION');
    SpreadsheetApp.getUi().alert('Session data corrupted. Please start a new session.');
    return;
  }
  
  // If session is complete, return to origin
  if (session.complete) {
    const ss = SpreadsheetApp.getActive();
    ss.toast(`Returning to: ${session.originVendor}`, '🔄 Returning', 2);
    Utilities.sleep(500);
    loadVendorData(session.originIdx);
    
    // Clear the session
    props.deleteProperty('BS_SKIP5_SESSION');
    
    ss.toast(`Back to ${session.originVendor} after ${session.changedFound} changed, ${session.skippedCount} skipped`, '✅ Done', 3);
    return;
  }
  
  // Otherwise, continue searching
  skip5AndReturn();
}

/**
 * Cancel an active Skip 5 & Return session
 */
function cancelSkip5Session() {
  const props = PropertiesService.getScriptProperties();
  const sessionStr = props.getProperty('BS_SKIP5_SESSION');
  
  if (!sessionStr) {
    SpreadsheetApp.getUi().alert('No Skip 5 & Return session active.');
    return;
  }
  
  props.deleteProperty('BS_SKIP5_SESSION');
  SpreadsheetApp.getUi().alert('Skip 5 & Return session cancelled.');
}

/**
 * Auto-traverse through vendors one at a time, loading each on the page
 * Unlike Skip Unchanged, this loads EVERY vendor (not just changed ones)
 * Colors List rows: yellow=no changes, green=has changes
 * Stops after 30 vendors to avoid timeout (session tracked in properties)
 * Press the button repeatedly to advance through vendors
 */
function autoTraverseVendors() {
  const ss = SpreadsheetApp.getActive();
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
  const props = PropertiesService.getScriptProperties();

  if (!listSh) {
    ss.toast('List sheet not found.', '❌ Error', 3);
    return;
  }

  const listData = listSh.getDataRange().getValues();
  const totalVendors = listData.length - 1;

  // Check for existing auto-traverse session
  let traverseSession = null;
  const sessionData = props.getProperty('BS_AUTOTRAVERSE_SESSION');
  if (sessionData) {
    try {
      traverseSession = JSON.parse(sessionData);
    } catch (e) {
      traverseSession = null;
    }
  }

  // Get next vendor index (start from current + 1, or continue session)
  let currentIdx = (getCurrentVendorIndex_() || 0) + 1;
  let traverseCount = 0;

  if (traverseSession) {
    traverseCount = traverseSession.count || 0;
  }

  // Make sure we don't start past the end
  if (currentIdx > totalVendors) {
    ss.toast('Already at the last vendor.', '🔁 Auto-Traverse', 3);
    props.deleteProperty('BS_AUTOTRAVERSE_SESSION');
    return;
  }

  // Safeguard: After 30 vendors traversed, show warning and reset
  const MAX_TRAVERSE = 30;
  if (traverseCount >= MAX_TRAVERSE) {
    ss.toast(`Session ended: ${traverseCount} vendors traversed (limit reached)`, '⚠️ Auto-Traverse Limit', 5);
    props.deleteProperty('BS_AUTOTRAVERSE_SESSION');
    traverseCount = 0;
  }

  const vendor = listData[currentIdx][BS_CFG.L_VENDOR];
  const source = listData[currentIdx][BS_CFG.L_SOURCE] || '';
  const listRow = currentIdx + 1;

  // Use the centralized change detection helper
  const changeResult = checkVendorForChanges_(vendor, listRow, source);

  // Increment traverse count
  traverseCount++;

  // Update session
  props.setProperty('BS_AUTOTRAVERSE_SESSION', JSON.stringify({
    count: traverseCount,
    startedAt: traverseSession?.startedAt || new Date().toISOString()
  }));

  if (changeResult.hasChanges) {
    // Show what changed
    const changeLabel = formatChangeType_(changeResult.changeType);
    ss.toast(`${vendor} (${traverseCount}/${MAX_TRAVERSE})\n${changeLabel}`, '🔁 Auto-Traverse', 3);

    loadVendorData(currentIdx, { forceChanged: true, changeType: changeResult.changeType });
    setListRowColor_(listSh, listRow, BS_CFG.COLOR_ROW_CHANGED);
  } else {
    // No changes - still load the vendor, but mark as no changes
    ss.toast(`${vendor} (${traverseCount}/${MAX_TRAVERSE})\nNo changes`, '🔁 Auto-Traverse', 3);

    loadVendorData(currentIdx, { forceChanged: false, changeType: 'No changes detected' });
    setListRowColor_(listSh, listRow, BS_CFG.COLOR_ROW_SKIPPED);
  }

  // Check if we've reached the end of the list
  if (currentIdx >= totalVendors) {
    ss.toast(`Reached end of list! ${traverseCount} vendors traversed.`, '✅ Auto-Traverse Complete', 5);
    props.deleteProperty('BS_AUTOTRAVERSE_SESSION');
  }
}

/**
 * Reset the Auto-Traverse session counter
 */
function resetAutoTraverseSession() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('BS_AUTOTRAVERSE_SESSION');
  SpreadsheetApp.getActive().toast('Auto-Traverse session reset.', '🔁 Reset', 3);
}

/**
 * TEST FUNCTION: Debug Google Drive folder search
 * Run this from the Apps Script editor to see why folders aren't being found
 */
function testGDriveSearch() {
  const vendorName = 'Ion Solar';  // Changed to test Ion Solar
  const parentFolderId = BS_CFG.GDRIVE_VENDORS_FOLDER_ID;
  
  Logger.log('='.repeat(60));
  Logger.log('GOOGLE DRIVE DEBUG TEST');
  Logger.log('='.repeat(60));
  Logger.log(`Vendor: ${vendorName}`);
  Logger.log(`Parent Folder ID: ${parentFolderId}`);
  Logger.log('');
  
  // Test 1: Can we access the parent folder?
  Logger.log('--- TEST 1: Access Parent Folder ---');
  try {
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    Logger.log(`✅ Parent folder found: ${parentFolder.getName()}`);
    Logger.log(`   URL: ${parentFolder.getUrl()}`);
  } catch (e) {
    Logger.log(`❌ Cannot access parent folder: ${e.message}`);
    return;
  }
  
  // Test 2: List ALL subfolders in parent (first 20)
  Logger.log('');
  Logger.log('--- TEST 2: List Subfolders (first 20) ---');
  try {
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    const folders = parentFolder.getFolders();
    let count = 0;
    while (folders.hasNext() && count < 20) {
      const folder = folders.next();
      Logger.log(`  ${count + 1}. "${folder.getName()}" (ID: ${folder.getId()})`);
      count++;
    }
    Logger.log(`   Total shown: ${count}`);
  } catch (e) {
    Logger.log(`❌ Error listing folders: ${e.message}`);
  }
  
  // Test 3: Search for folder using searchFolders
  Logger.log('');
  Logger.log('--- TEST 3: searchFolders() with vendor name ---');
  try {
    const query1 = `title contains '${vendorName}' and '${parentFolderId}' in parents`;
    Logger.log(`Query: ${query1}`);
    const results1 = DriveApp.searchFolders(query1);
    let found1 = 0;
    while (results1.hasNext()) {
      const folder = results1.next();
      Logger.log(`  ✅ Found: "${folder.getName()}" (ID: ${folder.getId()})`);
      found1++;
    }
    if (found1 === 0) Logger.log('  ❌ No results');
  } catch (e) {
    Logger.log(`❌ Error: ${e.message}`);
  }
  
  // Test 4: Search without parent restriction
  Logger.log('');
  Logger.log('--- TEST 4: searchFolders() without parent restriction ---');
  try {
    const query2 = `title contains '${vendorName}' and mimeType = 'application/vnd.google-apps.folder'`;
    Logger.log(`Query: ${query2}`);
    const results2 = DriveApp.searchFolders(query2);
    let found2 = 0;
    while (results2.hasNext() && found2 < 10) {
      const folder = results2.next();
      Logger.log(`  Found: "${folder.getName()}" (ID: ${folder.getId()})`);
      found2++;
    }
    if (found2 === 0) Logger.log('  ❌ No results');
  } catch (e) {
    Logger.log(`❌ Error: ${e.message}`);
  }
  
  // Test 5: Try to access the known folder directly
  Logger.log('');
  Logger.log('--- TEST 5: Direct access to known SunPower folder ---');
  const knownFolderId = '1UbRPIvG6ZsO3nNJ12h2YgG0W5fJzR5ae';
  try {
    const folder = DriveApp.getFolderById(knownFolderId);
    Logger.log(`✅ Direct access works: "${folder.getName()}"`);
    Logger.log(`   URL: ${folder.getUrl()}`);
    
    // Check its parent
    const parents = folder.getParents();
    while (parents.hasNext()) {
      const parent = parents.next();
      Logger.log(`   Parent: "${parent.getName()}" (ID: ${parent.getId()})`);
    }
  } catch (e) {
    Logger.log(`❌ Cannot access folder directly: ${e.message}`);
  }
  
  // Test 6: List files in the known folder
  Logger.log('');
  Logger.log('--- TEST 6: Files in known SunPower folder ---');
  try {
    const folder = DriveApp.getFolderById(knownFolderId);
    const files = folder.getFiles();
    let fileCount = 0;
    while (files.hasNext() && fileCount < 10) {
      const file = files.next();
      Logger.log(`  ${fileCount + 1}. "${file.getName()}"`);
      fileCount++;
    }
    Logger.log(`   Total shown: ${fileCount}`);
  } catch (e) {
    Logger.log(`❌ Error: ${e.message}`);
  }
  
  Logger.log('');
  Logger.log('='.repeat(60));
  Logger.log('TEST COMPLETE - Check logs above');
  Logger.log('='.repeat(60));
}

/************************************************************
 * MONDAY.COM BOARD SYNC
 * Pull data directly from monday.com boards to update
 * the affiliates monday.com and buyers monday.com sheets
 ************************************************************/

/**
 * Main function to sync monday.com boards to sheets
 */
function syncMondayComBoards() {
  const ss = SpreadsheetApp.getActive();

  ss.toast('Starting sync...', '🔄 Syncing', 30);
  
  try {
    // Sync Buyers board
    ss.toast('Syncing Buyers board...', '🔄 Syncing', 30);
    const buyersResult = syncBuyersBoard_();
    Logger.log(`Buyers sync complete: ${buyersResult.count} items`);
    
    // Sync Affiliates board
    ss.toast('Syncing Affiliates board...', '🔄 Syncing', 30);
    const affiliatesResult = syncAffiliatesBoard_();
    Logger.log(`Affiliates sync complete: ${affiliatesResult.count} items`);
    
    ss.toast(`Sync complete! Buyers: ${buyersResult.count}, Affiliates: ${affiliatesResult.count}`, '✅ Done', 5);
    
  } catch (e) {
    Logger.log(`Sync error: ${e.message}`);
    ss.toast(`Error: ${e.message}`, '❌ Error', 5);
  }
}

/**
 * Sync Buyers board to buyers monday.com sheet
 */
function syncBuyersBoard_() {
  const ss = SpreadsheetApp.getActive();
  const sheetName = 'buyers monday.com';
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  
  // Column IDs from testGetBoardColumns() output
  const columns = [
    { id: 'name', header: 'Name', type: 'name' },
    { id: 'subtasks_mktcnxe', header: 'Subitems', type: 'subtasks' },
    { id: 'color_mktqyter', header: 'Buyer Type', type: 'status' },
    { id: 'text_mkqnvsqh', header: 'BUYER NOTES MAIN', type: 'text' },
    { id: 'rating_mkv5ztjb', header: 'Opportunity Size', type: 'rating' },
    { id: 'tag_mkskgt84', header: 'Live Verticals', type: 'tags' },
    { id: 'tag_mkskewmq', header: 'Other Verticals', type: 'tags' },
    { id: 'board_relation_mky0bt0z', header: 'Contacts', type: 'board_relation' },
    { id: 'tag_mkskfmf3', header: 'Live Modalities', type: 'tags' },
    { id: 'tag_mkskassa', header: 'Other Modalities', type: 'tags' },
    { id: 'link_mksmwprd', header: 'Phonexa Link', type: 'link' },
    { id: 'link_mksmgg2h', header: 'Company URL', type: 'link' },
    { id: 'pulse_log_mkthvn03', header: 'Creation log', type: 'creation_log' },
    { id: 'board_relation_mkvd98v7', header: 'Sourcing', type: 'board_relation' },
    { id: 'text_mkvkr178', header: 'Other Name', type: 'text' },
    { id: 'pulse_updated_mkvqtmew', header: 'Last updated', type: 'last_updated' },
    { id: 'numeric_mkwp5np4', header: 'Rank', type: 'numbers' },
    { id: 'board_relation_mky0j7qj', header: 'link to Helpful Links', type: 'board_relation' },
    { id: 'dropdown_mkyam4qw', header: 'State(s)', type: 'dropdown' },
    { id: 'dropdown_mkyazy2j', header: 'Dead States(s)', type: 'dropdown' }
  ];
  
  return syncBoardToSheet_(BS_CFG.BUYERS_BOARD_ID, sheet, columns, 'Buyers');
}

/**
 * Sync Affiliates board to affiliates monday.com sheet
 */
function syncAffiliatesBoard_() {
  const ss = SpreadsheetApp.getActive();
  const sheetName = 'affiliates monday.com';
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  
  // Column IDs from testGetBoardColumns() output
  const columns = [
    { id: 'name', header: 'Name', type: 'name' },
    { id: 'multiple_person_mkt9k1n1', header: 'AM', type: 'people' },
    { id: 'text_mkrdahqz', header: 'AFFILIATE NOTES MAIN', type: 'text' },
    { id: 'rating_mkv5k453', header: 'Opportunity Size', type: 'rating' },
    { id: 'board_relation_mky0n0rf', header: 'Contacts', type: 'board_relation' },
    { id: 'tag_mksm70xw', header: 'Traffic Sources', type: 'tags' },
    { id: 'tag_mkskrddx', header: 'Live Verticals', type: 'tags' },
    { id: 'tag_mkskfs70', header: 'Other Verticals', type: 'tags' },
    { id: 'tag_mksk7whx', header: 'Live Modalities', type: 'tags' },
    { id: 'tag_mkskkszw', header: 'Other Modalities', type: 'tags' },
    { id: 'link_mksmgnc0', header: 'Phonexa Link', type: 'link' },
    { id: 'text_mksmcrpw', header: 'Other Name', type: 'text' },
    { id: 'board_relation_mksnsmsg', header: 'Sourcing', type: 'board_relation' },
    { id: 'color_mksy6tak', header: 'Priority', type: 'status' },
    { id: 'board_relation_mkthwkt6', header: 'URLs - Affiliates', type: 'board_relation' },
    { id: 'subtasks_mkvgk8ab', header: 'Subitems', type: 'subtasks' },
    { id: 'pulse_updated_mkvq53b1', header: 'Last updated', type: 'last_updated' },
    { id: 'boolean_mkxb61bz', header: 'imp', type: 'checkbox' },
    { id: 'board_relation_mky04azc', header: 'link to Helpful Links', type: 'board_relation' }
  ];
  
  return syncBoardToSheet_(BS_CFG.AFFILIATES_BOARD_ID, sheet, columns, 'Affiliates');
}

/**
 * Generic function to sync a monday.com board to a sheet
 */
function syncBoardToSheet_(boardId, sheet, columns, boardName) {
  const apiToken = BS_CFG.MONDAY_API_TOKEN;
  
  // Build column IDs for query (excluding 'name' which is automatic)
  const columnIds = columns
    .filter(c => c.id !== 'name')
    .map(c => `"${c.id}"`)
    .join(', ');
  
  // Query all items with pagination
  const allItems = [];
  let cursor = null;
  let pageCount = 0;
  const maxPages = 20; // Safety limit
  
  do {
    pageCount++;
    Logger.log(`Fetching ${boardName} page ${pageCount}...`);
    
    const cursorPart = cursor ? `, cursor: "${cursor}"` : '';
    const query = `
      query {
        boards(ids: [${boardId}]) {
          groups {
            id
            title
          }
          items_page(limit: 500${cursorPart}) {
            cursor
            items {
              id
              name
              group {
                id
                title
              }
              column_values(ids: [${columnIds}]) {
                id
                text
                value
                ... on MirrorValue {
                  display_value
                }
                ... on BoardRelationValue {
                  display_value
                }
                ... on StatusValue {
                  label
                }
                ... on DropdownValue {
                  text
                }
                ... on TagsValue {
                  text
                }
                ... on PersonValue {
                  text
                }
              }
            }
          }
        }
      }
    `;
    
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': apiToken },
      payload: JSON.stringify({ query: query }),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch('https://api.monday.com/v2', options);
    const result = JSON.parse(response.getContentText());
    
    if (result.errors) {
      Logger.log(`API Error: ${JSON.stringify(result.errors)}`);
      throw new Error(result.errors[0].message);
    }
    
    const itemsPage = result.data?.boards?.[0]?.items_page;
    if (!itemsPage) {
      Logger.log('No items_page in response');
      break;
    }
    
    allItems.push(...itemsPage.items);
    cursor = itemsPage.cursor;
    
    Logger.log(`  Fetched ${itemsPage.items.length} items (total: ${allItems.length})`);
    
  } while (cursor && pageCount < maxPages);
  
  Logger.log(`Total ${boardName} items fetched: ${allItems.length}`);
  
  // Build sheet data
  const headers = columns.map(c => c.header);
  // Add Group column at the end
  headers.push('Group');
  
  const rows = allItems.map(item => {
    const row = columns.map(col => {
      if (col.id === 'name') {
        return item.name;
      }
      
      const colValue = item.column_values.find(cv => cv.id === col.id);
      if (!colValue) return '';
      
      // Extract value based on type
      switch (col.type) {
        case 'status':
          return colValue.label || colValue.text || '';
        case 'board_relation':
          return colValue.display_value || colValue.text || '';
        case 'tags':
        case 'dropdown':
        case 'people':
          return colValue.text || '';
        case 'checkbox':
          // Parse checkbox JSON to get checked state
          try {
            const checkData = JSON.parse(colValue.value || '{}');
            return checkData.checked === 'true' || checkData.checked === true ? 'Yes' : '';
          } catch {
            return colValue.text || '';
          }
        case 'rating':
          // Parse rating JSON to get rating value
          try {
            const ratingData = JSON.parse(colValue.value || '{}');
            return ratingData.rating || colValue.text || '';
          } catch {
            return colValue.text || '';
          }
        case 'link':
          // Parse link JSON to get URL
          try {
            const linkData = JSON.parse(colValue.value || '{}');
            return linkData.url || colValue.text || '';
          } catch {
            return colValue.text || '';
          }
        case 'creation_log':
        case 'last_updated':
          // Parse to get date
          try {
            const logData = JSON.parse(colValue.value || '{}');
            if (logData.created_at) {
              return formatMondayDate_(logData.created_at);
            }
            if (logData.updated_at) {
              return formatMondayDate_(logData.updated_at);
            }
            return colValue.text || '';
          } catch {
            return colValue.text || '';
          }
        default:
          return colValue.text || '';
      }
    });
    
    // Add group name
    row.push(item.group?.title || '');
    
    return row;
  });
  
  // Clear sheet and write data
  sheet.clear();
  
  // Write headers
  if (headers.length > 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#4a86e8')
      .setFontColor('white');
  }
  
  // Write data rows
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  
  // Auto-resize columns
  for (let i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }
  
  // Freeze header row
  sheet.setFrozenRows(1);
  
  return { count: rows.length };
}

/**
 * Format monday.com date to readable format
 */
function formatMondayDate_(dateStr) {
  try {
    const date = new Date(dateStr);
    // Format as "MMM D, YYYY H:MM AM/PM"
    return Utilities.formatDate(date, 'America/Los_Angeles', 'MMM d, yyyy h:mm a');
  } catch {
    return dateStr;
  }
}

/**
 * Test function to get column IDs from a board
 * Run this to discover column IDs for new boards
 */
function testGetBoardColumns() {
  const boardId = BS_CFG.AFFILIATES_BOARD_ID; // Change to BUYERS_BOARD_ID or AFFILIATES_BOARD_ID
  const apiToken = BS_CFG.MONDAY_API_TOKEN;
  
  const query = `
    query {
      boards(ids: [${boardId}]) {
        name
        columns {
          id
          title
          type
        }
      }
    }
  `;
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': apiToken },
    payload: JSON.stringify({ query: query }),
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch('https://api.monday.com/v2', options);
  const result = JSON.parse(response.getContentText());
  
  Logger.log('Board: ' + result.data?.boards?.[0]?.name);
  Logger.log('Columns:');
  
  const columns = result.data?.boards?.[0]?.columns || [];
  columns.forEach(col => {
    Logger.log(`  { id: '${col.id}', header: '${col.title}', type: '${col.type}' },`);
  });
}

/************************************************************
 * DUPLICATE VENDOR CHECK
 * Identify vendors appearing in both Buyers and Affiliates boards
 ************************************************************/

/**
 * Check for vendors appearing in both Buyers and Affiliates
 * Also checks for duplicates within the same board
 * Shows results and allows adding to exclusion list
 */
function checkDuplicateVendors() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  
  ss.toast('Checking for duplicate vendors...', '🔍 Scanning', 5);
  
  // Get vendors from both sheets
  const buyersSheet = ss.getSheetByName('buyers monday.com');
  const affiliatesSheet = ss.getSheetByName('affiliates monday.com');
  
  if (!buyersSheet || !affiliatesSheet) {
    ui.alert('Error', 'Please run "Sync monday.com Data" first to populate the buyers and affiliates sheets.', ui.ButtonSet.OK);
    return;
  }
  
  // Get vendor names from column A (skip header)
  const buyersData = buyersSheet.getRange(2, 1, buyersSheet.getLastRow() - 1, 1).getValues();
  const affiliatesData = affiliatesSheet.getRange(2, 1, affiliatesSheet.getLastRow() - 1, 1).getValues();
  
  // Find duplicates within each board
  const buyerDuplicatesWithinBoard = findDuplicatesInList_(buyersData);
  const affiliateDuplicatesWithinBoard = findDuplicatesInList_(affiliatesData);
  
  const buyerNames = new Set(buyersData.map(row => String(row[0]).trim().toLowerCase()).filter(n => n));
  const affiliateNames = new Set(affiliatesData.map(row => String(row[0]).trim().toLowerCase()).filter(n => n));
  
  // Get exclusion list from Settings
  const exclusions = getDuplicateExclusions_();
  const exclusionSet = new Set(exclusions.map(e => e.toLowerCase()));
  
  // Find cross-board duplicates (in both, not in exclusion list)
  const crossBoardDuplicates = [];
  for (const name of buyerNames) {
    if (affiliateNames.has(name) && !exclusionSet.has(name)) {
      // Get the proper case version
      const properCase = buyersData.find(row => String(row[0]).trim().toLowerCase() === name)?.[0] || name;
      crossBoardDuplicates.push(properCase);
    }
  }
  
  // Also find excluded ones for reference
  const excludedDuplicates = [];
  for (const name of buyerNames) {
    if (affiliateNames.has(name) && exclusionSet.has(name)) {
      const properCase = buyersData.find(row => String(row[0]).trim().toLowerCase() === name)?.[0] || name;
      excludedDuplicates.push(properCase);
    }
  }
  
  Logger.log(`Found ${crossBoardDuplicates.length} cross-board duplicates (${excludedDuplicates.length} excluded)`);
  Logger.log(`Found ${buyerDuplicatesWithinBoard.length} duplicates within Buyers board`);
  Logger.log(`Found ${affiliateDuplicatesWithinBoard.length} duplicates within Affiliates board`);
  
  const totalIssues = crossBoardDuplicates.length + buyerDuplicatesWithinBoard.length + affiliateDuplicatesWithinBoard.length;
  
  if (totalIssues === 0) {
    let message = '✅ No duplicate vendors found!';
    if (excludedDuplicates.length > 0) {
      message += `\n\n(${excludedDuplicates.length} known duplicates are excluded via Settings)`;
    }
    ui.alert('Duplicate Check Complete', message, ui.ButtonSet.OK);
    return;
  }
  
  // Show results
  let message = '';
  
  if (crossBoardDuplicates.length > 0) {
    message += `⚠️ ${crossBoardDuplicates.length} vendor(s) in BOTH Buyers and Affiliates:\n`;
    message += crossBoardDuplicates.slice(0, 10).join(', ');
    if (crossBoardDuplicates.length > 10) {
      message += `, ... +${crossBoardDuplicates.length - 10} more`;
    }
    message += '\n\n';
  }
  
  if (buyerDuplicatesWithinBoard.length > 0) {
    message += `🔵 ${buyerDuplicatesWithinBoard.length} duplicate(s) within Buyers board:\n`;
    message += buyerDuplicatesWithinBoard.slice(0, 10).join(', ');
    if (buyerDuplicatesWithinBoard.length > 10) {
      message += `, ... +${buyerDuplicatesWithinBoard.length - 10} more`;
    }
    message += '\n\n';
  }
  
  if (affiliateDuplicatesWithinBoard.length > 0) {
    message += `🔵 ${affiliateDuplicatesWithinBoard.length} duplicate(s) within Affiliates board:\n`;
    message += affiliateDuplicatesWithinBoard.slice(0, 10).join(', ');
    if (affiliateDuplicatesWithinBoard.length > 10) {
      message += `, ... +${affiliateDuplicatesWithinBoard.length - 10} more`;
    }
    message += '\n\n';
  }
  
  if (excludedDuplicates.length > 0) {
    message += `(${excludedDuplicates.length} known cross-board duplicates already excluded)\n\n`;
  }
  message += 'See "Duplicate Vendors" sheet for full list.';
  
  ui.alert('Duplicate Vendors Found', message, ui.ButtonSet.OK);
  
  // Write to sheet for easier review
  writeDuplicatesToSheet_(crossBoardDuplicates, excludedDuplicates, buyerDuplicatesWithinBoard, affiliateDuplicatesWithinBoard);
}

/**
 * Find duplicate names within a single list
 * Returns array of names that appear more than once
 */
function findDuplicatesInList_(data) {
  const nameCounts = {};
  const duplicates = [];
  
  for (const row of data) {
    const name = String(row[0]).trim();
    const nameLower = name.toLowerCase();
    if (!nameLower) continue;
    
    if (nameCounts[nameLower]) {
      nameCounts[nameLower].count++;
      // Only add to duplicates once (when we see it the second time)
      if (nameCounts[nameLower].count === 2) {
        duplicates.push(nameCounts[nameLower].properCase);
      }
    } else {
      nameCounts[nameLower] = { count: 1, properCase: name };
    }
  }
  
  return duplicates;
}

/**
 * Get duplicate exclusions from Settings sheet
 * Format in Settings:
 * Row: "Duplicate Exclusions" | (empty)
 * Row: "Vendor Name" | "Reason"
 * Row: "SomeVendor" | "Legitimately both buyer and affiliate"
 */
function getDuplicateExclusions_() {
  const ss = SpreadsheetApp.getActive();
  const settingsSh = ss.getSheetByName('Settings');
  
  if (!settingsSh) {
    return [];
  }
  
  const data = settingsSh.getDataRange().getValues();
  const exclusions = [];
  let inExclusionsSection = false;
  let startCol = -1;
  
  for (let i = 0; i < data.length; i++) {
    // Search for "Duplicate Exclusions" header in any column
    if (!inExclusionsSection) {
      for (let col = 0; col < data[i].length; col++) {
        if (String(data[i][col] || '').trim().toLowerCase() === 'duplicate exclusions') {
          inExclusionsSection = true;
          startCol = col;
          break;
        }
      }
      continue;
    }
    
    const vendorCell = String(data[i][startCol] || '').trim();
    
    // Skip header row
    if (vendorCell.toLowerCase() === 'vendor name') {
      continue;
    }
    
    // Exit if we hit an empty row
    if (vendorCell === '') {
      break;
    }
    
    exclusions.push(vendorCell);
  }
  
  Logger.log(`Loaded ${exclusions.length} duplicate exclusions from Settings`);
  return exclusions;
}

/**
 * Write duplicates to a sheet for easy review
 */
function writeDuplicatesToSheet_(crossBoardDuplicates, excludedDuplicates, buyerDuplicatesWithinBoard, affiliateDuplicatesWithinBoard) {
  const ss = SpreadsheetApp.getActive();
  const sheetName = 'Duplicate Vendors';
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
  }
  
  // Headers
  sheet.getRange(1, 1, 1, 4).setValues([['Vendor Name', 'Status', 'Type', 'Notes']]);
  sheet.getRange(1, 1, 1, 4)
    .setFontWeight('bold')
    .setBackground('#4a86e8')
    .setFontColor('white');
  
  let row = 2;
  
  // Write cross-board duplicates (yellow)
  for (const vendor of crossBoardDuplicates) {
    sheet.getRange(row, 1).setValue(vendor);
    sheet.getRange(row, 2).setValue('⚠️ DUPLICATE');
    sheet.getRange(row, 3).setValue('Cross-Board');
    sheet.getRange(row, 4).setValue('In both Buyers AND Affiliates');
    sheet.getRange(row, 1, 1, 4).setBackground('#fff2cc'); // Yellow
    row++;
  }
  
  // Write same-board duplicates - Buyers (blue)
  for (const vendor of buyerDuplicatesWithinBoard) {
    sheet.getRange(row, 1).setValue(vendor);
    sheet.getRange(row, 2).setValue('🔵 DUPLICATE');
    sheet.getRange(row, 3).setValue('Within Buyers');
    sheet.getRange(row, 4).setValue('Appears multiple times in Buyers board');
    sheet.getRange(row, 1, 1, 4).setBackground('#cfe2f3'); // Light blue
    row++;
  }
  
  // Write same-board duplicates - Affiliates (blue)
  for (const vendor of affiliateDuplicatesWithinBoard) {
    sheet.getRange(row, 1).setValue(vendor);
    sheet.getRange(row, 2).setValue('🔵 DUPLICATE');
    sheet.getRange(row, 3).setValue('Within Affiliates');
    sheet.getRange(row, 4).setValue('Appears multiple times in Affiliates board');
    sheet.getRange(row, 1, 1, 4).setBackground('#cfe2f3'); // Light blue
    row++;
  }
  
  // Write excluded duplicates (green)
  for (const vendor of excludedDuplicates) {
    sheet.getRange(row, 1).setValue(vendor);
    sheet.getRange(row, 2).setValue('✓ Excluded');
    sheet.getRange(row, 3).setValue('Cross-Board');
    sheet.getRange(row, 4).setValue('In Settings exclusion list');
    sheet.getRange(row, 1, 1, 4).setBackground('#d9ead3'); // Green
    row++;
  }
  
  // Auto-resize
  sheet.autoResizeColumns(1, 4);
  sheet.setFrozenRows(1);
  
  // Add legend and instructions
  const totalRows = crossBoardDuplicates.length + buyerDuplicatesWithinBoard.length + 
                    affiliateDuplicatesWithinBoard.length + excludedDuplicates.length;
  if (totalRows > 0) {
    row += 2;
    sheet.getRange(row, 1).setValue('LEGEND:').setFontWeight('bold');
    row++;
    sheet.getRange(row, 1, 1, 2).setValues([['🟡 Yellow', 'Cross-board duplicate (in both Buyers and Affiliates)']]);
    sheet.getRange(row, 1, 1, 2).setBackground('#fff2cc');
    row++;
    sheet.getRange(row, 1, 1, 2).setValues([['🔵 Blue', 'Same-board duplicate (appears multiple times on one board)']]);
    sheet.getRange(row, 1, 1, 2).setBackground('#cfe2f3');
    row++;
    sheet.getRange(row, 1, 1, 2).setValues([['🟢 Green', 'Excluded (in Settings exclusion list)']]);
    sheet.getRange(row, 1, 1, 2).setBackground('#d9ead3');
    row += 2;
    
    sheet.getRange(row, 1).setValue('To exclude cross-board duplicates, add to Settings sheet:').setFontWeight('bold');
    row++;
    sheet.getRange(row, 1).setValue('Duplicate Exclusions');
    sheet.getRange(row, 1).setFontWeight('bold');
    row++;
    sheet.getRange(row, 1).setValue('Vendor Name');
    sheet.getRange(row, 2).setValue('Reason');
    sheet.getRange(row, 1, 1, 2).setFontWeight('bold').setBackground('#f3f3f3');
    row++;
    sheet.getRange(row, 1).setValue('Example Vendor');
    sheet.getRange(row, 2).setValue('Legitimately both buyer and affiliate');
    sheet.getRange(row, 1, 1, 2).setFontStyle('italic');
  }
  
  ss.toast(`Results written to "${sheetName}" sheet`, '✅ Done', 3);
}

/************************************************************
 * EMAIL RESPONSE GENERATION
 * Generate email responses using Claude AI and create Gmail drafts
 ************************************************************/

/**
 * Get the selected email thread from the A(I)DEN sheet
 * User should have their cursor on an email row in the EMAILS section
 */
function getSelectedEmailThread_() {
  const ss = SpreadsheetApp.getActive();
  const bsSh = ss.getSheetByName(BS_CFG.BATTLE_SHEET);

  if (!bsSh) {
    throw new Error('A(I)DEN sheet not found');
  }

  // Check if we're in the emails section by looking for the subject column (column A)
  // and checking if the row is after the EMAILS header
  const values = bsSh.getDataRange().getValues();

  // Find the EMAILS header row
  let emailsHeaderRow = -1;
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).includes('📧 EMAILS')) {
      emailsHeaderRow = i + 1; // 1-based
      break;
    }
  }

  if (emailsHeaderRow === -1) {
    throw new Error('Could not find EMAILS section');
  }

  // The header row with Subject/Date/Last/Labels is emailsHeaderRow + 1
  // Email rows start at emailsHeaderRow + 2
  const emailDataStartRow = emailsHeaderRow + 2;

  // Determine which row to use - selected row if valid, otherwise first email row
  const selection = bsSh.getSelection();
  const activeCell = selection.getCurrentCell();
  let row = emailDataStartRow; // Default to first email (newest)

  if (activeCell) {
    const selectedRow = activeCell.getRow();
    if (selectedRow >= emailDataStartRow) {
      row = selectedRow; // Use selected row if it's in the email section
    }
  }

  // Get the subject from the row
  const subject = String(bsSh.getRange(row, 1).getValue() || '').trim();

  if (!subject || subject === 'No emails found') {
    throw new Error('No emails available');
  }

  // Try to get the thread ID from the hyperlink formula
  const formula = bsSh.getRange(row, 1).getFormula();
  let threadId = null;

  if (formula) {
    // Extract thread ID from HYPERLINK formula like: =HYPERLINK("https://mail.google.com/mail/u/0/#inbox/THREADID", "Subject")
    const match = formula.match(/#inbox\/([^"]+)/);
    if (match) {
      threadId = match[1];
    }
  }

  if (!threadId) {
    throw new Error('Could not find thread ID for selected email. Make sure you clicked on a valid email row.');
  }

  // Fetch the full thread
  const thread = GmailApp.getThreadById(threadId);
  if (!thread) {
    throw new Error('Could not fetch email thread');
  }

  return {
    threadId: threadId,
    subject: subject,
    thread: thread
  };
}

/**
 * Get Email Response settings from Settings sheet
 * Format in Settings:
 * Row: "Email Response Settings" | (empty)
 * Row: "Response Type" | "Custom Instructions"
 * Row: "Cold Outreach - Follow Up" | "Be persistent but polite"
 * Row: "Schedule a Call" | "Always offer specific times"
 * Row: "General" | "Instructions that apply to all response types"
 * Note: Signature is fetched dynamically from Gmail API, not from settings
 */
function getEmailResponseSettings_() {
  const ss = SpreadsheetApp.getActive();
  const settingsSh = ss.getSheetByName('Settings');

  if (!settingsSh) {
    return { typeInstructions: {}, generalInstructions: '' };
  }

  const data = settingsSh.getDataRange().getValues();
  const typeInstructions = {};
  let generalInstructions = '';
  let inSection = false;
  let startCol = -1;

  for (let i = 0; i < data.length; i++) {
    // Search for "Email Response Settings" header in any column
    if (!inSection) {
      for (let col = 0; col < data[i].length; col++) {
        if (String(data[i][col] || '').trim().toLowerCase() === 'email response settings') {
          inSection = true;
          startCol = col;
          break;
        }
      }
      continue;
    }

    const typeCell = String(data[i][startCol] || '').trim();
    const instructionsCell = String(data[i][startCol + 1] || '').trim();

    // Skip header row
    if (typeCell.toLowerCase() === 'response type') {
      continue;
    }

    // Exit if we hit an empty row
    if (typeCell === '' && instructionsCell === '') {
      break;
    }

    // Parse settings - skip "signature" row (signature comes from Gmail API now)
    if (typeCell !== '' && instructionsCell !== '') {
      const typeLower = typeCell.toLowerCase();
      if (typeLower === 'general') {
        generalInstructions = instructionsCell;
      } else if (typeLower !== 'signature') {
        typeInstructions[typeCell] = instructionsCell;
      }
    }
  }

  return { typeInstructions, generalInstructions };
}

/**
 * Get task analysis settings from Settings sheet
 * Format: "Task Analysis Settings" header, then "Task Name" | "What to Look For"
 */
function getTaskAnalysisSettings_() {
  const ss = SpreadsheetApp.getActive();
  const settingsSh = ss.getSheetByName('Settings');

  if (!settingsSh) {
    return { taskInstructions: {}, generalInstructions: '' };
  }

  const data = settingsSh.getDataRange().getValues();
  const taskInstructions = {};
  let generalInstructions = '';
  let inSection = false;
  let startCol = -1;

  for (let i = 0; i < data.length; i++) {
    // Search for "Task Analysis Settings" header in any column
    if (!inSection) {
      for (let col = 0; col < data[i].length; col++) {
        if (String(data[i][col] || '').trim().toLowerCase() === 'task analysis settings') {
          inSection = true;
          startCol = col;
          break;
        }
      }
      continue;
    }

    const taskCell = String(data[i][startCol] || '').trim();
    const instructionsCell = String(data[i][startCol + 1] || '').trim();

    // Skip header row
    if (taskCell.toLowerCase() === 'task name') {
      continue;
    }

    // Exit if we hit an empty row
    if (taskCell === '' && instructionsCell === '') {
      break;
    }

    // Parse settings
    if (taskCell !== '' && instructionsCell !== '') {
      if (taskCell.toLowerCase() === 'general') {
        generalInstructions = instructionsCell;
      } else {
        taskInstructions[taskCell.toLowerCase()] = instructionsCell;
      }
    }
  }

  return { taskInstructions, generalInstructions };
}

/**
 * Get the full thread content formatted for Claude
 */
function getThreadContent_(thread) {
  const messages = thread.getMessages();
  const myEmail = Session.getActiveUser().getEmail().toLowerCase();
  let content = '';

  for (const msg of messages) {
    const from = msg.getFrom();
    const to = msg.getTo();
    const date = msg.getDate();
    const body = msg.getPlainBody();

    content += `\n\n=== MESSAGE ===\n`;
    content += `From: ${from}\n`;
    content += `To: ${to}\n`;
    content += `Date: ${date}\n`;
    content += `---\n`;
    content += body.substring(0, 3000); // Limit each message
  }

  // Check if I sent the last message (meaning this is a follow-up, not a reply)
  const lastMessage = messages[messages.length - 1];
  const lastFrom = lastMessage.getFrom().toLowerCase();
  const lastSenderIsMe = lastFrom.includes(myEmail) || lastFrom.includes('andy');

  return { content, lastSenderIsMe };
}

/**
 * Show dialog for extra directions and get user input
 */
function showDirectionsDialog_(responseType) {
  const ui = SpreadsheetApp.getUi();

  const result = ui.prompt(
    `📧 ${responseType}`,
    'Add any extra directions or context for the response:\n(Leave blank for default response)',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() !== ui.Button.OK) {
    return null;
  }

  return result.getResponseText().trim();
}

/**
 * Generate email response using Claude
 */
function generateEmailWithClaude_(threadContent, subject, responseType, extraDirections, lastSenderIsMe) {
  const claudeApiKey = BS_CFG.CLAUDE_API_KEY;

  if (!claudeApiKey || claudeApiKey === 'YOUR_ANTHROPIC_API_KEY_HERE') {
    throw new Error('Please set your Anthropic API key in BS_CFG.CLAUDE_API_KEY');
  }

  // Load custom settings from Settings sheet
  const settings = getEmailResponseSettings_();

  // Build the prompt based on response type
  let systemPrompt = `You are drafting an email AS Andy Worford, Director of Business Development at Profitise. Write in first person as Andy.

Key info:
- Company: Profitise - delivers exclusive homeowner leads in real-time with TCPA consent
- Andy's scheduling link: https://calendar.app.google/68Fk8pb9mUokSiaW8
- Phone: 888-400-4868 x 8117 | Cell: 949.351.2300

CRITICAL STYLE RULES - You must follow these:
- Write like a real human, not like an AI assistant
- NEVER use phrases like "Perfect!", "Absolutely!", "Great question!", "I'd be happy to", "Hope this helps!"
- NEVER start with enthusiastic affirmations - just get to the point
- Be direct, professional, casual-friendly (not corporate-stiff)
- Keep it brief - 2-4 sentences when possible
- Don't over-explain or be overly formal
- No signature block (added automatically)
- No extra blank line between greeting and body (e.g., "Hi Catie,\\nJust following up..." not "Hi Catie,\\n\\nJust following up...")

CONTEXT AWARENESS:
- Pay attention to WHO sent the last message in the thread
- If Andy (me) sent the last message, this is a FOLLOW-UP because they haven't replied yet
- If someone else sent the last message, this is a REPLY to their message`;

  // Add general custom instructions from Settings if present
  if (settings.generalInstructions) {
    systemPrompt += `\n\nADDITIONAL INSTRUCTIONS FROM ANDY:\n${settings.generalInstructions}`;
  }

  // Add type-specific instructions from Settings if present
  if (settings.typeInstructions[responseType]) {
    systemPrompt += `\n\nSPECIFIC INSTRUCTIONS FOR "${responseType}":\n${settings.typeInstructions[responseType]}`;
  }

  // Add context about whether this is a follow-up or reply
  let contextNote = '';
  if (lastSenderIsMe) {
    contextNote = `\n\nIMPORTANT: I (Andy) sent the last message and haven't received a reply. This is a FOLLOW-UP, not a reply to them. Don't act like they said something they didn't.`;
  }

  let userPrompt = `Response Type: ${responseType}${contextNote}

Email Thread:
${threadContent}

`;

  if (extraDirections) {
    userPrompt += `Extra directions: ${extraDirections}\n\n`;
  }

  userPrompt += `Draft the email body only (no subject line).`;

  const payload = {
    model: 'claude-sonnet-4-20250514',
    max_tokens: 1024,
    messages: [
      { role: 'user', content: userPrompt }
    ],
    system: systemPrompt
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': claudeApiKey,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', options);
  const result = JSON.parse(response.getContentText());

  if (result.error) {
    throw new Error(`Claude API error: ${result.error.message}`);
  }

  return result.content[0].text;
}

/**
 * Create Gmail draft and return the URL
 * Uses createDraftReplyAll for proper threading, then updates content via Gmail API
 */
function createDraftAndGetUrl_(thread, responseBody) {
  const messages = thread.getMessages();
  const lastMessage = messages[messages.length - 1];

  // Get my email address to exclude from recipients
  const myEmail = Session.getActiveUser().getEmail().toLowerCase();

  // Get signature from Gmail settings (includes images, fonts, formatting)
  let signature = '';
  try {
    const sendAsSettings = Gmail.Users.Settings.SendAs.list('me');
    if (sendAsSettings && sendAsSettings.sendAs) {
      // Find the primary send-as (or the one matching user's email)
      const primarySendAs = sendAsSettings.sendAs.find(s => s.isPrimary) ||
                            sendAsSettings.sendAs.find(s => s.sendAsEmail.toLowerCase() === myEmail) ||
                            sendAsSettings.sendAs[0];
      if (primarySendAs && primarySendAs.signature) {
        signature = primarySendAs.signature;
      }
    }
  } catch (e) {
    Logger.log('Could not fetch Gmail signature: ' + e.message);
  }

  // Internal domains - people in these domains go to CC
  const internalDomains = ['profitise.com', 'zeroparallel.com', 'phonexa.com'];

  // Helper to check if email is internal
  const isInternalEmail = (email) => {
    const emailLower = email.toLowerCase();
    return internalDomains.some(domain => emailLower.includes('@' + domain));
  };

  // Helper to extract email from "Name <email>" format
  const extractEmail = (str) => {
    const match = str.match(/<([^>]+)>/);
    return match ? match[1].toLowerCase() : str.toLowerCase().trim();
  };

  // Helper to split email addresses respecting quoted names (handles "Name, Inc" <email>)
  const splitAddresses = (str) => {
    if (!str) return [];
    const addresses = [];
    let current = '';
    let inQuotes = false;

    for (let i = 0; i < str.length; i++) {
      const char = str[i];
      if (char === '"') {
        inQuotes = !inQuotes;
        current += char;
      } else if (char === ',' && !inQuotes) {
        if (current.trim()) addresses.push(current.trim());
        current = '';
      } else {
        current += char;
      }
    }
    if (current.trim()) addresses.push(current.trim());
    return addresses;
  };

  // Collect all unique participants from the thread (excluding me)
  const allParticipants = new Set();

  for (const msg of messages) {
    const from = msg.getFrom() || '';
    const to = msg.getTo() || '';
    const cc = msg.getCc() || '';

    // Parse all addresses (respecting quoted names with commas)
    const allAddresses = [
      ...splitAddresses(from),
      ...splitAddresses(to),
      ...splitAddresses(cc)
    ];

    for (const addr of allAddresses) {
      const email = extractEmail(addr);
      if (email && email !== myEmail && !email.includes('sales@profitise.com')) {
        allParticipants.add(addr.trim());
      }
    }
  }

  // Separate into external (To) and internal (CC)
  const externalRecipients = [];
  const internalRecipients = [];

  for (const addr of allParticipants) {
    const email = extractEmail(addr);
    if (isInternalEmail(email)) {
      internalRecipients.push(addr);
    } else {
      externalRecipients.push(addr);
    }
  }

  // Build recipient strings
  // If no external recipients, internal go to To instead of Cc
  let toRecipients = '';
  let ccRecipients = '';
  let bccRecipients = '';

  if (externalRecipients.length > 0) {
    // Normal case: external to To, internal to Cc
    toRecipients = externalRecipients.join(', ');
    ccRecipients = internalRecipients.join(', ');
  } else {
    // All internal: put them in To
    toRecipients = internalRecipients.join(', ');
  }

  // Check if sales@profitise.com should be added
  const allRecipientsStr = [...allParticipants].join(',').toLowerCase();
  const salesAlreadyIncluded = allRecipientsStr.includes('sales@profitise.com');

  if (!salesAlreadyIncluded) {
    if (ccRecipients) {
      bccRecipients = 'sales@profitise.com';
    } else if (toRecipients) {
      bccRecipients = 'sales@profitise.com';
    } else {
      toRecipients = 'sales@profitise.com';
    }
  }

  // Log computed recipients
  Logger.log('=== RECIPIENT CALCULATION ===');
  Logger.log(`My email: ${myEmail}`);
  Logger.log(`All participants (${allParticipants.size}): ${[...allParticipants].join('; ')}`);
  Logger.log(`External (To): ${toRecipients || '(none)'}`);
  Logger.log(`Internal (Cc): ${ccRecipients || '(none)'}`);
  Logger.log(`Bcc: ${bccRecipients || '(none)'}`);
  Logger.log('=============================');

  // Build the quoted original message
  const lastMsgDate = lastMessage.getDate();
  const lastMsgFrom = lastMessage.getFrom();
  // Use getBody() to preserve HTML formatting (images, fonts, styling)
  const lastMsgBody = lastMessage.getBody() || '';

  // Format date for quote header
  const dateStr = Utilities.formatDate(lastMsgDate, Session.getScriptTimeZone(), "EEE, MMM d, yyyy 'at' h:mm a");

  // Helper to escape HTML special characters (for our text, not the original email)
  const escapeHtml = (text) => {
    return text
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  };

  // Convert response body to HTML (preserve line breaks)
  const responseHtml = escapeHtml(responseBody).replace(/\n/g, '<br>');

  // Build quoted message using Gmail's standard format
  const quoteHeaderHtml = `<div class="gmail_quote"><div dir="ltr" class="gmail_attr">On ${escapeHtml(dateStr)}, ${escapeHtml(lastMsgFrom)} wrote:<br></div><blockquote class="gmail_quote" style="margin:0px 0px 0px 0.8ex;border-left:1px solid rgb(204,204,204);padding-left:1ex">`;
  const quotedBodyHtml = `${lastMsgBody}</blockquote></div>`;

  // Full email body with signature and quote (HTML format)
  let fullBodyHtml = `<div dir="ltr">${responseHtml}`;
  if (signature) {
    fullBodyHtml += `<br><br>${signature}`;
  }
  fullBodyHtml += `</div><br>${quoteHeaderHtml}${quotedBodyHtml}`;

  // Step 1: Create a draft reply using GmailApp (this ensures proper threading)
  Logger.log('Creating draft reply with proper threading...');
  const draftReply = thread.createDraftReplyAll('Placeholder body - will be replaced');
  const draft = draftReply.getMessage();
  const draftId = draftReply.getId();

  // Step 2: Get threading headers from the draft (these were set by createDraftReplyAll)
  const inReplyTo = draft.getHeader('In-Reply-To') || '';
  const references = draft.getHeader('References') || '';
  Logger.log(`Thread ID: ${thread.getId()}`);
  Logger.log(`Thread first subject: ${thread.getFirstMessageSubject()}`);
  Logger.log(`Draft subject: ${draft.getSubject()}`);
  Logger.log(`Threading - In-Reply-To: ${inReplyTo || '(empty)'}`);
  Logger.log(`Threading - References: ${references || '(empty)'}`);

  // Step 3: Get the draft via Gmail API
  const gmailDraft = Gmail.Users.Drafts.get('me', draftId);
  const messageId = gmailDraft.message.id;

  // Step 4: Build the updated message preserving threading headers
  // Use thread's first message subject to avoid [External] tags added by recipients
  let subject = thread.getFirstMessageSubject();
  if (!subject.toLowerCase().startsWith('re:')) {
    subject = 'Re: ' + subject;
  }

  let rawHeaders = `From: ${myEmail}\r\n`;
  rawHeaders += `To: ${toRecipients}\r\n`;
  if (ccRecipients) rawHeaders += `Cc: ${ccRecipients}\r\n`;
  if (bccRecipients) rawHeaders += `Bcc: ${bccRecipients}\r\n`;
  rawHeaders += `Subject: ${subject}\r\n`;
  // CRITICAL: Include threading headers for proper reply threading
  if (inReplyTo) rawHeaders += `In-Reply-To: ${inReplyTo}\r\n`;
  if (references) rawHeaders += `References: ${references}\r\n`;
  rawHeaders += `MIME-Version: 1.0\r\n`;
  rawHeaders += `Content-Type: text/html; charset="UTF-8"\r\n`;
  rawHeaders += `\r\n`;

  const updateResource = {
    message: {
      // Use proper UTF-8 encoding for special characters in signature
      raw: Utilities.base64EncodeWebSafe(Utilities.newBlob(rawHeaders + fullBodyHtml).getBytes()),
      threadId: thread.getId()  // Also explicitly set threadId
    }
  };

  // Step 4: Update the draft with correct content and recipients
  Logger.log('Updating draft with HTML content and recipients...');
  const updatedDraft = Gmail.Users.Drafts.update(updateResource, 'me', draftId);

  // Get the message ID from the updated draft for the URL
  const updatedMessageId = updatedDraft.message.id;
  Logger.log(`Draft updated, message ID: ${updatedMessageId}`);

  // Track this thread for auto-archive after sending
  addPendingArchiveThread_(thread.getId());
  Logger.log(`Added thread ${thread.getId()} to pending archive list`);

  // Use the draft message ID to open compose window
  const gmailUrl = `https://mail.google.com/mail/u/0/#drafts?compose=${updatedMessageId}`;
  return gmailUrl;
}

/**
 * Add a thread ID to the pending archive list (stored in Script Properties)
 * These threads will be archived on next email refresh once Last=ME
 */
function addPendingArchiveThread_(threadId) {
  const props = PropertiesService.getScriptProperties();
  const pendingJson = props.getProperty('pendingArchiveThreads') || '[]';
  const pending = JSON.parse(pendingJson);

  if (!pending.includes(threadId)) {
    pending.push(threadId);
    props.setProperty('pendingArchiveThreads', JSON.stringify(pending));
  }
}

/**
 * Get pending archive thread IDs
 */
function getPendingArchiveThreads_() {
  const props = PropertiesService.getScriptProperties();
  const pendingJson = props.getProperty('pendingArchiveThreads') || '[]';
  return JSON.parse(pendingJson);
}

/**
 * Remove a thread ID from pending archive list (after archiving)
 */
function removePendingArchiveThread_(threadId) {
  const props = PropertiesService.getScriptProperties();
  const pendingJson = props.getProperty('pendingArchiveThreads') || '[]';
  const pending = JSON.parse(pendingJson);
  const updated = pending.filter(id => id !== threadId);
  props.setProperty('pendingArchiveThreads', JSON.stringify(updated));
}

/**
 * Check and archive threads that were responded to via Email Response
 * Called during email refresh - archives threads where Last=ME
 * Returns true if any threads were archived (to trigger refresh)
 */
function checkAndArchivePendingThreads_() {
  const pendingIds = getPendingArchiveThreads_();
  if (pendingIds.length === 0) return false;

  Logger.log(`Checking ${pendingIds.length} pending archive threads...`);
  const myEmail = Session.getActiveUser().getEmail().toLowerCase();
  let archivedCount = 0;

  for (const threadId of pendingIds) {
    try {
      const thread = GmailApp.getThreadById(threadId);
      if (!thread) {
        // Thread no longer exists, remove from pending
        removePendingArchiveThread_(threadId);
        continue;
      }

      const messages = thread.getMessages();
      if (messages.length === 0) continue;

      const lastMessage = messages[messages.length - 1];
      const lastSender = lastMessage.getFrom().toLowerCase();

      // Check if I sent the last message
      if (lastSender.includes(myEmail)) {
        Logger.log(`Archiving thread ${threadId} - Last sender is ME`);
        thread.moveToArchive();
        removePendingArchiveThread_(threadId);
        archivedCount++;
      }
    } catch (e) {
      Logger.log(`Error checking thread ${threadId}: ${e.message}`);
      // Remove from pending to avoid repeated errors
      removePendingArchiveThread_(threadId);
    }
  }

  if (archivedCount > 0) {
    Logger.log(`Archived ${archivedCount} threads`);
    SpreadsheetApp.getActive().toast(`Archived ${archivedCount} sent email(s)`, '📬 Auto-Archive', 3);
  }

  return archivedCount > 0;
}

/**
 * Check if there's a future meeting today for the current vendor
 * Returns warning message if found, null otherwise
 */
function checkForFutureMeetingToday_() {
  try {
    const ss = SpreadsheetApp.getActive();
    const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
    if (!listSh) return null;

    const currentIndex = getCurrentVendorIndex_();
    if (!currentIndex) return null;

    const listRow = currentIndex + 1; // +1 for header row
    const vendor = String(listSh.getRange(listRow, BS_CFG.L_VENDOR + 1).getValue() || '').trim();
    if (!vendor) return null;

    // Get contact emails for calendar search
    const contactEmails = [];
    // Try to get from cached/stored data first, or just search by vendor name

    // Get meetings for this vendor
    const meetingsResult = getUpcomingMeetingsForVendor_(vendor, contactEmails);
    const meetings = meetingsResult.meetings || [];

    // Find meetings that are today but in the future
    const now = new Date();
    const futureMeetingsToday = meetings.filter(m => {
      if (!m.isToday || m.isPast) return false;
      // Double-check: meeting start time is in the future
      if (m.startTime && m.startTime > now) return true;
      return false;
    });

    if (futureMeetingsToday.length > 0) {
      const meetingList = futureMeetingsToday.map(m => `• ${m.title} at ${m.time}`).join('\n');
      return `You have a meeting with ${vendor} later today:\n\n${meetingList}`;
    }

    return null;
  } catch (e) {
    Logger.log(`Error checking for future meetings: ${e.message}`);
    return null; // Don't block on errors
  }
}

/**
 * Main function to handle email response generation
 */
function generateEmailResponse_(responseType) {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  try {
    // Check for future meeting today before proceeding
    const futureMeetingWarning = checkForFutureMeetingToday_();
    if (futureMeetingWarning) {
      const response = ui.alert(
        '⚠️ Meeting Today',
        `${futureMeetingWarning}\n\nDo you still want to send an email response?`,
        ui.ButtonSet.YES_NO
      );
      if (response !== ui.Button.YES) {
        return; // User cancelled
      }
    }

    ss.toast('Getting selected email...', '📧 Email Response', 2);

    // Get selected email thread
    const emailData = getSelectedEmailThread_();

    // Show directions dialog
    const extraDirections = showDirectionsDialog_(responseType);
    if (extraDirections === null) {
      return; // User cancelled
    }

    ss.toast('Reading email thread...', '📧 Email Response', 2);

    // Get full thread content and context
    const { content: threadContent, lastSenderIsMe } = getThreadContent_(emailData.thread);

    ss.toast('Generating response with Claude...', '🤖 AI Working', 5);

    // Generate response with Claude (no draft yet)
    const responseBody = generateEmailWithClaude_(
      threadContent,
      emailData.subject,
      responseType,
      extraDirections,
      lastSenderIsMe
    );

    // Store context for potential revision or draft creation
    const revisionContext = {
      threadId: emailData.threadId,
      responseType: responseType,
      originalDirections: extraDirections,
      previousResponse: responseBody
    };
    PropertiesService.getUserProperties().setProperty('emailRevisionContext', JSON.stringify(revisionContext));

    // Show preview - draft only created when user confirms
    showDraftPreviewDialog_(responseBody, emailData.threadId);

  } catch (e) {
    ui.alert('Error', e.message, ui.ButtonSet.OK);
  }
}

/**
 * Show the draft preview dialog with revision option
 * Draft is only created when user clicks "Create Draft"
 */
function showDraftPreviewDialog_(responseBody, threadId) {
  const ui = SpreadsheetApp.getUi();

  // Build Gmail thread URL
  const threadUrl = threadId ? `https://mail.google.com/mail/u/0/#inbox/${threadId}` : '';

  // Safely escape the response for embedding in HTML
  const escapedResponse = responseBody
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/\n/g, '<br>');

  const html = `<!DOCTYPE html>
<html>
<head>
  <base target="_blank">
  <style>
    body { font-family: Arial, sans-serif; padding: 20px; }
    .header { color: #202124; font-size: 16px; margin-bottom: 15px; }
    .buttons { display: flex; gap: 10px; margin-bottom: 15px; }
    .btn {
      display: inline-block;
      padding: 12px 24px;
      border-radius: 4px;
      font-size: 14px;
      cursor: pointer;
      border: none;
    }
    .btn-primary { background: #1a73e8; color: white; }
    .btn-primary:hover { background: #1557b0; }
    .btn-primary:disabled { background: #ccc; cursor: not-allowed; }
    .btn-secondary { background: #f1f3f4; color: #5f6368; }
    .btn-secondary:hover { background: #e8eaed; }
    .preview {
      background: #f8f9fa;
      padding: 15px;
      border-radius: 8px;
      margin-top: 15px;
      font-size: 13px;
      max-height: 180px;
      overflow-y: auto;
      line-height: 1.5;
    }
    .revision-section {
      display: none;
      margin-top: 15px;
      padding-top: 15px;
      border-top: 1px solid #e0e0e0;
    }
    .revision-section.show { display: block; }
    .revision-input {
      width: 100%;
      padding: 10px;
      border: 1px solid #ddd;
      border-radius: 4px;
      font-size: 13px;
      margin-bottom: 10px;
      box-sizing: border-box;
    }
    .revision-label { font-size: 13px; color: #5f6368; margin-bottom: 5px; }
    .loading { color: #5f6368; font-style: italic; margin-top: 10px; }
  </style>
</head>
<body>
  <div class="header">Preview Generated Response ${threadUrl ? `<a href="${threadUrl}" target="_blank" style="font-size: 12px; color: #1a73e8; margin-left: 10px;">📧 View Original</a>` : ''}</div>
  <div class="buttons">
    <button id="createBtn" class="btn btn-primary" onclick="doCreateDraft()">Create Draft and Open</button>
    <button id="reviseBtn" class="btn btn-secondary" onclick="doShowRevision()">Revise</button>
  </div>
  <div id="previewContent" class="preview">${escapedResponse}</div>

  <div id="revisionSection" class="revision-section">
    <div class="revision-label">What would you like to change?</div>
    <input type="text" id="revisionInput" class="revision-input" placeholder="e.g., add actual scheduling link, make it shorter...">
    <button id="regenerateBtn" class="btn btn-primary" onclick="doRegenerate()">Regenerate</button>
  </div>

  <div id="loadingMsg" class="loading" style="display:none;"></div>

  <script>
    function doCreateDraft() {
      document.getElementById('createBtn').disabled = true;
      document.getElementById('reviseBtn').disabled = true;
      document.getElementById('loadingMsg').textContent = 'Creating draft...';
      document.getElementById('loadingMsg').style.display = 'block';

      google.script.run
        .withSuccessHandler(function(url) {
          window.open(url, '_blank');
          google.script.host.close();
        })
        .withFailureHandler(function(err) {
          alert('Error: ' + (err.message || err));
          document.getElementById('createBtn').disabled = false;
          document.getElementById('reviseBtn').disabled = false;
          document.getElementById('loadingMsg').style.display = 'none';
        })
        .createDraftFromPreview();
    }

    function doShowRevision() {
      document.getElementById('revisionSection').classList.add('show');
      document.getElementById('revisionInput').focus();
    }

    function doRegenerate() {
      var feedback = document.getElementById('revisionInput').value.trim();
      if (!feedback) {
        alert('Please enter what you want to change');
        return;
      }

      document.getElementById('regenerateBtn').disabled = true;
      document.getElementById('createBtn').disabled = true;
      document.getElementById('loadingMsg').textContent = 'Regenerating...';
      document.getElementById('loadingMsg').style.display = 'block';

      google.script.run
        .withSuccessHandler(function(newResponse) {
          var el = document.getElementById('previewContent');
          var safe = newResponse.replace(/&/g, '&amp;');
          safe = safe.replace(/[<]/g, '&lt;');
          safe = safe.replace(/[>]/g, '&gt;');
          safe = safe.replace(/\\n/g, '<br>');
          el.innerHTML = safe;
          document.getElementById('revisionInput').value = '';
          document.getElementById('regenerateBtn').disabled = false;
          document.getElementById('createBtn').disabled = false;
          document.getElementById('loadingMsg').style.display = 'none';
          document.getElementById('revisionInput').focus();
        })
        .withFailureHandler(function(err) {
          alert('Error: ' + (err.message || err));
          document.getElementById('regenerateBtn').disabled = false;
          document.getElementById('createBtn').disabled = false;
          document.getElementById('loadingMsg').style.display = 'none';
        })
        .reviseEmailDraft(feedback);
    }

    document.getElementById('revisionInput').addEventListener('keypress', function(e) {
      if (e.key === 'Enter') doRegenerate();
    });
  </script>
</body>
</html>`;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(700)
    .setHeight(600);

  ui.showModalDialog(htmlOutput, 'Email Response');
}

/**
 * Create draft from the previewed response (no underscore - must be callable from client)
 */
function createDraftFromPreview() {
  const ss = SpreadsheetApp.getActive();

  // Get stored context
  const contextJson = PropertiesService.getUserProperties().getProperty('emailRevisionContext');
  if (!contextJson) {
    throw new Error('No response context found. Please generate a new response.');
  }

  const context = JSON.parse(contextJson);

  ss.toast('Creating Gmail draft...', '📧 Creating Draft', 2);

  // Get the thread
  const thread = GmailApp.getThreadById(context.threadId);
  if (!thread) {
    throw new Error('Could not find email thread');
  }

  // Create the draft
  const draftUrl = createDraftAndGetUrl_(thread, context.previousResponse);

  // Clear the context
  PropertiesService.getUserProperties().deleteProperty('emailRevisionContext');

  return draftUrl;
}

/**
 * Revise the email response based on user feedback (no underscore - must be callable from client)
 * Returns the new response body to update the dialog in-place
 */
function reviseEmailDraft(feedback) {
  const ss = SpreadsheetApp.getActive();

  // Get stored context
  const contextJson = PropertiesService.getUserProperties().getProperty('emailRevisionContext');
  if (!contextJson) {
    throw new Error('No response context found. Please generate a new response.');
  }

  const context = JSON.parse(contextJson);

  ss.toast('Regenerating with your feedback...', '🤖 AI Working', 5);

  // Get the thread again
  const thread = GmailApp.getThreadById(context.threadId);
  if (!thread) {
    throw new Error('Could not find email thread');
  }

  // Get thread content
  const { content: threadContent, lastSenderIsMe } = getThreadContent_(thread);

  // Build revision directions combining original + feedback about previous response
  const revisionDirections = `Previous draft was:
---
${context.previousResponse}
---

User feedback on that draft: ${feedback}

Please generate an improved version addressing this feedback.`;

  // Generate new response
  const responseBody = generateEmailWithClaude_(
    threadContent,
    thread.getFirstMessageSubject(),
    context.responseType,
    revisionDirections,
    lastSenderIsMe
  );

  // Update stored context with new response
  context.previousResponse = responseBody;
  PropertiesService.getUserProperties().setProperty('emailRevisionContext', JSON.stringify(context));

  ss.toast('Done! Review updated response.', '✅', 3);

  // Return the new response - dialog will update in-place
  return responseBody;
}

// Menu item functions for each response type
function emailResponseColdFollowUp() {
  generateEmailResponse_('Cold Outreach - Follow Up');
}

function emailResponseScheduleCall() {
  generateEmailResponse_('Schedule a Call');
}

function emailResponsePaymentFollowUp() {
  generateEmailResponse_('Payment/Invoice Follow Up');
}

function emailResponseGeneralFollowUp() {
  generateEmailResponse_('General Follow Up');
}

function emailResponseCustom() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    '✍️ Custom Response Type',
    'Enter the type of response you want to generate:',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() === ui.Button.OK) {
    const customType = result.getResponseText().trim();
    if (customType) {
      generateEmailResponse_(customType);
    }
  }
}

/**
 * Email Response: Missed Meeting
 * Gets highlighted meeting from UPCOMING MEETINGS and generates follow-up
 */
function emailResponseMissedMeeting() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const bsSh = ss.getSheetByName(BS_CFG.BATTLE_SHEET);

  if (!bsSh) {
    ui.alert('Error', 'Battle Station sheet not found.', ui.ButtonSet.OK);
    return;
  }

  try {
    // Get the selected cell/row to find the highlighted meeting
    const selection = bsSh.getActiveRange();
    const selectedRow = selection.getRow();

    // Read data from the selected row to find meeting info
    const rowData = bsSh.getRange(selectedRow, 1, 1, 4).getValues()[0];
    const meetingTitle = String(rowData[0] || '').trim();
    const meetingDate = String(rowData[1] || '').trim();
    const meetingTime = String(rowData[2] || '').trim();

    // Validate we have meeting data
    if (!meetingTitle || !meetingDate) {
      ui.alert('Error', 'Please highlight a meeting row in the UPCOMING MEETINGS section.\n\nThe row should have the meeting title, date, and time.', ui.ButtonSet.OK);
      return;
    }

    // Check if this looks like a meeting row (not a header)
    if (meetingTitle.includes('UPCOMING MEETINGS') || meetingTitle === 'Meeting' || meetingTitle === 'Type') {
      ui.alert('Error', 'Please highlight a specific meeting row, not the header.', ui.ButtonSet.OK);
      return;
    }

    // Get current vendor
    const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
    const currentIndex = getCurrentVendorIndex_();
    if (!currentIndex || !listSh) {
      ui.alert('Error', 'Could not determine current vendor.', ui.ButtonSet.OK);
      return;
    }
    const listRow = currentIndex + 1;
    const vendor = String(listSh.getRange(listRow, BS_CFG.L_VENDOR + 1).getValue() || '').trim();

    // Get selected email thread (to reply to)
    ss.toast('Getting selected email...', '📧 Missed Meeting', 2);
    const emailData = getSelectedEmailThread_();

    // Build directions for Claude with meeting context
    const meetingContext = `The meeting "${meetingTitle}" was scheduled for ${meetingDate} at ${meetingTime}.\n\nWe waited but the contact didn't show up. We understand things come up, so we want to send a friendly follow-up to reschedule.\n\nInclude Andy's scheduling link for them to book a new time: https://calendar.app.google/68Fk8pb9mUokSiaW8`;

    ss.toast('Reading email thread...', '📧 Missed Meeting', 2);

    // Get full thread content
    const { content: threadContent, lastSenderIsMe } = getThreadContent_(emailData.thread);

    ss.toast('Generating response with Claude...', '🤖 AI Working', 5);

    // Generate response with Claude
    const responseBody = generateEmailWithClaude_(
      threadContent,
      emailData.subject,
      'Missed Meeting Follow Up',
      meetingContext,
      lastSenderIsMe
    );

    // Store context for potential revision
    const revisionContext = {
      threadId: emailData.threadId,
      responseType: 'Missed Meeting Follow Up',
      originalDirections: meetingContext,
      previousResponse: responseBody
    };
    PropertiesService.getUserProperties().setProperty('emailRevisionContext', JSON.stringify(revisionContext));

    // Show preview
    showDraftPreviewDialog_(responseBody, emailData.threadId);

  } catch (e) {
    ui.alert('Error', e.message, ui.ButtonSet.OK);
  }
}

/************************************************************
 * CANNED RESPONSES
 * Pre-written email templates without Claude processing
 ************************************************************/

/**
 * Canned Response: Referral Program
 */
function cannedResponseReferralProgram() {
  generateCannedResponse_('REFERRAL_PROGRAM', 'Referral Program');
}

/**
 * Canned Response: Initial Call Follow-up
 */
function cannedResponseInitialCallFollowup() {
  generateCannedResponse_('INITIAL_CALL_FOLLOWUP', 'Initial Call Follow-up');
}

/**
 * Get contacts for the current vendor from the Battle Station sheet
 * Returns array of contact names
 */
function getContactsForCurrentVendor_() {
  const ss = SpreadsheetApp.getActive();
  const bsSh = ss.getSheetByName(BS_CFG.BATTLE_SHEET);
  if (!bsSh) return [];

  const contacts = [];
  const data = bsSh.getDataRange().getValues();

  // Find CONTACTS section and extract names
  let inContactsSection = false;
  for (let i = 0; i < data.length; i++) {
    const cellA = String(data[i][0] || '').trim();

    // Check if we're entering CONTACTS section
    if (cellA.includes('CONTACTS')) {
      inContactsSection = true;
      continue;
    }

    // Check if we're leaving CONTACTS section (hit another section header)
    if (inContactsSection && (cellA.includes('UPCOMING MEETINGS') || cellA.includes('BOX DOCUMENTS') ||
        cellA.includes('GOOGLE DRIVE') || cellA.includes('CRYSTAL BALL') || cellA === '')) {
      // Empty row might just be spacing, but major headers signal end
      if (cellA.includes('UPCOMING') || cellA.includes('BOX') || cellA.includes('GOOGLE') || cellA.includes('CRYSTAL')) {
        break;
      }
    }

    // Extract contact names (they appear in column A after CONTACTS header)
    if (inContactsSection && cellA && !cellA.includes(':') && !cellA.startsWith('📧') &&
        !cellA.startsWith('📞') && !cellA.startsWith('💼')) {
      // This looks like a contact name
      contacts.push(cellA);
    }
  }

  return contacts;
}

/**
 * TEST: Run this function directly to check canned response config
 */
function testCannedResponseConfig() {
  Logger.log('=== Testing Canned Response Config ===');
  Logger.log('BS_CFG.CANNED_RESPONSE_DOCS = ' + JSON.stringify(BS_CFG.CANNED_RESPONSE_DOCS));
  Logger.log('REFERRAL_PROGRAM docId = ' + (BS_CFG.CANNED_RESPONSE_DOCS ? BS_CFG.CANNED_RESPONSE_DOCS.REFERRAL_PROGRAM : 'CANNED_RESPONSE_DOCS is undefined'));
  Logger.log('BS_CFG.REFERRAL_CONTRACT_FILE_ID = ' + BS_CFG.REFERRAL_CONTRACT_FILE_ID);

  SpreadsheetApp.getUi().alert('Check View > Logs for output');
}

/**
 * Get canned response template from Google Doc
 * Returns { text: plainText, html: htmlContent } or null if not found
 */
function getCannedResponseTemplate_(templateKey) {
  Logger.log(`Looking for template: ${templateKey}`);
  Logger.log(`CANNED_RESPONSE_DOCS: ${JSON.stringify(BS_CFG.CANNED_RESPONSE_DOCS)}`);

  const docId = BS_CFG.CANNED_RESPONSE_DOCS ? BS_CFG.CANNED_RESPONSE_DOCS[templateKey] : null;

  Logger.log(`Found docId: ${docId}`);

  if (!docId) {
    Logger.log(`No Google Doc configured for template: ${templateKey}`);
    return null;
  }

  try {
    // Get the document
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();

    // Get plain text for preview
    const text = body.getText();

    // Build HTML with inline styles directly from document structure
    // This is more reliable than parsing Google's CSS export (Gmail strips <style> blocks)
    let html = '';
    const numChildren = body.getNumChildren();

    for (let i = 0; i < numChildren; i++) {
      const child = body.getChild(i);
      const type = child.getType();

      if (type === DocumentApp.ElementType.PARAGRAPH) {
        const para = child.asParagraph();
        html += buildParagraphHtml_(para);
      } else if (type === DocumentApp.ElementType.LIST_ITEM) {
        const listItem = child.asListItem();
        const itemHtml = buildListItemHtml_(listItem);
        html += itemHtml;
      }
    }

    return { text, html };
  } catch (e) {
    Logger.log(`Error reading template doc ${templateKey}: ${e.message}`);
    return null;
  }
}

/**
 * Build HTML from a paragraph with inline styles
 */
function buildParagraphHtml_(para) {
  const textHtml = buildTextElementsHtml_(para);

  // Check for heading styles
  const heading = para.getHeading();
  if (heading === DocumentApp.ParagraphHeading.HEADING1) {
    return `<p style="font-size:20px;font-weight:bold;margin:16px 0 8px 0;">${textHtml}</p>`;
  } else if (heading === DocumentApp.ParagraphHeading.HEADING2) {
    return `<p style="font-size:16px;font-weight:bold;margin:14px 0 6px 0;">${textHtml}</p>`;
  } else if (heading === DocumentApp.ParagraphHeading.HEADING3) {
    return `<p style="font-size:14px;font-weight:bold;margin:12px 0 4px 0;">${textHtml}</p>`;
  }

  // Empty paragraph = line break
  if (!textHtml.trim()) {
    return '<br>';
  }

  return `<p style="margin:0 0 8px 0;">${textHtml}</p>`;
}

/**
 * Build HTML from a list item with inline styles
 */
function buildListItemHtml_(listItem) {
  const textHtml = buildTextElementsHtml_(listItem);
  const nestingLevel = listItem.getNestingLevel();
  const indent = 20 + (nestingLevel * 20);
  return `<p style="margin:0 0 4px ${indent}px;">• ${textHtml}</p>`;
}

/**
 * Build HTML from text elements with inline styles for bold, italic, underline
 */
function buildTextElementsHtml_(element) {
  let html = '';
  const text = element.getText();
  if (!text) return '';

  // Get the text element to check formatting
  const numChildren = element.getNumChildren();

  if (numChildren === 0) {
    // Simple text without children - check if element itself has formatting
    return escapeHtml_(text);
  }

  // Process each child (usually Text elements)
  for (let i = 0; i < numChildren; i++) {
    const child = element.getChild(i);
    if (child.getType() === DocumentApp.ElementType.TEXT) {
      html += buildTextRunHtml_(child.asText());
    }
  }

  return html;
}

/**
 * Build HTML from a Text element, handling formatting runs
 */
function buildTextRunHtml_(textElement) {
  const text = textElement.getText();
  if (!text) return '';

  let html = '';
  let i = 0;

  while (i < text.length) {
    // Get formatting at this position
    const isBold = textElement.isBold(i);
    const isItalic = textElement.isItalic(i);
    const isUnderline = textElement.isUnderline(i);
    const link = textElement.getLinkUrl(i);
    const fontSize = textElement.getFontSize(i);
    const bgColor = textElement.getBackgroundColor(i);

    // Find end of this formatting run
    let j = i + 1;
    while (j < text.length) {
      if (textElement.isBold(j) !== isBold ||
          textElement.isItalic(j) !== isItalic ||
          textElement.isUnderline(j) !== isUnderline ||
          textElement.getLinkUrl(j) !== link ||
          textElement.getFontSize(j) !== fontSize ||
          textElement.getBackgroundColor(j) !== bgColor) {
        break;
      }
      j++;
    }

    // Get and escape the text for this run
    let runText = escapeHtml_(text.substring(i, j));

    // Apply formatting with inline styles
    let styles = [];
    if (isBold) styles.push('font-weight:bold');
    if (isItalic) styles.push('font-style:italic');
    if (isUnderline) styles.push('text-decoration:underline');
    // Add font size if different from default (11pt)
    if (fontSize && fontSize !== 11) {
      styles.push(`font-size:${fontSize}pt`);
    }
    // Add background color if set
    if (bgColor) {
      styles.push(`background-color:${bgColor}`);
    }

    if (styles.length > 0) {
      runText = `<span style="${styles.join(';')}">${runText}</span>`;
    }

    // Wrap in link if needed
    if (link) {
      runText = `<a href="${link}" style="color:#1a73e8">${runText}</a>`;
    }

    html += runText;
    i = j;
  }

  return html;
}

/**
 * Escape HTML special characters
 */
function escapeHtml_(text) {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

/**
 * Main function to handle canned response generation
 * Shows dialog for contact selection, then creates draft
 */
function generateCannedResponse_(templateKey, templateName) {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  try {
    // Check for future meeting today before proceeding
    const futureMeetingWarning = checkForFutureMeetingToday_();
    if (futureMeetingWarning) {
      const response = ui.alert(
        '⚠️ Meeting Today',
        `${futureMeetingWarning}\n\nDo you still want to send an email response?`,
        ui.ButtonSet.YES_NO
      );
      if (response !== ui.Button.YES) {
        return;
      }
    }

    // Get template from Google Doc
    const templateData = getCannedResponseTemplate_(templateKey);
    if (!templateData) {
      ui.alert('Error', `Template "${templateKey}" not configured.\n\nPlease set up a Google Doc with the template and add its ID to CANNED_RESPONSE_DOCS in the config.`, ui.ButtonSet.OK);
      return;
    }
    const template = templateData.text; // Use plain text for preview

    // Get current vendor
    const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
    const currentIndex = getCurrentVendorIndex_();
    if (!currentIndex || !listSh) {
      ui.alert('Error', 'Could not determine current vendor.', ui.ButtonSet.OK);
      return;
    }
    const listRow = currentIndex + 1;
    const vendor = String(listSh.getRange(listRow, BS_CFG.L_VENDOR + 1).getValue() || '').trim();

    // Get contacts for this vendor
    const contacts = getContactsForCurrentVendor_();

    // Try to get selected email thread (optional - can create fresh draft if none)
    ss.toast('Checking for selected email...', '📨 Canned Response', 2);
    let emailData = null;
    let threadId = null;
    try {
      emailData = getSelectedEmailThread_();
      threadId = emailData.threadId;
    } catch (e) {
      // No email selected - will create fresh draft
      Logger.log('No email selected, will create fresh draft: ' + e.message);
    }

    // Show contact selection dialog
    // Extract first names only for dropdown
    const contactOptions = contacts.length > 0
      ? contacts.map(c => {
          const firstName = c.split(' ')[0]; // Get first name only
          return `<option value="${escapeHtml_(firstName)}">${escapeHtml_(c)} (${escapeHtml_(firstName)})</option>`;
        }).join('')
      : '<option value="">No contacts found</option>';

    const replyInfo = threadId
      ? `<div class="info reply-info">📧 Replying to: <strong>${escapeHtml_(emailData.subject || 'Selected email')}</strong></div>`
      : '<div class="info new-info">📝 No email selected - will create a NEW draft</div>';

    const html = `<!DOCTYPE html>
<html>
<head>
  <base target="_blank">
  <style>
    body { font-family: Arial, sans-serif; padding: 20px; }
    .header { color: #202124; font-size: 16px; margin-bottom: 15px; }
    label { display: block; margin-bottom: 5px; font-weight: bold; }
    select, input {
      width: 100%;
      padding: 10px;
      margin-bottom: 15px;
      border: 1px solid #ddd;
      border-radius: 4px;
      font-size: 14px;
      box-sizing: border-box;
    }
    .buttons { display: flex; gap: 10px; margin-top: 15px; }
    .btn {
      padding: 12px 24px;
      border-radius: 4px;
      font-size: 14px;
      cursor: pointer;
      border: none;
    }
    .btn-primary { background: #1a73e8; color: white; }
    .btn-primary:hover { background: #1557b0; }
    .btn-primary:disabled { background: #ccc; cursor: not-allowed; }
    .btn-secondary { background: #f1f3f4; color: #5f6368; }
    .btn-secondary:hover { background: #e8eaed; }
    .preview {
      background: #f8f9fa;
      padding: 15px;
      border-radius: 8px;
      margin-top: 15px;
      white-space: pre-wrap;
      font-family: Arial, sans-serif;
      font-size: 13px;
      max-height: 250px;
      overflow-y: auto;
      border: 1px solid #e0e0e0;
    }
    .info { color: #5f6368; font-size: 12px; margin-bottom: 10px; }
    .reply-info { color: #1a73e8; }
    .new-info { color: #f9a825; }
    #status { margin-top: 10px; font-size: 13px; }
  </style>
</head>
<body>
  <div class="header">📨 ${escapeHtml_(templateName)}</div>
  <div class="info">Vendor: <strong>${escapeHtml_(vendor)}</strong></div>
  ${replyInfo}

  <label for="contactSelect">Select Contact (first name will be used):</label>
  <select id="contactSelect" onchange="updatePreview()">
    ${contactOptions}
  </select>

  <label for="customContact">Or enter custom first name:</label>
  <input type="text" id="customContact" placeholder="Enter first name..." oninput="updatePreview()">

  <div class="preview" id="previewBox"></div>

  <div class="buttons">
    <button class="btn btn-primary" id="createBtn" onclick="createDraft()">✉️ Create Draft</button>
    <button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
  </div>
  <div id="status"></div>

  <script>
    const template = ${JSON.stringify(template)};
    const vendor = ${JSON.stringify(vendor)};
    const threadId = ${JSON.stringify(threadId)};
    const templateKey = ${JSON.stringify(templateKey)};

    function getContactName() {
      const custom = document.getElementById('customContact').value.trim();
      if (custom) return custom.split(' ')[0]; // First name only
      const select = document.getElementById('contactSelect');
      return select.value || 'there';
    }

    function updatePreview() {
      const contactName = getContactName();
      let preview = template
        .replace(/<CONTACT_NAME>/g, contactName)
        .replace(/<VENDOR_NAME>/g, vendor);
      document.getElementById('previewBox').textContent = preview;
    }

    function createDraft() {
      const btn = document.getElementById('createBtn');
      const status = document.getElementById('status');
      btn.disabled = true;
      btn.textContent = '⏳ Creating...';
      status.textContent = 'Creating draft...';
      status.style.color = '#666';

      const contactName = getContactName();

      google.script.run
        .withSuccessHandler(function(result) {
          if (result && result.success) {
            status.textContent = '✅ Draft created! Opening...';
            status.style.color = 'green';
            // Open draft in new tab
            if (result.draftUrl) {
              window.open(result.draftUrl, '_blank');
            }
            setTimeout(function() {
              google.script.host.close();
            }, 500);
          } else {
            status.textContent = '❌ Error: ' + (result ? result.error : 'Unknown error');
            status.style.color = 'red';
            btn.disabled = false;
            btn.textContent = '✉️ Create Draft';
          }
        })
        .withFailureHandler(function(error) {
          status.textContent = '❌ Error: ' + error.message;
          status.style.color = 'red';
          btn.disabled = false;
          btn.textContent = '✉️ Create Draft';
        })
        .createCannedResponseDraft(threadId, templateKey, contactName, vendor);
    }

    // Initial preview
    updatePreview();
  </script>
</body>
</html>`;

    const htmlOutput = HtmlService.createHtmlOutput(html)
      .setWidth(650)
      .setHeight(750);
    ui.showModalDialog(htmlOutput, 'Canned Response');

  } catch (e) {
    ui.alert('Error', e.message, ui.ButtonSet.OK);
  }
}

/**
 * Create draft from canned response template
 * Called from dialog
 * If threadId is null, creates a fresh draft instead of a reply
 */
function createCannedResponseDraft(threadId, templateKey, contactName, vendor) {
  try {
    // Get template from Google Doc (with HTML formatting)
    const templateData = getCannedResponseTemplate_(templateKey);
    if (!templateData) {
      return { success: false, error: 'Template not configured. Please set up a Google Doc.' };
    }

    // Replace placeholders in both plain text and HTML
    let plainBody = templateData.text
      .replace(/<CONTACT_NAME>/g, contactName)
      .replace(/<VENDOR_NAME>/g, vendor);

    let htmlBody = templateData.html
      .replace(/&lt;CONTACT_NAME&gt;/g, contactName)
      .replace(/&lt;VENDOR_NAME&gt;/g, vendor)
      .replace(/<CONTACT_NAME>/g, contactName)
      .replace(/<VENDOR_NAME>/g, vendor);

    // Get Gmail signature and append it
    const myEmail = Session.getActiveUser().getEmail().toLowerCase();
    let signature = '';
    try {
      const sendAsSettings = Gmail.Users.Settings.SendAs.list('me');
      if (sendAsSettings && sendAsSettings.sendAs) {
        const primarySendAs = sendAsSettings.sendAs.find(s => s.isPrimary) ||
                              sendAsSettings.sendAs.find(s => s.sendAsEmail.toLowerCase() === myEmail) ||
                              sendAsSettings.sendAs[0];
        if (primarySendAs && primarySendAs.signature) {
          signature = primarySendAs.signature;
        }
      }
    } catch (e) {
      Logger.log('Could not fetch Gmail signature: ' + e.message);
    }

    // Append signature to body
    if (signature) {
      plainBody += '\n\n' + signature.replace(/<[^>]*>/g, ''); // Strip HTML for plain text
      htmlBody += '<br><br>' + signature;
    }

    // Get attachment if configured for this template
    let attachments = [];
    let attachmentFileId = null;
    if (templateKey === 'REFERRAL_PROGRAM' && BS_CFG.REFERRAL_CONTRACT_FILE_ID) {
      attachmentFileId = BS_CFG.REFERRAL_CONTRACT_FILE_ID;
    } else if (templateKey === 'INITIAL_CALL_FOLLOWUP' && BS_CFG.INITIAL_CALL_FOLLOWUP_FILE_ID) {
      attachmentFileId = BS_CFG.INITIAL_CALL_FOLLOWUP_FILE_ID;
    }

    if (attachmentFileId) {
      try {
        const file = DriveApp.getFileById(attachmentFileId);
        attachments.push(file.getBlob());
        Logger.log('Attached file: ' + file.getName());
      } catch (e) {
        Logger.log('Could not attach file: ' + e.message);
      }
    }

    let draftUrl;

    if (threadId) {
      // Reply to existing thread - use Gmail API for proper threading and quoted content
      const thread = GmailApp.getThreadById(threadId);
      if (!thread) {
        return { success: false, error: 'Email thread not found' };
      }

      const messages = thread.getMessages();
      const lastMessage = messages[messages.length - 1];

      // Build the quoted original message
      const lastMsgDate = lastMessage.getDate();
      const lastMsgFrom = lastMessage.getFrom();
      const lastMsgBody = lastMessage.getBody() || '';

      // Format date for quote header
      const dateStr = Utilities.formatDate(lastMsgDate, Session.getScriptTimeZone(), "EEE, MMM d, yyyy 'at' h:mm a");

      // Helper to escape HTML special characters
      const escapeHtml = (text) => {
        return text
          .replace(/&/g, '&amp;')
          .replace(/</g, '&lt;')
          .replace(/>/g, '&gt;')
          .replace(/"/g, '&quot;');
      };

      // Build quoted message using Gmail's standard format
      const quoteHeaderHtml = `<div class="gmail_quote"><div dir="ltr" class="gmail_attr">On ${escapeHtml(dateStr)}, ${escapeHtml(lastMsgFrom)} wrote:<br></div><blockquote class="gmail_quote" style="margin:0px 0px 0px 0.8ex;border-left:1px solid rgb(204,204,204);padding-left:1ex">`;
      const quotedBodyHtml = `${lastMsgBody}</blockquote></div>`;

      // Full email body: template + signature + quoted message
      let fullBodyHtml = `<div dir="ltr">${htmlBody}</div><br>${quoteHeaderHtml}${quotedBodyHtml}`;

      // Step 1: Create placeholder draft for threading headers
      const draftReply = thread.createDraftReplyAll('Placeholder body - will be replaced');
      const draftMsg = draftReply.getMessage();
      const draftId = draftReply.getId();

      // Get threading headers
      const inReplyTo = draftMsg.getHeader('In-Reply-To') || '';
      const references = draftMsg.getHeader('References') || '';

      // Get subject from thread
      let subject = thread.getFirstMessageSubject();
      if (!subject.toLowerCase().startsWith('re:')) {
        subject = 'Re: ' + subject;
      }

      // Build recipients from last message
      const lastFrom = lastMessage.getFrom();
      const lastTo = lastMessage.getTo() || '';
      const lastCc = lastMessage.getCc() || '';
      const allRecipients = [lastFrom, lastTo, lastCc].filter(r => r).join(', ');

      // Extract email addresses, exclude ourselves and profitise emails (they go in CC/BCC)
      const emailRegex = /[\w.-]+@[\w.-]+\.\w+/g;
      const allEmails = allRecipients.match(emailRegex) || [];
      const toEmails = [];
      allEmails.forEach(email => {
        const lowerEmail = email.toLowerCase();
        if (lowerEmail !== myEmail &&
            !lowerEmail.includes('sales@profitise.com') &&
            !lowerEmail.includes('aden@profitise.com') &&
            !toEmails.includes(lowerEmail)) {
          toEmails.push(email);
        }
      });
      const toRecipients = toEmails.join(', ') || lastFrom;

      // Build raw email with headers
      let rawHeaders = `From: ${myEmail}\r\n`;
      rawHeaders += `To: ${toRecipients}\r\n`;
      rawHeaders += `Cc: aden@profitise.com\r\n`;
      rawHeaders += `Bcc: sales@profitise.com\r\n`;
      rawHeaders += `Subject: ${subject}\r\n`;
      if (inReplyTo) rawHeaders += `In-Reply-To: ${inReplyTo}\r\n`;
      if (references) rawHeaders += `References: ${references}\r\n`;
      rawHeaders += `MIME-Version: 1.0\r\n`;

      // Handle attachments with MIME multipart
      if (attachments.length > 0) {
        const boundary = 'boundary_' + Utilities.getUuid();
        rawHeaders += `Content-Type: multipart/mixed; boundary="${boundary}"\r\n\r\n`;

        let mimeBody = `--${boundary}\r\n`;
        mimeBody += `Content-Type: text/html; charset="UTF-8"\r\n`;
        mimeBody += `Content-Transfer-Encoding: base64\r\n\r\n`;
        // Encode HTML content as base64 for proper UTF-8 handling
        mimeBody += Utilities.base64Encode(Utilities.newBlob(fullBodyHtml).getBytes()) + '\r\n';

        attachments.forEach(blob => {
          mimeBody += `--${boundary}\r\n`;
          mimeBody += `Content-Type: ${blob.getContentType()}; name="${blob.getName()}"\r\n`;
          mimeBody += `Content-Disposition: attachment; filename="${blob.getName()}"\r\n`;
          mimeBody += `Content-Transfer-Encoding: base64\r\n\r\n`;
          mimeBody += Utilities.base64Encode(blob.getBytes()) + '\r\n';
        });
        mimeBody += `--${boundary}--`;

        // Convert full message to UTF-8 bytes for proper encoding
        const rawEmail = rawHeaders + mimeBody;
        const updateResource = {
          message: {
            raw: Utilities.base64EncodeWebSafe(Utilities.newBlob(rawEmail).getBytes()),
            threadId: thread.getId()
          }
        };
        const updatedDraft = Gmail.Users.Drafts.update(updateResource, 'me', draftId);
        draftUrl = `https://mail.google.com/mail/u/0/#drafts?compose=${updatedDraft.message.id}`;
      } else {
        rawHeaders += `Content-Type: text/html; charset="UTF-8"\r\n\r\n`;

        // Convert full message to UTF-8 bytes for proper encoding of special characters
        const rawEmail = rawHeaders + fullBodyHtml;
        const updateResource = {
          message: {
            raw: Utilities.base64EncodeWebSafe(Utilities.newBlob(rawEmail).getBytes()),
            threadId: thread.getId()
          }
        };
        const updatedDraft = Gmail.Users.Drafts.update(updateResource, 'me', draftId);
        draftUrl = `https://mail.google.com/mail/u/0/#drafts?compose=${updatedDraft.message.id}`;
      }

      // Track this thread for auto-archive after sending
      addPendingArchiveThread_(thread.getId());
      Logger.log(`Added thread ${thread.getId()} to pending archive list (canned response)`);

    } else {
      // Create fresh draft (no reply)
      const subject = `Referral Partnership - ${vendor}`;
      const options = {
        cc: 'aden@profitise.com',
        bcc: 'sales@profitise.com',
        htmlBody: htmlBody
      };
      if (attachments.length > 0) {
        options.attachments = attachments;
      }

      const draft = GmailApp.createDraft('', subject, plainBody, options);
      const draftId = draft.getMessage().getId();
      draftUrl = `https://mail.google.com/mail/u/0/#drafts?compose=${draftId}`;
    }

    SpreadsheetApp.getActive().toast('Draft created! Opening in Gmail...', '✅ Draft Ready', 3);

    return { success: true, draftUrl: draftUrl };
  } catch (e) {
    Logger.log(`Error creating canned response draft: ${e.message}`);
    return { success: false, error: e.message };
  }
}

/************************************************************
 * CONTACT DISCOVERY FROM GMAIL
 * Searches Gmail for potential new contacts or updated info
 * based on domains of existing contacts
 ************************************************************/

/**
 * Discover potential new contacts or updates from Gmail
 * Searches for emails from the same domains as existing contacts
 */
function discoverContactsFromGmail() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);

  if (!listSh) {
    ui.alert('Vendor list not found.');
    return;
  }

  const currentIndex = getCurrentVendorIndex_();
  const listRow = currentIndex + 1;
  const vendor = listSh.getRange(listRow, 1).getValue();
  const source = listSh.getRange(listRow, 2).getValue() || '';

  ss.toast('Analyzing contacts and searching Gmail...', '🔍 Contact Discovery', 5);

  // Get existing contacts for this vendor
  const contactData = getVendorContacts_(vendor, listRow);
  const existingContacts = contactData.contacts || [];

  // Extract unique domains from existing contacts
  const domains = new Set();
  const existingEmails = new Set();
  const existingPhones = new Set();

  for (const contact of existingContacts) {
    if (contact.email) {
      existingEmails.add(contact.email.toLowerCase());
      const domain = contact.email.split('@')[1];
      if (domain && !isGenericDomain_(domain)) {
        domains.add(domain.toLowerCase());
      }
    }
    if (contact.phone) {
      // Normalize phone to just digits for comparison
      const normalizedPhone = contact.phone.replace(/\D/g, '');
      if (normalizedPhone.length >= 10) {
        existingPhones.add(normalizedPhone.slice(-10)); // Last 10 digits
      }
    }
  }

  // Also extract domain from source/website field
  if (source) {
    const sourceDomain = extractDomainFromSource_(source);
    if (sourceDomain && !isGenericDomain_(sourceDomain)) {
      domains.add(sourceDomain.toLowerCase());
      Logger.log(`Added source domain: ${sourceDomain}`);
    }
  }

  // Also try to extract domain from vendor name (e.g., "HomeFix" -> "homefix.com")
  const vendorDomain = guessVendorDomain_(vendor);
  if (vendorDomain && !isGenericDomain_(vendorDomain)) {
    domains.add(vendorDomain.toLowerCase());
    Logger.log(`Added guessed vendor domain: ${vendorDomain}`);
  }

  if (domains.size === 0) {
    ui.alert('No domains found to search. Add a contact with a company email or update the source field.');
    return;
  }

  Logger.log(`=== CONTACT DISCOVERY ===`);
  Logger.log(`Vendor: ${vendor}`);
  Logger.log(`Existing contacts: ${existingContacts.length}`);
  Logger.log(`Domains to search: ${[...domains].join(', ') || '(none - searching specific emails)'}`);
  Logger.log(`Existing emails: ${[...existingEmails].join(', ')}`);

  // Search Gmail for emails from these domains
  const potentialNewContacts = [];
  const potentialUpdates = [];
  const emailsSearched = new Set();

  // If we have company domains, search by domain
  // Helper to extract all emails from a header field (handles "Name <email>, Name2 <email2>" format)
  const extractEmailsFromField = (field) => {
    if (!field) return [];
    const results = [];

    // Preprocess: normalize whitespace (replace newlines, tabs, multiple spaces with single space)
    const normalized = field.replace(/[\r\n\t]+/g, ' ').replace(/\s+/g, ' ').trim();

    // Split by comma (but be careful of commas in quoted names)
    const parts = normalized.split(/,(?=(?:[^"]*"[^"]*")*[^"]*$)/);

    for (const part of parts) {
      const trimmed = part.trim();
      if (!trimmed) continue;

      // Try to match "Name" <email> or Name <email> format
      const angleMatch = trimmed.match(/^(?:"?([^"<]+)"?\s*)?<([^>]+)>$/);
      if (angleMatch) {
        let name = angleMatch[1] ? angleMatch[1].trim() : '';
        const email = angleMatch[2].toLowerCase().trim();

        // Validate email has proper format (at least 2 chars before @, valid domain)
        if (email.includes('@')) {
          const [localPart, domain] = email.split('@');
          // Skip if local part is too short (likely malformed) or domain looks incomplete
          if (localPart && localPart.length >= 2 && domain && domain.includes('.')) {
            results.push({ name, email });
          } else {
            Logger.log(`Skipping malformed email from header: "${email}" (parsed from: "${trimmed}")`);
          }
        }
        continue;
      }

      // Try standalone email
      const emailOnlyMatch = trimmed.match(/^([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})$/);
      if (emailOnlyMatch) {
        const email = emailOnlyMatch[1].toLowerCase();
        const [localPart] = email.split('@');
        // Validate local part is reasonable length
        if (localPart && localPart.length >= 2) {
          results.push({ name: '', email });
        } else {
          Logger.log(`Skipping malformed standalone email: "${email}"`);
        }
      }
    }
    return results;
  };

  // Helper to clean contact name - only allow letters, spaces, hyphens, apostrophes
  const cleanContactName = (name) => {
    if (!name) return '';
    // Remove any non-name characters (keep letters, spaces, hyphens, apostrophes)
    let cleaned = name.replace(/[^a-zA-Z\s\-']/g, '').trim();
    // Collapse multiple spaces
    cleaned = cleaned.replace(/\s+/g, ' ');
    // Capitalize first letter of each word
    cleaned = cleaned.replace(/\b\w/g, c => c.toUpperCase());
    return cleaned;
  };

  for (const domain of domains) {
    ss.toast(`Searching emails with @${domain}...`, '🔍 Scanning', 3);

    try {
      // Search for emails involving this domain (from OR to)
      const query = `{from:@${domain} to:@${domain} cc:@${domain}}`;
      const threads = GmailApp.search(query, 0, 50); // Get up to 50 threads

      Logger.log(`Found ${threads.length} threads involving @${domain}`);

      for (const thread of threads) {
        const messages = thread.getMessages();

        for (const message of messages) {
          const from = message.getFrom() || '';
          const to = message.getTo() || '';
          const cc = message.getCc() || '';
          const body = message.getPlainBody() || '';
          const date = message.getDate();

          // Extract all email addresses from From, To, and CC fields
          const allAddresses = [
            ...extractEmailsFromField(from),
            ...extractEmailsFromField(to),
            ...extractEmailsFromField(cc)
          ];

          for (const addr of allAddresses) {
            const email = addr.email;

            // Skip if we've already processed this email address
            if (emailsSearched.has(email)) continue;
            emailsSearched.add(email);

            // Only consider contacts from the vendor's domains (not internal or other external)
            const emailDomain = email.split('@')[1];
            if (!emailDomain || !domains.has(emailDomain.toLowerCase())) continue;

            // Use extracted name or derive from email, then clean it
            const rawName = addr.name || email.split('@')[0].replace(/[._]/g, ' ');
            const name = cleanContactName(rawName);

            // Check if this is a new contact
            if (!existingEmails.has(email)) {
              // Try to find signature info for this contact (only reliable for senders)
              let phones = [];
              let jobTitle = '';

              // If this person sent a message, try to extract from their signature
              if (from.toLowerCase().includes(email)) {
                const signature = body.slice(-500);
                phones = extractPhoneNumbers_(signature);
                jobTitle = extractJobTitle_(signature, name);
              }

              potentialNewContacts.push({
                name: name,
                email: email,
                phone: phones.length > 0 ? phones[0] : '',
                jobTitle: jobTitle,
                lastSeen: Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
                domain: domain
              });

              Logger.log(`Potential new contact: ${name} <${email}> - ${jobTitle || 'no title'}`);
            } else {
              // Existing contact - check for phone updates and job title (only from their sent messages)
              if (!from.toLowerCase().includes(email)) continue;

              const signature = body.slice(-500);
              const phones = extractPhoneNumbers_(signature);
              const jobTitle = extractJobTitle_(signature, name);

              // Check for job title updates
              const existingContact = existingContacts.find(c =>
                c.email && c.email.toLowerCase() === email
              );

              if (existingContact && jobTitle && !existingContact.notes?.includes(jobTitle)) {
                potentialUpdates.push({
                  name: existingContact.name,
                  email: email,
                  updateType: 'jobTitle',
                  currentValue: existingContact.notes || '(none)',
                  newValue: jobTitle,
                  foundIn: Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd')
                });
                Logger.log(`Potential job title update for ${existingContact.name}: ${jobTitle}`);
              }

              for (const phone of phones) {
                const normalized = phone.replace(/\D/g, '').slice(-10);
                if (normalized.length === 10 && !existingPhones.has(normalized)) {
                  // Found a new phone number for an existing contact
                  if (existingContact) {
                    potentialUpdates.push({
                      name: existingContact.name,
                      email: email,
                      updateType: 'phone',
                      currentValue: existingContact.phone || '(none)',
                      newValue: phone,
                      foundIn: Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd')
                    });

                    Logger.log(`Potential phone update for ${existingContact.name}: ${phone}`);
                  }
                }
              }
            }
          }
        }
      }
    } catch (e) {
      Logger.log(`Error searching domain ${domain}: ${e.message}`);
    }
  }

  // If no company domains were found, search for specific existing email addresses
  // This allows us to find updates (job titles, phones) for contacts with generic domains
  if (domains.size === 0 && existingEmails.size > 0) {
    ss.toast('Searching for existing contacts in Gmail...', '🔍 Scanning', 3);

    for (const email of existingEmails) {
      if (emailsSearched.has(email)) continue;
      emailsSearched.add(email);

      try {
        // Search for emails from this specific address
        const query = `from:${email}`;
        const threads = GmailApp.search(query, 0, 20);

        Logger.log(`Found ${threads.length} threads from ${email}`);

        // Find the existing contact for this email
        const existingContact = existingContacts.find(c => c.email && c.email.toLowerCase() === email);
        if (!existingContact) continue;

        for (const thread of threads) {
          const messages = thread.getMessages();

          for (const message of messages) {
            const from = message.getFrom() || '';
            const body = message.getPlainBody() || '';
            const date = message.getDate();

            // Make sure this message is from the contact we're searching for
            const msgEmailMatch = from.match(/<([^>]+)>/) || [null, from.trim()];
            const msgEmail = msgEmailMatch[1].toLowerCase();
            if (msgEmail !== email) continue;

            // Extract info from signature
            const signature = body.slice(-500);

            // Check for job title update
            if (!existingContact.name?.includes('|')) {
              const jobTitle = extractJobTitle_(signature, existingContact.name);
              if (jobTitle && jobTitle.length > 3) {
                // Check if we already have this update
                const existingUpdate = potentialUpdates.find(u =>
                  u.email === email && u.updateType === 'jobTitle'
                );
                if (!existingUpdate) {
                  potentialUpdates.push({
                    contactId: existingContact.id,
                    name: existingContact.name,
                    email: email,
                    updateType: 'jobTitle',
                    currentValue: '(none)',
                    newValue: jobTitle,
                    foundIn: Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd')
                  });
                  Logger.log(`Potential job title update for ${existingContact.name}: ${jobTitle}`);
                }
              }
            }

            // Check for phone update
            if (!existingContact.phone) {
              const phones = extractPhoneNumbers_(signature);
              if (phones.length > 0) {
                const phone = phones[0];
                const normalizedNewPhone = phone.replace(/\D/g, '').slice(-10);
                if (!existingPhones.has(normalizedNewPhone)) {
                  const existingUpdate = potentialUpdates.find(u =>
                    u.email === email && u.updateType === 'phone'
                  );
                  if (!existingUpdate) {
                    potentialUpdates.push({
                      contactId: existingContact.id,
                      name: existingContact.name,
                      email: email,
                      updateType: 'phone',
                      currentValue: '(none)',
                      newValue: phone,
                      foundIn: Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd')
                    });
                    Logger.log(`Potential phone update for ${existingContact.name}: ${phone}`);
                  }
                }
              }
            }
          }
        }
      } catch (e) {
        Logger.log(`Error searching email ${email}: ${e.message}`);
      }
    }
  }

  // Sort results
  potentialNewContacts.sort((a, b) => b.lastSeen.localeCompare(a.lastSeen)); // Most recent first
  potentialUpdates.sort((a, b) => b.foundIn.localeCompare(a.foundIn));

  // Get Monday.com IDs for linking contacts
  const mondayItemId = contactData.mondayItemId;
  const boardId = contactData.boardId;

  // Prepare data for the dialog
  const dialogData = {
    vendor: vendor,
    mondayItemId: mondayItemId,
    boardId: boardId,
    domains: [...domains],
    existingContacts: existingContacts,
    potentialNewContacts: potentialNewContacts.slice(0, 20),
    potentialUpdates: potentialUpdates.slice(0, 20)
  };

  // Build results HTML with interactive elements
  let html = `
    <style>
      body { font-family: Arial, sans-serif; font-size: 13px; padding: 10px; }
      h2 { color: #1a73e8; margin-top: 20px; margin-bottom: 10px; }
      h3 { color: #333; margin-top: 15px; margin-bottom: 8px; }
      table { border-collapse: collapse; width: 100%; margin-bottom: 15px; }
      th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
      th { background-color: #f8f9fa; font-weight: bold; }
      tr:nth-child(even) { background-color: #f9f9f9; }
      .new-contact { background-color: #e8f5e9; }
      .update { background-color: #fff3e0; }
      .none { color: #666; font-style: italic; }
      .summary { background: #f0f0f0; padding: 10px; border-radius: 5px; margin-bottom: 15px; }
      .btn { padding: 8px 16px; margin: 5px; cursor: pointer; border: none; border-radius: 4px; font-size: 13px; }
      .btn-primary { background: #1a73e8; color: white; }
      .btn-primary:hover { background: #1557b0; }
      .btn-primary:disabled { background: #ccc; cursor: not-allowed; }
      .btn-secondary { background: #ff9800; color: white; }
      .btn-secondary:hover { background: #f57c00; }
      .btn-secondary:disabled { background: #ccc; cursor: not-allowed; }
      .checkbox-cell { width: 30px; text-align: center; }
      .job-title { color: #666; font-size: 12px; }
      .status-msg { padding: 10px; margin: 10px 0; border-radius: 4px; }
      .status-success { background: #d4edda; color: #155724; }
      .status-error { background: #f8d7da; color: #721c24; }
    </style>

    <h2>📇 Contact Discovery for ${escapeHtml_(vendor)}</h2>

    <div class="summary">
      <strong>Searched:</strong> ${domains.size > 0 ? [...domains].map(d => '@' + d).join(', ') : 'Specific email addresses (no company domains)'}<br>
      <strong>Existing contacts:</strong> ${existingContacts.length}<br>
      <strong>Potential new contacts found:</strong> ${potentialNewContacts.length}<br>
      <strong>Potential updates found:</strong> ${potentialUpdates.length}
    </div>

    <div id="statusMsg"></div>
  `;

  // New contacts section with checkboxes
  html += `<h3>🆕 Potential New Contacts (${potentialNewContacts.length})</h3>`;

  if (potentialNewContacts.length === 0) {
    html += `<p class="none">No new contacts discovered.</p>`;
  } else {
    html += `
      <div style="margin-bottom: 10px;">
        <button class="btn btn-primary" onclick="addSelectedContacts()" id="addBtn">➕ Add Selected to Monday.com</button>
        <label style="margin-left: 15px;"><input type="checkbox" id="selectAll" onchange="toggleSelectAll()"> Select All</label>
      </div>
      <table id="newContactsTable">
        <tr><th class="checkbox-cell"></th><th>Name</th><th>Email</th><th>Phone</th><th>Job Title</th><th>Last Seen</th></tr>
    `;
    for (let i = 0; i < Math.min(potentialNewContacts.length, 20); i++) {
      const contact = potentialNewContacts[i];
      html += `
        <tr class="new-contact">
          <td class="checkbox-cell"><input type="checkbox" class="contact-cb" data-index="${i}"></td>
          <td>${escapeHtml_(contact.name)}</td>
          <td>${escapeHtml_(contact.email)}</td>
          <td>${contact.phone || '<span class="none">-</span>'}</td>
          <td class="job-title">${contact.jobTitle ? escapeHtml_(contact.jobTitle) : '<span class="none">-</span>'}</td>
          <td>${contact.lastSeen}</td>
        </tr>
      `;
    }
    html += `</table>`;
    if (potentialNewContacts.length > 20) {
      html += `<p class="none">...and ${potentialNewContacts.length - 20} more</p>`;
    }
  }

  // Updates section (phone and job title)
  const phoneUpdates = potentialUpdates.filter(u => u.updateType === 'phone');
  const titleUpdates = potentialUpdates.filter(u => u.updateType === 'jobTitle');

  html += `<h3>📝 Potential Updates (${potentialUpdates.length})</h3>`;

  if (potentialUpdates.length === 0) {
    html += `<p class="none">No updates discovered.</p>`;
  } else {
    html += `
      <div style="margin-bottom: 10px;">
        <button class="btn btn-secondary" onclick="applySelectedUpdates()" id="updateBtn">📝 Apply Selected Updates</button>
        <label style="margin-left: 15px;"><input type="checkbox" id="selectAllUpdates" onchange="toggleSelectAllUpdates()"> Select All</label>
      </div>
      <table id="updatesTable">
        <tr><th class="checkbox-cell"></th><th>Name</th><th>Email</th><th>Field</th><th>Current</th><th>New Value</th><th>Found</th></tr>
    `;
    for (let i = 0; i < Math.min(potentialUpdates.length, 20); i++) {
      const update = potentialUpdates[i];
      const typeLabel = update.updateType === 'phone' ? '📱 Phone' : '💼 Job Title';
      html += `
        <tr class="update">
          <td class="checkbox-cell"><input type="checkbox" class="update-cb" data-index="${i}"></td>
          <td>${escapeHtml_(update.name)}</td>
          <td>${escapeHtml_(update.email)}</td>
          <td>${typeLabel}</td>
          <td>${escapeHtml_(update.currentValue)}</td>
          <td><strong>${escapeHtml_(update.newValue)}</strong></td>
          <td>${update.foundIn}</td>
        </tr>
      `;
    }
    html += `</table>`;
    if (potentialUpdates.length > 20) {
      html += `<p class="none">...and ${potentialUpdates.length - 20} more</p>`;
    }
  }

  html += `
    <h3>👤 Existing Contacts</h3>
    <table>
      <tr><th>Name</th><th>Email</th><th>Phone</th><th>Type</th><th>Status</th></tr>
  `;
  for (const contact of existingContacts) {
    html += `
      <tr>
        <td>${escapeHtml_(contact.name || '')}</td>
        <td>${escapeHtml_(contact.email || '')}</td>
        <td>${escapeHtml_(contact.phone || '')}</td>
        <td>${escapeHtml_(contact.contactType || '')}</td>
        <td>${escapeHtml_(contact.status || '')}</td>
      </tr>
    `;
  }
  html += `</table>`;

  // Add Done button for hard refresh
  html += `
    <div style="margin-top: 20px; padding-top: 15px; border-top: 1px solid #ddd; text-align: center;">
      <button id="done-btn" class="btn btn-primary" style="padding: 10px 30px;">
        ✓ Done - Refresh View
      </button>
    </div>
  `;

  // Add JavaScript for interactivity
  html += `
    <script>
      const dialogData = ${JSON.stringify(dialogData)};

      function toggleSelectAll() {
        const selectAll = document.getElementById('selectAll').checked;
        document.querySelectorAll('.contact-cb').forEach(cb => cb.checked = selectAll);
      }

      function addSelectedContacts() {
        const selected = [];
        document.querySelectorAll('.contact-cb:checked').forEach(cb => {
          const idx = parseInt(cb.dataset.index);
          selected.push(dialogData.potentialNewContacts[idx]);
        });

        if (selected.length === 0) {
          alert('Please select at least one contact to add.');
          return;
        }

        document.getElementById('addBtn').disabled = true;
        document.getElementById('addBtn').textContent = 'Adding...';
        document.getElementById('statusMsg').innerHTML = '<div class="status-msg">Adding ' + selected.length + ' contact(s)...</div>';

        google.script.run
          .withSuccessHandler(function(result) {
            document.getElementById('statusMsg').innerHTML = '<div class="status-msg status-success">' + result + '</div>';
            document.getElementById('addBtn').textContent = '✓ Done';
            // Uncheck added contacts
            document.querySelectorAll('.contact-cb:checked').forEach(cb => {
              cb.checked = false;
              cb.closest('tr').style.opacity = '0.5';
            });
          })
          .withFailureHandler(function(error) {
            document.getElementById('statusMsg').innerHTML = '<div class="status-msg status-error">Error: ' + error.message + '</div>';
            document.getElementById('addBtn').disabled = false;
            document.getElementById('addBtn').textContent = '➕ Add Selected to Monday.com';
          })
          .addContactsToMonday(selected, dialogData.mondayItemId, dialogData.boardId);
      }

      function toggleSelectAllUpdates() {
        const selectAll = document.getElementById('selectAllUpdates').checked;
        document.querySelectorAll('.update-cb').forEach(cb => cb.checked = selectAll);
      }

      function applySelectedUpdates() {
        const selected = [];
        document.querySelectorAll('.update-cb:checked').forEach(cb => {
          const idx = parseInt(cb.dataset.index);
          selected.push(dialogData.potentialUpdates[idx]);
        });

        if (selected.length === 0) {
          alert('Please select at least one update to apply.');
          return;
        }

        document.getElementById('updateBtn').disabled = true;
        document.getElementById('updateBtn').textContent = 'Applying...';
        document.getElementById('statusMsg').innerHTML = '<div class="status-msg">Applying ' + selected.length + ' update(s)...</div>';

        google.script.run
          .withSuccessHandler(function(result) {
            document.getElementById('statusMsg').innerHTML = '<div class="status-msg status-success">' + result + '</div>';
            document.getElementById('updateBtn').textContent = '✓ Done';
            // Uncheck applied updates
            document.querySelectorAll('.update-cb:checked').forEach(cb => {
              cb.checked = false;
              cb.closest('tr').style.opacity = '0.5';
            });
          })
          .withFailureHandler(function(error) {
            document.getElementById('statusMsg').innerHTML = '<div class="status-msg status-error">Error: ' + error.message + '</div>';
            document.getElementById('updateBtn').disabled = false;
            document.getElementById('updateBtn').textContent = '📝 Apply Selected Updates';
          })
          .applyContactUpdates(selected, dialogData.existingContacts);
      }

      // Done button - trigger hard refresh and close dialog
      document.getElementById('done-btn').addEventListener('click', function() {
        this.disabled = true;
        this.textContent = '⏳ Refreshing...';
        google.script.run
          .withSuccessHandler(function() {
            google.script.host.close();
          })
          .withFailureHandler(function(e) {
            alert('Refresh failed: ' + e.message);
            google.script.host.close();
          })
          .battleStationHardRefresh();
      });
    </script>
  `;

  // Show dialog
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(900)
    .setHeight(700);

  ui.showModalDialog(htmlOutput, '📇 Contact Discovery');

  // Store the last run date for this vendor
  const props = PropertiesService.getScriptProperties();
  const lastRunKey = `BS_CONTACT_DISCOVERY_${vendor.replace(/[^a-zA-Z0-9]/g, '_')}`;
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  props.setProperty(lastRunKey, today);

  Logger.log(`Contact discovery complete: ${potentialNewContacts.length} new, ${potentialUpdates.length} updates`);
}

/**
 * Add contacts to Monday.com from the Contact Discovery dialog
 * @param {Array} contacts - Array of contact objects with name, email, phone, jobTitle
 * @param {string} vendorItemId - Monday.com item ID of the vendor
 * @param {string} vendorBoardId - Monday.com board ID of the vendor (Buyers or Affiliates)
 * @returns {string} Success message
 */
function addContactsToMonday(contacts, vendorItemId, vendorBoardId) {
  const apiToken = BS_CFG.MONDAY_API_TOKEN;
  const contactsBoardId = BS_CFG.CONTACTS_BOARD_ID;

  Logger.log(`=== ADDING CONTACTS TO MONDAY.COM ===`);
  Logger.log(`Adding ${contacts.length} contact(s)`);
  Logger.log(`Vendor item: ${vendorItemId} on board ${vendorBoardId}`);

  const createdContactIds = [];
  const errors = [];

  for (const contact of contacts) {
    try {
      // Build column values JSON - Monday.com requires specific formats for each column type
      const columnValues = {};

      // Email column - Monday.com email column format
      if (contact.email) {
        columnValues['email_mkrk53z4'] = { email: contact.email, text: contact.email };
      }

      // Phone column - Monday.com phone column format
      if (contact.phone) {
        const phoneDigits = contact.phone.replace(/\D/g, '');
        if (phoneDigits.length >= 10) {
          columnValues['phone_mkrkzxq2'] = { phone: phoneDigits, countryShortName: 'US' };
        }
      }

      // Job Title column - text field
      if (contact.jobTitle) {
        columnValues['text_mkz5rj9h'] = contact.jobTitle;
      }

      const columnValuesJson = JSON.stringify(JSON.stringify(columnValues));

      // Create the contact item
      const createMutation = `
        mutation {
          create_item (
            board_id: ${contactsBoardId},
            item_name: "${contact.name.replace(/"/g, '\\"')}",
            column_values: ${columnValuesJson}
          ) {
            id
            name
          }
        }
      `;

      Logger.log(`Creating contact: ${contact.name}`);

      const createResult = mondayApiRequest_(createMutation, apiToken);

      if (createResult.errors && createResult.errors.length > 0) {
        Logger.log(`Error creating contact ${contact.name}: ${createResult.errors[0].message}`);
        errors.push(`${contact.name}: ${createResult.errors[0].message}`);
        continue;
      }

      if (createResult.data?.create_item?.id) {
        const newContactId = createResult.data.create_item.id;
        createdContactIds.push(newContactId);
        Logger.log(`Created contact ${contact.name} with ID: ${newContactId}`);
      }

    } catch (e) {
      Logger.log(`Exception creating contact ${contact.name}: ${e.message}`);
      errors.push(`${contact.name}: ${e.message}`);
    }
  }

  // Link created contacts to the vendor item
  if (createdContactIds.length > 0 && vendorItemId && vendorBoardId) {
    try {
      // Determine which contacts column to use based on vendor board
      const contactsColumnId = vendorBoardId === BS_CFG.BUYERS_BOARD_ID
        ? BS_CFG.BUYERS_CONTACTS_COLUMN
        : BS_CFG.AFFILIATES_CONTACTS_COLUMN;

      // Get existing linked contacts first
      const existingQuery = `
        query {
          items(ids: [${vendorItemId}]) {
            column_values(ids: ["${contactsColumnId}"]) {
              ... on BoardRelationValue {
                linked_item_ids
              }
            }
          }
        }
      `;

      const existingResult = mondayApiRequest_(existingQuery, apiToken);
      let existingIds = [];

      if (existingResult.data?.items?.[0]?.column_values?.[0]?.linked_item_ids) {
        existingIds = existingResult.data.items[0].column_values[0].linked_item_ids;
      }

      // Combine existing and new contact IDs
      const allContactIds = [...existingIds, ...createdContactIds];

      // Update the board relation column with all contact IDs
      const linkValue = JSON.stringify({ item_ids: allContactIds });
      const escapedLinkValue = linkValue.replace(/"/g, '\\"');

      const linkMutation = `
        mutation {
          change_column_value (
            board_id: ${vendorBoardId},
            item_id: ${vendorItemId},
            column_id: "${contactsColumnId}",
            value: "${escapedLinkValue}"
          ) { id }
        }
      `;

      Logger.log(`Linking ${createdContactIds.length} new contacts to vendor`);
      const linkResult = mondayApiRequest_(linkMutation, apiToken);

      if (linkResult.errors && linkResult.errors.length > 0) {
        Logger.log(`Error linking contacts: ${linkResult.errors[0].message}`);
        errors.push(`Linking: ${linkResult.errors[0].message}`);
      } else {
        Logger.log(`Successfully linked contacts to vendor`);
      }

    } catch (e) {
      Logger.log(`Exception linking contacts: ${e.message}`);
      errors.push(`Linking: ${e.message}`);
    }
  }

  // Build result message
  if (errors.length > 0) {
    if (createdContactIds.length > 0) {
      return `Added ${createdContactIds.length} contact(s) with ${errors.length} error(s): ${errors.join('; ')}`;
    } else {
      throw new Error(`Failed to add contacts: ${errors.join('; ')}`);
    }
  }

  return `Successfully added ${createdContactIds.length} contact(s) to Monday.com and linked to vendor!`;
}

/**
 * Apply updates to existing contacts in Monday.com
 * @param {Array} updates - Array of update objects with name, email, updateType, newValue
 * @param {Array} existingContacts - Array of existing contact objects to find contact IDs
 * @returns {string} Success message
 */
function applyContactUpdates(updates, existingContacts) {
  const apiToken = BS_CFG.MONDAY_API_TOKEN;
  const contactsBoardId = BS_CFG.CONTACTS_BOARD_ID;

  Logger.log(`=== APPLYING CONTACT UPDATES ===`);
  Logger.log(`Applying ${updates.length} update(s)`);

  // First, we need to find the contact item IDs by searching the contacts board
  // Query contacts board to get item IDs for matching emails
  const emails = updates.map(u => u.email.toLowerCase());
  const uniqueEmails = [...new Set(emails)];

  Logger.log(`Looking up contacts for emails: ${uniqueEmails.join(', ')}`);

  // Query to find contact items
  const searchQuery = `
    query {
      boards(ids: [${contactsBoardId}]) {
        items_page(limit: 500) {
          items {
            id
            name
            column_values(ids: ["email_mkrk53z4", "phone_mkrkzxq2"]) {
              id
              text
            }
          }
        }
      }
    }
  `;

  const searchResult = mondayApiRequest_(searchQuery, apiToken);
  const contactMap = new Map(); // email -> { id, name, phone }

  if (searchResult.data?.boards?.[0]?.items_page?.items) {
    for (const item of searchResult.data.boards[0].items_page.items) {
      const emailCol = item.column_values.find(c => c.id === 'email_mkrk53z4');
      const phoneCol = item.column_values.find(c => c.id === 'phone_mkrkzxq2');
      if (emailCol?.text) {
        contactMap.set(emailCol.text.toLowerCase(), {
          id: item.id,
          name: item.name,
          phone: phoneCol?.text || ''
        });
      }
    }
  }

  Logger.log(`Found ${contactMap.size} contacts in Monday.com`);

  const successCount = { phone: 0, jobTitle: 0 };
  const errors = [];

  for (const update of updates) {
    const contact = contactMap.get(update.email.toLowerCase());
    if (!contact) {
      Logger.log(`Contact not found for email: ${update.email}`);
      errors.push(`${update.name}: contact not found`);
      continue;
    }

    try {
      if (update.updateType === 'phone') {
        // Update phone column - Monday.com requires JSON format for phone columns
        const phoneDigits = update.newValue.replace(/\D/g, '');
        const phoneJson = JSON.stringify({ phone: phoneDigits, countryShortName: 'US' });

        const mutation = `
          mutation {
            change_column_value (
              board_id: ${contactsBoardId},
              item_id: ${contact.id},
              column_id: "phone_mkrkzxq2",
              value: ${JSON.stringify(phoneJson)}
            ) { id }
          }
        `;

        Logger.log(`Updating phone for ${update.name} (${contact.id}): ${phoneDigits}`);
        const result = mondayApiRequest_(mutation, apiToken);

        if (result.errors && result.errors.length > 0) {
          Logger.log(`Error updating phone: ${result.errors[0].message}`);
          errors.push(`${update.name} phone: ${result.errors[0].message}`);
        } else {
          successCount.phone++;
          Logger.log(`Successfully updated phone for ${update.name}`);
        }

      } else if (update.updateType === 'jobTitle') {
        // Update the dedicated Job Title column (text_mkz5rj9h)
        let titleOnly = update.newValue;
        if (titleOnly.includes('|')) {
          titleOnly = titleOnly.split('|').pop().trim();
        }
        // Strip any leading/trailing asterisks or special chars from signatures
        titleOnly = titleOnly.replace(/^[\*\s]+|[\*\s]+$/g, '').trim();

        const escapedTitle = titleOnly.replace(/"/g, '\\"');

        const mutation = `
          mutation {
            change_column_value (
              board_id: ${contactsBoardId},
              item_id: ${contact.id},
              column_id: "text_mkz5rj9h",
              value: "\\"${escapedTitle}\\""
            ) { id }
          }
        `;

        Logger.log(`Updating job title for ${update.name} (${contact.id}): ${titleOnly}`);
        const result = mondayApiRequest_(mutation, apiToken);

        if (result.errors && result.errors.length > 0) {
          Logger.log(`Error updating job title: ${result.errors[0].message}`);
          errors.push(`${update.name} title: ${result.errors[0].message}`);
        } else {
          successCount.jobTitle++;
          Logger.log(`Successfully updated job title for ${update.name}`);
        }
      }

    } catch (e) {
      Logger.log(`Exception updating ${update.name}: ${e.message}`);
      errors.push(`${update.name}: ${e.message}`);
    }
  }

  // Build result message
  const parts = [];
  if (successCount.phone > 0) parts.push(`${successCount.phone} phone update(s)`);
  if (successCount.jobTitle > 0) parts.push(`${successCount.jobTitle} job title update(s)`);

  if (errors.length > 0) {
    if (parts.length > 0) {
      return `Applied ${parts.join(', ')} with ${errors.length} error(s): ${errors.slice(0, 3).join('; ')}`;
    } else {
      throw new Error(`Failed to apply updates: ${errors.join('; ')}`);
    }
  }

  return `Successfully applied ${parts.join(' and ')}!`;
}

/**
 * Check if an email sender is a system/bounce message (not a real person)
 */
function isSystemOrBounceEmail_(sender) {
  if (!sender) return false;
  const lowerSender = sender.toLowerCase();

  // Common system/bounce email patterns
  const systemPatterns = [
    'mailer-daemon',
    'postmaster',
    'mail delivery subsystem',
    'noreply',
    'no-reply',
    'do-not-reply',
    'donotreply',
    'bounce',
    'auto-reply',
    'autoreply',
    'automated',
    'notification@',
    'notifications@',
    'alert@',
    'alerts@',
    'system@',
    'daemon@'
  ];

  return systemPatterns.some(pattern => lowerSender.includes(pattern));
}

/**
 * Check if a domain is a generic email provider (not company-specific)
 */
function isGenericDomain_(domain) {
  const genericDomains = [
    'gmail.com', 'yahoo.com', 'hotmail.com', 'outlook.com', 'aol.com',
    'icloud.com', 'me.com', 'mac.com', 'msn.com', 'live.com',
    'protonmail.com', 'proton.me', 'zoho.com', 'ymail.com',
    'mail.com', 'email.com', 'inbox.com', 'fastmail.com'
  ];
  return genericDomains.includes(domain.toLowerCase());
}

/**
 * Extract domain from a source/website field
 * Handles: URLs, domains with www, plain domains
 */
function extractDomainFromSource_(source) {
  if (!source) return null;

  let domain = source.trim().toLowerCase();

  // Remove protocol (http://, https://)
  domain = domain.replace(/^https?:\/\//, '');

  // Remove www.
  domain = domain.replace(/^www\./, '');

  // Remove path and query string (everything after first /)
  domain = domain.split('/')[0];

  // Remove port if present
  domain = domain.split(':')[0];

  // Basic validation: should have at least one dot and no spaces
  if (!domain.includes('.') || domain.includes(' ')) {
    return null;
  }

  return domain;
}

/**
 * Try to guess a company domain from the vendor name
 * e.g., "HomeFix" -> "homefix.com", "ABC Supply" -> "abcsupply.com"
 */
function guessVendorDomain_(vendor) {
  if (!vendor) return null;

  // Remove common suffixes like LLC, Inc, Corp, etc.
  let cleaned = vendor
    .replace(/,?\s*(LLC|Inc\.?|Corp\.?|Co\.?|Ltd\.?|Limited|Corporation|Company)$/i, '')
    .trim();

  // Remove special characters and spaces, convert to lowercase
  cleaned = cleaned
    .replace(/[^a-zA-Z0-9]/g, '')
    .toLowerCase();

  if (cleaned.length < 2) return null;

  return cleaned + '.com';
}

/**
 * Extract phone numbers from text (signature, body, etc.)
 */
function extractPhoneNumbers_(text) {
  if (!text) return [];

  // Blacklisted phone numbers (our company numbers - should not be picked up as contact info)
  const blacklistedPhones = [
    '8884004868',  // Profitise main line
    '18884004868', // With country code
  ];

  // Regex patterns for various phone formats
  const patterns = [
    /\b(\d{3}[-.\s]?\d{3}[-.\s]?\d{4})\b/g,           // 123-456-7890, 123.456.7890, 123 456 7890
    /\b\((\d{3})\)\s*(\d{3})[-.\s]?(\d{4})\b/g,       // (123) 456-7890
    /\b(\d{3})[-.\s](\d{4})\b/g,                       // 456-7890 (7 digit)
    /\+1[-.\s]?(\d{3})[-.\s]?(\d{3})[-.\s]?(\d{4})\b/g // +1 123-456-7890
  ];

  const phones = new Set();

  for (const pattern of patterns) {
    let match;
    while ((match = pattern.exec(text)) !== null) {
      // Normalize to just digits
      const digits = match[0].replace(/\D/g, '');
      if (digits.length >= 10) {
        const last10 = digits.slice(-10);

        // Skip blacklisted numbers
        if (blacklistedPhones.includes(last10) || blacklistedPhones.includes(digits)) {
          continue;
        }

        // Format nicely
        const formatted = `(${last10.slice(0,3)}) ${last10.slice(3,6)}-${last10.slice(6)}`;
        phones.add(formatted);
      }
    }
  }

  return [...phones];
}

/**
 * Extract job title from email signature
 * Looks for common job title patterns near the person's name
 */
function extractJobTitle_(text, name) {
  if (!text || !name) return '';

  // Common job title keywords - these should START the title or be key parts
  const titleKeywords = [
    'Director', 'Manager', 'President', 'VP', 'Vice President', 'CEO', 'CFO', 'COO', 'CTO', 'CMO',
    'Owner', 'Partner', 'Principal', 'Founder', 'Co-Founder',
    'Coordinator', 'Specialist', 'Analyst', 'Associate',
    'Executive', 'Administrator', 'Supervisor', 'Lead', 'Head',
    'Representative', 'Consultant', 'Advisor', 'Counsel',
    'Engineer', 'Developer', 'Architect', 'Designer',
    'Regional', 'National', 'Senior', 'Junior', 'Chief'
  ];

  // Phrases that indicate this is NOT a job title (confidentiality notices, email content, etc.)
  const excludePhrases = [
    'responsible for delivery',
    'intended recipient',
    'confidential',
    'dissemination',
    'strictly prohibited',
    'received this',
    'notify us',
    'delete it',
    'phone call',
    'let me know',
    'ready to',
    'buying from',
    'get started',
    'employee or agent',
    'click here',
    'unsubscribe',
    'view in browser'
  ];

  // Helper function to check if line is valid job title candidate
  const isValidTitleLine = (line) => {
    const lowerLine = line.toLowerCase();

    // Skip empty or too long lines (real job titles are usually under 50 chars)
    if (!line || line.length > 50 || line.length < 5) return false;

    // Skip lines starting with quote markers, bullets, or asterisks
    if (/^[>\*\-•\[]/.test(line.trim())) return false;

    // Skip lines with email addresses, URLs, or phone patterns in context
    if (line.includes('@') || line.includes('http') || line.includes('www.')) return false;

    // Skip lines containing exclusion phrases
    for (const phrase of excludePhrases) {
      if (lowerLine.includes(phrase)) return false;
    }

    // Must contain at least one title keyword
    let hasKeyword = false;
    for (const keyword of titleKeywords) {
      if (lowerLine.includes(keyword.toLowerCase())) {
        hasKeyword = true;
        break;
      }
    }
    if (!hasKeyword) return false;

    // Title should be mostly letters and spaces (at least 85% alphanumeric)
    const alphaRatio = (line.match(/[a-zA-Z\s,&]/g) || []).length / line.length;
    if (alphaRatio < 0.85) return false;

    return true;
  };

  const firstName = name.split(/\s+/)[0];
  const lines = text.split(/[\n\r]+/);

  // First pass: look for title right after the person's name
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();

    // If this line contains the person's name
    if (line.toLowerCase().includes(firstName.toLowerCase()) && line.length < 100) {
      // Check if there's a pipe separator with title after name (e.g., "John Smith | Director of Sales")
      if (line.includes('|')) {
        const parts = line.split('|');
        for (let j = 1; j < parts.length; j++) {
          const part = parts[j].trim();
          if (isValidTitleLine(part)) {
            return part;
          }
        }
      }

      // Check the next line for a title
      if (i + 1 < lines.length) {
        const nextLine = lines[i + 1].trim();
        if (isValidTitleLine(nextLine)) {
          return nextLine;
        }
      }
    }
  }

  // Second pass: look for standalone title lines near the top (signature block area)
  for (let i = 0; i < Math.min(lines.length, 15); i++) {
    const line = lines[i].trim();
    if (isValidTitleLine(line)) {
      return line;
    }
  }

  return '';
}

/**
 * Escape HTML special characters
 */
function escapeHtml_(text) {
  if (!text) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/************************************************************
 * TASK STATUS ANALYSIS FROM EMAILS
 * Uses Claude to analyze emails and suggest task status updates
 ************************************************************/

/**
 * Analyze emails to suggest task status updates
 * Uses AI to determine if tasks should be updated based on email content
 */
function analyzeTasksFromEmails() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);

  if (!listSh) {
    ui.alert('Vendor list not found.');
    return;
  }

  const claudeApiKey = BS_CFG.CLAUDE_API_KEY;
  if (!claudeApiKey || claudeApiKey === 'YOUR_ANTHROPIC_API_KEY_HERE') {
    ui.alert('Please set your Anthropic API key in BS_CFG.CLAUDE_API_KEY');
    return;
  }

  const currentIndex = getCurrentVendorIndex_();
  const listRow = currentIndex + 1;
  const vendor = listSh.getRange(listRow, 1).getValue();
  const source = listSh.getRange(listRow, 2).getValue() || '';

  ss.toast('Fetching emails and tasks...', '🤖 Task Analysis', 3);

  // Get vendor info including status
  const contactData = getVendorContacts_(vendor, listRow);
  const vendorStatus = contactData.liveStatus || 'Unknown';
  const hasContacts = (contactData.contacts || []).length > 0;
  const hasPhonexaLink = !!contactData.phonexaLink;

  // Get ALL emails from vendor's Gmail sublabel (not filtered search)
  // This gives full context including processed/dealt-with emails
  const emails = getAllEmailsFromVendorLabel_(listRow, 50);
  if (!emails || emails.length === 0) {
    ui.alert('No emails found in vendor label.');
    return;
  }

  // Get tasks for this vendor
  const tasks = getTasksForVendor_(vendor, listRow);
  const openTasks = tasks.filter(t => !t.isDone);

  if (openTasks.length === 0) {
    ui.alert('No open tasks found for this vendor.');
    return;
  }

  ss.toast('Analyzing with Claude AI...', '🤖 Task Analysis', 5);

  // Get task analysis settings from Settings sheet
  const taskSettings = getTaskAnalysisSettings_();

  // Build email summaries with content snippets (most recent first, limit to 25 for full context)
  // Also build a map of email index to thread ID for linking
  const emailThreadMap = {};
  const emailSummaries = emails.slice(0, 25).map((e, i) => {
    // Handle labels - could be array or string
    const labelsStr = Array.isArray(e.labels) ? e.labels.join(', ') : (e.labels || '');
    // Include snippet for context (truncate if too long)
    const snippet = (e.snippet || '').substring(0, 300);
    const msgCount = e.messageCount ? ` [${e.messageCount} msgs in thread]` : '';
    // Store thread ID for linking
    emailThreadMap[i + 1] = e.threadId;
    return `EMAIL ${i + 1} (${e.date})${msgCount}:
Subject: ${e.subject}
From: ${e.from || 'Unknown'}
Labels: ${labelsStr}
Preview: ${snippet}
---`;
  }).join('\n\n');

  // Build task list with custom instructions from settings
  const taskList = openTasks.map((t, i) => {
    // Extract base task name (without vendor suffix) for matching settings
    const baseTaskName = t.subject.replace(/ - [^-]+$/, '').trim().toLowerCase();
    const customInstruction = taskSettings.taskInstructions[baseTaskName] || '';
    const instructionNote = customInstruction ? `\n   [LOOK FOR: ${customInstruction}]` : '';
    return `${i + 1}. "${t.subject}" - Status: ${t.status}${t.taskDate ? ` (Date: ${t.taskDate})` : ''}${instructionNote}`;
  }).join('\n');

  // Build general instructions section
  let customGuidance = '';
  if (taskSettings.generalInstructions) {
    customGuidance = `\nCUSTOM GUIDANCE FROM USER:\n${taskSettings.generalInstructions}\n`;
  }

  // Build the prompt
  const prompt = `You are analyzing a vendor's email thread and their monday.com onboarding tasks to suggest status updates.

VENDOR: ${vendor}
VENDOR STATUS: ${vendorStatus}
HAS CONTACTS IN MONDAY.COM: ${hasContacts ? 'Yes' : 'No'}
HAS PHONEXA LINK: ${hasPhonexaLink ? 'Yes' : 'No'}

CURRENT OPEN TASKS:
${taskList}

POSSIBLE TASK STATUSES:
- "Waiting on Profitise" - We need to do something
- "Waiting on Client" - We're waiting for them to do something
- "Waiting on Phonexa" - We're waiting for Phonexa (our platform team) to do something
- "Done" - Task is complete

RECENT EMAILS (most recent first):
${emailSummaries}
${customGuidance}
IMPORTANT CONTEXT:
- Pay close attention to any [LOOK FOR: ...] notes on tasks - these are custom instructions from the user
- Onboarding tasks are typically done in a linear sequence
- If later tasks are being worked on, earlier tasks may already be complete
- Look for evidence in emails that suggests task status changes:
  * If we sent them something and are waiting for response → "Waiting on Client"
  * If they requested something and we haven't done it → "Waiting on Profitise"
  * If we mention Phonexa needs to do something → "Waiting on Phonexa"
  * If something is confirmed done in emails → "Done"

Analyze the emails and provide your recommendations in this EXACT format:

TASK UPDATES:
[For each task that should change status, use this format:]
- "Task Name" → New Status
  Reason: [Brief explanation based on email evidence]

NO CHANGE NEEDED:
[List any tasks where current status seems correct]

SUMMARY:
[2-3 sentences summarizing what's happening with this vendor based on emails]`;

  // Call Claude API
  const response = callClaudeAPI_(prompt, claudeApiKey);

  if (response.error) {
    ui.alert(`Claude API Error: ${response.error}`);
    return;
  }

  Logger.log(`=== TASK ANALYSIS ===`);
  Logger.log(`Vendor: ${vendor}`);
  Logger.log(`Open tasks: ${openTasks.length}`);
  Logger.log(`Emails analyzed: ${Math.min(emails.length, 15)}`);
  Logger.log(`Claude response: ${response.content}`);

  // Parse Claude's response to extract suggested updates
  let suggestedUpdates = parseTaskSuggestions_(response.content, openTasks);
  Logger.log(`Parsed ${suggestedUpdates.length} suggested updates (before filtering)`);

  // Filter out suggestions where status isn't actually changing
  suggestedUpdates = suggestedUpdates.filter(u => u.currentStatus !== u.newStatus);
  Logger.log(`After filtering same-status: ${suggestedUpdates.length} updates`);

  // Store task data for the dialog to use
  const taskDataForDialog = openTasks.map(t => ({
    itemId: t.itemId,
    statusColumnId: t.statusColumnId,
    subject: t.subject,
    currentStatus: t.status
  }));
  PropertiesService.getUserProperties().setProperty('taskAnalysisData', JSON.stringify(taskDataForDialog));

  // Build suggested updates HTML with Apply buttons
  const allStatuses = ['Waiting on Profitise', 'Waiting on Client', 'Waiting on Phonexa', 'Done'];

  // Helper to linkify "EMAIL X" references in reason text
  const linkifyEmailRefs = (text) => {
    return escapeHtml_(text).replace(/EMAIL\s*(\d+)/gi, (match, num) => {
      const threadId = emailThreadMap[parseInt(num)];
      if (threadId) {
        return `<a href="https://mail.google.com/mail/u/0/#inbox/${threadId}" target="_blank" style="color: #1a73e8;">EMAIL ${num}</a>`;
      }
      return match;
    });
  };

  let updatesHtml = '';
  if (suggestedUpdates.length > 0) {
    updatesHtml = '<h3 style="color: #1a73e8; margin-top: 15px;">📋 SUGGESTED UPDATES:</h3>';
    for (const update of suggestedUpdates) {
      const taskLink = `https://profitise-company.monday.com/boards/${BS_CFG.TASKS_BOARD_ID}?term=${encodeURIComponent(update.taskName)}`;
      const statusBg = getStatusColor_(update.newStatus);

      // Build alternative status buttons (exclude current and suggested)
      const altStatuses = allStatuses.filter(s =>
        s !== update.currentStatus && s !== update.newStatus
      );
      const escapedTaskName = escapeHtml_(update.taskName).replace(/'/g, "\\'");
      const altButtonsHtml = altStatuses.map(s => {
        const bg = getStatusColor_(s);
        return `<button class="alt-btn" style="background: ${bg};" onclick="applyOverride('${update.itemId}', '${update.statusColumnId}', '${s}', '${escapedTaskName}', this)">${s}</button>`;
      }).join('');

      updatesHtml += `
        <div class="update-row" id="update-${update.itemId}">
          <div class="update-task">
            <a href="${taskLink}" target="_blank" class="task-link">🔗 ${escapeHtml_(update.taskName)}</a>
          </div>
          <div class="update-change">
            <span class="old-status">${escapeHtml_(update.currentStatus)}</span> →
            <span style="background: ${statusBg}; padding: 2px 6px; border-radius: 3px;">${escapeHtml_(update.newStatus)}</span>
          </div>
          <div class="update-reason"><em>Reason:</em> ${linkifyEmailRefs(update.reason)}</div>
          <div class="update-actions">
            <button class="apply-btn" onclick="applyUpdate('${update.itemId}', '${update.statusColumnId}', '${escapeHtml_(update.newStatus)}', this)">✓ Apply</button>
            <button class="skip-btn" onclick="skipUpdate('${escapedTaskName}', '${escapeHtml_(update.currentStatus)}', this)">✗ Skip</button>
            ${altButtonsHtml}
          </div>
        </div>
      `;
    }
  } else {
    updatesHtml = '<h3 style="color: #666; margin-top: 15px;">✓ No status changes suggested</h3><p style="color: #888; font-size: 13px;">All tasks appear to have appropriate statuses based on the email history.</p>';
  }


  const html = `
    <style>
      body { font-family: Arial, sans-serif; font-size: 13px; padding: 15px; line-height: 1.6; }
      h2 { color: #1a73e8; margin-bottom: 10px; }
      h3 { margin-top: 20px; margin-bottom: 10px; }
      .summary-box { background: #f8f9fa; padding: 12px; border-radius: 5px; margin-bottom: 15px; }
      .update-row { background: #fff; border: 1px solid #ddd; padding: 12px; margin-bottom: 10px; border-radius: 5px; }
      .update-row.applied { background: #e8f5e9; border-color: #4caf50; }
      .update-row.skipped { background: #f5f5f5; opacity: 0.6; }
      .update-task { font-weight: bold; margin-bottom: 5px; }
      .update-change { margin-bottom: 5px; }
      .update-reason { color: #666; font-size: 12px; margin-bottom: 8px; }
      .update-actions { display: flex; gap: 8px; }
      .apply-btn { background: #4caf50; color: white; border: none; padding: 6px 12px; border-radius: 4px; cursor: pointer; }
      .apply-btn:hover { background: #45a049; }
      .apply-btn:disabled { background: #ccc; cursor: not-allowed; }
      .skip-btn { background: #f5f5f5; color: #666; border: 1px solid #ddd; padding: 6px 12px; border-radius: 4px; cursor: pointer; }
      .skip-btn:hover { background: #eee; }
      .alt-btn { color: #333; border: 1px solid #ccc; padding: 4px 8px; border-radius: 4px; cursor: pointer; font-size: 11px; }
      .alt-btn:hover { opacity: 0.8; }
      .alt-btn:disabled { opacity: 0.5; cursor: not-allowed; }
      .old-status { color: #666; text-decoration: line-through; }
      .task-link { color: #1a73e8; text-decoration: none; }
      .task-link:hover { text-decoration: underline; }
      .task-list { background: #fafafa; padding: 10px; margin-top: 15px; border-radius: 5px; }
      .task-item { padding: 5px 0; border-bottom: 1px solid #eee; }
      .summary-section { background: #f0f7ff; padding: 12px; border-radius: 5px; margin-top: 15px; }
      .status-msg { padding: 8px; margin-top: 10px; border-radius: 4px; display: none; }
      .status-msg.success { background: #e8f5e9; color: #2e7d32; display: block; }
      .status-msg.error { background: #ffebee; color: #c62828; display: block; }
    </style>

    <h2>🤖 Task Status Analysis for ${escapeHtml_(vendor)}</h2>

    <div class="summary-box">
      <strong>Analyzed:</strong> ${Math.min(emails.length, 25)} emails from label (${emails.length} total), ${openTasks.length} open tasks
    </div>

    <div id="status-message" class="status-msg"></div>

    ${updatesHtml}

    <div style="margin-top: 20px; padding-top: 15px; border-top: 1px solid #ddd; text-align: center;">
      <button id="done-btn" style="background: #1a73e8; color: white; border: none; padding: 10px 30px; border-radius: 5px; font-size: 14px; cursor: pointer;">
        ✓ Done - Refresh View
      </button>
    </div>

    <script>
      function applyUpdate(itemId, statusColumnId, newStatus, btn) {
        btn.disabled = true;
        btn.textContent = '⏳ Updating...';

        google.script.run
          .withSuccessHandler(function(result) {
            if (result.success) {
              btn.textContent = '✓ Applied';
              btn.parentElement.parentElement.classList.add('applied');
              showStatus('Updated: ' + newStatus, 'success');
            } else {
              btn.textContent = '✗ Failed';
              btn.disabled = false;
              showStatus('Error: ' + result.error, 'error');
            }
          })
          .withFailureHandler(function(error) {
            btn.textContent = '✗ Failed';
            btn.disabled = false;
            showStatus('Error: ' + error.message, 'error');
          })
          .updateTaskStatus(itemId, statusColumnId, newStatus);
      }

      function skipUpdate(taskName, currentStatus, btn) {
        var reason = prompt('Why are you skipping this suggestion?\\n\\nThis will be saved to Settings for future reference.');
        if (reason === null) return; // Cancelled

        // Save skip reason to Settings
        if (reason && reason.trim() !== '') {
          google.script.run
            .withFailureHandler(function(e) { console.log('Failed to save skip reason: ' + e); })
            .saveStatusOverride(taskName, 'SKIP (keep ' + currentStatus + ')', reason);
        }

        btn.parentElement.parentElement.classList.add('skipped');
        btn.parentElement.innerHTML = '<em>Skipped</em>';
      }

      function applyOverride(itemId, statusColumnId, newStatus, taskName, btn) {
        var comment = prompt('Why is "' + newStatus + '" more appropriate?\\n\\nThis will be saved to Settings for future reference.');
        if (comment === null) return; // Cancelled

        btn.disabled = true;
        btn.textContent = '⏳...';

        // First save the override comment to Settings
        if (comment && comment.trim() !== '') {
          google.script.run
            .withFailureHandler(function(e) { console.log('Failed to save override: ' + e); })
            .saveStatusOverride(taskName, newStatus, comment);
        }

        // Then apply the status update
        google.script.run
          .withSuccessHandler(function(result) {
            if (result.success) {
              btn.textContent = '✓ Applied';
              btn.parentElement.parentElement.classList.add('applied');
              showStatus('Override applied: ' + newStatus, 'success');
            } else {
              btn.textContent = '✗ Failed';
              btn.disabled = false;
              showStatus('Error: ' + result.error, 'error');
            }
          })
          .withFailureHandler(function(error) {
            btn.textContent = '✗ Failed';
            btn.disabled = false;
            showStatus('Error: ' + error.message, 'error');
          })
          .updateTaskStatus(itemId, statusColumnId, newStatus);
      }

      function showStatus(message, type) {
        var el = document.getElementById('status-message');
        el.textContent = message;
        el.className = 'status-msg ' + type;
        setTimeout(function() { el.className = 'status-msg'; }, 5000);
      }

      // Done button - trigger hard refresh and close dialog
      document.getElementById('done-btn').addEventListener('click', function() {
        this.disabled = true;
        this.textContent = '⏳ Refreshing...';
        google.script.run
          .withSuccessHandler(function() {
            google.script.host.close();
          })
          .withFailureHandler(function(e) {
            alert('Refresh failed: ' + e.message);
            google.script.host.close();
          })
          .battleStationHardRefresh();
      });
    </script>
  `;

  // Show dialog
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(850)
    .setHeight(750);

  ui.showModalDialog(htmlOutput, '🤖 Task Status Analysis');
}

/**
 * Parse Claude's response to extract task update suggestions
 */
function parseTaskSuggestions_(claudeResponse, openTasks) {
  const suggestions = [];

  // Match patterns like: - "Task Name" → New Status
  const updatePattern = /-\s*"([^"]+)"\s*→\s*(Waiting on Client|Waiting on Profitise|Waiting on Phonexa|Done)/gi;

  // Also capture the reason that follows
  const lines = claudeResponse.split('\n');
  let currentUpdate = null;

  for (const line of lines) {
    const match = line.match(/-\s*"([^"]+)"\s*→\s*(Waiting on Client|Waiting on Profitise|Waiting on Phonexa|Done)/i);
    if (match) {
      // Find the matching task to get itemId and statusColumnId
      const taskName = match[1];
      const newStatus = match[2];

      const matchingTask = openTasks.find(t =>
        t.subject.toLowerCase().includes(taskName.toLowerCase()) ||
        taskName.toLowerCase().includes(t.subject.replace(/ - [^-]+$/, '').toLowerCase())
      );

      if (matchingTask) {
        currentUpdate = {
          taskName: matchingTask.subject,
          itemId: matchingTask.itemId,
          statusColumnId: matchingTask.statusColumnId,
          currentStatus: matchingTask.status,
          newStatus: newStatus,
          reason: ''
        };
        suggestions.push(currentUpdate);
      }
    } else if (currentUpdate && line.toLowerCase().includes('reason:')) {
      // Extract reason
      currentUpdate.reason = line.replace(/.*reason:\s*/i, '').trim();
    }
  }

  return suggestions;
}

/**
 * Get background color for a status
 */
function getStatusColor_(status) {
  const statusLower = (status || '').toLowerCase();
  if (statusLower.includes('client')) return '#fff2cc';
  if (statusLower.includes('profitise')) return '#e3f2fd';
  if (statusLower.includes('phonexa')) return '#ffcdd2';
  if (statusLower === 'done') return '#c8e6c9';
  return '#f5f5f5';
}

/**
 * Update a task's status in Monday.com (called from dialog)
 */
function updateTaskStatus(itemId, statusColumnId, newStatus) {
  const apiToken = BS_CFG.MONDAY_API_TOKEN;
  const boardId = BS_CFG.TASKS_BOARD_ID;

  Logger.log(`Updating task ${itemId} status to: ${newStatus}`);

  // Monday.com status columns use label values, need to format as JSON
  const statusValue = JSON.stringify({ label: newStatus });
  const escapedValue = statusValue.replace(/"/g, '\\"');

  const mutation = `
    mutation {
      change_column_value (
        board_id: ${boardId},
        item_id: ${itemId},
        column_id: "${statusColumnId}",
        value: "${escapedValue}"
      ) { id }
    }
  `;

  try {
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': apiToken },
      payload: JSON.stringify({ query: mutation }),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch('https://api.monday.com/v2', options);
    const result = JSON.parse(response.getContentText());

    Logger.log(`Monday.com response: ${JSON.stringify(result)}`);

    if (result.errors && result.errors.length > 0) {
      return { success: false, error: result.errors[0].message };
    }

    if (result.data?.change_column_value?.id) {
      return { success: true };
    }

    return { success: false, error: 'Unexpected API response' };

  } catch (e) {
    Logger.log(`Error updating task: ${e.message}`);
    return { success: false, error: e.message };
  }
}

/**
 * Save a status override comment to Settings sheet columns S and T (Task Analysis Settings)
 * Appends the feedback to the existing "What to Look For" value for the matching task
 */
function saveStatusOverride(taskName, chosenStatus, comment) {
  const ss = SpreadsheetApp.getActive();
  const settingsSh = ss.getSheetByName('Settings');

  if (!settingsSh) {
    Logger.log('Settings sheet not found, cannot save override');
    return;
  }

  // Column S = 19 (Task Name), Column T = 20 (What to Look For)
  const colS = 19;
  const colT = 20;

  const data = settingsSh.getDataRange().getValues();
  const taskNameLower = taskName.toLowerCase();
  let matchedRow = -1;
  let inSection = false;

  // Find the task row in "Task Analysis Settings" section
  for (let i = 0; i < data.length; i++) {
    const cellS = String(data[i][colS - 1] || '').trim();

    // Check for section header
    if (cellS.toLowerCase() === 'task analysis settings') {
      inSection = true;
      continue;
    }

    if (!inSection) continue;

    // Skip column header row
    if (cellS.toLowerCase() === 'task name') continue;

    // Exit section on empty row
    if (cellS === '' && String(data[i][colT - 1] || '').trim() === '') {
      break;
    }

    // Check if this row matches our task (partial match)
    if (cellS !== '' && taskNameLower.includes(cellS.toLowerCase())) {
      matchedRow = i + 1; // 1-indexed
      break;
    }
  }

  if (matchedRow === -1) {
    Logger.log(`Task not found in Settings: ${taskName}`);
    return;
  }

  // Append feedback to existing value in column T
  const existingValue = String(settingsSh.getRange(matchedRow, colT).getValue() || '');
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd');
  const feedback = `[${timestamp}: ${chosenStatus}] ${comment}`;

  const newValue = existingValue
    ? `${existingValue}; ${feedback}`
    : feedback;

  settingsSh.getRange(matchedRow, colT).setValue(newValue);

  Logger.log(`Appended override feedback to row ${matchedRow}: "${feedback}"`);
}
