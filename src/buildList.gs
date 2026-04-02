/************************************************************
 * BUILD LIST - Vendor list builder with Gmail and monday.com integration
 *
 * Last Updated: 2026-01-07 1:15PM PST
 ************************************************************/

/***** CONFIG *****/
const SHEET_BUYERS_L1M       = 'Buyers L1M';
const SHEET_BUYERS_L6M       = 'Buyers L6M';
const SHEET_AFFILIATES_L1M   = 'Affiliates L1M';
const SHEET_AFFILIATES_L6M   = 'Affiliates L6M';
const SHEET_MON_BUYERS       = 'buyers monday.com';
const SHEET_MON_AFFILIATES   = 'affiliates monday.com';

const SHEET_OUT              = 'List';
const SHEET_SETTINGS         = 'Settings';

// Header candidates to auto-detect vendor and TTL
const VENDOR_HEADER_CANDIDATES = ['Buyer', 'Affiliate', 'Publisher', 'Name', 'Vendor'];
const TTL_HEADER_CANDIDATES    = ['TTL , USD', 'TTL, USD', 'TTL USD'];

// monday.com special columns (0-based indexes)
const ALIAS_COL_INDEX = 15;           // Column P (1-based 16) -> 0-based 15

// Header names to search for (will auto-detect column index)
const STATUS_HEADER_CANDIDATES = ['Group', 'Status', 'Group Title'];
const NOTES_HEADER_CANDIDATES = ['Notes', 'Note', 'Comments'];

// Status priority used when TTL = 0
const STATUS_PRIORITY = [
  'live',
  'onboarding',
  'paused',
  'preonboarding',
  'early talks',
  'top 500 remodelers',
  'other',
  'dead'
];
const STATUS_RANK = STATUS_PRIORITY.reduce((m, s, i) => (m[s] = i, m), {});

// Skip reason tracking (for debugging)
const SKIP_REASONS = [];

/**
 * NOTE: Menu is defined in BattleStation.gs onOpen() function
 * to avoid collision with multiple onOpen() declarations
 */

/**
 * Build List with Gmail links and monday.com notes
 * Generates columns: Vendor | TTL, USD | Source | Status | Notes | Gmail Link | no snoozing | processed? | Tranche
 *
 * SORT ORDER:
 *   Base sort: Group DESC → Buyer>Affiliate → Rank DESC → Opportunity Size DESC → Buyer Type ASC → Creation log ASC → Name ASC
 *   Then: vendors with emails in last 7 days move to top (same sort within)
 *   Then: vendors in (inbox OR 00.received) AND -is:snoozed move to top (same sort)
 *   Then: vendors in label:inbox move to top (same sort)
 *   Then: if Profitise Internal has 00.received emails, it goes to very top
 */
function buildListWithGmailAndNotes() {
  const ss = SpreadsheetApp.getActive();

  console.log('=== BUILD LIST WITH GMAIL & NOTES START ===');

  // Refresh batch caches since buildList usually means user is about to review vendors
  console.log('Refreshing batch caches for vendor review...');
  refreshBatchCaches_();

  const blacklist = readBlacklist_(ss);

  // Read vendors from monday.com sheets only (metric sheets commented out)
  const shMonB = ss.getSheetByName(SHEET_MON_BUYERS);
  const shMonA = ss.getSheetByName(SHEET_MON_AFFILIATES);

  /* // TODO: Re-enable metric sheets when actively used
  const shBuyL1 = mustGetSheet_(ss, SHEET_BUYERS_L1M);
  const shBuyL6 = mustGetSheet_(ss, SHEET_BUYERS_L6M);
  const shAffL1 = mustGetSheet_(ss, SHEET_AFFILIATES_L1M);
  const shAffL6 = mustGetSheet_(ss, SHEET_AFFILIATES_L6M);
  const buyersL1M = readMetricSheet_(shBuyL1, 'Buyer', 'Buyers L1M', blacklist);
  const buyersL6M = readMetricSheet_(shBuyL6, 'Buyer', 'Buyers L6M', blacklist);
  const affL1M = readMetricSheet_(shAffL1, 'Affiliate', 'Affiliates L1M', blacklist);
  const affL6M = readMetricSheet_(shAffL6, 'Affiliate', 'Affiliates L6M', blacklist);
  */

  const existingSet = new Set();
  const buyersMon = shMonB ? readMondaySheet_(shMonB, 'Buyer', 'buyers monday.com', blacklist, existingSet) : [];
  const affMon = shMonA ? readMondaySheet_(shMonA, 'Affiliate', 'affiliates monday.com', blacklist, existingSet) : [];

  console.log('Data loaded:', { buyersMon: buyersMon.length, affMon: affMon.length });

  // Combine all vendors
  const all = [...buyersMon, ...affMon];

  console.log('Total vendors:', all.length);

  // Read extra sort columns from monday.com sheets
  const sortData = readMondaySortColumns_(shMonB, shMonA);

  // Attach sort fields to each vendor
  for (const r of all) {
    const key = r.name.toLowerCase();
    const extra = sortData.get(key) || {};
    r.groupRank = getStatusRank_(r.status);
    r.rank = extra.rank || 0;
    r.opportunitySize = extra.opportunitySize || 0;
    r.buyerType = extra.buyerType || '';
    r.creationDate = extra.creationDate || new Date('2099-01-01');
    r.isBuyer = (r.type || '').toLowerCase().startsWith('buyer') ? 0 : 1; // 0=buyer, 1=affiliate
  }

  // BASE SORT: Group ASC (Live first) → Buyer>Affiliate → Rank DESC → Opportunity Size DESC → Buyer Type ASC → Creation log ASC → Name ASC
  all.sort((a, b) => {
    // Group rank ASC (lower = higher priority: Live=0, Onboarding=1, etc.)
    if (a.groupRank !== b.groupRank) return a.groupRank - b.groupRank;
    // Buyer before Affiliate
    if (a.isBuyer !== b.isBuyer) return a.isBuyer - b.isBuyer;
    // Rank DESC (higher rank first)
    if (b.rank !== a.rank) return b.rank - a.rank;
    // Opportunity Size DESC (higher first)
    if (b.opportunitySize !== a.opportunitySize) return b.opportunitySize - a.opportunitySize;
    // Buyer Type ASC
    if (a.buyerType !== b.buyerType) return a.buyerType.localeCompare(b.buyerType);
    // Creation log ASC (oldest first)
    if (a.creationDate.getTime() !== b.creationDate.getTime()) return a.creationDate - b.creationDate;
    // Name ASC
    return a.name.localeCompare(b.name);
  });

  // Set default tranche
  for (const r of all) {
    r.tranche = 'Normal';
  }

  // LAYER 1: Move vendors with emails in last 7 days to top, sorted by oldest thread date ASC
  console.log('Detecting recent email activity...');
  const { recentSet, recentDates } = getRecentEmailVendors_(all);
  console.log('Recent email vendors (7 days):', recentSet.size);

  const recentVendors = [];
  const nonRecentVendors = [];
  for (const r of all) {
    if (recentSet.has(r.name.toLowerCase())) {
      r.tranche = '🔥 Recent';
      r.threadDate = recentDates[r.name.toLowerCase()] || new Date();
      recentVendors.push(r);
    } else {
      nonRecentVendors.push(r);
    }
  }
  // Sort recent vendors by oldest thread date first
  recentVendors.sort((a, b) => a.threadDate - b.threadDate);

  // LAYER 2: Move vendors in (inbox OR 00.received) AND -is:snoozed to top, sorted by oldest thread date ASC
  console.log('Detecting actionable emails...');
  const { actionableSet, actionableDates, inboxOnlySet, inboxDates, profitiseInternal } = getActionableVendors_(all);
  console.log('Actionable vendors:', actionableSet.size);
  console.log('Inbox-only vendors:', inboxOnlySet.size);
  console.log('Profitise Internal actionable:', profitiseInternal);

  const actionableVendors = [];
  const restAfterActionable = [];
  for (const r of [...recentVendors, ...nonRecentVendors]) {
    if (actionableSet.has(r.name.toLowerCase())) {
      r.tranche = '📥 Actionable';
      r.threadDate = actionableDates[r.name.toLowerCase()] || r.threadDate || new Date();
      actionableVendors.push(r);
    } else {
      restAfterActionable.push(r);
    }
  }
  // Sort actionable by oldest thread date first
  actionableVendors.sort((a, b) => a.threadDate - b.threadDate);

  // LAYER 3: Move vendors in label:inbox to top of actionable, sorted by oldest thread date ASC
  const inboxVendors = [];
  const actionableNotInbox = [];
  for (const r of actionableVendors) {
    if (inboxOnlySet.has(r.name.toLowerCase())) {
      r.tranche = '📥 Inbox';
      r.threadDate = inboxDates[r.name.toLowerCase()] || r.threadDate || new Date();
      inboxVendors.push(r);
    } else {
      actionableNotInbox.push(r);
    }
  }
  // Sort inbox by oldest thread date first
  inboxVendors.sort((a, b) => a.threadDate - b.threadDate);

  // LAYER 4: If Profitise Internal has 00.received emails, put at very top
  let profitiseInternalVendor = null;
  if (profitiseInternal) {
    // Check if it's already in one of the lists, or create a placeholder
    const allLists = [inboxVendors, actionableNotInbox, restAfterActionable];
    for (const list of allLists) {
      for (let i = 0; i < list.length; i++) {
        if (list[i].name.toLowerCase().includes('profitise internal')) {
          profitiseInternalVendor = list.splice(i, 1)[0];
          profitiseInternalVendor.tranche = '🔴 Internal';
          break;
        }
      }
      if (profitiseInternalVendor) break;
    }
  }

  // Assemble final list
  const finalList = [];
  if (profitiseInternalVendor) finalList.push(profitiseInternalVendor);
  finalList.push(...inboxVendors, ...actionableNotInbox, ...restAfterActionable);

  console.log('Total vendors for output:', finalList.length);

  // Write to List sheet
  const shOut = ensureSheet_(ss, SHEET_OUT);

  // Clear all content and formatting first (removes green/yellow row highlighting)
  const lastRow = shOut.getMaxRows();
  const lastCol = shOut.getMaxColumns();
  if (lastRow > 0 && lastCol > 0) {
    shOut.getRange(1, 1, lastRow, lastCol).clear();
  }

  // Write headers
  shOut.getRange(1, 1, 1, 9).setValues([[
    'Vendor', 'TTL , USD', 'Source', 'Status', 'Notes',
    'Gmail Link', 'no snoozing', 'processed?', 'Tranche'
  ]]).setFontWeight('bold').setBackground('#e8f0fe');

  if (finalList.length > 0) {
    // Write data (columns A-E)
    const data = finalList.map(r => [
      r.name,
      r.ttl || 0,
      r.source,
      r.status || '',
      r.notes || ''
    ]);

    shOut.getRange(2, 1, data.length, 5).setValues(data);
    shOut.getRange(2, 2, data.length, 1).setNumberFormat('$#,##0.00;($#,##0.00)');

    // Read Gmail sublabel mappings from Settings sheet
    const gmailSublabelMap = readGmailSublabelMap_(ss);

    // Add Gmail links (plain URLs) and processed checkbox in columns F, G, H
    const gmailData = finalList.map(v => {
      // Check if we have a custom sublabel mapping, otherwise generate from name
      const vendorKey = v.name.toLowerCase();
      let vendorSlug;

      if (gmailSublabelMap.has(vendorKey)) {
        vendorSlug = gmailSublabelMap.get(vendorKey);
      } else {
        // Fallback: generate slug from vendor name (preserve hyphens and periods)
        vendorSlug = v.name.toLowerCase()
          .replace(/[^a-z0-9.-]+/g, '-')
          .replace(/^-+|-+$/g, '');
      }

      const gmailAll = `https://mail.google.com/mail/u/0/#search/label%3A00.received+AND+label%3Azzzvendors-${vendorSlug}+-label%3A03.noInbox`;
      const gmailNoSnooze = `https://mail.google.com/mail/u/0/#search/label%3A00.received+AND+label%3Azzzvendors-${vendorSlug}+-is%3Asnoozed+-label%3A03.noInbox`;

      return [gmailAll, gmailNoSnooze, false, v.tranche || 'Normal'];
    });

    shOut.getRange(2, 6, gmailData.length, 4).setValues(gmailData);
  }

  // Auto-resize columns
  shOut.autoResizeColumns(1, 5);
  shOut.setColumnWidth(6, 100);
  shOut.setColumnWidth(7, 100);
  shOut.setColumnWidth(8, 100);
  shOut.autoResizeColumn(9);  // Tranche column

  console.log('=== BUILD LIST WITH GMAIL & NOTES END ===');

  const trancheCounts = {};
  for (const r of finalList) {
    trancheCounts[r.tranche] = (trancheCounts[r.tranche] || 0) + 1;
  }
  trancheCounts['Total'] = finalList.length;

  console.log('Final counts:', trancheCounts);

  ss.toast(
    Object.entries(trancheCounts).map(([k,v]) => `${k}: ${v}`).join(' • '),
    '✅ List Built',
    8
  );
}


/** ========== SORT COLUMN READER ========== **/

/**
 * Read extra sort columns (Rank, Opportunity Size, Buyer Type, Creation log)
 * from the monday.com sheets. Returns a Map of lowercase vendor name -> { rank, opportunitySize, buyerType, creationDate }
 */
function readMondaySortColumns_(shMonB, shMonA) {
  const map = new Map();

  const readSheet = (sh) => {
    if (!sh) return;
    const allValues = sh.getDataRange().getValues();
    if (allValues.length < 2) return;

    const headers = allValues[0].map(h => String(h || '').trim().toLowerCase());

    // Find column indices
    const rankIdx = headers.findIndex(h => h === 'rank');
    const oppSizeIdx = headers.findIndex(h => h === 'opportunity size');
    const buyerTypeIdx = headers.findIndex(h => h === 'buyer type');
    const creationIdx = headers.findIndex(h => h === 'creation log' || h === 'creation_log');

    console.log(`[Sort Columns] Sheet: ${sh.getName()}, Rank: ${rankIdx}, OppSize: ${oppSizeIdx}, BuyerType: ${buyerTypeIdx}, Creation: ${creationIdx}`);

    for (let i = 1; i < allValues.length; i++) {
      const row = allValues[i];
      const name = normalizeName_(row[0]);
      if (!name) continue;
      const key = name.toLowerCase();

      const existing = map.get(key) || {};
      map.set(key, {
        rank: (rankIdx >= 0 ? toNumber_(row[rankIdx]) : 0) || existing.rank || 0,
        opportunitySize: (oppSizeIdx >= 0 ? toNumber_(row[oppSizeIdx]) : 0) || existing.opportunitySize || 0,
        buyerType: (buyerTypeIdx >= 0 ? String(row[buyerTypeIdx] || '').trim() : '') || existing.buyerType || '',
        creationDate: parseCreationDate_(creationIdx >= 0 ? row[creationIdx] : null) || existing.creationDate || new Date('2099-01-01')
      });
    }
  };

  readSheet(shMonB);
  readSheet(shMonA);

  console.log(`[Sort Columns] Loaded sort data for ${map.size} vendors`);
  return map;
}

/** Parse a creation log date from monday.com (can be Date object, ISO string, or timestamp) */
function parseCreationDate_(val) {
  if (!val) return null;
  if (val instanceof Date && !isNaN(val.getTime())) return val;
  const d = new Date(val);
  return isNaN(d.getTime()) ? null : d;
}


/** ========== EMAIL ACTIVITY DETECTION ========== **/

/**
 * Get vendors with email activity in the last 7 days.
 * Searches sent + received threads from last 7 days.
 */
function getRecentEmailVendors_(allVendors) {
  const recentSet = new Set();
  const recentDates = {};  // vendor (lowercase) -> oldest thread date

  try {
    const vendorMap = buildVendorLabelMap_(allVendors);
    const threads = GmailApp.search('newer_than:7d (label:00.received OR label:sent)', 0, 300);
    console.log(`Recent email search: ${threads.length} threads`);

    for (const thread of threads) {
      const vendor = matchVendorFromThread_(thread, vendorMap, allVendors);
      if (vendor) {
        recentSet.add(vendor);
        const threadDate = thread.getLastMessageDate();
        if (!recentDates[vendor] || threadDate < recentDates[vendor]) {
          recentDates[vendor] = threadDate;
        }
      }
    }
  } catch (e) {
    console.log('Error searching recent emails:', e.message);
  }

  return { recentSet, recentDates };
}

/**
 * Get vendors with actionable emails.
 * Returns { actionableSet, actionableDates, inboxOnlySet, inboxDates, profitiseInternal }
 */
function getActionableVendors_(allVendors) {
  const actionableSet = new Set();
  const actionableDates = {};  // vendor (lowercase) -> oldest thread date
  const inboxOnlySet = new Set();
  const inboxDates = {};  // vendor (lowercase) -> oldest inbox thread date
  let profitiseInternal = false;

  try {
    const vendorMap = buildVendorLabelMap_(allVendors);

    // Search: (inbox OR 00.received) AND -is:snoozed
    const actionableThreads = GmailApp.search('(label:inbox OR label:00.received) AND -is:snoozed', 0, 500);
    console.log(`Actionable email search: ${actionableThreads.length} threads`);

    // Search: inbox only (subset for top-tier priority)
    const inboxThreadMap = new Map();  // threadId -> thread
    const inboxThreads = GmailApp.search('label:inbox', 0, 300);
    for (const t of inboxThreads) inboxThreadMap.set(t.getId(), t);

    // Check for Profitise Internal
    const internalThreads = GmailApp.search('label:00.received AND label:zzzVendors/Profitise Internal', 0, 5);
    if (internalThreads.length > 0) profitiseInternal = true;

    for (const thread of actionableThreads) {
      const vendor = matchVendorFromThread_(thread, vendorMap, allVendors);
      if (vendor) {
        actionableSet.add(vendor);
        const threadDate = thread.getLastMessageDate();
        if (!actionableDates[vendor] || threadDate < actionableDates[vendor]) {
          actionableDates[vendor] = threadDate;
        }
        if (inboxThreadMap.has(thread.getId())) {
          inboxOnlySet.add(vendor);
          if (!inboxDates[vendor] || threadDate < inboxDates[vendor]) {
            inboxDates[vendor] = threadDate;
          }
        }
      }
    }

    // Also check inbox-only threads that might not be in actionable search
    for (const [tid, thread] of inboxThreadMap) {
      const vendor = matchVendorFromThread_(thread, vendorMap, allVendors);
      if (vendor && !inboxDates[vendor]) {
        inboxOnlySet.add(vendor);
        inboxDates[vendor] = thread.getLastMessageDate();
      }
    }
  } catch (e) {
    console.log('Error searching actionable emails:', e.message);
  }

  return { actionableSet, actionableDates, inboxOnlySet, inboxDates, profitiseInternal };
}

/** Build a map of lowercase vendor name -> true for fast lookup */
function buildVendorLabelMap_(allVendors) {
  const map = new Map();
  for (const v of allVendors) {
    map.set(v.name.toLowerCase(), v.name);
  }
  return map;
}

/** Match a vendor from a Gmail thread using zzzVendors labels or name matching */
function matchVendorFromThread_(thread, vendorMap, allVendors) {
  const labels = thread.getLabels();
  for (const label of labels) {
    const labelName = label.getName();
    if (labelName.startsWith('zzzVendors/')) {
      const vendorFromLabel = labelName.substring('zzzVendors/'.length).toLowerCase();
      // Direct match
      if (vendorMap.has(vendorFromLabel)) return vendorFromLabel;
      // Partial match
      for (const [key] of vendorMap) {
        if (vendorFromLabel.includes(key) || key.includes(vendorFromLabel)) return key;
      }
    }
  }

  // Fallback: name match in subject/from
  try {
    const subject = thread.getFirstMessageSubject().toLowerCase();
    const messages = thread.getMessages();
    let emailText = subject;
    if (messages.length > 0) {
      emailText += ' ' + messages[0].getFrom().toLowerCase();
    }
    for (const [key] of vendorMap) {
      if (key.length >= 4 && emailText.includes(key)) return key;
    }
  } catch (e) { /* skip */ }

  return null;
}


/** ========== MONTHLY RETURNS DETECTION ========== **/

/**
 * Get vendors with open Monthly Returns tasks
 * Looks at monday.com tasks where Project = "Monthly Returns" and status is not "Done"
 * Returns a Set of vendor names (lowercased) for matching
 */
function getVendorsWithOpenMonthlyReturns_() {
  const vendorSet = new Set();

  try {
    // Get all tasks from cache (already filtered by group)
    const allTasks = getAllMondayTasks_();

    if (!allTasks || allTasks.length === 0) {
      console.log('No monday.com tasks found for Monthly Returns check');
      return vendorSet;
    }

    // Filter to Monthly Returns tasks that are not Done
    const openMonthlyReturns = allTasks.filter(task => {
      const project = (task.project || '').toLowerCase();
      const status = (task.status || '').toLowerCase();
      return project === 'monthly returns' && status !== 'done';
    });

    console.log(`Found ${openMonthlyReturns.length} open Monthly Returns tasks`);

    // Extract vendor names from task names
    // Format: "Monthly Returns (Month Year) - VendorName"
    for (const task of openMonthlyReturns) {
      const taskName = task.name || '';
      const dashIndex = taskName.lastIndexOf(' - ');
      if (dashIndex > 0) {
        const vendorName = taskName.substring(dashIndex + 3).trim();
        if (vendorName) {
          vendorSet.add(vendorName.toLowerCase());
          console.log(`  Monthly Returns vendor: ${vendorName}`);
        }
      }
    }

    console.log(`Extracted ${vendorSet.size} unique vendors with open Monthly Returns`);
  } catch (e) {
    console.log('Error getting Monthly Returns vendors:', e.message);
  }

  return vendorSet;
}


/** ========== HOT ZONE DETECTION ========== **/

/**
 * Search Gmail for threads that indicate "hot" vendors
 * A vendor is "hot" if they have:
 *   1. Emails with label:00.received, OR
 *   2. ANY emails currently in the inbox
 *
 * Return a Set of vendor names (lowercased) that have email activity
 *
 * Detection methods (in priority order):
 * 1. Gmail label "zzzVendors/<vendor_name>" (most accurate)
 * 2. Exact vendor name match in subject/sender/recipient
 */
function getHotVendorsFromGmail_(allVendors) {
  const inboxSet = new Set();  // Highest priority - vendors with emails in inbox
  const hotSet = new Set();    // Other hot vendors (00.received, recent sent)
  const inboxOldestDate = {};  // vendor name (lowercase) -> oldest inbox email date

  try {
    // Build vendor name lookup for fast matching
    const vendorMap = new Map();
    const vendorNames = allVendors.map(v => {
      const nameLower = v.name.toLowerCase();
      vendorMap.set(nameLower, v.name);
      return {
        name: v.name,
        nameLower: nameLower
      };
    });

    let labelMatches = 0;
    let exactMatches = 0;

    // SEARCH 1: Inbox + 00.received (combined, excluding snoozed) — the actionable pool
    console.log('Searching inbox + 00.received for vendor emails (excluding snoozed)...');
    const actionableThreads = GmailApp.search('(label:inbox OR label:00.received) AND -is:snoozed', 0, 500);
    console.log(`Found ${actionableThreads.length} actionable threads`);

    // SEARCH 2: Sent emails in last 188 hours (approx 7.8 days) — for "hot" detection only
    console.log('Searching for recently sent emails...');
    const sentThreads = GmailApp.search('label:sent newer_than:188h', 0, 200);
    console.log(`Found ${sentThreads.length} sent threads in last 188h`);

    // Track actionable thread IDs (inbox + 00.received, unsnoozed)
    const actionableThreadIds = new Set(actionableThreads.map(t => t.getId()));

    // Process all threads
    const processedThreadIds = new Set();

    // Helper function to match vendor from thread
    const matchVendorFromThread = (thread) => {
      const labels = thread.getLabels();
      for (const label of labels) {
        const labelName = label.getName();
        if (labelName.startsWith('zzzVendors/')) {
          const vendorNameFromLabel = labelName.substring('zzzVendors/'.length).toLowerCase();
          if (vendorMap.has(vendorNameFromLabel)) {
            return { vendor: vendorNameFromLabel, matchType: 'label' };
          }
          for (const vendor of vendorNames) {
            if (vendorNameFromLabel.includes(vendor.nameLower) ||
                vendor.nameLower.includes(vendorNameFromLabel)) {
              return { vendor: vendor.nameLower, matchType: 'label-partial' };
            }
          }
        }
      }

      // Try exact name match in subject/sender/recipient
      const subject = thread.getFirstMessageSubject().toLowerCase();
      const messages = thread.getMessages();
      let emailText = subject;
      if (messages.length > 0) {
        const firstMsg = messages[0];
        emailText += ' ' + firstMsg.getFrom().toLowerCase();
        emailText += ' ' + firstMsg.getTo().toLowerCase();
      }

      for (const vendor of vendorNames) {
        if (emailText.includes(vendor.nameLower)) {
          return { vendor: vendor.nameLower, matchType: 'exact' };
        }
      }

      return null;
    };

    // Combine all threads for processing
    const allThreads = [...actionableThreads, ...sentThreads];

    for (const thread of allThreads) {
      try {
        const threadId = thread.getId();
        if (processedThreadIds.has(threadId)) continue;
        processedThreadIds.add(threadId);

        const match = matchVendorFromThread(thread);
        if (match) {
          // If this thread is actionable (inbox OR 00.received, unsnoozed)
          if (actionableThreadIds.has(threadId)) {
            inboxSet.add(match.vendor);
            // Track oldest email date per vendor for sorting
            var threadDate = thread.getLastMessageDate();
            if (!inboxOldestDate[match.vendor] || threadDate < inboxOldestDate[match.vendor]) {
              inboxOldestDate[match.vendor] = threadDate;
            }
            console.log(`ACTIONABLE: ${vendorMap.get(match.vendor) || match.vendor} (${match.matchType}, date: ${threadDate})`);
          } else {
            // Sent-only threads go to hot (vendor we've emailed but no inbox/received thread)
            if (!inboxSet.has(match.vendor)) {
              hotSet.add(match.vendor);
              console.log(`HOT: ${vendorMap.get(match.vendor) || match.vendor} (${match.matchType})`);
            }
          }
          if (match.matchType === 'label' || match.matchType === 'label-partial') {
            labelMatches++;
          } else {
            exactMatches++;
          }
        }
      } catch (e) {
        console.log(`Error processing thread: ${e.message}`);
      }
    }

    console.log(`Vendor detection summary: ${labelMatches} label matches, ${exactMatches} exact matches`);
    console.log(`Inbox vendors: ${inboxSet.size}, Hot vendors: ${hotSet.size}`);
    console.log(`Total unique threads processed: ${processedThreadIds.size}`);

  } catch (e) {
    console.log(`Error searching Gmail: ${e.message}`);
  }

  return { inboxSet, hotSet, inboxOldestDate };
}


/** ========== STATUS & NOTES LOOKUP ========== **/

/**
 * Build status lookup maps from monday.com sheets
 * Returns { buyers: Map, affiliates: Map }
 */
function buildStatusMaps_(shMonB, shMonA) {
  const buyersMap = new Map();
  const affiliatesMap = new Map();

  if (shMonB) {
    const buyersData = shMonB.getDataRange().getValues();
    const headers = buyersData[0].map(h => String(h || '').trim().toLowerCase());
    const statusIdx = findHeaderIndex_(headers, STATUS_HEADER_CANDIDATES);
    console.log(`[Buyers] Status column index: ${statusIdx} (header: ${statusIdx >= 0 ? buyersData[0][statusIdx] : 'NOT FOUND'})`);

    for (let i = 1; i < buyersData.length; i++) { // Skip header
      const row = buyersData[i];
      const vendor = normalizeName_(row[0]);
      const status = (statusIdx >= 0 && statusIdx < row.length) ? String(row[statusIdx] || '').trim() : '';
      if (vendor) {
        buyersMap.set(vendor.toLowerCase(), status);
      }
    }
  }

  if (shMonA) {
    const affiliatesData = shMonA.getDataRange().getValues();
    const headers = affiliatesData[0].map(h => String(h || '').trim().toLowerCase());
    const statusIdx = findHeaderIndex_(headers, STATUS_HEADER_CANDIDATES);
    console.log(`[Affiliates] Status column index: ${statusIdx} (header: ${statusIdx >= 0 ? affiliatesData[0][statusIdx] : 'NOT FOUND'})`);

    for (let i = 1; i < affiliatesData.length; i++) { // Skip header
      const row = affiliatesData[i];
      const vendor = normalizeName_(row[0]);
      const status = (statusIdx >= 0 && statusIdx < row.length) ? String(row[statusIdx] || '').trim() : '';
      if (vendor) {
        affiliatesMap.set(vendor.toLowerCase(), status);
      }
    }
  }

  return { buyers: buyersMap, affiliates: affiliatesMap };
}

/**
 * Build notes lookup maps from monday.com sheets
 * Returns { buyers: Map, affiliates: Map }
 */
function buildNotesMaps_(shMonB, shMonA) {
  const buyersMap = new Map();
  const affiliatesMap = new Map();

  if (shMonB) {
    const buyersData = shMonB.getDataRange().getValues();
    const headers = buyersData[0].map(h => String(h || '').trim().toLowerCase());
    const notesIdx = findHeaderIndex_(headers, NOTES_HEADER_CANDIDATES);
    console.log(`[Buyers] Notes column index: ${notesIdx} (header: ${notesIdx >= 0 ? buyersData[0][notesIdx] : 'NOT FOUND'})`);

    for (let i = 1; i < buyersData.length; i++) { // Skip header
      const row = buyersData[i];
      const vendor = normalizeName_(row[0]);
      const notes = (notesIdx >= 0 && notesIdx < row.length) ? String(row[notesIdx] || '').trim() : '';
      if (vendor) {
        buyersMap.set(vendor.toLowerCase(), notes);
      }
    }
  }

  if (shMonA) {
    const affiliatesData = shMonA.getDataRange().getValues();
    const headers = affiliatesData[0].map(h => String(h || '').trim().toLowerCase());
    const notesIdx = findHeaderIndex_(headers, NOTES_HEADER_CANDIDATES);
    console.log(`[Affiliates] Notes column index: ${notesIdx} (header: ${notesIdx >= 0 ? affiliatesData[0][notesIdx] : 'NOT FOUND'})`);

    for (let i = 1; i < affiliatesData.length; i++) { // Skip header
      const row = affiliatesData[i];
      const vendor = normalizeName_(row[0]);
      const notes = (notesIdx >= 0 && notesIdx < row.length) ? String(row[notesIdx] || '').trim() : '';
      if (vendor) {
        affiliatesMap.set(vendor.toLowerCase(), notes);
      }
    }
  }

  return { buyers: buyersMap, affiliates: affiliatesMap };
}

/**
 * Lookup status for a vendor from the status maps
 * Returns the raw Group/Status value (not normalized)
 */
function lookupStatus_(name, type, statusMaps) {
  const key = name.toLowerCase();
  const isBuyer = (type || '').toLowerCase().startsWith('buyer');

  if (isBuyer && statusMaps.buyers.has(key)) {
    return statusMaps.buyers.get(key);
  }
  if (!isBuyer && statusMaps.affiliates.has(key)) {
    return statusMaps.affiliates.get(key);
  }
  // Try the other map as fallback
  if (statusMaps.buyers.has(key)) {
    return statusMaps.buyers.get(key);
  }
  if (statusMaps.affiliates.has(key)) {
    return statusMaps.affiliates.get(key);
  }
  return '';
}

/**
 * Lookup notes for a vendor from the notes maps
 */
function lookupNotes_(name, type, notesMaps) {
  const key = name.toLowerCase();
  const isBuyer = (type || '').toLowerCase().startsWith('buyer');

  if (isBuyer && notesMaps.buyers.has(key)) {
    return notesMaps.buyers.get(key);
  }
  if (!isBuyer && notesMaps.affiliates.has(key)) {
    return notesMaps.affiliates.get(key);
  }
  // Try the other map as fallback
  if (notesMaps.buyers.has(key)) {
    return notesMaps.buyers.get(key);
  }
  if (notesMaps.affiliates.has(key)) {
    return notesMaps.affiliates.get(key);
  }
  return '';
}


/** ========== GMAIL SUBLABEL MAPPING ========== **/

/**
 * Read Gmail sublabel mappings from Settings sheet
 * Returns a Map of lowercase vendor name -> gmail sublabel
 */
function readGmailSublabelMap_(ss) {
  const sh = ss.getSheetByName(SHEET_SETTINGS);
  const map = new Map();

  if (!sh) return map;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return map;

  // Assuming Column A = Vendor Name, Column B = Gmail Sublabel
  const data = sh.getRange(2, 1, lastRow - 1, 2).getValues();

  for (const row of data) {
    const vendorName = String(row[0] || '').trim();
    const sublabel = String(row[1] || '').trim();

    if (vendorName && sublabel) {
      map.set(vendorName.toLowerCase(), sublabel);
    }
  }

  console.log(`Loaded ${map.size} Gmail sublabel mappings from Settings`);
  return map;
}


/** ========== READERS & HELPERS ========== **/

/**
 * Read a metric sheet (L1M/L6M) and return [{name, ttl, type, source}, ...]
 * Uses header index detection for vendor and TTL columns.
 */
function readMetricSheet_(sh, type, sourceLabel, blacklist) {
  const { columns, rows } = readObjectsFromSheet_(sh);
  console.log(`[${sourceLabel}] Total rows:`, rows.length);

  if (!rows.length) return [];
  const vHeader = firstExistingHeader_(columns, VENDOR_HEADER_CANDIDATES, sh.getName(), 'vendor');
  const tHeader = firstExistingHeader_(columns, TTL_HEADER_CANDIDATES, sh.getName(), 'TTL , USD');

  const out = [];
  for (const r of rows) {
    const name = normalizeName_(r[vHeader]);
    if (!name) continue;

    const key = name.toLowerCase();
    if (blacklist.has(key)) {
      recordSkipReason_({ name, source: sourceLabel, reason: 'blacklist' });
      continue;
    }

    out.push({
      name,
      ttl: toNumber_(r[tHeader]),
      type,
      source: sourceLabel
    });
  }
  return dedupeKeepMaxTTL_(out);
}

/**
 * Read a monday.com sheet:
 * - Name from Column A (index 0)
 * - Aliases in Column P (index 15)
 * - TTL optional via recognized TTL header (or specific column)
 * - Status from Column S (index 18)
 * - Notes from Column D (index 3)
 * - Dedupes vs existingSet and blacklist
 */
function readMondaySheet_(sh, type, sourceLabel, blacklist, existingSet) {
  console.log(`\n[${sourceLabel}] START READ`);

  // DEBUG: Set a vendor name to trace through the entire process (case-insensitive)
  const DEBUG_VENDOR = 'purity';  // Change this to trace a specific vendor

  // Get raw array data instead of objects for monday.com sheets
  const allValues = sh.getDataRange().getValues();
  if (allValues.length < 2) return []; // Need at least header + 1 row

  const headers = allValues[0].map(h => String(h || '').trim());
  const headersLower = headers.map(h => h.toLowerCase());
  console.log(`[${sourceLabel}] Total rows:`, allValues.length - 1);

  // For monday.com exports: name is always in column A (index 0)
  const nameIdx = 0; // Column A

  const ttlIdx = headers.findIndex(h =>
    eq_(h, 'TTL , USD') || eq_(h, 'TTL, USD') || eq_(h, 'TTL USD')
  );
  const typeIdx = headers.findIndex(h => eq_(h, 'Type'));

  // Auto-detect Status/Group and Notes columns by header name
  const statusIdx = findHeaderIndex_(headersLower, STATUS_HEADER_CANDIDATES);
  const notesIdx = findHeaderIndex_(headersLower, NOTES_HEADER_CANDIDATES);
  console.log(`[${sourceLabel}] Status column: ${statusIdx >= 0 ? headers[statusIdx] : 'NOT FOUND'} (idx ${statusIdx})`);
  console.log(`[${sourceLabel}] Notes column: ${notesIdx >= 0 ? headers[notesIdx] : 'NOT FOUND'} (idx ${notesIdx})`)

  const out = [];
  let skippedBlacklist = 0;
  let skippedExisting = 0;
  let skippedAlias = 0;
  let skippedNoName = 0;

  // Process rows starting from index 1 (skip header)
  for (let i = 1; i < allValues.length; i++) {
    const row = allValues[i];

    // Name is always in column A
    const rawName = row[nameIdx] || '';
    const name = normalizeName_(rawName);
    const isDebugVendor = DEBUG_VENDOR && name.toLowerCase().includes(DEBUG_VENDOR.toLowerCase());

    if (isDebugVendor) {
      console.log(`\n🔍 DEBUG [${sourceLabel}] Found "${name}" (raw: "${rawName}") at row ${i + 1}`);
    }

    if (!name) {
      skippedNoName++;
      if (isDebugVendor) console.log(`  ❌ SKIPPED: No name after normalization`);
      continue;
    }

    const key = name.toLowerCase();

    if (blacklist.has(key)) {
      skippedBlacklist++;
      if (isDebugVendor) console.log(`  ❌ SKIPPED: In blacklist`);
      continue;
    }

    if (existingSet.has(key)) {
      skippedExisting++;
      if (isDebugVendor) console.log(`  ❌ SKIPPED: Already exists in L1M/L6M set`);
      continue;
    }

    // Aliases in Column P (index 15)
    const aliasRaw = (ALIAS_COL_INDEX < row.length) ? (row[ALIAS_COL_INDEX] || '') : '';
    const aliases = String(aliasRaw).split(',')
      .map(s => normalizeName_(s).toLowerCase())
      .filter(Boolean);

    if (isDebugVendor && aliases.length > 0) {
      console.log(`  Aliases: ${aliases.join(', ')}`);
    }

    if (aliases.some(a => existingSet.has(a) || blacklist.has(a))) {
      skippedAlias++;
      if (isDebugVendor) console.log(`  ❌ SKIPPED: Alias matches existing/blacklist`);
      recordSkipReason_({ name, source: sourceLabel, reason: 'alias-duplicate' });
      continue;
    }

    const ttl = (ttlIdx !== -1 && ttlIdx < row.length) ? toNumber_(row[ttlIdx]) : 0;
    const explicitType = (typeIdx !== -1 && typeIdx < row.length) ? String(row[typeIdx] || '').trim() : '';
    const finalType = explicitType || type;
    // Use raw status value (Group name) instead of normalizing
    const status = (statusIdx >= 0 && statusIdx < row.length) ? String(row[statusIdx] || '').trim() : '';
    const notes = (notesIdx >= 0 && notesIdx < row.length) ? String(row[notesIdx] || '').trim() : '';

    if (isDebugVendor) {
      console.log(`  ✅ ADDED: name="${name}", ttl=${ttl}, type="${finalType}", status="${status}"`);
    }

    out.push({ name, ttl, type: finalType, source: sourceLabel, status, notes });
    existingSet.add(key);
  }

  console.log(`[${sourceLabel}] Skip summary:`, {
    noName: skippedNoName,
    blacklist: skippedBlacklist,
    existing: skippedExisting,
    alias: skippedAlias,
    added: out.length
  });
  console.log(`[${sourceLabel}] END READ\n`);

  return dedupeKeepMaxTTL_(out);
}

/** Get status rank for sorting (maps raw status to priority rank) */
function getStatusRank_(rawStatus) {
  const v = String(rawStatus || '').trim().toLowerCase();
  if (!v) return STATUS_RANK['other'];

  if (v.includes('live')) return STATUS_RANK['live'];
  if (v.includes('onboard')) return STATUS_RANK['onboarding'];
  if (v.includes('pause')) return STATUS_RANK['paused'];
  if (v.includes('pre')) return STATUS_RANK['preonboarding'];
  if (v.includes('early')) return STATUS_RANK['early talks'];
  if (v.includes('top') && v.includes('500')) return STATUS_RANK['top 500 remodelers'];
  if (v.includes('dead') || (v.includes('no') && v.includes('go')) || v.includes('closed')) return STATUS_RANK['dead'];

  return STATUS_RANK['other'];
}

/** Normalize buyer or affiliate status to our buckets */
function normalizeStatus_(s) {
  const v = String(s || '').trim().toLowerCase();
  if (!v) return 'other';

  if (v.includes('live')) return 'live';
  if (v.includes('onboard')) return 'onboarding';
  if (v.includes('pause')) return 'paused';
  if (v.includes('pre')) return 'preonboarding';
  if (v.includes('early')) return 'early talks';
  if (v.includes('top') && v.includes('500')) return 'top 500 remodelers';
  if (v.includes('dead') || (v.includes('no') && v.includes('go')) || v.includes('closed')) return 'dead';

  return 'other';
}

/** Reads Settings column E "blacklist" and returns a Set of normalized, lowercased names */
function readBlacklist_(ss) {
  const sh = ss.getSheetByName(SHEET_SETTINGS);
  const set = new Set();
  if (!sh) return set;

  const last = sh.getLastRow();
  if (last < 2) return set;

  const header = String(sh.getRange(1, 5).getValue() || '').trim().toLowerCase(); // E1
  if (header !== 'blacklist' && header !== 'vendor blacklist') return set;

  const vals = sh.getRange(2, 5, last - 1, 1).getValues().flat(); // E2:E
  for (const v of vals) {
    const norm = normalizeName_(v);
    if (norm) set.add(norm.toLowerCase());
  }
  return set;
}

/** Data readers */
function readObjectsFromSheet_(sh) {
  const values = sh.getDataRange().getValues();
  if (!values.length) return { columns: [], rows: [] };
  const headers = values[0].map(h => String(h || '').trim());
  const rows = values.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => (obj[h] = row[i]));
    return obj;
  });
  return { columns: headers, rows };
}

/** De-dupe by lowercased name keeping max TTL */
function dedupeKeepMaxTTL_(rows) {
  const map = new Map();
  for (const r of rows) {
    const key = r.name.toLowerCase();
    const prev = map.get(key);
    if (!prev || (r.ttl || 0) > (prev.ttl || 0)) {
      map.set(key, r);
    }
  }
  return Array.from(map.values());
}

/** Normalize display name by removing leading "[123]" token if present */
function normalizeName_(v) {
  if (v == null) return '';
  let s = String(v).trim();
  s = s.replace(/^\[\s*\d+\s*\]\s*/, '');
  return s.trim();
}

/** Robust currency parsing */
function toNumber_(v) {
  if (v == null || v === '') return 0;
  if (typeof v === 'number') return v;
  let s = String(v).trim();
  let neg = false;
  if (s.startsWith('(') && s.endsWith(')')) {
    neg = true; s = s.slice(1, -1);
  }
  s = s.replace(/[$,\s]/g, '');
  const n = parseFloat(s);
  return isNaN(n) ? 0 : (neg ? -n : n);
}

/** Case-insensitive string equality */
function eq_(a, b) {
  return String(a || '').trim().toLowerCase() === String(b || '').trim().toLowerCase();
}

/** Find the first header that exists from a list of candidates */
function firstExistingHeader_(columns, candidates, sheetName, purpose) {
  for (const c of candidates) {
    if (columns.some(col => eq_(col, c))) {
      return c;
    }
  }
  // Return first candidate as fallback
  console.log(`Warning: No matching header found for ${purpose} in ${sheetName}. Using first candidate: ${candidates[0]}`);
  return candidates[0];
}

/** Find header index (optional, returns null if not found) */
function findHeaderOptionalIndex_(columns, header) {
  const idx = columns.findIndex(col => eq_(col, header));
  return idx === -1 ? null : idx;
}

/** Find the first matching header index from a list of candidates (lowercased headers) */
function findHeaderIndex_(headersLower, candidates) {
  for (const candidate of candidates) {
    const candidateLower = candidate.toLowerCase();
    const idx = headersLower.findIndex(h => h === candidateLower || h.includes(candidateLower));
    if (idx >= 0) return idx;
  }
  return -1;
}

/** Ensure a sheet exists, create if not */
function ensureSheet_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  }
  return sh;
}

/** Get a sheet, throw if not found */
function mustGetSheet_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) {
    throw new Error(`Required sheet "${name}" not found`);
  }
  return sh;
}

/** Record skip reasons for debugging */
function recordSkipReason_(obj) {
  SKIP_REASONS.push(obj);
}

/** Get skip reasons (for debugging) */
function getSkipReasons() {
  return SKIP_REASONS;
}
