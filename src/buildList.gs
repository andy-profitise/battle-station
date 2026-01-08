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
 * Generates columns: Vendor | TTL, USD | Source | Status | Notes | Gmail Link | no snoozing | processed?
 *
 * HOT ZONE: Vendors with recent emails (label:00.received last 7 days OR in inbox) at top
 * NORMAL ZONE: All other vendors sorted by TTL, type, alpha
 */
function buildListWithGmailAndNotes() {
  const ss = SpreadsheetApp.getActive();

  console.log('=== BUILD LIST WITH GMAIL & NOTES START ===');

  // Get all vendors using existing buildVendorList logic
  const shBuyL1 = mustGetSheet_(ss, SHEET_BUYERS_L1M);
  const shBuyL6 = mustGetSheet_(ss, SHEET_BUYERS_L6M);
  const shAffL1 = mustGetSheet_(ss, SHEET_AFFILIATES_L1M);
  const shAffL6 = mustGetSheet_(ss, SHEET_AFFILIATES_L6M);
  const shMonB  = ss.getSheetByName(SHEET_MON_BUYERS);
  const shMonA  = ss.getSheetByName(SHEET_MON_AFFILIATES);

  const blacklist = readBlacklist_(ss);

  // Read all vendors
  const buyersL1M = readMetricSheet_(shBuyL1, 'Buyer', 'Buyers L1M', blacklist);
  const buyersL1MSet = new Set(buyersL1M.map(r => r.name.toLowerCase()));

  const buyersL6MAll = readMetricSheet_(shBuyL6, 'Buyer', 'Buyers L6M', blacklist);
  const buyersL6M = buyersL6MAll.filter(r => !buyersL1MSet.has(r.name.toLowerCase()));

  const buyersExisting = new Set([...buyersL1MSet, ...buyersL6M.map(r => r.name.toLowerCase())]);
  const buyersMon = shMonB ? readMondaySheet_(shMonB, 'Buyer', 'buyers monday.com', blacklist, buyersExisting) : [];

  const affL1M = readMetricSheet_(shAffL1, 'Affiliate', 'Affiliates L1M', blacklist);
  const affL1MSet = new Set(affL1M.map(r => r.name.toLowerCase()));

  const affL6MAll = readMetricSheet_(shAffL6, 'Affiliate', 'Affiliates L6M', blacklist);
  const affL6M = affL6MAll.filter(r => !affL1MSet.has(r.name.toLowerCase()));

  const affExisting = new Set([...affL1MSet, ...affL6M.map(r => r.name.toLowerCase())]);
  const affMon = shMonA ? readMondaySheet_(shMonA, 'Affiliate', 'affiliates monday.com', blacklist, affExisting) : [];

  console.log('Data loaded:', {
    buyersL1M: buyersL1M.length,
    buyersL6M: buyersL6M.length,
    buyersMon: buyersMon.length,
    affL1M: affL1M.length,
    affL6M: affL6M.length,
    affMon: affMon.length
  });

  // Split by TTL
  const gt0 = [];
  const z_buyL6 = [], z_affL6 = [], z_buyL1 = [], z_affL1 = [], z_mon = [];

  const pushByTtl = (arr, zeroTarget) => {
    for (const r of arr) {
      if ((r.ttl || 0) > 0) gt0.push(r);
      else zeroTarget.push(r);
    }
  };

  pushByTtl(buyersL6M, z_buyL6);
  pushByTtl(affL6M, z_affL6);
  pushByTtl(buyersL1M, z_buyL1);
  pushByTtl(affL1M, z_affL1);

  for (const r of buyersMon) ((r.ttl || 0) > 0 ? gt0 : z_mon).push(r);
  for (const r of affMon) ((r.ttl || 0) > 0 ? gt0 : z_mon).push(r);

  // Sort >0 TTL by TTL desc, then type (buyers first), then alpha
  gt0.sort((a, b) => {
    const ttlDiff = (b.ttl || 0) - (a.ttl || 0);
    if (ttlDiff !== 0) return ttlDiff;
    const rankA = (a.type || '').toLowerCase().startsWith('buyer') ? 0 : 1;
    const rankB = (b.type || '').toLowerCase().startsWith('buyer') ? 0 : 1;
    if (rankA !== rankB) return rankA - rankB;
    return String(a.name).localeCompare(String(b.name));
  });

  // Sort zero-TTL groups alphabetically
  const alpha = (a, b) => String(a.name).localeCompare(String(b.name));
  z_buyL6.sort(alpha);
  z_affL6.sort(alpha);
  z_buyL1.sort(alpha);
  z_affL1.sort(alpha);

  // Sort monday.com zero-TTL by status rank, then type, then alpha
  z_mon.sort((a, b) => {
    const sA = getStatusRank_(a.status);
    const sB = getStatusRank_(b.status);
    if (sA !== sB) return sA - sB;
    const rankA = (a.type || '').toLowerCase().startsWith('buyer') ? 0 : 1;
    const rankB = (b.type || '').toLowerCase().startsWith('buyer') ? 0 : 1;
    if (rankA !== rankB) return rankA - rankB;
    return String(a.name).localeCompare(String(b.name));
  });

  // Assemble all vendors
  const all = [...gt0, ...z_buyL6, ...z_affL6, ...z_buyL1, ...z_affL1, ...z_mon];

  console.log('Total before status/notes lookup:', all.length);

  // Build status and notes maps from monday.com sheets
  const statusMaps = buildStatusMaps_(shMonB, shMonA);
  const notesMaps = buildNotesMaps_(shMonB, shMonA);

  console.log('Maps built:', {
    buyersStatus: statusMaps.buyers.size,
    affiliatesStatus: statusMaps.affiliates.size,
    buyersNotes: notesMaps.buyers.size,
    affiliatesNotes: notesMaps.affiliates.size
  });

  // Lookup status and notes for each vendor
  for (const r of all) {
    r.status = lookupStatus_(r.name, r.type, statusMaps);
    r.notes = lookupNotes_(r.name, r.type, notesMaps);
  }

  // PRIORITY ZONES: Detect vendors with emails or chat activity
  // 1. INBOX (highest) - vendors with emails currently in inbox
  // 2. CHAT - vendors detected via OCR from Teams/Telegram/WhatsApp
  // 3. HOT - vendors with 00.received or recent sent emails
  // 4. NORMAL - everything else
  console.log('Detecting priority vendors...');
  const { inboxSet, hotSet } = getHotVendorsFromGmail_(all);
  console.log('Inbox vendors found:', inboxSet.size);
  console.log('Hot vendors found:', hotSet.size);

  // Get OCR-detected vendors (from chat platforms)
  let chatSet = new Set();
  try {
    const ocrVendors = getOcrDetectedVendors();
    chatSet = new Set(ocrVendors.keys());
    console.log('Chat (OCR) vendors found:', chatSet.size);
  } catch (e) {
    console.log('Error getting OCR vendors:', e.message);
  }

  const inboxZone = [];
  const chatZone = [];
  const hotZone = [];
  const normalZone = [];

  for (const r of all) {
    const nameLower = r.name.toLowerCase();
    if (inboxSet.has(nameLower)) {
      inboxZone.push(r);
    } else if (chatSet.has(nameLower)) {
      chatZone.push(r);
    } else if (hotSet.has(nameLower)) {
      hotZone.push(r);
    } else {
      normalZone.push(r);
    }
  }

  // Final list: Inbox at top, then chat zone, then hot zone, then normal zone
  // Each zone keeps same sort order as `all` (already sorted)
  const finalList = [...inboxZone, ...chatZone, ...hotZone, ...normalZone];

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
  shOut.getRange(1, 1, 1, 8).setValues([[
    'Vendor', 'TTL , USD', 'Source', 'Status', 'Notes',
    'Gmail Link', 'no snoozing', 'processed?'
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

      return [gmailAll, gmailNoSnooze, false];
    });

    shOut.getRange(2, 6, gmailData.length, 3).setValues(gmailData);
  }

  // Auto-resize columns
  shOut.autoResizeColumns(1, 5);
  shOut.setColumnWidth(6, 100);
  shOut.setColumnWidth(7, 100);
  shOut.setColumnWidth(8, 100);

  console.log('=== BUILD LIST WITH GMAIL & NOTES END ===');

  const counts = {
    '🔥 HOT (recent emails)': hotZone.length,
    'Total vendors': finalList.length,
    'With status': finalList.filter(v => v.status).length,
    'With notes': finalList.filter(v => v.notes).length
  };

  console.log('Final counts:', counts);

  ss.toast(
    Object.entries(counts).map(([k,v]) => `${k}: ${v}`).join(' • '),
    '✅ List Built',
    8
  );
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

    // SEARCH 1: Inbox emails (highest priority)
    console.log('Searching inbox for vendor emails...');
    const inboxThreads = GmailApp.search('label:inbox', 0, 200);
    console.log(`Found ${inboxThreads.length} inbox threads`);

    // SEARCH 2: Emails with label:00.received (hot but not inbox priority)
    console.log('Searching for emails with label:00.received...');
    const recentThreads = GmailApp.search('label:00.received', 0, 200);
    console.log(`Found ${recentThreads.length} threads with label:00.received`);

    // SEARCH 3: Sent emails in last 188 hours (approx 7.8 days)
    console.log('Searching for recently sent emails...');
    const sentThreads = GmailApp.search('label:sent newer_than:188h', 0, 200);
    console.log(`Found ${sentThreads.length} sent threads in last 188h`);

    // Track inbox thread IDs separately for priority detection
    const inboxThreadIds = new Set(inboxThreads.map(t => t.getId()));

    // Process inbox threads first (these get highest priority)
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
    const allThreads = [...inboxThreads, ...recentThreads, ...sentThreads];

    for (const thread of allThreads) {
      try {
        const threadId = thread.getId();
        if (processedThreadIds.has(threadId)) continue;
        processedThreadIds.add(threadId);

        const match = matchVendorFromThread(thread);
        if (match) {
          // If this thread is in inbox, add to inbox set (highest priority)
          if (inboxThreadIds.has(threadId)) {
            inboxSet.add(match.vendor);
            console.log(`INBOX: ${vendorMap.get(match.vendor) || match.vendor} (${match.matchType})`);
          } else {
            // Only add to hotSet if not already in inboxSet
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

  return { inboxSet, hotSet };
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
