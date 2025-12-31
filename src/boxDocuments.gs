/************************************************************
 * BOX DOCUMENTS - Search Box.com for vendor-related documents
 *
 * Standalone module for Battle Station
 * Searches Box.com API for files matching vendor names (fuzzy)
 *
 * SETUP:
 * 1. Create a Box Developer App at https://app.box.com/developers/console
 *    - Choose "Custom App" with "User Authentication (OAuth 2.0)"
 *    - Set redirect URI to: https://script.google.com/macros/d/{SCRIPT_ID}/usercallback
 *    - Note your Client ID and Client Secret
 * 2. Run authorizeBox() from the menu to complete OAuth flow
 * 3. Test with testBoxConnection() or searchBoxForVendor("Company Name")
 *
 * After testing, integrate with Battle Station by adding to getBoxDocuments_()
 ************************************************************/

const BOX_CFG = {
  // ========== BOX APP CREDENTIALS ==========
  CLIENT_ID: 'mjf6skqa0e6ixkpcw5uhzdeksecedihi',
  CLIENT_SECRET: 'BpQl2SZlx3uHPZmftc5V70ajJfdIMYuJ',

  // OAuth endpoints
  AUTH_URL: 'https://account.box.com/api/oauth2/authorize',
  TOKEN_URL: 'https://api.box.com/oauth2/token',

  // API base
  API_BASE: 'https://api.box.com/2.0',

  // Search settings
  DEFAULT_LIMIT: 20,
  FILE_TYPES: ['pdf', 'docx', 'doc', 'xlsx', 'xls', 'pptx', 'ppt'],  // Optional filter

  // Property keys for storing tokens
  PROP_ACCESS_TOKEN: 'BOX_ACCESS_TOKEN',
  PROP_REFRESH_TOKEN: 'BOX_REFRESH_TOKEN',
  PROP_TOKEN_EXPIRY: 'BOX_TOKEN_EXPIRY'
};


/************************************************************
 * OAUTH 2.0 AUTHENTICATION
 ************************************************************/

/**
 * Get the OAuth2 service for Box
 * Uses the OAuth2 library: 1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF
 */
function getBoxService_() {
  return OAuth2.createService('box')
    .setAuthorizationBaseUrl(BOX_CFG.AUTH_URL)
    .setTokenUrl(BOX_CFG.TOKEN_URL)
    .setClientId(BOX_CFG.CLIENT_ID)
    .setClientSecret(BOX_CFG.CLIENT_SECRET)
    .setCallbackFunction('boxAuthCallback')
    .setPropertyStore(PropertiesService.getUserProperties())
    .setScope('root_readonly')  // Read access to files and folders
    .setParam('access_type', 'offline');  // Get refresh token
}

/**
 * Start the Box authorization flow
 * Run this from the Script Editor or menu
 */
function authorizeBox() {
  const service = getBoxService_();

  if (service.hasAccess()) {
    SpreadsheetApp.getUi().alert('✅ Box is already authorized!\n\nRun testBoxConnection() to verify.');
    return;
  }

  const authUrl = service.getAuthorizationUrl();

  const html = HtmlService.createHtmlOutput(
    '<h2>Box Authorization Required</h2>' +
    '<p>Click the link below to authorize access to your Box account:</p>' +
    '<p><a href="' + authUrl + '" target="_blank" style="font-size: 18px; color: blue;">🔗 Authorize Box Access</a></p>' +
    '<p>After authorizing, close this dialog and run <code>testBoxConnection()</code></p>'
  )
  .setWidth(450)
  .setHeight(200);

  SpreadsheetApp.getUi().showModalDialog(html, 'Box Authorization');
}

/**
 * OAuth callback handler
 */
function boxAuthCallback(request) {
  const service = getBoxService_();
  const authorized = service.handleCallback(request);

  if (authorized) {
    return HtmlService.createHtmlOutput(
      '<h2>✅ Success!</h2>' +
      '<p>Box has been authorized. You can close this window.</p>' +
      '<p>Run <code>testBoxConnection()</code> to verify the connection.</p>'
    );
  } else {
    return HtmlService.createHtmlOutput(
      '<h2>❌ Authorization Failed</h2>' +
      '<p>Please try again or check your Box app configuration.</p>'
    );
  }
}

/**
 * Revoke Box authorization (for testing/reset)
 */
function revokeBoxAuth() {
  const service = getBoxService_();
  service.reset();
  SpreadsheetApp.getUi().alert('Box authorization has been revoked.\n\nRun authorizeBox() to re-authorize.');
}


/************************************************************
 * BOX API FUNCTIONS
 ************************************************************/

/**
 * Make an authenticated request to Box API
 */
function boxApiRequest_(endpoint, options = {}) {
  const service = getBoxService_();

  if (!service.hasAccess()) {
    throw new Error('Box is not authorized. Run authorizeBox() first.');
  }

  const url = endpoint.startsWith('http') ? endpoint : BOX_CFG.API_BASE + endpoint;

  const fetchOptions = {
    method: options.method || 'get',
    headers: {
      'Authorization': 'Bearer ' + service.getAccessToken(),
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };

  if (options.payload) {
    fetchOptions.payload = JSON.stringify(options.payload);
  }

  const response = UrlFetchApp.fetch(url, fetchOptions);
  const code = response.getResponseCode();
  const text = response.getContentText();

  if (code >= 200 && code < 300) {
    return JSON.parse(text);
  } else {
    Logger.log(`Box API Error (${code}): ${text}`);
    throw new Error(`Box API Error (${code}): ${text.substring(0, 200)}`);
  }
}

/**
 * Search Box for files matching a query (vendor name)
 * Searches both file NAMES and file CONTENT
 *
 * @param {string} query - The search query (vendor name)
 * @param {object} options - Optional settings
 * @returns {array} Array of matching files with metadata
 */
function searchBox(query, options = {}) {
  const limit = options.limit || BOX_CFG.DEFAULT_LIMIT;
  const type = options.type || 'file';  // 'file', 'folder', or 'web_link'

  // Build search URL with parameters
  // Search both name and content to find files that mention the vendor
  let searchUrl = `/search?query=${encodeURIComponent(query)}&limit=${limit}`;

  // Optional: filter by type
  if (type) {
    searchUrl += `&type=${type}`;
  }

  // Optional: filter by file extensions
  if (options.fileExtensions && options.fileExtensions.length > 0) {
    searchUrl += `&file_extensions=${options.fileExtensions.join(',')}`;
  }

  // Optional: limit to specific folder(s)
  if (options.folderIds && options.folderIds.length > 0) {
    searchUrl += `&ancestor_folder_ids=${options.folderIds.join(',')}`;
  }

  // Request additional fields including parent folder info
  searchUrl += '&fields=id,name,type,description,created_at,modified_at,size,parent,path_collection,shared_link';

  try {
    const result = boxApiRequest_(searchUrl);

    return result.entries.map(item => ({
      id: item.id,
      name: item.name,
      type: item.type,
      description: item.description || '',
      createdAt: item.created_at,
      modifiedAt: item.modified_at,
      size: item.size,
      folderPath: item.path_collection?.entries?.map(e => e.name).join('/') || '',
      parentFolder: item.parent?.name || '',
      parentFolderId: item.parent?.id || '',
      parentFolderUrl: item.parent?.id ? `https://app.box.com/folder/${item.parent.id}` : '',
      sharedLink: item.shared_link?.url || null,
      webUrl: item.type === 'folder' ? `https://app.box.com/folder/${item.id}` : `https://app.box.com/file/${item.id}`
    }));

  } catch (e) {
    Logger.log(`Box search error: ${e.message}`);
    return [];
  }
}

/**
 * Search Box with fuzzy matching for vendor name
 * Tries multiple search strategies to find relevant documents
 *
 * @param {string} vendorName - The vendor name to search for
 * @returns {array} Array of matching documents
 */
function searchBoxForVendor(vendorName) {
  if (!vendorName || vendorName.trim() === '') {
    Logger.log('No vendor name provided');
    return [];
  }

  const cleanName = vendorName.trim();

  // Remove common suffixes for alternate search
  const nameWithoutSuffix = cleanName
    .replace(/,?\s*(LLC|Inc\.?|Corp\.?|Corporation|Company|Co\.|L\.?L\.?C\.?)$/i, '')
    .trim();

  Logger.log(`Searching Box for exact phrase: "${cleanName}"`);

  let results = [];

  // Strategy 1: Exact phrase search (quoted) - primary name
  results = searchBox(`"${cleanName}"`);

  if (results.length > 0) {
    Logger.log(`Found ${results.length} results with exact search for "${cleanName}"`);
    return results;
  }

  // Strategy 2: Try without suffix (e.g., "Ion Solar" instead of "Ion Solar, LLC")
  if (nameWithoutSuffix !== cleanName) {
    Logger.log(`Trying exact search without suffix: "${nameWithoutSuffix}"`);
    results = searchBox(`"${nameWithoutSuffix}"`);

    if (results.length > 0) {
      Logger.log(`Found ${results.length} results with exact search for "${nameWithoutSuffix}"`);
      return results;
    }
  }

  Logger.log('No Box documents found for vendor');
  return [];
}

/**
 * Get current user info (useful for testing)
 */
function getBoxCurrentUser() {
  return boxApiRequest_('/users/me');
}


/************************************************************
 * TEST FUNCTIONS
 ************************************************************/

/**
 * Test Box connection and display user info
 */
function testBoxConnection() {
  try {
    const user = getBoxCurrentUser();

    const msg = `✅ Box Connected!\n\n` +
      `User: ${user.name}\n` +
      `Email: ${user.login}\n` +
      `Space Used: ${(user.space_used / 1024 / 1024 / 1024).toFixed(2)} GB\n` +
      `Space Total: ${(user.space_amount / 1024 / 1024 / 1024).toFixed(2)} GB`;

    Logger.log(msg);
    SpreadsheetApp.getUi().alert(msg);

    return user;

  } catch (e) {
    const msg = `❌ Box Connection Failed\n\n${e.message}\n\nRun authorizeBox() to authorize.`;
    Logger.log(msg);
    SpreadsheetApp.getUi().alert(msg);
    return null;
  }
}

/**
 * Test search for a specific vendor
 * Run this with different vendor names from your list
 */
function testSearchVendor() {
  // Change this to test different vendors
  const testVendor = 'EnergyPal';  // Or any vendor from your List

  Logger.log(`\n========== TESTING BOX SEARCH ==========`);
  Logger.log(`Vendor: ${testVendor}`);
  Logger.log(`=========================================\n`);

  const results = searchBoxForVendor(testVendor);

  if (results.length === 0) {
    Logger.log('No documents found.');
    return;
  }

  Logger.log(`Found ${results.length} document(s):\n`);

  results.forEach((doc, i) => {
    Logger.log(`${i + 1}. ${doc.name}`);
    Logger.log(`   Type: ${doc.type}`);
    Logger.log(`   Path: ${doc.folderPath}`);
    Logger.log(`   Modified: ${doc.modifiedAt}`);
    Logger.log(`   URL: ${doc.webUrl}`);
    Logger.log('');
  });

  return results;
}

/**
 * Test with multiple vendors from your list
 * Pulls vendors from the List sheet and tests each
 */
function testMultipleVendors() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const listSheet = ss.getSheetByName('List');

  if (!listSheet) {
    Logger.log('List sheet not found');
    return;
  }

  // Get first 10 vendors for testing
  const vendors = listSheet.getRange('A2:A11').getValues()
    .flat()
    .filter(v => v && v.toString().trim() !== '');

  Logger.log(`\n========== TESTING ${vendors.length} VENDORS ==========\n`);

  const summary = [];

  vendors.forEach(vendor => {
    const results = searchBoxForVendor(vendor);
    summary.push({
      vendor: vendor,
      found: results.length,
      firstDoc: results[0]?.name || '-'
    });

    Logger.log(`${vendor}: ${results.length} document(s) found`);

    // Rate limiting - Box has 10 req/sec limit
    Utilities.sleep(200);
  });

  Logger.log('\n========== SUMMARY ==========');
  summary.forEach(s => {
    Logger.log(`${s.vendor}: ${s.found} docs${s.found > 0 ? ` (first: ${s.firstDoc})` : ''}`);
  });

  return summary;
}


/************************************************************
 * BATTLE STATION INTEGRATION
 *
 * These functions are designed to integrate with Battle Station
 * Once tested, add calls to these from your main BattleStation.gs
 ************************************************************/

/**
 * Get Box documents for a vendor - formatted for Battle Station display
 *
 * @param {string} vendorName - The vendor name
 * @returns {object} Object with documents array and formatted HTML
 */
function getBoxDocumentsForBattleStation(vendorName) {
  const docs = searchBoxForVendor(vendorName);

  return {
    documents: docs,
    count: docs.length,
    hasDocuments: docs.length > 0,
    // Pre-formatted for display in Battle Station
    displayRows: docs.map(doc => ({
      name: doc.name,
      folder: doc.parentFolder || doc.folderPath.split('/').pop() || 'Root',
      modified: doc.modifiedAt ? new Date(doc.modifiedAt).toLocaleDateString() : '',
      url: doc.webUrl
    }))
  };
}


/************************************************************
 * BATCH BOX OPERATIONS
 *
 * For Turbo Traverse - fetch all documents once, filter locally
 ************************************************************/

/**
 * List all items in a Box folder (non-recursive)
 * @param {string} folderId - The folder ID (use '0' for root)
 * @param {number} limit - Max items to return per page
 * @returns {array} Array of items in the folder
 */
function listBoxFolder_(folderId, limit = 1000) {
  const items = [];
  let offset = 0;
  let hasMore = true;

  while (hasMore) {
    const url = `/folders/${folderId}/items?limit=${limit}&offset=${offset}&fields=id,name,type,description,created_at,modified_at,size,parent,path_collection,shared_link`;

    try {
      const result = boxApiRequest_(url);

      for (const item of result.entries) {
        items.push({
          id: item.id,
          name: item.name,
          type: item.type,
          description: item.description || '',
          createdAt: item.created_at,
          modifiedAt: item.modified_at,
          size: item.size,
          folderPath: item.path_collection?.entries?.map(e => e.name).join('/') || '',
          parentFolder: item.parent?.name || '',
          parentFolderId: item.parent?.id || '',
          parentFolderUrl: item.parent?.id ? `https://app.box.com/folder/${item.parent.id}` : '',
          sharedLink: item.shared_link?.url || null,
          webUrl: item.type === 'folder' ? `https://app.box.com/folder/${item.id}` : `https://app.box.com/file/${item.id}`
        });
      }

      offset += result.entries.length;
      hasMore = result.entries.length === limit && offset < result.total_count;

    } catch (e) {
      Logger.log(`Error listing Box folder ${folderId}: ${e.message}`);
      hasMore = false;
    }
  }

  return items;
}

/**
 * Recursively list all files in Box from root folder
 * @param {string} folderId - Starting folder ID (default '0' for root)
 * @param {number} maxDepth - Maximum folder depth to traverse
 * @param {number} currentDepth - Current depth (internal)
 * @returns {array} All files found
 */
function listAllBoxFiles_(folderId = '0', maxDepth = 5, currentDepth = 0) {
  if (currentDepth >= maxDepth) {
    return [];
  }

  const items = listBoxFolder_(folderId);
  const allFiles = [];
  const subfolders = [];

  for (const item of items) {
    if (item.type === 'file') {
      allFiles.push(item);
    } else if (item.type === 'folder') {
      // Skip certain folders that are unlikely to have vendor docs
      const skipFolders = ['Trash', 'Archive', 'Old', 'Backup'];
      if (!skipFolders.some(skip => item.name.toLowerCase().includes(skip.toLowerCase()))) {
        subfolders.push(item);
      }
    }
  }

  // Recursively get files from subfolders
  for (const folder of subfolders) {
    const subFiles = listAllBoxFiles_(folder.id, maxDepth, currentDepth + 1);
    allFiles.push(...subFiles);
  }

  return allFiles;
}

/**
 * Get all Box documents for batch processing (with caching)
 * Fetches all files from Box once and caches for the session
 * @returns {array} All Box files
 */
function getAllBoxDocuments_() {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'box_all_documents';

  // Check script-level cache first (faster than sheet cache)
  const cached = cache.get(cacheKey);
  if (cached) {
    try {
      const docs = JSON.parse(cached);
      Logger.log(`Box batch: loaded ${docs.length} documents from cache`);
      return docs;
    } catch (e) {
      // Cache corrupted, continue to fetch
    }
  }

  Logger.log('Box batch: fetching all documents from Box...');

  const boxService = getBoxService_();
  if (!boxService.hasAccess()) {
    Logger.log('Box not authorized - cannot batch fetch');
    return [];
  }

  const startTime = Date.now();
  const allDocs = listAllBoxFiles_('0', 4); // Max 4 levels deep
  const elapsed = Math.round((Date.now() - startTime) / 1000);

  Logger.log(`Box batch: fetched ${allDocs.length} documents in ${elapsed}s`);

  // Cache for 30 minutes (1800 seconds) - CacheService max is 6 hours
  // Split into chunks if too large (100KB limit per key)
  try {
    const jsonStr = JSON.stringify(allDocs);
    if (jsonStr.length < 90000) {
      cache.put(cacheKey, jsonStr, 1800);
    } else {
      // Too large for single cache entry - store count only and skip caching
      Logger.log(`Box batch: ${allDocs.length} docs too large for cache (${Math.round(jsonStr.length/1024)}KB)`);
    }
  } catch (e) {
    Logger.log(`Box batch cache error: ${e.message}`);
  }

  return allDocs;
}

/**
 * Filter Box documents for a specific vendor (local matching)
 * @param {array} allDocs - All Box documents (from getAllBoxDocuments_)
 * @param {string} vendorName - Vendor name to search for
 * @param {string} otherName - Optional alternate vendor name
 * @returns {array} Matching documents
 */
function filterBoxDocumentsForVendor_(allDocs, vendorName, otherName) {
  if (!allDocs || allDocs.length === 0 || !vendorName) {
    return [];
  }

  const matches = [];
  const seenIds = new Set();

  // Build search terms
  const searchTerms = [vendorName.toLowerCase().trim()];

  // Add version without suffix
  const withoutSuffix = vendorName
    .replace(/,?\s*(LLC|Inc\.?|Corp\.?|Corporation|Company|Co\.|L\.?L\.?C\.?)$/i, '')
    .trim().toLowerCase();
  if (withoutSuffix !== searchTerms[0]) {
    searchTerms.push(withoutSuffix);
  }

  // Add other name variants
  if (otherName) {
    const otherNames = otherName.split(',').map(n => n.trim().toLowerCase()).filter(n => n);
    searchTerms.push(...otherNames);

    // Add without suffix for each
    for (const name of otherNames) {
      const clean = name.replace(/,?\s*(LLC|Inc\.?|Corp\.?|Corporation|Company|Co\.|L\.?L\.?C\.?)$/i, '').trim();
      if (clean !== name && !searchTerms.includes(clean)) {
        searchTerms.push(clean);
      }
    }
  }

  for (const doc of allDocs) {
    if (seenIds.has(doc.id)) continue;

    // Check file name and folder path
    const searchText = (doc.name + ' ' + doc.folderPath + ' ' + doc.parentFolder).toLowerCase();

    for (const term of searchTerms) {
      if (searchText.includes(term)) {
        doc.matchedOn = term;
        matches.push(doc);
        seenIds.add(doc.id);
        break;
      }
    }
  }

  return matches;
}

/**
 * Clear the Box batch cache
 */
function clearBoxBatchCache() {
  const cache = CacheService.getScriptCache();
  cache.remove('box_all_documents');
  Logger.log('Box batch cache cleared');
}
