function getAllDataAssetsOptimized() {
  var cacheKey = 'ALL_DATA_ASSETS_CACHE';
  try {
    var cached = CacheService.getScriptCache().get(cacheKey);
    if (cached) return JSON.parse(cached);
  } catch (e) {}
  var assets = getAllDataAssets();
  try {
    CacheService.getScriptCache().put(cacheKey, JSON.stringify(assets), 900);
  } catch (e) {}
  return assets;
}

function getDataAssetDetailsList() {
  try {
    var assets = getAllDataAssetsOptimized();
    return { success: true, assets: assets };
  } catch (error) {
    console.error('getDataAssetDetailsList failed:', error);
    return { success: false, error: error.message, assets: [] };
  }
}

function getDataAssetDetails(assetId) {
  try {
    var asset = getDataAssetById(assetId);
    if (!asset) return { success: false, error: 'Data asset not found' };
    return { success: true, asset: asset };
  } catch (error) {
    console.error('getDataAssetDetails failed:', error);
    return { success: false, error: error.message };
  }
}

function saveNewDataAsset(assetData) {
  try {
    var asset = createDataAsset(assetData);
    invalidateDataAssetCache();
    return { success: true, asset: asset };
  } catch (error) {
    console.error('saveNewDataAsset failed:', error);
    return { success: false, error: error.message };
  }
}

function updateDataAssetDetails(payload) {
  try {
    var assetId = payload.id;
    delete payload.id;
    var asset = updateDataAsset(assetId, payload);
    invalidateDataAssetCache();
    return { success: true, asset: asset };
  } catch (error) {
    console.error('updateDataAssetDetails failed:', error);
    return { success: false, error: error.message };
  }
}

function deleteExistingDataAsset(assetId) {
  try {
    deleteDataAsset(assetId);
    invalidateDataAssetCache();
    return { success: true };
  } catch (error) {
    console.error('deleteExistingDataAsset failed:', error);
    return { success: false, error: error.message };
  }
}

function invalidateDataAssetCache() {
  try {
    CacheService.getScriptCache().remove('ALL_DATA_ASSETS_CACHE');
    invalidateCache('dataAsset', null, 'update');
  } catch (e) {}
}

function getDataAssetFormOptions() {
  try {
    var users = getActiveUsersOptimized()
      .filter(function(u) { return u.email; })
      .map(function(u) { return { email: u.email, name: u.name || u.email.split('@')[0], role: u.role || '' }; });
    var projects = getAllProjectsOptimized()
      .map(function(p) { return { id: p.id, name: p.name || '' }; });
    var assetTypes = (CONFIG && CONFIG.DATA_ASSET_TYPES) ? CONFIG.DATA_ASSET_TYPES.slice() : [];
    var buckets = [];
    try {
      buckets = (typeof getAllDataAssetBuckets === 'function')
        ? getAllDataAssetBuckets().map(function(b) { return { id: b.id, name: b.name }; })
        : [];
    } catch (e) {
      console.error('getDataAssetFormOptions: bucket lookup failed:', e);
    }
    return { success: true, users: users, projects: projects, assetTypes: assetTypes, buckets: buckets };
  } catch (error) {
    console.error('getDataAssetFormOptions failed:', error);
    return { success: false, error: error.message, users: [], projects: [], assetTypes: [], buckets: [] };
  }
}

function importWorkLogDataAssets(workbookId) {
  try {
    var result = importDataAssetsFromWorkLog(workbookId);
    return result;
  } catch (error) {
    console.error('importWorkLogDataAssets failed:', error);
    return { success: false, error: error.message };
  }
}

function getStoredDataAssetsWorkbookId() {
  try {
    return { success: true, workbookId: getDataAssetsWorkbookId() };
  } catch (error) {
    console.error('getStoredDataAssetsWorkbookId failed:', error);
    return { success: false, workbookId: '' };
  }
}

function saveDataAssetsWorkbookId(workbookId) {
  try {
    PermissionGuard.requirePermission('admin:settings');
    setDataAssetsWorkbookId(workbookId);
    return { success: true };
  } catch (error) {
    console.error('saveDataAssetsWorkbookId failed:', error);
    return { success: false, error: error.message };
  }
}

function previewDataAssetSource(sourceUrl) {
  try {
    PermissionGuard.requirePermission('dataasset:read');
    var fileId = _extractSheetIdFromInput_(sourceUrl);
    if (!fileId) {
      return { success: false, error: 'No Google Sheet URL detected. Only Google Sheets can be previewed.' };
    }
    var sheetUrl = 'https://docs.google.com/spreadsheets/d/' + fileId + '/edit';
    var ss;
    try {
      ss = SpreadsheetApp.openById(fileId);
    } catch (e) {
      var msg = String((e && e.message) || e);
      var serviceEmail = '';
      try { serviceEmail = Session.getEffectiveUser().getEmail() || ''; } catch (err) {}
      return {
        success: false,
        code: 'NO_ACCESS',
        error: msg,
        serviceEmail: serviceEmail,
        fileUrl: sheetUrl,
        fileId: fileId
      };
    }
    var fileName = '';
    var fileUrl = '';
    try { fileName = ss.getName(); } catch (e) {}
    try { fileUrl = ss.getUrl(); } catch (e) {}
    var allSheets = [];
    try { allSheets = ss.getSheets(); } catch (e) {}
    var SHEET_CAP = 20;
    var COL_CAP = 30;
    var ROW_CAP = 5;
    var truncated = allSheets.length > SHEET_CAP;
    var sheetsOut = [];
    for (var i = 0; i < Math.min(allSheets.length, SHEET_CAP); i++) {
      var sheet = allSheets[i];
      var lastRow = 0;
      var lastCol = 0;
      try { lastRow = sheet.getLastRow(); } catch (e) {}
      try { lastCol = sheet.getLastColumn(); } catch (e) {}
      if (lastRow === 0 || lastCol === 0) continue;
      var maxCols = Math.min(lastCol, COL_CAP);
      var headers = [];
      var rows = [];
      try {
        var headerVals = sheet.getRange(1, 1, 1, maxCols).getDisplayValues();
        headers = headerVals[0] || [];
      } catch (e) {}
      var rowsAvailable = Math.max(0, lastRow - 1);
      var rowsToRead = Math.min(ROW_CAP, rowsAvailable);
      if (rowsToRead > 0) {
        try { rows = sheet.getRange(2, 1, rowsToRead, maxCols).getDisplayValues(); } catch (e) {}
      }
      sheetsOut.push({
        name: sheet.getName(),
        gid: sheet.getSheetId(),
        rowCount: lastRow,
        colCount: lastCol,
        truncatedCols: lastCol > COL_CAP,
        headers: headers,
        rows: rows
      });
    }
    return {
      success: true,
      fileId: fileId,
      fileName: fileName,
      fileUrl: fileUrl,
      sheetsTotal: allSheets.length,
      sheetsTruncated: truncated,
      sheets: sheetsOut
    };
  } catch (error) {
    console.error('previewDataAssetSource failed:', error);
    return { success: false, error: error.message };
  }
}

function _extractSheetIdFromInput_(input) {
  if (!input) return '';
  var s = String(input);
  var m = s.match(/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (m) return m[1];
  if (/^[a-zA-Z0-9-_]{20,}$/.test(s.trim())) return s.trim();
  return '';
}
