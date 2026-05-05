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

function _da_filterAssetsForCurrentUser_(assets) {
  try {
    if (!Array.isArray(assets)) return [];
    var role = (typeof getCurrentUserRole === 'function') ? getCurrentUserRole() : null;
    if (role === 'admin' || role === 'manager') return assets;
    if (typeof getManagerContractFromRole === 'function' && getManagerContractFromRole(role)) {
      var mgrAllowed = getManagerAccessibleProjectIds(role);
      return assets.filter(function(a) {
        if (!a) return false;
        var ids = String(a.relatedProjects || '').split(/[,;]/).map(function(s) { return s.trim(); }).filter(Boolean);
        if (ids.length === 0) return false;
        for (var i = 0; i < ids.length; i++) { if (mgrAllowed[ids[i]]) return true; }
        return false;
      });
    }
    if (role === 'client') {
      var email = (typeof getCurrentUserEmailOptimized === 'function') ? getCurrentUserEmailOptimized() : null;
      if (!email) return [];
      var clientAllowed = getClientAccessibleProjectIds(email);
      return assets.filter(function(a) {
        if (!a) return false;
        var ids = String(a.relatedProjects || '').split(/[,;]/).map(function(s) { return s.trim(); }).filter(Boolean);
        if (ids.length === 0) return false;
        for (var i = 0; i < ids.length; i++) { if (clientAllowed[ids[i]]) return true; }
        return false;
      });
    }
    return assets;
  } catch (e) {
    console.error('_da_filterAssetsForCurrentUser_ failed:', e);
    return [];
  }
}

function _da_authorizeAssetAccess_(asset) {
  if (!asset) return false;
  try {
    var role = (typeof getCurrentUserRole === 'function') ? getCurrentUserRole() : null;
    if (role === 'admin' || role === 'manager') return true;
    var ids = String(asset.relatedProjects || '').split(/[,;]/).map(function(s) { return s.trim(); }).filter(Boolean);
    if (typeof getManagerContractFromRole === 'function' && getManagerContractFromRole(role)) {
      var mgrAllowed = getManagerAccessibleProjectIds(role);
      for (var i = 0; i < ids.length; i++) { if (mgrAllowed[ids[i]]) return true; }
      return false;
    }
    if (role === 'client') {
      var email = (typeof getCurrentUserEmailOptimized === 'function') ? getCurrentUserEmailOptimized() : null;
      if (!email) return false;
      var clientAllowed = getClientAccessibleProjectIds(email);
      for (var j = 0; j < ids.length; j++) { if (clientAllowed[ids[j]]) return true; }
      return false;
    }
    return true;
  } catch (e) {
    console.error('_da_authorizeAssetAccess_ failed:', e);
    return false;
  }
}

function _da_safePayload_(payload) {
  if (!payload || typeof payload !== 'object') return {};
  var safe = {};
  Object.keys(payload).slice(0, 30).forEach(function(k) {
    var v = payload[k];
    if (v === null || v === undefined) { safe[k] = v; return; }
    if (typeof v === 'string') { safe[k] = v.length > 200 ? v.slice(0, 200) + '...' : v; }
    else if (typeof v === 'number' || typeof v === 'boolean') { safe[k] = v; }
    else { safe[k] = '[object]'; }
  });
  return safe;
}

function getDataAssetDetailsList() {
  try {
    var assets = getAllDataAssetsOptimized();
    var filtered = _da_filterAssetsForCurrentUser_(assets);
    return { success: true, assets: filtered };
  } catch (error) {
    console.error('getDataAssetDetailsList failed:', error);
    return { success: false, error: error.message, assets: [] };
  }
}

function getDataAssetDetails(assetId) {
  try {
    var asset = getDataAssetById(assetId);
    if (!asset) return { success: false, error: 'Data asset not found' };
    if (!_da_authorizeAssetAccess_(asset)) {
      return { success: false, error: 'Data asset not found' };
    }
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
    try {
      if (typeof logAuditEvent === 'function') {
        logAuditEvent('dataasset.create', 'dataAsset', asset.id, asset.name || '', { assetType: asset.assetType });
      }
    } catch (e) {}
    return { success: true, asset: asset };
  } catch (error) {
    console.error('saveNewDataAsset failed:', error);
    return { success: false, error: error.message };
  }
}

function updateDataAssetDetails(payload) {
  try {
    var assetId = payload.id;
    var _existingAsset = null;
    try { _existingAsset = getDataAssetById(assetId); } catch (e) {}
    if (!_existingAsset) return { success: false, error: 'Data asset not found' };
    if (!_da_authorizeAssetAccess_(_existingAsset)) {
      return { success: false, error: 'Permission denied' };
    }
    delete payload.id;
    var asset = updateDataAsset(assetId, payload);
    invalidateDataAssetCache();
    try {
      if (typeof logAuditEvent === 'function') {
        logAuditEvent('dataasset.update', 'dataAsset', assetId, (asset && asset.name) || '', _da_safePayload_(payload));
      }
    } catch (e) {}
    return { success: true, asset: asset };
  } catch (error) {
    console.error('updateDataAssetDetails failed:', error);
    return { success: false, error: error.message };
  }
}

function deleteExistingDataAsset(assetId) {
  try {
    var _existingAsset = null;
    try { _existingAsset = getDataAssetById(assetId); } catch (e) {}
    if (!_existingAsset) return { success: false, error: 'Data asset not found' };
    if (!_da_authorizeAssetAccess_(_existingAsset)) {
      return { success: false, error: 'Permission denied' };
    }
    deleteDataAsset(assetId);
    invalidateDataAssetCache();
    try {
      if (typeof logAuditEvent === 'function') {
        logAuditEvent('dataasset.delete', 'dataAsset', assetId, (_existingAsset && _existingAsset.name) || '', {});
      }
    } catch (e) {}
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

function setProjectDataAssetLinks(projectId, assetIds) {
  try {
    PermissionGuard.requirePermission('dataasset:update');
    if (!projectId) return { success: false, error: 'projectId required' };
    var desired = {};
    (assetIds || []).forEach(function(id) { if (id) desired[String(id)] = true; });
    var assets = getAllDataAssets();
    var changed = 0;
    for (var i = 0; i < assets.length; i++) {
      var a = assets[i];
      if (!a || !a.id) continue;
      var ids = String(a.relatedProjects || '').split(/[,;]/).map(function(s) { return s.trim(); }).filter(Boolean);
      var hasLink = ids.indexOf(projectId) !== -1;
      var shouldHaveLink = !!desired[a.id];
      if (hasLink === shouldHaveLink) continue;
      var nextIds = shouldHaveLink ? ids.concat([projectId]) : ids.filter(function(x) { return x !== projectId; });
      var seen = {};
      var dedup = [];
      for (var j = 0; j < nextIds.length; j++) {
        var v = nextIds[j];
        if (v && !seen[v]) { seen[v] = true; dedup.push(v); }
      }
      updateDataAsset(a.id, { relatedProjects: dedup.join(',') });
      changed++;
    }
    if (changed > 0) invalidateDataAssetCache();
    var refreshed = changed > 0 ? getAllDataAssetsOptimized() : assets;
    return { success: true, changed: changed, assets: refreshed };
  } catch (error) {
    console.error('setProjectDataAssetLinks failed:', error);
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
    if (!_da_isFileIdRegistered_(fileId)) {
      try {
        if (typeof logAuditEvent === 'function') {
          logAuditEvent('dataasset.preview_denied', 'dataAsset', fileId, '', { reason: 'unregistered_file_id' });
        }
      } catch (e) {}
      return { success: false, error: 'File is not a registered data asset. Preview is restricted to data assets registered in COLONY.' };
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

function _da_isFileIdRegistered_(fileId) {
  try {
    if (!fileId) return false;
    var assets = getAllDataAssetsOptimized();
    for (var i = 0; i < assets.length; i++) {
      var a = assets[i];
      if (!a) continue;
      var sourceUrl = String(a.sourceUrl || a.fileUrl || a.url || a.driveFileId || a.fileId || '');
      if (!sourceUrl) {
        var raw = '';
        try { raw = JSON.stringify(a); } catch (e) { raw = ''; }
        if (raw.indexOf(fileId) !== -1) return true;
        continue;
      }
      var assetFileId = _extractSheetIdFromInput_(sourceUrl);
      if (assetFileId && assetFileId === fileId) return true;
      if (sourceUrl.indexOf(fileId) !== -1) return true;
    }
    return false;
  } catch (e) {
    console.error('_da_isFileIdRegistered_ failed:', e);
    return false;
  }
}
