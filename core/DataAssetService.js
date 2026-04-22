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
    return { success: true, users: users, projects: projects };
  } catch (error) {
    console.error('getDataAssetFormOptions failed:', error);
    return { success: false, error: error.message, users: [], projects: [] };
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
