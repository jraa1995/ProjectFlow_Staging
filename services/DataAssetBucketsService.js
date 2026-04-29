var DataAssetBucketsService = (function() {

  function _sanitizeName(value) {
    return sanitize(String(value || '').trim());
  }

  function _summarize(bucket, members) {
    var count = members.length;
    var typeCounts = {};
    members.forEach(function(m) {
      var t = m.assetType || 'Untyped';
      typeCounts[t] = (typeCounts[t] || 0) + 1;
    });
    return {
      id: bucket.id,
      name: bucket.name,
      description: bucket.description || '',
      ownerEmail: bucket.ownerEmail || '',
      primaryStakeholder: bucket.primaryStakeholder || '',
      createdBy: bucket.createdBy || '',
      createdAt: bucket.createdAt,
      updatedAt: bucket.updatedAt,
      memberCount: count,
      memberTypeCounts: typeCounts
    };
  }

  function _membersForBucket(bucketId, allAssets) {
    var src = allAssets || getAllDataAssetsOptimized();
    return src.filter(function(a) { return a.bucketId === bucketId; });
  }

  function listBuckets() {
    PermissionGuard.requirePermission('dataasset:read');
    var buckets = getAllDataAssetBuckets();
    if (!buckets.length) return [];
    var allAssets = getAllDataAssetsOptimized();
    return buckets.map(function(b) { return _summarize(b, _membersForBucket(b.id, allAssets)); });
  }

  function getBucketDetails(bucketId) {
    PermissionGuard.requirePermission('dataasset:read');
    if (!bucketId) throw new Error('bucketId is required');
    var bucket = getDataAssetBucketById(bucketId);
    if (!bucket) throw new Error('Bucket not found: ' + bucketId);
    var allAssets = getAllDataAssetsOptimized();
    var members = _membersForBucket(bucketId, allAssets);
    var summary = _summarize(bucket, members);
    summary.members = members.map(function(m) {
      return {
        id: m.id,
        assetName: m.assetName,
        assetType: m.assetType || '',
        status: m.status,
        assetOwner: m.assetOwner || '',
        relatedProjects: m.relatedProjects || ''
      };
    });
    return summary;
  }

  function createBucket(payload) {
    PermissionGuard.requirePermission('dataasset:create');
    if (!payload || !payload.name || !String(payload.name).trim()) {
      throw new Error('Bucket name is required');
    }
    var name = _sanitizeName(payload.name);
    var existing = getAllDataAssetBuckets();
    var dup = existing.find(function(b) { return (b.name || '').toLowerCase() === name.toLowerCase(); });
    if (dup) throw new Error('A bucket with this name already exists');
    var sheet = getDataAssetBucketsSheet();
    var columns = CONFIG.DATA_ASSET_BUCKET_COLUMNS;
    var currentUser = getCurrentUserEmail();
    var bucket = {
      id: generateId('DAB'),
      name: name,
      description: sanitize(payload.description || ''),
      ownerEmail: sanitize(payload.ownerEmail || currentUser),
      primaryStakeholder: sanitize(payload.primaryStakeholder || ''),
      createdBy: currentUser,
      createdAt: now(),
      updatedAt: now(),
      jsonData: payload.jsonData || ''
    };
    sheet.appendRow(objectToRow(bucket, columns));
    SpreadsheetApp.flush();
    logActivity(currentUser, 'created', 'asset_bucket', bucket.id, { name: bucket.name });
    return _summarize(bucket, []);
  }

  function updateBucket(payload) {
    PermissionGuard.requirePermission('dataasset:update');
    if (!payload || !payload.id) throw new Error('Bucket id is required');
    var sheet = getDataAssetBucketsSheet();
    var data = sheet.getDataRange().getValues();
    var columns = CONFIG.DATA_ASSET_BUCKET_COLUMNS;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === payload.id) {
        var existing = rowToObject(data[i], columns);
        if (Object.prototype.hasOwnProperty.call(payload, 'name')) {
          var trimmed = _sanitizeName(payload.name);
          if (!trimmed) throw new Error('Bucket name cannot be empty');
          if (trimmed.toLowerCase() !== (existing.name || '').toLowerCase()) {
            var dup = getAllDataAssetBuckets().find(function(b) {
              return b.id !== existing.id && (b.name || '').toLowerCase() === trimmed.toLowerCase();
            });
            if (dup) throw new Error('A bucket with this name already exists');
          }
          existing.name = trimmed;
        }
        if (Object.prototype.hasOwnProperty.call(payload, 'description')) {
          existing.description = sanitize(payload.description || '');
        }
        if (Object.prototype.hasOwnProperty.call(payload, 'ownerEmail')) {
          existing.ownerEmail = sanitize(payload.ownerEmail || '');
        }
        if (Object.prototype.hasOwnProperty.call(payload, 'primaryStakeholder')) {
          existing.primaryStakeholder = sanitize(payload.primaryStakeholder || '');
        }
        existing.updatedAt = now();
        sheet.getRange(i + 1, 1, 1, columns.length).setValues([objectToRow(existing, columns)]);
        SpreadsheetApp.flush();
        logActivity(getCurrentUserEmail(), 'updated', 'asset_bucket', existing.id, { name: existing.name });
        var members = _membersForBucket(existing.id);
        return _summarize(existing, members);
      }
    }
    throw new Error('Bucket not found: ' + payload.id);
  }

  function deleteBucket(bucketId) {
    PermissionGuard.requirePermission('dataasset:delete');
    if (!bucketId) throw new Error('Bucket id is required');
    var lock = LockService.getScriptLock();
    lock.waitLock(5000);
    try {
      var bucketSheet = getDataAssetBucketsSheet();
      var bucketData = bucketSheet.getDataRange().getValues();
      var bucketRow = -1;
      var bucketName = '';
      for (var i = 1; i < bucketData.length; i++) {
        if (bucketData[i][0] === bucketId) {
          bucketRow = i + 1;
          bucketName = bucketData[i][1];
          break;
        }
      }
      if (bucketRow === -1) throw new Error('Bucket not found: ' + bucketId);

      var assetSheet = getDataAssetsSheet();
      var assetData = assetSheet.getDataRange().getValues();
      var assetCols = CONFIG.DATA_ASSET_COLUMNS;
      var bucketIdx = assetCols.indexOf('bucketId');
      var clearedAssets = [];
      if (bucketIdx !== -1) {
        var timestamp = now();
        var currentUser = getCurrentUserEmail();
        var updatedAtIdx = assetCols.indexOf('updatedAt');
        var lastUpdatedByIdx = assetCols.indexOf('lastUpdatedBy');
        for (var j = 1; j < assetData.length; j++) {
          if (assetData[j][bucketIdx] === bucketId) {
            assetData[j][bucketIdx] = '';
            if (updatedAtIdx !== -1) assetData[j][updatedAtIdx] = timestamp;
            if (lastUpdatedByIdx !== -1) assetData[j][lastUpdatedByIdx] = currentUser;
            assetSheet.getRange(j + 1, 1, 1, assetCols.length).setValues([assetData[j]]);
            clearedAssets.push(String(assetData[j][0]));
          }
        }
      }
      bucketSheet.deleteRow(bucketRow);
      SpreadsheetApp.flush();
      invalidateDataAssetCache();
      logActivity(getCurrentUserEmail(), 'deleted', 'asset_bucket', bucketId, { name: bucketName, clearedMembers: clearedAssets.length });
      return { id: bucketId, clearedAssets: clearedAssets };
    } finally {
      lock.releaseLock();
    }
  }

  function setAssetsBucket(payload) {
    PermissionGuard.requirePermission('dataasset:update');
    if (!payload) throw new Error('payload is required');
    var assetIds = Array.isArray(payload.assetIds) ? payload.assetIds.filter(Boolean) : [];
    if (!assetIds.length) throw new Error('At least one asset id is required');
    var rawBucketId = sanitize(payload.bucketId || '');
    if (rawBucketId && !getDataAssetBucketById(rawBucketId)) {
      throw new Error('Bucket not found: ' + rawBucketId);
    }
    var requested = {};
    assetIds.forEach(function(id) { requested[id] = true; });
    var lock = LockService.getScriptLock();
    lock.waitLock(5000);
    try {
      var sheet = getDataAssetsSheet();
      var data = sheet.getDataRange().getValues();
      var columns = CONFIG.DATA_ASSET_COLUMNS;
      var bucketIdx = columns.indexOf('bucketId');
      if (bucketIdx === -1) throw new Error('bucketId column missing');
      var updatedAtIdx = columns.indexOf('updatedAt');
      var lastUpdatedByIdx = columns.indexOf('lastUpdatedBy');
      var timestamp = now();
      var currentUser = getCurrentUserEmail();
      var updatedIds = [];
      var found = {};
      for (var i = 1; i < data.length; i++) {
        var rowId = data[i][0];
        if (!requested[rowId]) continue;
        found[rowId] = true;
        if (data[i][bucketIdx] === rawBucketId) continue;
        data[i][bucketIdx] = rawBucketId;
        if (updatedAtIdx !== -1) data[i][updatedAtIdx] = timestamp;
        if (lastUpdatedByIdx !== -1) data[i][lastUpdatedByIdx] = currentUser;
        sheet.getRange(i + 1, 1, 1, columns.length).setValues([data[i]]);
        updatedIds.push(rowId);
      }
      var notFound = assetIds.filter(function(id) { return !found[id]; });
      SpreadsheetApp.flush();
      updatedIds.forEach(function(id) {
        logActivity(currentUser, rawBucketId ? 'bucket_assigned' : 'bucket_cleared', 'dataasset', id, { bucketId: rawBucketId });
      });
      invalidateDataAssetCache();
      return { updated: updatedIds.length, updatedIds: updatedIds, notFound: notFound, bucketId: rawBucketId };
    } finally {
      lock.releaseLock();
    }
  }

  return {
    listBuckets: listBuckets,
    getBucketDetails: getBucketDetails,
    createBucket: createBucket,
    updateBucket: updateBucket,
    deleteBucket: deleteBucket,
    setAssetsBucket: setAssetsBucket
  };
})();

function getDataAssetBuckets() {
  try {
    return { success: true, buckets: DataAssetBucketsService.listBuckets() };
  } catch (error) {
    console.error('getDataAssetBuckets failed:', error);
    return { success: false, error: error.message, buckets: [] };
  }
}

function getDataAssetBucketDetails(bucketId) {
  try {
    return { success: true, bucket: DataAssetBucketsService.getBucketDetails(bucketId) };
  } catch (error) {
    console.error('getDataAssetBucketDetails failed:', error);
    return { success: false, error: error.message };
  }
}

function saveDataAssetBucket(payload) {
  try {
    return { success: true, bucket: DataAssetBucketsService.createBucket(payload) };
  } catch (error) {
    console.error('saveDataAssetBucket failed:', error);
    return { success: false, error: error.message };
  }
}

function updateDataAssetBucket(payload) {
  try {
    return { success: true, bucket: DataAssetBucketsService.updateBucket(payload) };
  } catch (error) {
    console.error('updateDataAssetBucket failed:', error);
    return { success: false, error: error.message };
  }
}

function deleteDataAssetBucket(bucketId) {
  try {
    var result = DataAssetBucketsService.deleteBucket(bucketId);
    return { success: true, id: result.id, clearedAssets: result.clearedAssets };
  } catch (error) {
    console.error('deleteDataAssetBucket failed:', error);
    return { success: false, error: error.message };
  }
}

function setDataAssetsBucket(payload) {
  try {
    var result = DataAssetBucketsService.setAssetsBucket(payload);
    return { success: true, updated: result.updated, updatedIds: result.updatedIds, notFound: result.notFound, bucketId: result.bucketId };
  } catch (error) {
    console.error('setDataAssetsBucket failed:', error);
    return { success: false, error: error.message };
  }
}
