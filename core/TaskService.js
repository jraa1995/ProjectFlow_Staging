function loadMyBoard(projectId) {
  return getMyBoardOptimized(projectId || null);
}

function loadMasterBoard(projectId) {
  try {
    var userRole = getCurrentUserRole();
    if (userRole !== 'admin' && userRole !== 'manager') {
      throw new Error('Permission denied: Master Board requires admin or manager access.');
    }

    var batchData = getBatchDataFast();
    var filteredTasks = projectId
      ? batchData.tasks.filter(function(task) { return task.projectId === projectId; })
      : batchData.tasks;

    var board = buildBoardData(filteredTasks, projectId, { view: 'master' });

    return {
      ...board,
      projects: batchData.projects,
      users: batchData.users,
      taskCount: filteredTasks.length
    };
  } catch (error) {
    console.error('loadMasterBoard failed:', error);
    throw error;
  }
}

function getMyBoardOptimized(projectId, userEmail) {
  try {
    var currentUser = userEmail || getCurrentUserEmailOptimized();

    if (!currentUser) {
      throw new Error('User not authenticated');
    }

    var batchData = getBatchDataFast();
    var userTasks = batchData.tasks.filter(function(task) {
      return task.assignee && task.assignee.toLowerCase() === currentUser.toLowerCase();
    });

    var filteredTasks = projectId
      ? userTasks.filter(function(task) { return task.projectId === projectId; })
      : userTasks;

    var board = buildBoardData(filteredTasks, projectId, {
      view: 'my',
      userEmail: currentUser
    });

    return {
      ...board,
      projects: batchData.projects,
      users: batchData.users,
      taskCount: filteredTasks.length
    };
  } catch (error) {
    console.error('getMyBoardOptimized failed:', error);
    throw error;
  }
}

function saveNewTask(taskData) {
  try {
    const result = createTask(taskData);
    invalidateTaskCache(result.id, 'create');
    return result;
  } catch (error) {
    console.error('saveNewTask failed:', error);
    throw error;
  }
}

function saveTaskUpdate(taskId, updates) {
  if (isValidTaskUid(taskId)) {
    var resolved = getTaskByUid(taskId);
    if (resolved) taskId = resolved.id;
  }
  const result = updateTask(taskId, updates);
  invalidateTaskCache(taskId, 'update');
  try {
    if (result.dueDate) {
      const currentUser = getCurrentUserEmail();
      if (!result.assignee || result.assignee.toLowerCase() === currentUser.toLowerCase()) {
        syncTaskToCalendar(result);
      }
    }
  } catch (e) {
    console.error('saveTaskUpdate: calendar sync failed for task ' + taskId + ':', e);
  }
  return result;
}

function moveTaskToStatus(taskId, newStatus, newPosition) {
  if (isValidTaskUid(taskId)) {
    var resolved = getTaskByUid(taskId);
    if (resolved) taskId = resolved.id;
  }
  const status = denormalizeStatusId(newStatus);
  const result = moveTask(taskId, status, newPosition);
  invalidateTaskCache(taskId, 'update');
  return result;
}

function removeTask(taskId) {
  if (isValidTaskUid(taskId)) {
    var resolved = getTaskByUid(taskId);
    if (resolved) taskId = resolved.id;
  }
  const result = deleteTask(taskId);
  invalidateTaskCache(taskId, 'delete');
  return result;
}

function loadTask(taskId) {
  try {
    if (!taskId || typeof taskId !== 'string' || taskId.trim() === '') {
      return null;
    }
    var id = taskId.trim();
    if (isValidTaskUid(id)) {
      var resolved = getTaskByUid(id);
      return resolved || null;
    }
    var task = getTaskById(id);
    return task || null;
  } catch (error) {
    console.error('loadTask error for taskId ' + taskId + ':', error);
    return null;
  }
}

function searchAllTasks(query, projectId) {
  return searchTasks(query, projectId || null);
}

function executeBatchOperations(operations) {
  if (!Array.isArray(operations) || operations.length === 0) {
    return { results: [], errors: [] };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var updatesList = [];
    operations.forEach(function(op) {
      var changes = {};
      switch (op.type) {
        case 'moveTask':
          changes = { status: op.params.status, position: op.params.position || 0 };
          break;
        case 'updateTask':
          changes = op.params.changes;
          break;
        case 'updatePriority':
          changes = { priority: op.params.priority };
          break;
        case 'updateAssignee':
          changes = { assignee: op.params.assignee };
          break;
        default:
          break;
      }
      if (op.params.taskId) {
        updatesList.push({ taskId: op.params.taskId, changes: changes });
      }
    });
    var results = batchUpdateTasks(updatesList);
    invalidateTaskCache(null, 'update');
    return {
      results: results.map(function(r, i) { return { index: i, success: true, result: r }; }),
      errors: []
    };
  } catch (e) {
    console.error('executeBatchOperations failed:', e.message);
    return { results: [], errors: [{ index: 0, success: false, error: e.message }] };
  } finally {
    lock.releaseLock();
  }
}

function updateTaskWithVersion(taskId, updates, expectedVersion) {
  try {
    if (typeof LockManager === 'undefined') {
      const task = updateTask(taskId, updates);
      return { success: true, task: task };
    }

    const task = LockManager.updateTaskWithLocking(taskId, updates, expectedVersion);
    invalidateTaskCache(taskId, 'update');
    return { success: true, task: task };
  } catch (error) {
    console.error('updateTaskWithVersion failed:', error);

    if (error.message.includes('VERSION_CONFLICT')) {
      try {
        const conflictData = JSON.parse(error.message);
        return {
          success: false,
          conflict: true,
          code: 'VERSION_CONFLICT',
          currentVersion: conflictData.currentVersion,
          expectedVersion: conflictData.expectedVersion,
          lastModifiedBy: conflictData.lastModifiedBy,
          error: 'This task was modified by another user'
        };
      } catch (e) {
        // Parse failed, fall through
      }
    }

    return { success: false, error: error.message };
  }
}

function acquireTaskEditLock(taskId) {
  try {
    if (typeof LockManager === 'undefined') {
      return { success: true, lock: null };
    }

    return LockManager.acquireEditLock(taskId);
  } catch (error) {
    console.error('acquireTaskEditLock failed:', error);
    return { success: false, error: error.message };
  }
}

function auditTaskIntegrity() {
  var columns = CONFIG.TASK_COLUMNS;
  var idCol = columns.indexOf('id');
  var depsCol = columns.indexOf('dependencies');
  var sheet = getTasksSheet();
  var data = sheet.getDataRange().getValues();
  var totalRows = data.length - 1;
  var idMap = {};
  var duplicates = [];
  var allIds = new Set();
  for (var i = 1; i < data.length; i++) {
    var id = data[i][idCol];
    if (!id) continue;
    allIds.add(id);
    if (idMap[id]) {
      duplicates.push({ id: id, rows: [idMap[id], i + 1] });
    } else {
      idMap[id] = i + 1;
    }
  }
  var invalidDeps = [];
  if (depsCol !== -1) {
    for (var i = 1; i < data.length; i++) {
      var deps = data[i][depsCol];
      if (!deps) continue;
      var depList = String(deps).split(',').map(function(d) { return d.trim(); }).filter(Boolean);
      depList.forEach(function(depId) {
        if (!allIds.has(depId)) {
          invalidDeps.push({ taskId: data[i][idCol], invalidDep: depId, row: i + 1 });
        }
      });
    }
  }
  var orphanComments = [];
  try {
    var commentsSheet = getCommentsSheet();
    var commentData = commentsSheet.getDataRange().getValues();
    var commentColumns = CONFIG.COMMENT_COLUMNS;
    var commentTaskCol = commentColumns.indexOf('taskId');
    if (commentTaskCol !== -1) {
      for (var i = 1; i < commentData.length; i++) {
        var commentTaskId = commentData[i][commentTaskCol];
        if (commentTaskId && !allIds.has(commentTaskId)) {
          orphanComments.push({ commentRow: i + 1, taskId: commentTaskId });
        }
      }
    }
  } catch (e) {
    console.error('auditTaskIntegrity: comments scan failed:', e);
  }
  console.log('=== TASK INTEGRITY AUDIT ===');
  console.log('Tasks scanned: ' + totalRows);
  console.log('Unique IDs: ' + allIds.size);
  console.log('Duplicate IDs: ' + duplicates.length);
  console.log('Invalid dependencies: ' + invalidDeps.length);
  console.log('Orphan comments: ' + orphanComments.length);
  if (duplicates.length > 0) console.log('Duplicates: ' + JSON.stringify(duplicates));
  if (invalidDeps.length > 0) console.log('Invalid deps: ' + JSON.stringify(invalidDeps));
  if (orphanComments.length > 0) console.log('Orphan comments: ' + JSON.stringify(orphanComments));
  if (duplicates.length === 0 && invalidDeps.length === 0 && orphanComments.length === 0) {
    console.log('ALL CLEAR — no integrity issues found');
  }
  return { duplicates: duplicates, invalidDeps: invalidDeps, orphanComments: orphanComments };
}

function backfillTaskUids() {
  var sheet = getTasksSheet();
  var columns = CONFIG.TASK_COLUMNS;
  var uidCol = columns.indexOf('taskUid');
  if (uidCol === -1) throw new Error('taskUid column not found in TASK_COLUMNS');
  var data = sheet.getDataRange().getValues();
  var count = 0;
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    if (!data[i][uidCol] || !isValidTaskUid(String(data[i][uidCol]))) {
      data[i][uidCol] = generateTaskUid();
      count++;
    }
  }
  if (count > 0) {
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    SpreadsheetApp.flush();
  }
  console.log('backfillTaskUids: backfilled ' + count + ' of ' + (data.length - 1) + ' tasks');
  return { backfilledCount: count, totalRows: data.length - 1 };
}

function repairTaskIdentity() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getTasksSheet();
    var columns = CONFIG.TASK_COLUMNS;
    var idCol = columns.indexOf('id');
    var uidCol = columns.indexOf('taskUid');
    var depsCol = columns.indexOf('dependencies');
    var projectIdCol = columns.indexOf('projectId');
    var data = sheet.getDataRange().getValues();
    var idMap = {};
    for (var i = 1; i < data.length; i++) {
      var id = data[i][idCol];
      if (!id) continue;
      if (!idMap[id]) idMap[id] = [];
      idMap[id].push(i);
    }
    var remaps = [];
    for (var id in idMap) {
      if (idMap[id].length <= 1) continue;
      var rows = idMap[id];
      for (var j = 1; j < rows.length; j++) {
        var rowIdx = rows[j];
        var projectId = data[rowIdx][projectIdCol] || 'TASK';
        var newId = generateTaskIdUnderLock_(projectId);
        remaps.push({ oldId: id, newId: newId, rowIndex: rowIdx + 1 });
        data[rowIdx][idCol] = newId;
        if (data[rowIdx][depsCol]) {
          data[rowIdx][depsCol] = String(data[rowIdx][depsCol]).split(',').map(function(d) {
            return d.trim() === id ? newId : d.trim();
          }).join(',');
        }
      }
    }
    for (var i = 1; i < data.length; i++) {
      if (!data[i][idCol]) continue;
      if (data[i][depsCol]) {
        var changed = false;
        var deps = String(data[i][depsCol]).split(',').map(function(d) {
          var trimmed = d.trim();
          for (var r = 0; r < remaps.length; r++) {
            if (remaps[r].oldId === trimmed) { changed = true; return remaps[r].newId; }
          }
          return trimmed;
        }).join(',');
        if (changed) data[i][depsCol] = deps;
      }
    }
    var uidBackfillCount = 0;
    for (var i = 1; i < data.length; i++) {
      if (!data[i][idCol]) continue;
      if (!data[i][uidCol] || !isValidTaskUid(String(data[i][uidCol]))) {
        data[i][uidCol] = generateTaskUid();
        uidBackfillCount++;
      }
    }
    if (remaps.length > 0 || uidBackfillCount > 0) {
      sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
      SpreadsheetApp.flush();
    }
    remaps.forEach(function(remap) {
      cascadeTaskIdRemap_(remap.oldId, remap.newId);
    });
    var currentUser = getCurrentUserEmail();
    remaps.forEach(function(remap) {
      logActivity(currentUser, 'repair_identity', 'task', remap.newId, { oldId: remap.oldId, reason: 'duplicate_id_repair' });
    });
    console.log('repairTaskIdentity: repaired ' + remaps.length + ' duplicates, backfilled ' + uidBackfillCount + ' UIDs');
    return { duplicatesRepaired: remaps.length, uidsBackfilled: uidBackfillCount, remaps: remaps };
  } finally {
    lock.releaseLock();
  }
}

function cascadeTaskIdRemap_(oldId, newId) {
  try {
    var commentsSheet = getCommentsSheet();
    var commentCols = CONFIG.COMMENT_COLUMNS;
    var commentTaskCol = commentCols.indexOf('taskId');
    if (commentTaskCol !== -1) {
      cascadeColumnValue_(commentsSheet, commentTaskCol, oldId, newId);
    }
  } catch (e) { console.error('cascadeTaskIdRemap_ comments failed:', e); }
  try {
    var mentionsSheet = getMentionsSheet();
    var mentionCols = CONFIG.MENTION_COLUMNS;
    var mentionTaskCol = mentionCols.indexOf('taskId');
    if (mentionTaskCol !== -1) {
      cascadeColumnValue_(mentionsSheet, mentionTaskCol, oldId, newId);
    }
  } catch (e) { console.error('cascadeTaskIdRemap_ mentions failed:', e); }
  try {
    var notifSheet = getNotificationsSheet();
    var notifCols = CONFIG.NOTIFICATION_COLUMNS;
    var entityIdCol = notifCols.indexOf('entityId');
    var entityTypeCol = notifCols.indexOf('entityType');
    if (entityIdCol !== -1 && entityTypeCol !== -1) {
      var data = notifSheet.getDataRange().getValues();
      var changed = false;
      for (var i = 1; i < data.length; i++) {
        if (data[i][entityTypeCol] === 'task' && data[i][entityIdCol] === oldId) {
          data[i][entityIdCol] = newId;
          changed = true;
        }
      }
      if (changed) { notifSheet.getRange(1, 1, data.length, data[0].length).setValues(data); }
    }
  } catch (e) { console.error('cascadeTaskIdRemap_ notifications failed:', e); }
  try {
    var depsSheet = getSheet(CONFIG.SHEETS.TASK_DEPENDENCIES, CONFIG.TASK_DEPENDENCY_COLUMNS || ['id', 'successorId', 'predecessorId', 'type', 'lag', 'createdAt']);
    var depsData = depsSheet.getDataRange().getValues();
    var depsChanged = false;
    for (var i = 1; i < depsData.length; i++) {
      if (depsData[i][1] === oldId) { depsData[i][1] = newId; depsChanged = true; }
      if (depsData[i][2] === oldId) { depsData[i][2] = newId; depsChanged = true; }
    }
    if (depsChanged) { depsSheet.getRange(1, 1, depsData.length, depsData[0].length).setValues(depsData); }
  } catch (e) { console.error('cascadeTaskIdRemap_ dependencies failed:', e); }
  try {
    var actSheet = getSheet(CONFIG.SHEETS.ACTIVITY, CONFIG.ACTIVITY_COLUMNS);
    var actCols = CONFIG.ACTIVITY_COLUMNS;
    var actEntityIdCol = actCols.indexOf('entityId');
    var actEntityTypeCol = actCols.indexOf('entityType');
    if (actEntityIdCol !== -1 && actEntityTypeCol !== -1) {
      var actData = actSheet.getDataRange().getValues();
      var actChanged = false;
      for (var i = 1; i < actData.length; i++) {
        if (actData[i][actEntityTypeCol] === 'task' && actData[i][actEntityIdCol] === oldId) {
          actData[i][actEntityIdCol] = newId;
          actChanged = true;
        }
      }
      if (actChanged) { actSheet.getRange(1, 1, actData.length, actData[0].length).setValues(actData); }
    }
  } catch (e) { console.error('cascadeTaskIdRemap_ activity failed:', e); }
}

function cascadeColumnValue_(sheet, colIndex, oldValue, newValue) {
  var data = sheet.getDataRange().getValues();
  var changed = false;
  for (var i = 1; i < data.length; i++) {
    if (data[i][colIndex] === oldValue) {
      data[i][colIndex] = newValue;
      changed = true;
    }
  }
  if (changed) {
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  }
}

function checkTaskIntegrity() {
  var sheet = getTasksSheet();
  var columns = CONFIG.TASK_COLUMNS;
  var idCol = columns.indexOf('id');
  var uidCol = columns.indexOf('taskUid');
  var data = sheet.getDataRange().getValues();
  var ids = {};
  var duplicateCount = 0;
  var missingUidCount = 0;
  for (var i = 1; i < data.length; i++) {
    var id = data[i][idCol];
    if (!id) continue;
    if (ids[id]) duplicateCount++;
    else ids[id] = true;
    if (!data[i][uidCol]) missingUidCount++;
  }
  return { healthy: duplicateCount === 0 && missingUidCount === 0, duplicateCount: duplicateCount, missingUidCount: missingUidCount };
}

function resetTaskSystem() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getTasksSheet();
    var columns = CONFIG.TASK_COLUMNS;
    var idCol = columns.indexOf('id');
    var uidCol = columns.indexOf('taskUid');
    var projectIdCol = columns.indexOf('projectId');
    var data = sheet.getDataRange().getValues();
    var seqMap = {};
    var uidCount = 0;
    for (var i = 1; i < data.length; i++) {
      var id = String(data[i][idCol] || '');
      if (!id) continue;
      var match = id.match(/^(.+)-(\d+)$/);
      if (match) {
        var prefix = match[1];
        var seq = parseInt(match[2]);
        if (!seqMap[prefix] || seq > seqMap[prefix]) seqMap[prefix] = seq;
      }
      if (!data[i][uidCol] || !isValidTaskUid(String(data[i][uidCol]))) {
        data[i][uidCol] = generateTaskUid();
        uidCount++;
      }
    }
    if (uidCount > 0) {
      sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
      SpreadsheetApp.flush();
    }
    var props = PropertiesService.getScriptProperties();
    var seqResets = [];
    for (var prefix in seqMap) {
      var key = 'TASK_SEQ_' + prefix;
      var oldVal = props.getProperty(key) || '0';
      props.setProperty(key, String(seqMap[prefix]));
      seqResets.push({ projectId: prefix, oldSeq: parseInt(oldVal), newSeq: seqMap[prefix] });
    }
    try { invalidateTaskCache(null, 'update'); } catch (e) {}
    try { RowIndexCache.invalidateSheet('Tasks'); } catch (e) {}
    try {
      var scriptCache = CacheService.getScriptCache();
      scriptCache.removeAll(['ALL_TASKS_CACHE', 'BATCH_DATA_CACHE', 'BATCH_DATA_DIRTY']);
    } catch (e) {}
    console.log('resetTaskSystem: processed ' + (data.length - 1) + ' tasks, reset ' + seqResets.length + ' sequences, backfilled ' + uidCount + ' UIDs');
    return { tasksProcessed: data.length - 1, sequencesReset: seqResets, uidsBackfilled: uidCount, cacheCleared: true };
  } finally {
    lock.releaseLock();
  }
}

function resetTaskSequences() {
  var sheet = getTasksSheet();
  var columns = CONFIG.TASK_COLUMNS;
  var idCol = columns.indexOf('id');
  var data = sheet.getDataRange().getValues();
  var seqMap = {};
  for (var i = 1; i < data.length; i++) {
    var id = String(data[i][idCol] || '');
    if (!id) continue;
    var match = id.match(/^(.+)-(\d+)$/);
    if (match) {
      var prefix = match[1];
      var seq = parseInt(match[2]);
      if (!seqMap[prefix] || seq > seqMap[prefix]) seqMap[prefix] = seq;
    }
  }
  var props = PropertiesService.getScriptProperties();
  var resets = [];
  for (var prefix in seqMap) {
    var key = 'TASK_SEQ_' + prefix;
    var oldVal = props.getProperty(key) || '0';
    props.setProperty(key, String(seqMap[prefix]));
    resets.push({ projectId: prefix, oldSeq: parseInt(oldVal), newSeq: seqMap[prefix] });
  }
  console.log('resetTaskSequences: reset ' + resets.length + ' sequences');
  return { sequencesReset: resets };
}

function exportTaskSnapshot() {
  var sheet = getTasksSheet();
  var columns = CONFIG.TASK_COLUMNS;
  var data = sheet.getDataRange().getValues();
  var tasks = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    tasks.push(rowToObject(data[i], columns));
  }
  try {
    var ss = getColonySpreadsheet_();
    var backupName = 'Task_Backup_' + new Date().toISOString().replace(/[:.]/g, '-').substring(0, 19);
    var backupSheet = ss.insertSheet(backupName);
    backupSheet.getRange(1, 1, 1, columns.length).setValues([columns]).setFontWeight('bold');
    if (data.length > 1) {
      backupSheet.getRange(2, 1, data.length - 1, data[0].length).setValues(data.slice(1));
    }
    console.log('exportTaskSnapshot: created backup sheet "' + backupName + '" with ' + tasks.length + ' tasks');
    return { success: true, backupSheet: backupName, taskCount: tasks.length, tasks: tasks };
  } catch (e) {
    console.error('exportTaskSnapshot: failed to create backup sheet:', e);
    return { success: true, backupSheet: null, taskCount: tasks.length, tasks: tasks };
  }
}

function releaseTaskEditLock(taskId) {
  try {
    if (typeof LockManager === 'undefined') {
      return { success: true };
    }

    return LockManager.releaseEditLock(taskId);
  } catch (error) {
    console.error('releaseTaskEditLock failed:', error);
    return { success: false, error: error.message };
  }
}

function checkTaskEditLock(taskId) {
  try {
    if (typeof LockManager === 'undefined') {
      return { locked: false };
    }

    return LockManager.checkEditLock(taskId);
  } catch (error) {
    console.error('checkTaskEditLock failed:', error);
    return { locked: false };
  }
}
