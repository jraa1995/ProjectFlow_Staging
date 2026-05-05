function _ts_authorizeTaskRead_(task) {
  if (!task) return false;
  try {
    var role = (typeof getCurrentUserRole === 'function') ? getCurrentUserRole() : null;
    if (role === 'admin' || role === 'manager') return true;
    if (typeof getManagerContractFromRole === 'function' && getManagerContractFromRole(role)) {
      var allowed = getManagerAccessibleProjectIds(role);
      return !!(task.projectId && allowed[task.projectId]);
    }
    if (role === 'client') {
      var email = (typeof getCurrentUserEmailOptimized === 'function') ? getCurrentUserEmailOptimized() : null;
      if (!email) return false;
      var allowedC = getClientAccessibleProjectIds(email);
      if (task.projectId && allowedC[task.projectId]) return true;
      var emailLc = String(email).toLowerCase();
      if (task.assignee && String(task.assignee).toLowerCase() === emailLc) return true;
      if (task.reporter && String(task.reporter).toLowerCase() === emailLc) return true;
      return false;
    }
    return PermissionGuard.canReadTask(task);
  } catch (e) {
    console.error('_ts_authorizeTaskRead_ failed:', e);
    return false;
  }
}

function _ts_authorizeTaskWrite_(task, action) {
  if (!task) return false;
  try {
    var role = (typeof getCurrentUserRole === 'function') ? getCurrentUserRole() : null;
    if (role === 'admin') return true;
    if (typeof getManagerContractFromRole === 'function' && getManagerContractFromRole(role)) {
      var allowed = getManagerAccessibleProjectIds(role);
      return !!(task.projectId && allowed[task.projectId]);
    }
    if (role === 'client') {
      var email = (typeof getCurrentUserEmailOptimized === 'function') ? getCurrentUserEmailOptimized() : null;
      if (!email) return false;
      var allowedC = getClientAccessibleProjectIds(email);
      if (!task.projectId || !allowedC[task.projectId]) return false;
      var emailLc = String(email).toLowerCase();
      if (action === 'delete') {
        return !!(task.reporter && String(task.reporter).toLowerCase() === emailLc);
      }
      return !!(
        (task.assignee && String(task.assignee).toLowerCase() === emailLc) ||
        (task.reporter && String(task.reporter).toLowerCase() === emailLc)
      );
    }
    if (action === 'update') return PermissionGuard.canUpdateTask(task);
    if (action === 'delete') return PermissionGuard.canDeleteTask(task);
    return PermissionGuard.canReadTask(task);
  } catch (e) {
    console.error('_ts_authorizeTaskWrite_ failed:', e);
    return false;
  }
}

function _ts_authorizeProjectCreate_(projectId) {
  try {
    if (!projectId) return true;
    var pidUpper = String(projectId).toUpperCase();
    if (pidUpper === 'ADHOC' || pidUpper === 'TASK') return true;
    var role = (typeof getCurrentUserRole === 'function') ? getCurrentUserRole() : null;
    if (role === 'admin' || role === 'manager') return true;
    if (typeof getManagerContractFromRole === 'function' && getManagerContractFromRole(role)) {
      var allowed = getManagerAccessibleProjectIds(role);
      return !!allowed[projectId];
    }
    if (role === 'client') {
      var email = (typeof getCurrentUserEmailOptimized === 'function') ? getCurrentUserEmailOptimized() : null;
      if (!email) return false;
      var allowedC = getClientAccessibleProjectIds(email);
      return !!allowedC[projectId];
    }
    return true;
  } catch (e) {
    console.error('_ts_authorizeProjectCreate_ failed:', e);
    return false;
  }
}

function _ts_safeUpdates_(updates) {
  if (!updates || typeof updates !== 'object') return {};
  var safe = {};
  Object.keys(updates).slice(0, 30).forEach(function(k) {
    var v = updates[k];
    if (v === null || v === undefined) { safe[k] = v; return; }
    if (typeof v === 'string') { safe[k] = v.length > 200 ? v.slice(0, 200) + '...' : v; }
    else if (typeof v === 'number' || typeof v === 'boolean') { safe[k] = v; }
    else { safe[k] = '[object]'; }
  });
  return safe;
}

function loadMyBoard(projectId) {
  return getMyBoardOptimized(projectId || null);
}

function loadMasterBoard(projectId) {
  try {
    var userRole = getCurrentUserRole();
    if (userRole !== 'admin' && !isManagerRole(userRole)) {
      throw new Error('Permission denied: Master Board requires admin or manager access.');
    }

    var batchData = getBatchDataFast();
    var filteredTasks = projectId
      ? batchData.tasks.filter(function(task) { return task.projectId === projectId; })
      : batchData.tasks;
    if (getManagerContractFromRole(userRole)) {
      var userEmail = getCurrentUserEmailOptimized();
      filteredTasks = filterTasksByUserRole(filteredTasks, userEmail, userRole);
    }

    var board = buildBoardData(filteredTasks, projectId, { view: 'master' });

    var integrityResult = null;
    try { integrityResult = checkTaskIntegrity(); } catch (e) {}
    return {
      ...board,
      projects: batchData.projects,
      users: batchData.users,
      taskCount: filteredTasks.length,
      integrityResult: integrityResult
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

    var integrityResult = null;
    try { integrityResult = checkTaskIntegrity(); } catch (e) {}
    return {
      ...board,
      projects: batchData.projects,
      users: batchData.users,
      taskCount: filteredTasks.length,
      integrityResult: integrityResult
    };
  } catch (error) {
    console.error('getMyBoardOptimized failed:', error);
    throw error;
  }
}

function saveNewTask(taskData) {
  try {
    taskData = taskData || {};
    var title = typeof taskData.title === 'string' ? taskData.title.trim() : '';
    if (!title) throw new Error('Title is required');
    if (!taskData.assignee) throw new Error('Assignee is required');
    var projectId = taskData.projectId || '';
    var daId = taskData.dataAssetId || '';
    if (!projectId && !daId) throw new Error('Project is required');
    if (!_ts_authorizeProjectCreate_(projectId)) {
      throw new Error('Permission denied: cannot create task in project ' + projectId);
    }
    taskData.title = title;
    const result = createTask(taskData);
    patchTaskCache(result.id, result, 'create');
    invalidateTaskCache(result.id, 'create');
    try {
      if (typeof logAuditEvent === 'function') {
        logAuditEvent('task.create', 'task', result.id, result.title || '', { projectId: result.projectId, assignee: result.assignee });
      }
    } catch (e) {}
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
  var _existingTask = null;
  try { _existingTask = getTaskById(taskId); } catch (e) {}
  if (_existingTask && !_ts_authorizeTaskWrite_(_existingTask, 'update')) {
    throw new Error('Permission denied: cannot update this task');
  }
  const result = updateTask(taskId, updates);
  invalidateTaskCache(taskId, 'update');
  try {
    if (typeof logAuditEvent === 'function') {
      logAuditEvent('task.update', 'task', taskId, (result && result.title) || '', _ts_safeUpdates_(updates));
    }
  } catch (e) {}
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
  var _existingTask = null;
  try { _existingTask = getTaskById(taskId); } catch (e) {}
  if (_existingTask && !_ts_authorizeTaskWrite_(_existingTask, 'update')) {
    throw new Error('Permission denied: cannot update this task');
  }
  const status = denormalizeStatusId(newStatus);
  const result = moveTask(taskId, status, newPosition);
  invalidateTaskCache(taskId, 'update');
  try {
    if (typeof logAuditEvent === 'function') {
      logAuditEvent('task.move', 'task', taskId, (result && result.title) || '', { status: status, position: newPosition });
    }
  } catch (e) {}
  return result;
}

function removeTask(taskId) {
  if (isValidTaskUid(taskId)) {
    var resolved = getTaskByUid(taskId);
    if (resolved) taskId = resolved.id;
  }
  var _existingTask = null;
  try { _existingTask = getTaskById(taskId); } catch (e) {}
  if (_existingTask && !_ts_authorizeTaskWrite_(_existingTask, 'delete')) {
    throw new Error('Permission denied: cannot delete this task');
  }
  const result = deleteTask(taskId);
  invalidateTaskCache(taskId, 'delete');
  try {
    if (typeof logAuditEvent === 'function') {
      logAuditEvent('task.delete', 'task', taskId, (_existingTask && _existingTask.title) || '', { projectId: _existingTask && _existingTask.projectId });
    }
  } catch (e) {}
  return result;
}

function loadTask(taskId) {
  try {
    if (!taskId || typeof taskId !== 'string' || taskId.trim() === '') {
      return null;
    }
    var id = taskId.trim();
    var task;
    if (isValidTaskUid(id)) {
      task = getTaskByUid(id);
    } else {
      task = getTaskById(id);
    }
    if (!task) return null;
    if (!_ts_authorizeTaskRead_(task)) return null;
    return task;
  } catch (error) {
    console.error('loadTask error for taskId ' + taskId + ':', error);
    return null;
  }
}

function loadTaskDetail(taskId) {
  try {
    if (!taskId || typeof taskId !== 'string' || taskId.trim() === '') return null;
    var id = taskId.trim();
    var task;
    if (isValidTaskUid(id)) {
      task = getTaskByUid(id);
    } else {
      task = getTaskById(id);
    }
    if (!task) return null;
    if (!_ts_authorizeTaskRead_(task)) return null;
    var comments = [];
    try { comments = loadComments(task.id); } catch (e) { console.error('loadTaskDetail comments error:', e); }
    var dependencies = [];
    try { dependencies = getTaskDependenciesWithDetails(task.id); } catch (e) { console.error('loadTaskDetail dependencies error:', e); }
    var mentionUsers = [];
    try { mentionUsers = getAllUsersForMentions(false); } catch (e) { console.error('loadTaskDetail mentionUsers error:', e); }
    return {
      task: task,
      comments: comments,
      dependencies: dependencies,
      mentionUsers: mentionUsers
    };
  } catch (error) {
    console.error('loadTaskDetail failed for ' + taskId + ':', error);
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

  var _batchResults = null;
  var _batchNotifSpecs = [];
  var _batchDenied = [];
  var _batchError = null;
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
        var _opTask = null;
        try { _opTask = getTaskById(op.params.taskId); } catch (e) {}
        if (_opTask && !_ts_authorizeTaskWrite_(_opTask, 'update')) {
          _batchDenied.push({ taskId: op.params.taskId, error: 'Permission denied' });
          return;
        }
        updatesList.push({ taskId: op.params.taskId, changes: changes });
      }
    });
    _batchResults = batchUpdateTasks(updatesList);
    if (_batchResults && _batchResults.notificationSpecs) {
      _batchNotifSpecs = _batchResults.notificationSpecs;
    }
    invalidateTaskCache(null, 'update');
    try {
      if (typeof logAuditEvent === 'function') {
        logAuditEvent('task.batch_update', 'task', '', '', { count: _batchResults.length, denied: _batchDenied.length });
      }
    } catch (e) {}
  } catch (e) {
    console.error('executeBatchOperations failed:', e.message);
    _batchError = e.message;
  } finally {
    lock.releaseLock();
  }
  if (_batchError) {
    return { results: [], errors: [{ index: 0, success: false, error: _batchError }] };
  }
  if (typeof NotificationEngine !== 'undefined' && _batchNotifSpecs.length > 0) {
    _batchNotifSpecs.forEach(function(spec) {
      try { NotificationEngine.createNotification(spec); }
      catch (e) { console.error('Failed to dispatch batch notification:', e); }
    });
  }
  return {
    results: (_batchResults || []).map(function(r, i) { return { index: i, success: true, result: r }; }),
    errors: _batchDenied
  };
}

function bulkUpdateStatus(taskIds, newStatus) {
  try {
    if (!Array.isArray(taskIds) || taskIds.length === 0) {
      return { success: false, error: 'No tasks selected' };
    }
    var status = denormalizeStatusId(newStatus);
    var updated = [];
    var failed = [];
    taskIds.forEach(function(rawId) {
      try {
        var resolvedId = rawId;
        if (isValidTaskUid(rawId)) {
          var resolvedTask = getTaskByUid(rawId);
          if (resolvedTask) resolvedId = resolvedTask.id;
        }
        var existingTask = getTaskById(resolvedId);
        if (!existingTask) { failed.push({ taskId: resolvedId, error: 'Not found' }); return; }
        if (!_ts_authorizeTaskWrite_(existingTask, 'update')) {
          failed.push({ taskId: resolvedId, error: 'Permission denied' });
          return;
        }
        moveTask(resolvedId, status);
        invalidateTaskCache(resolvedId, 'update');
        updated.push({ taskId: resolvedId, status: status });
      } catch (e) {
        failed.push({ taskId: rawId, error: e.message || 'Unknown error' });
      }
    });
    try {
      if (typeof logAuditEvent === 'function') {
        logAuditEvent('task.bulk_status', 'task', '', '', { count: updated.length, denied: failed.length, status: status });
      }
    } catch (e) {}
    return { success: true, updated: updated, failed: failed };
  } catch (error) {
    console.error('bulkUpdateStatus failed:', error);
    return { success: false, error: error.message };
  }
}

function bulkDeleteTasks(taskIds) {
  try {
    if (!Array.isArray(taskIds) || taskIds.length === 0) {
      return { success: false, error: 'No tasks selected' };
    }
    var deleted = [];
    var failed = [];
    taskIds.forEach(function(rawId) {
      try {
        var resolvedId = rawId;
        if (isValidTaskUid(rawId)) {
          var resolvedTask = getTaskByUid(rawId);
          if (resolvedTask) resolvedId = resolvedTask.id;
        }
        var existingTask = getTaskById(resolvedId);
        if (!existingTask) { failed.push({ taskId: resolvedId, error: 'Not found' }); return; }
        if (!_ts_authorizeTaskWrite_(existingTask, 'delete')) {
          failed.push({ taskId: resolvedId, error: 'Permission denied' });
          return;
        }
        deleteTask(resolvedId);
        invalidateTaskCache(resolvedId, 'delete');
        deleted.push({ taskId: resolvedId });
      } catch (e) {
        failed.push({ taskId: rawId, error: e.message || 'Unknown error' });
      }
    });
    try {
      if (typeof logAuditEvent === 'function') {
        logAuditEvent('task.bulk_delete', 'task', '', '', { count: deleted.length, denied: failed.length });
      }
    } catch (e) {}
    return { success: true, deleted: deleted, failed: failed };
  } catch (error) {
    console.error('bulkDeleteTasks failed:', error);
    return { success: false, error: error.message };
  }
}

function updateTaskField(taskId, field, newValue) {
  try {
    if (!taskId || !field) return { success: false, error: 'taskId and field required' };
    var allowedFields = {
      'title': true, 'description': true, 'status': true, 'priority': true,
      'assignee': true, 'reporter': true, 'dueDate': true, 'startDate': true,
      'storyPoints': true, 'estimatedHrs': true, 'actualHrs': true,
      'sprint': true, 'type': true, 'labels': true, 'parentId': true,
      'projectId': true, 'milestoneDate': true
    };
    if (!allowedFields[field]) {
      return { success: false, error: 'Field not editable: ' + field };
    }
    var updates = {};
    updates[field] = newValue;
    var result = saveTaskUpdate(taskId, updates);
    return { success: true, task: result };
  } catch (error) {
    console.error('updateTaskField failed:', error);
    return { success: false, error: error.message };
  }
}

function updateTaskWithVersion(taskId, updates, expectedVersion) {
  try {
    var _existingTask = null;
    try { _existingTask = getTaskById(taskId); } catch (e) {}
    if (_existingTask && !_ts_authorizeTaskWrite_(_existingTask, 'update')) {
      throw new Error('Permission denied: cannot update this task');
    }

    if (typeof LockManager === 'undefined') {
      const task = updateTask(taskId, updates);
      try {
        if (typeof logAuditEvent === 'function') {
          logAuditEvent('task.update', 'task', taskId, (task && task.title) || '', _ts_safeUpdates_(updates));
        }
      } catch (e) {}
      return { success: true, task: task };
    }

    const task = LockManager.updateTaskWithLocking(taskId, updates, expectedVersion);
    invalidateTaskCache(taskId, 'update');
    try {
      if (typeof logAuditEvent === 'function') {
        logAuditEvent('task.update', 'task', taskId, (task && task.title) || '', _ts_safeUpdates_(updates));
      }
    } catch (e) {}
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
  return {
    summary: {
      tasksScanned: totalRows,
      uniqueIds: allIds.size,
      duplicateCount: duplicates.length,
      invalidDepCount: invalidDeps.length,
      orphanCommentCount: orphanComments.length,
      allClear: duplicates.length === 0 && invalidDeps.length === 0 && orphanComments.length === 0
    },
    duplicates: duplicates,
    invalidDeps: invalidDeps,
    orphanComments: orphanComments
  };
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
    try { clearAllCaches(); } catch (e) { console.error('repairTaskIdentity: cache clear failed:', e); }
    try { RowIndexCache.invalidateSheet('Tasks'); } catch (e) {}
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
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get('INTEGRITY_CHECK_RESULT');
    if (cached) return JSON.parse(cached);
  } catch (e) {}
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
  var result = { healthy: duplicateCount === 0 && missingUidCount === 0, duplicateCount: duplicateCount, missingUidCount: missingUidCount };
  try { CacheService.getScriptCache().put('INTEGRITY_CHECK_RESULT', JSON.stringify(result), 300); } catch (e) {}
  return result;
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
