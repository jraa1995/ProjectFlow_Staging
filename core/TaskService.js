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
  const status = denormalizeStatusId(newStatus);
  const result = moveTask(taskId, status, newPosition);
  invalidateTaskCache(taskId, 'update');
  return result;
}

function removeTask(taskId) {
  const result = deleteTask(taskId);
  invalidateTaskCache(taskId, 'delete');
  return result;
}

function loadTask(taskId) {
  try {
    if (!taskId || typeof taskId !== 'string' || taskId.trim() === '') {
      return null;
    }

    const task = getTaskById(taskId.trim());
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
