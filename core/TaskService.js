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
  const results = [];
  const errors = [];

  if (!Array.isArray(operations) || operations.length === 0) {
    return { results: [], errors: [] };
  }

  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    operations.forEach((op, index) => {
      try {
        let result;
        switch (op.type) {
          case 'moveTask':
            result = moveTask(op.params.taskId, op.params.status, op.params.position || 0);
            break;
          case 'updateTask':
            result = updateTask(op.params.taskId, op.params.changes);
            break;
          case 'updatePriority':
            result = updateTask(op.params.taskId, { priority: op.params.priority });
            break;
          case 'updateAssignee':
            result = updateTask(op.params.taskId, { assignee: op.params.assignee });
            break;
          default:
            throw new Error('Unknown operation type: ' + op.type);
        }
        results.push({ index, success: true, result });
      } catch (e) {
        console.error('Batch operation', index, 'failed:', e.message);
        errors.push({ index, success: false, error: e.message });
      }
    });

    if (results.length > 0) {
      invalidateTaskCache(null, 'update');
    }
  } finally {
    lock.releaseLock();
  }

  return { results, errors };
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
