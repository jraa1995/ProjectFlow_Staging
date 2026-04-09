function getSpreadsheet() {
  return getColonySpreadsheet_();
}

function getBatchDataFast() {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get('BATCH_DATA_CACHE');
    if (cached) {
      try { return JSON.parse(cached); } catch (e) {}
    }
    var jsonCached = readJsonCache('BATCH_DATA', 15);
    if (jsonCached) {
      try {
        var jsonStr = JSON.stringify(jsonCached);
        if (jsonStr.length < 100000) cache.put('BATCH_DATA_CACHE', jsonStr, 1800);
      } catch (e) {}
      return jsonCached;
    }
    return rebuildBatchCache() || buildBatchDataFallback();
  } catch (error) {
    console.error('getBatchDataFast failed:', error);
    return buildBatchDataFallback();
  }
}

function rebuildBatchCache() {
  const cache = CacheService.getScriptCache();
  const lockKey = 'BATCH_REBUILD_LOCK';
  if (cache.get(lockKey)) return null;
  cache.put(lockKey, '1', 10);
  try {
    const ss = getSpreadsheet();
    const tasksSheet = ss.getSheetByName(CONFIG.SHEETS.TASKS);
    const projectsSheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);
    const usersSheet = ss.getSheetByName(CONFIG.SHEETS.USERS);

    const tasksData = tasksSheet ? tasksSheet.getDataRange().getValues() : [];
    const projectsData = projectsSheet ? projectsSheet.getDataRange().getValues() : [];
    const usersData = usersSheet ? usersSheet.getDataRange().getValues() : [];

    const taskJsonIdx = CONFIG.TASK_COLUMNS.indexOf('jsonData');
    const projJsonIdx = CONFIG.PROJECT_COLUMNS.indexOf('jsonData');
    const userJsonIdx = CONFIG.USER_COLUMNS.indexOf('jsonData');

    const tasks = [];
    if (tasksData.length > 1) {
      for (let i = 1; i < tasksData.length; i++) {
        const row = tasksData[i];
        if (!row[0] && !row[2]) continue;
        var task;
        if (taskJsonIdx !== -1 && row[taskJsonIdx]) {
          try { task = JSON.parse(row[taskJsonIdx]); } catch (e) { task = rowToObject(row, CONFIG.TASK_COLUMNS); }
        } else {
          task = rowToObject(row, CONFIG.TASK_COLUMNS);
        }
        task.labels = typeof task.labels === 'string' && task.labels
          ? task.labels.split(',').map(l => l.trim()).filter(l => l)
          : [];
        task.storyPoints = parseInt(task.storyPoints) || 0;
        task.estimatedHrs = parseFloat(task.estimatedHrs) || 0;
        task.actualHrs = parseFloat(task.actualHrs) || 0;
        task.position = parseInt(task.position) || 0;
        tasks.push(task);
      }
    }

    const projects = [];
    if (projectsData.length > 1) {
      for (let i = 1; i < projectsData.length; i++) {
        const row = projectsData[i];
        if (!row[0]) continue;
        if (projJsonIdx !== -1 && row[projJsonIdx]) {
          try { projects.push(JSON.parse(row[projJsonIdx])); } catch (e) { projects.push(rowToObject(row, CONFIG.PROJECT_COLUMNS)); }
        } else {
          projects.push(rowToObject(row, CONFIG.PROJECT_COLUMNS));
        }
      }
    }

    const users = [];
    if (usersData.length > 1) {
      for (let i = 1; i < usersData.length; i++) {
        const row = usersData[i];
        if (!row[0]) continue;
        var user;
        if (userJsonIdx !== -1 && row[userJsonIdx]) {
          try { user = JSON.parse(row[userJsonIdx]); } catch (e) { user = rowToObject(row, CONFIG.USER_COLUMNS); }
        } else {
          user = rowToObject(row, CONFIG.USER_COLUMNS);
        }
        if (user.status === 'active' || !user.status) {
          users.push({
            email: user.email,
            name: user.name,
            role: user.role,
            avatar: user.avatar
          });
        }
      }
    }

    const result = {
      tasks,
      projects,
      users,
      config: {
        statuses: CONFIG.STATUSES,
        priorities: CONFIG.PRIORITIES,
        types: CONFIG.TYPES,
        colors: CONFIG.COLORS
      }
    };

    var jsonStr = JSON.stringify(result);
    if (jsonStr.length < 100000) {
      cache.put('BATCH_DATA_CACHE', jsonStr, 1800);
    }
    if (jsonStr.length < 45000) {
      writeJsonCache('BATCH_DATA', result);
    }
    return result;
  } finally {
    cache.remove(lockKey);
  }
}

function buildBatchDataFallback() {
  return {
    tasks: getAllTasks(),
    projects: getAllProjects(),
    users: getActiveUsers(),
    config: {
      statuses: CONFIG.STATUSES,
      priorities: CONFIG.PRIORITIES,
      types: CONFIG.TYPES,
      colors: CONFIG.COLORS
    }
  };
}

function getAllTasksOptimized() {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'ALL_TASKS_CACHE';
    const cached = cache.get(cacheKey);

    if (cached) {
      try {
        return JSON.parse(cached);
      } catch (e) {
      }
    }

    const tasks = getAllTasks();
    const jsonStr = JSON.stringify(tasks);
    if (jsonStr.length < 100000) {
      cache.put(cacheKey, jsonStr, 900);
    }
    return tasks;
  } catch (error) {
    console.error('getAllTasksOptimized failed:', error);
    return getAllTasks();
  }
}

function getAllProjectsOptimized() {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'ALL_PROJECTS_CACHE';
    const cached = cache.get(cacheKey);

    if (cached) {
      try {
        return JSON.parse(cached);
      } catch (e) {
      }
    }

    const projects = getAllProjects();
    cache.put(cacheKey, JSON.stringify(projects), 1800);
    return projects;
  } catch (error) {
    console.error('getAllProjectsOptimized failed:', error);
    return getAllProjects();
  }
}

function getActiveUsersOptimized(forceRefresh) {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'ACTIVE_USERS_CACHE';

    if (!forceRefresh) {
      const cached = cache.get(cacheKey);
      if (cached) {
        try {
          return JSON.parse(cached);
        } catch (e) {
        }
      }
    }

    const users = getActiveUsers();
    cache.put(cacheKey, JSON.stringify(users), 1800);
    return users;
  } catch (error) {
    console.error('getActiveUsersOptimized failed:', error);
    return getActiveUsers();
  }
}

function loadUsersFresh() {
  return getActiveUsersOptimized(true);
}

function clearAllCaches() {
  try {
    const cache = CacheService.getScriptCache();
    cache.removeAll(['ALL_TASKS_CACHE', 'ALL_PROJECTS_CACHE', 'ACTIVE_USERS_CACHE', 'ALL_USERS_MENTIONS_CACHE', 'BATCH_DATA_CACHE']);
    RequestCache.clear();
  } catch (error) {
    console.error('clearAllCaches failed:', error);
  }
}

function clearUserCache() {
  try {
    const cache = CacheService.getScriptCache();
    cache.removeAll(['ACTIVE_USERS_CACHE', 'ALL_USERS_MENTIONS_CACHE']);
    return { success: true };
  } catch (error) {
    console.error('clearUserCache failed:', error);
    return { success: false, error: error.message };
  }
}

function patchTaskCache(taskId, updatedTask, changeType) {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'ALL_TASKS_CACHE';
    const cached = cache.get(cacheKey);

    if (cached) {
      const tasks = JSON.parse(cached);
      const index = tasks.findIndex(t => t.id === taskId);

      if (changeType === 'delete') {
        if (index >= 0) tasks.splice(index, 1);
      } else if (changeType === 'create' && updatedTask) {
        tasks.push(updatedTask);
      } else if (updatedTask && index >= 0) {
        tasks[index] = updatedTask;
      }

      const jsonStr = JSON.stringify(tasks);
      if (jsonStr.length < 100000) {
        cache.put(cacheKey, jsonStr, 300);
      } else {
        cache.remove(cacheKey);
      }
    }

    cache.remove(`TASK_${taskId}`);
    cache.remove(`row_Tasks_${taskId}`);
    logChange('task', taskId, changeType || 'update');

    const batchCached = cache.get('BATCH_DATA_CACHE');
    if (batchCached) {
      try {
        const batchData = JSON.parse(batchCached);
        const batchIndex = batchData.tasks.findIndex(t => t.id === taskId);
        if (changeType === 'delete') {
          if (batchIndex >= 0) batchData.tasks.splice(batchIndex, 1);
        } else if (changeType === 'create' && updatedTask) {
          batchData.tasks.push(updatedTask);
        } else if (updatedTask && batchIndex >= 0) {
          batchData.tasks[batchIndex] = updatedTask;
        }
        const batchJson = JSON.stringify(batchData);
        if (batchJson.length < 100000) {
          cache.put('BATCH_DATA_CACHE', batchJson, 300);
        } else {
          cache.remove('BATCH_DATA_CACHE');
        }
      } catch (e) {
        try { rebuildBatchCache(); } catch (e2) { cache.remove('BATCH_DATA_CACHE'); }
      }
    } else {
      try { rebuildBatchCache(); } catch (e) {}
    }

    if (RequestCache && RequestCache._tasks) {
      if (changeType === 'delete') {
        RequestCache._tasks = RequestCache._tasks.filter(t => t.id !== taskId);
      } else if (changeType === 'create' && updatedTask) {
        RequestCache._tasks.push(updatedTask);
      } else if (updatedTask) {
        const idx = RequestCache._tasks.findIndex(t => t.id === taskId);
        if (idx >= 0) RequestCache._tasks[idx] = updatedTask;
      }
    }
  } catch (error) {
    console.error('patchTaskCache failed:', error);
    invalidateTaskCache(taskId, changeType);
  }
}

function invalidateTaskCache(taskId, changeType) {
  try {
    const cache = CacheService.getScriptCache();
    cache.remove('ALL_TASKS_CACHE');

    if (taskId) {
      cache.remove(`TASK_${taskId}`);
      cache.remove(`row_Tasks_${taskId}`);
      logChange('task', taskId, changeType || 'update');
    } else {
      incrementGlobalVersion();
    }

    if (RequestCache && RequestCache._tasks) {
      RequestCache._tasks = null;
    }

    try {
      rebuildBatchCache();
    } catch (e) {
      cache.remove('BATCH_DATA_CACHE');
    }
  } catch (error) {
    console.error('invalidateTaskCache failed:', error);
  }
}

function invalidateProjectCache(projectId) {
  try {
    const cache = CacheService.getScriptCache();
    cache.remove('ALL_PROJECTS_CACHE');

    if (projectId) {
      cache.remove(`PROJECT_${projectId}`);
    }

    if (RequestCache && RequestCache._projects !== undefined) {
      RequestCache._projects = null;
    }
  } catch (error) {
    console.error('invalidateProjectCache failed:', error);
  }
}

function invalidateUserCache() {
  try {
    const cache = CacheService.getScriptCache();
    cache.remove('ACTIVE_USERS_CACHE');

    if (RequestCache && RequestCache._users !== undefined) {
      RequestCache._users = null;
    }
  } catch (error) {
    console.error('invalidateUserCache failed:', error);
  }
}

function invalidateDependencyCache() {
  try {
    if (RequestCache) {
      RequestCache._dependencies = null;
    }
  } catch (error) {
    console.error('invalidateDependencyCache failed:', error);
  }
}

function getGlobalVersion() {
  try {
    const cache = CacheService.getScriptCache();
    let version = cache.get('GLOBAL_DATA_VERSION');

    if (!version) {
      version = PropertiesService.getScriptProperties().getProperty('GLOBAL_DATA_VERSION') || '0';
      cache.put('GLOBAL_DATA_VERSION', version, 600);
    }

    return parseInt(version, 10);
  } catch (error) {
    console.error('getGlobalVersion failed:', error);
    return 0;
  }
}

function incrementGlobalVersion() {
  try {
    const props = PropertiesService.getScriptProperties();
    const current = parseInt(props.getProperty('GLOBAL_DATA_VERSION') || '0', 10);
    const next = current + 1;
    props.setProperty('GLOBAL_DATA_VERSION', next.toString());
    CacheService.getScriptCache().put('GLOBAL_DATA_VERSION', next.toString(), 600);
    return next;
  } catch (error) {
    console.error('incrementGlobalVersion failed:', error);
    return 0;
  }
}

function logChange(entityType, entityId, changeType) {
  try {
    const version = incrementGlobalVersion();
    const cache = CacheService.getScriptCache();
    const recentChangesKey = 'RECENT_CHANGES';
    let recentChanges = [];

    const cached = cache.get(recentChangesKey);
    if (cached) {
      try {
        recentChanges = JSON.parse(cached);
      } catch (e) {
        recentChanges = [];
      }
    }

    recentChanges.push({
      entityType,
      entityId,
      changeType,
      version,
      timestamp: new Date().toISOString()
    });

    if (recentChanges.length > 100) {
      recentChanges = recentChanges.slice(-100);
    }

    cache.put(recentChangesKey, JSON.stringify(recentChanges), 600);
    return version;
  } catch (error) {
    console.error('logChange failed:', error);
    return 0;
  }
}

function getChangesSince(sinceVersion) {
  try {
    const currentVersion = getGlobalVersion();

    if (sinceVersion >= currentVersion) {
      return { version: currentVersion, changedIds: [], changes: [] };
    }

    const cache = CacheService.getScriptCache();
    const cached = cache.get('RECENT_CHANGES');
    let recentChanges = [];

    if (cached) {
      try {
        recentChanges = JSON.parse(cached);
      } catch (e) {
        recentChanges = [];
      }
    }

    const relevantChanges = recentChanges.filter(c => c.version > sinceVersion);
    const changedTaskIds = [...new Set(
      relevantChanges
        .filter(c => c.entityType === 'task')
        .map(c => c.entityId)
    )];

    return {
      version: currentVersion,
      changedIds: changedTaskIds,
      changes: relevantChanges
    };
  } catch (error) {
    console.error('getChangesSince failed:', error);
    return { version: 0, changedIds: [], changes: [] };
  }
}

function getTasksByIds(taskIds) {
  try {
    if (!taskIds || taskIds.length === 0) {
      return [];
    }

    const allTasks = getAllTasksOptimized();
    const taskIdSet = new Set(taskIds);
    return allTasks.filter(task => taskIdSet.has(task.id));
  } catch (error) {
    console.error('getTasksByIds failed:', error);
    return [];
  }
}
