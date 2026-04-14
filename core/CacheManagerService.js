function getSpreadsheet() {
  return getColonySpreadsheet_();
}

var VersionedCache = {
  _cache: null,
  _versionForRequest: null,

  _getCache: function() {
    if (!this._cache) this._cache = CacheService.getScriptCache();
    return this._cache;
  },

  _currentVersion: function() {
    if (this._versionForRequest === null) {
      this._versionForRequest = getGlobalVersion();
    }
    return this._versionForRequest;
  },

  resetRequestVersion: function() {
    this._versionForRequest = null;
  },

  put: function(key, data, ttl) {
    try {
      var version = this._currentVersion();
      var envelope = JSON.stringify({ v: version, d: data });
      if (envelope.length > 95000) return false;
      this._getCache().put(key, envelope, ttl || 900);
      return true;
    } catch (e) {
      console.error('VersionedCache.put failed:', e);
      return false;
    }
  },

  get: function(key) {
    try {
      var raw = this._getCache().get(key);
      if (!raw) return null;
      var envelope = JSON.parse(raw);
      if (envelope.v !== undefined && envelope.d !== undefined) {
        if (envelope.v < this._currentVersion()) return null;
        return envelope.d;
      }
      return envelope;
    } catch (e) {
      return null;
    }
  },

  remove: function(key) {
    this._getCache().remove(key);
  }
};

function invalidateCache(entityType, entityId, changeType) {
  try {
    logChange(entityType || 'unknown', entityId || '*', changeType || 'update');
    VersionedCache.resetRequestVersion();
    if (typeof RequestCache !== 'undefined') {
      RequestCache.clear();
    }
    CacheService.getScriptCache().put('BATCH_DATA_DIRTY', '1', 900);
  } catch (error) {
    console.error('invalidateCache failed:', error);
  }
}

function getBatchDataFast() {
  try {
    var cached = VersionedCache.get('BATCH_DATA_CACHE');
    if (cached) return cached;
    if (!CacheService.getScriptCache().get('BATCH_DATA_DIRTY')) {
      var jsonCached = readJsonCache('BATCH_DATA', 15);
      if (jsonCached) {
        VersionedCache.put('BATCH_DATA_CACHE', jsonCached, 1800);
        return jsonCached;
      }
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
    const isDeletedIdx = CONFIG.TASK_COLUMNS.indexOf('isDeleted');

    const tasks = [];
    if (tasksData.length > 1) {
      for (let i = 1; i < tasksData.length; i++) {
        const row = tasksData[i];
        if (!row[0] && !row[2]) continue;
        if (isDeletedIdx !== -1 && (row[isDeletedIdx] === true || row[isDeletedIdx] === 'true' || row[isDeletedIdx] === 'TRUE')) continue;
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
        task.isMilestone = task.isMilestone === true || task.isMilestone === 'true' || task.isMilestone === 'TRUE' || task.isMilestone === 1;
        if (task.taskUid && task.id) UidIndexCache.set(task.taskUid, task.id);
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

    VersionedCache.put('BATCH_DATA_CACHE', result, 1800);
    var jsonStr = JSON.stringify(result);
    if (jsonStr.length < 45000) {
      writeJsonCache('BATCH_DATA', result);
    }
    cache.remove('BATCH_DATA_DIRTY');
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
    var cached = VersionedCache.get('ALL_TASKS_CACHE');
    if (cached) return cached;
    const tasks = getAllTasks();
    VersionedCache.put('ALL_TASKS_CACHE', tasks, 900);
    return tasks;
  } catch (error) {
    console.error('getAllTasksOptimized failed:', error);
    return getAllTasks();
  }
}

function getAllProjectsOptimized() {
  try {
    var cached = VersionedCache.get('ALL_PROJECTS_CACHE');
    if (cached) return cached;
    const projects = getAllProjects();
    VersionedCache.put('ALL_PROJECTS_CACHE', projects, 1800);
    return projects;
  } catch (error) {
    console.error('getAllProjectsOptimized failed:', error);
    return getAllProjects();
  }
}

function getActiveUsersOptimized(forceRefresh) {
  try {
    if (!forceRefresh) {
      var cached = VersionedCache.get('ACTIVE_USERS_CACHE');
      if (cached) return cached;
    }
    const users = getActiveUsers();
    VersionedCache.put('ACTIVE_USERS_CACHE', users, 1800);
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
    incrementGlobalVersion();
    VersionedCache.resetRequestVersion();
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
    var tasks = VersionedCache.get('ALL_TASKS_CACHE');
    if (tasks) {
      var index = tasks.findIndex(function(t) { return t.id === taskId || t.taskUid === taskId; });
      if (changeType === 'delete') {
        if (index >= 0) tasks.splice(index, 1);
      } else if (changeType === 'create' && updatedTask) {
        var dupeIdx = tasks.findIndex(function(t) { return t.id === updatedTask.id || (updatedTask.taskUid && t.taskUid === updatedTask.taskUid); });
        if (dupeIdx >= 0) { tasks[dupeIdx] = updatedTask; } else { tasks.push(updatedTask); }
      } else if (updatedTask && index >= 0) {
        tasks[index] = updatedTask;
      }
      VersionedCache.put('ALL_TASKS_CACHE', tasks, 300);
    }

    var batchData = VersionedCache.get('BATCH_DATA_CACHE');
    if (batchData && batchData.tasks) {
      var batchIndex = batchData.tasks.findIndex(function(t) { return t.id === taskId || t.taskUid === taskId; });
      if (changeType === 'delete') {
        if (batchIndex >= 0) batchData.tasks.splice(batchIndex, 1);
      } else if (changeType === 'create' && updatedTask) {
        var batchDupeIdx = batchData.tasks.findIndex(function(t) { return t.id === updatedTask.id || (updatedTask.taskUid && t.taskUid === updatedTask.taskUid); });
        if (batchDupeIdx >= 0) { batchData.tasks[batchDupeIdx] = updatedTask; } else { batchData.tasks.push(updatedTask); }
      } else if (updatedTask && batchIndex >= 0) {
        batchData.tasks[batchIndex] = updatedTask;
      }
      VersionedCache.put('BATCH_DATA_CACHE', batchData, 300);
    }

    if (RequestCache && RequestCache._tasks) {
      if (changeType === 'delete') {
        RequestCache._tasks = RequestCache._tasks.filter(function(t) { return t.id !== taskId && t.taskUid !== taskId; });
      } else if (changeType === 'create' && updatedTask) {
        var reqDupeIdx = RequestCache._tasks.findIndex(function(t) { return t.id === updatedTask.id || (updatedTask.taskUid && t.taskUid === updatedTask.taskUid); });
        if (reqDupeIdx >= 0) { RequestCache._tasks[reqDupeIdx] = updatedTask; } else { RequestCache._tasks.push(updatedTask); }
      } else if (updatedTask) {
        var idx = RequestCache._tasks.findIndex(function(t) { return t.id === taskId || t.taskUid === taskId; });
        if (idx >= 0) RequestCache._tasks[idx] = updatedTask;
      }
    }
  } catch (error) {
    console.error('patchTaskCache failed:', error);
    invalidateTaskCache(taskId, changeType);
  }
}

function invalidateTaskCache(taskId, changeType) {
  invalidateCache('task', taskId, changeType);
}

function invalidateProjectCache(projectId) {
  invalidateCache('project', projectId, 'update');
}

function invalidateUserCache() {
  invalidateCache('user', null, 'update');
}

function invalidateDependencyCache() {
  invalidateCache('dependency', null, 'update');
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

function getSyncUpdate(sinceVersion) {
  try {
    var changeResult = getChangesSince(sinceVersion);
    if (changeResult.changedIds.length === 0) {
      return { version: changeResult.version, changedIds: [], changedTasks: [], changes: changeResult.changes };
    }
    var tasks = getTasksByIds(changeResult.changedIds);
    return {
      version: changeResult.version,
      changedIds: changeResult.changedIds,
      changedTasks: tasks,
      changes: changeResult.changes
    };
  } catch (error) {
    console.error('getSyncUpdate failed:', error);
    return { version: 0, changedIds: [], changedTasks: [], changes: [] };
  }
}
