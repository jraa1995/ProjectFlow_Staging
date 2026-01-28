const LockManager = {
  EDIT_LOCK_TTL_SECONDS: 300,
  SCRIPT_LOCK_TIMEOUT_MS: 30000,

  updateTaskWithLocking(taskId, updates, expectedVersion) {
    const sheet = getTasksSheet();
    const data = sheet.getDataRange().getValues();
    const columns = CONFIG.TASK_COLUMNS;
    const idIndex = 0;
    const versionIndex = columns.indexOf('version');
    const lastModifiedByIndex = columns.indexOf('lastModifiedBy');

    for (let i = 1; i < data.length; i++) {
      if (data[i][idIndex] === taskId) {
        const task = rowToObject(data[i], columns);
        const currentVersion = parseInt(task.version) || 0;

        if (expectedVersion !== undefined && expectedVersion !== null) {
          if (currentVersion !== expectedVersion) {
            throw new Error(JSON.stringify({
              code: 'VERSION_CONFLICT',
              message: 'This task was modified by another user. Please refresh and try again.',
              currentVersion: currentVersion,
              expectedVersion: expectedVersion,
              lastModifiedBy: task.lastModifiedBy
            }));
          }
        }

        Object.assign(task, updates);
        task.version = currentVersion + 1;
        task.lastModifiedBy = getCurrentUserEmail();
        task.updatedAt = now();

        if (updates.status === 'Done' && !task.completedAt) {
          task.completedAt = now();
        }

        const newRow = objectToRow(task, columns);
        sheet.getRange(i + 1, 1, 1, columns.length).setValues([newRow]);
        this.invalidateTaskCaches(taskId);
        return task;
      }
    }

    throw new Error('Task not found: ' + taskId);
  },

  createTaskWithVersion(taskData) {
    taskData.version = 1;
    taskData.lastModifiedBy = getCurrentUserEmail();
    return createTask(taskData);
  },

  acquireEditLock(taskId, userId) {
    try {
      userId = userId || getCurrentUserEmail();
      const cache = CacheService.getScriptCache();
      const lockKey = 'EDIT_LOCK_' + taskId;
      const existingLock = cache.get(lockKey);

      if (existingLock) {
        try {
          const lock = JSON.parse(existingLock);
          if (lock.userId === userId) {
            lock.expiresAt = new Date(Date.now() + this.EDIT_LOCK_TTL_SECONDS * 1000).toISOString();
            cache.put(lockKey, JSON.stringify(lock), this.EDIT_LOCK_TTL_SECONDS);
            return {
              success: true,
              lock: lock,
              extended: true
            };
          }
          if (new Date(lock.expiresAt) > new Date()) {
            return {
              success: false,
              error: 'Task is being edited by ' + (lock.userName || lock.userId),
              lock: lock
            };
          }
        } catch (e) {
        }
      }

      const user = getUserByEmail(userId);
      const lock = {
        taskId: taskId,
        userId: userId,
        userName: user?.name || userId,
        acquiredAt: new Date().toISOString(),
        expiresAt: new Date(Date.now() + this.EDIT_LOCK_TTL_SECONDS * 1000).toISOString()
      };

      cache.put(lockKey, JSON.stringify(lock), this.EDIT_LOCK_TTL_SECONDS);

      return {
        success: true,
        lock: lock
      };
    } catch (error) {
      console.error('acquireEditLock failed:', error);
      return {
        success: false,
        error: error.message
      };
    }
  },

  releaseEditLock(taskId, userId) {
    try {
      userId = userId || getCurrentUserEmail();
      const cache = CacheService.getScriptCache();
      const lockKey = 'EDIT_LOCK_' + taskId;
      const existingLock = cache.get(lockKey);

      if (existingLock) {
        try {
          const lock = JSON.parse(existingLock);
          if (lock.userId !== userId) {
            return {
              success: false,
              error: 'You do not own this lock'
            };
          }
        } catch (e) {
        }
      }

      cache.remove(lockKey);
      return { success: true };
    } catch (error) {
      console.error('releaseEditLock failed:', error);
      return {
        success: false,
        error: error.message
      };
    }
  },

  checkEditLock(taskId) {
    try {
      const cache = CacheService.getScriptCache();
      const lockKey = 'EDIT_LOCK_' + taskId;
      const existingLock = cache.get(lockKey);

      if (!existingLock) {
        return { locked: false };
      }

      try {
        const lock = JSON.parse(existingLock);
        if (new Date(lock.expiresAt) > new Date()) {
          return {
            locked: true,
            lock: lock,
            isOwnLock: lock.userId === getCurrentUserEmail()
          };
        }
      } catch (e) {
      }

      return { locked: false };
    } catch (error) {
      console.error('checkEditLock failed:', error);
      return { locked: false };
    }
  },

  acquireScriptLock() {
    const lock = LockService.getScriptLock();
    const acquired = lock.tryLock(this.SCRIPT_LOCK_TIMEOUT_MS);
    if (!acquired) {
      throw new Error('Could not acquire script lock. Another operation is in progress.');
    }
    return lock;
  },

  withScriptLock(fn) {
    const lock = this.acquireScriptLock();
    try {
      return fn();
    } finally {
      lock.releaseLock();
    }
  },

  invalidateTaskCaches(taskId) {
    try {
      const scriptCache = CacheService.getScriptCache();
      scriptCache.remove('ALL_TASKS_CACHE');
      if (typeof RequestCache !== 'undefined') {
        RequestCache.clear();
      }
    } catch (error) {
      console.error('Cache invalidation failed:', error);
    }
  },

  invalidateAllCaches() {
    try {
      const scriptCache = CacheService.getScriptCache();
      scriptCache.removeAll([
        'ALL_TASKS_CACHE',
        'ALL_PROJECTS_CACHE',
        'ACTIVE_USERS_CACHE'
      ]);
      if (typeof RequestCache !== 'undefined') {
        RequestCache.clear();
      }
    } catch (error) {
      console.error('Full cache invalidation failed:', error);
    }
  },

  getCached(key, fetchFn, options = {}) {
    const scriptTTL = options.scriptTTL || 300;
    const userTTL = options.userTTL || 1800;

    if (typeof RequestCache !== 'undefined') {
      const requestCached = RequestCache[key];
      if (requestCached !== undefined) {
        return requestCached;
      }
    }

    const scriptCache = CacheService.getScriptCache();
    const scriptCached = scriptCache.get(key);

    if (scriptCached) {
      try {
        const data = JSON.parse(scriptCached);
        if (typeof RequestCache !== 'undefined') {
          RequestCache[key] = data;
        }
        return data;
      } catch (e) {
        console.warn('Script cache parse error for', key);
      }
    }

    const data = fetchFn();
    const jsonData = JSON.stringify(data);

    if (jsonData.length < 100000) {
      scriptCache.put(key, jsonData, scriptTTL);
    }

    if (typeof RequestCache !== 'undefined') {
      RequestCache[key] = data;
    }

    return data;
  }
};

if (typeof module !== 'undefined' && module.exports) {
  module.exports = LockManager;
}
