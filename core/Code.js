function doGet(e) {
  try {
    initializeSystem();
    const template = HtmlService.createTemplateFromFile('ui/Index');
    return template.evaluate()
      .setTitle('ProjectFlow')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  } catch (error) {
    console.error('doGet error:', error);
    return HtmlService.createHtmlOutput(getErrorPage(error.message));
  }
}

function include(filename) {
  const cleanFilename = filename.includes('/') ? filename.split('/').pop() : filename;
  try {
    return HtmlService.createHtmlOutputFromFile(cleanFilename).getContent();
  } catch (e) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  }
}

function getErrorPage(message) {
  return `
  <!DOCTYPE html>
  <html>
  <head>
  <title>ProjectFlow - Setup Required</title>
  <style>
  body { font-family: -apple-system, sans-serif; padding: 40px; background: #f8fafc; }
  .container { max-width: 500px; margin: 0 auto; background: white; padding: 40px;
    border-radius: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); text-align: center; }
  h1 { color: #f59e0b; }
  .error { color: #dc2626; font-size: 14px; background: #fef2f2; padding: 10px; border-radius: 4px; margin: 20px 0; }
  button { padding: 12px 24px; background: #f59e0b; color: white; border: none;
    border-radius: 8px; cursor: pointer; font-size: 16px; }
  button:hover { background: #d97706; }
  code { background: #f1f5f9; padding: 2px 8px; border-radius: 4px; }
  </style>
  </head>
  <body>
  <div class="container">
  <h1>⚙️ Setup Required</h1>
  <p>Please run <code>quickSetup()</code> in the Apps Script editor.</p>
  <div class="error">${message}</div>
  <button onclick="location.reload()">Retry</button>
  </div>
  </body>
  </html>
  `;
}

function initializeSystem() {
  try {
    if (typeof CONFIG === 'undefined') {
      throw new Error('CONFIG object is undefined. Check Config.gs for syntax errors.');
    }

    const sheetCreators = [
      { fn: getTasksSheet, name: 'Tasks' },
      { fn: getUsersSheet, name: 'Users' },
      { fn: getProjectsSheet, name: 'Projects' },
      { fn: getCommentsSheet, name: 'Comments' },
      { fn: getActivitySheet, name: 'Activity' },
      { fn: getMentionsSheet, name: 'Mentions' },
      { fn: getNotificationsSheet, name: 'Notifications' },
      { fn: getAnalyticsCacheSheet, name: 'Analytics Cache' },
      { fn: getTaskDependenciesSheet, name: 'Task Dependencies' },
      { fn: getFunnelStagingSheet, name: 'Funnel Staging' }
    ];

    sheetCreators.forEach(({ fn, name }) => {
      try {
        fn();
      } catch (e) {
        console.error(`Failed to create ${name} sheet:`, e.message);
        throw e;
      }
    });

    getCurrentUser();
  } catch (error) {
    console.error('initializeSystem failed:', error);
    throw error;
  }
}

const RequestCache = {
  _tasks: null,
  _projects: null,
  _users: null,
  _activity: null,

  getTasks() {
    if (this._tasks === null) {
      this._tasks = getAllTasks();
    }
    return this._tasks;
  },

  getProjects() {
    if (this._projects === null) {
      this._projects = getAllProjects();
    }
    return this._projects;
  },

  getUsers() {
    if (this._users === null) {
      this._users = getActiveUsers();
    }
    return this._users;
  },

  getActivity(limit) {
    if (this._activity === null) {
      this._activity = getRecentActivity(limit || 20);
    }
    return this._activity;
  },

  clear() {
    this._tasks = null;
    this._projects = null;
    this._users = null;
    this._activity = null;
  }
};

function getBatchDataFast() {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'BATCH_DATA_CACHE';
    const cached = cache.get(cacheKey);

    if (cached) {
      try {
        return JSON.parse(cached);
      } catch (e) {
        console.warn('Batch cache parse error');
      }
    }

    const ss = getSpreadsheet();
    const tasksSheet = ss.getSheetByName(CONFIG.SHEETS.TASKS);
    const projectsSheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);
    const usersSheet = ss.getSheetByName(CONFIG.SHEETS.USERS);

    const tasksData = tasksSheet ? tasksSheet.getDataRange().getValues() : [];
    const projectsData = projectsSheet ? projectsSheet.getDataRange().getValues() : [];
    const usersData = usersSheet ? usersSheet.getDataRange().getValues() : [];

    const tasks = [];
    if (tasksData.length > 1) {
      for (let i = 1; i < tasksData.length; i++) {
        const row = tasksData[i];
        if (!row[0] && !row[2]) continue;
        const task = rowToObject(row, CONFIG.TASK_COLUMNS);
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
        projects.push(rowToObject(row, CONFIG.PROJECT_COLUMNS));
      }
    }

    const users = [];
    if (usersData.length > 1) {
      for (let i = 1; i < usersData.length; i++) {
        const row = usersData[i];
        if (!row[0]) continue;
        const user = rowToObject(row, CONFIG.USER_COLUMNS);
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

    try {
      const jsonStr = JSON.stringify(result);
      if (jsonStr.length < 100000) {
        cache.put(cacheKey, jsonStr, 300);
      }
    } catch (e) {
      console.warn('Batch cache write failed:', e);
    }

    return result;
  } catch (error) {
    console.error('getBatchDataFast failed:', error);
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
        console.warn('Cache parse error, fetching fresh data');
      }
    }

    const tasks = getAllTasks();
    const jsonStr = JSON.stringify(tasks);
    if (jsonStr.length < 100000) {
      cache.put(cacheKey, jsonStr, 300);
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
        console.warn('Cache parse error for projects');
      }
    }

    const projects = getAllProjects();
    cache.put(cacheKey, JSON.stringify(projects), 600);
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
          console.warn('Cache parse error for users');
        }
      }
    }

    const users = getActiveUsers();
    cache.put(cacheKey, JSON.stringify(users), 300);
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
    cache.removeAll(['ALL_TASKS_CACHE', 'ALL_PROJECTS_CACHE', 'ACTIVE_USERS_CACHE', 'ALL_USERS_MENTIONS_CACHE']);
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

    if (RequestCache && RequestCache.tasks) {
      RequestCache.tasks = null;
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

    if (RequestCache && RequestCache.projects) {
      RequestCache.projects = null;
    }
  } catch (error) {
    console.error('invalidateProjectCache failed:', error);
  }
}

function invalidateUserCache() {
  try {
    const cache = CacheService.getScriptCache();
    cache.remove('ACTIVE_USERS_CACHE');

    if (RequestCache && RequestCache.users) {
      RequestCache.users = null;
    }
  } catch (error) {
    console.error('invalidateUserCache failed:', error);
  }
}

function getGlobalVersion() {
  try {
    const cache = CacheService.getScriptCache();
    let version = cache.get('GLOBAL_DATA_VERSION');

    if (!version) {
      version = PropertiesService.getScriptProperties().getProperty('GLOBAL_DATA_VERSION') || '0';
      cache.put('GLOBAL_DATA_VERSION', version, 300);
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
    CacheService.getScriptCache().put('GLOBAL_DATA_VERSION', next.toString(), 300);
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

function loadMyBoard(projectId) {
  return getMyBoardOptimized(projectId || null);
}

function loadMasterBoard(projectId) {
  try {
    const allTasks = getAllTasksOptimized();
    const filteredTasks = projectId
      ? allTasks.filter(task => task.projectId === projectId)
      : allTasks;

    const projects = getAllProjectsOptimized();
    const users = getActiveUsersOptimized();
    const board = buildBoardData(filteredTasks, projectId, { view: 'master' });

    return {
      ...board,
      projects: projects,
      users: users,
      taskCount: filteredTasks.length
    };
  } catch (error) {
    console.error('loadMasterBoard failed:', error);
    throw error;
  }
}

function getMyBoardOptimized(projectId, userEmail) {
  try {
    const currentUser = userEmail || getCurrentUserEmailOptimized();

    if (!currentUser) {
      throw new Error('User not authenticated');
    }

    const allTasks = getAllTasksOptimized();
    const userTasks = allTasks.filter(task =>
      task.assignee && task.assignee.toLowerCase() === currentUser.toLowerCase()
    );

    const filteredTasks = projectId
      ? userTasks.filter(task => task.projectId === projectId)
      : userTasks;

    const projects = getAllProjectsOptimized();
    const users = getActiveUsersOptimized();

    const board = buildBoardData(filteredTasks, projectId, {
      view: 'my',
      userEmail: currentUser
    });

    return {
      ...board,
      projects: projects,
      users: users,
      taskCount: filteredTasks.length
    };
  } catch (error) {
    console.error('getMyBoardOptimized failed:', error);
    throw error;
  }
}

function getAnalyticsData(days) {
  try {
    const endDate = new Date();
    const startDate = new Date(endDate.getTime() - (days * 24 * 60 * 60 * 1000));
    const now = new Date();

    const allTasks = getAllTasksOptimized();
    const allProjects = getAllProjectsOptimized();
    const allUsers = getActiveUsersOptimized();

    const metrics = {
      total: 0,
      completed: 0,
      inProgress: 0,
      overdue: 0
    };

    const byStatus = {};
    const byPriority = {};
    const byAssignee = {};
    const projectTaskCount = {};
    const projectCompletedCount = {};
    const weeklyData = {};

    CONFIG.STATUSES.forEach(s => byStatus[s] = 0);
    CONFIG.PRIORITIES.forEach(p => byPriority[p] = 0);

    allProjects.forEach(p => {
      projectTaskCount[p.id] = 0;
      projectCompletedCount[p.id] = 0;
    });

    allTasks.forEach(task => {
      const taskDate = new Date(task.createdAt || task.updatedAt);
      const inDateRange = taskDate >= startDate && taskDate <= endDate;

      if (task.status) {
        byStatus[task.status] = (byStatus[task.status] || 0) + 1;
      }

      if (task.priority) {
        byPriority[task.priority] = (byPriority[task.priority] || 0) + 1;
      }

      const assignee = task.assignee || 'Unassigned';
      byAssignee[assignee] = (byAssignee[assignee] || 0) + 1;

      if (task.projectId) {
        projectTaskCount[task.projectId] = (projectTaskCount[task.projectId] || 0) + 1;
        if (task.status === 'Done') {
          projectCompletedCount[task.projectId] = (projectCompletedCount[task.projectId] || 0) + 1;
        }
      }

      if (inDateRange) {
        metrics.total++;

        if (task.status === 'Done') {
          metrics.completed++;
          const weekKey = getWeekKey(new Date(task.updatedAt || task.createdAt));
          weeklyData[weekKey] = weeklyData[weekKey] || { completed: 0, total: 0 };
          weeklyData[weekKey].completed++;
        } else if (task.status === 'In Progress') {
          metrics.inProgress++;
        }

        if (task.dueDate && task.status !== 'Done') {
          if (new Date(task.dueDate) < now) {
            metrics.overdue++;
          }
        }

        const weekKey = getWeekKey(taskDate);
        weeklyData[weekKey] = weeklyData[weekKey] || { completed: 0, total: 0 };
        weeklyData[weekKey].total++;
      }
    });

    const completionTrend = Object.entries(weeklyData)
      .sort((a, b) => a[0].localeCompare(b[0]))
      .map(([date, data]) => ({
        date: date,
        completed: data.completed,
        total: data.total
      }));

    const teamProductivity = allUsers.map(user => {
      const userTasks = byAssignee[user.email] || 0;
      const userCompleted = allTasks.filter(t =>
        t.assignee === user.email && t.status === 'Done'
      ).length;

      return {
        name: user.name || user.email,
        email: user.email,
        totalTasks: userTasks,
        tasksCompleted: userCompleted,
        productivity: userTasks > 0 ? Math.round((userCompleted / userTasks) * 100) : 0
      };
    }).sort((a, b) => b.productivity - a.productivity);

    const projectHealth = allProjects.map(project => {
      const total = projectTaskCount[project.id] || 0;
      const completed = projectCompletedCount[project.id] || 0;
      const progress = total > 0 ? (completed / total) * 100 : 0;

      let health = 'good';
      if (progress < 30) health = 'poor';
      else if (progress < 60) health = 'at-risk';
      else if (progress >= 90) health = 'excellent';

      return {
        name: project.name,
        progress: Math.round(progress),
        health: health,
        totalTasks: total,
        tasksCompleted: completed
      };
    });

    const priorityDistribution = Object.entries(byPriority).map(([priority, count]) => ({
      priority: priority,
      count: count
    }));

    const recentActivity = getRecentActivity(20).map(activity => ({
      type: activity.action || 'update',
      description: activity.description || `${activity.action || 'Updated'} ${activity.entityType} ${activity.entityId}`,
      user: getNameFromEmail(activity.userId) || activity.userId,
      timestamp: activity.createdAt
    }));

    return {
      metrics: metrics,
      teamProductivity: teamProductivity,
      projectHealth: projectHealth,
      recentActivity: recentActivity,
      completionTrend: completionTrend,
      priorityDistribution: priorityDistribution
    };
  } catch (error) {
    console.error('getAnalyticsData failed:', error);
    throw error;
  }
}

function getWeekKey(date) {
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);
  d.setDate(d.getDate() - d.getDay());
  return d.toISOString().split('T')[0];
}

function getNameFromEmail(email) {
  if (!email) return null;
  const users = getActiveUsersOptimized();
  const user = users.find(u => u.email === email);
  return user ? user.name : email.split('@')[0];
}

function getTimelineData(projectId) {
  const logs = [];

  try {
    if (typeof TimelineEngine === 'undefined') {
      return {
        tasks: [],
        milestones: [],
        dependencies: [],
        criticalPath: [],
        dateRange: { start: new Date(), end: new Date() },
        projectId: projectId,
        error: 'TimelineEngine not loaded',
        logs: logs
      };
    }

    const result = TimelineEngine.generateProjectTimeline(projectId || null);

    if (result) {
      try {
        JSON.stringify(result);
      } catch (serError) {
        return {
          tasks: [],
          milestones: [],
          dependencies: [],
          criticalPath: [],
          dateRange: { start: new Date().toISOString(), end: new Date().toISOString() },
          projectId: projectId,
          error: 'Result serialization failed: ' + serError.message,
          logs: logs
        };
      }

      result.logs = logs;
      return result;
    }

    return {
      tasks: [],
      milestones: [],
      dependencies: [],
      criticalPath: [],
      dateRange: { start: new Date().toISOString(), end: new Date().toISOString() },
      projectId: projectId,
      error: 'TimelineEngine returned null',
      logs: logs
    };
  } catch (error) {
    return {
      tasks: [],
      milestones: [],
      dependencies: [],
      criticalPath: [],
      dateRange: { start: new Date(), end: new Date() },
      projectId: projectId,
      error: error.message,
      logs: logs
    };
  }
}

function getFilteredTimeline(projectId, filters) {
  try {
    if (typeof TimelineEngine === 'undefined') {
      console.error('getFilteredTimeline: TimelineEngine is not defined!');
      return {
        tasks: [],
        milestones: [],
        dependencies: [],
        criticalPath: [],
        dateRange: { start: new Date(), end: new Date() },
        projectId: projectId,
        error: 'TimelineEngine not loaded'
      };
    }

    const timelineData = TimelineEngine.generateProjectTimeline(projectId || null);

    if (!timelineData) {
      console.warn('getFilteredTimeline: TimelineEngine returned null data');
      return {
        tasks: [],
        milestones: [],
        dependencies: [],
        criticalPath: [],
        dateRange: { start: new Date(), end: new Date() },
        projectId: projectId
      };
    }

    if (!timelineData.tasks) {
      timelineData.tasks = [];
    }

    if (!filters || Object.keys(filters).length === 0) {
      return timelineData;
    }

    let tasks = timelineData.tasks;

    if (filters.status && filters.status.length > 0) {
      tasks = tasks.filter(t => filters.status.includes(t.status));
    }

    if (filters.priority && filters.priority.length > 0) {
      tasks = tasks.filter(t => filters.priority.includes(t.priority));
    }

    if (filters.assignee && filters.assignee.length > 0) {
      tasks = tasks.filter(t => filters.assignee.includes(t.assignee));
    }

    if (filters.search) {
      const searchLower = filters.search.toLowerCase();
      tasks = tasks.filter(t =>
        (t.name && t.name.toLowerCase().includes(searchLower)) ||
        (t.id && t.id.toLowerCase().includes(searchLower))
      );
    }

    return {
      ...timelineData,
      tasks: tasks
    };
  } catch (error) {
    console.error('getFilteredTimeline failed:', error);
    return {
      tasks: [],
      milestones: [],
      dependencies: [],
      criticalPath: [],
      dateRange: { start: new Date(), end: new Date() },
      projectId: projectId,
      error: error.message
    };
  }
}

function diagnosticTimeline() {
  const result = {
    timelineEngineAvailable: typeof TimelineEngine !== 'undefined',
    timelineEngineType: typeof TimelineEngine,
    getAllTasksAvailable: typeof getAllTasks !== 'undefined',
    sampleTaskCount: 0,
    sampleTask: null,
    errors: []
  };

  try {
    const tasks = getAllTasks();
    result.sampleTaskCount = tasks.length;

    if (tasks.length > 0) {
      const sample = tasks[0];
      result.sampleTask = {
        id: sample.id,
        title: sample.title,
        startDate: sample.startDate,
        startDateType: typeof sample.startDate,
        dueDate: sample.dueDate,
        dueDateType: typeof sample.dueDate,
        createdAt: sample.createdAt
      };
    }

    if (typeof TimelineEngine !== 'undefined' && tasks.length > 0) {
      const testDate = '1/16/2026';
      const parsed = TimelineEngine.parseFlexibleDate(testDate);
      result.dateParseTest = {
        input: testDate,
        output: parsed,
        outputType: typeof parsed,
        isDate: parsed instanceof Date,
        isValid: parsed && !isNaN(parsed.getTime())
      };
    }
  } catch (error) {
    result.errors.push(error.message);
  }

  return result;
}

function getMilestonesForView() {
  try {
    const milestones = getMilestones();
    const projects = getAllProjectsOptimized();

    return {
      success: true,
      milestones: milestones,
      projects: projects
    };
  } catch (error) {
    console.error('getMilestonesForView failed:', error);
    return {
      success: false,
      error: error.message,
      milestones: [],
      projects: []
    };
  }
}

function saveNewTask(taskData) {
  const result = createTask(taskData);
  invalidateTaskCache(result.id, 'create');
  return result;
}

function saveTaskUpdate(taskId, updates) {
  const result = updateTask(taskId, updates);
  invalidateTaskCache(taskId, 'update');
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
      console.warn('loadTask: Invalid taskId provided:', taskId);
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

function saveNewProject(projectData) {
  const result = createProject(projectData);
  invalidateProjectCache();
  return result;
}

function loadProjects() {
  return getAllProjectsOptimized();
}

function loadUsers() {
  return getActiveUsersOptimized();
}

function getAllUsersForMentions(forceRefresh) {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'ALL_USERS_MENTIONS_CACHE';

    if (!forceRefresh) {
      const cached = cache.get(cacheKey);
      if (cached) {
        try {
          return JSON.parse(cached);
        } catch (e) {
          console.warn('Cache parse error for mentions users');
        }
      }
    }

    const allUsers = getAllUsers();
    const users = allUsers
      .filter(u => u.email)
      .map(u => ({
        email: u.email,
        name: u.name || u.email.split('@')[0],
        active: u.active
      }));

    cache.put(cacheKey, JSON.stringify(users), 300);
    return users;
  } catch (error) {
    console.error('getAllUsersForMentions failed:', error);
    const allUsers = getAllUsers();
    return allUsers
      .filter(u => u.email)
      .map(u => ({
        email: u.email,
        name: u.name || u.email.split('@')[0],
        active: u.active
      }));
  }
}

function loadComments(taskId) {
  if (!taskId) {
    console.error('loadComments: No taskId provided');
    return [];
  }

  const comments = getCommentsForTask(taskId);

  return comments.map(comment => {
    const mentionedUsers = comment.mentionedUsers
      ? (typeof comment.mentionedUsers === 'string'
        ? comment.mentionedUsers.split(',').map(e => e.trim()).filter(e => e)
        : comment.mentionedUsers)
      : [];

    const commenter = getUserByEmail(comment.userId);
    const userName = commenter ? commenter.name : comment.userId;

    return {
      ...comment,
      userName: userName,
      formattedContent: formatCommentContent(comment.content, mentionedUsers),
      mentionedUsers: mentionedUsers
    };
  });
}

function saveComment(taskId, content) {
  try {
    const result = addCommentWithMentions(taskId, content);
    return {
      success: true,
      comment: result,
      mentionCount: result.mentionedUsers ? result.mentionedUsers.length : 0,
      notificationCount: result.notifications ? result.notifications.length : 0
    };
  } catch (error) {
    console.error('saveComment: Error - ' + error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

function formatCommentContent(content, mentionedUserEmails) {
  if (!content) return '';

  let formatted = content;

  formatted = formatted.replace(/@\[([^\]]+)\]\([^)]+\)/g, function(match, name) {
    return '<span class="mention">@' + name + '</span>';
  });

  mentionedUserEmails.forEach(email => {
    const user = getUserByEmail(email);
    const displayName = user ? user.name : email;
    const regex = new RegExp('@' + email.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
    formatted = formatted.replace(regex, '<span class="mention">@' + displayName + '</span>');
  });

  formatted = formatted.replace(/@([A-Za-z][A-Za-z0-9 ]*[A-Za-z0-9])(?![^<]*<\/span>)(?=\s|$|[.,!?])/g, function(match, name) {
    return '<span class="mention">@' + name + '</span>';
  });

  return formatted;
}

function getCurrentUserEmailOptimized() {
  try {
    const email = PropertiesService.getScriptProperties().getProperty('CURRENT_USER_EMAIL');
    const timestamp = PropertiesService.getScriptProperties().getProperty('LOGIN_TIMESTAMP');

    if (email && timestamp) {
      const loginTime = parseInt(timestamp);
      const now = new Date().getTime();
      const sessionDuration = 24 * 60 * 60 * 1000;

      if (now - loginTime < sessionDuration) {
        return email;
      }
    }

    return null;
  } catch (error) {
    console.error('getCurrentUserEmailOptimized failed:', error);
    return null;
  }
}

function setCurrentUserEmailOptimized(email) {
  try {
    PropertiesService.getScriptProperties().setProperty('CURRENT_USER_EMAIL', email);
    PropertiesService.getScriptProperties().setProperty('LOGIN_TIMESTAMP', new Date().getTime().toString());
  } catch (error) {
    console.error('setCurrentUserEmailOptimized failed:', error);
  }
}

function getUserByEmailOptimized(email) {
  try {
    const sheet = getUsersSheet();
    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) return null;

    const headers = data[0];
    const emailIndex = headers.indexOf('email');

    if (emailIndex === -1) return null;

    for (let i = 1; i < data.length; i++) {
      if (data[i][emailIndex] && data[i][emailIndex].toLowerCase() === email.toLowerCase()) {
        const user = {};
        headers.forEach((header, index) => {
          user[header] = data[i][index];
        });
        return user;
      }
    }

    return null;
  } catch (error) {
    console.error('getUserByEmailOptimized failed:', error);
    return null;
  }
}

function getInitialData() {
  try {
    const userEmail = getCurrentUserEmailOptimized();

    if (!userEmail) {
      throw new Error('No active session. Please login.');
    }

    const user = getUserByEmailOptimized(userEmail);

    if (!user) {
      throw new Error('User session invalid. Please login again.');
    }

    const boardData = getMyBoardOptimized(null, userEmail);

    return {
      user: user,
      board: boardData,
      config: {
        statuses: CONFIG.STATUSES,
        priorities: CONFIG.PRIORITIES,
        types: CONFIG.TYPES,
        colors: CONFIG.COLORS
      }
    };
  } catch (error) {
    console.error('getInitialData failed:', error);
    throw error;
  }
}

function getInitialDataFast() {
  try {
    const userEmail = getCurrentUserEmailOptimized();

    if (!userEmail) {
      throw new Error('No active session. Please login.');
    }

    const user = getUserByEmailOptimized(userEmail);

    if (!user) {
      throw new Error('User session invalid. Please login again.');
    }

    const batchData = getBatchDataFast();

    const userTasks = batchData.tasks.filter(task =>
      task.assignee && task.assignee.toLowerCase() === userEmail.toLowerCase()
    );

    const columns = CONFIG.STATUSES.map(status => ({
      id: status.toLowerCase().replace(/\s+/g, '-'),
      title: status,
      color: CONFIG.COLORS[status] || '#6B7280',
      tasks: userTasks.filter(t => t.status === status)
    }));

    const total = userTasks.length;
    const completed = userTasks.filter(t => t.status === 'Done').length;
    const now = new Date();

    const dueSoon = userTasks.filter(t => {
      if (!t.dueDate || t.status === 'Done') return false;
      const due = new Date(t.dueDate);
      const diff = (due - now) / (1000 * 60 * 60 * 24);
      return diff >= 0 && diff <= 3;
    }).length;

    const overdue = userTasks.filter(t => {
      if (!t.dueDate || t.status === 'Done') return false;
      return new Date(t.dueDate) < now;
    }).length;

    return {
      user: user,
      board: {
        columns: columns,
        projects: batchData.projects,
        users: batchData.users,
        stats: {
          total: total,
          completed: completed,
          progress: total > 0 ? Math.round((completed / total) * 100) : 0,
          dueSoon: dueSoon,
          overdue: overdue
        },
        taskCount: userTasks.length
      },
      config: batchData.config
    };
  } catch (error) {
    console.error('getInitialDataFast failed:', error);
    throw error;
  }
}

function getListViewData() {
  try {
    const tasks = getAllTasksOptimized();
    const projects = getAllProjectsOptimized();
    const users = getActiveUsersOptimized();

    return {
      success: true,
      tasks: tasks,
      projects: projects,
      users: users
    };
  } catch (error) {
    console.error('getListViewData failed:', error);
    return {
      success: false,
      error: error.message,
      tasks: [],
      projects: [],
      users: []
    };
  }
}

function getCalendarData() {
  try {
    const tasks = getAllTasksOptimized();
    const projects = getAllProjectsOptimized();
    const users = getActiveUsersOptimized();

    return {
      success: true,
      tasks: tasks,
      projects: projects,
      users: users
    };
  } catch (error) {
    console.error('getCalendarData failed:', error);
    return {
      success: false,
      error: error.message,
      tasks: [],
      projects: [],
      users: []
    };
  }
}

function loginWithEmail(email) {
  try {
    if (!email || !email.includes('@')) {
      throw new Error('Please enter a valid email address');
    }

    email = email.toLowerCase().trim();
    const user = getUserByEmailOptimized(email);

    if (!user) {
      throw new Error('User not found. Please contact your administrator for access.');
    }

    setCurrentUserEmailOptimized(email);
    const boardData = getMyBoardOptimized(null, email);

    return {
      user: user,
      board: boardData,
      config: {
        statuses: CONFIG.STATUSES,
        priorities: CONFIG.PRIORITIES,
        types: CONFIG.TYPES,
        colors: CONFIG.COLORS
      }
    };
  } catch (error) {
    console.error('loginWithEmail failed:', error);
    throw error;
  }
}

function logout() {
  try {
    const userEmail = getCurrentUserEmailOptimized();

    if (typeof AuthService !== 'undefined' && userEmail) {
      AuthService.invalidateAllUserSessions(userEmail);
    }

    PropertiesService.getScriptProperties().deleteProperty('CURRENT_USER_EMAIL');
    PropertiesService.getScriptProperties().deleteProperty('LOGIN_TIMESTAMP');
    clearAllCaches();

    return { success: true };
  } catch (error) {
    console.error('logout failed:', error);
    throw error;
  }
}

function loginWithPassword(email, password, options) {
  try {
    if (typeof AuthService === 'undefined') {
      throw new Error('AuthService not available');
    }

    const result = AuthService.authenticate(email, password, options || {});

    if (!result.success) {
      return result;
    }

    setCurrentUserEmailOptimized(email);
    const boardData = getMyBoardOptimized(null, email);

    return {
      success: true,
      user: result.user,
      session: result.session,
      board: boardData,
      config: {
        statuses: CONFIG.STATUSES,
        priorities: CONFIG.PRIORITIES,
        types: CONFIG.TYPES,
        colors: CONFIG.COLORS
      }
    };
  } catch (error) {
    console.error('loginWithPassword failed:', error);
    return {
      success: false,
      error: error.message || 'Authentication failed'
    };
  }
}

function setUserPassword(email, newPassword, currentPassword) {
  try {
    if (typeof AuthService === 'undefined') {
      throw new Error('AuthService not available');
    }

    if (!email) {
      email = getCurrentUserEmailOptimized();
    }

    const currentUserEmail = getCurrentUserEmailOptimized();

    if (currentUserEmail !== email) {
      const currentUser = getUserByEmailOptimized(currentUserEmail);
      if (!currentUser || currentUser.role !== 'admin') {
        return { success: false, error: 'Permission denied' };
      }
    }

    return AuthService.setPassword(email, newPassword, currentPassword);
  } catch (error) {
    console.error('setUserPassword failed:', error);
    return { success: false, error: error.message };
  }
}

function validateSessionToken(token, userId) {
  try {
    if (typeof AuthService === 'undefined') {
      const email = getCurrentUserEmailOptimized();
      return {
        valid: email === userId,
        reason: email ? 'Legacy session' : 'No session'
      };
    }

    return AuthService.validateSession(token, userId);
  } catch (error) {
    console.error('validateSessionToken failed:', error);
    return { valid: false, reason: error.message };
  }
}

function getLoginRateLimitStatus(email) {
  try {
    if (typeof AuthService === 'undefined') {
      return { allowed: true, remainingAttempts: 5 };
    }

    return AuthService.checkRateLimit(email);
  } catch (error) {
    console.error('getLoginRateLimitStatus failed:', error);
    return { allowed: true, remainingAttempts: 5 };
  }
}

function logoutAllDevices() {
  try {
    const userEmail = getCurrentUserEmailOptimized();

    if (!userEmail) {
      return { success: false, error: 'Not logged in' };
    }

    if (typeof AuthService !== 'undefined') {
      AuthService.invalidateAllUserSessions(userEmail);
    }

    PropertiesService.getScriptProperties().deleteProperty('CURRENT_USER_EMAIL');
    PropertiesService.getScriptProperties().deleteProperty('LOGIN_TIMESTAMP');
    clearAllCaches();

    return { success: true };
  } catch (error) {
    console.error('logoutAllDevices failed:', error);
    return { success: false, error: error.message };
  }
}

function requestMfaCode(email) {
  try {
    if (typeof AuthService === 'undefined') {
      return { success: false, error: 'MFA not available' };
    }

    const user = getUserByEmailOptimized(email);

    if (!user) {
      return { success: false, error: 'User not found' };
    }

    AuthService.sendMfaCode(user);
    return { success: true, message: 'MFA code sent to your email' };
  } catch (error) {
    console.error('requestMfaCode failed:', error);
    return { success: false, error: error.message };
  }
}

function setMfaEnabled(enabled) {
  try {
    const userEmail = getCurrentUserEmailOptimized();

    if (!userEmail) {
      return { success: false, error: 'Not logged in' };
    }

    updateUser(userEmail, { mfaEnabled: enabled });
    return { success: true };
  } catch (error) {
    console.error('setMfaEnabled failed:', error);
    return { success: false, error: error.message };
  }
}

function getUserOrganizations() {
  try {
    const userEmail = getCurrentUserEmailOptimized();

    if (!userEmail) {
      return { success: false, error: 'Not logged in', organizations: [] };
    }

    const user = getUserByEmailOptimized(userEmail);
    const orgs = getAllOrganizations();
    const orgMembers = [];

    orgs.forEach(org => {
      const members = getOrganizationMembers(org.id);
      if (members.some(m => m.userId === userEmail)) {
        orgMembers.push({
          ...org,
          memberRole: members.find(m => m.userId === userEmail)?.orgRole
        });
      }
    });

    return {
      success: true,
      organizations: orgMembers,
      currentOrgId: user?.organizationId
    };
  } catch (error) {
    console.error('getUserOrganizations failed:', error);
    return { success: false, error: error.message, organizations: [] };
  }
}

function getUserTeams() {
  try {
    const userEmail = getCurrentUserEmailOptimized();

    if (!userEmail) {
      return { success: false, error: 'Not logged in', teams: [] };
    }

    const user = getUserByEmailOptimized(userEmail);
    const teams = getTeamsForUser(userEmail);

    return {
      success: true,
      teams: teams.map(team => ({
        ...team,
        memberCount: getTeamMembers(team.id).length
      })),
      currentTeamId: user?.teamId
    };
  } catch (error) {
    console.error('getUserTeams failed:', error);
    return { success: false, error: error.message, teams: [] };
  }
}

function createNewOrganization(orgData) {
  try {
    const org = createOrganization(orgData);
    return { success: true, organization: org };
  } catch (error) {
    console.error('createNewOrganization failed:', error);
    return { success: false, error: error.message };
  }
}

function createNewTeam(teamData) {
  try {
    const team = createTeam(teamData);
    return { success: true, team: team };
  } catch (error) {
    console.error('createNewTeam failed:', error);
    return { success: false, error: error.message };
  }
}

function getOrganizationDetails(orgId) {
  try {
    const org = getOrganizationById(orgId);

    if (!org) {
      return { success: false, error: 'Organization not found' };
    }

    const members = getOrganizationMembers(orgId);
    const teams = getAllTeams(orgId);

    const enrichedMembers = members.map(m => {
      const user = getUserByEmail(m.userId);
      return {
        ...m,
        name: user?.name || m.userId,
        email: m.userId,
        role: user?.role,
        active: user?.active
      };
    });

    return {
      success: true,
      organization: org,
      members: enrichedMembers,
      teams: teams
    };
  } catch (error) {
    console.error('getOrganizationDetails failed:', error);
    return { success: false, error: error.message };
  }
}

function getTeamDetails(teamId) {
  try {
    const team = getTeamById(teamId);

    if (!team) {
      return { success: false, error: 'Team not found' };
    }

    const members = getTeamMembers(teamId);

    const enrichedMembers = members.map(m => {
      const user = getUserByEmail(m.userId);
      return {
        ...m,
        name: user?.name || m.userId,
        email: m.userId,
        role: user?.role,
        active: user?.active
      };
    });

    return {
      success: true,
      team: team,
      members: enrichedMembers
    };
  } catch (error) {
    console.error('getTeamDetails failed:', error);
    return { success: false, error: error.message };
  }
}

function addOrgMember(orgId, userEmail, role) {
  try {
    const member = addOrganizationMember(orgId, userEmail, role);
    return { success: true, member: member };
  } catch (error) {
    console.error('addOrgMember failed:', error);
    return { success: false, error: error.message };
  }
}

function addTeamMemberToTeam(teamId, userEmail, role) {
  try {
    const member = addTeamMember(teamId, userEmail, role);
    return { success: true, member: member };
  } catch (error) {
    console.error('addTeamMemberToTeam failed:', error);
    return { success: false, error: error.message };
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

function getAutomationRules(projectId) {
  try {
    if (typeof AutomationEngine === 'undefined') {
      return { success: false, error: 'AutomationEngine not available', rules: [] };
    }

    const rules = AutomationEngine.getRules(projectId);
    return { success: true, rules: rules };
  } catch (error) {
    console.error('getAutomationRules failed:', error);
    return { success: false, error: error.message, rules: [] };
  }
}

function createAutomationRule(ruleData) {
  try {
    if (typeof AutomationEngine === 'undefined') {
      return { success: false, error: 'AutomationEngine not available' };
    }

    const rule = AutomationEngine.createRule(ruleData);
    return { success: true, rule: rule };
  } catch (error) {
    console.error('createAutomationRule failed:', error);
    return { success: false, error: error.message };
  }
}

function updateAutomationRule(ruleId, updates) {
  try {
    if (typeof AutomationEngine === 'undefined') {
      return { success: false, error: 'AutomationEngine not available' };
    }

    const rule = AutomationEngine.updateRule(ruleId, updates);
    return { success: true, rule: rule };
  } catch (error) {
    console.error('updateAutomationRule failed:', error);
    return { success: false, error: error.message };
  }
}

function deleteAutomationRule(ruleId) {
  try {
    if (typeof AutomationEngine === 'undefined') {
      return { success: false, error: 'AutomationEngine not available' };
    }

    const success = AutomationEngine.deleteRule(ruleId);
    return { success: success };
  } catch (error) {
    console.error('deleteAutomationRule failed:', error);
    return { success: false, error: error.message };
  }
}

function getUserWorkload(userEmail, options) {
  try {
    if (typeof CapacityEngine === 'undefined') {
      return { success: false, error: 'CapacityEngine not available' };
    }

    const email = userEmail || getCurrentUserEmailOptimized();
    const workload = CapacityEngine.calculateUserWorkload(email, options || {});
    return { success: true, workload: workload };
  } catch (error) {
    console.error('getUserWorkload failed:', error);
    return { success: false, error: error.message };
  }
}

function getTeamWorkload(options) {
  try {
    if (typeof CapacityEngine === 'undefined') {
      return { success: false, error: 'CapacityEngine not available' };
    }

    const report = CapacityEngine.getTeamWorkloadReport(options || {});
    return { success: true, report: report };
  } catch (error) {
    console.error('getTeamWorkload failed:', error);
    return { success: false, error: error.message };
  }
}

function setCapacity(capacityData) {
  try {
    if (typeof CapacityEngine === 'undefined') {
      return { success: false, error: 'CapacityEngine not available' };
    }

    const config = CapacityEngine.setUserCapacity(capacityData);
    return { success: true, config: config };
  } catch (error) {
    console.error('setCapacity failed:', error);
    return { success: false, error: error.message };
  }
}

function findAvailableAssignees(options) {
  try {
    if (typeof CapacityEngine === 'undefined') {
      const users = getActiveUsersOptimized();
      return {
        success: true,
        users: users.map(u => ({ userId: u.email, userName: u.name, available: true }))
      };
    }

    const users = CapacityEngine.findAvailableUsers(options || {});
    return { success: true, users: users };
  } catch (error) {
    console.error('findAvailableAssignees failed:', error);
    return { success: false, error: error.message, users: [] };
  }
}

function getTaskSlaStatus(taskId) {
  try {
    if (typeof SLAEngine === 'undefined') {
      return { success: false, error: 'SLAEngine not available' };
    }

    const task = getTaskById(taskId);

    if (!task) {
      return { success: false, error: 'Task not found' };
    }

    const slaStatus = SLAEngine.calculateTaskSlaStatus(task);
    return { success: true, slaStatus: slaStatus };
  } catch (error) {
    console.error('getTaskSlaStatus failed:', error);
    return { success: false, error: error.message };
  }
}

function getSlaComplianceReport(options) {
  try {
    if (typeof SLAEngine === 'undefined') {
      return { success: false, error: 'SLAEngine not available' };
    }

    const report = SLAEngine.getSlaComplianceReport(options || {});
    return { success: true, report: report };
  } catch (error) {
    console.error('getSlaComplianceReport failed:', error);
    return { success: false, error: error.message };
  }
}

function getSlaConfigs() {
  try {
    if (typeof SLAEngine === 'undefined') {
      return { success: false, error: 'SLAEngine not available', configs: [] };
    }

    const configs = SLAEngine.getSlaConfigs();
    return { success: true, configs: configs };
  } catch (error) {
    console.error('getSlaConfigs failed:', error);
    return { success: false, error: error.message, configs: [] };
  }
}

function getTriageQueue(filters) {
  try {
    if (typeof TriageEngine === 'undefined') {
      return { success: false, error: 'TriageEngine not available', items: [] };
    }

    const items = TriageEngine.getTriageQueue(filters || {});
    return { success: true, items: items };
  } catch (error) {
    console.error('getTriageQueue failed:', error);
    return { success: false, error: error.message, items: [] };
  }
}

function assignTriageItem(itemId, taskData) {
  try {
    if (typeof TriageEngine === 'undefined') {
      return { success: false, error: 'TriageEngine not available' };
    }

    const task = TriageEngine.assignTriageItem(itemId, taskData);
    invalidateTaskCache(task?.id, 'create');
    return { success: true, task: task };
  } catch (error) {
    console.error('assignTriageItem failed:', error);
    return { success: false, error: error.message };
  }
}

function rejectTriageItem(itemId, reason) {
  try {
    if (typeof TriageEngine === 'undefined') {
      return { success: false, error: 'TriageEngine not available' };
    }

    const item = TriageEngine.rejectTriageItem(itemId, reason);
    return { success: true, item: item };
  } catch (error) {
    console.error('rejectTriageItem failed:', error);
    return { success: false, error: error.message };
  }
}

function submitExternalTicket(ticketData) {
  try {
    if (typeof TriageEngine === 'undefined') {
      return { success: false, error: 'TriageEngine not available' };
    }

    const item = TriageEngine.submitTicket(ticketData);
    return { success: true, item: item };
  } catch (error) {
    console.error('submitExternalTicket failed:', error);
    return { success: false, error: error.message };
  }
}

function getNotifications() {
  try {
    const userEmail = getCurrentUserEmailOptimized();
    if (!userEmail) return [];
    return getNotificationsForUser(userEmail, 50);
  } catch (error) {
    console.error('getNotifications failed:', error);
    return [];
  }
}

function getNotificationsWithCount() {
  try {
    const userEmail = getCurrentUserEmailOptimized();
    if (!userEmail) return { notifications: [], unreadCount: 0 };

    const notifications = getNotificationsForUser(userEmail, 50);
    const unreadCount = notifications.filter(n => !n.read).length;

    return {
      notifications: notifications,
      unreadCount: unreadCount
    };
  } catch (error) {
    console.error('getNotificationsWithCount failed:', error);
    return { notifications: [], unreadCount: 0 };
  }
}

function getUnreadNotificationCount() {
  try {
    const userEmail = getCurrentUserEmailOptimized();
    if (!userEmail) return 0;

    const notifications = getNotificationsForUser(userEmail, 100, true);
    return notifications.length;
  } catch (error) {
    console.error('getUnreadNotificationCount failed:', error);
    return 0;
  }
}

function markNotificationRead(notificationId) {
  return markNotificationAsRead(notificationId);
}

function markAllNotificationsRead() {
  try {
    const userEmail = getCurrentUserEmailOptimized();
    if (!userEmail) return { success: false, error: 'Not authenticated' };

    const sheet = getNotificationsSheet();
    const data = sheet.getDataRange().getValues();
    const columns = CONFIG.NOTIFICATION_COLUMNS;
    const userIdIndex = columns.indexOf('userId');
    const readIndex = columns.indexOf('read');
    let markedCount = 0;

    for (let i = 1; i < data.length; i++) {
      if (data[i][userIdIndex] === userEmail && data[i][readIndex] !== true) {
        sheet.getRange(i + 1, readIndex + 1).setValue(true);
        markedCount++;
      }
    }

    return { success: true, markedCount: markedCount };
  } catch (error) {
    console.error('markAllNotificationsRead failed:', error);
    return { success: false, error: error.message };
  }
}

function updateNotificationPreferences(preferences) {
  const userEmail = getCurrentUserEmailOptimized();
  if (!userEmail) throw new Error('Not authenticated');
  return saveNotificationPreferences(userEmail, preferences);
}

function saveNotificationPreferences(userEmail, preferences) {
  try {
    const sheet = getUserPreferencesSheet();
    const data = sheet.getDataRange().getValues();
    const columns = CONFIG.USER_PREFERENCE_COLUMNS;
    let found = false;

    for (let i = 1; i < data.length; i++) {
      const pref = rowToObject(data[i], columns);
      if (pref.userId === userEmail && pref.category === 'notifications') {
        sheet.getRange(i + 1, columns.indexOf('settings') + 1).setValue(JSON.stringify(preferences));
        sheet.getRange(i + 1, columns.indexOf('updatedAt') + 1).setValue(now());
        found = true;
        break;
      }
    }

    if (!found) {
      const newPref = {
        id: generateId('pref'),
        userId: userEmail,
        category: 'notifications',
        settings: JSON.stringify(preferences),
        updatedAt: now()
      };
      sheet.appendRow(objectToRow(newPref, columns));
    }

    return { success: true };
  } catch (error) {
    console.error('saveNotificationPreferences failed:', error);
    return { success: false, error: error.message };
  }
}

function getNotificationPreferences(userEmail) {
  try {
    userEmail = userEmail || getCurrentUserEmailOptimized();
    if (!userEmail) return {};

    const sheet = getUserPreferencesSheet();
    const data = sheet.getDataRange().getValues();
    const columns = CONFIG.USER_PREFERENCE_COLUMNS;

    for (let i = 1; i < data.length; i++) {
      const pref = rowToObject(data[i], columns);
      if (pref.userId === userEmail && pref.category === 'notifications') {
        try {
          return JSON.parse(pref.settings);
        } catch (e) {
          return {};
        }
      }
    }

    return {
      mention: { email: true, in_app: true },
      task_assigned: { email: true, in_app: true },
      task_updated: { email: false, in_app: true },
      deadline_approaching: { email: true, in_app: true },
      comment_added: { email: false, in_app: true }
    };
  } catch (error) {
    console.error('getNotificationPreferences failed:', error);
    return {};
  }
}

function getUserSuggestions(query, excludeUsers) {
  try {
    return MentionEngine.getUserSuggestions(query, excludeUsers || []);
  } catch (error) {
    console.error('getUserSuggestions failed:', error);
    return getBasicUserSuggestions(query, excludeUsers || []);
  }
}

function getBasicUserSuggestions(query, excludeUsers) {
  if (!query || query.length < 1) {
    return [];
  }

  const activeUsers = getActiveUsersOptimized();
  const queryLower = query.toLowerCase();
  const excludeSet = new Set((excludeUsers || []).map(email => email.toLowerCase()));

  const suggestions = activeUsers
    .filter(user => {
      if (excludeSet.has(user.email.toLowerCase())) {
        return false;
      }
      const emailMatch = user.email.toLowerCase().includes(queryLower);
      const nameMatch = user.name && user.name.toLowerCase().includes(queryLower);
      return emailMatch || nameMatch;
    })
    .map(user => ({
      email: user.email,
      name: user.name || user.email.split('@')[0],
      displayText: `${user.name || user.email.split('@')[0]} (${user.email})`
    }))
    .slice(0, 10);

  return suggestions;
}

function addDependency(successorId, predecessorId, dependencyType, lag) {
  try {
    const result = DependencyEngine.addDependency(successorId, predecessorId, dependencyType || 'finish_to_start', lag || 0);
    return {
      success: true,
      dependency: result
    };
  } catch (error) {
    console.error('addDependency failed:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

function removeDependency(dependencyId) {
  try {
    DependencyEngine.removeDependency(dependencyId);
    return {
      success: true
    };
  } catch (error) {
    console.error('removeDependency failed:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

function getTaskDependenciesWithDetails(taskId) {
  try {
    const dependencies = DependencyEngine.getTaskDependencies(taskId);

    const result = {
      predecessors: dependencies.predecessors.map(dep => {
        const task = getTaskById(dep.predecessorId);
        return {
          id: dep.id,
          predecessorId: dep.predecessorId,
          dependencyType: dep.dependencyType,
          lag: dep.lag,
          task: task ? {
            id: task.id,
            title: task.title,
            status: task.status,
            priority: task.priority,
            assignee: task.assignee
          } : null
        };
      }),
      successors: dependencies.successors.map(dep => {
        const task = getTaskById(dep.successorId);
        return {
          id: dep.id,
          successorId: dep.successorId,
          dependencyType: dep.dependencyType,
          lag: dep.lag,
          task: task ? {
            id: task.id,
            title: task.title,
            status: task.status,
            priority: task.priority,
            assignee: task.assignee
          } : null
        };
      })
    };

    return result;
  } catch (error) {
    console.error('getTaskDependenciesWithDetails failed:', error);
    return {
      predecessors: [],
      successors: []
    };
  }
}

function getTasksForDependencyPicker(currentTaskId) {
  try {
    const allTasks = getAllTasks();
    const projects = getAllProjects();
    const availableTasks = allTasks.filter(task => task.id !== currentTaskId);

    return {
      success: true,
      tasks: availableTasks,
      projects: projects
    };
  } catch (error) {
    console.error('getTasksForDependencyPicker failed:', error);
    return {
      success: false,
      error: error.message,
      tasks: [],
      projects: []
    };
  }
}

function calculateCriticalPathForProject(projectId) {
  try {
    const result = DependencyEngine.calculateCriticalPath(projectId);
    return {
      success: true,
      criticalPath: result.criticalPath || [],
      totalDuration: result.totalDuration || 0,
      projectStart: result.projectStart,
      projectEnd: result.projectEnd
    };
  } catch (error) {
    console.error('calculateCriticalPathForProject failed:', error);
    return {
      success: false,
      error: error.message,
      criticalPath: [],
      totalDuration: 0
    };
  }
}

function canTaskStart(taskId) {
  try {
    const result = DependencyEngine.canStartTask(taskId);
    return {
      success: true,
      canStart: result.canStart,
      blockingTasks: result.blockingTasks || []
    };
  } catch (error) {
    console.error('canTaskStart failed:', error);
    return {
      success: false,
      error: error.message,
      canStart: true,
      blockingTasks: []
    };
  }
}

function getBlockedTasksForProject(projectId) {
  try {
    return DependencyEngine.getBlockedTasks(projectId);
  } catch (error) {
    console.error('getBlockedTasksForProject failed:', error);
    return [];
  }
}
