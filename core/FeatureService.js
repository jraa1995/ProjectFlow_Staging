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

function syncRequestsFromWorkbook(workbookId) {
  try {
    const result = importFromRequestsWorkbook(workbookId || '');
    return result;
  } catch (error) {
    console.error('syncRequestsFromWorkbook failed:', error);
    return { success: false, error: error.message, stagedCount: 0, skippedCount: 0 };
  }
}

function getStoredRequestsWorkbookId() {
  try {
    return { success: true, workbookId: getRequestsWorkbookId() };
  } catch (error) {
    console.error('getStoredRequestsWorkbookId failed:', error);
    return { success: false, workbookId: '' };
  }
}

function saveRequestsWorkbookId(workbookId) {
  try {
    setRequestsWorkbookId(workbookId);
    return { success: true };
  } catch (error) {
    console.error('saveRequestsWorkbookId failed:', error);
    return { success: false, error: error.message };
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

function addDependency(successorId, predecessorId, dependencyType, lag) {
  try {
    const result = DependencyEngine.addDependency(successorId, predecessorId, dependencyType || 'finish_to_start', lag || 0);
    invalidateDependencyCache();
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
    invalidateDependencyCache();
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

function toggleCalendarSync(taskId, enabled) {
  try {
    const task = getTaskById(taskId);
    if (!task) return { success: false, error: 'Task not found' };
    if (enabled && task.dueDate) {
      syncTaskToCalendar(task);
      return { success: true, synced: true };
    } else if (!enabled) {
      removeCalendarEvent(task);
      return { success: true, synced: false };
    }
    return { success: true, synced: false };
  } catch (error) {
    console.error('toggleCalendarSync failed:', error);
    return { success: false, error: error.message };
  }
}

function processRecurringTasksWrapper() {
  try {
    processRecurringTasks();
    return { success: true };
  } catch (error) {
    console.error('processRecurringTasksWrapper failed:', error);
    return { success: false, error: error.message };
  }
}

function setupRecurringTrigger() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const existing = triggers.find(t => t.getHandlerFunction() === 'processRecurringTasksWrapper');
    if (existing) return { success: true, message: 'Trigger already exists' };
    ScriptApp.newTrigger('processRecurringTasksWrapper')
      .timeBased()
      .everyHours(24)
      .create();
    return { success: true, message: 'Daily recurring trigger created' };
  } catch (error) {
    console.error('setupRecurringTrigger failed:', error);
    return { success: false, error: error.message };
  }
}

function logTimeEntry(taskId, entry) {
  try {
    var task = getTaskById(taskId);
    if (!task) return { success: false, error: 'Task not found' };
    var entries = [];
    try { entries = task.timeEntries ? JSON.parse(task.timeEntries) : []; } catch (e) { entries = []; }
    entries.push({
      start: entry.start,
      end: entry.end,
      hours: entry.hours || 0,
      user: getCurrentUserEmail()
    });
    var newActualHrs = Math.round((parseFloat(task.actualHrs || 0) + parseFloat(entry.hours || 0)) * 100) / 100;
    updateTask(taskId, {
      timeEntries: JSON.stringify(entries),
      actualHrs: newActualHrs
    });
    invalidateTaskCache(taskId, 'update');
    return { success: true, actualHrs: newActualHrs };
  } catch (error) {
    console.error('logTimeEntry failed:', error);
    return { success: false, error: error.message };
  }
}

function exportViewToSheets(viewType, filters) {
  try {
    filters = filters || {};
    const allTasks = getAllTasks({}, { skipPermissionCheck: false });
    let tasks = allTasks;
    if (filters.projectId) {
      tasks = tasks.filter(t => t.projectId === filters.projectId);
    }
    if (filters.status) {
      tasks = tasks.filter(t => t.status === filters.status);
    }
    if (filters.assignee) {
      tasks = tasks.filter(t => t.assignee && t.assignee.toLowerCase() === filters.assignee.toLowerCase());
    }
    if (filters.sprint) {
      tasks = tasks.filter(t => t.sprint === filters.sprint);
    }
    if (viewType === 'board') {
      tasks = tasks.filter(t => !t.isMilestone);
    } else if (viewType === 'sprint') {
      tasks = tasks.filter(t => t.sprint && t.sprint !== '');
    }
    const headers = ['ID', 'Project', 'Title', 'Status', 'Priority', 'Type', 'Assignee', 'Reporter', 'Due Date', 'Sprint', 'Story Points', 'Labels', 'Created At'];
    const rows = tasks.map(t => [
      t.id || '',
      t.projectId || '',
      t.title || '',
      t.status || '',
      t.priority || '',
      t.type || '',
      t.assignee || '',
      t.reporter || '',
      t.dueDate || '',
      t.sprint || '',
      t.storyPoints || 0,
      Array.isArray(t.labels) ? t.labels.join(', ') : (t.labels || ''),
      t.createdAt || ''
    ]);
    const exportTitle = 'COLONY Export - ' + (viewType || 'list') + ' - ' + new Date().toLocaleDateString();
    const ss = SpreadsheetApp.create(exportTitle);
    const sheet = ss.getActiveSheet();
    sheet.setName(viewType || 'Export');
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#1e293b');
    headerRange.setFontColor('white');
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
    return { success: true, spreadsheetUrl: ss.getUrl() };
  } catch (error) {
    console.error('exportViewToSheets failed:', error);
    return { success: false, error: error.message };
  }
}
