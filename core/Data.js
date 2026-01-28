function getSheet(sheetName, columns) {
  if (!columns || !Array.isArray(columns)) {
    console.error(`Invalid columns parameter for sheet ${sheetName}:`, columns);
    throw new Error(`Invalid columns parameter for sheet ${sheetName}. Expected array, got: ${typeof columns}`);
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, columns.length)
      .setValues([columns])
      .setFontWeight('bold')
      .setBackground('#1e293b')
      .setFontColor('white');
    sheet.setFrozenRows(1);
  } else {
    const existingHeaders = sheet.getRange(1, 1, 1, columns.length).getValues()[0];
    const headersMatch = columns.every((col, index) => existingHeaders[index] === col);
    if (!headersMatch) {
      sheet.getRange(1, 1, 1, columns.length)
        .setValues([columns])
        .setFontWeight('bold')
        .setBackground('#1e293b')
        .setFontColor('white');
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

function getTasksSheet() {
  if (!CONFIG || !CONFIG.TASK_COLUMNS) {
    throw new Error('CONFIG.TASK_COLUMNS is undefined. Check Config.gs for syntax errors.');
  }
  return getSheet(CONFIG.SHEETS.TASKS, CONFIG.TASK_COLUMNS);
}

function getUsersSheet() {
  if (!CONFIG || !CONFIG.USER_COLUMNS) {
    throw new Error('CONFIG.USER_COLUMNS is undefined. Check Config.gs for syntax errors.');
  }
  return getSheet(CONFIG.SHEETS.USERS, CONFIG.USER_COLUMNS);
}

function getProjectsSheet() {
  if (!CONFIG || !CONFIG.PROJECT_COLUMNS) {
    throw new Error('CONFIG.PROJECT_COLUMNS is undefined. Check Config.gs for syntax errors.');
  }
  return getSheet(CONFIG.SHEETS.PROJECTS, CONFIG.PROJECT_COLUMNS);
}

function getCommentsSheet() {
  if (!CONFIG || !CONFIG.COMMENT_COLUMNS) {
    throw new Error('CONFIG.COMMENT_COLUMNS is undefined. Check Config.gs for syntax errors.');
  }
  return getSheet(CONFIG.SHEETS.COMMENTS, CONFIG.COMMENT_COLUMNS);
}

function getActivitySheet() {
  if (!CONFIG || !CONFIG.ACTIVITY_COLUMNS) {
    throw new Error('CONFIG.ACTIVITY_COLUMNS is undefined. Check Config.gs for syntax errors.');
  }
  return getSheet(CONFIG.SHEETS.ACTIVITY, CONFIG.ACTIVITY_COLUMNS);
}

function getMentionsSheet() {
  if (!CONFIG || !CONFIG.MENTION_COLUMNS) {
    console.error('CONFIG or CONFIG.MENTION_COLUMNS is undefined');
    throw new Error('CONFIG.MENTION_COLUMNS is undefined. Check Config.gs for syntax errors.');
  }
  return getSheet(CONFIG.SHEETS.MENTIONS, CONFIG.MENTION_COLUMNS);
}

function getNotificationsSheet() {
  if (!CONFIG || !CONFIG.NOTIFICATION_COLUMNS) {
    throw new Error('CONFIG.NOTIFICATION_COLUMNS is undefined. Check Config.gs for syntax errors.');
  }
  return getSheet(CONFIG.SHEETS.NOTIFICATIONS, CONFIG.NOTIFICATION_COLUMNS);
}

function getAnalyticsCacheSheet() {
  if (!CONFIG || !CONFIG.ANALYTICS_CACHE_COLUMNS) {
    throw new Error('CONFIG.ANALYTICS_CACHE_COLUMNS is undefined. Check Config.gs for syntax errors.');
  }
  return getSheet(CONFIG.SHEETS.ANALYTICS_CACHE, CONFIG.ANALYTICS_CACHE_COLUMNS);
}

function getTaskDependenciesSheet() {
  if (!CONFIG || !CONFIG.TASK_DEPENDENCY_COLUMNS) {
    throw new Error('CONFIG.TASK_DEPENDENCY_COLUMNS is undefined. Check Config.gs for syntax errors.');
  }
  return getSheet(CONFIG.SHEETS.TASK_DEPENDENCIES, CONFIG.TASK_DEPENDENCY_COLUMNS);
}

function getSessionTokensSheet() {
  if (!CONFIG || !CONFIG.SESSION_TOKEN_COLUMNS) {
    throw new Error('CONFIG.SESSION_TOKEN_COLUMNS is undefined. Check Config.gs for syntax errors.');
  }
  return getSheet(CONFIG.SHEETS.SESSION_TOKENS, CONFIG.SESSION_TOKEN_COLUMNS);
}

function getMfaCodesSheet() {
  if (!CONFIG || !CONFIG.MFA_CODE_COLUMNS) {
    throw new Error('CONFIG.MFA_CODE_COLUMNS is undefined. Check Config.gs for syntax errors.');
  }
  return getSheet(CONFIG.SHEETS.MFA_CODES, CONFIG.MFA_CODE_COLUMNS);
}

function getRolesSheet() {
  if (!CONFIG || !CONFIG.ROLE_COLUMNS) {
    throw new Error('CONFIG.ROLE_COLUMNS is undefined. Check Config.gs for syntax errors.');
  }
  return getSheet(CONFIG.SHEETS.ROLES, CONFIG.ROLE_COLUMNS);
}

function getProjectMembersSheet() {
  if (!CONFIG || !CONFIG.PROJECT_MEMBER_COLUMNS) {
    throw new Error('CONFIG.PROJECT_MEMBER_COLUMNS is undefined. Check Config.gs for syntax errors.');
  }
  return getSheet(CONFIG.SHEETS.PROJECT_MEMBERS, CONFIG.PROJECT_MEMBER_COLUMNS);
}

function getOrganizationsSheet() {
  if (!CONFIG || !CONFIG.ORGANIZATION_COLUMNS) {
    throw new Error('CONFIG.ORGANIZATION_COLUMNS is undefined. Check Config.gs for syntax errors.');
  }
  return getSheet(CONFIG.SHEETS.ORGANIZATIONS, CONFIG.ORGANIZATION_COLUMNS);
}

function getOrganizationMembersSheet() {
  if (!CONFIG || !CONFIG.ORGANIZATION_MEMBER_COLUMNS) {
    throw new Error('CONFIG.ORGANIZATION_MEMBER_COLUMNS is undefined. Check Config.gs for syntax errors.');
  }
  return getSheet(CONFIG.SHEETS.ORGANIZATION_MEMBERS, CONFIG.ORGANIZATION_MEMBER_COLUMNS);
}

function getTeamsSheet() {
  if (!CONFIG || !CONFIG.TEAM_COLUMNS) {
    throw new Error('CONFIG.TEAM_COLUMNS is undefined. Check Config.gs for syntax errors.');
  }
  return getSheet(CONFIG.SHEETS.TEAMS, CONFIG.TEAM_COLUMNS);
}

function getTeamMembersSheet() {
  if (!CONFIG || !CONFIG.TEAM_MEMBER_COLUMNS) {
    throw new Error('CONFIG.TEAM_MEMBER_COLUMNS is undefined. Check Config.gs for syntax errors.');
  }
  return getSheet(CONFIG.SHEETS.TEAM_MEMBERS, CONFIG.TEAM_MEMBER_COLUMNS);
}

function getUserPreferencesSheet() {
  if (!CONFIG || !CONFIG.USER_PREFERENCE_COLUMNS) {
    throw new Error('CONFIG.USER_PREFERENCE_COLUMNS is undefined. Check Config.gs for syntax errors.');
  }
  return getSheet(CONFIG.SHEETS.USER_PREFERENCES, CONFIG.USER_PREFERENCE_COLUMNS);
}

function getAutomationRulesSheet() {
  if (!CONFIG || !CONFIG.AUTOMATION_RULE_COLUMNS) {
    throw new Error('CONFIG.AUTOMATION_RULE_COLUMNS is undefined. Check Config.gs for syntax errors.');
  }
  return getSheet(CONFIG.SHEETS.AUTOMATION_RULES, CONFIG.AUTOMATION_RULE_COLUMNS);
}

function getTeamCapacitySheet() {
  if (!CONFIG || !CONFIG.TEAM_CAPACITY_COLUMNS) {
    throw new Error('CONFIG.TEAM_CAPACITY_COLUMNS is undefined. Check Config.gs for syntax errors.');
  }
  return getSheet(CONFIG.SHEETS.TEAM_CAPACITY, CONFIG.TEAM_CAPACITY_COLUMNS);
}

function getWebhookSubscriptionsSheet() {
  if (!CONFIG || !CONFIG.WEBHOOK_SUBSCRIPTION_COLUMNS) {
    throw new Error('CONFIG.WEBHOOK_SUBSCRIPTION_COLUMNS is undefined. Check Config.gs for syntax errors.');
  }
  return getSheet(CONFIG.SHEETS.WEBHOOK_SUBSCRIPTIONS, CONFIG.WEBHOOK_SUBSCRIPTION_COLUMNS);
}

function getSlaConfigSheet() {
  if (!CONFIG || !CONFIG.SLA_CONFIG_COLUMNS) {
    throw new Error('CONFIG.SLA_CONFIG_COLUMNS is undefined. Check Config.gs for syntax errors.');
  }
  return getSheet(CONFIG.SHEETS.SLA_CONFIG, CONFIG.SLA_CONFIG_COLUMNS);
}

function getTriageQueueSheet() {
  if (!CONFIG || !CONFIG.TRIAGE_QUEUE_COLUMNS) {
    throw new Error('CONFIG.TRIAGE_QUEUE_COLUMNS is undefined. Check Config.gs for syntax errors.');
  }
  return getSheet(CONFIG.SHEETS.TRIAGE_QUEUE, CONFIG.TRIAGE_QUEUE_COLUMNS);
}

const RowIndexCache = {
  CACHE_TTL: 600,

  get(sheetName, id) {
    try {
      const cache = CacheService.getScriptCache();
      const key = `row_${sheetName}_${id}`;
      const cached = cache.get(key);
      if (cached) {
        return parseInt(cached, 10);
      }
      return null;
    } catch (e) {
      console.warn('RowIndexCache.get failed:', e.message);
      return null;
    }
  },

  set(sheetName, id, rowIndex) {
    try {
      const cache = CacheService.getScriptCache();
      const key = `row_${sheetName}_${id}`;
      cache.put(key, rowIndex.toString(), this.CACHE_TTL);
    } catch (e) {
      console.warn('RowIndexCache.set failed:', e.message);
    }
  },

  invalidate(sheetName, id) {
    try {
      const cache = CacheService.getScriptCache();
      const key = `row_${sheetName}_${id}`;
      cache.remove(key);
    } catch (e) {
      console.warn('RowIndexCache.invalidate failed:', e.message);
    }
  },

  invalidateSheet(sheetName) {
  }
};

function findRowWithCache(sheet, sheetName, id, idColumnIndex) {
  let rowIndex = RowIndexCache.get(sheetName, id);
  if (rowIndex) {
    try {
      const idCell = sheet.getRange(rowIndex, idColumnIndex + 1).getValue();
      if (idCell === id) {
        return rowIndex;
      }
      RowIndexCache.invalidate(sheetName, id);
      rowIndex = null;
    } catch (e) {
      RowIndexCache.invalidate(sheetName, id);
      rowIndex = null;
    }
  }
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][idColumnIndex] === id) {
      rowIndex = i + 1;
      RowIndexCache.set(sheetName, id, rowIndex);
      return rowIndex;
    }
  }
  return null;
}

function getAllOrganizations() {
  const sheet = getOrganizationsSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const columns = CONFIG.ORGANIZATION_COLUMNS;
  const orgs = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const org = rowToObject(row, columns);
    if (typeof org.settings === 'string' && org.settings) {
      try {
        org.settings = JSON.parse(org.settings);
      } catch (e) {
        org.settings = {};
      }
    } else {
      org.settings = {};
    }
    orgs.push(org);
  }
  return orgs;
}

function getOrganizationById(orgId) {
  const orgs = getAllOrganizations();
  return orgs.find(o => o.id === orgId) || null;
}

function createOrganization(orgData) {
  if (typeof PermissionGuard !== 'undefined') {
    PermissionGuard.requirePermission('admin:settings');
  }
  const sheet = getOrganizationsSheet();
  const currentUser = getCurrentUserEmail();
  const org = {
    id: orgData.id || generateId('org'),
    name: sanitize(orgData.name || 'New Organization'),
    slug: orgData.slug || orgData.name?.toLowerCase().replace(/\s+/g, '-').replace(/[^a-z0-9-]/g, '') || generateId('org'),
    description: sanitize(orgData.description || ''),
    settings: JSON.stringify(orgData.settings || {}),
    plan: orgData.plan || 'free',
    createdAt: now(),
    ownerId: orgData.ownerId || currentUser
  };
  sheet.appendRow(objectToRow(org, CONFIG.ORGANIZATION_COLUMNS));
  addOrganizationMember(org.id, org.ownerId, 'owner');
  logActivity(currentUser, 'created', 'organization', org.id, { name: org.name });
  return org;
}

function updateOrganization(orgId, updates) {
  if (typeof PermissionGuard !== 'undefined') {
    PermissionGuard.requirePermission('admin:settings');
  }
  const sheet = getOrganizationsSheet();
  const data = sheet.getDataRange().getValues();
  const columns = CONFIG.ORGANIZATION_COLUMNS;
  const idIndex = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][idIndex] === orgId) {
      const org = rowToObject(data[i], columns);
      if (updates.settings && typeof updates.settings !== 'string') {
        updates.settings = JSON.stringify(updates.settings);
      }
      Object.assign(org, updates);
      const newRow = objectToRow(org, columns);
      sheet.getRange(i + 1, 1, 1, columns.length).setValues([newRow]);
      return org;
    }
  }
  throw new Error('Organization not found: ' + orgId);
}

function addOrganizationMember(orgId, userEmail, role) {
  const sheet = getOrganizationMembersSheet();
  const columns = CONFIG.ORGANIZATION_MEMBER_COLUMNS;
  const currentUser = getCurrentUserEmail();
  const member = {
    id: generateId('om'),
    organizationId: orgId,
    userId: userEmail,
    orgRole: role || 'member',
    joinedAt: now(),
    invitedBy: currentUser
  };
  sheet.appendRow(objectToRow(member, columns));
  try {
    updateUser(userEmail, { organizationId: orgId });
  } catch (e) {
    console.error('Failed to update user org ID:', e);
  }
  return member;
}

function removeOrganizationMember(orgId, userEmail) {
  const sheet = getOrganizationMembersSheet();
  const data = sheet.getDataRange().getValues();
  const columns = CONFIG.ORGANIZATION_MEMBER_COLUMNS;
  for (let i = 1; i < data.length; i++) {
    const member = rowToObject(data[i], columns);
    if (member.organizationId === orgId && member.userId === userEmail) {
      sheet.deleteRow(i + 1);
      try {
        updateUser(userEmail, { organizationId: '' });
      } catch (e) {
        console.error('Failed to clear user org ID:', e);
      }
      return true;
    }
  }
  return false;
}

function getOrganizationMembers(orgId) {
  const sheet = getOrganizationMembersSheet();
  const data = sheet.getDataRange().getValues();
  const columns = CONFIG.ORGANIZATION_MEMBER_COLUMNS;
  const members = [];
  for (let i = 1; i < data.length; i++) {
    const member = rowToObject(data[i], columns);
    if (member.organizationId === orgId) {
      members.push(member);
    }
  }
  return members;
}

function getAllTeams(orgId) {
  const sheet = getTeamsSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const columns = CONFIG.TEAM_COLUMNS;
  let teams = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    teams.push(rowToObject(row, columns));
  }
  if (orgId) {
    teams = teams.filter(t => t.organizationId === orgId);
  }
  return teams;
}

function getTeamById(teamId) {
  const teams = getAllTeams();
  return teams.find(t => t.id === teamId) || null;
}

function createTeam(teamData) {
  const sheet = getTeamsSheet();
  const currentUser = getCurrentUserEmail();
  const team = {
    id: teamData.id || generateId('team'),
    organizationId: teamData.organizationId || '',
    name: sanitize(teamData.name || 'New Team'),
    description: sanitize(teamData.description || ''),
    leaderId: teamData.leaderId || currentUser,
    parentTeamId: teamData.parentTeamId || '',
    createdAt: now()
  };
  sheet.appendRow(objectToRow(team, CONFIG.TEAM_COLUMNS));
  addTeamMember(team.id, team.leaderId, 'leader');
  logActivity(currentUser, 'created', 'team', team.id, { name: team.name });
  return team;
}

function updateTeam(teamId, updates) {
  const sheet = getTeamsSheet();
  const data = sheet.getDataRange().getValues();
  const columns = CONFIG.TEAM_COLUMNS;
  const idIndex = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][idIndex] === teamId) {
      const team = rowToObject(data[i], columns);
      Object.assign(team, updates);
      const newRow = objectToRow(team, columns);
      sheet.getRange(i + 1, 1, 1, columns.length).setValues([newRow]);
      return team;
    }
  }
  throw new Error('Team not found: ' + teamId);
}

function deleteTeam(teamId) {
  const sheet = getTeamsSheet();
  const data = sheet.getDataRange().getValues();
  const idIndex = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][idIndex] === teamId) {
      sheet.deleteRow(i + 1);
      removeAllTeamMembers(teamId);
      logActivity(getCurrentUserEmail(), 'deleted', 'team', teamId, {});
      return true;
    }
  }
  return false;
}

function addTeamMember(teamId, userEmail, role) {
  const sheet = getTeamMembersSheet();
  const columns = CONFIG.TEAM_MEMBER_COLUMNS;
  const member = {
    id: generateId('tm'),
    teamId: teamId,
    userId: userEmail,
    teamRole: role || 'member',
    joinedAt: now()
  };
  sheet.appendRow(objectToRow(member, columns));
  try {
    updateUser(userEmail, { teamId: teamId });
  } catch (e) {
    console.error('Failed to update user team ID:', e);
  }
  return member;
}

function removeTeamMember(teamId, userEmail) {
  const sheet = getTeamMembersSheet();
  const data = sheet.getDataRange().getValues();
  const columns = CONFIG.TEAM_MEMBER_COLUMNS;
  for (let i = 1; i < data.length; i++) {
    const member = rowToObject(data[i], columns);
    if (member.teamId === teamId && member.userId === userEmail) {
      sheet.deleteRow(i + 1);
      try {
        updateUser(userEmail, { teamId: '' });
      } catch (e) {
        console.error('Failed to clear user team ID:', e);
      }
      return true;
    }
  }
  return false;
}

function removeAllTeamMembers(teamId) {
  const sheet = getTeamMembersSheet();
  const data = sheet.getDataRange().getValues();
  const columns = CONFIG.TEAM_MEMBER_COLUMNS;
  for (let i = data.length - 1; i >= 1; i--) {
    const member = rowToObject(data[i], columns);
    if (member.teamId === teamId) {
      sheet.deleteRow(i + 1);
    }
  }
}

function getTeamMembers(teamId) {
  const sheet = getTeamMembersSheet();
  const data = sheet.getDataRange().getValues();
  const columns = CONFIG.TEAM_MEMBER_COLUMNS;
  const members = [];
  for (let i = 1; i < data.length; i++) {
    const member = rowToObject(data[i], columns);
    if (member.teamId === teamId) {
      members.push(member);
    }
  }
  return members;
}

function getTeamsForUser(userEmail) {
  const sheet = getTeamMembersSheet();
  const data = sheet.getDataRange().getValues();
  const columns = CONFIG.TEAM_MEMBER_COLUMNS;
  const teamIds = [];
  for (let i = 1; i < data.length; i++) {
    const member = rowToObject(data[i], columns);
    if (member.userId === userEmail) {
      teamIds.push(member.teamId);
    }
  }
  const allTeams = getAllTeams();
  return allTeams.filter(t => teamIds.includes(t.id));
}

function rowToObject(row, columns) {
  const obj = {};
  columns.forEach((col, i) => {
    let value = row[i];
    if (value instanceof Date) {
      value = value.toISOString();
    }
    obj[col] = value !== undefined && value !== null ? value : '';
  });
  return obj;
}

function objectToRow(obj, columns) {
  return columns.map(col => {
    const value = obj[col];
    if (Array.isArray(value)) {
      return value.join(', ');
    }
    return value !== undefined ? value : '';
  });
}

function getAllTasks(filters = {}, options = {}) {
  const sheet = getTasksSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const columns = CONFIG.TASK_COLUMNS;
  let tasks = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] && !row[2]) continue;
    const task = rowToObject(row, columns);
    if (typeof task.labels === 'string' && task.labels) {
      task.labels = task.labels.split(',').map(l => l.trim()).filter(l => l);
    } else {
      task.labels = [];
    }
    task.storyPoints = parseInt(task.storyPoints) || 0;
    task.estimatedHrs = parseFloat(task.estimatedHrs) || 0;
    task.actualHrs = parseFloat(task.actualHrs) || 0;
    task.position = parseInt(task.position) || 0;
    task.isMilestone = task.isMilestone === true || task.isMilestone === 'true' || task.isMilestone === 'TRUE' || task.isMilestone === 1;
    tasks.push(task);
  }
  if (filters.assignee) {
    tasks = tasks.filter(t => t.assignee?.toLowerCase() === filters.assignee.toLowerCase());
  }
  if (filters.projectId) {
    tasks = tasks.filter(t => t.projectId === filters.projectId);
  }
  if (filters.status) {
    tasks = tasks.filter(t => t.status === filters.status);
  }
  if (filters.search) {
    const q = filters.search.toLowerCase();
    tasks = tasks.filter(t =>
      t.title?.toLowerCase().includes(q) ||
      t.description?.toLowerCase().includes(q) ||
      t.id?.toLowerCase().includes(q)
    );
  }
  if (!options.skipPermissionCheck && typeof PermissionGuard !== 'undefined') {
    tasks = PermissionGuard.filterTasksByPermission(tasks);
  }
  return tasks;
}

function getTaskById(taskId) {
  const tasks = getAllTasks();
  return tasks.find(t => t.id === taskId) || null;
}

function getTasksForUser(email) {
  return getAllTasks({ assignee: email });
}

function createTask(taskData) {
  if (typeof PermissionGuard !== 'undefined') {
    PermissionGuard.requirePermission('task:create', { projectId: taskData.projectId });
  }
  const sheet = getTasksSheet();
  const existingTasks = getAllTasks({}, { skipPermissionCheck: true });
  const currentUser = getCurrentUserEmail();
  const timestamp = now();
  const taskId = generateTaskId(taskData.projectId, existingTasks);
  const sameStatusTasks = existingTasks.filter(t => t.status === (taskData.status || 'To Do'));
  const maxPosition = sameStatusTasks.reduce((max, t) => Math.max(max, t.position || 0), 0);
  const isMilestone = taskData.isMilestone === true || taskData.isMilestone === 'true';
  if (isMilestone && !taskData.milestoneDate) {
    throw new Error('Milestone date is required for milestone tasks');
  }
  const task = {
    id: taskId,
    projectId: taskData.projectId || '',
    title: sanitize(taskData.title || 'New Task'),
    description: sanitize(taskData.description || ''),
    status: taskData.status || 'To Do',
    priority: taskData.priority || 'Medium',
    type: taskData.type || 'Task',
    assignee: taskData.assignee || currentUser,
    reporter: taskData.reporter || currentUser,
    dueDate: taskData.dueDate || '',
    startDate: taskData.startDate || '',
    sprint: taskData.sprint || '',
    storyPoints: parseInt(taskData.storyPoints) || 0,
    estimatedHrs: parseFloat(taskData.estimatedHrs) || 0,
    actualHrs: 0,
    labels: Array.isArray(taskData.labels) ? taskData.labels :
      (taskData.labels ? taskData.labels.split(',').map(l => l.trim()) : []),
    parentId: taskData.parentId || '',
    position: maxPosition + 1,
    createdAt: timestamp,
    updatedAt: timestamp,
    completedAt: '',
    dependencies: taskData.dependencies || '',
    timeEntries: taskData.timeEntries || '',
    customFields: taskData.customFields || '',
    templateId: taskData.templateId || '',
    recurringConfig: taskData.recurringConfig || '',
    isMilestone: isMilestone,
    milestoneDate: taskData.milestoneDate || '',
    milestoneType: taskData.milestoneType || ''
  };
  sheet.appendRow(objectToRow(task, CONFIG.TASK_COLUMNS));
  SpreadsheetApp.flush();
  const newRowIndex = sheet.getLastRow();
  RowIndexCache.set('Tasks', task.id, newRowIndex);
  logActivity(currentUser, 'created', 'task', task.id, { title: task.title });
  return task;
}

function updateTask(taskId, updates) {
  const sheet = getTasksSheet();
  const columns = CONFIG.TASK_COLUMNS;
  const idIndex = 0;
  const rowIndex = findRowWithCache(sheet, 'Tasks', taskId, idIndex);
  if (!rowIndex) {
    throw new Error('Task not found: ' + taskId);
  }
  const rowData = sheet.getRange(rowIndex, 1, 1, columns.length).getValues()[0];
  const task = rowToObject(rowData, columns);
  if (typeof PermissionGuard !== 'undefined') {
    if (!PermissionGuard.canUpdateTask(task)) {
      throw new Error('Permission denied: You cannot update this task');
    }
  }
  const isMilestone = updates.hasOwnProperty('isMilestone') ?
    (updates.isMilestone === true || updates.isMilestone === 'true') :
    (task.isMilestone === true || task.isMilestone === 'true');
  if (isMilestone) {
    const milestoneDate = updates.milestoneDate || task.milestoneDate;
    if (!milestoneDate) {
      throw new Error('Milestone date is required for milestone tasks');
    }
  }
  const changes = {};
  Object.keys(updates).forEach(key => {
    if (task[key] !== updates[key]) {
      changes[key] = { from: task[key], to: updates[key] };
    }
  });
  Object.assign(task, updates);
  task.updatedAt = now();
  if (updates.status === 'Done' && !task.completedAt) {
    task.completedAt = now();
  }
  const newRow = objectToRow(task, columns);
  sheet.getRange(rowIndex, 1, 1, columns.length).setValues([newRow]);
  SpreadsheetApp.flush();
  if (Object.keys(changes).length > 0) {
    logActivity(getCurrentUserEmail(), 'updated', 'task', taskId, changes);
  }
  return task;
}

function moveTask(taskId, newStatus, newPosition) {
  const updates = { status: newStatus };
  if (newPosition !== undefined) {
    updates.position = newPosition;
  }
  return updateTask(taskId, updates);
}

function deleteTask(taskId) {
  const sheet = getTasksSheet();
  const columns = CONFIG.TASK_COLUMNS;
  const idIndex = 0;
  const rowIndex = findRowWithCache(sheet, 'Tasks', taskId, idIndex);
  if (!rowIndex) {
    return false;
  }
  const rowData = sheet.getRange(rowIndex, 1, 1, columns.length).getValues()[0];
  if (typeof PermissionGuard !== 'undefined') {
    const task = rowToObject(rowData, columns);
    if (!PermissionGuard.canDeleteTask(task)) {
      throw new Error('Permission denied: You cannot delete this task');
    }
  }
  sheet.deleteRow(rowIndex);
  SpreadsheetApp.flush();
  RowIndexCache.invalidate('Tasks', taskId);
  logActivity(getCurrentUserEmail(), 'deleted', 'task', taskId, {});
  return true;
}

function getMilestones(projectId) {
  let tasks = getAllTasks();
  tasks = tasks.filter(t =>
    t.isMilestone === true ||
    t.isMilestone === 'true' ||
    t.isMilestone === 'TRUE' ||
    t.isMilestone === 1 ||
    t.isMilestone === '1'
  );
  if (projectId) {
    tasks = tasks.filter(t => t.projectId === projectId);
  }
  tasks.sort((a, b) => {
    const dateA = a.milestoneDate ? new Date(a.milestoneDate) : new Date(9999, 11, 31);
    const dateB = b.milestoneDate ? new Date(b.milestoneDate) : new Date(9999, 11, 31);
    return dateA - dateB;
  });
  return tasks;
}

function getUpcomingMilestones(days = 30) {
  const milestones = getMilestones();
  const now = new Date();
  const futureDate = new Date(now.getTime() + (days * 24 * 60 * 60 * 1000));
  return milestones.filter(m => {
    if (!m.milestoneDate) return false;
    const milestoneDate = new Date(m.milestoneDate);
    return milestoneDate >= now && milestoneDate <= futureDate;
  });
}

function getOverdueMilestones() {
  const milestones = getMilestones();
  const now = new Date();
  return milestones.filter(m => {
    if (!m.milestoneDate) return false;
    if (m.status === 'Done') return false;
    const milestoneDate = new Date(m.milestoneDate);
    return milestoneDate < now;
  });
}

function getCompletedMilestones() {
  const milestones = getMilestones();
  return milestones.filter(m => m.status === 'Done');
}

function getCurrentUserEmail() {
  try {
    const manualEmail = getManualUserEmail();
    if (manualEmail) {
      return manualEmail;
    }
    return Session.getActiveUser().getEmail() || 'anonymous@example.com';
  } catch (e) {
    return 'anonymous@example.com';
  }
}

function setCurrentUserEmail(email) {
  try {
    const properties = PropertiesService.getScriptProperties();
    properties.setProperty('CURRENT_USER_EMAIL', email);
    return true;
  } catch (e) {
    console.error('Failed to set current user email:', e);
    return false;
  }
}

function getManualUserEmail() {
  try {
    const properties = PropertiesService.getScriptProperties();
    return properties.getProperty('CURRENT_USER_EMAIL');
  } catch (e) {
    return null;
  }
}

function clearCurrentUserEmail() {
  try {
    const properties = PropertiesService.getScriptProperties();
    properties.deleteProperty('CURRENT_USER_EMAIL');
    return true;
  } catch (e) {
    console.error('Failed to clear current user email:', e);
    return false;
  }
}

function getCurrentUser() {
  const email = getCurrentUserEmail();
  if (email === 'anonymous@example.com') {
    throw new Error('No authenticated user found. Please login manually.');
  }
  let user = getUserByEmail(email);
  if (!user) {
    user = createUser({
      email: email,
      name: email.split('@')[0],
      role: 'admin'
    });
  }
  return user;
}

function getAllUsers() {
  const sheet = getUsersSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const columns = CONFIG.USER_COLUMNS;
  const users = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;

    const user = rowToObject(row, columns);
    const activeVal = user.active;
    user.active = activeVal === true ||
                  activeVal === 'true' ||
                  activeVal === 'TRUE' ||
                  activeVal === 'True' ||
                  activeVal === 'yes' ||
                  activeVal === 'YES' ||
                  activeVal === 'Yes' ||
                  activeVal === 1 ||
                  activeVal === '1' ||
                  activeVal === 'active' ||
                  activeVal === 'Active' ||
                  activeVal === 'ACTIVE';
    users.push(user);
  }

  return users;
}

function getUserByEmail(email) {
  const users = getAllUsers();
  return users.find(u => u.email?.toLowerCase() === email?.toLowerCase()) || null;
}

function createUser(userData) {
  const sheet = getUsersSheet();
  if (getUserByEmail(userData.email)) {
    return getUserByEmail(userData.email);
  }
  const user = {
    email: userData.email.toLowerCase().trim(),
    name: userData.name || userData.email.split('@')[0],
    role: userData.role || 'member',
    active: true,
    workbookId: userData.workbookId || '',
    createdAt: now()
  };
  sheet.appendRow(objectToRow(user, CONFIG.USER_COLUMNS));
  return user;
}

function updateUser(email, updates) {
  const sheet = getUsersSheet();
  const data = sheet.getDataRange().getValues();
  const columns = CONFIG.USER_COLUMNS;
  const emailIndex = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][emailIndex]?.toLowerCase() === email?.toLowerCase()) {
      const user = rowToObject(data[i], columns);
      Object.assign(user, updates);
      const newRow = objectToRow(user, columns);
      sheet.getRange(i + 1, 1, 1, columns.length).setValues([newRow]);
      return user;
    }
  }
  throw new Error('User not found: ' + email);
}

function getActiveUsers() {
  return getAllUsers().filter(u => u.active);
}

function getAllProjects() {
  const sheet = getProjectsSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const columns = CONFIG.PROJECT_COLUMNS;
  const projects = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    projects.push(rowToObject(row, columns));
  }
  return projects;
}

function getProjectById(projectId) {
  const projects = getAllProjects();
  return projects.find(p => p.id === projectId) || null;
}

function createProject(projectData) {
  const sheet = getProjectsSheet();
  const existingProjects = getAllProjects();
  const currentUser = getCurrentUserEmail();
  const projectId = projectData.id || generateProjectAcronym(projectData.name, existingProjects);
  const project = {
    id: projectId,
    name: sanitize(projectData.name || 'New Project'),
    description: sanitize(projectData.description || ''),
    status: projectData.status || 'active',
    ownerId: projectData.ownerId || currentUser,
    startDate: projectData.startDate || now().split('T')[0],
    endDate: projectData.endDate || '',
    createdAt: now()
  };
  sheet.appendRow(objectToRow(project, CONFIG.PROJECT_COLUMNS));
  SpreadsheetApp.flush();
  logActivity(currentUser, 'created', 'project', project.id, { name: project.name });
  return project;
}

function updateProject(projectId, updates) {
  const sheet = getProjectsSheet();
  const data = sheet.getDataRange().getValues();
  const columns = CONFIG.PROJECT_COLUMNS;
  const idIndex = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][idIndex] === projectId) {
      const project = rowToObject(data[i], columns);
      Object.assign(project, updates);
      const newRow = objectToRow(project, columns);
      sheet.getRange(i + 1, 1, 1, columns.length).setValues([newRow]);
      SpreadsheetApp.flush();
      return project;
    }
  }
  throw new Error('Project not found: ' + projectId);
}

function getCommentsForTask(taskId) {
  if (!taskId) {
    return [];
  }

  const sheet = getCommentsSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const columns = CONFIG.COMMENT_COLUMNS;
  const taskIdIndex = columns.indexOf('taskId');

  if (taskIdIndex === -1) {
    console.error('getCommentsForTask: taskId column not found in CONFIG.COMMENT_COLUMNS');
    return [];
  }

  const comments = [];
  const taskIdStr = String(taskId).trim().toLowerCase();

  for (let i = 1; i < data.length; i++) {
    const rowTaskId = String(data[i][taskIdIndex] || '').trim().toLowerCase();
    if (rowTaskId === taskIdStr) {
      comments.push(rowToObject(data[i], columns));
    }
  }

  return comments.sort((a, b) => new Date(a.createdAt) - new Date(b.createdAt));
}

function addComment(taskId, content) {
  const sheet = getCommentsSheet();
  const currentUser = getCurrentUserEmail();
  const timestamp = now();
  const comment = {
    id: generateId('cmt'),
    taskId: taskId,
    userId: currentUser,
    content: sanitize(content),
    createdAt: timestamp,
    updatedAt: timestamp
  };
  sheet.appendRow(objectToRow(comment, CONFIG.COMMENT_COLUMNS));
  logActivity(currentUser, 'commented', 'task', taskId, { preview: content.substring(0, 50) });
  return comment;
}

function deleteComment(commentId) {
  const sheet = getCommentsSheet();
  const data = sheet.getDataRange().getValues();
  const idIndex = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][idIndex] === commentId) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

function createMention(mentionData) {
  const sheet = getMentionsSheet();
  const timestamp = now();
  const mention = {
    id: generateId('men'),
    commentId: mentionData.commentId,
    mentionedUserId: mentionData.mentionedUserId,
    mentionedByUserId: mentionData.mentionedByUserId,
    taskId: mentionData.taskId,
    createdAt: timestamp,
    notificationSent: false,
    acknowledged: false
  };
  sheet.appendRow(objectToRow(mention, CONFIG.MENTION_COLUMNS));
  return mention;
}

function getMentionsForUser(userId, limit = 50) {
  const sheet = getMentionsSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const columns = CONFIG.MENTION_COLUMNS;
  const mentions = [];
  for (let i = 1; i < data.length; i++) {
    const mention = rowToObject(data[i], columns);
    if (mention.mentionedUserId === userId) {
      mentions.push(mention);
    }
  }
  mentions.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
  return mentions.slice(0, limit);
}

function acknowledgeMention(mentionId) {
  const sheet = getMentionsSheet();
  const data = sheet.getDataRange().getValues();
  const columns = CONFIG.MENTION_COLUMNS;
  const idIndex = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][idIndex] === mentionId) {
      const mention = rowToObject(data[i], columns);
      mention.acknowledged = true;
      const newRow = objectToRow(mention, columns);
      sheet.getRange(i + 1, 1, 1, columns.length).setValues([newRow]);
      return mention;
    }
  }
  throw new Error('Mention not found: ' + mentionId);
}

function createNotification(notificationData) {
  const sheet = getNotificationsSheet();
  const timestamp = now();
  const notification = {
    id: generateId('not'),
    userId: notificationData.userId,
    type: notificationData.type,
    title: sanitize(notificationData.title || ''),
    message: sanitize(notificationData.message || ''),
    entityType: notificationData.entityType || '',
    entityId: notificationData.entityId || '',
    read: false,
    createdAt: timestamp,
    scheduledFor: notificationData.scheduledFor || timestamp,
    channels: Array.isArray(notificationData.channels) ?
      notificationData.channels.join(',') :
      (notificationData.channels || 'in_app')
  };
  sheet.appendRow(objectToRow(notification, CONFIG.NOTIFICATION_COLUMNS));
  return notification;
}

function getNotificationsForUser(userId, limit = 50, unreadOnly = false) {
  try {
    const sheet = getNotificationsSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    const columns = CONFIG.NOTIFICATION_COLUMNS;
    const notifications = [];
    for (let i = 1; i < data.length; i++) {
      const notification = rowToObject(data[i], columns);
      if (notification.userId === userId) {
        notification.read = notification.read === true || notification.read === 'true';
        if (unreadOnly && notification.read) continue;
        if (typeof notification.channels === 'string' && notification.channels) {
          notification.channels = notification.channels.split(',').map(c => c.trim());
        } else {
          notification.channels = ['in_app'];
        }
        notifications.push(notification);
      }
    }
    notifications.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
    return notifications.slice(0, limit);
  } catch (error) {
    console.error('Error getting notifications for user:', error);
    return [];
  }
}

function markNotificationAsRead(notificationId) {
  const sheet = getNotificationsSheet();
  const data = sheet.getDataRange().getValues();
  const columns = CONFIG.NOTIFICATION_COLUMNS;
  const idIndex = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][idIndex] === notificationId) {
      const notification = rowToObject(data[i], columns);
      notification.read = true;
      const newRow = objectToRow(notification, columns);
      sheet.getRange(i + 1, 1, 1, columns.length).setValues([newRow]);
      return notification;
    }
  }
  throw new Error('Notification not found: ' + notificationId);
}

function getPendingNotifications(limit = 100) {
  const sheet = getNotificationsSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const columns = CONFIG.NOTIFICATION_COLUMNS;
  const notifications = [];
  const now = new Date();
  for (let i = 1; i < data.length; i++) {
    const notification = rowToObject(data[i], columns);
    const scheduledFor = new Date(notification.scheduledFor);
    if (scheduledFor <= now && !notification.processed) {
      notifications.push(notification);
    }
  }
  return notifications.slice(0, limit);
}

function getCachedAnalytics(cacheKey) {
  const sheet = getAnalyticsCacheSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return null;
  const columns = CONFIG.ANALYTICS_CACHE_COLUMNS;
  const now = new Date();
  for (let i = 1; i < data.length; i++) {
    const cache = rowToObject(data[i], columns);
    if (cache.cacheKey === cacheKey) {
      const expiresAt = new Date(cache.expiresAt);
      if (expiresAt > now) {
        try {
          return JSON.parse(cache.data);
        } catch (e) {
          console.error('Failed to parse cached data:', e);
          return null;
        }
      } else {
        sheet.deleteRow(i + 1);
        return null;
      }
    }
  }
  return null;
}

function setCachedAnalytics(cacheKey, data, ttlSeconds = 3600) {
  const sheet = getAnalyticsCacheSheet();
  const timestamp = now();
  const expiresAt = new Date(Date.now() + (ttlSeconds * 1000)).toISOString();
  const existingData = sheet.getDataRange().getValues();
  for (let i = 1; i < existingData.length; i++) {
    if (existingData[i][1] === cacheKey) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  const cache = {
    id: generateId('cache'),
    cacheKey: cacheKey,
    data: JSON.stringify(data),
    expiresAt: expiresAt,
    createdAt: timestamp
  };
  sheet.appendRow(objectToRow(cache, CONFIG.ANALYTICS_CACHE_COLUMNS));
  return cache;
}

function clearExpiredCache() {
  const sheet = getAnalyticsCacheSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return 0;
  const now = new Date();
  let cleared = 0;
  for (let i = data.length - 1; i >= 1; i--) {
    const expiresAt = new Date(data[i][3]);
    if (expiresAt <= now) {
      sheet.deleteRow(i + 1);
      cleared++;
    }
  }
  return cleared;
}

function createTaskDependency(dependencyData) {
  const sheet = getTaskDependenciesSheet();
  const timestamp = now();
  const dependency = {
    id: generateId('dep'),
    predecessorId: dependencyData.predecessorId,
    successorId: dependencyData.successorId,
    dependencyType: dependencyData.dependencyType || 'finish_to_start',
    lag: parseFloat(dependencyData.lag) || 0,
    createdAt: timestamp
  };
  sheet.appendRow(objectToRow(dependency, CONFIG.TASK_DEPENDENCY_COLUMNS));
  return dependency;
}

function getTaskDependencies(taskId) {
  const sheet = getTaskDependenciesSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { predecessors: [], successors: [] };
  const columns = CONFIG.TASK_DEPENDENCY_COLUMNS;
  const predecessors = [];
  const successors = [];
  for (let i = 1; i < data.length; i++) {
    const dependency = rowToObject(data[i], columns);
    if (dependency.successorId === taskId) {
      predecessors.push(dependency);
    }
    if (dependency.predecessorId === taskId) {
      successors.push(dependency);
    }
  }
  return { predecessors, successors };
}

function deleteTaskDependency(dependencyId) {
  const sheet = getTaskDependenciesSheet();
  const data = sheet.getDataRange().getValues();
  const idIndex = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][idIndex] === dependencyId) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

function logActivity(userId, action, entityType, entityId, details) {
  try {
    const sheet = getActivitySheet();
    const activity = {
      id: generateId('act'),
      userId: userId,
      action: action,
      entityType: entityType,
      entityId: entityId,
      details: JSON.stringify(details),
      timestamp: now()
    };
    sheet.appendRow(objectToRow(activity, CONFIG.ACTIVITY_COLUMNS));
  } catch (e) {
    console.error('Activity log error:', e);
  }
}

function getRecentActivity(limit = 50, entityId = null) {
  const sheet = getActivitySheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const columns = CONFIG.ACTIVITY_COLUMNS;
  let activities = [];
  for (let i = 1; i < data.length; i++) {
    const activity = rowToObject(data[i], columns);
    try {
      activity.details = JSON.parse(activity.details || '{}');
    } catch (e) {
      activity.details = {};
    }
    if (entityId && activity.entityId !== entityId) continue;
    activities.push(activity);
  }
  activities.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
  return activities.slice(0, limit);
}
