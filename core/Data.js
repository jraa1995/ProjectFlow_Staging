function getSheet(sheetName, columns) {
  if (!columns || !Array.isArray(columns)) {
    console.error(`Invalid columns parameter for sheet ${sheetName}:`, columns);
    throw new Error(`Invalid columns parameter for sheet ${sheetName}. Expected array, got: ${typeof columns}`);
  }
  const ss = getColonySpreadsheet_();
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

  _getGenKey(sheetName) {
    return 'rowgen_' + sheetName;
  },

  _getGeneration(sheetName) {
    try {
      var cache = CacheService.getScriptCache();
      return cache.get(this._getGenKey(sheetName)) || '0';
    } catch (e) {
      return '0';
    }
  },

  get(sheetName, id) {
    try {
      var cache = CacheService.getScriptCache();
      var key = 'row_' + sheetName + '_' + id;
      var cached = cache.get(key);
      if (cached) {
        var parsed = JSON.parse(cached);
        if (parsed.gen === this._getGeneration(sheetName)) {
          return parsed.row;
        }
        cache.remove(key);
      }
      return null;
    } catch (e) {
      return null;
    }
  },

  set(sheetName, id, rowIndex) {
    try {
      var cache = CacheService.getScriptCache();
      var key = 'row_' + sheetName + '_' + id;
      var gen = this._getGeneration(sheetName);
      cache.put(key, JSON.stringify({ row: rowIndex, gen: gen }), this.CACHE_TTL);
    } catch (e) {
    }
  },

  invalidate(sheetName, id) {
    try {
      var cache = CacheService.getScriptCache();
      var key = 'row_' + sheetName + '_' + id;
      cache.remove(key);
    } catch (e) {
    }
  },

  invalidateSheet(sheetName) {
    try {
      var cache = CacheService.getScriptCache();
      var newGen = Date.now().toString(36);
      cache.put(this._getGenKey(sheetName), newGen, this.CACHE_TTL);
    } catch (e) {
    }
  }
};

const UserCache = {
  CACHE_KEY: 'USER_MAP_CACHE',
  CACHE_TTL: 300,
  _memoryCache: null,
  _memoryCacheTime: 0,

  getMap() {
    const now = Date.now();
    if (this._memoryCache && (now - this._memoryCacheTime) < 60000) {
      return this._memoryCache;
    }
    try {
      const cache = CacheService.getScriptCache();
      const cached = cache.get(this.CACHE_KEY);
      if (cached) {
        this._memoryCache = JSON.parse(cached);
        this._memoryCacheTime = now;
        return this._memoryCache;
      }
    } catch (e) {
    }
    return null;
  },

  buildAndCache() {
    const sheet = getUsersSheet();
    const data = sheet.getDataRange().getValues();
    const columns = CONFIG.USER_COLUMNS;
    const userMap = {};
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      const user = rowToObject(row, columns);
      if (user.email) {
        userMap[user.email.toLowerCase()] = user;
      }
    }
    try {
      const cache = CacheService.getScriptCache();
      cache.put(this.CACHE_KEY, JSON.stringify(userMap), this.CACHE_TTL);
    } catch (e) {
    }
    this._memoryCache = userMap;
    this._memoryCacheTime = Date.now();
    return userMap;
  },

  get(email) {
    if (!email) return null;
    const emailLower = email.toLowerCase();
    let userMap = this.getMap();
    if (!userMap) {
      userMap = this.buildAndCache();
    }
    return userMap[emailLower] || null;
  },

  invalidate() {
    try {
      const cache = CacheService.getScriptCache();
      cache.remove(this.CACHE_KEY);
      this._memoryCache = null;
      this._memoryCacheTime = 0;
    } catch (e) {
    }
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
  PermissionGuard.requirePermission('admin:settings');
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
  PermissionGuard.requirePermission('admin:settings');
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
  const jsonDataIndex = columns.indexOf('jsonData');
  const row = columns.map((col, i) => {
    if (i === jsonDataIndex) return '';
    const value = obj[col];
    if (Array.isArray(value)) return value.join(', ');
    return value !== undefined ? value : '';
  });
  if (jsonDataIndex !== -1) {
    try {
      var cleanObj = {};
      columns.forEach(col => { if (col !== 'jsonData') cleanObj[col] = obj[col] !== undefined ? obj[col] : ''; });
      row[jsonDataIndex] = JSON.stringify(cleanObj);
    } catch (e) {
      console.error('objectToRow jsonData serialization failed:', e);
    }
  }
  return row;
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
  if (filters.sprint) {
    tasks = tasks.filter(t => t.sprint === filters.sprint);
  }
  if (filters.search) {
    const q = filters.search.toLowerCase();
    tasks = tasks.filter(t =>
      t.title?.toLowerCase().includes(q) ||
      t.description?.toLowerCase().includes(q) ||
      t.id?.toLowerCase().includes(q)
    );
  }
  if (!options.skipPermissionCheck) {
    tasks = PermissionGuard.filterTasksByPermission(tasks);
  }
  return tasks;
}

function getTaskById(taskId) {
  if (!taskId) return null;
  const sheet = getTasksSheet();
  const columns = CONFIG.TASK_COLUMNS;
  const rowIndex = findRowWithCache(sheet, 'Tasks', taskId, 0);
  if (!rowIndex) return null;
  const rowData = sheet.getRange(rowIndex, 1, 1, columns.length).getValues()[0];
  const task = rowToObject(rowData, columns);
  task.labels = typeof task.labels === 'string' && task.labels
    ? task.labels.split(',').map(l => l.trim()).filter(l => l)
    : [];
  return task;
}

function getTasksForUser(email) {
  return getAllTasks({ assignee: email });
}

function calculateStoryPoints(startDate, dueDate) {
  if (!startDate || !dueDate) return { storyPoints: 0, estimatedHrs: 0 };
  var start = new Date(startDate);
  var end = new Date(dueDate);
  if (isNaN(start.getTime()) || isNaN(end.getTime())) return { storyPoints: 0, estimatedHrs: 0 };
  if (end < start) return { storyPoints: 0, estimatedHrs: 0 };
  var workingDays = 0;
  var current = new Date(start);
  while (current <= end) {
    var day = current.getDay();
    if (day !== 0 && day !== 6) workingDays++;
    current.setDate(current.getDate() + 1);
  }
  return { storyPoints: workingDays, estimatedHrs: workingDays * 8 };
}

function createTask(taskData) {
  PermissionGuard.requirePermission('task:create', { projectId: taskData.projectId });
  const sheet = getTasksSheet();
  const currentUser = getCurrentUserEmail();
  const timestamp = now();
  const taskId = generateTaskId(taskData.projectId);
  const maxPosition = Date.now();
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
    assignee: taskData.assignee !== undefined ? (taskData.assignee || '') : currentUser,
    watchers: Array.isArray(taskData.watchers) ? taskData.watchers.join(',') : (taskData.watchers || ''),
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
  if (task.startDate && task.dueDate) {
    var calc = calculateStoryPoints(task.startDate, task.dueDate);
    task.storyPoints = calc.storyPoints;
    task.estimatedHrs = calc.estimatedHrs;
  }
  sheet.appendRow(objectToRow(task, CONFIG.TASK_COLUMNS));
  const newRowIndex = sheet.getLastRow();
  RowIndexCache.set('Tasks', task.id, newRowIndex);
  logActivity(currentUser, 'created', 'task', task.id, { title: task.title });
  if (task.assignee && task.assignee !== currentUser && typeof NotificationEngine !== 'undefined') {
    try {
      NotificationEngine.createNotification({
        userId: task.assignee,
        type: 'task_assigned',
        title: 'Task Assigned',
        message: 'You have been assigned to ' + task.id + ': ' + task.title,
        entityType: 'task',
        entityId: task.id,
        channels: ['in_app', 'email']
      });
    } catch (e) {
      console.error('Failed to send assignment notification:', e);
    }
  }
  return task;
}

function updateTask(taskId, updates) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const sheet = getTasksSheet();
    const columns = CONFIG.TASK_COLUMNS;
    const idIndex = 0;
    const rowIndex = findRowWithCache(sheet, 'Tasks', taskId, idIndex);
    if (!rowIndex) {
      throw new Error('Task not found: ' + taskId);
    }
    const rowData = sheet.getRange(rowIndex, 1, 1, columns.length).getValues()[0];
    const task = rowToObject(rowData, columns);
    if (!PermissionGuard.canUpdateTask(task)) {
      throw new Error('Permission denied: You cannot update this task');
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
    if (updates.startDate !== undefined || updates.dueDate !== undefined) {
      if (task.startDate && task.dueDate) {
        var calc = calculateStoryPoints(task.startDate, task.dueDate);
        task.storyPoints = calc.storyPoints;
        task.estimatedHrs = calc.estimatedHrs;
      }
    }
    task.updatedAt = now();
    if (updates.status === 'Done' && !task.completedAt) {
      task.completedAt = now();
    }
    const newRow = objectToRow(task, columns);
    sheet.getRange(rowIndex, 1, 1, columns.length).setValues([newRow]);
    const currentUser = getCurrentUserEmail();
    if (Object.keys(changes).length > 0) {
      logActivity(currentUser, 'updated', 'task', taskId, changes);
    }
    if (changes.assignee && updates.assignee && updates.assignee !== currentUser && typeof NotificationEngine !== 'undefined') {
      try {
        NotificationEngine.createNotification({
          userId: updates.assignee,
          type: 'task_assigned',
          title: 'Task Assigned',
          message: 'You have been assigned to ' + taskId + ': ' + task.title,
          entityType: 'task',
          entityId: taskId,
          channels: ['in_app', 'email']
        });
      } catch (e) {
        console.error('Failed to send assignment notification:', e);
      }
    }
    return task;
  } finally {
    lock.releaseLock();
  }
}

function moveTask(taskId, newStatus, newPosition) {
  const updates = { status: newStatus };
  if (newPosition !== undefined) {
    updates.position = newPosition;
  }
  return updateTask(taskId, updates);
}

function deleteTask(taskId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const sheet = getTasksSheet();
    const columns = CONFIG.TASK_COLUMNS;
    const idIndex = 0;
    const rowIndex = findRowWithCache(sheet, 'Tasks', taskId, idIndex);
    if (!rowIndex) {
      return false;
    }
    const rowData = sheet.getRange(rowIndex, 1, 1, columns.length).getValues()[0];
    const task = rowToObject(rowData, columns);
    if (!PermissionGuard.canDeleteTask(task)) {
      throw new Error('Permission denied: You cannot delete this task');
    }
    sheet.deleteRow(rowIndex);
    SpreadsheetApp.flush();
    RowIndexCache.invalidate('Tasks', taskId);
    RowIndexCache.invalidateSheet('Tasks');
    logActivity(getCurrentUserEmail(), 'deleted', 'task', taskId, {});
    return true;
  } finally {
    lock.releaseLock();
  }
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
    // Use per-user keyed cache instead of shared ScriptProperties
    const userKey = Session.getTemporaryActiveUserKey();
    if (userKey) {
      const cache = CacheService.getScriptCache();
      cache.put('ACTIVE_SESSION_EMAIL_' + userKey, email.toLowerCase().trim(), 86400);
    }
    return true;
  } catch (e) {
    console.error('Failed to set current user email:', e);
    return false;
  }
}

function getManualUserEmail() {
  try {
    // Use per-user keyed cache instead of shared ScriptProperties
    const userKey = Session.getTemporaryActiveUserKey();
    if (userKey) {
      const cache = CacheService.getScriptCache();
      return cache.get('ACTIVE_SESSION_EMAIL_' + userKey);
    }
    return null;
  } catch (e) {
    return null;
  }
}

function clearCurrentUserEmail() {
  try {
    const userKey = Session.getTemporaryActiveUserKey();
    if (userKey) {
      const cache = CacheService.getScriptCache();
      cache.remove('ACTIVE_SESSION_EMAIL_' + userKey);
    }
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
    // Auto-create as 'member' role (not admin) for new domain users
    user = createUser({
      email: email,
      name: email.split('@')[0],
      role: 'member'
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
  return UserCache.get(email);
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
  UserCache.invalidate();
  return user;
}

function updateUser(email, updates) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
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
        UserCache.invalidate();
        return user;
      }
    }
    throw new Error('User not found: ' + email);
  } finally {
    lock.releaseLock();
  }
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
  PermissionGuard.requirePermission('project:create');
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
    createdAt: now(),
    updatedAt: now(),
    version: projectData.version || '0.1.0',
    repoUrl: projectData.repoUrl || '',
    releaseNotes: projectData.releaseNotes || '[]',
    changelog: projectData.changelog || '[]',
    tags: projectData.tags || '',
    settings: projectData.settings || '{}',
    lastUpdatedBy: currentUser
  };
  sheet.appendRow(objectToRow(project, CONFIG.PROJECT_COLUMNS));
  SpreadsheetApp.flush();
  logActivity(currentUser, 'created', 'project', project.id, { name: project.name });
  return project;
}

function updateProject(projectId, updates) {
  PermissionGuard.requirePermission('project:update', { projectId: projectId });
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const sheet = getProjectsSheet();
    const data = sheet.getDataRange().getValues();
    const columns = CONFIG.PROJECT_COLUMNS;
    const idIndex = 0;
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIndex] === projectId) {
        const project = rowToObject(data[i], columns);
        Object.assign(project, updates);
        project.updatedAt = now();
        project.lastUpdatedBy = getCurrentUserEmail();
        const newRow = objectToRow(project, columns);
        sheet.getRange(i + 1, 1, 1, columns.length).setValues([newRow]);
        SpreadsheetApp.flush();
        return project;
      }
    }
    throw new Error('Project not found: ' + projectId);
  } finally {
    lock.releaseLock();
  }
}

function deleteProject(projectId) {
  PermissionGuard.requirePermission('project:delete');
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const sheet = getProjectsSheet();
    const data = sheet.getDataRange().getValues();
    const idIndex = 0;
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIndex] === projectId) {
        sheet.deleteRow(i + 1);
        SpreadsheetApp.flush();
        logActivity(getCurrentUserEmail(), 'deleted', 'project', projectId, {});
        return { success: true };
      }
    }
    throw new Error('Project not found: ' + projectId);
  } finally {
    lock.releaseLock();
  }
}

function addReleaseNote(projectId, releaseData) {
  PermissionGuard.requirePermission('project:update', { projectId: projectId });
  const project = getProjectById(projectId);
  if (!project) throw new Error('Project not found: ' + projectId);
  var notes = [];
  try { notes = typeof project.releaseNotes === 'string' ? JSON.parse(project.releaseNotes) : (project.releaseNotes || []); } catch (e) { notes = []; }
  var entry = {
    version: releaseData.version || project.version || '0.0.0',
    date: releaseData.date || now(),
    type: releaseData.type || 'patch',
    summary: sanitize(releaseData.summary || '')
  };
  if (releaseData.sections && typeof releaseData.sections === 'object') {
    entry.sections = {};
    ['added', 'changed', 'fixed', 'removed', 'security'].forEach(function(key) {
      if (Array.isArray(releaseData.sections[key]) && releaseData.sections[key].length > 0) {
        entry.sections[key] = releaseData.sections[key].map(function(item) { return sanitize(String(item)); });
      }
    });
  }
  if (releaseData.notes) entry.notes = sanitize(releaseData.notes);
  notes.unshift(entry);
  return updateProject(projectId, {
    releaseNotes: JSON.stringify(notes),
    version: releaseData.version || project.version
  });
}

function addChangelogEntry(projectId, entry) {
  PermissionGuard.requirePermission('project:update', { projectId: projectId });
  const project = getProjectById(projectId);
  if (!project) throw new Error('Project not found: ' + projectId);
  var log = [];
  try { log = typeof project.changelog === 'string' ? JSON.parse(project.changelog) : (project.changelog || []); } catch (e) { log = []; }
  log.unshift({
    date: entry.date || now(),
    author: entry.author || getCurrentUserEmail(),
    type: entry.type || 'update',
    message: sanitize(entry.message || ''),
    relatedVersion: entry.relatedVersion || ''
  });
  return updateProject(projectId, { changelog: JSON.stringify(log) });
}

function getOpenSurgeTasks() {
  const tasks = getAllTasks({}, { skipPermissionCheck: true });
  return tasks
    .filter(t => {
      const labels = Array.isArray(t.labels) ? t.labels : (t.labels ? String(t.labels).split(',').map(l => l.trim()) : []);
      return labels.includes('surge') && !t.assignee;
    })
    .sort((a, b) => (b.createdAt || '').localeCompare(a.createdAt || ''));
}

function pickUpTask(taskId) {
  PermissionGuard.requirePermission('task:update:own', { ownerId: getCurrentUserEmail() });
  const task = getTaskById(taskId);
  if (!task) throw new Error('Task not found: ' + taskId);
  const labels = Array.isArray(task.labels) ? task.labels : (task.labels ? String(task.labels).split(',').map(l => l.trim()) : []);
  if (!labels.includes('surge')) throw new Error('Task is not a surge task');
  const newLabels = labels.filter(l => l !== 'surge');
  return updateTask(taskId, {
    assignee: getCurrentUserEmail(),
    labels: newLabels.join(', ')
  });
}

function getJsonCacheSheet() {
  const ss = getColonySpreadsheet_();
  let sheet = ss.getSheetByName(CONFIG.SHEETS.JSON_CACHE);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.JSON_CACHE);
    sheet.appendRow(CONFIG.JSON_CACHE_COLUMNS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function writeJsonCache(key, data) {
  try {
    const sheet = getJsonCacheSheet();
    const allData = sheet.getDataRange().getValues();
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][0] === key) {
        sheet.getRange(i + 1, 2, 1, 2).setValues([[JSON.stringify(data), now()]]);
        return;
      }
    }
    sheet.appendRow([key, JSON.stringify(data), now()]);
  } catch (e) {
    console.error('writeJsonCache failed:', e);
  }
}

function readJsonCache(key, maxAgeMinutes) {
  try {
    const sheet = getJsonCacheSheet();
    const allData = sheet.getDataRange().getValues();
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][0] === key) {
        if (maxAgeMinutes) {
          const updatedAt = new Date(allData[i][2]);
          const age = (Date.now() - updatedAt.getTime()) / 60000;
          if (age > maxAgeMinutes) return null;
        }
        return JSON.parse(allData[i][1]);
      }
    }
    return null;
  } catch (e) {
    console.error('readJsonCache failed:', e);
    return null;
  }
}

function backfillJsonData(sheetName, columns) {
  try {
    const ss = getColonySpreadsheet_();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, error: 'Sheet not found: ' + sheetName };
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, updated: 0 };
    const jsonDataIndex = columns.indexOf('jsonData');
    if (jsonDataIndex === -1) return { success: false, error: 'No jsonData column in ' + sheetName };
    var updated = 0;
    var pendingValues = [];
    var pendingRows = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][jsonDataIndex]) continue;
      var obj = rowToObject(data[i], columns);
      var cleanObj = {};
      columns.forEach(function(col) { if (col !== 'jsonData') cleanObj[col] = obj[col] || ''; });
      pendingRows.push(i + 1);
      pendingValues.push(JSON.stringify(cleanObj));
      updated++;
    }
    if (pendingRows.length > 0) {
      var colNum = jsonDataIndex + 1;
      var batchStart = 0;
      while (batchStart < pendingRows.length) {
        var batchEnd = batchStart;
        while (batchEnd + 1 < pendingRows.length &&
               pendingRows[batchEnd + 1] === pendingRows[batchEnd] + 1) {
          batchEnd++;
        }
        var vals = [];
        for (var b = batchStart; b <= batchEnd; b++) {
          vals.push([pendingValues[b]]);
        }
        sheet.getRange(pendingRows[batchStart], colNum, vals.length, 1).setValues(vals);
        batchStart = batchEnd + 1;
      }
    }
    SpreadsheetApp.flush();
    return { success: true, updated: updated };
  } catch (e) {
    console.error('backfillJsonData failed:', e);
    return { success: false, error: e.message };
  }
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

function getAllDependenciesMap() {
  if (RequestCache._dependencies !== undefined && RequestCache._dependencies !== null) {
    return RequestCache._dependencies;
  }
  const sheet = getTaskDependenciesSheet();
  const data = sheet.getDataRange().getValues();
  const bySuccessor = {};
  const byPredecessor = {};
  if (data.length > 1) {
    const columns = CONFIG.TASK_DEPENDENCY_COLUMNS;
    for (let i = 1; i < data.length; i++) {
      const dep = rowToObject(data[i], columns);
      if (!bySuccessor[dep.successorId]) bySuccessor[dep.successorId] = [];
      bySuccessor[dep.successorId].push(dep);
      if (!byPredecessor[dep.predecessorId]) byPredecessor[dep.predecessorId] = [];
      byPredecessor[dep.predecessorId].push(dep);
    }
  }
  const result = { bySuccessor, byPredecessor };
  RequestCache._dependencies = result;
  return result;
}

function getTaskDependencies(taskId) {
  const map = getAllDependenciesMap();
  return {
    predecessors: map.bySuccessor[taskId] || [],
    successors: map.byPredecessor[taskId] || []
  };
}

function batchUpdatePositions(updates) {
  if (!updates || updates.length === 0) return;
  const sheet = getTasksSheet();
  const columns = CONFIG.TASK_COLUMNS;
  const data = sheet.getDataRange().getValues();
  const idIndex = 0;
  const positionIndex = columns.indexOf('position');
  if (positionIndex === -1) return;
  const updateMap = {};
  updates.forEach(u => { updateMap[u.taskId] = u.position; });
  const updatedAt = now();
  const updatedAtIndex = columns.indexOf('updatedAt');
  const rangesToWrite = [];
  for (let i = 1; i < data.length; i++) {
    const rowId = data[i][idIndex];
    if (updateMap.hasOwnProperty(rowId)) {
      data[i][positionIndex] = updateMap[rowId];
      if (updatedAtIndex !== -1) data[i][updatedAtIndex] = updatedAt;
      rangesToWrite.push({ rowIndex: i + 1, rowData: data[i] });
    }
  }
  if (rangesToWrite.length === 0) return;
  rangesToWrite.sort(function(a, b) { return a.rowIndex - b.rowIndex; });
  var batchStart = 0;
  while (batchStart < rangesToWrite.length) {
    var batchEnd = batchStart;
    while (batchEnd + 1 < rangesToWrite.length &&
           rangesToWrite[batchEnd + 1].rowIndex === rangesToWrite[batchEnd].rowIndex + 1) {
      batchEnd++;
    }
    var startRow = rangesToWrite[batchStart].rowIndex;
    var numRows = batchEnd - batchStart + 1;
    var batchData = [];
    for (var b = batchStart; b <= batchEnd; b++) {
      batchData.push(rangesToWrite[b].rowData);
    }
    sheet.getRange(startRow, 1, numRows, columns.length).setValues(batchData);
    batchStart = batchEnd + 1;
  }
  SpreadsheetApp.flush();
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

function getSprintsSheet() {
  if (!CONFIG || !CONFIG.SPRINT_COLUMNS) {
    throw new Error('CONFIG.SPRINT_COLUMNS is undefined. Check Config.gs for syntax errors.');
  }
  return getSheet(CONFIG.SHEETS.SPRINTS, CONFIG.SPRINT_COLUMNS);
}

function getAllSprints(projectId) {
  const sheet = getSprintsSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const columns = CONFIG.SPRINT_COLUMNS;
  let sprints = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    sprints.push(rowToObject(row, columns));
  }
  if (projectId) {
    sprints = sprints.filter(s => s.projectId === projectId);
  }
  return sprints;
}

function getSprintById(sprintId) {
  const sprints = getAllSprints();
  return sprints.find(s => s.id === sprintId) || null;
}

function createSprint(sprintData) {
  PermissionGuard.requirePermission('project:update', { projectId: sprintData.projectId });
  const sheet = getSprintsSheet();
  const currentUser = getCurrentUserEmail();
  const projectId = sprintData.projectId || '';
  if (projectId) {
    const activeSprint = getAllSprints(projectId).find(s => s.status === 'active');
    if (activeSprint) {
      throw new Error('Project already has an active sprint: ' + activeSprint.id);
    }
  }
  const sprint = {
    id: generateId('SPR'),
    projectId: projectId,
    name: sanitize(sprintData.name || 'New Sprint'),
    goal: sanitize(sprintData.goal || ''),
    startDate: sprintData.startDate || '',
    endDate: sprintData.endDate || '',
    status: 'planning',
    createdAt: now(),
    createdBy: currentUser,
    completedAt: ''
  };
  sheet.appendRow(objectToRow(sprint, CONFIG.SPRINT_COLUMNS));
  logActivity(currentUser, 'created', 'sprint', sprint.id, { name: sprint.name });
  return sprint;
}

function updateSprint(sprintId, updates) {
  const sheet = getSprintsSheet();
  const data = sheet.getDataRange().getValues();
  const columns = CONFIG.SPRINT_COLUMNS;
  const idIndex = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][idIndex] === sprintId) {
      const sprint = rowToObject(data[i], columns);
      Object.assign(sprint, updates);
      sheet.getRange(i + 1, 1, 1, columns.length).setValues([objectToRow(sprint, columns)]);
      return sprint;
    }
  }
  throw new Error('Sprint not found: ' + sprintId);
}

function startSprint(sprintId) {
  const sprint = getSprintById(sprintId);
  if (!sprint) throw new Error('Sprint not found: ' + sprintId);
  PermissionGuard.requirePermission('project:update', { projectId: sprint.projectId });
  if (sprint.status !== 'planning') {
    throw new Error('Only sprints in planning status can be started');
  }
  const activeSprint = getAllSprints(sprint.projectId).find(s => s.status === 'active' && s.id !== sprintId);
  if (activeSprint) {
    throw new Error('Project already has an active sprint: ' + activeSprint.id);
  }
  return updateSprint(sprintId, { status: 'active', startDate: now() });
}

function completeSprint(sprintId) {
  const sprint = getSprintById(sprintId);
  if (!sprint) throw new Error('Sprint not found: ' + sprintId);
  PermissionGuard.requirePermission('project:update', { projectId: sprint.projectId });
  if (sprint.status !== 'active') {
    throw new Error('Only active sprints can be completed');
  }
  const incompleteTasks = getAllTasks({ sprint: sprintId }, { skipPermissionCheck: true })
    .filter(t => t.status !== 'Done');
  incompleteTasks.forEach(task => {
    try {
      updateTask(task.id, { status: 'To Do', sprint: '' });
    } catch (e) {
      console.error('completeSprint: failed to roll back task ' + task.id + ':', e);
    }
  });
  return updateSprint(sprintId, { status: 'completed', completedAt: now() });
}

function getSprintBurndown(sprintId) {
  const sprint = getSprintById(sprintId);
  if (!sprint) throw new Error('Sprint not found: ' + sprintId);
  const sprintTasks = getAllTasks({ sprint: sprintId }, { skipPermissionCheck: true });
  const totalPoints = sprintTasks.reduce((sum, t) => sum + (parseInt(t.storyPoints) || 0), 0);
  if (!sprint.startDate) return { sprintId: sprintId, totalPoints: totalPoints, days: [] };
  const startDate = new Date(sprint.startDate);
  const endDate = sprint.completedAt ? new Date(sprint.completedAt) : new Date();
  const days = [];
  const cursor = new Date(startDate);
  cursor.setHours(0, 0, 0, 0);
  while (cursor <= endDate) {
    const dayEnd = new Date(cursor);
    dayEnd.setHours(23, 59, 59, 999);
    const dayIso = cursor.toISOString().split('T')[0];
    const burnedPoints = sprintTasks
      .filter(t => t.status === 'Done' && t.completedAt && new Date(t.completedAt) <= dayEnd)
      .reduce((sum, t) => sum + (parseInt(t.storyPoints) || 0), 0);
    days.push({ date: dayIso, remaining: totalPoints - burnedPoints });
    cursor.setDate(cursor.getDate() + 1);
  }
  return { sprintId: sprintId, totalPoints: totalPoints, days: days };
}

function assignProjectTasksToSprint(sprintId, projectId) {
  var allTasks = getAllTasks({}, { skipPermissionCheck: true });
  var eligibleTasks = allTasks.filter(function(t) {
    return t.projectId === projectId && (!t.sprint || t.sprint === '') && t.status !== 'Done';
  });
  var count = 0;
  eligibleTasks.forEach(function(task) {
    try {
      updateTask(task.id, { sprint: sprintId });
      count++;
    } catch (e) {
      console.error('assignProjectTasksToSprint: failed to assign task ' + task.id + ':', e);
    }
  });
  return count;
}

function syncTaskToCalendar(task) {
  if (!task || !task.dueDate) return;
  try {
    const calendar = CalendarApp.getDefaultCalendar();
    let customFields = {};
    if (task.customFields) {
      try {
        customFields = typeof task.customFields === 'string'
          ? JSON.parse(task.customFields)
          : task.customFields;
      } catch (e) {
        customFields = {};
      }
    }
    const title = '[' + (task.id || 'TASK') + '] ' + (task.title || 'Task');
    const description = (task.description || '') + '\n\nProject: ' + (task.projectId || '') + '\nStatus: ' + (task.status || '');
    const dueDate = new Date(task.dueDate);
    const existingEventId = customFields.calendarEventId || null;
    if (existingEventId) {
      try {
        const existingEvent = calendar.getEventById(existingEventId);
        if (existingEvent) {
          existingEvent.setTitle(title);
          existingEvent.setDescription(description);
          existingEvent.setAllDayDate(dueDate);
          return;
        }
      } catch (e) {
        console.error('syncTaskToCalendar: event lookup failed, creating new:', e);
      }
    }
    const newEvent = calendar.createAllDayEvent(title, dueDate, { description: description });
    customFields.calendarEventId = newEvent.getId();
    updateTask(task.id, { customFields: JSON.stringify(customFields) });
  } catch (e) {
    console.error('syncTaskToCalendar failed for task ' + (task.id || '?') + ':', e);
  }
}

function removeCalendarEvent(task) {
  if (!task) return;
  try {
    let customFields = {};
    if (task.customFields) {
      try {
        customFields = typeof task.customFields === 'string'
          ? JSON.parse(task.customFields)
          : task.customFields;
      } catch (e) {
        customFields = {};
      }
    }
    const eventId = customFields.calendarEventId;
    if (!eventId) return;
    const calendar = CalendarApp.getDefaultCalendar();
    try {
      const event = calendar.getEventById(eventId);
      if (event) event.deleteEvent();
    } catch (e) {
      console.error('removeCalendarEvent: event delete failed for task ' + (task.id || '?') + ':', e);
    }
    delete customFields.calendarEventId;
    updateTask(task.id, { customFields: JSON.stringify(customFields) });
  } catch (e) {
    console.error('removeCalendarEvent failed for task ' + (task.id || '?') + ':', e);
  }
}

function processRecurringTasks() {
  const doneTasks = getAllTasks({ status: 'Done' }, { skipPermissionCheck: true })
    .filter(t => t.recurringConfig && t.recurringConfig !== '');
  const timestamp = now();
  doneTasks.forEach(task => {
    try {
      let config = {};
      try {
        config = typeof task.recurringConfig === 'string'
          ? JSON.parse(task.recurringConfig)
          : task.recurringConfig;
      } catch (e) {
        return;
      }
      const { freq, day, interval, endDate, count } = config;
      if (!freq) return;
      if (endDate && new Date(endDate) < new Date()) return;
      const lastRecurrence = task.lastRecurrenceAt ? new Date(task.lastRecurrenceAt) : new Date(task.completedAt || task.updatedAt);
      const intervalMs = (parseInt(interval) || 1) * (freq === 'daily' ? 86400000 : freq === 'weekly' ? 604800000 : freq === 'monthly' ? 2592000000 : 0);
      if (!intervalMs) return;
      const nextDue = new Date(lastRecurrence.getTime() + intervalMs);
      if (nextDue > new Date()) return;
      if (count !== undefined && count !== null && count !== '') {
        const spawnedCount = parseInt(config._spawnedCount) || 0;
        if (spawnedCount >= parseInt(count)) return;
        config._spawnedCount = spawnedCount + 1;
      }
      const newDueDate = task.dueDate ? new Date(new Date(task.dueDate).getTime() + intervalMs).toISOString() : '';
      const newStartDate = task.startDate ? new Date(new Date(task.startDate).getTime() + intervalMs).toISOString() : '';
      const newTaskData = {
        projectId: task.projectId,
        title: task.title,
        description: task.description,
        priority: task.priority,
        type: task.type,
        assignee: task.assignee,
        reporter: task.reporter,
        dueDate: newDueDate,
        startDate: newStartDate,
        sprint: task.sprint,
        storyPoints: task.storyPoints,
        estimatedHrs: task.estimatedHrs,
        labels: task.labels,
        parentId: task.parentId,
        customFields: task.customFields,
        templateId: task.templateId,
        recurringConfig: task.recurringConfig,
        status: 'To Do'
      };
      createTask(newTaskData);
      updateTask(task.id, { lastRecurrenceAt: timestamp, recurringConfig: JSON.stringify(config) });
    } catch (e) {
      console.error('processRecurringTasks: failed to process task ' + task.id + ':', e);
    }
  });
}
