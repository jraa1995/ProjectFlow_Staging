function doGet(e) {
  try {
    initializeSystem();
    const template = HtmlService.createTemplateFromFile('ui/Index');
    var params = (e && e.parameter) || {};
    template.deepLink = {
      projectId: params.projectId || '',
      tab: params.tab || '',
      taskId: params.taskId || '',
      dataAssetId: params.dataAssetId || '',
      view: params.view || ''
    };
    return template.evaluate()
      .setTitle('COLONY')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  } catch (error) {
    console.error('doGet error:', error);
    return HtmlService.createHtmlOutput(getErrorPage(error.message));
  }
}

function getPickerOAuthToken() {
  return ScriptApp.getOAuthToken();
}

function getPickerConfig() {
  try {
    var props = PropertiesService.getScriptProperties();
    return {
      token: ScriptApp.getOAuthToken(),
      developerKey: props.getProperty('PICKER_API_KEY') || '',
      appId: props.getProperty('PICKER_APP_ID') || ''
    };
  } catch (error) {
    console.error('getPickerConfig failed:', error);
    return { token: '', developerKey: '', appId: '' };
  }
}

function include(filename) {
  var attempts = [filename];
  if (filename.includes('/')) {
    attempts.push(filename.split('/').pop());
  } else {
    attempts.push('ui/' + filename);
  }
  for (var i = 0; i < attempts.length; i++) {
    try {
      return HtmlService.createTemplateFromFile(attempts[i]).evaluate().getContent();
    } catch (e) {
      if (i === attempts.length - 1) {
        console.error('include failed for: ' + filename + ' (tried: ' + attempts.join(', ') + ')');
        return '<script>window._includeFailed = window._includeFailed || []; window._includeFailed.push("' + filename + '");</script>';
      }
    }
  }
  return '';
}

function testIncludes() {
  var files = ['ui/HomeView', 'ui/FunnelView', 'ui/ListView', 'ui/ProjectView', 'ui/Scripts', 'ui/Utils',
               'ui/DataSync', 'ui/BoardView', 'ui/LoginModule', 'ui/Dialogs', 'ui/Sidebar',
               'ui/AnalyticsModule', 'ui/Shortcuts',
               'ui/SurgeView',
               'ui/TaskDetailPanel', 'ui/TaskEditModal', 'ui/MentionAutocomplete',
               'ui/DependencyPicker', 'ui/NotificationCenter', 'ui/CacheService',
               'ui/DataAssetView', 'ui/DataAssetViewScript',
               'ui/AdHocView', 'ui/BadgeShowcase'];
  var results = {};
  files.forEach(function(f) {
    try {
      var content = HtmlService.createTemplateFromFile(f).evaluate().getContent();
      var hasScript = content.indexOf('<script') !== -1 || content.indexOf('function ') !== -1;
      results[f] = 'OK (' + content.length + ' chars, scripts: ' + hasScript + ')';
    } catch (e) {
      results[f] = 'FAIL: ' + e.message;
    }
  });
  Logger.log(JSON.stringify(results, null, 2));
  return results;
}

function getErrorPage(message) {
  const safeMessage = String(message || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#x27;');

  return `
  <!DOCTYPE html>
  <html>
  <head>
  <title>COLONY - Setup Required</title>
  <style>
  body { font-family: -apple-system, sans-serif; padding: 40px; background: #f8fafc; }
  .container { max-width: 500px; margin: 0 auto; background: white; padding: 40px;
    border-radius: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); text-align: center; }
  h1 { color: #171717; }
  .error { color: #dc2626; font-size: 14px; background: #fef2f2; padding: 10px; border-radius: 4px; margin: 20px 0; }
  button { padding: 12px 24px; background: #404040; color: white; border: none;
    border-radius: 8px; cursor: pointer; font-size: 16px; }
  button:hover { background: #262626; }
  code { background: #f1f5f9; padding: 2px 8px; border-radius: 4px; }
  </style>
  </head>
  <body>
  <div class="container">
  <h1>Setup Required</h1>
  <p>Please run <code>quickSetup()</code> in the Apps Script editor.</p>
  <div class="error">${safeMessage}</div>
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

    var cache = CacheService.getScriptCache();
    if (cache.get('SYSTEM_INITIALIZED')) {
      getCurrentUser();
      return;
    }

    WORKBOOK_SHEET_SPEC.forEach(function(spec) {
      try {
        var columns = CONFIG[spec.columnsKey];
        if (!Array.isArray(columns)) {
          console.error('Missing columns for ' + spec.name + ' (' + spec.columnsKey + ')');
          return;
        }
        getSheet(spec.name, columns);
      } catch (e) {
        console.error('Failed to ensure sheet ' + spec.name + ':', e.message);
        throw e;
      }
    });

    try {
      PermissionGuard.initializeDefaultRoles();
    } catch (e) {
      console.error('Failed to seed default roles:', e.message);
    }

    cache.put('SYSTEM_INITIALIZED', '1', 3600);
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
  _dependencies: null,
  _taskIndex: null,
  _dataAssets: null,

  getTasks() {
    if (this._tasks === null) {
      this._tasks = getAllTasks();
    }
    return this._tasks;
  },

  getTaskIndex() {
    if (this._taskIndex) return this._taskIndex;
    var tasks = this.getTasks();
    var idx = { byProject: {}, byAssignee: {}, byStatus: {}, byId: {} };
    for (var i = 0; i < tasks.length; i++) {
      var t = tasks[i];
      idx.byId[t.id] = t;
      var p = t.projectId || '__none__';
      if (!idx.byProject[p]) idx.byProject[p] = [];
      idx.byProject[p].push(t);
      var a = (t.assignee || 'Unassigned').toLowerCase();
      if (!idx.byAssignee[a]) idx.byAssignee[a] = [];
      idx.byAssignee[a].push(t);
      var s = t.status || 'Backlog';
      if (!idx.byStatus[s]) idx.byStatus[s] = [];
      idx.byStatus[s].push(t);
    }
    this._taskIndex = idx;
    return idx;
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

  getDependencies() {
    if (this._dependencies === null) {
      this._dependencies = getAllDependenciesMap();
    }
    return this._dependencies;
  },

  getDataAssets() {
    if (this._dataAssets === null) {
      this._dataAssets = getAllDataAssetsOptimized();
    }
    return this._dataAssets;
  },

  clear() {
    this._tasks = null;
    this._projects = null;
    this._users = null;
    this._activity = null;
    this._dependencies = null;
    this._taskIndex = null;
    this._dataAssets = null;
  }
};

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

function escapeHtml_(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#x27;');
}

function formatCommentContent(content, mentionedUserEmails) {
  if (!content) return '';

  let formatted = content;

  formatted = formatted.replace(/@\[([^\]]+)\]\([^)]+\)/g, function(match, name) {
    return '<span class="mention">@' + escapeHtml_(name) + '</span>';
  });

  mentionedUserEmails.forEach(email => {
    const user = getUserByEmail(email);
    const displayName = user ? user.name : email;
    const regex = new RegExp('@' + email.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
    formatted = formatted.replace(regex, '<span class="mention">@' + escapeHtml_(displayName) + '</span>');
  });

  formatted = formatted.replace(/@([A-Za-z][A-Za-z0-9 ]*[A-Za-z0-9])(?![^<]*<\/span>)(?=\s|$|[.,!?])/g, function(match, name) {
    return '<span class="mention">@' + escapeHtml_(name) + '</span>';
  });

  return formatted;
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
      user: sanitizeUserForClient(user),
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

    const columns = CONFIG.STATUSES.map((status, index) => {
      const columnTasks = userTasks.filter(t => t.status === status);
      return {
        id: status.toLowerCase().replace(/\s+/g, '_'),
        name: status,
        color: CONFIG.COLORS[status] || '#6B7280',
        order: index,
        tasks: columnTasks,
        count: columnTasks.length
      };
    });

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
      user: sanitizeUserForClient(user),
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
    var userEmail = getCurrentUserEmailOptimized();
    var userRole = getCurrentUserRole();
    var allTasks = getAllTasksOptimized();
    var tasks = filterTasksByUserRole(allTasks, userEmail, userRole);
    tasks = tasks.filter(function(t) {
      var pid = (t.projectId || '').toUpperCase();
      return pid !== 'ADHOC' && pid !== 'TASK' && pid !== '';
    });
    var allProjects = getAllProjectsOptimized();
    var projects = filterProjectsByUserRole(allProjects, tasks, userEmail, userRole);
    var users = getActiveUsersOptimized();
    var dataAssets = getAllDataAssetsOptimized();

    return {
      success: true,
      tasks: tasks,
      projects: projects,
      users: users,
      dataAssets: dataAssets,
      userRole: userRole
    };
  } catch (error) {
    console.error('getListViewData failed:', error);
    return {
      success: false,
      error: error.message,
      tasks: [],
      projects: [],
      users: [],
      dataAssets: []
    };
  }
}

function getAdHocData() {
  try {
    var userEmail = getCurrentUserEmailOptimized();
    var userRole = getCurrentUserRole();
    var allTasks = getAllTasksOptimized();
    var tasks = filterTasksByUserRole(allTasks, userEmail, userRole);
    tasks = tasks.filter(function(t) {
      var pid = (t.projectId || '').toUpperCase();
      return pid === 'ADHOC' || pid === 'TASK' || pid === '';
    });
    var users = getActiveUsersOptimized();
    return { success: true, tasks: tasks, users: users };
  } catch (error) {
    console.error('getAdHocData failed:', error);
    return { success: false, error: error.message, tasks: [], users: [] };
  }
}

function getRefreshData() {
  try {
    var batchData = getBatchDataFast();
    return {
      success: true,
      tasks: batchData.tasks,
      projects: batchData.projects,
      users: batchData.users,
      config: batchData.config
    };
  } catch (error) {
    console.error('getRefreshData failed:', error);
    return { success: false, error: error.message, tasks: [], projects: [], users: [] };
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

function registerNewUser() {
  return { success: false, error: 'Self-registration is no longer available. Please request access through the application.' };
}

function getSessionUser() {
  try {
    var activeUser = Session.getActiveUser();
    if (!activeUser) {
      return { authenticated: false, reason: 'no_session' };
    }
    var email = activeUser.getEmail();
    if (!email) {
      return { authenticated: false, reason: 'no_session' };
    }
    email = email.toLowerCase().trim();

    var user = getUserByEmailOptimized(email);

    if (!user) {
      return { authenticated: false, reason: 'no_account', email: email };
    }

    var isActive = user.active === true || user.active === 'true' || user.active === 'TRUE'
      || user.active === 'yes' || user.active === 1 || user.active === '1' || user.active === 'active';
    if (!isActive && user.active !== '' && user.active !== undefined) {
      return { authenticated: false, reason: 'disabled' };
    }

    try {
      var sheet = getUsersSheet();
      var data = sheet.getDataRange().getValues();
      var columns = CONFIG.USER_COLUMNS;
      var emailIndex = 0;
      var lastLoginIndex = columns.indexOf('lastLogin');
      if (lastLoginIndex !== -1) {
        for (var i = 1; i < data.length; i++) {
          if (data[i][emailIndex] && data[i][emailIndex].toString().toLowerCase() === email) {
            sheet.getRange(i + 1, lastLoginIndex + 1).setValue(new Date().toISOString());
            break;
          }
        }
      }
    } catch (e) {
      console.error('Failed to update lastLogin:', e);
    }

    try { BadgeEngine.evaluateLoginStreak(email); } catch (e) {}

    var payload = buildLoginPayload_(user, email);
    payload.authenticated = true;
    return payload;
  } catch (error) {
    console.error('getSessionUser failed:', error);
    return { authenticated: false, reason: 'error', error: error.message };
  }
}

function requestAccess() {
  try {
    var activeUser = Session.getActiveUser();
    if (!activeUser) {
      return { success: false, error: 'Could not identify your Google account' };
    }
    var email = activeUser.getEmail();
    if (!email) {
      return { success: false, error: 'Could not identify your Google account' };
    }
    email = email.toLowerCase().trim();

    var existing = getUserByEmailOptimized(email);
    if (existing) {
      return { success: false, error: 'Account already exists' };
    }

    var sheet = getAccessRequestsSheet();
    var data = sheet.getDataRange().getValues();
    var columns = CONFIG.ACCESS_REQUEST_COLUMNS;
    for (var i = 1; i < data.length; i++) {
      var row = rowToObject(data[i], columns);
      if (row.email && row.email.toString().toLowerCase() === email && row.status === 'pending') {
        return { success: false, error: 'Request already pending. Please wait for admin approval.' };
      }
    }

    var request = {
      id: generateId('AR'),
      email: email,
      name: email.split('@')[0],
      requestedAt: now(),
      status: 'pending',
      reviewedBy: '',
      reviewedAt: '',
      reason: ''
    };
    sheet.appendRow(objectToRow(request, columns));

    try {
      var admins = getActiveUsers().filter(function(u) { return u.role === 'admin'; });
      admins.forEach(function(admin) {
        GmailApp.sendEmail(admin.email, 'COLONY — New Access Request', '', {
          htmlBody: '<div style="font-family: sans-serif; max-width: 600px; margin: 0 auto;">' +
            '<div style="background: #f8fafc; padding: 20px; border-radius: 8px; margin-bottom: 20px;">' +
            '<h2 style="color: #525252; margin: 0;">COLONY</h2></div>' +
            '<div style="padding: 20px;">' +
            '<p><strong>' + email + '</strong> is requesting access to COLONY.</p>' +
            '<p style="color: #6b7280; font-size: 14px;">Requested: ' + new Date().toLocaleString() + '</p>' +
            '<p>Open COLONY and navigate to <strong>Settings → Access Requests</strong> to approve or reject this request.</p>' +
            '</div></div>',
          name: 'COLONY'
        });
      });
    } catch (e) {
      console.error('Failed to notify admins of access request:', e);
    }

    return { success: true, message: 'Access request submitted. You will receive an email when approved.' };
  } catch (error) {
    console.error('requestAccess failed:', error);
    return { success: false, error: error.message || 'Failed to submit access request' };
  }
}

function getPendingAccessRequests() {
  try {
    PermissionGuard.requirePermission('admin:settings');
    var sheet = getAccessRequestsSheet();
    var data = sheet.getDataRange().getValues();
    var columns = CONFIG.ACCESS_REQUEST_COLUMNS;
    var requests = [];
    for (var i = 1; i < data.length; i++) {
      var row = rowToObject(data[i], columns);
      if (row.email && row.status === 'pending') {
        requests.push(row);
      }
    }
    return { success: true, requests: requests };
  } catch (error) {
    console.error('getPendingAccessRequests failed:', error);
    return { success: false, error: error.message };
  }
}

function approveAccessRequest(email) {
  try {
    PermissionGuard.requirePermission('admin:settings');
    if (!email) {
      return { success: false, error: 'Email is required' };
    }
    email = email.toLowerCase().trim();
    var adminEmail = getCurrentUserEmailOptimized();

    var sheet = getAccessRequestsSheet();
    var data = sheet.getDataRange().getValues();
    var columns = CONFIG.ACCESS_REQUEST_COLUMNS;
    var foundRow = -1;
    for (var i = 1; i < data.length; i++) {
      var row = rowToObject(data[i], columns);
      if (row.email && row.email.toString().toLowerCase() === email && row.status === 'pending') {
        foundRow = i;
        break;
      }
    }
    if (foundRow === -1) {
      return { success: false, error: 'No pending request found for this email' };
    }

    var newUser = createUser({ email: email, name: email.split('@')[0], role: 'member' });
    if (!newUser || !newUser.email) {
      return { success: false, error: 'Failed to create user account for ' + email };
    }

    var approvedRow = rowToObject(data[foundRow], columns);
    approvedRow.status = 'approved';
    approvedRow.reviewedBy = adminEmail;
    approvedRow.reviewedAt = now();
    var newRowData = objectToRow(approvedRow, columns);
    sheet.getRange(foundRow + 1, 1, 1, columns.length).setValues([newRowData]);

    try {
      GmailApp.sendEmail(email, 'COLONY — Access Granted', '', {
        htmlBody: '<div style="font-family: sans-serif; max-width: 600px; margin: 0 auto;">' +
          '<div style="background: #f8fafc; padding: 20px; border-radius: 8px; margin-bottom: 20px;">' +
          '<h2 style="color: #525252; margin: 0;">COLONY</h2></div>' +
          '<div style="padding: 20px;">' +
          '<p>Your access to COLONY has been approved.</p>' +
          '<p>You can now open the application and start using it.</p>' +
          '</div></div>',
        name: 'COLONY'
      });
    } catch (e) {
      console.error('Failed to send approval email:', e);
    }

    return { success: true };
  } catch (error) {
    console.error('approveAccessRequest failed:', error);
    return { success: false, error: error.message };
  }
}

function rejectAccessRequest(email, reason) {
  try {
    PermissionGuard.requirePermission('admin:settings');
    if (!email) {
      return { success: false, error: 'Email is required' };
    }
    email = email.toLowerCase().trim();
    var adminEmail = getCurrentUserEmailOptimized();

    var sheet = getAccessRequestsSheet();
    var data = sheet.getDataRange().getValues();
    var columns = CONFIG.ACCESS_REQUEST_COLUMNS;
    var found = false;
    for (var i = 1; i < data.length; i++) {
      var row = rowToObject(data[i], columns);
      if (row.email && row.email.toString().toLowerCase() === email && row.status === 'pending') {
        row.status = 'rejected';
        row.reviewedBy = adminEmail;
        row.reviewedAt = now();
        row.reason = reason || '';
        var newRow = objectToRow(row, columns);
        sheet.getRange(i + 1, 1, 1, columns.length).setValues([newRow]);
        found = true;
        break;
      }
    }
    if (!found) {
      return { success: false, error: 'No pending request found for this email' };
    }

    try {
      GmailApp.sendEmail(email, 'COLONY — Access Request Update', '', {
        htmlBody: '<div style="font-family: sans-serif; max-width: 600px; margin: 0 auto;">' +
          '<div style="background: #f8fafc; padding: 20px; border-radius: 8px; margin-bottom: 20px;">' +
          '<h2 style="color: #525252; margin: 0;">COLONY</h2></div>' +
          '<div style="padding: 20px;">' +
          '<p>We need additional verification for your access request.</p>' +
          (reason ? '<p><strong>Details:</strong> ' + reason + '</p>' : '') +
          '<p>Please reach out to your administrator if you believe this is an error.</p>' +
          '</div></div>',
        name: 'COLONY'
      });
    } catch (e) {
      console.error('Failed to send rejection email:', e);
    }

    return { success: true };
  } catch (error) {
    console.error('rejectAccessRequest failed:', error);
    return { success: false, error: error.message };
  }
}
