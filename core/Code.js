function doGet(e) {
  try {
    initializeSystem();
    const template = HtmlService.createTemplateFromFile('ui/Index');
    return template.evaluate()
      .setTitle('COLONY')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  } catch (error) {
    console.error('doGet error:', error);
    return HtmlService.createHtmlOutput(getErrorPage(error.message));
  }
}

function include(filename) {
  var attempts = [filename];
  if (filename.includes('/')) {
    attempts.push(filename.split('/').pop());
  }
  for (var i = 0; i < attempts.length; i++) {
    try {
      return HtmlService.createHtmlOutputFromFile(attempts[i]).getContent();
    } catch (e) {
      if (i === attempts.length - 1) {
        console.error('include failed for: ' + filename);
        return '';
      }
    }
  }
  return '';
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
  _dependencies: null,

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

  getDependencies() {
    if (this._dependencies === null) {
      this._dependencies = getAllDependenciesMap();
    }
    return this._dependencies;
  },

  clear() {
    this._tasks = null;
    this._projects = null;
    this._users = null;
    this._activity = null;
    this._dependencies = null;
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
    var allProjects = getAllProjectsOptimized();
    var projects = filterProjectsByUserRole(allProjects, tasks, userEmail, userRole);
    var users = getActiveUsersOptimized();

    return {
      success: true,
      tasks: tasks,
      projects: projects,
      users: users,
      userRole: userRole
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
