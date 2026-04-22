var UserMetricsService = (function() {

  var EMAIL_INDEX = 0;

  function _blankMetrics(email) {
    return {
      userEmail: email,
      tasksCreated: 0,
      tasksCompleted: 0,
      tasksCompletedWithinHour: 0,
      tasksAssignedToOthers: 0,
      uniqueAssigneesList: '[]',
      uniqueAssigneesCount: 0,
      unblockerCount: 0,
      inProgressPeak: 0,
      commentsPosted: 0,
      projectsCreated: 0,
      dataAssetsCreated: 0,
      orgsOwnedCount: 0,
      teamsJoinedCount: 0,
      completionDays: '[]',
      currentCompletionStreak: 0,
      longestCompletionStreak: 0,
      lastCompletedAt: '',
      loginDays: '[]',
      currentLoginStreak: 0,
      longestLoginStreak: 0,
      lastLoginAt: '',
      totalBadges: 0,
      lastEvaluatedAt: '',
      updatedAt: now()
    };
  }

  function _parseJsonField(value, fallback) {
    if (!value) return fallback;
    if (Array.isArray(value) || typeof value === 'object') return value;
    try { return JSON.parse(value); } catch (e) { return fallback; }
  }

  function _serializeJsonField(value) {
    if (typeof value === 'string') return value;
    try { return JSON.stringify(value || []); } catch (e) { return '[]'; }
  }

  function _findRow(sheet, email) {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var rowEmail = data[i][EMAIL_INDEX];
      if (rowEmail && String(rowEmail).toLowerCase() === email.toLowerCase()) {
        return { rowIndex: i + 1, row: data[i] };
      }
    }
    return null;
  }

  function _rowToMetrics(row, columns) {
    var obj = rowToObject(row, columns);
    ['tasksCreated', 'tasksCompleted', 'tasksCompletedWithinHour', 'tasksAssignedToOthers',
     'uniqueAssigneesCount', 'unblockerCount', 'inProgressPeak', 'commentsPosted',
     'projectsCreated', 'dataAssetsCreated', 'orgsOwnedCount', 'teamsJoinedCount',
     'currentCompletionStreak', 'longestCompletionStreak',
     'currentLoginStreak', 'longestLoginStreak', 'totalBadges'
    ].forEach(function(k) {
      obj[k] = parseInt(obj[k], 10) || 0;
    });
    return obj;
  }

  function getUserMetrics(email) {
    if (!email) return null;
    email = email.toLowerCase().trim();
    var sheet = getUserMetricsSheet();
    var columns = CONFIG.USER_METRIC_COLUMNS;
    var found = _findRow(sheet, email);
    if (!found) return null;
    return _rowToMetrics(found.row, columns);
  }

  function getOrCreateMetrics(email) {
    email = email.toLowerCase().trim();
    var existing = getUserMetrics(email);
    if (existing) return existing;

    var lock = LockService.getScriptLock();
    lock.waitLock(5000);
    try {
      existing = getUserMetrics(email);
      if (existing) return existing;
      var sheet = getUserMetricsSheet();
      var columns = CONFIG.USER_METRIC_COLUMNS;
      var metrics = _blankMetrics(email);
      sheet.appendRow(objectToRow(metrics, columns));
      return metrics;
    } finally {
      lock.releaseLock();
    }
  }

  function _writeMetrics(email, mutator) {
    email = email.toLowerCase().trim();
    var lock = LockService.getScriptLock();
    lock.waitLock(5000);
    try {
      var sheet = getUserMetricsSheet();
      var columns = CONFIG.USER_METRIC_COLUMNS;
      var found = _findRow(sheet, email);
      var metrics;
      var rowIndex;
      if (found) {
        metrics = _rowToMetrics(found.row, columns);
        rowIndex = found.rowIndex;
      } else {
        metrics = _blankMetrics(email);
        rowIndex = sheet.getLastRow() + 1;
      }
      mutator(metrics);
      metrics.userEmail = email;
      metrics.updatedAt = now();
      var newRow = objectToRow(metrics, columns);
      if (found) {
        sheet.getRange(rowIndex, 1, 1, columns.length).setValues([newRow]);
      } else {
        sheet.appendRow(newRow);
      }
      return metrics;
    } finally {
      lock.releaseLock();
    }
  }

  function recordTaskCreated(email) {
    if (!email) return null;
    return _writeMetrics(email, function(m) {
      m.tasksCreated = (m.tasksCreated || 0) + 1;
    });
  }

  function recordTaskCompleted(email, task) {
    if (!email) return null;
    return _writeMetrics(email, function(m) {
      m.tasksCompleted = (m.tasksCompleted || 0) + 1;
      if (task && task.createdAt && task.completedAt) {
        var diff = new Date(task.completedAt).getTime() - new Date(task.createdAt).getTime();
        if (diff > 0 && diff <= 3600000) {
          m.tasksCompletedWithinHour = (m.tasksCompletedWithinHour || 0) + 1;
        }
      }
      var completedAt = (task && task.completedAt) ? task.completedAt : now();
      var day = new Date(completedAt).toISOString().slice(0, 10);
      var days = _parseJsonField(m.completionDays, []);
      if (days.indexOf(day) === -1) {
        days.push(day);
        days.sort();
        if (days.length > 180) days = days.slice(-180);
      }
      m.completionDays = _serializeJsonField(days);
      var streak = _computeTrailingStreak(days);
      m.currentCompletionStreak = streak;
      if (streak > (m.longestCompletionStreak || 0)) m.longestCompletionStreak = streak;
      m.lastCompletedAt = completedAt;
    });
  }

  function recordAssigneeUsed(email, assigneeEmail) {
    if (!email || !assigneeEmail) return null;
    if (email.toLowerCase() === assigneeEmail.toLowerCase()) return null;
    return _writeMetrics(email, function(m) {
      m.tasksAssignedToOthers = (m.tasksAssignedToOthers || 0) + 1;
      var list = _parseJsonField(m.uniqueAssigneesList, []);
      var normalized = assigneeEmail.toLowerCase();
      if (list.indexOf(normalized) === -1) {
        list.push(normalized);
      }
      m.uniqueAssigneesList = _serializeJsonField(list);
      m.uniqueAssigneesCount = list.length;
    });
  }

  function recordCommentPosted(email) {
    if (!email) return null;
    return _writeMetrics(email, function(m) {
      m.commentsPosted = (m.commentsPosted || 0) + 1;
    });
  }

  function recordProjectCreated(email) {
    if (!email) return null;
    return _writeMetrics(email, function(m) {
      m.projectsCreated = (m.projectsCreated || 0) + 1;
    });
  }

  function recordDataAssetCreated(email) {
    if (!email) return null;
    return _writeMetrics(email, function(m) {
      m.dataAssetsCreated = (m.dataAssetsCreated || 0) + 1;
    });
  }

  function recordLogin(email) {
    if (!email) return null;
    return _writeMetrics(email, function(m) {
      var today = new Date().toISOString().slice(0, 10);
      var days = _parseJsonField(m.loginDays, []);
      if (days[days.length - 1] !== today) {
        days.push(today);
        days.sort();
        if (days.length > 90) days = days.slice(-90);
      }
      m.loginDays = _serializeJsonField(days);
      var streak = _computeTrailingStreak(days);
      m.currentLoginStreak = streak;
      if (streak > (m.longestLoginStreak || 0)) m.longestLoginStreak = streak;
      m.lastLoginAt = now();
    });
  }

  function recordInProgressCount(email, currentInProgress) {
    if (!email) return null;
    return _writeMetrics(email, function(m) {
      if ((currentInProgress || 0) > (m.inProgressPeak || 0)) {
        m.inProgressPeak = currentInProgress;
      }
    });
  }

  function setBadgeCount(email, totalBadges) {
    if (!email) return null;
    return _writeMetrics(email, function(m) {
      m.totalBadges = totalBadges || 0;
      m.lastEvaluatedAt = now();
    });
  }

  function _computeTrailingStreak(sortedDays) {
    if (!sortedDays || !sortedDays.length) return 0;
    var streak = 1;
    for (var i = sortedDays.length - 1; i > 0; i--) {
      var curr = new Date(sortedDays[i]);
      var prev = new Date(sortedDays[i - 1]);
      var diff = (curr.getTime() - prev.getTime()) / 86400000;
      if (Math.round(diff) === 1) {
        streak++;
      } else if (Math.round(diff) === 0) {
        continue;
      } else {
        break;
      }
    }
    return streak;
  }

  function rebuildUserMetrics(email) {
    if (!email) return null;
    email = email.toLowerCase().trim();

    var allTasks = getAllTasks();
    var allProjects = getAllProjects();

    var metrics = _blankMetrics(email);

    var myTasks = allTasks.filter(function(t) {
      return t.assignee && t.assignee.toLowerCase() === email;
    });
    var myCompleted = myTasks.filter(function(t) { return t.status === 'Done'; });

    metrics.tasksCompleted = myCompleted.length;

    metrics.tasksCompletedWithinHour = myCompleted.reduce(function(acc, t) {
      if (!t.createdAt || !t.completedAt) return acc;
      var diff = new Date(t.completedAt).getTime() - new Date(t.createdAt).getTime();
      return (diff > 0 && diff <= 3600000) ? acc + 1 : acc;
    }, 0);

    metrics.tasksCreated = allTasks.filter(function(t) {
      return t.reporter && t.reporter.toLowerCase() === email;
    }).length;

    var assignees = {};
    var tasksToOthers = 0;
    allTasks.forEach(function(t) {
      if (t.reporter && t.reporter.toLowerCase() === email &&
          t.assignee && t.assignee.toLowerCase() !== email) {
        tasksToOthers++;
        assignees[t.assignee.toLowerCase()] = true;
      }
    });
    metrics.tasksAssignedToOthers = tasksToOthers;
    var assigneeList = Object.keys(assignees);
    metrics.uniqueAssigneesList = _serializeJsonField(assigneeList);
    metrics.uniqueAssigneesCount = assigneeList.length;

    var unblocker = 0;
    myCompleted.forEach(function(t) {
      var wasBlocking = allTasks.some(function(other) {
        if (!other.dependencies) return false;
        var deps = typeof other.dependencies === 'string'
          ? other.dependencies.split(',').map(function(d) { return d.trim(); })
          : (Array.isArray(other.dependencies) ? other.dependencies : []);
        return deps.indexOf(t.id) !== -1;
      });
      if (wasBlocking) unblocker++;
    });
    metrics.unblockerCount = unblocker;

    metrics.inProgressPeak = myTasks.filter(function(t) {
      return t.status === 'In Progress';
    }).length;

    metrics.projectsCreated = allProjects.filter(function(p) {
      var owner = p.createdBy || p.ownerId;
      return owner && owner.toLowerCase() === email;
    }).length;

    try {
      var assets = typeof getAllDataAssets === 'function' ? getAllDataAssets() : [];
      metrics.dataAssetsCreated = assets.filter(function(a) {
        return a.assetOwner && a.assetOwner.toLowerCase() === email;
      }).length;
    } catch (e) {}

    try {
      var orgs = typeof getAllOrganizations === 'function' ? getAllOrganizations() : [];
      metrics.orgsOwnedCount = orgs.filter(function(o) {
        return o.ownerId && o.ownerId.toLowerCase() === email;
      }).length;
    } catch (e) {}

    try {
      if (typeof getTeamMembershipsByUser === 'function') {
        var memberships = getTeamMembershipsByUser(email) || [];
        metrics.teamsJoinedCount = memberships.length;
      }
    } catch (e) {}

    var daySet = {};
    myCompleted.forEach(function(t) {
      if (t.completedAt) {
        var day = new Date(t.completedAt).toISOString().slice(0, 10);
        daySet[day] = true;
      }
    });
    var completionDays = Object.keys(daySet).sort();
    metrics.completionDays = _serializeJsonField(completionDays.slice(-180));
    metrics.currentCompletionStreak = _computeTrailingStreak(completionDays);
    metrics.longestCompletionStreak = _longestStreak(completionDays);
    if (myCompleted.length) {
      var lastCompleted = myCompleted
        .map(function(t) { return t.completedAt; })
        .filter(Boolean)
        .sort();
      metrics.lastCompletedAt = lastCompleted[lastCompleted.length - 1] || '';
    }

    try {
      var user = getUserByEmail(email);
      if (user && user.jsonData) {
        var jd = typeof user.jsonData === 'string' ? JSON.parse(user.jsonData) : user.jsonData;
        if (jd && Array.isArray(jd.loginDays)) {
          var loginDays = jd.loginDays.slice().sort();
          metrics.loginDays = _serializeJsonField(loginDays.slice(-90));
          metrics.currentLoginStreak = _computeTrailingStreak(loginDays);
          metrics.longestLoginStreak = _longestStreak(loginDays);
          if (user.lastLogin) metrics.lastLoginAt = user.lastLogin;
        }
        if (jd && Array.isArray(jd.badges)) {
          metrics.totalBadges = jd.badges.length;
        }
      }
    } catch (e) {}

    try {
      if (typeof BadgeEngine !== 'undefined') {
        var sheetBadges = BadgeEngine.getUserBadges(email) || [];
        if (sheetBadges.length > (metrics.totalBadges || 0)) {
          metrics.totalBadges = sheetBadges.length;
        }
      }
    } catch (e) {}

    metrics.lastEvaluatedAt = now();
    metrics.updatedAt = now();

    var lock = LockService.getScriptLock();
    lock.waitLock(5000);
    try {
      var sheet = getUserMetricsSheet();
      var columns = CONFIG.USER_METRIC_COLUMNS;
      var found = _findRow(sheet, email);
      var newRow = objectToRow(metrics, columns);
      if (found) {
        sheet.getRange(found.rowIndex, 1, 1, columns.length).setValues([newRow]);
      } else {
        sheet.appendRow(newRow);
      }
    } finally {
      lock.releaseLock();
    }

    return metrics;
  }

  function _longestStreak(sortedDays) {
    if (!sortedDays || !sortedDays.length) return 0;
    var longest = 1;
    var current = 1;
    for (var i = 1; i < sortedDays.length; i++) {
      var curr = new Date(sortedDays[i]);
      var prev = new Date(sortedDays[i - 1]);
      var diff = Math.round((curr.getTime() - prev.getTime()) / 86400000);
      if (diff === 1) {
        current++;
        if (current > longest) longest = current;
      } else if (diff === 0) {
        continue;
      } else {
        current = 1;
      }
    }
    return longest;
  }

  function rebuildAllMetrics() {
    PermissionGuard.requirePermission('admin:settings');
    var users = getAllUsers();
    var rebuilt = 0;
    var errors = [];
    users.forEach(function(u) {
      if (!u.email) return;
      try {
        rebuildUserMetrics(u.email);
        rebuilt++;
      } catch (e) {
        console.error('rebuildAllMetrics failed for ' + u.email + ':', e);
        errors.push({ email: u.email, error: e.message });
      }
    });
    return { rebuilt: rebuilt, errors: errors, totalUsers: users.length };
  }

  function getAllMetrics() {
    var sheet = getUserMetricsSheet();
    var data = sheet.getDataRange().getValues();
    var columns = CONFIG.USER_METRIC_COLUMNS;
    var out = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][EMAIL_INDEX]) continue;
      out.push(_rowToMetrics(data[i], columns));
    }
    return out;
  }

  return {
    getUserMetrics: getUserMetrics,
    getOrCreateMetrics: getOrCreateMetrics,
    getAllMetrics: getAllMetrics,
    recordTaskCreated: recordTaskCreated,
    recordTaskCompleted: recordTaskCompleted,
    recordAssigneeUsed: recordAssigneeUsed,
    recordCommentPosted: recordCommentPosted,
    recordProjectCreated: recordProjectCreated,
    recordDataAssetCreated: recordDataAssetCreated,
    recordLogin: recordLogin,
    recordInProgressCount: recordInProgressCount,
    setBadgeCount: setBadgeCount,
    rebuildUserMetrics: rebuildUserMetrics,
    rebuildAllMetrics: rebuildAllMetrics
  };

})();

function getMyMetrics() {
  var email = getCurrentUserEmailOptimized();
  if (!email) return null;
  return UserMetricsService.getUserMetrics(email) || UserMetricsService.getOrCreateMetrics(email);
}

function adminRebuildMetrics(email) {
  var caller = getUserByEmail(getCurrentUserEmailOptimized());
  if (!caller || caller.role !== 'admin') {
    throw new Error('Permission denied: only admins can rebuild metrics');
  }
  if (email) return UserMetricsService.rebuildUserMetrics(email);
  return UserMetricsService.rebuildAllMetrics();
}
