var BadgeEngine = (function() {

  var BADGES = {
    pioneer: {
      id: 'pioneer',
      name: 'Pioneer',
      icon: 'fa-rocket',
      color: '#8b5cf6',
      description: 'Part of the Alpha tester pool',
      category: 'identity',
      rarity: 'legendary',
      manual: true
    },
    founder: {
      id: 'founder',
      name: 'Founder',
      icon: 'fa-flag',
      color: '#f59e0b',
      description: 'Created an organization',
      category: 'identity',
      rarity: 'rare'
    },
    architect: {
      id: 'architect',
      name: 'Architect',
      icon: 'fa-drafting-compass',
      color: '#3b82f6',
      description: 'Created 10+ projects',
      category: 'identity',
      rarity: 'rare',
      threshold: 10
    },
    closer: {
      id: 'closer',
      name: 'Closer',
      icon: 'fa-check-double',
      color: '#10b981',
      description: 'Completed 50 tasks',
      category: 'productivity',
      rarity: 'uncommon',
      threshold: 50
    },
    centurion: {
      id: 'centurion',
      name: 'Centurion',
      icon: 'fa-shield-alt',
      color: '#ef4444',
      description: 'Completed 100 tasks',
      category: 'productivity',
      rarity: 'rare',
      threshold: 100
    },
    titan: {
      id: 'titan',
      name: 'Titan',
      icon: 'fa-mountain',
      color: '#dc2626',
      description: 'Completed 500 tasks',
      category: 'productivity',
      rarity: 'legendary',
      threshold: 500
    },
    streak_7: {
      id: 'streak_7',
      name: 'On Fire',
      icon: 'fa-fire',
      color: '#f97316',
      description: '7-day task completion streak',
      category: 'productivity',
      rarity: 'uncommon',
      threshold: 7
    },
    streak_30: {
      id: 'streak_30',
      name: 'Unstoppable',
      icon: 'fa-fire-alt',
      color: '#dc2626',
      description: '30-day task completion streak',
      category: 'productivity',
      rarity: 'epic',
      threshold: 30
    },
    speed_demon: {
      id: 'speed_demon',
      name: 'Speed Demon',
      icon: 'fa-bolt',
      color: '#eab308',
      description: 'Completed a task within 1 hour of creation',
      category: 'productivity',
      rarity: 'uncommon'
    },
    zero_inbox: {
      id: 'zero_inbox',
      name: 'Zero Inbox',
      icon: 'fa-inbox',
      color: '#06b6d4',
      description: 'Cleared all assigned tasks to Done',
      category: 'productivity',
      rarity: 'rare'
    },
    mentor: {
      id: 'mentor',
      name: 'Mentor',
      icon: 'fa-chalkboard-teacher',
      color: '#8b5cf6',
      description: 'Assigned tasks to 10+ different people',
      category: 'collaboration',
      rarity: 'rare',
      threshold: 10
    },
    first_responder: {
      id: 'first_responder',
      name: 'First Responder',
      icon: 'fa-hand-paper',
      color: '#ef4444',
      description: 'Picked up 10+ unassigned tasks',
      category: 'collaboration',
      rarity: 'uncommon',
      threshold: 10
    },
    unlocker: {
      id: 'unlocker',
      name: 'Unlocker',
      icon: 'fa-unlock-alt',
      color: '#10b981',
      description: 'Resolved 10+ tasks that were blocking others',
      category: 'collaboration',
      rarity: 'rare',
      threshold: 10
    },
    team_player: {
      id: 'team_player',
      name: 'Team Player',
      icon: 'fa-users',
      color: '#3b82f6',
      description: 'Member of 3+ teams',
      category: 'collaboration',
      rarity: 'uncommon',
      threshold: 3
    },
    night_owl: {
      id: 'night_owl',
      name: 'Night Owl',
      icon: 'fa-moon',
      color: '#6366f1',
      description: 'Completed a task between midnight and 5 AM',
      category: 'rare',
      rarity: 'extremely_rare'
    },
    early_bird: {
      id: 'early_bird',
      name: 'Early Bird',
      icon: 'fa-sun',
      color: '#f59e0b',
      description: 'Completed a task between 5 AM and 7 AM',
      category: 'rare',
      rarity: 'uncommon'
    },
    perfectionist: {
      id: 'perfectionist',
      name: 'Perfectionist',
      icon: 'fa-gem',
      color: '#ec4899',
      description: 'Had zero overdue tasks for 30 consecutive days',
      category: 'rare',
      rarity: 'epic'
    },
    marathon: {
      id: 'marathon',
      name: 'Marathon',
      icon: 'fa-running',
      color: '#14b8a6',
      description: 'Logged in 30 consecutive days',
      category: 'rare',
      rarity: 'epic',
      threshold: 30
    },
    legend: {
      id: 'legend',
      name: 'Legend',
      icon: 'fa-crown',
      color: '#f59e0b',
      description: 'Earned 10+ badges',
      category: 'rare',
      rarity: 'legendary',
      threshold: 10
    },
    multitasker: {
      id: 'multitasker',
      name: 'Multitasker',
      icon: 'fa-layer-group',
      color: '#7c3aed',
      description: 'Had 20+ tasks in progress simultaneously',
      category: 'productivity',
      rarity: 'uncommon',
      threshold: 20
    },
    critic: {
      id: 'critic',
      name: 'Critic',
      icon: 'fa-comment-dots',
      color: '#64748b',
      description: 'Left 50+ comments on tasks',
      category: 'collaboration',
      rarity: 'uncommon',
      threshold: 50
    },
    data_hoarder: {
      id: 'data_hoarder',
      name: 'Data Hoarder',
      icon: 'fa-database',
      color: '#0ea5e9',
      description: 'Created 25+ data assets',
      category: 'identity',
      rarity: 'rare',
      threshold: 25
    },
    enigma: {
      id: 'enigma',
      name: 'The Enigma',
      icon: 'fa-user-secret',
      color: '#1e1b4b',
      description: 'A convergence of rare discipline — completion, consistency, and creation.',
      category: 'rare',
      rarity: 'enigmatic',
      secret: true
    }
  };

  var RARITY_ORDER = { common: 0, uncommon: 1, rare: 2, epic: 3, extremely_rare: 4, legendary: 5, enigmatic: 6 };

  function getAllBadgeDefinitions() {
    var defs = [];
    Object.keys(BADGES).forEach(function(key) {
      defs.push(BADGES[key]);
    });
    return defs;
  }

  function getUserBadges(email) {
    if (!email) return [];
    email = email.toLowerCase();
    try {
      var sheet = getUserBadgesSheet();
      var data = sheet.getDataRange().getValues();
      var columns = CONFIG.USER_BADGE_COLUMNS;
      var emailCol = columns.indexOf('userEmail');
      var badgeIdCol = columns.indexOf('badgeId');
      var earnedCol = columns.indexOf('earnedAt');
      var out = [];
      for (var i = 1; i < data.length; i++) {
        if (data[i][emailCol] && String(data[i][emailCol]).toLowerCase() === email) {
          out.push({
            id: data[i][badgeIdCol],
            earnedAt: data[i][earnedCol] instanceof Date
              ? data[i][earnedCol].toISOString()
              : data[i][earnedCol]
          });
        }
      }
      if (out.length) return out;
    } catch (e) {
      console.error('getUserBadges sheet read failed:', e);
    }
    var user = getUserByEmail(email);
    if (!user) return [];
    var jd = _parseJsonData(user);
    return jd.badges || [];
  }

  function _hasBadge(email, badgeId) {
    var badges = getUserBadges(email);
    return badges.some(function(b) { return b.id === badgeId; });
  }

  function awardBadge(email, badgeId, options) {
    if (!BADGES[badgeId]) return null;
    if (!email) return null;
    options = options || {};
    email = email.toLowerCase();
    if (_hasBadge(email, badgeId)) return null;

    var def = BADGES[badgeId];
    var now = new Date().toISOString();
    var record = {
      id: generateId('UB'),
      userEmail: email,
      badgeId: badgeId,
      badgeName: def.name,
      category: def.category || '',
      rarity: def.rarity || '',
      earnedAt: now,
      awardedBy: options.awardedBy || 'system',
      triggerMetric: options.triggerMetric || '',
      triggerValue: options.triggerValue != null ? options.triggerValue : '',
      notes: options.notes || ''
    };

    var lock = LockService.getScriptLock();
    lock.waitLock(5000);
    try {
      var sheet = getUserBadgesSheet();
      sheet.appendRow(objectToRow(record, CONFIG.USER_BADGE_COLUMNS));
      _mirrorToUserJson_(email, { id: badgeId, earnedAt: now }, 'add');
      try {
        var badges = getUserBadges(email);
        if (typeof UserMetricsService !== 'undefined') {
          UserMetricsService.setBadgeCount(email, badges.length);
        }
      } catch (e) {
        console.error('setBadgeCount failed:', e);
      }
    } finally {
      lock.releaseLock();
    }

    try {
      if (typeof NotificationEngine !== 'undefined') {
        NotificationEngine.createNotification({
          userId: email,
          type: 'badge_earned',
          title: 'Badge earned: ' + def.name,
          message: def.description || ('You earned the ' + def.name + ' badge.'),
          entityType: 'badge',
          entityId: badgeId,
          channels: ['in_app']
        });
      }
    } catch (e) {
      console.error('Badge notification failed:', e);
    }

    return { id: badgeId, earnedAt: now };
  }

  function revokeBadge(email, badgeId) {
    if (!email) return false;
    email = email.toLowerCase();
    var lock = LockService.getScriptLock();
    lock.waitLock(5000);
    try {
      var sheet = getUserBadgesSheet();
      var data = sheet.getDataRange().getValues();
      var columns = CONFIG.USER_BADGE_COLUMNS;
      var emailCol = columns.indexOf('userEmail');
      var badgeIdCol = columns.indexOf('badgeId');
      var removed = false;
      for (var i = data.length - 1; i >= 1; i--) {
        if (data[i][emailCol] && String(data[i][emailCol]).toLowerCase() === email
            && data[i][badgeIdCol] === badgeId) {
          sheet.deleteRow(i + 1);
          removed = true;
        }
      }
      if (removed) {
        _mirrorToUserJson_(email, { id: badgeId }, 'remove');
        try {
          var badges = getUserBadges(email);
          if (typeof UserMetricsService !== 'undefined') {
            UserMetricsService.setBadgeCount(email, badges.length);
          }
        } catch (e) {}
      }
      return removed;
    } finally {
      lock.releaseLock();
    }
  }

  function _mirrorToUserJson_(email, award, mode) {
    try {
      var sheet = getUsersSheet();
      var data = sheet.getDataRange().getValues();
      var columns = CONFIG.USER_COLUMNS;
      for (var i = 1; i < data.length; i++) {
        if (data[i][0] && String(data[i][0]).toLowerCase() === email) {
          var user = rowToObject(data[i], columns);
          var jd = _parseJsonData(user);
          if (!jd.badges) jd.badges = [];
          if (mode === 'add') {
            if (!jd.badges.some(function(b) { return b.id === award.id; })) {
              jd.badges.push(award);
            }
          } else if (mode === 'remove') {
            jd.badges = jd.badges.filter(function(b) { return b.id !== award.id; });
          }
          user.jsonData = JSON.stringify(jd);
          var newRow = objectToRow(user, columns);
          sheet.getRange(i + 1, 1, 1, columns.length).setValues([newRow]);
          UserCache.invalidate();
          return;
        }
      }
    } catch (e) {
      console.error('_mirrorToUserJson_ failed:', e);
    }
  }

  function evaluateBadges(email) {
    if (!email) return [];
    email = email.toLowerCase();
    var awarded = [];

    var metrics = null;
    try {
      metrics = typeof UserMetricsService !== 'undefined'
        ? (UserMetricsService.getUserMetrics(email) || UserMetricsService.rebuildUserMetrics(email))
        : null;
    } catch (e) {
      console.error('evaluateBadges metrics load failed:', e);
    }
    if (!metrics) return [];

    var existing = getUserBadges(email).map(function(b) { return b.id; });

    function tryAward(badgeId, metricName, value, threshold) {
      if (existing.indexOf(badgeId) !== -1) return;
      if (value < threshold) return;
      var opts = {
        awardedBy: 'system',
        triggerMetric: metricName,
        triggerValue: String(value)
      };
      var r = awardBadge(email, badgeId, opts);
      if (r) {
        awarded.push(r);
        existing.push(badgeId);
      }
    }

    try {
      tryAward('closer',       'tasksCompleted',            metrics.tasksCompleted,            50);
      tryAward('centurion',    'tasksCompleted',            metrics.tasksCompleted,            100);
      tryAward('titan',        'tasksCompleted',            metrics.tasksCompleted,            500);
      tryAward('architect',    'projectsCreated',           metrics.projectsCreated,           10);
      tryAward('speed_demon',  'tasksCompletedWithinHour',  metrics.tasksCompletedWithinHour,  1);
      tryAward('unlocker',     'unblockerCount',            metrics.unblockerCount,            10);
      tryAward('mentor',       'uniqueAssigneesCount',      metrics.uniqueAssigneesCount,      10);
      tryAward('multitasker',  'inProgressPeak',            metrics.inProgressPeak,            20);
      tryAward('streak_7',     'currentCompletionStreak',   metrics.currentCompletionStreak,   7);
      tryAward('streak_30',    'currentCompletionStreak',   metrics.currentCompletionStreak,   30);
      tryAward('marathon',     'currentLoginStreak',        metrics.currentLoginStreak,        30);
      tryAward('team_player',  'teamsJoinedCount',          metrics.teamsJoinedCount,          3);
      tryAward('data_hoarder', 'dataAssetsCreated',         metrics.dataAssetsCreated,         25);
      tryAward('founder',      'orgsOwnedCount',            metrics.orgsOwnedCount,            1);

      if (existing.indexOf('zero_inbox') === -1 && metrics.tasksCompleted > 0) {
        var myOpen = getAllTasks().filter(function(t) {
          return t.assignee && t.assignee.toLowerCase() === email && t.status !== 'Done' && !t.isDeleted;
        });
        if (myOpen.length === 0) {
          var r = awardBadge(email, 'zero_inbox', { awardedBy: 'system', triggerMetric: 'openTasks', triggerValue: '0' });
          if (r) { awarded.push(r); existing.push('zero_inbox'); }
        }
      }

      if (existing.indexOf('night_owl') === -1 || existing.indexOf('early_bird') === -1) {
        if (metrics.lastCompletedAt) {
          var lastHour = new Date(metrics.lastCompletedAt).getHours();
          if (existing.indexOf('night_owl') === -1 && lastHour >= 0 && lastHour < 5) {
            var r = awardBadge(email, 'night_owl', { awardedBy: 'system', triggerMetric: 'lastCompletedHour', triggerValue: String(lastHour) });
            if (r) { awarded.push(r); existing.push('night_owl'); }
          }
          if (existing.indexOf('early_bird') === -1 && lastHour >= 5 && lastHour < 7) {
            var r = awardBadge(email, 'early_bird', { awardedBy: 'system', triggerMetric: 'lastCompletedHour', triggerValue: String(lastHour) });
            if (r) { awarded.push(r); existing.push('early_bird'); }
          }
        }
      }

      tryAward('legend', 'totalBadges', existing.length, 10);

      if (existing.indexOf('enigma') === -1
          && metrics.tasksCompleted >= 25
          && metrics.currentLoginStreak >= 30) {
        var projectsOwnedOrCreated = metrics.projectsCreated || 0;
        try {
          var ownedCount = 0;
          var allProjects = getAllProjects();
          allProjects.forEach(function(p) {
            var owner = (p.ownerId || p.createdBy || '').toLowerCase();
            if (owner === email) ownedCount++;
          });
          if (ownedCount > projectsOwnedOrCreated) projectsOwnedOrCreated = ownedCount;
        } catch (e) { console.error('enigma project count failed:', e); }

        if (projectsOwnedOrCreated >= 5) {
          var r = awardBadge(email, 'enigma', {
            awardedBy: 'system',
            triggerMetric: 'multi',
            triggerValue: 'tasks=' + metrics.tasksCompleted
              + ', loginStreak=' + metrics.currentLoginStreak
              + ', projects=' + projectsOwnedOrCreated
          });
          if (r) { awarded.push(r); existing.push('enigma'); }
        }
      }

    } catch (e) {
      console.error('BadgeEngine.evaluateBadges failed:', e);
    }

    try {
      if (typeof UserMetricsService !== 'undefined') {
        UserMetricsService.setBadgeCount(email, existing.length);
      }
    } catch (e) {}

    return awarded;
  }

  function evaluateLoginStreak(email) {
    if (!email) return null;
    email = email.toLowerCase();

    var metrics = null;
    try {
      if (typeof UserMetricsService !== 'undefined') {
        metrics = UserMetricsService.recordLogin(email);
      }
    } catch (e) {
      console.error('evaluateLoginStreak metrics update failed:', e);
    }

    if (_hasBadge(email, 'marathon')) return null;

    var streak = metrics ? metrics.currentLoginStreak : 0;
    if (streak >= 30) {
      return awardBadge(email, 'marathon', {
        awardedBy: 'system',
        triggerMetric: 'currentLoginStreak',
        triggerValue: String(streak)
      });
    }
    return null;
  }

  function getUserBadgeProfile(email) {
    if (!email) return { badges: [], definitions: getAllBadgeDefinitions() };
    email = email.toLowerCase();
    var earned = getUserBadges(email);
    var loginStreak = 0;
    var metrics = null;
    try {
      if (typeof UserMetricsService !== 'undefined') {
        metrics = UserMetricsService.getUserMetrics(email);
        if (metrics) loginStreak = metrics.currentLoginStreak || 0;
      }
    } catch (e) {}
    return {
      badges: earned,
      definitions: getAllBadgeDefinitions(),
      loginStreak: loginStreak,
      metrics: metrics,
      totalEarned: earned.length,
      totalAvailable: Object.keys(BADGES).length
    };
  }

  function _parseJsonData(user) {
    if (!user) return {};
    if (typeof user.jsonData === 'string' && user.jsonData) {
      try { return JSON.parse(user.jsonData); } catch (e) { return {}; }
    }
    return {};
  }

  function migrateBadgesFromJsonData() {
    var users = getAllUsers();
    var migrated = 0;
    users.forEach(function(u) {
      if (!u.email) return;
      var jd = _parseJsonData(u);
      if (!jd.badges || !jd.badges.length) return;
      jd.badges.forEach(function(b) {
        if (!b || !b.id || !BADGES[b.id]) return;
        if (_hasBadge(u.email, b.id)) return;
        var def = BADGES[b.id];
        var record = {
          id: generateId('UB'),
          userEmail: u.email.toLowerCase(),
          badgeId: b.id,
          badgeName: def.name,
          category: def.category || '',
          rarity: def.rarity || '',
          earnedAt: b.earnedAt || new Date().toISOString(),
          awardedBy: 'migration',
          triggerMetric: '',
          triggerValue: '',
          notes: 'Migrated from user.jsonData'
        };
        try {
          var sheet = getUserBadgesSheet();
          sheet.appendRow(objectToRow(record, CONFIG.USER_BADGE_COLUMNS));
          migrated++;
        } catch (e) {
          console.error('Badge migration failed for ' + u.email + '/' + b.id + ':', e);
        }
      });
    });
    return { migrated: migrated, users: users.length };
  }

  return {
    BADGES: BADGES,
    RARITY_ORDER: RARITY_ORDER,
    getAllBadgeDefinitions: getAllBadgeDefinitions,
    getUserBadges: getUserBadges,
    getUserBadgeProfile: getUserBadgeProfile,
    awardBadge: awardBadge,
    revokeBadge: revokeBadge,
    evaluateBadges: evaluateBadges,
    evaluateLoginStreak: evaluateLoginStreak,
    migrateBadgesFromJsonData: migrateBadgesFromJsonData
  };

})();

function getBadgeProfile(email) {
  return BadgeEngine.getUserBadgeProfile(email || getCurrentUserEmailOptimized());
}

function evaluateMyBadges() {
  var email = getCurrentUserEmailOptimized();
  if (!email) return [];
  return BadgeEngine.evaluateBadges(email);
}

function adminAwardBadge(email, badgeId) {
  var caller = getUserByEmail(getCurrentUserEmailOptimized());
  if (!caller || caller.role !== 'admin') {
    throw new Error('Permission denied: only admins can award badges');
  }
  return BadgeEngine.awardBadge(email, badgeId);
}

function adminRevokeBadge(email, badgeId) {
  var caller = getUserByEmail(getCurrentUserEmailOptimized());
  if (!caller || caller.role !== 'admin') {
    throw new Error('Permission denied: only admins can revoke badges');
  }
  return BadgeEngine.revokeBadge(email, badgeId);
}
