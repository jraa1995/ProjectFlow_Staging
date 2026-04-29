var ContributionWeights = {
  TASK_TYPE: { 'Task': 2, 'Bug': 3, 'Feature': 4, 'Enhancement': 3, 'Story': 4, 'Epic': 6, 'Spike': 2 },
  TASK_TYPE_DEFAULT: 2,
  TASK_SHIPPED_BONUS: 5,
  COMMENT: 1,
  COMMENT_DAILY_CAP: 3,
  WAGGLES_QUESTION: 2,
  WAGGLES_ANSWER: 4,
  WAGGLES_ANSWER_ACCEPTED_BONUS: 5,
  WAGGLES_VOTE_RECEIVED: 1,
  WAGGLES_VOTE_DAILY_CAP: 5,
  ASSIGNEE_PICKUP: 2
};

function _contribDateKey_(value) {
  if (!value) return '';
  var d = (value instanceof Date) ? value : new Date(value);
  if (isNaN(d.getTime())) return '';
  var tz = (CONFIG && CONFIG.TIMEZONE) || Session.getScriptTimeZone() || 'America/New_York';
  return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
}

function _contribInRange_(dateKey, startKey, endKey) {
  if (!dateKey) return false;
  return dateKey >= startKey && dateKey <= endKey;
}

function _contribWindowKeys_(year) {
  var tz = (CONFIG && CONFIG.TIMEZONE) || Session.getScriptTimeZone() || 'America/New_York';
  var endDate, startDate;
  if (year && /^\d{4}$/.test(String(year))) {
    var y = parseInt(year, 10);
    startDate = new Date(y, 0, 1);
    endDate = new Date(y, 11, 31, 23, 59, 59);
  } else {
    endDate = new Date();
    startDate = new Date(endDate.getTime() - 365 * 24 * 60 * 60 * 1000);
  }
  return {
    startKey: Utilities.formatDate(startDate, tz, 'yyyy-MM-dd'),
    endKey: Utilities.formatDate(endDate, tz, 'yyyy-MM-dd')
  };
}

function _contribCollectTasks_(email, startKey, endKey) {
  var sheet = getTasksSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 1, lastRow - 1, CONFIG.TASK_COLUMNS.length).getValues();
  var meLower = String(email).toLowerCase();
  var events = [];
  for (var i = 0; i < data.length; i++) {
    if (!data[i][0]) continue;
    var t = rowToObject(data[i], CONFIG.TASK_COLUMNS);
    if (t.isDeleted === true || t.isDeleted === 'true') continue;
    if (String(t.assignee || '').toLowerCase() !== meLower) continue;
    if (t.status === 'Done' && t.completedAt) {
      var doneKey = _contribDateKey_(t.completedAt);
      if (_contribInRange_(doneKey, startKey, endKey)) {
        var typeWeight = ContributionWeights.TASK_TYPE[t.type] || ContributionWeights.TASK_TYPE_DEFAULT;
        events.push({
          date: doneKey,
          source: 'task',
          points: typeWeight,
          label: 'Closed: ' + (t.title || t.id),
          taskId: t.id,
          projectId: t.projectId
        });
        if (t.shippedRef) {
          var shipKey = _contribDateKey_(t.shippedAt || t.completedAt);
          if (_contribInRange_(shipKey, startKey, endKey)) {
            events.push({
              date: shipKey,
              source: 'ship',
              points: ContributionWeights.TASK_SHIPPED_BONUS,
              label: 'Shipped: ' + (t.title || t.id),
              taskId: t.id,
              shippedRef: t.shippedRef,
              projectId: t.projectId
            });
          }
        }
      }
    }
  }
  return events;
}

function _contribCollectAssigneePickups_(email, startKey, endKey) {
  var meLower = String(email).toLowerCase();
  try {
    var sheet = getActivitySheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    var data = sheet.getRange(2, 1, lastRow - 1, CONFIG.ACTIVITY_COLUMNS.length).getValues();
    var actionIdx = CONFIG.ACTIVITY_COLUMNS.indexOf('action');
    var userIdx = CONFIG.ACTIVITY_COLUMNS.indexOf('userId');
    var entityTypeIdx = CONFIG.ACTIVITY_COLUMNS.indexOf('entityType');
    var entityIdIdx = CONFIG.ACTIVITY_COLUMNS.indexOf('entityId');
    var detailsIdx = CONFIG.ACTIVITY_COLUMNS.indexOf('details');
    var tsIdx = CONFIG.ACTIVITY_COLUMNS.indexOf('timestamp');
    var events = [];
    for (var i = 0; i < data.length; i++) {
      if (!data[i][0]) continue;
      var action = String(data[i][actionIdx] || '');
      if (action !== 'updated') continue;
      if (String(data[i][entityTypeIdx]) !== 'task') continue;
      if (String(data[i][userIdx] || '').toLowerCase() !== meLower) continue;
      var details = data[i][detailsIdx];
      if (!details) continue;
      var parsed = null;
      try { parsed = (typeof details === 'string') ? JSON.parse(details) : details; } catch (e) { continue; }
      if (!parsed || !parsed.assignee) continue;
      var to = String(parsed.assignee.to || '').toLowerCase();
      var from = String(parsed.assignee.from || '').toLowerCase();
      if (to !== meLower) continue;
      if (from === meLower) continue;
      var key = _contribDateKey_(data[i][tsIdx]);
      if (!_contribInRange_(key, startKey, endKey)) continue;
      events.push({
        date: key,
        source: 'pickup',
        points: ContributionWeights.ASSIGNEE_PICKUP,
        label: 'Picked up task ' + data[i][entityIdIdx],
        taskId: data[i][entityIdIdx]
      });
    }
    return events;
  } catch (e) {
    console.error('_contribCollectAssigneePickups_ failed:', e);
    return [];
  }
}

function _contribCollectComments_(email, startKey, endKey) {
  try {
    var sheet = getCommentsSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    var data = sheet.getRange(2, 1, lastRow - 1, CONFIG.COMMENT_COLUMNS.length).getValues();
    var userIdx = CONFIG.COMMENT_COLUMNS.indexOf('userId');
    var createdIdx = CONFIG.COMMENT_COLUMNS.indexOf('createdAt');
    var taskIdx = CONFIG.COMMENT_COLUMNS.indexOf('taskId');
    var meLower = String(email).toLowerCase();
    var events = [];
    for (var i = 0; i < data.length; i++) {
      if (!data[i][0]) continue;
      if (String(data[i][userIdx] || '').toLowerCase() !== meLower) continue;
      var key = _contribDateKey_(data[i][createdIdx]);
      if (!_contribInRange_(key, startKey, endKey)) continue;
      events.push({
        date: key,
        source: 'comment',
        points: ContributionWeights.COMMENT,
        label: 'Comment',
        taskId: data[i][taskIdx]
      });
    }
    return events;
  } catch (e) {
    console.error('_contribCollectComments_ failed:', e);
    return [];
  }
}

function _contribCollectWaggles_(email, startKey, endKey) {
  var events = [];
  var meLower = String(email).toLowerCase();
  try {
    var qSheet = getSheet(CONFIG.SHEETS.WAGGLES_QUESTIONS, CONFIG.WAGGLES_QUESTION_COLUMNS);
    var qLast = qSheet.getLastRow();
    if (qLast >= 2) {
      var qData = qSheet.getRange(2, 1, qLast - 1, CONFIG.WAGGLES_QUESTION_COLUMNS.length).getValues();
      for (var i = 0; i < qData.length; i++) {
        if (!qData[i][0]) continue;
        var q = rowToObject(qData[i], CONFIG.WAGGLES_QUESTION_COLUMNS);
        if (q.status === 'deleted') continue;
        if (String(q.authorEmail || '').toLowerCase() !== meLower) continue;
        var key = _contribDateKey_(q.createdAt);
        if (!_contribInRange_(key, startKey, endKey)) continue;
        events.push({
          date: key,
          source: 'waggles_question',
          points: ContributionWeights.WAGGLES_QUESTION,
          label: 'Asked: ' + (q.title || q.id),
          questionId: q.id
        });
      }
    }
  } catch (e) {
    console.error('_contribCollectWaggles_ questions failed:', e);
  }
  try {
    var aSheet = getSheet(CONFIG.SHEETS.WAGGLES_ANSWERS, CONFIG.WAGGLES_ANSWER_COLUMNS);
    var aLast = aSheet.getLastRow();
    if (aLast >= 2) {
      var aData = aSheet.getRange(2, 1, aLast - 1, CONFIG.WAGGLES_ANSWER_COLUMNS.length).getValues();
      for (var j = 0; j < aData.length; j++) {
        if (!aData[j][0]) continue;
        var a = rowToObject(aData[j], CONFIG.WAGGLES_ANSWER_COLUMNS);
        if (String(a.authorEmail || '').toLowerCase() !== meLower) continue;
        var aKey = _contribDateKey_(a.createdAt);
        if (!_contribInRange_(aKey, startKey, endKey)) continue;
        var pts = ContributionWeights.WAGGLES_ANSWER;
        if (a.isAccepted === true || a.isAccepted === 'true') {
          pts += ContributionWeights.WAGGLES_ANSWER_ACCEPTED_BONUS;
        }
        events.push({
          date: aKey,
          source: 'waggles_answer',
          points: pts,
          label: (a.isAccepted ? 'Accepted answer' : 'Posted answer'),
          answerId: a.id,
          questionId: a.questionId
        });
      }
    }
  } catch (e) {
    console.error('_contribCollectWaggles_ answers failed:', e);
  }
  try {
    var vSheet = getSheet(CONFIG.SHEETS.WAGGLES_VOTES, CONFIG.WAGGLES_VOTE_COLUMNS);
    var vLast = vSheet.getLastRow();
    if (vLast < 2) return events;
    var vData = vSheet.getRange(2, 1, vLast - 1, CONFIG.WAGGLES_VOTE_COLUMNS.length).getValues();
    var vTargetTypeIdx = CONFIG.WAGGLES_VOTE_COLUMNS.indexOf('targetType');
    var vTargetIdIdx = CONFIG.WAGGLES_VOTE_COLUMNS.indexOf('targetId');
    var vValueIdx = CONFIG.WAGGLES_VOTE_COLUMNS.indexOf('value');
    var vCreatedIdx = CONFIG.WAGGLES_VOTE_COLUMNS.indexOf('createdAt');
    var vVoterIdx = CONFIG.WAGGLES_VOTE_COLUMNS.indexOf('voterEmail');

    var qAuthorByQId = {};
    try {
      var qSheet2 = getSheet(CONFIG.SHEETS.WAGGLES_QUESTIONS, CONFIG.WAGGLES_QUESTION_COLUMNS);
      var qLast2 = qSheet2.getLastRow();
      if (qLast2 >= 2) {
        var qData2 = qSheet2.getRange(2, 1, qLast2 - 1, CONFIG.WAGGLES_QUESTION_COLUMNS.length).getValues();
        var qIdIdx = CONFIG.WAGGLES_QUESTION_COLUMNS.indexOf('id');
        var qAuthorIdx = CONFIG.WAGGLES_QUESTION_COLUMNS.indexOf('authorEmail');
        for (var qi = 0; qi < qData2.length; qi++) {
          if (qData2[qi][qIdIdx]) qAuthorByQId[qData2[qi][qIdIdx]] = String(qData2[qi][qAuthorIdx] || '').toLowerCase();
        }
      }
    } catch (e) {}

    var aAuthorByAId = {};
    try {
      var aSheet2 = getSheet(CONFIG.SHEETS.WAGGLES_ANSWERS, CONFIG.WAGGLES_ANSWER_COLUMNS);
      var aLast2 = aSheet2.getLastRow();
      if (aLast2 >= 2) {
        var aData2 = aSheet2.getRange(2, 1, aLast2 - 1, CONFIG.WAGGLES_ANSWER_COLUMNS.length).getValues();
        var aIdIdx = CONFIG.WAGGLES_ANSWER_COLUMNS.indexOf('id');
        var aAuthorIdx = CONFIG.WAGGLES_ANSWER_COLUMNS.indexOf('authorEmail');
        for (var ai = 0; ai < aData2.length; ai++) {
          if (aData2[ai][aIdIdx]) aAuthorByAId[aData2[ai][aIdIdx]] = String(aData2[ai][aAuthorIdx] || '').toLowerCase();
        }
      }
    } catch (e) {}

    for (var k = 0; k < vData.length; k++) {
      if (!vData[k][0]) continue;
      var voter = String(vData[k][vVoterIdx] || '').toLowerCase();
      if (voter === meLower) continue;
      var targetType = vData[k][vTargetTypeIdx];
      var targetId = vData[k][vTargetIdIdx];
      var ownerEmail = '';
      if (targetType === 'question') ownerEmail = qAuthorByQId[targetId] || '';
      else if (targetType === 'answer') ownerEmail = aAuthorByAId[targetId] || '';
      if (ownerEmail !== meLower) continue;
      var voteValue = parseInt(vData[k][vValueIdx], 10) || 0;
      if (voteValue !== 1) continue;
      var voteKey = _contribDateKey_(vData[k][vCreatedIdx]);
      if (!_contribInRange_(voteKey, startKey, endKey)) continue;
      events.push({
        date: voteKey,
        source: 'waggles_vote_received',
        points: ContributionWeights.WAGGLES_VOTE_RECEIVED,
        label: 'Upvote received',
        targetType: targetType,
        targetId: targetId
      });
    }
  } catch (e) {
    console.error('_contribCollectWaggles_ votes failed:', e);
  }
  return events;
}

function _contribAggregate_(events) {
  var byDay = {};
  var totals = {};
  var counts = {};
  var capCounter = {};
  for (var i = 0; i < events.length; i++) {
    var ev = events[i];
    if (!byDay[ev.date]) byDay[ev.date] = { date: ev.date, score: 0, breakdown: {}, events: [] };
    var capKey = ev.date + '|' + ev.source;
    var dailyCap = null;
    if (ev.source === 'comment') dailyCap = ContributionWeights.COMMENT_DAILY_CAP;
    else if (ev.source === 'waggles_vote_received') dailyCap = ContributionWeights.WAGGLES_VOTE_DAILY_CAP;
    if (dailyCap !== null) {
      capCounter[capKey] = (capCounter[capKey] || 0) + 1;
      if (capCounter[capKey] > dailyCap) continue;
    }
    byDay[ev.date].score += ev.points;
    byDay[ev.date].breakdown[ev.source] = (byDay[ev.date].breakdown[ev.source] || 0) + ev.points;
    if (byDay[ev.date].events.length < 8) byDay[ev.date].events.push(ev);
    totals[ev.source] = (totals[ev.source] || 0) + ev.points;
    totals._score = (totals._score || 0) + ev.points;
    counts[ev.source] = (counts[ev.source] || 0) + 1;
    counts._events = (counts._events || 0) + 1;
  }
  return { byDay: byDay, totals: totals, counts: counts };
}

function _contribGridDays_(startKey, endKey) {
  var out = [];
  var start = new Date(startKey + 'T00:00:00');
  var end = new Date(endKey + 'T00:00:00');
  var tz = (CONFIG && CONFIG.TIMEZONE) || Session.getScriptTimeZone() || 'America/New_York';
  for (var d = new Date(start.getTime()); d.getTime() <= end.getTime(); d.setDate(d.getDate() + 1)) {
    out.push(Utilities.formatDate(d, tz, 'yyyy-MM-dd'));
  }
  return out;
}

function getMyContributionHeatmap(year) {
  try {
    var email = getCurrentUserEmailOptimized();
    if (!email) return { success: false, error: 'Not authenticated' };
    var win = _contribWindowKeys_(year);

    var events = []
      .concat(_contribCollectTasks_(email, win.startKey, win.endKey))
      .concat(_contribCollectAssigneePickups_(email, win.startKey, win.endKey))
      .concat(_contribCollectComments_(email, win.startKey, win.endKey))
      .concat(_contribCollectWaggles_(email, win.startKey, win.endKey));

    var agg = _contribAggregate_(events);
    var grid = _contribGridDays_(win.startKey, win.endKey);
    var days = grid.map(function(dateKey) {
      var entry = agg.byDay[dateKey];
      if (!entry) return { date: dateKey, score: 0, breakdown: {}, events: [] };
      return entry;
    });

    var streak = _contribComputeStreak_(grid, agg.byDay);

    return {
      success: true,
      email: email,
      windowStart: win.startKey,
      windowEnd: win.endKey,
      days: days,
      totals: agg.totals,
      counts: agg.counts,
      currentStreak: streak.current,
      longestStreak: streak.longest,
      weights: ContributionWeights
    };
  } catch (error) {
    console.error('getMyContributionHeatmap failed:', error);
    return { success: false, error: error.message };
  }
}

function _contribComputeStreak_(grid, byDay) {
  var current = 0;
  var longest = 0;
  var run = 0;
  for (var i = 0; i < grid.length; i++) {
    var d = byDay[grid[i]];
    if (d && d.score > 0) {
      run++;
      if (run > longest) longest = run;
    } else {
      run = 0;
    }
  }
  for (var j = grid.length - 1; j >= 0; j--) {
    var dd = byDay[grid[j]];
    if (dd && dd.score > 0) current++;
    else break;
  }
  return { current: current, longest: longest };
}

function getMyContributionRecentActivity(limit) {
  try {
    var email = getCurrentUserEmailOptimized();
    if (!email) return { success: false, error: 'Not authenticated' };
    var cap = Math.max(5, Math.min(50, parseInt(limit, 10) || 20));
    var win = _contribWindowKeys_(null);
    var events = []
      .concat(_contribCollectTasks_(email, win.startKey, win.endKey))
      .concat(_contribCollectWaggles_(email, win.startKey, win.endKey));

    var collapsed = [];
    var taskClosedByDayId = {};
    for (var i = 0; i < events.length; i++) {
      var ev = events[i];
      if (ev.source === 'task' && ev.taskId) {
        taskClosedByDayId[ev.date + '|' + ev.taskId] = collapsed.length;
        collapsed.push(ev);
      } else if (ev.source === 'ship' && ev.taskId) {
        var key = ev.date + '|' + ev.taskId;
        var idx = taskClosedByDayId[key];
        if (typeof idx === 'number') {
          collapsed[idx].shippedRef = ev.shippedRef;
          collapsed[idx].label = 'Shipped: ' + (ev.label || '').replace(/^Shipped:\s*/, '');
          collapsed[idx].source = 'task_shipped';
        } else {
          collapsed.push(ev);
        }
      } else {
        collapsed.push(ev);
      }
    }

    collapsed.sort(function(a, b) {
      if (a.date === b.date) return 0;
      return a.date < b.date ? 1 : -1;
    });
    return { success: true, items: collapsed.slice(0, cap) };
  } catch (error) {
    console.error('getMyContributionRecentActivity failed:', error);
    return { success: false, error: error.message };
  }
}
