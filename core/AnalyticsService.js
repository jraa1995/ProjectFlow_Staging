function getAnalyticsData(days) {
  var dash = getAnalyticsDashboard({ days: days });
  return _legacyAnalyticsShape_(dash);
}

function getAnalyticsDashboard(opts) {
  try {
    opts = opts || {};
    var userRole = getCurrentUserRole();
    if (userRole !== 'admin' && !isManagerRole(userRole)) {
      throw new Error('Permission denied: Analytics requires admin or manager access.');
    }
    var days = parseInt(opts.days, 10);
    if (!days || days < 1) days = 30;
    var endDate = new Date();
    var startDate = new Date(endDate.getTime() - (days * 86400000));

    var ctx = _buildAnalyticsContext_(userRole, opts);
    var allTasks = ctx.tasks;
    var allProjects = ctx.projects;
    var allUsers = ctx.users;

    var unfilteredProjects = allProjects;
    var unfilteredUsers = allUsers;
    if (opts.projectId) {
      allTasks = allTasks.filter(function(t) { return t.projectId === opts.projectId; });
      allProjects = allProjects.filter(function(p) { return p.id === opts.projectId; });
    }
    if (opts.userEmail) {
      var emailLc = String(opts.userEmail).toLowerCase();
      allTasks = allTasks.filter(function(t) {
        return (t.assignee && String(t.assignee).toLowerCase() === emailLc) ||
               (t.reporter && String(t.reporter).toLowerCase() === emailLc);
      });
      allUsers = allUsers.filter(function(u) {
        return u && u.email && u.email.toLowerCase() === emailLc;
      });
    }

    var slaConfigs = _getSlaConfigsSafe_();
    var perTaskSla = _annotateTasksWithSla_(allTasks, slaConfigs);
    var metrics = _computeOverallMetrics_(allTasks, perTaskSla, startDate, endDate);
    var throughput = _computeThroughputBuckets_(allTasks, startDate, endDate);
    var leadTime = _computeLeadTimeStats_(allTasks);
    var statusDistribution = _groupCount_(allTasks, 'status', CONFIG.STATUSES);
    var priorityDistribution = _groupCount_(allTasks, 'priority', CONFIG.PRIORITIES);
    var typeDistribution = _groupCount_(allTasks, 'type', CONFIG.TYPES);
    var ageingWip = _computeAgeingWip_(allTasks, 20);
    var slaSummary = _computeSlaSummary_(allTasks, perTaskSla, allProjects, allUsers);
    var users = _buildUserPerformance_(allTasks, perTaskSla, allUsers, days);
    var projects = _buildProjectPerformance_(allTasks, perTaskSla, allProjects, allUsers);
    var recentActivity = _buildRecentActivity_(20, allUsers, ctx.allowedProjectIds, userRole);

    return {
      asOf: new Date().toISOString(),
      scope: {
        days: days,
        projectId: opts.projectId || null,
        userEmail: opts.userEmail || null,
        role: userRole,
        contract: getManagerContractFromRole(userRole) || null,
        canSeeAll: userRole === 'admin'
      },
      filters: {
        projects: unfilteredProjects.map(function(p) { return { id: p.id, name: p.name }; }),
        users: unfilteredUsers.map(function(u) { return { email: u.email, name: u.name || u.email }; })
      },
      metrics: metrics,
      sla: slaSummary,
      throughput: throughput,
      leadTime: leadTime,
      statusDistribution: statusDistribution,
      priorityDistribution: priorityDistribution,
      typeDistribution: typeDistribution,
      ageingWip: ageingWip,
      users: users,
      projects: projects,
      recentActivity: recentActivity,
      slaConfigs: slaConfigs
    };
  } catch (error) {
    console.error('getAnalyticsDashboard failed:', error);
    throw error;
  }
}

function getUserPerformanceDetail(email, days) {
  try {
    var userRole = getCurrentUserRole();
    if (userRole !== 'admin' && !isManagerRole(userRole)) {
      throw new Error('Permission denied: Analytics requires admin or manager access.');
    }
    if (!email) throw new Error('email is required');
    days = parseInt(days, 10) || 30;
    var endDate = new Date();
    var startDate = new Date(endDate.getTime() - (days * 86400000));
    var emailLc = String(email).toLowerCase();

    var ctx = _buildAnalyticsContext_(userRole, {});
    var allTasks = ctx.tasks;
    var slaConfigs = _getSlaConfigsSafe_();
    var perTaskSla = _annotateTasksWithSla_(allTasks, slaConfigs);
    var userTasks = allTasks.filter(function(t) {
      return (t.assignee && String(t.assignee).toLowerCase() === emailLc) ||
             (t.reporter && String(t.reporter).toLowerCase() === emailLc);
    });
    var assignedTasks = userTasks.filter(function(t) {
      return t.assignee && String(t.assignee).toLowerCase() === emailLc;
    });
    var perTaskSlaForUser = {};
    Object.keys(perTaskSla).forEach(function(id) {
      perTaskSlaForUser[id] = perTaskSla[id];
    });

    var profile = ctx.users.find(function(u) { return u.email && u.email.toLowerCase() === emailLc; });
    if (!profile) {
      var allUsersFallback = getAllUsers();
      profile = allUsersFallback.find(function(u) { return u.email && u.email.toLowerCase() === emailLc; }) || { email: email };
    }

    var metrics = _computeOverallMetrics_(assignedTasks, perTaskSla, startDate, endDate);
    var throughput = _computeThroughputBuckets_(assignedTasks, startDate, endDate);
    var leadTime = _computeLeadTimeStats_(assignedTasks);
    var statusDistribution = _groupCount_(assignedTasks, 'status', CONFIG.STATUSES);
    var priorityDistribution = _groupCount_(assignedTasks, 'priority', CONFIG.PRIORITIES);
    var typeDistribution = _groupCount_(assignedTasks, 'type', CONFIG.TYPES);
    var slaSummary = _computeSlaSummary_(assignedTasks, perTaskSla, ctx.projects, ctx.users);
    var openTasks = assignedTasks
      .filter(function(t) { return t.status !== 'Done'; })
      .map(function(t) {
        var s = perTaskSla[t.id];
        return {
          id: t.id,
          title: t.title,
          status: t.status,
          priority: t.priority,
          projectId: t.projectId,
          dueDate: t.dueDate,
          slaStatus: s ? s.status : null,
          ageDays: _ageDays_(t)
        };
      })
      .sort(function(a, b) { return (b.ageDays || 0) - (a.ageDays || 0); })
      .slice(0, 50);

    var projectBreakdown = _buildProjectPerformance_(assignedTasks, perTaskSla, ctx.projects, ctx.users)
      .filter(function(p) { return p.totalTasks > 0; })
      .slice(0, 25);
    var serverMetrics = null;
    try { serverMetrics = UserMetricsService.getUserMetrics(emailLc) || null; } catch (e) {}

    return {
      asOf: new Date().toISOString(),
      user: {
        email: profile.email,
        name: profile.name || profile.email,
        role: profile.role || null,
        active: profile.active !== false
      },
      scope: { days: days, role: userRole, contract: getManagerContractFromRole(userRole) || null },
      metrics: metrics,
      sla: slaSummary,
      throughput: throughput,
      leadTime: leadTime,
      statusDistribution: statusDistribution,
      priorityDistribution: priorityDistribution,
      typeDistribution: typeDistribution,
      openTasks: openTasks,
      projects: projectBreakdown,
      serverMetrics: serverMetrics,
      tasksReportedCount: userTasks.filter(function(t) { return t.reporter && t.reporter.toLowerCase() === emailLc; }).length
    };
  } catch (error) {
    console.error('getUserPerformanceDetail failed:', error);
    throw error;
  }
}

function getProjectPerformanceDetail(projectId, days) {
  try {
    var userRole = getCurrentUserRole();
    if (userRole !== 'admin' && !isManagerRole(userRole)) {
      throw new Error('Permission denied: Analytics requires admin or manager access.');
    }
    if (!projectId) throw new Error('projectId is required');
    days = parseInt(days, 10) || 30;
    var endDate = new Date();
    var startDate = new Date(endDate.getTime() - (days * 86400000));

    var ctx = _buildAnalyticsContext_(userRole, {});
    if (ctx.allowedProjectIds && !ctx.allowedProjectIds[projectId]) {
      throw new Error('Permission denied: project not in your scope.');
    }
    var project = ctx.projects.find(function(p) { return p.id === projectId; });
    if (!project) throw new Error('Project not found');

    var projectTasks = ctx.tasks.filter(function(t) { return t.projectId === projectId; });
    var slaConfigs = _getSlaConfigsSafe_();
    var perTaskSla = _annotateTasksWithSla_(projectTasks, slaConfigs);

    var metrics = _computeOverallMetrics_(projectTasks, perTaskSla, startDate, endDate);
    var throughput = _computeThroughputBuckets_(projectTasks, startDate, endDate);
    var leadTime = _computeLeadTimeStats_(projectTasks);
    var statusDistribution = _groupCount_(projectTasks, 'status', CONFIG.STATUSES);
    var priorityDistribution = _groupCount_(projectTasks, 'priority', CONFIG.PRIORITIES);
    var typeDistribution = _groupCount_(projectTasks, 'type', CONFIG.TYPES);
    var slaSummary = _computeSlaSummary_(projectTasks, perTaskSla, ctx.projects, ctx.users);
    var contributors = _buildUserPerformance_(projectTasks, perTaskSla, ctx.users, days)
      .filter(function(u) { return u.totalTasks > 0; });
    var ageingWip = _computeAgeingWip_(projectTasks, 30);

    return {
      asOf: new Date().toISOString(),
      project: {
        id: project.id,
        name: project.name,
        status: project.status,
        ownerId: project.ownerId,
        startDate: project.startDate,
        endDate: project.endDate
      },
      scope: { days: days, role: userRole, contract: getManagerContractFromRole(userRole) || null },
      metrics: metrics,
      sla: slaSummary,
      throughput: throughput,
      leadTime: leadTime,
      statusDistribution: statusDistribution,
      priorityDistribution: priorityDistribution,
      typeDistribution: typeDistribution,
      ageingWip: ageingWip,
      contributors: contributors
    };
  } catch (error) {
    console.error('getProjectPerformanceDetail failed:', error);
    throw error;
  }
}

function _buildAnalyticsContext_(userRole, opts) {
  var allTasks = _getAnalyticsTasks_();
  var allProjects;
  try {
    allProjects = getAllProjectsOptimized() || [];
  } catch (e) {
    allProjects = getAllProjects();
  }
  var allUsers;
  try {
    allUsers = getActiveUsersOptimized() || [];
  } catch (e) {
    allUsers = getActiveUsers();
  }
  var allowedProjectIds = null;
  if (typeof getManagerContractFromRole === 'function' && getManagerContractFromRole(userRole)) {
    allowedProjectIds = getManagerAccessibleProjectIds(userRole, allProjects);
    allProjects = allProjects.filter(function(p) { return p && p.id && allowedProjectIds[p.id]; });
    allTasks = allTasks.filter(function(t) { return t && t.projectId && allowedProjectIds[t.projectId]; });
  }
  return {
    tasks: allTasks,
    projects: allProjects,
    users: allUsers,
    allowedProjectIds: allowedProjectIds
  };
}

function _getAnalyticsTasks_() {
  try {
    return getAllTasks({}, { includeArchived: true, skipPermissionCheck: true });
  } catch (e) {
    console.error('_getAnalyticsTasks_ fallback:', e);
    try {
      return getAllTasksOptimized() || [];
    } catch (ee) {
      return [];
    }
  }
}

function _getSlaConfigsSafe_() {
  try {
    if (typeof SLAEngine === 'undefined') return [];
    return SLAEngine.getSlaConfigs() || [];
  } catch (e) {
    console.error('_getSlaConfigsSafe_ failed:', e);
    return [];
  }
}

function _slaForPriority_(slaConfigs, priority) {
  if (!priority || !slaConfigs || !slaConfigs.length) return null;
  for (var i = 0; i < slaConfigs.length; i++) {
    var c = slaConfigs[i];
    if (c && c.enabled && c.priority === priority) return c;
  }
  return null;
}

function _annotateTasksWithSla_(tasks, slaConfigs) {
  var out = {};
  if (!tasks || !tasks.length) return out;
  for (var i = 0; i < tasks.length; i++) {
    var t = tasks[i];
    if (!t || !t.id) continue;
    out[t.id] = _calcTaskSla_(t, slaConfigs);
  }
  return out;
}

function _calcTaskSla_(task, slaConfigs) {
  var sla = _slaForPriority_(slaConfigs, task.priority);
  if (!sla) return { hasSla: false };
  var createdAt = new Date(task.createdAt || task.updatedAt || Date.now());
  var nowDate = new Date();
  var elapsedHrs = (nowDate.getTime() - createdAt.getTime()) / 3600000;
  var wasResponded = ['Backlog', 'To Do'].indexOf(task.status) === -1;
  var isResolved = task.status === 'Done';
  var responseDeadline = new Date(createdAt.getTime() + sla.responseTime * 3600000);
  var resolutionDeadline = new Date(createdAt.getTime() + sla.resolutionTime * 3600000);
  var responseProgress = sla.responseTime > 0 ? Math.min(100, (elapsedHrs / sla.responseTime) * 100) : 100;
  var resolutionProgress = sla.resolutionTime > 0 ? Math.min(100, (elapsedHrs / sla.resolutionTime) * 100) : 100;
  var responseBreached = !wasResponded && nowDate > responseDeadline;
  var resolutionBreached = !isResolved && nowDate > resolutionDeadline;
  var resolvedOnTime = false;
  var responseOnTime = false;
  if (isResolved && task.completedAt) {
    var completedAt = new Date(task.completedAt);
    resolvedOnTime = completedAt <= resolutionDeadline;
    responseOnTime = !responseBreached || (wasResponded && completedAt <= responseDeadline);
  } else if (wasResponded) {
    responseOnTime = !responseBreached;
  }
  var status = 'on_track';
  if (responseBreached || resolutionBreached) status = 'breached';
  else if (responseProgress > 75 || resolutionProgress > 75) status = 'at_risk';
  return {
    hasSla: true,
    slaConfig: { id: sla.id, name: sla.name, priority: sla.priority, responseTime: sla.responseTime, resolutionTime: sla.resolutionTime },
    status: status,
    responseDeadline: responseDeadline.toISOString(),
    resolutionDeadline: resolutionDeadline.toISOString(),
    responseProgress: Math.round(responseProgress),
    resolutionProgress: Math.round(resolutionProgress),
    responseBreached: responseBreached,
    resolutionBreached: resolutionBreached,
    responseCompleted: wasResponded,
    resolutionCompleted: isResolved,
    resolvedOnTime: resolvedOnTime,
    responseOnTime: responseOnTime
  };
}

function _computeOverallMetrics_(tasks, perTaskSla, startDate, endDate) {
  var nowDate = new Date();
  var dueSoonCutoff = new Date(nowDate.getTime() + 3 * 86400000);
  var total = 0;
  var completed = 0;
  var inProgress = 0;
  var todo = 0;
  var review = 0;
  var testing = 0;
  var backlog = 0;
  var overdue = 0;
  var dueSoon = 0;
  var unassigned = 0;
  var slaBreached = 0;
  var slaAtRisk = 0;
  var slaOnTrack = 0;
  var slaTracked = 0;
  var leadTimes = [];
  var onTimeCompleted = 0;
  var totalCompletedWithDue = 0;
  var createdInRange = 0;
  var completedInRange = 0;

  for (var i = 0; i < tasks.length; i++) {
    var t = tasks[i];
    if (!t) continue;
    total++;
    if (!t.assignee) unassigned++;
    if (t.status === 'Done') completed++;
    else if (t.status === 'In Progress') inProgress++;
    else if (t.status === 'To Do') todo++;
    else if (t.status === 'Backlog') backlog++;
    else if (t.status === 'Review') review++;
    else if (t.status === 'Testing') testing++;
    if (t.dueDate && t.status !== 'Done') {
      var dueD = new Date(t.dueDate);
      if (dueD < nowDate) overdue++;
      else if (dueD <= dueSoonCutoff) dueSoon++;
    }
    if (t.status === 'Done' && t.dueDate && t.completedAt) {
      totalCompletedWithDue++;
      if (new Date(t.completedAt) <= new Date(t.dueDate)) onTimeCompleted++;
    }
    if (t.status === 'Done' && t.createdAt && t.completedAt) {
      var lt = (new Date(t.completedAt).getTime() - new Date(t.createdAt).getTime()) / 86400000;
      if (lt >= 0 && lt < 365) leadTimes.push(lt);
    }
    var s = perTaskSla[t.id];
    if (s && s.hasSla) {
      slaTracked++;
      if (s.status === 'breached') slaBreached++;
      else if (s.status === 'at_risk') slaAtRisk++;
      else slaOnTrack++;
    }
    var createdAt = t.createdAt ? new Date(t.createdAt) : null;
    var completedAt = t.completedAt ? new Date(t.completedAt) : null;
    if (createdAt && createdAt >= startDate && createdAt <= endDate) createdInRange++;
    if (completedAt && completedAt >= startDate && completedAt <= endDate) completedInRange++;
  }
  var avgLead = leadTimes.length ? _avg_(leadTimes) : 0;
  var medLead = leadTimes.length ? _percentile_(leadTimes, 0.5) : 0;
  var p90Lead = leadTimes.length ? _percentile_(leadTimes, 0.9) : 0;
  var weeks = Math.max(1, Math.round((endDate.getTime() - startDate.getTime()) / 86400000 / 7));
  return {
    total: total,
    completed: completed,
    inProgress: inProgress,
    todo: todo,
    backlog: backlog,
    review: review,
    testing: testing,
    overdue: overdue,
    dueSoon: dueSoon,
    unassigned: unassigned,
    slaTracked: slaTracked,
    slaBreached: slaBreached,
    slaAtRisk: slaAtRisk,
    slaOnTrack: slaOnTrack,
    slaBreachRate: slaTracked ? Math.round((slaBreached / slaTracked) * 100) : 0,
    slaComplianceRate: slaTracked ? Math.round((slaOnTrack / slaTracked) * 100) : 100,
    completionRate: total ? Math.round((completed / total) * 100) : 0,
    onTimeDeliveryRate: totalCompletedWithDue ? Math.round((onTimeCompleted / totalCompletedWithDue) * 100) : 0,
    avgLeadTimeDays: Math.round(avgLead * 10) / 10,
    medianLeadTimeDays: Math.round(medLead * 10) / 10,
    p90LeadTimeDays: Math.round(p90Lead * 10) / 10,
    throughputPerWeek: Math.round((completedInRange / weeks) * 10) / 10,
    createdInRange: createdInRange,
    completedInRange: completedInRange
  };
}

function _computeThroughputBuckets_(tasks, startDate, endDate) {
  var weekMs = 7 * 86400000;
  var startTime = startDate.getTime();
  var endTime = endDate.getTime();
  var numWeeks = Math.max(1, Math.ceil((endTime - startTime) / weekMs));
  var weeks = [];
  for (var w = 0; w < numWeeks; w++) {
    weeks.push({
      date: new Date(startTime + w * weekMs).toISOString().slice(0, 10),
      created: 0,
      completed: 0,
      breached: 0
    });
  }
  for (var i = 0; i < tasks.length; i++) {
    var t = tasks[i];
    if (!t) continue;
    if (t.createdAt) {
      var c = new Date(t.createdAt).getTime();
      var ci = Math.floor((c - startTime) / weekMs);
      if (ci >= 0 && ci < numWeeks) weeks[ci].created++;
    }
    if (t.status === 'Done' && t.completedAt) {
      var d = new Date(t.completedAt).getTime();
      var di = Math.floor((d - startTime) / weekMs);
      if (di >= 0 && di < numWeeks) weeks[di].completed++;
    }
  }
  return { weeks: weeks };
}

function _computeLeadTimeStats_(tasks) {
  var byPriority = {};
  CONFIG.PRIORITIES.forEach(function(p) { byPriority[p] = []; });
  var allLeadTimes = [];
  for (var i = 0; i < tasks.length; i++) {
    var t = tasks[i];
    if (!t || t.status !== 'Done' || !t.createdAt || !t.completedAt) continue;
    var lt = (new Date(t.completedAt).getTime() - new Date(t.createdAt).getTime()) / 86400000;
    if (lt < 0 || lt > 365) continue;
    allLeadTimes.push(lt);
    if (t.priority && byPriority[t.priority]) byPriority[t.priority].push(lt);
  }
  var summary = {
    sampleSize: allLeadTimes.length,
    avgDays: allLeadTimes.length ? Math.round(_avg_(allLeadTimes) * 10) / 10 : 0,
    medianDays: allLeadTimes.length ? Math.round(_percentile_(allLeadTimes, 0.5) * 10) / 10 : 0,
    p90Days: allLeadTimes.length ? Math.round(_percentile_(allLeadTimes, 0.9) * 10) / 10 : 0,
    byPriority: {}
  };
  Object.keys(byPriority).forEach(function(p) {
    var arr = byPriority[p];
    summary.byPriority[p] = {
      sampleSize: arr.length,
      avgDays: arr.length ? Math.round(_avg_(arr) * 10) / 10 : 0,
      medianDays: arr.length ? Math.round(_percentile_(arr, 0.5) * 10) / 10 : 0,
      p90Days: arr.length ? Math.round(_percentile_(arr, 0.9) * 10) / 10 : 0
    };
  });
  return summary;
}

function _computeSlaSummary_(tasks, perTaskSla, projects, users) {
  var byPriority = {};
  CONFIG.PRIORITIES.forEach(function(p) {
    byPriority[p] = { total: 0, breached: 0, atRisk: 0, onTrack: 0, responseMet: 0, responseBreach: 0, resolutionMet: 0, resolutionBreach: 0 };
  });
  var byUser = {};
  var byProject = {};
  var breachList = [];
  var atRiskList = [];
  var totals = { total: 0, tracked: 0, breached: 0, atRisk: 0, onTrack: 0, responseMet: 0, responseBreach: 0, resolutionMet: 0, resolutionBreach: 0 };

  var userMap = {};
  users.forEach(function(u) { if (u && u.email) userMap[u.email.toLowerCase()] = u; });
  var projectMap = {};
  projects.forEach(function(p) { if (p && p.id) projectMap[p.id] = p; });

  for (var i = 0; i < tasks.length; i++) {
    var t = tasks[i];
    if (!t) continue;
    totals.total++;
    var s = perTaskSla[t.id];
    if (!s || !s.hasSla) continue;
    totals.tracked++;
    var prio = t.priority || 'Medium';
    if (!byPriority[prio]) byPriority[prio] = { total: 0, breached: 0, atRisk: 0, onTrack: 0, responseMet: 0, responseBreach: 0, resolutionMet: 0, resolutionBreach: 0 };
    byPriority[prio].total++;
    if (s.status === 'breached') { byPriority[prio].breached++; totals.breached++; }
    else if (s.status === 'at_risk') { byPriority[prio].atRisk++; totals.atRisk++; }
    else { byPriority[prio].onTrack++; totals.onTrack++; }
    if (s.responseCompleted) {
      if (s.responseBreached) { byPriority[prio].responseBreach++; totals.responseBreach++; }
      else { byPriority[prio].responseMet++; totals.responseMet++; }
    }
    if (s.resolutionCompleted) {
      if (s.resolutionBreached) { byPriority[prio].resolutionBreach++; totals.resolutionBreach++; }
      else { byPriority[prio].resolutionMet++; totals.resolutionMet++; }
    }
    var assigneeKey = t.assignee ? String(t.assignee).toLowerCase() : '__unassigned__';
    if (!byUser[assigneeKey]) byUser[assigneeKey] = { email: assigneeKey, name: assigneeKey === '__unassigned__' ? 'Unassigned' : (userMap[assigneeKey] && userMap[assigneeKey].name) || assigneeKey, total: 0, breached: 0, atRisk: 0, onTrack: 0 };
    byUser[assigneeKey].total++;
    if (s.status === 'breached') byUser[assigneeKey].breached++;
    else if (s.status === 'at_risk') byUser[assigneeKey].atRisk++;
    else byUser[assigneeKey].onTrack++;

    var pid = t.projectId || '__none__';
    if (!byProject[pid]) byProject[pid] = { id: pid, name: pid === '__none__' ? '(no project)' : ((projectMap[pid] && projectMap[pid].name) || pid), total: 0, breached: 0, atRisk: 0, onTrack: 0 };
    byProject[pid].total++;
    if (s.status === 'breached') byProject[pid].breached++;
    else if (s.status === 'at_risk') byProject[pid].atRisk++;
    else byProject[pid].onTrack++;

    if (s.status === 'breached' && t.status !== 'Done') {
      breachList.push({
        taskId: t.id,
        title: t.title,
        priority: t.priority,
        status: t.status,
        assignee: t.assignee || '',
        assigneeName: (userMap[assigneeKey] && userMap[assigneeKey].name) || t.assignee || '',
        projectId: t.projectId || '',
        projectName: (projectMap[t.projectId] && projectMap[t.projectId].name) || '',
        ageDays: _ageDays_(t),
        responseBreached: s.responseBreached,
        resolutionBreached: s.resolutionBreached,
        resolutionDeadline: s.resolutionDeadline
      });
    } else if (s.status === 'at_risk' && t.status !== 'Done') {
      atRiskList.push({
        taskId: t.id,
        title: t.title,
        priority: t.priority,
        status: t.status,
        assignee: t.assignee || '',
        assigneeName: (userMap[assigneeKey] && userMap[assigneeKey].name) || t.assignee || '',
        projectId: t.projectId || '',
        projectName: (projectMap[t.projectId] && projectMap[t.projectId].name) || '',
        responseProgress: s.responseProgress,
        resolutionProgress: s.resolutionProgress,
        resolutionDeadline: s.resolutionDeadline
      });
    }
  }
  Object.keys(byPriority).forEach(function(p) {
    var b = byPriority[p];
    b.complianceRate = b.total ? Math.round((b.onTrack / b.total) * 100) : 100;
    b.breachRate = b.total ? Math.round((b.breached / b.total) * 100) : 0;
    var resTotal = b.resolutionMet + b.resolutionBreach;
    b.resolutionRate = resTotal ? Math.round((b.resolutionMet / resTotal) * 100) : 100;
    var respTotal = b.responseMet + b.responseBreach;
    b.responseRate = respTotal ? Math.round((b.responseMet / respTotal) * 100) : 100;
  });
  var byUserList = Object.keys(byUser).map(function(k) {
    var u = byUser[k];
    u.complianceRate = u.total ? Math.round((u.onTrack / u.total) * 100) : 100;
    u.breachRate = u.total ? Math.round((u.breached / u.total) * 100) : 0;
    return u;
  }).sort(function(a, b) { return b.breached - a.breached || b.total - a.total; });
  var byProjectList = Object.keys(byProject).map(function(k) {
    var p = byProject[k];
    p.complianceRate = p.total ? Math.round((p.onTrack / p.total) * 100) : 100;
    p.breachRate = p.total ? Math.round((p.breached / p.total) * 100) : 0;
    return p;
  }).sort(function(a, b) { return b.breached - a.breached || b.total - a.total; });
  breachList.sort(function(a, b) { return (b.ageDays || 0) - (a.ageDays || 0); });
  atRiskList.sort(function(a, b) { return (b.resolutionProgress || 0) - (a.resolutionProgress || 0); });

  var responseTotal = totals.responseMet + totals.responseBreach;
  var resolutionTotal = totals.resolutionMet + totals.resolutionBreach;
  return {
    overall: {
      total: totals.total,
      tracked: totals.tracked,
      breached: totals.breached,
      atRisk: totals.atRisk,
      onTrack: totals.onTrack,
      complianceRate: totals.tracked ? Math.round((totals.onTrack / totals.tracked) * 100) : 100,
      breachRate: totals.tracked ? Math.round((totals.breached / totals.tracked) * 100) : 0,
      responseRate: responseTotal ? Math.round((totals.responseMet / responseTotal) * 100) : 100,
      resolutionRate: resolutionTotal ? Math.round((totals.resolutionMet / resolutionTotal) * 100) : 100
    },
    byPriority: byPriority,
    byUser: byUserList.slice(0, 50),
    byProject: byProjectList.slice(0, 50),
    breachList: breachList.slice(0, 50),
    atRiskList: atRiskList.slice(0, 50)
  };
}

function _buildUserPerformance_(tasks, perTaskSla, users, days) {
  var nowDate = new Date();
  var byUser = {};
  users.forEach(function(u) {
    if (!u || !u.email) return;
    var key = u.email.toLowerCase();
    byUser[key] = {
      email: u.email,
      name: u.name || u.email,
      role: u.role || null,
      title: u.title || '',
      department: u.department || '',
      avatar: u.avatar || '',
      totalTasks: 0,
      tasksReported: 0,
      tasksCompleted: 0,
      tasksInProgress: 0,
      tasksTodo: 0,
      tasksBacklog: 0,
      tasksReview: 0,
      tasksTesting: 0,
      tasksOverdue: 0,
      tasksOnTime: 0,
      tasksCompletedWithDue: 0,
      slaTracked: 0,
      slaBreached: 0,
      slaAtRisk: 0,
      slaOnTrack: 0,
      leadTimes: [],
      lastActivityAt: null,
      ageingHours: []
    };
  });
  for (var i = 0; i < tasks.length; i++) {
    var t = tasks[i];
    if (!t) continue;
    var asgn = t.assignee ? String(t.assignee).toLowerCase() : null;
    var rep = t.reporter ? String(t.reporter).toLowerCase() : null;
    if (asgn && !byUser[asgn]) {
      byUser[asgn] = {
        email: t.assignee, name: t.assignee, role: null, title: '', department: '', avatar: '',
        totalTasks: 0, tasksReported: 0, tasksCompleted: 0, tasksInProgress: 0, tasksTodo: 0, tasksBacklog: 0, tasksReview: 0, tasksTesting: 0,
        tasksOverdue: 0, tasksOnTime: 0, tasksCompletedWithDue: 0, slaTracked: 0, slaBreached: 0, slaAtRisk: 0, slaOnTrack: 0, leadTimes: [], lastActivityAt: null, ageingHours: []
      };
    }
    if (rep && !byUser[rep]) {
      byUser[rep] = {
        email: t.reporter, name: t.reporter, role: null, title: '', department: '', avatar: '',
        totalTasks: 0, tasksReported: 0, tasksCompleted: 0, tasksInProgress: 0, tasksTodo: 0, tasksBacklog: 0, tasksReview: 0, tasksTesting: 0,
        tasksOverdue: 0, tasksOnTime: 0, tasksCompletedWithDue: 0, slaTracked: 0, slaBreached: 0, slaAtRisk: 0, slaOnTrack: 0, leadTimes: [], lastActivityAt: null, ageingHours: []
      };
    }
    if (rep && byUser[rep]) byUser[rep].tasksReported++;
    if (!asgn || !byUser[asgn]) continue;
    var u = byUser[asgn];
    u.totalTasks++;
    if (t.status === 'Done') u.tasksCompleted++;
    else if (t.status === 'In Progress') u.tasksInProgress++;
    else if (t.status === 'To Do') u.tasksTodo++;
    else if (t.status === 'Backlog') u.tasksBacklog++;
    else if (t.status === 'Review') u.tasksReview++;
    else if (t.status === 'Testing') u.tasksTesting++;
    if (t.dueDate && t.status !== 'Done' && new Date(t.dueDate) < nowDate) u.tasksOverdue++;
    if (t.status === 'Done' && t.dueDate && t.completedAt) {
      u.tasksCompletedWithDue++;
      if (new Date(t.completedAt) <= new Date(t.dueDate)) u.tasksOnTime++;
    }
    if (t.status === 'Done' && t.createdAt && t.completedAt) {
      var lt = (new Date(t.completedAt).getTime() - new Date(t.createdAt).getTime()) / 86400000;
      if (lt >= 0 && lt < 365) u.leadTimes.push(lt);
    }
    if (t.status !== 'Done') {
      var refDate = t.updatedAt ? new Date(t.updatedAt) : (t.createdAt ? new Date(t.createdAt) : null);
      if (refDate) u.ageingHours.push((nowDate.getTime() - refDate.getTime()) / 3600000);
    }
    var s = perTaskSla[t.id];
    if (s && s.hasSla) {
      u.slaTracked++;
      if (s.status === 'breached') u.slaBreached++;
      else if (s.status === 'at_risk') u.slaAtRisk++;
      else u.slaOnTrack++;
    }
    var ts = t.updatedAt ? new Date(t.updatedAt) : (t.completedAt ? new Date(t.completedAt) : null);
    if (ts && (!u.lastActivityAt || ts > new Date(u.lastActivityAt))) u.lastActivityAt = ts.toISOString();
  }
  var weeks = Math.max(1, Math.round((days || 30) / 7));
  return Object.keys(byUser).map(function(k) {
    var u = byUser[k];
    var completionRate = u.totalTasks ? Math.round((u.tasksCompleted / u.totalTasks) * 100) : 0;
    var onTimeRate = u.tasksCompletedWithDue ? Math.round((u.tasksOnTime / u.tasksCompletedWithDue) * 100) : 0;
    var slaRate = u.slaTracked ? Math.round((u.slaOnTrack / u.slaTracked) * 100) : 100;
    var avgLead = u.leadTimes.length ? Math.round(_avg_(u.leadTimes) * 10) / 10 : 0;
    var medLead = u.leadTimes.length ? Math.round(_percentile_(u.leadTimes, 0.5) * 10) / 10 : 0;
    var p90Lead = u.leadTimes.length ? Math.round(_percentile_(u.leadTimes, 0.9) * 10) / 10 : 0;
    var avgAge = u.ageingHours.length ? Math.round(_avg_(u.ageingHours) / 24 * 10) / 10 : 0;
    var throughputPerWeek = Math.round((u.tasksCompleted / weeks) * 10) / 10;
    var score = Math.round(
      (completionRate * 0.4) +
      (slaRate * 0.3) +
      (onTimeRate * 0.3)
    );
    return {
      email: u.email,
      name: u.name,
      role: u.role,
      title: u.title,
      department: u.department,
      avatar: u.avatar,
      totalTasks: u.totalTasks,
      tasksReported: u.tasksReported,
      tasksCompleted: u.tasksCompleted,
      tasksInProgress: u.tasksInProgress,
      tasksTodo: u.tasksTodo,
      tasksBacklog: u.tasksBacklog,
      tasksReview: u.tasksReview,
      tasksTesting: u.tasksTesting,
      tasksOverdue: u.tasksOverdue,
      completionRate: completionRate,
      onTimeRate: onTimeRate,
      slaRate: slaRate,
      slaTracked: u.slaTracked,
      slaBreached: u.slaBreached,
      slaAtRisk: u.slaAtRisk,
      avgLeadTimeDays: avgLead,
      medianLeadTimeDays: medLead,
      p90LeadTimeDays: p90Lead,
      avgAgeingDays: avgAge,
      throughputPerWeek: throughputPerWeek,
      lastActivityAt: u.lastActivityAt,
      score: score
    };
  }).sort(function(a, b) { return b.score - a.score || b.tasksCompleted - a.tasksCompleted; });
}

function _buildProjectPerformance_(tasks, perTaskSla, projects, users) {
  var nowDate = new Date();
  var byProject = {};
  projects.forEach(function(p) {
    if (!p || !p.id) return;
    byProject[p.id] = {
      id: p.id,
      name: p.name,
      status: p.status,
      ownerId: p.ownerId,
      startDate: p.startDate,
      endDate: p.endDate,
      totalTasks: 0,
      tasksCompleted: 0,
      tasksInProgress: 0,
      tasksOverdue: 0,
      tasksOnTime: 0,
      tasksCompletedWithDue: 0,
      slaTracked: 0,
      slaBreached: 0,
      slaAtRisk: 0,
      slaOnTrack: 0,
      contributors: {},
      leadTimes: []
    };
  });
  for (var i = 0; i < tasks.length; i++) {
    var t = tasks[i];
    if (!t || !t.projectId) continue;
    var p = byProject[t.projectId];
    if (!p) continue;
    p.totalTasks++;
    if (t.status === 'Done') p.tasksCompleted++;
    else if (t.status === 'In Progress') p.tasksInProgress++;
    if (t.dueDate && t.status !== 'Done' && new Date(t.dueDate) < nowDate) p.tasksOverdue++;
    if (t.status === 'Done' && t.dueDate && t.completedAt) {
      p.tasksCompletedWithDue++;
      if (new Date(t.completedAt) <= new Date(t.dueDate)) p.tasksOnTime++;
    }
    if (t.status === 'Done' && t.createdAt && t.completedAt) {
      var lt = (new Date(t.completedAt).getTime() - new Date(t.createdAt).getTime()) / 86400000;
      if (lt >= 0 && lt < 365) p.leadTimes.push(lt);
    }
    var s = perTaskSla[t.id];
    if (s && s.hasSla) {
      p.slaTracked++;
      if (s.status === 'breached') p.slaBreached++;
      else if (s.status === 'at_risk') p.slaAtRisk++;
      else p.slaOnTrack++;
    }
    if (t.assignee) {
      var key = String(t.assignee).toLowerCase();
      if (!p.contributors[key]) p.contributors[key] = { email: t.assignee, count: 0, completed: 0 };
      p.contributors[key].count++;
      if (t.status === 'Done') p.contributors[key].completed++;
    }
  }
  var userNameByEmail = {};
  users.forEach(function(u) { if (u && u.email) userNameByEmail[u.email.toLowerCase()] = u.name || u.email; });
  return Object.keys(byProject).map(function(id) {
    var p = byProject[id];
    var completionRate = p.totalTasks ? Math.round((p.tasksCompleted / p.totalTasks) * 100) : 0;
    var onTimeRate = p.tasksCompletedWithDue ? Math.round((p.tasksOnTime / p.tasksCompletedWithDue) * 100) : 0;
    var slaRate = p.slaTracked ? Math.round((p.slaOnTrack / p.slaTracked) * 100) : 100;
    var avgLead = p.leadTimes.length ? Math.round(_avg_(p.leadTimes) * 10) / 10 : 0;
    var top = Object.keys(p.contributors).map(function(k) {
      var c = p.contributors[k];
      return { email: c.email, name: userNameByEmail[k] || c.email, count: c.count, completed: c.completed };
    }).sort(function(a, b) { return b.count - a.count; }).slice(0, 5);
    var health = _scoreHealth_(completionRate, slaRate, onTimeRate);
    return {
      id: p.id,
      name: p.name,
      status: p.status,
      ownerId: p.ownerId,
      ownerName: userNameByEmail[(p.ownerId || '').toLowerCase()] || p.ownerId || '',
      startDate: p.startDate,
      endDate: p.endDate,
      totalTasks: p.totalTasks,
      tasksCompleted: p.tasksCompleted,
      tasksInProgress: p.tasksInProgress,
      tasksOverdue: p.tasksOverdue,
      completionRate: completionRate,
      progress: completionRate,
      onTimeRate: onTimeRate,
      slaRate: slaRate,
      slaTracked: p.slaTracked,
      slaBreached: p.slaBreached,
      slaAtRisk: p.slaAtRisk,
      avgLeadTimeDays: avgLead,
      health: health,
      topContributors: top
    };
  }).sort(function(a, b) { return b.totalTasks - a.totalTasks; });
}

function _scoreHealth_(completionRate, slaRate, onTimeRate) {
  if (slaRate >= 90 && completionRate >= 70) return 'excellent';
  if (slaRate >= 75 && (completionRate >= 50 || onTimeRate >= 70)) return 'good';
  if (slaRate >= 50 || completionRate >= 30) return 'at-risk';
  return 'poor';
}

function _computeAgeingWip_(tasks, limit) {
  var nowDate = new Date();
  var rows = [];
  for (var i = 0; i < tasks.length; i++) {
    var t = tasks[i];
    if (!t || t.status === 'Done') continue;
    var ref = t.updatedAt ? new Date(t.updatedAt) : (t.createdAt ? new Date(t.createdAt) : null);
    if (!ref) continue;
    var ageDays = (nowDate.getTime() - ref.getTime()) / 86400000;
    rows.push({
      taskId: t.id,
      title: t.title,
      status: t.status,
      priority: t.priority,
      assignee: t.assignee || '',
      projectId: t.projectId || '',
      ageDays: Math.round(ageDays * 10) / 10
    });
  }
  rows.sort(function(a, b) { return b.ageDays - a.ageDays; });
  return rows.slice(0, limit || 20);
}

function _buildRecentActivity_(limit, users, allowedProjectIds, userRole) {
  try {
    var raw = getRecentActivity(limit * 3);
    var nameByEmail = {};
    users.forEach(function(u) { if (u && u.email) nameByEmail[u.email.toLowerCase()] = u.name || u.email; });
    var allTasks = null;
    if (allowedProjectIds) {
      try { allTasks = _getAnalyticsTasks_(); } catch (e) {}
    }
    var taskProjectMap = {};
    if (allTasks) {
      allTasks.forEach(function(t) { if (t && t.id) taskProjectMap[t.id] = t.projectId; });
    }
    var out = [];
    for (var i = 0; i < raw.length && out.length < limit; i++) {
      var a = raw[i];
      if (!a) continue;
      if (allowedProjectIds && a.entityType === 'task' && a.entityId) {
        var pid = taskProjectMap[a.entityId];
        if (pid && !allowedProjectIds[pid]) continue;
      }
      out.push({
        id: a.id,
        type: a.action || 'update',
        action: a.action,
        entityType: a.entityType,
        entityId: a.entityId,
        description: a.description || ((a.action || 'updated') + ' ' + (a.entityType || '') + ' ' + (a.entityId || '')),
        userId: a.userId,
        user: nameByEmail[(a.userId || '').toLowerCase()] || (a.userId ? String(a.userId).split('@')[0] : a.userId),
        timestamp: a.timestamp || a.createdAt
      });
    }
    return out;
  } catch (error) {
    console.error('_buildRecentActivity_ failed:', error);
    return [];
  }
}

function _groupCount_(tasks, field, ordering) {
  var counts = {};
  if (Array.isArray(ordering)) ordering.forEach(function(k) { counts[k] = 0; });
  for (var i = 0; i < tasks.length; i++) {
    var v = tasks[i] && tasks[i][field];
    if (!v) continue;
    counts[v] = (counts[v] || 0) + 1;
  }
  var keys = Object.keys(counts);
  if (Array.isArray(ordering)) {
    keys.sort(function(a, b) {
      var ai = ordering.indexOf(a);
      var bi = ordering.indexOf(b);
      if (ai === -1 && bi === -1) return a.localeCompare(b);
      if (ai === -1) return 1;
      if (bi === -1) return -1;
      return ai - bi;
    });
  }
  return keys.map(function(k) {
    var label = field === 'priority' ? 'priority' : (field === 'status' ? 'status' : 'type');
    var obj = { count: counts[k] };
    obj[label] = k;
    return obj;
  });
}

function _ageDays_(task) {
  if (!task) return 0;
  var ref = task.updatedAt ? new Date(task.updatedAt) : (task.createdAt ? new Date(task.createdAt) : null);
  if (!ref) return 0;
  return Math.round(((Date.now() - ref.getTime()) / 86400000) * 10) / 10;
}

function _avg_(arr) {
  if (!arr || !arr.length) return 0;
  var sum = 0;
  for (var i = 0; i < arr.length; i++) sum += arr[i];
  return sum / arr.length;
}

function _percentile_(arr, p) {
  if (!arr || !arr.length) return 0;
  var sorted = arr.slice().sort(function(a, b) { return a - b; });
  var idx = Math.min(sorted.length - 1, Math.floor(p * (sorted.length - 1)));
  return sorted[idx];
}

function _legacyAnalyticsShape_(dash) {
  if (!dash) return null;
  var teamProductivity = (dash.users || []).map(function(u) {
    return {
      name: u.name,
      email: u.email,
      totalTasks: u.totalTasks,
      tasksCompleted: u.tasksCompleted,
      productivity: u.completionRate
    };
  });
  var projectHealth = (dash.projects || []).map(function(p) {
    return {
      name: p.name,
      progress: p.progress,
      health: p.health,
      totalTasks: p.totalTasks,
      tasksCompleted: p.tasksCompleted
    };
  });
  return {
    metrics: {
      total: dash.metrics.total,
      completed: dash.metrics.completed,
      inProgress: dash.metrics.inProgress,
      overdue: dash.metrics.overdue
    },
    teamProductivity: teamProductivity,
    projectHealth: projectHealth,
    recentActivity: dash.recentActivity || [],
    completionTrend: (dash.throughput && dash.throughput.weeks ? dash.throughput.weeks : []).map(function(w) {
      return { date: w.date, completed: w.completed, total: w.created };
    }),
    priorityDistribution: dash.priorityDistribution || []
  };
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
