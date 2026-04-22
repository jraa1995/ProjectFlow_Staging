function getAnalyticsData(days) {
  try {
    var userRole = getCurrentUserRole();
    if (userRole !== 'admin' && userRole !== 'manager') {
      throw new Error('Permission denied: Analytics requires admin or manager access.');
    }

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
    const byAssigneeCompleted = {};
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
      if (task.status === 'Done') {
        byAssigneeCompleted[assignee] = (byAssigneeCompleted[assignee] || 0) + 1;
      }

      if (task.projectId) {
        projectTaskCount[task.projectId] = (projectTaskCount[task.projectId] || 0) + 1;
        if (task.status === 'Done') {
          projectCompletedCount[task.projectId] = (projectCompletedCount[task.projectId] || 0) + 1;
        }
      }

      metrics.total++;
      if (task.status === 'Done') {
        metrics.completed++;
      } else if (task.status === 'In Progress') {
        metrics.inProgress++;
      }
      if (task.dueDate && task.status !== 'Done') {
        if (new Date(task.dueDate) < now) {
          metrics.overdue++;
        }
      }

      if (inDateRange) {
        if (task.status === 'Done') {
          const weekKey = getWeekKey(new Date(task.updatedAt || task.createdAt));
          weeklyData[weekKey] = weeklyData[weekKey] || { completed: 0, total: 0 };
          weeklyData[weekKey].completed++;
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
      const userCompleted = byAssigneeCompleted[user.email] || 0;

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

    const userNameByEmail = {};
    allUsers.forEach(u => {
      if (u && u.email) userNameByEmail[u.email] = u.name || u.email.split('@')[0];
    });
    const recentActivity = getRecentActivity(20).map(activity => ({
      type: activity.action || 'update',
      description: activity.description || `${activity.action || 'Updated'} ${activity.entityType} ${activity.entityId}`,
      user: userNameByEmail[activity.userId] || (activity.userId ? activity.userId.split('@')[0] : activity.userId),
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
