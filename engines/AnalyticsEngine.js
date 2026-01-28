class AnalyticsEngine {
  static calculateProductivityMetrics(dateRange, teamFilter) {
    const endDate = dateRange?.end || new Date();
    const startDate = dateRange?.start || new Date(endDate.getTime() - (30 * 24 * 60 * 60 * 1000));
    const allTasks = getAllTasks();
    const filteredTasks = allTasks.filter(task => {
      const taskDate = new Date(task.createdAt || task.updatedAt);
      const inDateRange = taskDate >= startDate && taskDate <= endDate;
      const matchesTeam = !teamFilter || task.assignee === teamFilter;
      return inDateRange && matchesTeam;
    });
    const totalTasks = filteredTasks.length;
    const completedTasks = filteredTasks.filter(task => task.status === 'Done').length;
    const inProgressTasks = filteredTasks.filter(task => task.status === 'In Progress').length;
    const overdueTasks = filteredTasks.filter(task => {
      return task.dueDate && new Date(task.dueDate) < new Date() && task.status !== 'Done';
    }).length;
    const completionRate = totalTasks > 0 ? (completedTasks / totalTasks) * 100 : 0;
    const completedTasksWithDates = filteredTasks.filter(task =>
      task.status === 'Done' && task.createdAt && task.updatedAt
    );
    let averageTaskTime = 0;
    if (completedTasksWithDates.length > 0) {
      const totalTime = completedTasksWithDates.reduce((sum, task) => {
        const created = new Date(task.createdAt);
        const completed = new Date(task.updatedAt);
        return sum + (completed - created);
      }, 0);
      averageTaskTime = totalTime / completedTasksWithDates.length / (1000 * 60 * 60 * 24);
    }
    const velocityTrend = this.calculateVelocityTrend(filteredTasks, startDate, endDate);
    const teamUtilization = this.calculateTeamUtilization(filteredTasks);
    return {
      total: totalTasks,
      completed: completedTasks,
      inProgress: inProgressTasks,
      overdue: overdueTasks,
      completionRate: Math.round(completionRate * 100) / 100,
      averageTaskTime: Math.round(averageTaskTime * 100) / 100,
      velocityTrend: velocityTrend,
      teamUtilization: teamUtilization,
      dateRange: {
        start: startDate.toISOString(),
        end: endDate.toISOString()
      }
    };
  }

  static calculateVelocityTrend(tasks, startDate, endDate) {
    const completedTasks = tasks.filter(task => task.status === 'Done' && task.updatedAt);
    const trend = [];
    const weekMs = 7 * 24 * 60 * 60 * 1000;
    let currentDate = new Date(startDate);
    while (currentDate < endDate) {
      const weekEnd = new Date(currentDate.getTime() + weekMs);
      const weekTasks = completedTasks.filter(task => {
        const taskDate = new Date(task.updatedAt);
        return taskDate >= currentDate && taskDate < weekEnd;
      });
      trend.push({
        date: currentDate.toISOString().split('T')[0],
        completed: weekTasks.length,
        total: tasks.filter(task => {
          const taskDate = new Date(task.createdAt || task.updatedAt);
          return taskDate >= currentDate && taskDate < weekEnd;
        }).length
      });
      currentDate = weekEnd;
    }
    return trend;
  }

  static calculateTeamUtilization(tasks) {
    const users = getActiveUsers();
    const utilization = [];
    users.forEach(user => {
      const userTasks = tasks.filter(task => task.assignee === user.email);
      const completedTasks = userTasks.filter(task => task.status === 'Done');
      const productivity = userTasks.length > 0 ?
        Math.round((completedTasks.length / userTasks.length) * 100) : 0;
      utilization.push({
        name: user.name || user.email,
        email: user.email,
        totalTasks: userTasks.length,
        tasksCompleted: completedTasks.length,
        productivity: productivity
      });
    });
    return utilization.sort((a, b) => b.productivity - a.productivity);
  }

  static generatePredictions(projectId, confidenceLevel = 0.8) {
    const tasks = projectId ? getAllTasks({ projectId }) : getAllTasks();
    const completedTasks = tasks.filter(task => task.status === 'Done');
    const remainingTasks = tasks.filter(task => task.status !== 'Done');
    if (completedTasks.length === 0) {
      return {
        estimatedCompletion: null,
        confidenceInterval: 0,
        riskFactors: ['Insufficient historical data'],
        recommendations: ['Complete more tasks to improve predictions']
      };
    }
    const completionTimes = completedTasks
      .filter(task => task.createdAt && task.updatedAt)
      .map(task => {
        const created = new Date(task.createdAt);
        const completed = new Date(task.updatedAt);
        return (completed - created) / (1000 * 60 * 60 * 24);
      });
    if (completionTimes.length === 0) {
      return {
        estimatedCompletion: null,
        confidenceInterval: 0,
        riskFactors: ['No completion time data available'],
        recommendations: ['Ensure task dates are properly tracked']
      };
    }
    const avgCompletionTime = completionTimes.reduce((sum, time) => sum + time, 0) / completionTimes.length;
    const variance = completionTimes.reduce((sum, time) => sum + Math.pow(time - avgCompletionTime, 2), 0) / completionTimes.length;
    const stdDev = Math.sqrt(variance);
    const estimatedDaysRemaining = remainingTasks.length * avgCompletionTime;
    const estimatedCompletion = new Date(Date.now() + (estimatedDaysRemaining * 24 * 60 * 60 * 1000));
    const zScore = confidenceLevel === 0.95 ? 1.96 : (confidenceLevel === 0.9 ? 1.645 : 1.28);
    const marginOfError = zScore * (stdDev / Math.sqrt(completionTimes.length));
    const riskFactors = this.identifyRiskFactors(tasks, avgCompletionTime, stdDev);
    const recommendations = this.generateRecommendations(tasks, riskFactors);
    return {
      estimatedCompletion: estimatedCompletion.toISOString(),
      confidenceInterval: Math.round(confidenceLevel * 100),
      marginOfErrorDays: Math.round(marginOfError * 100) / 100,
      riskFactors: riskFactors,
      recommendations: recommendations,
      historicalData: {
        avgCompletionTime: Math.round(avgCompletionTime * 100) / 100,
        standardDeviation: Math.round(stdDev * 100) / 100,
        sampleSize: completionTimes.length
      }
    };
  }

  static identifyBottlenecks(tasks, timeframe = 30) {
    const cutoffDate = new Date(Date.now() - (timeframe * 24 * 60 * 60 * 1000));
    const recentTasks = tasks.filter(task => {
      const taskDate = new Date(task.updatedAt || task.createdAt);
      return taskDate >= cutoffDate;
    });
    const statusCounts = {};
    const statusDurations = {};
    recentTasks.forEach(task => {
      const status = task.status;
      statusCounts[status] = (statusCounts[status] || 0) + 1;
      if (task.updatedAt) {
        const timeInStatus = (new Date() - new Date(task.updatedAt)) / (1000 * 60 * 60 * 24);
        if (!statusDurations[status]) statusDurations[status] = [];
        statusDurations[status].push(timeInStatus);
      }
    });
    const bottlenecks = [];
    Object.keys(statusCounts).forEach(status => {
      const count = statusCounts[status];
      const avgDuration = statusDurations[status] ?
        statusDurations[status].reduce((sum, d) => sum + d, 0) / statusDurations[status].length : 0;
      if (count > recentTasks.length * 0.2 && avgDuration > 7) {
        bottlenecks.push({
          status: status,
          taskCount: count,
          averageDuration: Math.round(avgDuration * 100) / 100,
          severity: avgDuration > 14 ? 'high' : 'medium'
        });
      }
    });
    const assigneeWorkload = {};
    recentTasks.forEach(task => {
      if (task.assignee) {
        assigneeWorkload[task.assignee] = (assigneeWorkload[task.assignee] || 0) + 1;
      }
    });
    const overloadedAssignees = Object.keys(assigneeWorkload)
      .filter(assignee => assigneeWorkload[assignee] > recentTasks.length * 0.3)
      .map(assignee => ({
        assignee: assignee,
        taskCount: assigneeWorkload[assignee],
        percentage: Math.round((assigneeWorkload[assignee] / recentTasks.length) * 100)
      }));
    return {
      statusBottlenecks: bottlenecks,
      overloadedAssignees: overloadedAssignees,
      totalTasksAnalyzed: recentTasks.length,
      timeframe: timeframe,
      recommendations: this.generateBottleneckRecommendations(bottlenecks, overloadedAssignees)
    };
  }

  static calculateBurndownData(sprintId, projectId) {
    let tasks = getAllTasks();
    if (projectId) {
      tasks = tasks.filter(task => task.projectId === projectId);
    }
    if (sprintId) {
      tasks = tasks.filter(task => task.sprint === sprintId);
    }
    const totalStoryPoints = tasks.reduce((sum, task) => sum + (parseInt(task.storyPoints) || 0), 0);
    const completedStoryPoints = tasks
      .filter(task => task.status === 'Done')
      .reduce((sum, task) => sum + (parseInt(task.storyPoints) || 0), 0);
    const sprintDays = 14;
    const idealBurndown = [];
    for (let day = 0; day <= sprintDays; day++) {
      idealBurndown.push({
        day: day,
        remaining: totalStoryPoints - (totalStoryPoints * day / sprintDays)
      });
    }
    const actualBurndown = [
      { day: 0, remaining: totalStoryPoints },
      { day: sprintDays, remaining: totalStoryPoints - completedStoryPoints }
    ];
    return {
      totalStoryPoints: totalStoryPoints,
      completedStoryPoints: completedStoryPoints,
      remainingStoryPoints: totalStoryPoints - completedStoryPoints,
      idealBurndown: idealBurndown,
      actualBurndown: actualBurndown,
      sprintProgress: totalStoryPoints > 0 ? (completedStoryPoints / totalStoryPoints) * 100 : 0
    };
  }

  static generateExecutiveDashboard(portfolioFilter) {
    const allTasks = getAllTasks();
    const allProjects = getAllProjects();
    let tasks = allTasks;
    let projects = allProjects;
    if (portfolioFilter) {
      projects = projects.filter(project => project.portfolio === portfolioFilter);
      const projectIds = projects.map(p => p.id);
      tasks = tasks.filter(task => projectIds.includes(task.projectId));
    }
    const totalProjects = projects.length;
    const activeProjects = projects.filter(project => project.status !== 'Completed').length;
    const totalTasks = tasks.length;
    const completedTasks = tasks.filter(task => task.status === 'Done').length;
    const projectHealth = projects.map(project => {
      const projectTasks = tasks.filter(task => task.projectId === project.id);
      const projectCompleted = projectTasks.filter(task => task.status === 'Done').length;
      const progress = projectTasks.length > 0 ? (projectCompleted / projectTasks.length) * 100 : 0;
      let health = 'good';
      if (progress < 30) health = 'poor';
      else if (progress < 60) health = 'at-risk';
      else if (progress < 90) health = 'good';
      else health = 'excellent';
      return {
        name: project.name,
        progress: Math.round(progress),
        health: health,
        totalTasks: projectTasks.length,
        tasksCompleted: projectCompleted
      };
    });
    const teamProductivity = this.calculateTeamUtilization(tasks);
    const recentActivity = getRecentActivity(20).map(activity => ({
      type: activity.action,
      description: activity.description || `${activity.action} on ${activity.entityType} ${activity.entityId}`,
      user: activity.userId,
      timestamp: activity.createdAt
    }));
    return {
      metrics: {
        totalProjects: totalProjects,
        activeProjects: activeProjects,
        total: totalTasks,
        completed: completedTasks,
        inProgress: tasks.filter(task => task.status === 'In Progress').length,
        overdue: tasks.filter(task =>
          task.dueDate && new Date(task.dueDate) < new Date() && task.status !== 'Done'
        ).length
      },
      projectHealth: projectHealth,
      teamProductivity: teamProductivity.slice(0, 10),
      recentActivity: recentActivity,
      completionTrend: this.calculateVelocityTrend(tasks,
        new Date(Date.now() - (30 * 24 * 60 * 60 * 1000)),
        new Date()
      ),
      priorityDistribution: this.calculatePriorityDistribution(tasks)
    };
  }

  static calculatePriorityDistribution(tasks) {
    const priorities = {};
    tasks.forEach(task => {
      const priority = task.priority || 'Medium';
      priorities[priority] = (priorities[priority] || 0) + 1;
    });
    return Object.keys(priorities).map(priority => ({
      priority: priority,
      count: priorities[priority]
    }));
  }

  static identifyRiskFactors(tasks, avgCompletionTime, stdDev) {
    const riskFactors = [];
    if (stdDev > avgCompletionTime * 0.5) {
      riskFactors.push('High variability in task completion times');
    }
    const overdueTasks = tasks.filter(task =>
      task.dueDate && new Date(task.dueDate) < new Date() && task.status !== 'Done'
    );
    if (overdueTasks.length > tasks.length * 0.2) {
      riskFactors.push('High number of overdue tasks');
    }
    const unassignedTasks = tasks.filter(task => !task.assignee);
    if (unassignedTasks.length > tasks.length * 0.1) {
      riskFactors.push('Many unassigned tasks');
    }
    const criticalTasks = tasks.filter(task =>
      (task.priority === 'Critical' || task.priority === 'High') &&
      task.status === 'To Do'
    );
    if (criticalTasks.length > 0) {
      riskFactors.push('High priority tasks not started');
    }
    return riskFactors;
  }

  static generateRecommendations(tasks, riskFactors) {
    const recommendations = [];
    riskFactors.forEach(risk => {
      switch (risk) {
        case 'High variability in task completion times':
          recommendations.push('Consider breaking down large tasks into smaller, more predictable units');
          break;
        case 'High number of overdue tasks':
          recommendations.push('Review and update task due dates, prioritize overdue items');
          break;
        case 'Many unassigned tasks':
          recommendations.push('Assign tasks to team members to improve accountability');
          break;
        case 'High priority tasks not started':
          recommendations.push('Prioritize high-priority tasks and ensure they are actively worked on');
          break;
      }
    });
    if (recommendations.length === 0) {
      recommendations.push('Project appears to be on track, continue current practices');
    }
    return recommendations;
  }

  static generateBottleneckRecommendations(bottlenecks, overloadedAssignees) {
    const recommendations = [];
    bottlenecks.forEach(bottleneck => {
      recommendations.push(`Address ${bottleneck.status} bottleneck: ${bottleneck.taskCount} tasks averaging ${bottleneck.averageDuration} days`);
    });
    overloadedAssignees.forEach(assignee => {
      recommendations.push(`Consider redistributing workload for ${assignee.assignee} (${assignee.percentage}% of tasks)`);
    });
    if (recommendations.length === 0) {
      recommendations.push('No significant bottlenecks detected');
    }
    return recommendations;
  }

  static detectAnomalies(tasks, timeframe = 30) {
    const cutoffDate = new Date(Date.now() - (timeframe * 24 * 60 * 60 * 1000));
    const recentTasks = tasks.filter(task => {
      const taskDate = new Date(task.updatedAt || task.createdAt);
      return taskDate >= cutoffDate;
    });
    const anomalies = [];
    const dailyCompletions = {};
    recentTasks.filter(task => task.status === 'Done').forEach(task => {
      const date = new Date(task.updatedAt).toISOString().split('T')[0];
      dailyCompletions[date] = (dailyCompletions[date] || 0) + 1;
    });
    const completionCounts = Object.values(dailyCompletions);
    if (completionCounts.length > 0) {
      const avgCompletions = completionCounts.reduce((sum, count) => sum + count, 0) / completionCounts.length;
      const maxCompletions = Math.max(...completionCounts);
      if (maxCompletions > avgCompletions * 3) {
        anomalies.push({
          type: 'unusual_completion_spike',
          description: `Unusual spike in task completions: ${maxCompletions} tasks in one day`,
          severity: 'medium'
        });
      }
    }
    const dailyCreations = {};
    recentTasks.forEach(task => {
      const date = new Date(task.createdAt).toISOString().split('T')[0];
      dailyCreations[date] = (dailyCreations[date] || 0) + 1;
    });
    const creationCounts = Object.values(dailyCreations);
    if (creationCounts.length > 0) {
      const avgCreations = creationCounts.reduce((sum, count) => sum + count, 0) / creationCounts.length;
      const maxCreations = Math.max(...creationCounts);
      if (maxCreations > avgCreations * 3) {
        anomalies.push({
          type: 'unusual_creation_spike',
          description: `Unusual spike in task creation: ${maxCreations} tasks in one day`,
          severity: 'low'
        });
      }
    }
    return {
      anomalies: anomalies,
      timeframe: timeframe,
      tasksAnalyzed: recentTasks.length
    };
  }
}
