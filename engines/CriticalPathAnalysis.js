function generateCriticalPathAnalysis(projectId = null, options = {}) {
  try {
    const timelineData = TimelineEngine.generateProjectTimeline(projectId);
    if (!timelineData.tasks.length) {
      return {
        criticalPath: [],
        analysis: {
          totalTasks: 0,
          criticalTasks: 0,
          projectDuration: 0,
          completionDate: null,
          riskLevel: 'low'
        }
      };
    }
    const criticalPath = timelineData.criticalPath;
    const criticalTasks = timelineData.tasks.filter(task => criticalPath.includes(task.id));
    const completionAnalysis = calculateProjectCompletion(timelineData.tasks, criticalTasks);
    const impactAnalysis = analyzeCriticalPathImpact(criticalTasks, timelineData.tasks);
    const riskAssessment = assessProjectRisk(criticalTasks, timelineData.tasks);
    const recommendations = generateCriticalPathRecommendations(criticalTasks, impactAnalysis, riskAssessment);
    return {
      criticalPath: criticalPath,
      criticalTasks: criticalTasks.map(task => ({
        ...task,
        criticalPathPosition: criticalPath.indexOf(task.id) + 1,
        impactScore: impactAnalysis.taskImpacts[task.id] || 0
      })),
      analysis: {
        totalTasks: timelineData.tasks.length,
        criticalTasks: criticalTasks.length,
        criticalPathPercentage: ((criticalTasks.length / timelineData.tasks.length) * 100).toFixed(1),
        projectDuration: completionAnalysis.totalDuration,
        earliestCompletion: completionAnalysis.earliestCompletion,
        latestCompletion: completionAnalysis.latestCompletion,
        currentCompletion: completionAnalysis.currentCompletion,
        completionConfidence: completionAnalysis.confidence,
        riskLevel: riskAssessment.overallRisk,
        riskFactors: riskAssessment.riskFactors,
        bottlenecks: impactAnalysis.bottlenecks,
        recommendations: recommendations
      },
      scenarios: generateCompletionScenarios(criticalTasks, completionAnalysis),
      dependencies: timelineData.dependencies.filter(dep =>
        criticalPath.includes(dep.from) || criticalPath.includes(dep.to)
      )
    };
  } catch (error) {
    throw new Error('Failed to generate critical path analysis: ' + error.message);
  }
}

function calculateProjectCompletion(allTasks, criticalTasks) {
  const now = new Date();
  let earliestCompletion = null;
  let totalDuration = 0;
  if (criticalTasks.length > 0) {
    const sortedCriticalTasks = [...criticalTasks].sort((a, b) => a.start - b.start);
    const firstTask = sortedCriticalTasks[0];
    const lastTask = sortedCriticalTasks[sortedCriticalTasks.length - 1];
    earliestCompletion = lastTask.end;
    totalDuration = Math.ceil((lastTask.end - firstTask.start) / (1000 * 60 * 60 * 24));
  }
  let latestCompletion = null;
  if (allTasks.length > 0) {
    latestCompletion = new Date(Math.max(...allTasks.map(t => t.end.getTime())));
  }
  const totalProgress = criticalTasks.reduce((sum, task) => sum + task.progress, 0);
  const averageProgress = criticalTasks.length > 0 ? totalProgress / criticalTasks.length : 0;
  let currentCompletion = null;
  if (earliestCompletion && averageProgress > 0) {
    const remainingProgress = 100 - averageProgress;
    const progressRate = averageProgress / Math.max(1, Math.ceil((now - criticalTasks[0]?.start || now) / (1000 * 60 * 60 * 24)));
    const remainingDays = remainingProgress / Math.max(0.1, progressRate);
    currentCompletion = new Date(now);
    currentCompletion.setDate(currentCompletion.getDate() + Math.ceil(remainingDays));
  }
  let confidence = 'medium';
  if (averageProgress > 80) {
    confidence = 'high';
  } else if (averageProgress < 30) {
    confidence = 'low';
  }
  return {
    totalDuration: totalDuration,
    earliestCompletion: earliestCompletion,
    latestCompletion: latestCompletion,
    currentCompletion: currentCompletion,
    averageProgress: averageProgress,
    confidence: confidence
  };
}

function analyzeCriticalPathImpact(criticalTasks, allTasks) {
  const taskImpacts = {};
  const bottlenecks = [];
  criticalTasks.forEach(task => {
    let impactScore = 0;
    impactScore += 10;
    const duration = Math.ceil((task.end - task.start) / (1000 * 60 * 60 * 24));
    impactScore += Math.min(20, duration * 2);
    impactScore += (100 - task.progress) * 0.3;
    const now = new Date();
    if (task.end < now && task.status !== 'Done') {
      const overdueDays = Math.ceil((now - task.end) / (1000 * 60 * 60 * 24));
      impactScore += overdueDays * 5;
    }
    const dependentCount = allTasks.filter(t =>
      t.dependencies && t.dependencies.includes(task.id)
    ).length;
    impactScore += dependentCount * 3;
    taskImpacts[task.id] = Math.round(impactScore);
    if (impactScore > 30 && task.progress < 50) {
      bottlenecks.push({
        taskId: task.id,
        taskName: task.name,
        impactScore: Math.round(impactScore),
        progress: task.progress,
        reason: 'High impact task with low progress'
      });
    }
  });
  return {
    taskImpacts: taskImpacts,
    bottlenecks: bottlenecks,
    highestImpactTask: Object.keys(taskImpacts).reduce((a, b) =>
      taskImpacts[a] > taskImpacts[b] ? a : b, Object.keys(taskImpacts)[0]
    )
  };
}

function assessProjectRisk(criticalTasks, allTasks) {
  const riskFactors = [];
  let riskScore = 0;
  const now = new Date();
  const overdueCriticalTasks = criticalTasks.filter(task =>
    task.end < now && task.status !== 'Done'
  );
  if (overdueCriticalTasks.length > 0) {
    riskFactors.push({
      type: 'overdue_critical_tasks',
      severity: 'high',
      description: `${overdueCriticalTasks.length} critical path tasks are overdue`,
      tasks: overdueCriticalTasks.map(t => t.id)
    });
    riskScore += overdueCriticalTasks.length * 15;
  }
  const lowProgressTasks = criticalTasks.filter(task =>
    task.progress < 25 && task.status !== 'Done'
  );
  if (lowProgressTasks.length > 0) {
    riskFactors.push({
      type: 'low_progress_critical_tasks',
      severity: 'medium',
      description: `${lowProgressTasks.length} critical path tasks have low progress`,
      tasks: lowProgressTasks.map(t => t.id)
    });
    riskScore += lowProgressTasks.length * 8;
  }
  const unassignedTasks = criticalTasks.filter(task => !task.assignee);
  if (unassignedTasks.length > 0) {
    riskFactors.push({
      type: 'unassigned_critical_tasks',
      severity: 'medium',
      description: `${unassignedTasks.length} critical path tasks are unassigned`,
      tasks: unassignedTasks.map(t => t.id)
    });
    riskScore += unassignedTasks.length * 10;
  }
  const criticalPathPercentage = (criticalTasks.length / allTasks.length) * 100;
  if (criticalPathPercentage > 60) {
    riskFactors.push({
      type: 'long_critical_path',
      severity: 'medium',
      description: `${criticalPathPercentage.toFixed(1)}% of tasks are on the critical path`,
      value: criticalPathPercentage
    });
    riskScore += 20;
  }
  const avgTaskDuration = criticalTasks.reduce((sum, task) => {
    const duration = Math.ceil((task.end - task.start) / (1000 * 60 * 60 * 24));
    return sum + duration;
  }, 0) / criticalTasks.length;
  if (avgTaskDuration < 2) {
    riskFactors.push({
      type: 'tight_deadlines',
      severity: 'low',
      description: 'Critical path tasks have very short durations',
      value: avgTaskDuration
    });
    riskScore += 5;
  }
  let overallRisk = 'low';
  if (riskScore > 50) {
    overallRisk = 'high';
  } else if (riskScore > 25) {
    overallRisk = 'medium';
  }
  return {
    overallRisk: overallRisk,
    riskScore: riskScore,
    riskFactors: riskFactors
  };
}

function generateCriticalPathRecommendations(criticalTasks, impactAnalysis, riskAssessment) {
  const recommendations = [];
  if (impactAnalysis.bottlenecks.length > 0) {
    recommendations.push({
      type: 'bottleneck_focus',
      priority: 'high',
      title: 'Focus on Bottleneck Tasks',
      description: `Prioritize completion of ${impactAnalysis.bottlenecks.length} bottleneck tasks to prevent project delays`,
      tasks: impactAnalysis.bottlenecks.map(b => b.taskId),
      action: 'Allocate additional resources or reassign team members to these tasks'
    });
  }
  riskAssessment.riskFactors.forEach(risk => {
    switch (risk.type) {
      case 'overdue_critical_tasks':
        recommendations.push({
          type: 'overdue_recovery',
          priority: 'critical',
          title: 'Address Overdue Critical Tasks',
          description: 'Immediate action required for overdue critical path tasks',
          tasks: risk.tasks,
          action: 'Reassess task scope, add resources, or adjust project timeline'
        });
        break;
      case 'unassigned_critical_tasks':
        recommendations.push({
          type: 'assignment_needed',
          priority: 'high',
          title: 'Assign Critical Tasks',
          description: 'Critical path tasks need immediate assignment',
          tasks: risk.tasks,
          action: 'Assign qualified team members to these tasks immediately'
        });
        break;
      case 'long_critical_path':
        recommendations.push({
          type: 'parallel_execution',
          priority: 'medium',
          title: 'Consider Parallel Execution',
          description: 'Look for opportunities to parallelize critical path tasks',
          action: 'Review dependencies and identify tasks that can be executed in parallel'
        });
        break;
    }
  });
  const avgProgress = criticalTasks.reduce((sum, task) => sum + task.progress, 0) / criticalTasks.length;
  if (avgProgress < 50) {
    recommendations.push({
      type: 'progress_acceleration',
      priority: 'medium',
      title: 'Accelerate Critical Path Progress',
      description: 'Critical path progress is below 50%',
      action: 'Consider daily standups for critical path tasks and remove blockers'
    });
  }
  const highImpactTask = criticalTasks.find(t => t.id === impactAnalysis.highestImpactTask);
  if (highImpactTask && highImpactTask.progress < 75) {
    recommendations.push({
      type: 'resource_optimization',
      priority: 'medium',
      title: 'Optimize Resources for High-Impact Task',
      description: `Task "${highImpactTask.name}" has the highest project impact`,
      tasks: [highImpactTask.id],
      action: 'Ensure this task has priority access to resources and expertise'
    });
  }
  return recommendations.sort((a, b) => {
    const priorityOrder = { critical: 4, high: 3, medium: 2, low: 1 };
    return priorityOrder[b.priority] - priorityOrder[a.priority];
  });
}

function generateCompletionScenarios(criticalTasks, completionAnalysis) {
  const scenarios = [];
  scenarios.push({
    name: 'Best Case',
    description: 'All critical path tasks complete on schedule',
    completionDate: completionAnalysis.earliestCompletion,
    probability: 'low',
    assumptions: ['No delays or blockers', 'All resources available', 'No scope changes']
  });
  if (completionAnalysis.currentCompletion) {
    scenarios.push({
      name: 'Current Trajectory',
      description: 'Based on current progress rate',
      completionDate: completionAnalysis.currentCompletion,
      probability: 'medium',
      assumptions: ['Current progress rate continues', 'No major blockers', 'Team availability remains stable']
    });
  }
  if (completionAnalysis.earliestCompletion) {
    const worstCase = new Date(completionAnalysis.earliestCompletion);
    worstCase.setDate(worstCase.getDate() + Math.ceil(completionAnalysis.totalDuration * 0.2));
    scenarios.push({
      name: 'Worst Case',
      description: 'Accounting for potential delays and risks',
      completionDate: worstCase,
      probability: 'low',
      assumptions: ['Multiple task delays', 'Resource constraints', 'Scope creep or technical challenges']
    });
  }
  return scenarios;
}
