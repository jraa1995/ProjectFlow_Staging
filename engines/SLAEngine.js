const SLAEngine = {
  getSlaConfigs() {
    try {
      const sheet = getSlaConfigSheet();
      const data = sheet.getDataRange().getValues();
      const columns = CONFIG.SLA_CONFIG_COLUMNS;
      const configs = [];
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (!row[0]) continue;
        const config = rowToObject(row, columns);
        config.enabled = config.enabled === true || config.enabled === 'true';
        config.responseTime = parseFloat(config.responseTime) || 0;
        config.resolutionTime = parseFloat(config.resolutionTime) || 0;
        if (typeof config.escalationRules === 'string' && config.escalationRules) {
          try {
            config.escalationRules = JSON.parse(config.escalationRules);
          } catch (e) {
            config.escalationRules = [];
          }
        } else {
          config.escalationRules = [];
        }
        configs.push(config);
      }
      return configs;
    } catch (error) {
      return this.getDefaultSlaConfigs();
    }
  },

  getDefaultSlaConfigs() {
    return [
      {
        id: 'sla_critical',
        name: 'Critical Priority SLA',
        priority: 'Critical',
        responseTime: 1,
        resolutionTime: 4,
        enabled: true,
        escalationRules: [
          { threshold: 50, action: 'notify_assignee' },
          { threshold: 75, action: 'notify_manager' },
          { threshold: 100, action: 'escalate_priority' }
        ]
      },
      {
        id: 'sla_highest',
        name: 'Highest Priority SLA',
        priority: 'Highest',
        responseTime: 4,
        resolutionTime: 24,
        enabled: true,
        escalationRules: [
          { threshold: 75, action: 'notify_assignee' },
          { threshold: 100, action: 'notify_manager' }
        ]
      },
      {
        id: 'sla_high',
        name: 'High Priority SLA',
        priority: 'High',
        responseTime: 8,
        resolutionTime: 48,
        enabled: true,
        escalationRules: [
          { threshold: 75, action: 'notify_assignee' }
        ]
      },
      {
        id: 'sla_medium',
        name: 'Medium Priority SLA',
        priority: 'Medium',
        responseTime: 24,
        resolutionTime: 120,
        enabled: true,
        escalationRules: []
      },
      {
        id: 'sla_low',
        name: 'Low Priority SLA',
        priority: 'Low',
        responseTime: 48,
        resolutionTime: 240,
        enabled: false,
        escalationRules: []
      }
    ];
  },

  getSlaForPriority(priority) {
    const configs = this.getSlaConfigs();
    return configs.find(c => c.priority === priority && c.enabled) || null;
  },

  saveSlaConfig(slaData) {
    const sheet = getSlaConfigSheet();
    const columns = CONFIG.SLA_CONFIG_COLUMNS;
    const data = sheet.getDataRange().getValues();
    const config = {
      id: slaData.id || generateId('sla'),
      name: slaData.name,
      priority: slaData.priority,
      responseTime: slaData.responseTime,
      resolutionTime: slaData.resolutionTime,
      escalationRules: JSON.stringify(slaData.escalationRules || []),
      enabled: slaData.enabled !== false,
      createdAt: slaData.createdAt || now()
    };
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === slaData.id) {
        const newRow = objectToRow(config, columns);
        sheet.getRange(i + 1, 1, 1, columns.length).setValues([newRow]);
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow(objectToRow(config, columns));
    }
    return config;
  },

  calculateTaskSlaStatus(task) {
    const sla = this.getSlaForPriority(task.priority);
    if (!sla) {
      return { hasSla: false };
    }
    const createdAt = new Date(task.createdAt);
    const now = new Date();
    const elapsedHrs = (now - createdAt) / (1000 * 60 * 60);
    const wasResponded = !['Backlog', 'To Do'].includes(task.status);
    const isResolved = task.status === 'Done';
    const responseDeadline = new Date(createdAt.getTime() + sla.responseTime * 60 * 60 * 1000);
    const responseRemaining = (responseDeadline - now) / (1000 * 60 * 60);
    const responseProgress = Math.min(100, (elapsedHrs / sla.responseTime) * 100);
    const responseBreached = !wasResponded && now > responseDeadline;
    const resolutionDeadline = new Date(createdAt.getTime() + sla.resolutionTime * 60 * 60 * 1000);
    const resolutionRemaining = (resolutionDeadline - now) / (1000 * 60 * 60);
    const resolutionProgress = Math.min(100, (elapsedHrs / sla.resolutionTime) * 100);
    const resolutionBreached = !isResolved && now > resolutionDeadline;
    let status = 'on_track';
    if (responseBreached || resolutionBreached) {
      status = 'breached';
    } else if (responseProgress > 75 || resolutionProgress > 75) {
      status = 'at_risk';
    }
    return {
      hasSla: true,
      slaConfig: sla,
      status: status,
      response: {
        deadline: responseDeadline.toISOString(),
        remainingHrs: wasResponded ? null : Math.max(0, responseRemaining),
        progress: wasResponded ? 100 : responseProgress,
        breached: responseBreached,
        completed: wasResponded
      },
      resolution: {
        deadline: resolutionDeadline.toISOString(),
        remainingHrs: isResolved ? null : Math.max(0, resolutionRemaining),
        progress: isResolved ? 100 : resolutionProgress,
        breached: resolutionBreached,
        completed: isResolved
      }
    };
  },

  checkAllSlaBreaches() {
    const tasks = getAllTasks().filter(t => t.status !== 'Done');
    const results = {
      checked: 0,
      breached: [],
      atRisk: [],
      onTrack: []
    };
    for (const task of tasks) {
      const slaStatus = this.calculateTaskSlaStatus(task);
      if (!slaStatus.hasSla) continue;
      results.checked++;
      const taskSummary = {
        taskId: task.id,
        title: task.title,
        priority: task.priority,
        assignee: task.assignee,
        status: task.status,
        slaStatus: slaStatus.status,
        responseBreached: slaStatus.response.breached,
        resolutionBreached: slaStatus.resolution.breached,
        responseRemaining: slaStatus.response.remainingHrs,
        resolutionRemaining: slaStatus.resolution.remainingHrs
      };
      if (slaStatus.status === 'breached') {
        results.breached.push(taskSummary);
      } else if (slaStatus.status === 'at_risk') {
        results.atRisk.push(taskSummary);
      } else {
        results.onTrack.push(taskSummary);
      }
    }
    return results;
  },

  processEscalations() {
    const breachReport = this.checkAllSlaBreaches();
    const results = {
      notificationsSent: 0,
      escalations: []
    };
    for (const task of breachReport.breached) {
      const sla = this.getSlaForPriority(task.priority);
      if (!sla || !sla.escalationRules) continue;
      const applicableRules = sla.escalationRules.filter(r => r.threshold <= 100);
      for (const rule of applicableRules) {
        this.executeEscalation(task, rule, 'breached');
        results.escalations.push({
          taskId: task.taskId,
          action: rule.action,
          reason: 'breached'
        });
      }
    }
    for (const task of breachReport.atRisk) {
      const sla = this.getSlaForPriority(task.priority);
      if (!sla || !sla.escalationRules) continue;
      const fullTask = getTaskById(task.taskId);
      const slaStatus = this.calculateTaskSlaStatus(fullTask);
      const maxProgress = Math.max(
        slaStatus.response.progress || 0,
        slaStatus.resolution.progress || 0
      );
      const applicableRules = sla.escalationRules.filter(r => r.threshold <= maxProgress);
      for (const rule of applicableRules) {
        this.executeEscalation(task, rule, 'at_risk');
        results.escalations.push({
          taskId: task.taskId,
          action: rule.action,
          reason: 'at_risk'
        });
      }
    }
    return results;
  },

  executeEscalation(taskSummary, rule, reason) {
    try {
      switch (rule.action) {
        case 'notify_assignee':
          if (taskSummary.assignee) {
            NotificationEngine.createNotification({
              userId: taskSummary.assignee,
              type: 'deadline_approaching',
              title: `SLA ${reason === 'breached' ? 'Breached' : 'At Risk'}: ${taskSummary.title}`,
              message: `Task ${taskSummary.taskId} requires attention. SLA status: ${reason}`,
              entityType: 'task',
              entityId: taskSummary.taskId,
              channels: ['in_app', 'email']
            });
          }
          break;
        case 'notify_manager':
          const task = getTaskById(taskSummary.taskId);
          if (task?.projectId) {
            const project = getProjectById(task.projectId);
            if (project?.ownerId) {
              NotificationEngine.createNotification({
                userId: project.ownerId,
                type: 'deadline_approaching',
                title: `SLA Escalation: ${taskSummary.title}`,
                message: `Task ${taskSummary.taskId} has ${reason === 'breached' ? 'breached' : 'reached'} SLA threshold`,
                entityType: 'task',
                entityId: taskSummary.taskId,
                channels: ['in_app', 'email']
              });
            }
          }
          break;
        case 'escalate_priority':
          const priorities = CONFIG.PRIORITIES;
          const currentIndex = priorities.indexOf(taskSummary.priority);
          if (currentIndex < priorities.length - 1) {
            updateTask(taskSummary.taskId, {
              priority: priorities[currentIndex + 1]
            });
          }
          break;
        default:
      }
    } catch (error) {
    }
  },

  getSlaComplianceReport(options = {}) {
    let tasks = getAllTasks();
    if (options.projectId) {
      tasks = tasks.filter(t => t.projectId === options.projectId);
    }
    if (options.startDate) {
      const startDate = new Date(options.startDate);
      tasks = tasks.filter(t => new Date(t.createdAt) >= startDate);
    }
    if (options.endDate) {
      const endDate = new Date(options.endDate);
      tasks = tasks.filter(t => new Date(t.createdAt) <= endDate);
    }
    const report = {
      totalTasks: 0,
      tasksWithSla: 0,
      responseCompliance: { met: 0, breached: 0, pending: 0 },
      resolutionCompliance: { met: 0, breached: 0, pending: 0 },
      byPriority: {}
    };
    for (const task of tasks) {
      const slaStatus = this.calculateTaskSlaStatus(task);
      report.totalTasks++;
      if (!slaStatus.hasSla) continue;
      report.tasksWithSla++;
      if (!report.byPriority[task.priority]) {
        report.byPriority[task.priority] = {
          total: 0,
          responseMet: 0,
          responseBreach: 0,
          resolutionMet: 0,
          resolutionBreach: 0
        };
      }
      report.byPriority[task.priority].total++;
      if (slaStatus.response.completed) {
        if (slaStatus.response.breached) {
          report.responseCompliance.breached++;
          report.byPriority[task.priority].responseBreach++;
        } else {
          report.responseCompliance.met++;
          report.byPriority[task.priority].responseMet++;
        }
      } else {
        report.responseCompliance.pending++;
      }
      if (slaStatus.resolution.completed) {
        if (slaStatus.resolution.breached) {
          report.resolutionCompliance.breached++;
          report.byPriority[task.priority].resolutionBreach++;
        } else {
          report.resolutionCompliance.met++;
          report.byPriority[task.priority].resolutionMet++;
        }
      } else {
        report.resolutionCompliance.pending++;
      }
    }
    const totalResponse = report.responseCompliance.met + report.responseCompliance.breached;
    const totalResolution = report.resolutionCompliance.met + report.resolutionCompliance.breached;
    report.responseComplianceRate = totalResponse > 0
      ? Math.round((report.responseCompliance.met / totalResponse) * 100)
      : 100;
    report.resolutionComplianceRate = totalResolution > 0
      ? Math.round((report.resolutionCompliance.met / totalResolution) * 100)
      : 100;
    report.overallComplianceRate = Math.round(
      (report.responseComplianceRate + report.resolutionComplianceRate) / 2
    );
    return report;
  }
};
