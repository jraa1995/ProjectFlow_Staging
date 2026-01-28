const CapacityEngine = {
  DEFAULT_HOURS_PER_DAY: 8,
  DEFAULT_WORK_DAYS: ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'],
  OVERLOAD_THRESHOLD: 1.2,

  getUserCapacity(userEmail) {
    try {
      const sheet = getTeamCapacitySheet();
      const data = sheet.getDataRange().getValues();
      const columns = CONFIG.TEAM_CAPACITY_COLUMNS;
      const now = new Date();
      for (let i = 1; i < data.length; i++) {
        const config = rowToObject(data[i], columns);
        if (config.userId !== userEmail) continue;
        const effectiveFrom = config.effectiveFrom ? new Date(config.effectiveFrom) : new Date(0);
        const effectiveTo = config.effectiveTo ? new Date(config.effectiveTo) : new Date(9999, 11, 31);
        if (now >= effectiveFrom && now <= effectiveTo) {
          if (typeof config.workDays === 'string' && config.workDays) {
            try {
              config.workDays = JSON.parse(config.workDays);
            } catch (e) {
              config.workDays = this.DEFAULT_WORK_DAYS;
            }
          }
          if (typeof config.vacationDays === 'string' && config.vacationDays) {
            try {
              config.vacationDays = JSON.parse(config.vacationDays);
            } catch (e) {
              config.vacationDays = [];
            }
          }
          config.hoursPerDay = parseFloat(config.hoursPerDay) || this.DEFAULT_HOURS_PER_DAY;
          return config;
        }
      }
      return {
        userId: userEmail,
        hoursPerDay: this.DEFAULT_HOURS_PER_DAY,
        workDays: this.DEFAULT_WORK_DAYS,
        vacationDays: []
      };
    } catch (error) {
      return {
        userId: userEmail,
        hoursPerDay: this.DEFAULT_HOURS_PER_DAY,
        workDays: this.DEFAULT_WORK_DAYS,
        vacationDays: []
      };
    }
  },

  setUserCapacity(capacityData) {
    const sheet = getTeamCapacitySheet();
    const columns = CONFIG.TEAM_CAPACITY_COLUMNS;
    const data = sheet.getDataRange().getValues();
    const config = {
      id: capacityData.id || generateId('cap'),
      userId: capacityData.userId,
      hoursPerDay: capacityData.hoursPerDay || this.DEFAULT_HOURS_PER_DAY,
      workDays: JSON.stringify(capacityData.workDays || this.DEFAULT_WORK_DAYS),
      vacationDays: JSON.stringify(capacityData.vacationDays || []),
      effectiveFrom: capacityData.effectiveFrom || now().split('T')[0],
      effectiveTo: capacityData.effectiveTo || ''
    };
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][columns.indexOf('userId')] === capacityData.userId &&
          data[i][columns.indexOf('id')] === capacityData.id) {
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

  calculateUserWorkload(userEmail, options = {}) {
    const capacity = this.getUserCapacity(userEmail);
    const tasks = getAllTasks({ assignee: userEmail });
    let userTasks = options.projectId
      ? tasks.filter(t => t.projectId === options.projectId)
      : tasks;
    const startDate = options.startDate ? new Date(options.startDate) : new Date();
    const endDate = options.endDate ? new Date(options.endDate) : new Date(startDate.getTime() + 14 * 24 * 60 * 60 * 1000);
    const activeTasks = userTasks.filter(t => t.status !== 'Done');
    const totalEstimatedHrs = activeTasks.reduce((sum, t) => sum + (parseFloat(t.estimatedHrs) || 0), 0);
    const workDays = this.getWorkDaysInRange(startDate, endDate, capacity);
    const availableHrs = workDays * capacity.hoursPerDay;
    const utilization = availableHrs > 0 ? totalEstimatedHrs / availableHrs : 0;
    const tasksByStatus = {};
    activeTasks.forEach(task => {
      if (!tasksByStatus[task.status]) {
        tasksByStatus[task.status] = { count: 0, hours: 0 };
      }
      tasksByStatus[task.status].count++;
      tasksByStatus[task.status].hours += parseFloat(task.estimatedHrs) || 0;
    });
    return {
      userId: userEmail,
      capacity: capacity,
      period: { startDate, endDate },
      workDays: workDays,
      availableHrs: availableHrs,
      totalEstimatedHrs: totalEstimatedHrs,
      utilization: Math.round(utilization * 100),
      isOverloaded: utilization > this.OVERLOAD_THRESHOLD,
      activeTasks: activeTasks.length,
      tasksByStatus: tasksByStatus,
      tasks: activeTasks.map(t => ({
        id: t.id,
        title: t.title,
        status: t.status,
        priority: t.priority,
        estimatedHrs: t.estimatedHrs,
        dueDate: t.dueDate
      }))
    };
  },

  getWorkDaysInRange(startDate, endDate, capacity) {
    const dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    const vacationSet = new Set((capacity.vacationDays || []).map(d => new Date(d).toDateString()));
    let workDays = 0;
    const current = new Date(startDate);
    while (current <= endDate) {
      const dayName = dayNames[current.getDay()];
      const dateString = current.toDateString();
      if (capacity.workDays.includes(dayName) && !vacationSet.has(dateString)) {
        workDays++;
      }
      current.setDate(current.getDate() + 1);
    }
    return workDays;
  },

  getTeamWorkloadReport(options = {}) {
    const users = options.teamId
      ? getTeamMembers(options.teamId).map(m => getUserByEmail(m.userId)).filter(Boolean)
      : getActiveUsers();
    const workloads = users.map(user => this.calculateUserWorkload(user.email, options));
    const totalAvailable = workloads.reduce((sum, w) => sum + w.availableHrs, 0);
    const totalEstimated = workloads.reduce((sum, w) => sum + w.totalEstimatedHrs, 0);
    const totalTasks = workloads.reduce((sum, w) => sum + w.activeTasks, 0);
    const overloadedUsers = workloads.filter(w => w.isOverloaded);
    const underutilizedUsers = workloads.filter(w => w.utilization < 50);
    return {
      teamSize: users.length,
      period: {
        startDate: options.startDate || new Date(),
        endDate: options.endDate || new Date(Date.now() + 14 * 24 * 60 * 60 * 1000)
      },
      totalAvailableHrs: totalAvailable,
      totalEstimatedHrs: totalEstimated,
      teamUtilization: totalAvailable > 0 ? Math.round((totalEstimated / totalAvailable) * 100) : 0,
      totalActiveTasks: totalTasks,
      memberWorkloads: workloads.map(w => ({
        userId: w.userId,
        availableHrs: w.availableHrs,
        estimatedHrs: w.totalEstimatedHrs,
        utilization: w.utilization,
        isOverloaded: w.isOverloaded,
        taskCount: w.activeTasks
      })),
      overloadedUsers: overloadedUsers.map(w => ({
        userId: w.userId,
        utilization: w.utilization,
        overBy: Math.round((w.utilization - 100))
      })),
      underutilizedUsers: underutilizedUsers.map(w => ({
        userId: w.userId,
        utilization: w.utilization,
        availableHrs: w.availableHrs - w.totalEstimatedHrs
      }))
    };
  },

  getDailyWorkloadDistribution(userEmail, days = 14) {
    const capacity = this.getUserCapacity(userEmail);
    const tasks = getAllTasks({ assignee: userEmail })
      .filter(t => t.status !== 'Done' && t.dueDate);
    const dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    const vacationSet = new Set((capacity.vacationDays || []).map(d => new Date(d).toDateString()));
    const distribution = [];
    const startDate = new Date();
    for (let i = 0; i < days; i++) {
      const date = new Date(startDate.getTime() + i * 24 * 60 * 60 * 1000);
      const dateString = date.toISOString().split('T')[0];
      const dayName = dayNames[date.getDay()];
      const isWorkDay = capacity.workDays.includes(dayName) && !vacationSet.has(date.toDateString());
      const dueTasks = tasks.filter(t => t.dueDate && t.dueDate.startsWith(dateString));
      const dueHours = dueTasks.reduce((sum, t) => sum + (parseFloat(t.estimatedHrs) || 0), 0);
      distribution.push({
        date: dateString,
        dayName: dayName,
        isWorkDay: isWorkDay,
        availableHrs: isWorkDay ? capacity.hoursPerDay : 0,
        scheduledHrs: dueHours,
        utilization: isWorkDay && capacity.hoursPerDay > 0
          ? Math.round((dueHours / capacity.hoursPerDay) * 100)
          : 0,
        tasksDue: dueTasks.length
      });
    }
    return distribution;
  },

  findAvailableUsers(options = {}) {
    const users = getActiveUsers();
    const availability = [];
    for (const user of users) {
      const workload = this.calculateUserWorkload(user.email, {
        endDate: options.dueDate
      });
      const availableHrs = workload.availableHrs - workload.totalEstimatedHrs;
      const canFit = !options.requiredHrs || availableHrs >= options.requiredHrs;
      availability.push({
        userId: user.email,
        userName: user.name,
        availableHrs: Math.max(0, availableHrs),
        currentUtilization: workload.utilization,
        canFit: canFit,
        activeTasks: workload.activeTasks
      });
    }
    availability.sort((a, b) => b.availableHrs - a.availableHrs);
    return availability;
  },

  checkUserAvailability(userEmail, estimatedHrs, dueDate) {
    const workload = this.calculateUserWorkload(userEmail, { endDate: dueDate });
    const availableHrs = workload.availableHrs - workload.totalEstimatedHrs;
    const projectedUtilization = workload.availableHrs > 0
      ? ((workload.totalEstimatedHrs + estimatedHrs) / workload.availableHrs) * 100
      : 100;
    return {
      canTakeTask: availableHrs >= estimatedHrs,
      availableHrs: Math.max(0, availableHrs),
      currentUtilization: workload.utilization,
      projectedUtilization: Math.round(projectedUtilization),
      wouldBeOverloaded: projectedUtilization > this.OVERLOAD_THRESHOLD * 100,
      message: availableHrs >= estimatedHrs
        ? `User has ${availableHrs.toFixed(1)} hours available`
        : `User would be over capacity by ${(estimatedHrs - availableHrs).toFixed(1)} hours`
    };
  }
};
