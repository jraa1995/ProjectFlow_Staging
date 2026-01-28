class TimelineEngine {
  static parseFlexibleDate(dateValue) {
    if (!dateValue) return null;
    if (dateValue instanceof Date) {
      return isNaN(dateValue.getTime()) ? null : dateValue;
    }
    if (typeof dateValue === 'number') {
      const excelEpoch = new Date(1899, 11, 30);
      const msPerDay = 24 * 60 * 60 * 1000;
      const date = new Date(excelEpoch.getTime() + dateValue * msPerDay);
      return isNaN(date.getTime()) ? null : date;
    }
    if (typeof dateValue === 'string') {
      const trimmed = dateValue.trim();
      if (!trimmed) return null;
      const usFormat = trimmed.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (usFormat) {
        const month = parseInt(usFormat[1], 10);
        const day = parseInt(usFormat[2], 10);
        const year = parseInt(usFormat[3], 10);
        if (month >= 1 && month <= 12 && day >= 1 && day <= 31 && year >= 1900 && year <= 2100) {
          const date = new Date(year, month - 1, day);
          if (!isNaN(date.getTime())) {
            return date;
          }
        }
      }
      const isoFormat = trimmed.match(/^(\d{4})-(\d{2})-(\d{2})$/);
      if (isoFormat) {
        const year = parseInt(isoFormat[1], 10);
        const month = parseInt(isoFormat[2], 10);
        const day = parseInt(isoFormat[3], 10);
        if (month >= 1 && month <= 12 && day >= 1 && day <= 31 && year >= 1900 && year <= 2100) {
          const date = new Date(year, month - 1, day);
          if (!isNaN(date.getTime())) {
            return date;
          }
        }
      }
      const date = new Date(trimmed);
      if (!isNaN(date.getTime()) && date.getFullYear() >= 1900 && date.getFullYear() <= 2100) {
        return date;
      }
    }
    return null;
  }

  static generateProjectTimeline(projectId = null, dateRange = null) {
    try {
      const filters = projectId ? { projectId: projectId } : {};
      let tasks = getAllTasks(filters);
      if (dateRange) {
        tasks = tasks.filter(task => {
          const taskStart = this.parseFlexibleDate(task.startDate);
          const taskDue = this.parseFlexibleDate(task.dueDate);
          if (taskStart && taskDue) {
            return (taskStart <= dateRange.end && taskDue >= dateRange.start);
          } else if (taskStart) {
            return taskStart >= dateRange.start && taskStart <= dateRange.end;
          } else if (taskDue) {
            return taskDue >= dateRange.start && taskDue <= dateRange.end;
          }
          return false;
        });
      }
      const ganttTasks = this.convertTasksToGanttFormat(tasks);
      const dependencies = this.processDependencies(tasks);
      const criticalPath = this.calculateCriticalPath(ganttTasks, dependencies);
      const milestones = this.generateMilestones(tasks, projectId);
      const ganttTasksSerialized = ganttTasks.map(task => {
        try {
          return {
            id: String(task.id || ''),
            name: String(task.name || ''),
            start: task.start ? task.start.toISOString() : new Date().toISOString(),
            end: task.end ? task.end.toISOString() : new Date().toISOString(),
            progress: Number(task.progress || 0),
            dependencies: Array.isArray(task.dependencies) ? task.dependencies : [],
            assignee: String(task.assignee || ''),
            priority: String(task.priority || 'Medium'),
            type: String(task.type || 'Task'),
            status: String(task.status || 'To Do'),
            projectId: String(task.projectId || ''),
            estimatedHrs: Number(task.estimatedHrs || 0),
            actualHrs: Number(task.actualHrs || 0),
            parentId: String(task.parentId || ''),
            labels: Array.isArray(task.labels) ? task.labels.map(l => String(l)) : []
          };
        } catch (e) {
          return null;
        }
      }).filter(t => t !== null);
      const milestonesSerialized = milestones.map(milestone => {
        try {
          return {
            id: String(milestone.id || ''),
            name: String(milestone.name || ''),
            date: milestone.date ? milestone.date.toISOString() : new Date().toISOString(),
            type: String(milestone.type || ''),
            taskId: String(milestone.taskId || ''),
            projectId: String(milestone.projectId || '')
          };
        } catch (e) {
          return null;
        }
      }).filter(m => m !== null);
      const calculatedDateRange = this.calculateTimelineRange(ganttTasks);
      const result = {
        tasks: ganttTasksSerialized,
        milestones: milestonesSerialized,
        dependencies: dependencies.map(dep => ({
          id: String(dep.id || ''),
          from: String(dep.from || ''),
          to: String(dep.to || ''),
          type: String(dep.type || 'finish_to_start'),
          lag: Number(dep.lag || 0)
        })),
        criticalPath: criticalPath.map(id => String(id)),
        dateRange: {
          start: calculatedDateRange.start ? calculatedDateRange.start.toISOString() : new Date().toISOString(),
          end: calculatedDateRange.end ? calculatedDateRange.end.toISOString() : new Date().toISOString()
        },
        projectId: projectId ? String(projectId) : null
      };
      return result;
    } catch (error) {
      throw new Error('Failed to generate timeline data: ' + error.message);
    }
  }

  static convertTasksToGanttFormat(tasks) {
    const validTasks = [];
    tasks.forEach(task => {
      let startDate = this.parseFlexibleDate(task.startDate);
      let endDate = this.parseFlexibleDate(task.dueDate);
      if (!startDate && !endDate) {
        const createdDate = this.parseFlexibleDate(task.createdAt);
        if (createdDate) {
          startDate = createdDate;
          const estimatedDays = Math.max(1, Math.ceil((task.estimatedHrs || 8) / 8));
          endDate = new Date(startDate);
          endDate.setDate(endDate.getDate() + estimatedDays);
        } else {
          return;
        }
      } else if (!startDate && endDate) {
        const estimatedDays = Math.max(1, Math.ceil((task.estimatedHrs || 8) / 8));
        startDate = new Date(endDate);
        startDate.setDate(startDate.getDate() - estimatedDays);
      } else if (startDate && !endDate) {
        const estimatedDays = Math.max(1, Math.ceil((task.estimatedHrs || 8) / 8));
        endDate = new Date(startDate);
        endDate.setDate(endDate.getDate() + estimatedDays);
      }
      if (!startDate || !endDate || isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
        return;
      }
      let progress = 0;
      if (task.status === 'Done') {
        progress = 100;
      } else if (task.status === 'In Progress' || task.status === 'Review' || task.status === 'Testing') {
        if (task.estimatedHrs > 0) {
          progress = Math.min(90, (task.actualHrs / task.estimatedHrs) * 100);
        } else {
          progress = 50;
        }
      } else if (task.status === 'To Do') {
        progress = 0;
      } else {
        progress = 0;
      }
      validTasks.push({
        id: task.id,
        name: task.title,
        start: startDate,
        end: endDate,
        progress: Math.round(progress),
        dependencies: this.parseTaskDependencies(task),
        assignee: task.assignee || '',
        priority: task.priority || 'Medium',
        type: task.type || 'Task',
        status: task.status || 'To Do',
        projectId: task.projectId || '',
        estimatedHrs: task.estimatedHrs || 0,
        actualHrs: task.actualHrs || 0,
        parentId: task.parentId || '',
        labels: task.labels || []
      });
    });
    return validTasks;
  }

  static parseTaskDependencies(task) {
    if (!task.dependencies) return [];
    if (typeof task.dependencies === 'string') {
      return task.dependencies.split(',').map(dep => dep.trim()).filter(dep => dep);
    }
    if (Array.isArray(task.dependencies)) {
      return task.dependencies;
    }
    return [];
  }

  static processDependencies(tasks) {
    const dependencies = [];
    const taskIds = new Set(tasks.map(t => t.id));
    const dependencySheet = getTaskDependenciesSheet();
    const dependencyData = dependencySheet.getDataRange().getValues();
    if (dependencyData.length > 1) {
      for (let i = 1; i < dependencyData.length; i++) {
        const row = dependencyData[i];
        const dependency = rowToObject(row, CONFIG.TASK_DEPENDENCY_COLUMNS);
        if (taskIds.has(dependency.predecessorId) && taskIds.has(dependency.successorId)) {
          dependencies.push({
            id: dependency.id,
            from: dependency.predecessorId,
            to: dependency.successorId,
            type: dependency.dependencyType || 'finish_to_start',
            lag: parseFloat(dependency.lag) || 0
          });
        }
      }
    }
    tasks.forEach(task => {
      const taskDeps = this.parseTaskDependencies(task);
      taskDeps.forEach(depId => {
        if (taskIds.has(depId)) {
          const exists = dependencies.some(d => d.from === depId && d.to === task.id);
          if (!exists) {
            dependencies.push({
              id: `dep_${depId}_${task.id}`,
              from: depId,
              to: task.id,
              type: 'finish_to_start',
              lag: 0
            });
          }
        }
      });
    });
    return dependencies;
  }

  static calculateCriticalPath(tasks, dependencies) {
    try {
      const taskMap = new Map();
      tasks.forEach(task => {
        taskMap.set(task.id, {
          ...task,
          duration: this.calculateTaskDuration(task),
          earliestStart: 0,
          earliestFinish: 0,
          latestStart: 0,
          latestFinish: 0,
          totalFloat: 0,
          predecessors: [],
          successors: []
        });
      });
      dependencies.forEach(dep => {
        const predecessor = taskMap.get(dep.from);
        const successor = taskMap.get(dep.to);
        if (predecessor && successor) {
          predecessor.successors.push({
            taskId: dep.to,
            type: dep.type,
            lag: dep.lag
          });
          successor.predecessors.push({
            taskId: dep.from,
            type: dep.type,
            lag: dep.lag
          });
        }
      });
      this.forwardPass(taskMap);
      this.backwardPass(taskMap);
      taskMap.forEach(task => {
        task.totalFloat = task.latestStart - task.earliestStart;
      });
      const criticalPath = [];
      taskMap.forEach(task => {
        if (Math.abs(task.totalFloat) < 0.01) {
          criticalPath.push(task.id);
        }
      });
      return criticalPath;
    } catch (error) {
      return [];
    }
  }

  static calculateTaskDuration(task) {
    if (task.start && task.end) {
      const diffTime = Math.abs(task.end - task.start);
      const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
      return Math.max(1, diffDays);
    }
    if (task.estimatedHrs > 0) {
      return Math.max(1, Math.ceil(task.estimatedHrs / 8));
    }
    return 1;
  }

  static forwardPass(taskMap) {
    const visited = new Set();
    const visiting = new Set();
    const visit = (taskId) => {
      if (visited.has(taskId)) return;
      if (visiting.has(taskId)) {
        return;
      }
      visiting.add(taskId);
      const task = taskMap.get(taskId);
      if (!task) return;
      let maxEarliestFinish = 0;
      task.predecessors.forEach(pred => {
        visit(pred.taskId);
        const predTask = taskMap.get(pred.taskId);
        if (predTask) {
          const finishTime = predTask.earliestFinish + pred.lag;
          maxEarliestFinish = Math.max(maxEarliestFinish, finishTime);
        }
      });
      task.earliestStart = maxEarliestFinish;
      task.earliestFinish = task.earliestStart + task.duration;
      visiting.delete(taskId);
      visited.add(taskId);
    };
    taskMap.forEach((task, taskId) => {
      visit(taskId);
    });
  }

  static backwardPass(taskMap) {
    let projectEndTime = 0;
    taskMap.forEach(task => {
      projectEndTime = Math.max(projectEndTime, task.earliestFinish);
    });
    const visited = new Set();
    const visiting = new Set();
    const visit = (taskId) => {
      if (visited.has(taskId)) return;
      if (visiting.has(taskId)) {
        return;
      }
      visiting.add(taskId);
      const task = taskMap.get(taskId);
      if (!task) return;
      if (task.successors.length === 0) {
        task.latestFinish = projectEndTime;
      } else {
        let minLatestStart = Infinity;
        task.successors.forEach(succ => {
          visit(succ.taskId);
          const succTask = taskMap.get(succ.taskId);
          if (succTask) {
            const startTime = succTask.latestStart - succ.lag;
            minLatestStart = Math.min(minLatestStart, startTime);
          }
        });
        task.latestFinish = minLatestStart === Infinity ? projectEndTime : minLatestStart;
      }
      task.latestStart = task.latestFinish - task.duration;
      visiting.delete(taskId);
      visited.add(taskId);
    };
    taskMap.forEach((task, taskId) => {
      visit(taskId);
    });
  }

  static generateMilestones(tasks, projectId) {
    const milestones = [];
    tasks.forEach(task => {
      if (task.status === 'Done' && task.completedAt) {
        const completedDate = this.parseFlexibleDate(task.completedAt);
        if (completedDate) {
          milestones.push({
            id: `milestone_${task.id}`,
            name: `Completed: ${task.title}`,
            date: completedDate,
            type: 'task_completion',
            taskId: task.id,
            projectId: task.projectId
          });
        }
      }
    });
    if (projectId) {
      try {
        const project = getProjectById(projectId);
        if (project && project.endDate) {
          const endDate = this.parseFlexibleDate(project.endDate);
          if (endDate) {
            milestones.push({
              id: `milestone_project_${projectId}`,
              name: `Project Deadline: ${project.name}`,
              date: endDate,
              type: 'project_deadline',
              projectId: projectId
            });
          }
        }
      } catch (error) {
      }
    }
    milestones.sort((a, b) => a.date - b.date);
    return milestones;
  }

  static calculateTimelineRange(tasks) {
    if (tasks.length === 0) {
      const now = new Date();
      return {
        start: now,
        end: new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000)
      };
    }
    let minDate = null;
    let maxDate = null;
    tasks.forEach(task => {
      if (task.start) {
        if (!minDate || task.start < minDate) {
          minDate = task.start;
        }
      }
      if (task.end) {
        if (!maxDate || task.end > maxDate) {
          maxDate = task.end;
        }
      }
    });
    if (minDate) {
      minDate = new Date(minDate.getTime() - 7 * 24 * 60 * 60 * 1000);
    }
    if (maxDate) {
      maxDate = new Date(maxDate.getTime() + 7 * 24 * 60 * 60 * 1000);
    }
    return {
      start: minDate || new Date(),
      end: maxDate || new Date()
    };
  }

  static validateDependencies(dependencies) {
    const graph = new Map();
    const taskIds = new Set();
    dependencies.forEach(dep => {
      taskIds.add(dep.from);
      taskIds.add(dep.to);
      if (!graph.has(dep.from)) {
        graph.set(dep.from, []);
      }
      graph.get(dep.from).push(dep.to);
    });
    const visited = new Set();
    const recursionStack = new Set();
    const circularDependencies = [];
    const hasCycle = (node, path = []) => {
      if (recursionStack.has(node)) {
        const cycleStart = path.indexOf(node);
        circularDependencies.push(path.slice(cycleStart).concat([node]));
        return true;
      }
      if (visited.has(node)) {
        return false;
      }
      visited.add(node);
      recursionStack.add(node);
      path.push(node);
      const neighbors = graph.get(node) || [];
      for (const neighbor of neighbors) {
        if (hasCycle(neighbor, [...path])) {
          return true;
        }
      }
      recursionStack.delete(node);
      return false;
    };
    let hasCircularDependency = false;
    taskIds.forEach(taskId => {
      if (!visited.has(taskId)) {
        if (hasCycle(taskId)) {
          hasCircularDependency = true;
        }
      }
    });
    return {
      isValid: !hasCircularDependency,
      circularDependencies: circularDependencies
    };
  }
}
