function addDependency(successorId, predecessorId, dependencyType, lag) {
  try {
    const successor = getTaskById(successorId);
    const predecessor = getTaskById(predecessorId);

    if (!successor) {
      throw new Error(`Successor task not found: ${successorId}`);
    }

    if (!predecessor) {
      throw new Error(`Predecessor task not found: ${predecessorId}`);
    }

    if (successorId === predecessorId) {
      throw new Error('A task cannot depend on itself');
    }

    const existing = getDependencyBetween(successorId, predecessorId);
    if (existing) {
      throw new Error('Dependency already exists between these tasks');
    }

    const circular = detectCircularDependency(successorId, predecessorId);
    if (circular) {
      throw new Error(`Circular dependency detected: ${circular.join(' → ')}`);
    }

    const dependency = createTaskDependency({
      successorId: successorId,
      predecessorId: predecessorId,
      dependencyType: dependencyType || 'finish_to_start',
      lag: lag || 0
    });

    return {
      success: true,
      dependency: dependency
    };
  } catch (error) {
    console.error('addDependency error:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

function removeDependency(dependencyId) {
  try {
    const success = deleteTaskDependency(dependencyId);
    if (success) {
      return { success: true };
    } else {
      throw new Error('Dependency not found');
    }
  } catch (error) {
    console.error('removeDependency error:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

function getTaskDependenciesWithDetails(taskId) {
  try {
    const deps = getTaskDependencies(taskId);
    const taskCache = {};
    const resolveTask = (id) => {
      if (!taskCache[id]) taskCache[id] = getTaskById(id);
      return taskCache[id];
    };

    return {
      predecessors: deps.predecessors.map(dep => ({ ...dep, task: resolveTask(dep.predecessorId) })),
      successors: deps.successors.map(dep => ({ ...dep, task: resolveTask(dep.successorId) }))
    };
  } catch (error) {
    console.error('getTaskDependenciesWithDetails error:', error);
    return { predecessors: [], successors: [] };
  }
}

function detectCircularDependency(successorId, predecessorId) {
  try {
    const depMap = getAllDependenciesMap();
    const visited = new Set();
    const path = [];

    function hasCycle(currentId, targetId) {
      if (currentId === targetId) {
        path.push(currentId);
        return true;
      }
      if (visited.has(currentId)) return false;
      visited.add(currentId);
      path.push(currentId);
      const preds = depMap.bySuccessor[currentId] || [];
      for (const dep of preds) {
        if (hasCycle(dep.predecessorId, targetId)) return true;
      }
      path.pop();
      return false;
    }

    if (hasCycle(predecessorId, successorId)) {
      return path;
    }
    return null;
  } catch (error) {
    console.error('detectCircularDependency error:', error);
    return null;
  }
}

function getDependencyBetween(successorId, predecessorId) {
  try {
    const deps = getTaskDependencies(successorId);
    return deps.predecessors.find(dep => dep.predecessorId === predecessorId) || null;
  } catch (error) {
    console.error('getDependencyBetween error:', error);
    return null;
  }
}

function canStartTask(taskId) {
  try {
    const task = getTaskById(taskId);
    if (!task) throw new Error('Task not found');

    const deps = getTaskDependencies(taskId);
    if (deps.predecessors.length === 0) {
      return { canStart: true, blockingTasks: [] };
    }

    const taskMap = {};
    deps.predecessors.forEach(dep => {
      if (!taskMap[dep.predecessorId]) {
        taskMap[dep.predecessorId] = getTaskById(dep.predecessorId);
      }
    });

    const blockingTasks = [];
    for (const dep of deps.predecessors) {
      const predecessor = taskMap[dep.predecessorId];
      if (!predecessor) continue;
      if (dep.dependencyType === 'finish_to_start') {
        if (predecessor.status !== 'Done') {
          blockingTasks.push({
            taskId: predecessor.id,
            title: predecessor.title,
            status: predecessor.status,
            dependencyType: dep.dependencyType
          });
        }
      } else if (dep.dependencyType === 'start_to_start') {
        if (predecessor.status === 'Backlog' || predecessor.status === 'To Do') {
          blockingTasks.push({
            taskId: predecessor.id,
            title: predecessor.title,
            status: predecessor.status,
            dependencyType: dep.dependencyType
          });
        }
      }
    }

    return {
      canStart: blockingTasks.length === 0,
      blockingTasks: blockingTasks
    };
  } catch (error) {
    console.error('canStartTask error:', error);
    return { canStart: true, blockingTasks: [] };
  }
}

function getBlockedTasks() {
  try {
    const allTasks = RequestCache.getTasks();
    const depMap = getAllDependenciesMap();
    const taskMap = {};
    allTasks.forEach(t => { taskMap[t.id] = t; });
    const blockedTasks = [];

    for (const task of allTasks) {
      if (task.status === 'Done') continue;
      const predecessors = depMap.bySuccessor[task.id] || [];
      if (predecessors.length === 0) continue;

      const blockingTasks = [];
      for (const dep of predecessors) {
        const predecessor = taskMap[dep.predecessorId];
        if (!predecessor) continue;
        if (dep.dependencyType === 'finish_to_start' && predecessor.status !== 'Done') {
          blockingTasks.push({
            taskId: predecessor.id,
            title: predecessor.title,
            status: predecessor.status,
            dependencyType: dep.dependencyType
          });
        } else if (dep.dependencyType === 'start_to_start' &&
          (predecessor.status === 'Backlog' || predecessor.status === 'To Do')) {
          blockingTasks.push({
            taskId: predecessor.id,
            title: predecessor.title,
            status: predecessor.status,
            dependencyType: dep.dependencyType
          });
        }
      }

      if (blockingTasks.length > 0) {
        blockedTasks.push({ ...task, blockingTasks });
      }
    }

    return blockedTasks;
  } catch (error) {
    console.error('getBlockedTasks error:', error);
    return [];
  }
}

function calculateCriticalPath(projectId) {
  try {
    const allTasks = RequestCache.getTasks();
    let tasks = projectId ? allTasks.filter(t => t.projectId === projectId) : allTasks;
    tasks = tasks.filter(t => t.startDate || t.dueDate);

    if (tasks.length === 0) {
      return {
        criticalPath: [],
        totalDuration: 0,
        projectStart: null,
        projectEnd: null
      };
    }

    const depMap = getAllDependenciesMap();
    const taskMap = {};

    tasks.forEach(t => {
      taskMap[t.id] = {
        task: t,
        duration: calculateTaskDuration(t),
        earliestStart: null,
        earliestFinish: null,
        latestStart: null,
        latestFinish: null,
        slack: null,
        isCritical: false,
        predecessors: [],
        successors: []
      };
    });

    tasks.forEach(t => {
      const preds = depMap.bySuccessor[t.id] || [];
      preds.forEach(dep => {
        if (taskMap[dep.predecessorId]) {
          taskMap[t.id].predecessors.push(dep.predecessorId);
        }
      });
      const succs = depMap.byPredecessor[t.id] || [];
      succs.forEach(dep => {
        if (taskMap[dep.successorId]) {
          taskMap[t.id].successors.push(dep.successorId);
        }
      });
    });

    const forwardPass = (taskId) => {
      const node = taskMap[taskId];
      if (node.earliestFinish !== null) return;
      if (node.predecessors.length === 0) {
        node.earliestStart = node.task.startDate ? new Date(node.task.startDate) : new Date();
      } else {
        let maxFinish = new Date(0);
        node.predecessors.forEach(predId => {
          forwardPass(predId);
          const pred = taskMap[predId];
          if (pred.earliestFinish > maxFinish) maxFinish = pred.earliestFinish;
        });
        node.earliestStart = new Date(maxFinish);
      }
      node.earliestFinish = new Date(node.earliestStart.getTime() + (node.duration * 24 * 60 * 60 * 1000));
    };

    Object.keys(taskMap).forEach(taskId => forwardPass(taskId));

    let projectEnd = new Date(0);
    Object.values(taskMap).forEach(node => {
      if (node.earliestFinish > projectEnd) projectEnd = node.earliestFinish;
    });

    const backwardPass = (taskId) => {
      const node = taskMap[taskId];
      if (node.latestStart !== null) return;
      if (node.successors.length === 0) {
        node.latestFinish = new Date(projectEnd);
      } else {
        let minStart = new Date(9999, 11, 31);
        node.successors.forEach(succId => {
          backwardPass(succId);
          const succ = taskMap[succId];
          if (succ.latestStart < minStart) minStart = succ.latestStart;
        });
        node.latestFinish = new Date(minStart);
      }
      node.latestStart = new Date(node.latestFinish.getTime() - (node.duration * 24 * 60 * 60 * 1000));
    };

    Object.keys(taskMap).forEach(taskId => backwardPass(taskId));

    Object.values(taskMap).forEach(node => {
      const slackDays = (node.latestStart - node.earliestStart) / (24 * 60 * 60 * 1000);
      node.slack = Math.round(slackDays);
      node.isCritical = node.slack === 0;
    });

    const criticalTasks = Object.values(taskMap)
      .filter(node => node.isCritical)
      .map(node => ({
        taskId: node.task.id,
        title: node.task.title,
        duration: node.duration,
        earliestStart: node.earliestStart,
        earliestFinish: node.earliestFinish,
        slack: node.slack
      }))
      .sort((a, b) => a.earliestStart - b.earliestStart);

    const projectStart = criticalTasks.length > 0 ? criticalTasks[0].earliestStart : new Date();
    const totalDuration = Math.ceil((projectEnd - projectStart) / (24 * 60 * 60 * 1000));

    return {
      criticalPath: criticalTasks,
      totalDuration: totalDuration,
      projectStart: projectStart,
      projectEnd: projectEnd,
      allTasks: Object.values(taskMap).map(node => ({
        taskId: node.task.id,
        title: node.task.title,
        duration: node.duration,
        slack: node.slack,
        isCritical: node.isCritical
      }))
    };
  } catch (error) {
    console.error('calculateCriticalPath error:', error);
    return {
      criticalPath: [],
      totalDuration: 0,
      projectStart: null,
      projectEnd: null,
      allTasks: []
    };
  }
}

function calculateTaskDuration(task) {
  if (task.startDate && task.dueDate) {
    const start = new Date(task.startDate);
    const end = new Date(task.dueDate);
    const days = Math.ceil((end - start) / (24 * 60 * 60 * 1000));
    return Math.max(days, 1);
  }
  if (task.estimatedHrs > 0) {
    return Math.ceil(task.estimatedHrs / 8);
  }
  return 1;
}

function getDependencyChain(taskId) {
  try {
    const depMap = getAllDependenciesMap();
    const visited = new Set();
    const chain = [];

    function traverse(currentId) {
      if (visited.has(currentId)) return;
      visited.add(currentId);
      chain.push(currentId);
      const preds = depMap.bySuccessor[currentId] || [];
      preds.forEach(dep => traverse(dep.predecessorId));
    }

    traverse(taskId);
    return chain.slice(1);
  } catch (error) {
    console.error('getDependencyChain error:', error);
    return [];
  }
}

function getImpactChain(taskId) {
  try {
    const depMap = getAllDependenciesMap();
    const visited = new Set();
    const chain = [];

    function traverse(currentId) {
      if (visited.has(currentId)) return;
      visited.add(currentId);
      chain.push(currentId);
      const succs = depMap.byPredecessor[currentId] || [];
      succs.forEach(dep => traverse(dep.successorId));
    }

    traverse(taskId);
    return chain.slice(1);
  } catch (error) {
    console.error('getImpactChain error:', error);
    return [];
  }
}

function validateAllDependencies() {
  try {
    const depMap = getAllDependenciesMap();
    const allDeps = [
      ...Object.values(depMap.bySuccessor).flat(),
      ...Object.values(depMap.byPredecessor).flat()
    ];
    const seen = new Set();
    const uniqueDeps = allDeps.filter(d => {
      if (seen.has(d.id)) return false;
      seen.add(d.id);
      return true;
    });

    if (uniqueDeps.length === 0) {
      return { valid: true, issues: [] };
    }

    const issues = [];
    for (const dependency of uniqueDeps) {
      const successor = getTaskById(dependency.successorId);
      const predecessor = getTaskById(dependency.predecessorId);

      if (!successor) {
        issues.push({
          dependencyId: dependency.id,
          issue: 'orphaned',
          message: `Successor task not found: ${dependency.successorId}`
        });
      }
      if (!predecessor) {
        issues.push({
          dependencyId: dependency.id,
          issue: 'orphaned',
          message: `Predecessor task not found: ${dependency.predecessorId}`
        });
      }
      if (successor && predecessor) {
        const circular = detectCircularDependency(dependency.successorId, dependency.predecessorId);
        if (circular) {
          issues.push({
            dependencyId: dependency.id,
            issue: 'circular',
            message: `Circular dependency: ${circular.join(' → ')}`
          });
        }
      }
    }

    return {
      valid: issues.length === 0,
      issues: issues,
      totalDependencies: uniqueDeps.length
    };
  } catch (error) {
    console.error('validateAllDependencies error:', error);
    return {
      valid: false,
      issues: [{ issue: 'error', message: error.message }]
    };
  }
}
