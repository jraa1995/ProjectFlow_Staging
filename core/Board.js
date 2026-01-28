function getMyBoard(projectId) {
  const userEmail = getCurrentUserEmail();
  const tasks = getTasksForUser(userEmail);
  return buildBoardData(tasks, projectId, {
    view: "my",
    userEmail: userEmail,
  });
}

function getMasterBoard(projectId) {
  const tasks = getAllTasks();
  return buildBoardData(tasks, projectId, {
    view: "master",
  });
}

function buildBoardData(allTasks, projectId, options = {}) {
  let tasks = allTasks;
  if (projectId) {
    tasks = tasks.filter((t) => t.projectId === projectId);
  }

  tasks = tasks.map(task => {
    try {
      const dependencies = getTaskDependencies(task.id);
      const hasDependencies = dependencies.predecessors.length > 0 || dependencies.successors.length > 0;
      const blockedBy = dependencies.predecessors.filter(dep => {
        const predecessorTask = getTaskById(dep.predecessorId);
        return predecessorTask && predecessorTask.status !== 'Done';
      });
      const isBlocked = blockedBy.length > 0;
      return {
        ...task,
        hasDependencies: hasDependencies,
        isBlocked: isBlocked,
        blockedByCount: blockedBy.length,
        blockingCount: dependencies.successors.length,
        dependencyMetadata: {
          predecessorCount: dependencies.predecessors.length,
          successorCount: dependencies.successors.length,
          blockedByTasks: blockedBy.map(dep => {
            const t = getTaskById(dep.predecessorId);
            return t ? { id: t.id, title: t.title, status: t.status } : null;
          }).filter(Boolean)
        }
      };
    } catch (error) {
      console.error('Failed to load dependencies for task ' + task.id + ':', error);
      return task;
    }
  });

  const columns = CONFIG.STATUSES.map((status, index) => {
    const columnTasks = tasks
      .filter((t) => t.status === status)
      .sort((a, b) => (a.position || 0) - (b.position || 0));
    return {
      id: normalizeStatusId(status),
      name: status,
      color: CONFIG.COLORS[status] || "#6B7280",
      order: index,
      tasks: columnTasks,
      count: columnTasks.length,
    };
  });

  const stats = calculateBoardStats(tasks);
  const projects = getAllProjects();
  const users = getActiveUsers();

  return {
    columns: columns,
    stats: stats,
    projects: projects,
    users: users,
    taskCount: tasks.length,
    view: options.view || "board",
    userEmail: options.userEmail || null,
    projectFilter: projectId || null,
    config: {
      statuses: CONFIG.STATUSES,
      priorities: CONFIG.PRIORITIES,
      types: CONFIG.TYPES,
      colors: CONFIG.COLORS,
    },
  };
}

function normalizeStatusId(status) {
  if (!status) return "backlog";
  return String(status).toLowerCase().replace(/\s+/g, "_");
}

function isValidStatus(status) {
  if (!status || typeof status !== 'string') return false;
  return CONFIG.STATUSES.includes(status);
}

function getValidStatus(status, defaultStatus) {
  if (isValidStatus(status)) {
    return status;
  }
  const fallback = defaultStatus || CONFIG.STATUSES[0] || 'Backlog';
  console.warn('getValidStatus: Invalid status "' + status + '", using default "' + fallback + '"');
  return fallback;
}

function denormalizeStatusId(columnId) {
  if (!columnId || typeof columnId !== 'string') {
    console.warn('denormalizeStatusId: Invalid columnId "' + columnId + '", using default');
    return CONFIG.STATUSES[0] || 'Backlog';
  }

  const statusMap = {};
  CONFIG.STATUSES.forEach((s) => {
    statusMap[normalizeStatusId(s)] = s;
  });

  const status = statusMap[columnId.toLowerCase()];
  if (!status) {
    console.warn('denormalizeStatusId: Unknown columnId "' + columnId + '", valid options: ' + Object.keys(statusMap).join(', '));
    if (CONFIG.STATUSES.includes(columnId)) {
      return columnId;
    }
    return CONFIG.STATUSES[0] || 'Backlog';
  }
  return status;
}

function calculateBoardStats(tasks) {
  const total = tasks.length;
  const now = new Date();
  const weekFromNow = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000);

  const byStatus = {};
  CONFIG.STATUSES.forEach((s) => (byStatus[s] = 0));

  const byPriority = {};
  CONFIG.PRIORITIES.forEach((p) => (byPriority[p] = 0));

  const byAssignee = {};

  const byType = {};
  CONFIG.TYPES.forEach((t) => (byType[t] = 0));

  let completed = 0;
  let dueSoon = 0;
  let overdue = 0;
  let totalPoints = 0;
  let completedPoints = 0;
  let totalEstimatedHrs = 0;
  let totalActualHrs = 0;

  tasks.forEach((task) => {
    byStatus[task.status] = (byStatus[task.status] || 0) + 1;
    if (task.status === "Done") {
      completed++;
      completedPoints += task.storyPoints || 0;
    }
    byPriority[task.priority] = (byPriority[task.priority] || 0) + 1;
    const assignee = task.assignee || "Unassigned";
    byAssignee[assignee] = (byAssignee[assignee] || 0) + 1;
    byType[task.type] = (byType[task.type] || 0) + 1;
    totalPoints += task.storyPoints || 0;
    totalEstimatedHrs += task.estimatedHrs || 0;
    totalActualHrs += task.actualHrs || 0;
    if (task.dueDate && task.status !== "Done") {
      const dueDate = new Date(task.dueDate);
      if (dueDate < now) {
        overdue++;
      } else if (dueDate <= weekFromNow) {
        dueSoon++;
      }
    }
  });

  const progress = total > 0 ? Math.round((completed / total) * 100) : 0;

  return {
    total,
    completed,
    progress,
    dueSoon,
    overdue,
    totalPoints,
    completedPoints,
    totalEstimatedHrs,
    totalActualHrs,
    byStatus,
    byPriority,
    byAssignee,
    byType,
  };
}

function searchTasks(query, projectId) {
  const allTasks = getAllTasks();
  const q = query.toLowerCase();

  let results = allTasks.filter((task) => {
    return (
      task.id?.toLowerCase().includes(q) ||
      task.title?.toLowerCase().includes(q) ||
      task.description?.toLowerCase().includes(q) ||
      task.assignee?.toLowerCase().includes(q) ||
      task.labels?.some((l) => l.toLowerCase().includes(q)) ||
      task.status?.toLowerCase().includes(q) ||
      task.priority?.toLowerCase().includes(q)
    );
  });

  if (projectId) {
    results = results.filter((t) => t.projectId === projectId);
  }

  return results;
}

function getFilteredTasks(filters) {
  let tasks = getAllTasks();

  if (filters.assignee) {
    tasks = tasks.filter((t) => t.assignee === filters.assignee);
  }

  if (filters.reporter) {
    tasks = tasks.filter((t) => t.reporter === filters.reporter);
  }

  if (filters.projectId) {
    tasks = tasks.filter((t) => t.projectId === filters.projectId);
  }

  if (filters.status && filters.status.length > 0) {
    tasks = tasks.filter((t) => filters.status.includes(t.status));
  }

  if (filters.priority && filters.priority.length > 0) {
    tasks = tasks.filter((t) => filters.priority.includes(t.priority));
  }

  if (filters.type && filters.type.length > 0) {
    tasks = tasks.filter((t) => filters.type.includes(t.type));
  }

  if (filters.sprint) {
    tasks = tasks.filter((t) => t.sprint === filters.sprint);
  }

  if (filters.labels && filters.labels.length > 0) {
    tasks = tasks.filter((t) => {
      const taskLabels = t.labels || [];
      return filters.labels.some((l) => taskLabels.includes(l));
    });
  }

  if (filters.hasParent !== undefined) {
    if (filters.hasParent) {
      tasks = tasks.filter((t) => t.parentId);
    } else {
      tasks = tasks.filter((t) => !t.parentId);
    }
  }

  if (filters.overdue) {
    const now = new Date();
    tasks = tasks.filter(
      (t) => t.dueDate && new Date(t.dueDate) < now && t.status !== "Done"
    );
  }

  if (filters.dueSoon) {
    const now = new Date();
    const weekFromNow = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000);
    tasks = tasks.filter((t) => {
      if (!t.dueDate || t.status === "Done") return false;
      const dueDate = new Date(t.dueDate);
      return dueDate >= now && dueDate <= weekFromNow;
    });
  }

  return tasks;
}

function reorderTasksInColumn(taskId, newPosition, status) {
  const tasks = getAllTasks().filter((t) => t.status === status);
  const taskIndex = tasks.findIndex((t) => t.id === taskId);
  if (taskIndex === -1) return false;

  const [movedTask] = tasks.splice(taskIndex, 1);
  tasks.splice(newPosition, 0, movedTask);

  tasks.forEach((task, index) => {
    if (task.position !== index) {
      updateTask(task.id, { position: index });
    }
  });

  return true;
}

function getAllLabels() {
  const tasks = getAllTasks();
  const labelSet = new Set();

  tasks.forEach((task) => {
    if (task.labels && Array.isArray(task.labels)) {
      task.labels.forEach((l) => labelSet.add(l));
    }
  });

  return Array.from(labelSet).sort();
}

function getAllSprints() {
  const tasks = getAllTasks();
  const sprintSet = new Set();

  tasks.forEach((task) => {
    if (task.sprint) {
      sprintSet.add(task.sprint);
    }
  });

  return Array.from(sprintSet).sort((a, b) => {
    const numA = parseInt(a) || 0;
    const numB = parseInt(b) || 0;
    return numA - numB;
  });
}

function getSubtasks(parentId) {
  return getAllTasks().filter((t) => t.parentId === parentId);
}

function getParentTasks() {
  return getAllTasks().filter((t) => t.type === "Epic" || t.type === "Story");
}
