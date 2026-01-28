function getFilteredTimelineData(projectId = null, filters = {}) {
  try {
    const timelineData = TimelineEngine.generateProjectTimeline(projectId, filters.dateRange);
    let filteredTasks = timelineData.tasks;
    if (filters.assignee) {
      filteredTasks = filteredTasks.filter(task => task.assignee === filters.assignee);
    }
    if (filters.status) {
      filteredTasks = filteredTasks.filter(task => task.status === filters.status);
    }
    if (filters.priority) {
      filteredTasks = filteredTasks.filter(task => task.priority === filters.priority);
    }
    if (filters.type) {
      filteredTasks = filteredTasks.filter(task => task.type === filters.type);
    }
    if (filters.search) {
      const searchTerm = filters.search.toLowerCase();
      filteredTasks = filteredTasks.filter(task =>
        task.name.toLowerCase().includes(searchTerm) ||
        task.id.toLowerCase().includes(searchTerm) ||
        (task.assignee && task.assignee.toLowerCase().includes(searchTerm)) ||
        (task.labels && task.labels.some(label => label.toLowerCase().includes(searchTerm)))
      );
    }
    if (filters.showOverdueOnly) {
      const now = new Date();
      filteredTasks = filteredTasks.filter(task =>
        task.end < now && task.status !== 'Done'
      );
    }
    if (filters.completionStatus) {
      if (filters.completionStatus === 'completed') {
        filteredTasks = filteredTasks.filter(task => task.status === 'Done');
      } else if (filters.completionStatus === 'incomplete') {
        filteredTasks = filteredTasks.filter(task => task.status !== 'Done');
      }
    }
    if (filters.taskDateRange) {
      filteredTasks = filteredTasks.filter(task => {
        const taskStart = task.start;
        const taskEnd = task.end;
        return (taskStart >= filters.taskDateRange.start && taskStart <= filters.taskDateRange.end) ||
          (taskEnd >= filters.taskDateRange.start && taskEnd <= filters.taskDateRange.end) ||
          (taskStart <= filters.taskDateRange.start && taskEnd >= filters.taskDateRange.end);
      });
    }
    const highlightedTasks = highlightOverdueTasks(filteredTasks);
    const filteredTaskIds = new Set(highlightedTasks.map(t => t.id));
    const filteredDependencies = timelineData.dependencies.filter(dep =>
      filteredTaskIds.has(dep.from) && filteredTaskIds.has(dep.to)
    );
    const filteredCriticalPath = TimelineEngine.calculateCriticalPath(
      highlightedTasks,
      filteredDependencies
    );
    return {
      ...timelineData,
      tasks: highlightedTasks,
      dependencies: filteredDependencies,
      criticalPath: filteredCriticalPath,
      filterStats: {
        totalTasks: timelineData.tasks.length,
        filteredTasks: highlightedTasks.length,
        overdueTasks: countOverdueTasks(highlightedTasks),
        completedTasks: countCompletedTasks(highlightedTasks),
        criticalPathTasks: filteredCriticalPath.length
      }
    };
  } catch (error) {
    throw new Error('Failed to filter timeline data: ' + error.message);
  }
}

function highlightOverdueTasks(tasks) {
  const now = new Date();
  return tasks.map(task => {
    const isOverdue = task.end < now && task.status !== 'Done';
    return {
      ...task,
      isOverdue: isOverdue,
      overdueBy: isOverdue ? Math.ceil((now - task.end) / (1000 * 60 * 60 * 24)) : 0,
      customClass: (task.customClass || '') + (isOverdue ? ' overdue' : '')
    };
  });
}

function countOverdueTasks(tasks) {
  const now = new Date();
  return tasks.filter(task => task.end < now && task.status !== 'Done').length;
}

function countCompletedTasks(tasks) {
  return tasks.filter(task => task.status === 'Done').length;
}

function searchTimelineTasks(query, projectId = null, searchOptions = {}) {
  try {
    const timelineData = TimelineEngine.generateProjectTimeline(projectId);
    if (!query || query.trim().length === 0) {
      return {
        tasks: timelineData.tasks,
        totalResults: timelineData.tasks.length,
        searchQuery: '',
        searchTime: 0
      };
    }
    const startTime = Date.now();
    const searchTerm = query.toLowerCase().trim();
    const searchCriteria = parseSearchQuery(searchTerm);
    let results = timelineData.tasks.filter(task => {
      if (searchCriteria.text) {
        const textMatch =
          task.name.toLowerCase().includes(searchCriteria.text) ||
          task.id.toLowerCase().includes(searchCriteria.text) ||
          (task.assignee && task.assignee.toLowerCase().includes(searchCriteria.text));
        if (!textMatch) return false;
      }
      if (searchCriteria.status && task.status.toLowerCase() !== searchCriteria.status) {
        return false;
      }
      if (searchCriteria.priority && task.priority.toLowerCase() !== searchCriteria.priority) {
        return false;
      }
      if (searchCriteria.assignee &&
          (!task.assignee || !task.assignee.toLowerCase().includes(searchCriteria.assignee))) {
        return false;
      }
      if (searchCriteria.label &&
          (!task.labels || !task.labels.some(label =>
            label.toLowerCase().includes(searchCriteria.label)))) {
        return false;
      }
      if (searchCriteria.overdue) {
        const now = new Date();
        const isOverdue = task.end < now && task.status !== 'Done';
        if (!isOverdue) return false;
      }
      return true;
    });
    if (searchCriteria.text) {
      results = sortSearchResults(results, searchCriteria.text);
    }
    const searchTime = Date.now() - startTime;
    return {
      tasks: results,
      totalResults: results.length,
      searchQuery: query,
      searchTime: searchTime,
      searchCriteria: searchCriteria
    };
  } catch (error) {
    throw new Error('Failed to search timeline tasks: ' + error.message);
  }
}

function parseSearchQuery(query) {
  const criteria = {
    text: '',
    status: null,
    priority: null,
    assignee: null,
    label: null,
    overdue: false
  };
  const quotedPhrases = [];
  let quotedQuery = query.replace(/"([^"]+)"/g, (match, phrase) => {
    quotedPhrases.push(phrase);
    return '';
  });
  const fieldFilters = [
    { pattern: /status:(\w+)/gi, field: 'status' },
    { pattern: /priority:(\w+)/gi, field: 'priority' },
    { pattern: /assignee:(\w+)/gi, field: 'assignee' },
    { pattern: /label:(\w+)/gi, field: 'label' },
    { pattern: /overdue:(true|false)/gi, field: 'overdue' }
  ];
  fieldFilters.forEach(filter => {
    let match;
    while ((match = filter.pattern.exec(quotedQuery)) !== null) {
      if (filter.field === 'overdue') {
        criteria[filter.field] = match[1].toLowerCase() === 'true';
      } else {
        criteria[filter.field] = match[1].toLowerCase();
      }
      quotedQuery = quotedQuery.replace(match[0], '');
    }
  });
  const remainingText = quotedQuery.trim();
  const allText = [remainingText, ...quotedPhrases].filter(t => t).join(' ').trim();
  if (allText) {
    criteria.text = allText;
  }
  return criteria;
}

function sortSearchResults(tasks, searchTerm) {
  return tasks.sort((a, b) => {
    const aScore = calculateRelevanceScore(a, searchTerm);
    const bScore = calculateRelevanceScore(b, searchTerm);
    return bScore - aScore;
  });
}

function calculateRelevanceScore(task, searchTerm) {
  let score = 0;
  const term = searchTerm.toLowerCase();
  if (task.name.toLowerCase() === term) {
    score += 100;
  } else if (task.name.toLowerCase().includes(term)) {
    score += 50;
  }
  if (task.id.toLowerCase() === term) {
    score += 80;
  } else if (task.id.toLowerCase().includes(term)) {
    score += 30;
  }
  if (task.assignee && task.assignee.toLowerCase().includes(term)) {
    score += 20;
  }
  if (task.labels && task.labels.some(label => label.toLowerCase().includes(term))) {
    score += 15;
  }
  if (task.priority === 'Critical' || task.priority === 'High') {
    score += 10;
  }
  const now = new Date();
  if (task.end < now && task.status !== 'Done') {
    score += 5;
  }
  return score;
}

function getTimelineFilterSuggestions(projectId = null) {
  try {
    const timelineData = TimelineEngine.generateProjectTimeline(projectId);
    const assignees = [...new Set(timelineData.tasks.map(t => t.assignee).filter(a => a))];
    const statuses = [...new Set(timelineData.tasks.map(t => t.status))];
    const priorities = [...new Set(timelineData.tasks.map(t => t.priority))];
    const types = [...new Set(timelineData.tasks.map(t => t.type))];
    const labels = [...new Set(timelineData.tasks.flatMap(t => t.labels || []))];
    const now = new Date();
    const overdueTasks = timelineData.tasks.filter(t => t.end < now && t.status !== 'Done');
    const completedTasks = timelineData.tasks.filter(t => t.status === 'Done');
    return {
      assignees: assignees.sort(),
      statuses: statuses.sort(),
      priorities: priorities.sort(),
      types: types.sort(),
      labels: labels.sort(),
      statistics: {
        totalTasks: timelineData.tasks.length,
        overdueTasks: overdueTasks.length,
        completedTasks: completedTasks.length,
        inProgressTasks: timelineData.tasks.filter(t =>
          t.status === 'In Progress' || t.status === 'Review' || t.status === 'Testing'
        ).length
      }
    };
  } catch (error) {
    return {
      assignees: [],
      statuses: [],
      priorities: [],
      types: [],
      labels: [],
      statistics: {
        totalTasks: 0,
        overdueTasks: 0,
        completedTasks: 0,
        inProgressTasks: 0
      }
    };
  }
}
