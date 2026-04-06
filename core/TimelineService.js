function getTimelineData(projectId) {
  const logs = [];

  try {
    if (typeof TimelineEngine === 'undefined') {
      return {
        tasks: [],
        milestones: [],
        dependencies: [],
        criticalPath: [],
        dateRange: { start: new Date(), end: new Date() },
        projectId: projectId,
        error: 'TimelineEngine not loaded',
        logs: logs
      };
    }

    const result = TimelineEngine.generateProjectTimeline(projectId || null);

    if (result) {
      try {
        JSON.stringify(result);
      } catch (serError) {
        return {
          tasks: [],
          milestones: [],
          dependencies: [],
          criticalPath: [],
          dateRange: { start: new Date().toISOString(), end: new Date().toISOString() },
          projectId: projectId,
          error: 'Result serialization failed: ' + serError.message,
          logs: logs
        };
      }

      result.logs = logs;
      return result;
    }

    return {
      tasks: [],
      milestones: [],
      dependencies: [],
      criticalPath: [],
      dateRange: { start: new Date().toISOString(), end: new Date().toISOString() },
      projectId: projectId,
      error: 'TimelineEngine returned null',
      logs: logs
    };
  } catch (error) {
    return {
      tasks: [],
      milestones: [],
      dependencies: [],
      criticalPath: [],
      dateRange: { start: new Date(), end: new Date() },
      projectId: projectId,
      error: error.message,
      logs: logs
    };
  }
}

function getFilteredTimeline(projectId, filters) {
  try {
    if (typeof TimelineEngine === 'undefined') {
      console.error('getFilteredTimeline: TimelineEngine is not defined!');
      return {
        tasks: [],
        milestones: [],
        dependencies: [],
        criticalPath: [],
        dateRange: { start: new Date(), end: new Date() },
        projectId: projectId,
        error: 'TimelineEngine not loaded'
      };
    }

    const timelineData = TimelineEngine.generateProjectTimeline(projectId || null);

    if (!timelineData) {
      return {
        tasks: [],
        milestones: [],
        dependencies: [],
        criticalPath: [],
        dateRange: { start: new Date(), end: new Date() },
        projectId: projectId
      };
    }

    if (!timelineData.tasks) {
      timelineData.tasks = [];
    }

    if (!filters || Object.keys(filters).length === 0) {
      return timelineData;
    }

    let tasks = timelineData.tasks;

    if (filters.status && filters.status.length > 0) {
      tasks = tasks.filter(t => filters.status.includes(t.status));
    }

    if (filters.priority && filters.priority.length > 0) {
      tasks = tasks.filter(t => filters.priority.includes(t.priority));
    }

    if (filters.assignee && filters.assignee.length > 0) {
      tasks = tasks.filter(t => filters.assignee.includes(t.assignee));
    }

    if (filters.search) {
      const searchLower = filters.search.toLowerCase();
      tasks = tasks.filter(t =>
        (t.name && t.name.toLowerCase().includes(searchLower)) ||
        (t.id && t.id.toLowerCase().includes(searchLower))
      );
    }

    return {
      ...timelineData,
      tasks: tasks
    };
  } catch (error) {
    console.error('getFilteredTimeline failed:', error);
    return {
      tasks: [],
      milestones: [],
      dependencies: [],
      criticalPath: [],
      dateRange: { start: new Date(), end: new Date() },
      projectId: projectId,
      error: error.message
    };
  }
}

function diagnosticTimeline() {
  const result = {
    timelineEngineAvailable: typeof TimelineEngine !== 'undefined',
    timelineEngineType: typeof TimelineEngine,
    getAllTasksAvailable: typeof getAllTasks !== 'undefined',
    sampleTaskCount: 0,
    sampleTask: null,
    errors: []
  };

  try {
    const tasks = getAllTasks();
    result.sampleTaskCount = tasks.length;

    if (tasks.length > 0) {
      const sample = tasks[0];
      result.sampleTask = {
        id: sample.id,
        title: sample.title,
        startDate: sample.startDate,
        startDateType: typeof sample.startDate,
        dueDate: sample.dueDate,
        dueDateType: typeof sample.dueDate,
        createdAt: sample.createdAt
      };
    }

    if (typeof TimelineEngine !== 'undefined' && tasks.length > 0) {
      const testDate = '1/16/2026';
      const parsed = TimelineEngine.parseFlexibleDate(testDate);
      result.dateParseTest = {
        input: testDate,
        output: parsed,
        outputType: typeof parsed,
        isDate: parsed instanceof Date,
        isValid: parsed && !isNaN(parsed.getTime())
      };
    }
  } catch (error) {
    result.errors.push(error.message);
  }

  return result;
}

function getMilestonesForView() {
  try {
    const milestones = getMilestones();
    const projects = getAllProjectsOptimized();

    return {
      success: true,
      milestones: milestones,
      projects: projects
    };
  } catch (error) {
    console.error('getMilestonesForView failed:', error);
    return {
      success: false,
      error: error.message,
      milestones: [],
      projects: []
    };
  }
}
