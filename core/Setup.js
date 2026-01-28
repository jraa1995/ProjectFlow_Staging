function quickSetup() {
  try {
    if (typeof CONFIG === 'undefined') {
      throw new Error('CONFIG object is undefined. Check Config.gs for syntax errors.');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const existingSheets = ss.getSheets().map(s => s.getName());

    initializeSystem();

    const newSheets = ss.getSheets().map(s => s.getName());
    const user = getCurrentUser();
    const projects = getAllProjects();

    if (projects.length === 0) {
      const project = createProject({
        name: 'Default Project',
        description: 'Your first project - rename or delete as needed'
      });
    }

    const tasks = getAllTasks();
    if (tasks.length === 0) {
      createSampleTasks();
    }

    const testResults = runQuickTests();

    return {
      success: true,
      user: user.email,
      projects: getAllProjects().length,
      tasks: getAllTasks().length,
      tests: testResults
    };
  } catch (error) {
    console.error('Setup failed:', error);
    console.error('Stack:', error.stack);

    return {
      success: false,
      error: error.message,
      troubleshooting: error.message.includes('CONFIG') ?
        'Check Config.gs for syntax errors' :
        'Check console logs for detailed error information'
    };
  }
}

function createSampleTasks() {
  const projects = getAllProjects();
  const projectId = projects[0]?.id || '';
  const userEmail = getCurrentUserEmail();

  const sampleTasks = [
    {
      title: 'Set up project repository',
      description: 'Initialize Git repository and configure CI/CD pipeline',
      status: 'Done',
      priority: 'High',
      type: 'Task',
      labels: ['setup', 'infrastructure']
    },
    {
      title: 'Design database schema',
      description: 'Create ERD and define all tables, relationships, and indexes',
      status: 'Done',
      priority: 'High',
      type: 'Task',
      labels: ['database', 'design']
    },
    {
      title: 'Implement user authentication',
      description: 'Build secure login system with OAuth integration',
      status: 'In Progress',
      priority: 'Critical',
      type: 'Feature',
      labels: ['security', 'backend'],
      storyPoints: 8
    },
    {
      title: 'Create dashboard UI',
      description: 'Design and implement the main dashboard with charts and metrics',
      status: 'In Progress',
      priority: 'High',
      type: 'Feature',
      labels: ['frontend', 'ui'],
      storyPoints: 13
    },
    {
      title: 'Fix login redirect bug',
      description: 'Users not being redirected properly after successful login',
      status: 'To Do',
      priority: 'Medium',
      type: 'Bug',
      labels: ['bug', 'frontend']
    },
    {
      title: 'Write API documentation',
      description: 'Document all API endpoints with examples and schemas',
      status: 'To Do',
      priority: 'Low',
      type: 'Task',
      labels: ['documentation']
    },
    {
      title: 'Set up monitoring',
      description: 'Configure alerts and dashboards for system health',
      status: 'Backlog',
      priority: 'Medium',
      type: 'Task',
      labels: ['devops', 'monitoring']
    },
    {
      title: 'Performance optimization sprint',
      description: 'Identify and fix performance bottlenecks',
      status: 'Backlog',
      priority: 'Low',
      type: 'Story',
      labels: ['performance'],
      storyPoints: 21
    },
    {
      title: 'Mobile responsive design',
      description: 'Ensure application works on all device sizes',
      status: 'Review',
      priority: 'High',
      type: 'Feature',
      labels: ['mobile', 'responsive'],
      storyPoints: 8
    },
    {
      title: 'Security audit',
      description: 'Review code for security vulnerabilities',
      status: 'Testing',
      priority: 'Critical',
      type: 'Task',
      labels: ['security', 'audit']
    }
  ];

  sampleTasks.forEach((taskData, index) => {
    taskData.projectId = projectId;
    taskData.assignee = userEmail;
    if (index % 2 === 0) {
      const dueDate = new Date();
      dueDate.setDate(dueDate.getDate() + (index * 2) + 3);
      taskData.dueDate = dueDate.toISOString().split('T')[0];
    }
    createTask(taskData);
  });

  return sampleTasks.length;
}

function cleanupDuplicateSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const requiredSheets = ['Tasks', 'Users', 'Projects', 'Comments', 'Activity', 'Mentions', 'Notifications', 'Analytics_Cache', 'Task_Dependencies'];
  const sheetCounts = {};
  const duplicates = [];

  sheets.forEach(sheet => {
    const name = sheet.getName();
    sheetCounts[name] = (sheetCounts[name] || 0) + 1;
    if (sheetCounts[name] > 1) {
      duplicates.push(sheet);
    }
  });

  let deletedCount = 0;

  duplicates.forEach(sheet => {
    try {
      if (ss.getSheets().length > 1) {
        ss.deleteSheet(sheet);
        deletedCount++;
      }
    } catch (e) {
      console.error(`Failed to delete ${sheet.getName()}:`, e.message);
    }
  });

  const finalSheets = ss.getSheets().map(s => s.getName());
  return { deleted: deletedCount, remaining: finalSheets };
}

function clearSampleData() {
  const tasks = getAllTasks();
  let cleared = 0;

  tasks.forEach(task => {
    const labels = task.labels || [];
    if (labels.includes('setup') || labels.includes('infrastructure') ||
      task.description?.includes('sample') || task.description?.includes('Sample')) {
      deleteTask(task.id);
      cleared++;
    }
  });

  return cleared;
}

function resetSystem() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Reset System',
    'This will DELETE ALL DATA. Are you sure?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return false;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsToDelete = ['Tasks', 'Users', 'Projects', 'Comments', 'Activity', 'Mentions', 'Notifications', 'Analytics_Cache', 'Task_Dependencies'];

  sheetsToDelete.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      if (ss.getSheets().length > 1) {
        ss.deleteSheet(sheet);
      } else {
        sheet.clear();
      }
    }
  });

  return true;
}

function testSystem() {
  const tests = [
    {
      name: 'Get Current User', fn: () => {
        const user = getCurrentUser();
        return user && user.email;
      }
    },
    {
      name: 'Get All Tasks', fn: () => {
        const tasks = getAllTasks();
        return Array.isArray(tasks);
      }
    },
    {
      name: 'Get All Users', fn: () => {
        const users = getAllUsers();
        return Array.isArray(users);
      }
    },
    {
      name: 'Get All Projects', fn: () => {
        const projects = getAllProjects();
        return Array.isArray(projects);
      }
    },
    {
      name: 'Build My Board', fn: () => {
        const board = getMyBoard(null);
        return board && board.columns && board.columns.length > 0;
      }
    },
    {
      name: 'Build Master Board', fn: () => {
        const board = getMasterBoard(null);
        return board && board.columns && board.columns.length > 0;
      }
    },
    {
      name: 'Create Task', fn: () => {
        const task = createTask({ title: 'Test Task ' + Date.now(), status: 'To Do' });
        const success = task && task.id;
        if (success) deleteTask(task.id);
        return success;
      }
    },
    {
      name: 'Search Tasks', fn: () => {
        const results = searchTasks('test');
        return Array.isArray(results);
      }
    },
    {
      name: 'Get Initial Data', fn: () => {
        const data = getInitialData();
        return data && data.user && data.board && data.config;
      }
    },
    {
      name: 'Calculate Stats', fn: () => {
        const tasks = getAllTasks();
        const stats = calculateBoardStats(tasks);
        return stats && typeof stats.total === 'number';
      }
    }
  ];

  let passed = 0;
  let failed = 0;
  const results = [];

  tests.forEach(test => {
    try {
      const startTime = Date.now();
      const result = test.fn();
      const duration = Date.now() - startTime;

      if (result) {
        passed++;
        results.push({ name: test.name, passed: true, duration });
      } else {
        failed++;
        results.push({ name: test.name, passed: false, error: 'Returned false' });
      }
    } catch (error) {
      failed++;
      results.push({ name: test.name, passed: false, error: error.message });
    }
  });

  return { passed, failed, total: tests.length, results };
}

function runQuickTests() {
  const tests = [
    { name: 'User', fn: () => getCurrentUser()?.email },
    { name: 'Tasks Sheet', fn: () => getTasksSheet()?.getName() },
    { name: 'My Board', fn: () => getMyBoard(null)?.columns?.length > 0 }
  ];

  let passed = 0;

  tests.forEach(t => {
    try {
      if (t.fn()) {
        passed++;
      }
    } catch (e) {
    }
  });

  return { passed, total: tests.length };
}

function diagnose() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const issues = [];
  const requiredSheets = ['Tasks', 'Users', 'Projects', 'Comments', 'Activity', 'Mentions', 'Notifications', 'Analytics_Cache', 'Task_Dependencies'];

  requiredSheets.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) {
      issues.push(`Missing sheet: ${name}`);
    }
  });

  try {
    const email = Session.getActiveUser().getEmail();
    const user = getUserByEmail(email);
    if (!user) {
      issues.push('User not in Users sheet');
    }
  } catch (e) {
    issues.push(`User error: ${e.message}`);
  }

  try {
    const tasks = getAllTasks();
    const users = getAllUsers();
    const projects = getAllProjects();
  } catch (e) {
    issues.push(`Data error: ${e.message}`);
  }

  return { issues, issueCount: issues.length };
}

function addTeamMember(email, name, role) {
  if (!email || !email.includes('@')) {
    throw new Error('Valid email required');
  }

  const existingUser = getUserByEmail(email);
  if (existingUser) {
    return existingUser;
  }

  const user = createUser({
    email: email.toLowerCase().trim(),
    name: name || email.split('@')[0],
    role: role || 'member'
  });

  return user;
}

function listTeamMembers() {
  const users = getAllUsers();
  return users;
}

function deactivateUser(email) {
  return updateUser(email, { active: false });
}

function activateUser(email) {
  return updateUser(email, { active: true });
}

function bulkMoveTasksToStatus(fromStatus, toStatus) {
  const tasks = getAllTasks({ status: fromStatus });
  let moved = 0;

  tasks.forEach(task => {
    try {
      updateTask(task.id, { status: toStatus });
      moved++;
    } catch (e) {
      console.error(`Failed to move ${task.id}: ${e.message}`);
    }
  });

  return moved;
}

function bulkAssignTasks(taskIds, assigneeEmail) {
  let assigned = 0;

  taskIds.forEach(taskId => {
    try {
      updateTask(taskId, { assignee: assigneeEmail });
      assigned++;
    } catch (e) {
      console.error(`Failed to assign ${taskId}: ${e.message}`);
    }
  });

  return assigned;
}

function archiveOldCompletedTasks(daysOld) {
  daysOld = daysOld || 30;
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - daysOld);

  const tasks = getAllTasks({ status: 'Done' });
  let archived = 0;

  tasks.forEach(task => {
    if (task.completedAt && new Date(task.completedAt) < cutoff) {
      deleteTask(task.id);
      archived++;
    }
  });

  return archived;
}

function generateReport() {
  const tasks = getAllTasks();
  const stats = calculateBoardStats(tasks);
  const users = getAllUsers();
  const projects = getAllProjects();

  return stats;
}

function exportTasksToCSV() {
  const tasks = getAllTasks();
  const headers = CONFIG.TASK_COLUMNS.join(',');

  const rows = tasks.map(task => {
    return CONFIG.TASK_COLUMNS.map(col => {
      let value = task[col];
      if (Array.isArray(value)) value = value.join(';');
      if (typeof value === 'string' && (value.includes(',') || value.includes('"'))) {
        value = `"${value.replace(/"/g, '""')}"`;
      }
      return value || '';
    }).join(',');
  });

  const csv = [headers, ...rows].join('\n');
  return csv;
}
