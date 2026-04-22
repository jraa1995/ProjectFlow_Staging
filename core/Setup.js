const WORKBOOK_SHEET_SPEC = [
  { name: 'Tasks',                 columnsKey: 'TASK_COLUMNS' },
  { name: 'Users',                 columnsKey: 'USER_COLUMNS' },
  { name: 'Projects',              columnsKey: 'PROJECT_COLUMNS' },
  { name: 'Comments',              columnsKey: 'COMMENT_COLUMNS' },
  { name: 'Activity',              columnsKey: 'ACTIVITY_COLUMNS' },
  { name: 'Mentions',              columnsKey: 'MENTION_COLUMNS' },
  { name: 'Notifications',         columnsKey: 'NOTIFICATION_COLUMNS' },
  { name: 'Analytics_Cache',       columnsKey: 'ANALYTICS_CACHE_COLUMNS' },
  { name: 'Task_Dependencies',     columnsKey: 'TASK_DEPENDENCY_COLUMNS' },
  { name: 'Funnel_Staging',        columnsKey: 'FUNNEL_STAGING_COLUMNS' },
  { name: 'Roles',                 columnsKey: 'ROLE_COLUMNS',              seed: 'defaultRoles' },
  { name: 'Project_Members',       columnsKey: 'PROJECT_MEMBER_COLUMNS' },
  { name: 'Organizations',         columnsKey: 'ORGANIZATION_COLUMNS' },
  { name: 'Organization_Members',  columnsKey: 'ORGANIZATION_MEMBER_COLUMNS' },
  { name: 'Teams',                 columnsKey: 'TEAM_COLUMNS' },
  { name: 'Team_Members',          columnsKey: 'TEAM_MEMBER_COLUMNS' },
  { name: 'User_Preferences',      columnsKey: 'USER_PREFERENCE_COLUMNS' },
  { name: 'Automation_Rules',      columnsKey: 'AUTOMATION_RULE_COLUMNS' },
  { name: 'Team_Capacity',         columnsKey: 'TEAM_CAPACITY_COLUMNS' },
  { name: 'SLA_Config',            columnsKey: 'SLA_CONFIG_COLUMNS' },
  { name: 'Triage_Queue',          columnsKey: 'TRIAGE_QUEUE_COLUMNS' },
  { name: 'JSON_Cache',            columnsKey: 'JSON_CACHE_COLUMNS' },
  { name: 'Data_Assets',           columnsKey: 'DATA_ASSET_COLUMNS' },
  { name: 'Access_Requests',       columnsKey: 'ACCESS_REQUEST_COLUMNS' },
  { name: 'User_Badges',           columnsKey: 'USER_BADGE_COLUMNS' },
  { name: 'User_Metrics',          columnsKey: 'USER_METRIC_COLUMNS',       seed: 'userMetrics' }
];

const WORKBOOK_DEPRECATED_SHEETS = ['Webhook_Subscriptions'];

function normalizeWorkbook(options) {
  PermissionGuard.requirePermission('admin:settings');
  options = options || {};
  const dryRun = options.dryRun === true;
  const archiveOrphans = options.archiveOrphans === true;

  const ss = getColonySpreadsheet_();
  const report = {
    dryRun: dryRun,
    timestamp: now(),
    createdSheets: [],
    columnsAdded: [],
    columnOrderWarnings: [],
    seeded: [],
    orphanSheets: [],
    archivedSheets: [],
    deprecatedFound: [],
    unchanged: [],
    errors: []
  };

  const expectedNames = {};
  WORKBOOK_SHEET_SPEC.forEach(s => { expectedNames[s.name] = true; });
  WORKBOOK_DEPRECATED_SHEETS.forEach(n => { expectedNames[n] = true; });

  WORKBOOK_SHEET_SPEC.forEach(spec => {
    try {
      const columns = CONFIG[spec.columnsKey];
      if (!Array.isArray(columns)) {
        report.errors.push({ sheet: spec.name, error: 'Missing CONFIG.' + spec.columnsKey });
        return;
      }

      let sheet = ss.getSheetByName(spec.name);
      if (!sheet) {
        if (!dryRun) {
          sheet = ss.insertSheet(spec.name);
          sheet.getRange(1, 1, 1, columns.length)
            .setValues([columns])
            .setFontWeight('bold')
            .setBackground('#1e293b')
            .setFontColor('white');
          sheet.setFrozenRows(1);
        }
        report.createdSheets.push(spec.name);
      } else {
        const lastCol = Math.max(1, sheet.getLastColumn());
        const headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
          .map(h => String(h).trim());

        const missing = columns.filter(c => headerRow.indexOf(c) === -1);
        if (missing.length > 0) {
          if (!dryRun) {
            let startCol = headerRow.length + 1;
            if (headerRow.length === 1 && !headerRow[0]) startCol = 1;
            sheet.getRange(1, startCol, 1, missing.length)
              .setValues([missing])
              .setFontWeight('bold')
              .setBackground('#1e293b')
              .setFontColor('white');
          }
          report.columnsAdded.push({ sheet: spec.name, columns: missing });
        }

        const presentExpected = columns.filter(c => headerRow.indexOf(c) !== -1);
        const orderInSheet = presentExpected.map(c => headerRow.indexOf(c));
        const isSorted = orderInSheet.every((v, i, a) => i === 0 || a[i - 1] <= v);
        if (!isSorted) {
          report.columnOrderWarnings.push({
            sheet: spec.name,
            message: 'Column order in sheet differs from CONFIG — rowToObject may misalign. Manual reorder required.',
            expected: columns,
            actual: headerRow
          });
        }

        if (missing.length === 0 && isSorted) {
          report.unchanged.push(spec.name);
        }
      }

      if (spec.seed === 'defaultRoles') {
        try {
          const rolesSheet = ss.getSheetByName(spec.name);
          const rowCount = rolesSheet ? Math.max(0, rolesSheet.getLastRow() - 1) : 0;
          const defaultRoleNames = Object.keys(PermissionGuard.DEFAULT_ROLES);
          if (dryRun) {
            report.seeded.push({ sheet: spec.name, wouldAdd: defaultRoleNames.length - rowCount });
          } else {
            const added = PermissionGuard.initializeDefaultRoles();
            if (added > 0) report.seeded.push({ sheet: spec.name, rowsAdded: added });
          }
        } catch (e) {
          console.error('Role seeding failed:', e);
          report.errors.push({ sheet: spec.name, error: 'Role seed: ' + e.message });
        }
      }

      if (spec.seed === 'userMetrics' && options.rebuildMetrics) {
        try {
          if (dryRun) {
            report.seeded.push({ sheet: spec.name, wouldRebuild: 'all users' });
          } else {
            const res = UserMetricsService.rebuildAllMetrics();
            report.seeded.push({ sheet: spec.name, rebuilt: res.rebuilt, errors: res.errors.length });
          }
        } catch (e) {
          console.error('Metrics rebuild failed:', e);
          report.errors.push({ sheet: spec.name, error: 'Metrics rebuild: ' + e.message });
        }
      }
    } catch (e) {
      console.error('normalizeWorkbook failed for ' + spec.name + ':', e);
      report.errors.push({ sheet: spec.name, error: e.message });
    }
  });

  ss.getSheets().forEach(s => {
    const name = s.getName();
    if (expectedNames[name]) return;
    if (name.indexOf('_ARCHIVED_') === 0) return;

    const rowCount = Math.max(0, s.getLastRow() - 1);

    if (WORKBOOK_DEPRECATED_SHEETS.indexOf(name) !== -1) {
      report.deprecatedFound.push({ sheet: name, rowCount: rowCount });
      return;
    }

    report.orphanSheets.push({ sheet: name, rowCount: rowCount });

    if (archiveOrphans && !dryRun && rowCount === 0) {
      const archivedName = '_ARCHIVED_' + name + '_' + Date.now();
      try {
        s.setName(archivedName);
        report.archivedSheets.push({ oldName: name, newName: archivedName });
      } catch (e) {
        report.errors.push({ sheet: name, error: 'Archive rename failed: ' + e.message });
      }
    }
  });

  if (options.migrateBadges) {
    try {
      if (dryRun) {
        report.seeded.push({ sheet: 'User_Badges', wouldMigrate: 'user.jsonData badges' });
      } else {
        const migRes = BadgeEngine.migrateBadgesFromJsonData();
        report.seeded.push({ sheet: 'User_Badges', migrated: migRes.migrated });
      }
    } catch (e) {
      console.error('Badge migration failed:', e);
      report.errors.push({ sheet: 'User_Badges', error: 'Badge migration: ' + e.message });
    }
  }

  if (!dryRun) {
    invalidateSystemInitCache_();
  }

  return report;
}

function normalizeWorkbookDryRun() {
  return normalizeWorkbook({ dryRun: true });
}

function normalizeWorkbookFull() {
  return normalizeWorkbook({ migrateBadges: true, rebuildMetrics: true });
}

function migrateBadgesFromJsonData() {
  PermissionGuard.requirePermission('admin:settings');
  return BadgeEngine.migrateBadgesFromJsonData();
}

function rebuildAllUserMetrics() {
  PermissionGuard.requirePermission('admin:settings');
  return UserMetricsService.rebuildAllMetrics();
}

function invalidateSystemInitCache_() {
  try {
    CacheService.getScriptCache().remove('SYSTEM_INITIALIZED');
  } catch (e) {
    console.error('invalidateSystemInitCache_ failed:', e);
  }
}

function quickSetup() {
  try {
    PermissionGuard.requirePermission('admin:settings');
    if (typeof CONFIG === 'undefined') {
      throw new Error('CONFIG object is undefined. Check Config.gs for syntax errors.');
    }

    const ss = getColonySpreadsheet_();
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
  PermissionGuard.requirePermission('admin:settings');
  const ss = getColonySpreadsheet_();
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
  PermissionGuard.requirePermission('admin:settings');
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
  PermissionGuard.requirePermission('admin:settings');
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Reset System',
    'This will DELETE ALL DATA. Are you sure?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return false;
  }

  const ss = getColonySpreadsheet_();
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
        const task = createTask({ title: 'Test Task ' + Date.now(), status: 'To Do', projectId: 'TEST' });
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
  const ss = getColonySpreadsheet_();
  const issues = [];

  WORKBOOK_SHEET_SPEC.forEach(spec => {
    const sheet = ss.getSheetByName(spec.name);
    if (!sheet) {
      issues.push(`Missing sheet: ${spec.name}`);
      return;
    }
    const columns = CONFIG[spec.columnsKey] || [];
    const lastCol = Math.max(1, sheet.getLastColumn());
    const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
    const missingCols = columns.filter(c => header.indexOf(c) === -1);
    if (missingCols.length > 0) {
      issues.push(`Sheet ${spec.name} missing columns: ${missingCols.join(', ')}`);
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

function createTeamMemberByEmail(email, name, role) {
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
  PermissionGuard.requirePermission('admin:settings');
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
  PermissionGuard.requirePermission('admin:settings');
  var sheet = getTasksSheet();
  var columns = CONFIG.TASK_COLUMNS;
  var data = sheet.getDataRange().getValues();
  var idIndex = 0;
  var assigneeIndex = columns.indexOf('assignee');
  var updatedAtIndex = columns.indexOf('updatedAt');
  if (assigneeIndex === -1) return 0;

  var idSet = {};
  taskIds.forEach(function(id) { idSet[id] = true; });

  var rangesToWrite = [];
  var timestamp = now();
  for (var i = 1; i < data.length; i++) {
    if (idSet[data[i][idIndex]]) {
      data[i][assigneeIndex] = assigneeEmail;
      if (updatedAtIndex !== -1) data[i][updatedAtIndex] = timestamp;
      rangesToWrite.push({ rowIndex: i + 1, rowData: data[i] });
    }
  }

  if (rangesToWrite.length === 0) return 0;
  rangesToWrite.sort(function(a, b) { return a.rowIndex - b.rowIndex; });

  var batchStart = 0;
  while (batchStart < rangesToWrite.length) {
    var batchEnd = batchStart;
    while (batchEnd + 1 < rangesToWrite.length &&
           rangesToWrite[batchEnd + 1].rowIndex === rangesToWrite[batchEnd].rowIndex + 1) {
      batchEnd++;
    }
    var startRow = rangesToWrite[batchStart].rowIndex;
    var batchData = [];
    for (var b = batchStart; b <= batchEnd; b++) {
      batchData.push(rangesToWrite[b].rowData);
    }
    sheet.getRange(startRow, 1, batchData.length, columns.length).setValues(batchData);
    batchStart = batchEnd + 1;
  }
  return rangesToWrite.length;
}

function archiveOldCompletedTasks(daysOld) {
  PermissionGuard.requirePermission('admin:settings');
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
