function saveNewProject(projectData) {
  const result = createProject(projectData);
  invalidateProjectCache();
  try {
    CacheService.getScriptCache().remove('BATCH_DATA_CACHE');
  } catch (e) { /* best effort */ }
  return result;
}

function loadProjects() {
  return getAllProjectsOptimized();
}

function fetchSurgeTasks() {
  try {
    const tasks = getOpenSurgeTasks();
    return { success: true, tasks: tasks };
  } catch (error) {
    console.error('fetchSurgeTasks failed:', error);
    return { success: false, error: error.message, tasks: [] };
  }
}

function pickUpSurgeTask(taskId) {
  try {
    const task = pickUpTask(taskId);
    invalidateTaskCache(taskId, 'update');
    return { success: true, task: task };
  } catch (error) {
    console.error('pickUpSurgeTask failed:', error);
    return { success: false, error: error.message };
  }
}

function getProjectDetailsList() {
  try {
    const projects = getAllProjectsOptimized();
    return { success: true, projects: projects };
  } catch (error) {
    console.error('getProjectDetailsList failed:', error);
    return { success: false, error: error.message, projects: [] };
  }
}

function getProjectDetails(projectId) {
  try {
    const project = getProjectById(projectId);
    if (!project) return { success: false, error: 'Project not found' };
    return { success: true, project: project };
  } catch (error) {
    console.error('getProjectDetails failed:', error);
    return { success: false, error: error.message };
  }
}

function updateProjectDetails(payload) {
  try {
    var projectId = payload.id;
    if (!projectId) return { success: false, error: 'Project ID is required' };
    var updates = Object.assign({}, payload);
    delete updates.id;
    const project = updateProject(projectId, updates);
    invalidateProjectCache(projectId);
    return { success: true, project: project };
  } catch (error) {
    console.error('updateProjectDetails failed:', error);
    return { success: false, error: error.message };
  }
}

function addProjectReleaseNote(projectId, releaseData) {
  try {
    const project = addReleaseNote(projectId, releaseData);
    invalidateProjectCache(projectId);
    return { success: true, project: project };
  } catch (error) {
    console.error('addProjectReleaseNote failed:', error);
    return { success: false, error: error.message };
  }
}

function addProjectChangelogEntry(projectId, entry) {
  try {
    const project = addChangelogEntry(projectId, entry);
    invalidateProjectCache(projectId);
    return { success: true, project: project };
  } catch (error) {
    console.error('addProjectChangelogEntry failed:', error);
    return { success: false, error: error.message };
  }
}

function archiveProject(projectId) {
  try {
    var result = updateProject(projectId, { status: 'archived' });
    invalidateProjectCache();
    return { success: true };
  } catch (error) {
    console.error('archiveProject failed:', error);
    return { success: false, error: error.message };
  }
}

function unarchiveProject(projectId) {
  try {
    var result = updateProject(projectId, { status: 'active' });
    invalidateProjectCache();
    return { success: true };
  } catch (error) {
    console.error('unarchiveProject failed:', error);
    return { success: false, error: error.message };
  }
}

function deleteExistingProject(projectId) {
  try {
    deleteProject(projectId);
    invalidateProjectCache();
    try { CacheService.getScriptCache().remove('BATCH_DATA_CACHE'); } catch (e) {}
    return { success: true };
  } catch (error) {
    console.error('deleteExistingProject failed:', error);
    return { success: false, error: error.message };
  }
}

function updateProjectAssignees(projectId, ownerId, secondaryUsers) {
  try {
    var project = getProjectById(projectId);
    if (!project) return { success: false, error: 'Project not found' };
    var settings = {};
    try { settings = typeof project.settings === 'string' ? JSON.parse(project.settings) : (project.settings || {}); } catch (e) { settings = {}; }
    settings.secondaryUsers = secondaryUsers || [];
    var updates = { settings: JSON.stringify(settings) };
    if (ownerId) updates.ownerId = ownerId;
    updateProject(projectId, updates);
    invalidateProjectCache();
    return { success: true };
  } catch (error) {
    console.error('updateProjectAssignees failed:', error);
    return { success: false, error: error.message };
  }
}

function importWorkLogProjects(workbookId) {
  try {
    var result = importProjectsFromWorkLog(workbookId);
    return result;
  } catch (error) {
    console.error('importWorkLogProjects failed:', error);
    return { success: false, error: error.message };
  }
}

function syncWorkLogProjects(workbookId) {
  try {
    var result = syncWorklogSettings(workbookId);
    return result;
  } catch (error) {
    console.error('syncWorkLogProjects failed:', error);
    return { success: false, error: error.message };
  }
}

function getStoredProjectsWorkbookId() {
  try {
    return { success: true, workbookId: getProjectsWorkbookId() };
  } catch (error) {
    console.error('getStoredProjectsWorkbookId failed:', error);
    return { success: false, workbookId: '' };
  }
}

function saveProjectsWorkbookId(workbookId) {
  try {
    PermissionGuard.requirePermission('admin:settings');
    setProjectsWorkbookId(workbookId);
    return { success: true };
  } catch (error) {
    console.error('saveProjectsWorkbookId failed:', error);
    return { success: false, error: error.message };
  }
}

