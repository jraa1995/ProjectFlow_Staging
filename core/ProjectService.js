function saveNewProject(projectData) {
  const result = createProject(projectData);
  invalidateProjectCache();
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

function getProjectTaxonomyValues() {
  try {
    var projects = getAllProjectsOptimized();
    var fields = CONFIG.TAXONOMY_FIELDS;
    var result = {};
    fields.forEach(function(f) { result[f] = {}; });
    projects.forEach(function(p) {
      var settings = {};
      try { settings = typeof p.settings === 'string' ? JSON.parse(p.settings) : (p.settings || {}); } catch(e) {}
      fields.forEach(function(f) {
        var val = settings[f];
        if (val && String(val).trim() && val !== 'N/A' && val !== 'FALSE') {
          result[f][String(val).trim()] = true;
        }
      });
    });
    var sorted = {};
    fields.forEach(function(f) { sorted[f] = Object.keys(result[f]).sort(); });
    return sorted;
  } catch (error) {
    console.error('getProjectTaxonomyValues failed:', error);
    return {};
  }
}

function getProjectHierarchy(projectId) {
  try {
    var projects = getAllProjectsOptimized();
    var project = null;
    for (var i = 0; i < projects.length; i++) {
      if (projects[i].id === projectId) { project = projects[i]; break; }
    }
    if (!project) return null;
    var parentId = null;
    try {
      var s = typeof project.settings === 'string' ? JSON.parse(project.settings) : (project.settings || {});
      parentId = s.linkedProjectId || null;
    } catch (e) {}
    var children = [];
    projects.forEach(function(p) {
      if (p.id === projectId) return;
      try {
        var ps = typeof p.settings === 'string' ? JSON.parse(p.settings) : (p.settings || {});
        if (ps.linkedProjectId === projectId) {
          children.push({ id: p.id, name: p.name, status: p.status });
        }
      } catch (e) {}
    });
    var parent = null;
    if (parentId) {
      for (var i = 0; i < projects.length; i++) {
        if (projects[i].id === parentId) {
          parent = { id: projects[i].id, name: projects[i].name, status: projects[i].status };
          break;
        }
      }
    }
    return { project: { id: project.id, name: project.name, status: project.status }, parentId: parentId, parent: parent, children: children };
  } catch (error) {
    console.error('getProjectHierarchy failed:', error);
    return null;
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

