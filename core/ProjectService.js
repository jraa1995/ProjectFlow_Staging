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
    const tasks = getAllTasksOptimized();
    var countMap = {};
    for (var i = 0; i < tasks.length; i++) {
      var pid = tasks[i].projectId;
      if (pid) countMap[pid] = (countMap[pid] || 0) + 1;
    }
    for (var j = 0; j < projects.length; j++) {
      projects[j].taskCount = countMap[projects[j].id] || 0;
    }
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
      fields.forEach(function(f) {
        var val = p[f];
        if (!val) {
          try {
            var s = typeof p.settings === 'string' ? JSON.parse(p.settings) : (p.settings || {});
            val = s[f];
          } catch (e) {}
        }
        if (val) {
          var arr = val;
          if (typeof arr === 'string') {
            try { var parsed = JSON.parse(arr); if (Array.isArray(parsed)) arr = parsed; } catch(e) {}
          }
          if (Array.isArray(arr)) {
            arr.forEach(function(v) { if (v && String(v).trim() && v !== 'N/A' && v !== 'FALSE') result[f][String(v).trim()] = true; });
          } else if (String(val).trim() && val !== 'N/A' && val !== 'FALSE') {
            result[f][String(val).trim()] = true;
          }
        }
      });
    });
    var hardcoded = CONFIG.TAXONOMY_OPTIONS || {};
    var sorted = {};
    fields.forEach(function(f) {
      var merged = Object.assign({}, result[f]);
      if (hardcoded[f]) {
        hardcoded[f].forEach(function(opt) { merged[opt] = true; });
      }
      sorted[f] = Object.keys(merged).sort();
    });
    sorted.biSmes = CONFIG.BI_SMES || [];
    sorted.aiInvolved = (CONFIG.TAXONOMY_OPTIONS && CONFIG.TAXONOMY_OPTIONS.aiInvolved) || [];
    sorted.intendedUsers = (CONFIG.TAXONOMY_OPTIONS && CONFIG.TAXONOMY_OPTIONS.intendedUsers) || [];
    sorted.contractOptions = (CONFIG.TAXONOMY_OPTIONS && CONFIG.TAXONOMY_OPTIONS.contractOptions) || [];
    sorted.dataCadence = (CONFIG.TAXONOMY_OPTIONS && CONFIG.TAXONOMY_OPTIONS.dataCadence) || [];
    sorted.projectTypeCode = CONFIG.PROJECT_TYPE_OPTIONS || [];
    return sorted;
  } catch (error) {
    console.error('getProjectTaxonomyValues failed:', error);
    return {};
  }
}

function getProjectPickerOptions() {
  try {
    var projects = getAllProjectsOptimized();
    return projects
      .filter(function(p) { return (p.status || '').toLowerCase() !== 'archived'; })
      .map(function(p) { return { id: p.id, name: p.name }; });
  } catch (error) {
    console.error('getProjectPickerOptions failed:', error);
    return [];
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
    var parentId = project.linkedProjectId || null;
    if (!parentId) {
      try {
        var s = typeof project.settings === 'string' ? JSON.parse(project.settings) : (project.settings || {});
        parentId = s.linkedProjectId || null;
      } catch (e) {}
    }
    var children = [];
    projects.forEach(function(p) {
      if (p.id === projectId) return;
      var pLinked = p.linkedProjectId || null;
      if (!pLinked) {
        try {
          var ps = typeof p.settings === 'string' ? JSON.parse(p.settings) : (p.settings || {});
          pLinked = ps.linkedProjectId || null;
        } catch (e) {}
      }
      if (pLinked === projectId) {
        children.push({ id: p.id, name: p.name, status: p.status });
      }
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

function migrateProjectOwnerNamesToEmails() {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var users = getAllUsers();
    var nameToEmail = {};
    var lowerNameToEmail = {};
    for (var i = 0; i < users.length; i++) {
      var u = users[i];
      if (u.email && u.name) {
        nameToEmail[u.name] = u.email;
        lowerNameToEmail[u.name.toLowerCase().trim()] = u.email;
      }
    }
    var sheet = getProjectsSheet();
    var data = sheet.getDataRange().getValues();
    var columns = data[0].map(function(h) { return String(h).trim(); });
    var ownerCol = columns.indexOf('ownerId');
    var settingsCol = columns.indexOf('settings');
    if (ownerCol === -1) return { success: false, error: 'ownerId column not found' };
    var migrated = [];
    var skipped = [];
    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var projectId = String(row[columns.indexOf('id')] || '');
      var currentOwner = String(row[ownerCol] || '').trim();
      var changed = false;
      if (currentOwner && currentOwner.indexOf('@') === -1 && currentOwner !== '-') {
        var resolved = nameToEmail[currentOwner] || lowerNameToEmail[currentOwner.toLowerCase().trim()];
        if (resolved) {
          data[r][ownerCol] = resolved;
          changed = true;
          migrated.push({ project: projectId, field: 'ownerId', from: currentOwner, to: resolved });
        } else {
          skipped.push({ project: projectId, field: 'ownerId', name: currentOwner });
        }
      }
      if (settingsCol !== -1) {
        var raw = row[settingsCol];
        if (raw) {
          try {
            var settings = typeof raw === 'string' ? JSON.parse(raw) : raw;
            if (Array.isArray(settings.secondaryUsers) && settings.secondaryUsers.length) {
              var updated = false;
              for (var s = 0; s < settings.secondaryUsers.length; s++) {
                var su = String(settings.secondaryUsers[s] || '').trim();
                if (su && su.indexOf('@') === -1) {
                  var resolvedSu = nameToEmail[su] || lowerNameToEmail[su.toLowerCase().trim()];
                  if (resolvedSu) {
                    migrated.push({ project: projectId, field: 'secondaryUser', from: su, to: resolvedSu });
                    settings.secondaryUsers[s] = resolvedSu;
                    updated = true;
                  } else {
                    skipped.push({ project: projectId, field: 'secondaryUser', name: su });
                  }
                }
              }
              if (updated) {
                data[r][settingsCol] = JSON.stringify(settings);
                changed = true;
              }
            }
          } catch (e) {}
        }
      }
    }
    if (migrated.length > 0) {
      sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
      SpreadsheetApp.flush();
      try { invalidateProjectCache(); } catch (e) {}
    }
    return { success: true, migrated: migrated, skipped: skipped };
  } finally {
    lock.releaseLock();
  }
}

function migrateProjectSchema() {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var ss = getColonySpreadsheet_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);
    if (!sheet) return { success: false, error: 'Projects sheet not found' };
    var data = sheet.getDataRange().getValues();
    if (data.length < 1) return { success: false, error: 'Sheet is empty' };
    var oldHeaders = data[0].map(function(h) { return String(h).trim(); });
    var newColumns = CONFIG.PROJECT_COLUMNS;
    var headersMatch = newColumns.length === oldHeaders.length && newColumns.every(function(col, i) { return oldHeaders[i] === col; });
    if (headersMatch) return { success: true, message: 'Schema already up to date', columns: newColumns.length };
    var oldColMap = {};
    for (var c = 0; c < oldHeaders.length; c++) {
      if (oldHeaders[c]) oldColMap[oldHeaders[c]] = c;
    }
    var newData = [newColumns];
    for (var r = 1; r < data.length; r++) {
      var newRow = [];
      for (var n = 0; n < newColumns.length; n++) {
        var col = newColumns[n];
        if (oldColMap.hasOwnProperty(col)) {
          newRow.push(data[r][oldColMap[col]]);
        } else {
          newRow.push('');
        }
      }
      newData.push(newRow);
    }
    sheet.clear();
    sheet.getRange(1, 1, newData.length, newColumns.length).setValues(newData);
    sheet.getRange(1, 1, 1, newColumns.length)
      .setFontWeight('bold')
      .setBackground('#1e293b')
      .setFontColor('white');
    sheet.setFrozenRows(1);
    SpreadsheetApp.flush();
    try { invalidateProjectCache(); } catch (e) {}
    try { RowIndexCache.invalidateSheet(CONFIG.SHEETS.PROJECTS); } catch (e) {}
    var added = newColumns.filter(function(col) { return !oldColMap.hasOwnProperty(col); });
    var removed = oldHeaders.filter(function(col) { return col && newColumns.indexOf(col) === -1; });
    return { success: true, oldColumnCount: oldHeaders.length, newColumnCount: newColumns.length, columnsAdded: added, columnsRemoved: removed };
  } finally {
    lock.releaseLock();
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

function getStoredCpmdWorkbookId() {
  try {
    return { success: true, workbookId: getCpmdWorkbookId() };
  } catch (error) {
    console.error('getStoredCpmdWorkbookId failed:', error);
    return { success: false, workbookId: '' };
  }
}

function saveCpmdWorkbookId(workbookId) {
  try {
    PermissionGuard.requirePermission('admin:settings');
    setCpmdWorkbookId(workbookId);
    return { success: true };
  } catch (error) {
    console.error('saveCpmdWorkbookId failed:', error);
    return { success: false, error: error.message };
  }
}

function loadCpmdPersonnel() {
  try {
    return { success: true, personnel: getCpmdPersonnel() };
  } catch (error) {
    console.error('loadCpmdPersonnel failed:', error);
    return { success: false, personnel: [] };
  }
}

function generateProjectCode(projectId) {
  try {
    var projects = getAllProjectsOptimized();
    var project = projects.find(function(p) { return p.id === projectId; });
    if (!project) return { success: false, error: 'Project not found' };
    var settings = {};
    try { settings = typeof project.settings === 'string' ? JSON.parse(project.settings) : (project.settings || {}); } catch(e) {}

    var missing = [];
    var startDate = project.startDate || settings.govRequestDate;
    var d = startDate ? new Date(startDate) : null;
    if (!startDate || startDate === CONFIG.TBD_DATE_SENTINEL || !d || isNaN(d.getTime())) {
      missing.push('startDate');
    }

    var workstream = settings.workstream || project.workstream || '';
    var wsCode = CONFIG.PROJECT_ID_WORKSTREAM_CODES[workstream];
    if (!wsCode) missing.push('workstream');

    var projectType = settings.projectType || project.projectType || '';
    var ptCode = CONFIG.PROJECT_ID_TYPE_CODES[projectType];
    if (!ptCode) missing.push('projectType');

    var contract = settings.contractInitial || '';
    var cCode = CONFIG.PROJECT_ID_CONTRACT_CODES[contract];
    if (!cCode) missing.push('contractInitial');

    if (missing.length > 0) {
      return {
        success: false,
        missing: missing,
        error: 'Missing required fields: ' + missing.join(', ')
      };
    }

    var month = d.getMonth();
    var year = d.getFullYear();
    var fy = month >= 9 ? (year + 1) % 100 : year % 100;
    var fyStr = String(fy).padStart(2, '0');

    var prefix = fyStr + '-' + wsCode + ptCode + '-' + cCode;
    var maxSeq = 0;
    projects.forEach(function(p) {
      var s = {};
      try { s = typeof p.settings === 'string' ? JSON.parse(p.settings) : (p.settings || {}); } catch(e) {}
      var code = s.projectCode || '';
      if (code && code.indexOf(prefix) === 0) {
        var seqPart = code.substring(prefix.length);
        var seqNum = parseInt(seqPart, 10);
        if (!isNaN(seqNum) && seqNum > maxSeq) maxSeq = seqNum;
      }
    });
    var nextSeq = String(maxSeq + 1).padStart(3, '0');
    var projectCode = prefix + nextSeq;

    var settingsObj = typeof project.settings === 'string' ? JSON.parse(project.settings) : (project.settings || {});
    settingsObj.projectCode = projectCode;
    var settingsStr = JSON.stringify(settingsObj);
    updateProject(project.id, { settings: settingsStr });

    return { success: true, projectCode: projectCode };
  } catch (error) {
    console.error('generateProjectCode failed:', error);
    return { success: false, error: error.message };
  }
}

function sendReleaseNotesEmail(projectId, payload) {
  try {
    PermissionGuard.requirePermission('project:update', { projectId: projectId });
    var project = getProjectById(projectId);
    if (!project) return { success: false, error: 'Project not found' };

    var recipients = payload.recipients;
    if (!recipients || !recipients.length) return { success: false, error: 'No recipients selected' };

    var releases = [];
    try { releases = typeof project.releaseNotes === 'string' ? JSON.parse(project.releaseNotes) : (project.releaseNotes || []); } catch (e) { releases = []; }
    if (!releases.length) return { success: false, error: 'No release notes to send' };

    var selectedVersion = payload.version || '';
    var release = selectedVersion ? releases.find(function(r) { return r.version === selectedVersion; }) : releases[0];
    if (!release) return { success: false, error: 'Release version not found' };

    var introText = payload.introText || '';
    var screenshotUrl = payload.screenshotUrl || '';
    var senderName = payload.senderName || 'COLONY';

    var sectionLabels = { added: 'Added', changed: 'Changed', fixed: 'Fixed', removed: 'Removed', security: 'Security' };
    var sectionsHtml = '';
    if (release.sections) {
      var sectionColors = { added: '#10B981', changed: '#3B82F6', fixed: '#8B5CF6', removed: '#EF4444', security: '#F59E0B' };
      Object.keys(sectionLabels).forEach(function(key) {
        if (release.sections[key] && release.sections[key].length) {
          var color = sectionColors[key] || '#6B7280';
          sectionsHtml += '<tr><td style="padding: 12px 0 4px 0;"><h3 style="margin: 0; font-size: 13px; font-weight: 600; color: ' + color + '; text-transform: uppercase; letter-spacing: 0.05em;">' + sectionLabels[key] + '</h3></td></tr>';
          release.sections[key].forEach(function(item) {
            sectionsHtml += '<tr><td style="padding: 2px 0 2px 16px; font-size: 14px; color: #374151; line-height: 1.5;">&#8226; ' + sanitize(item) + '</td></tr>';
          });
        }
      });
    } else if (release.notes) {
      sectionsHtml = '<tr><td style="padding: 8px 0; font-size: 14px; color: #374151; line-height: 1.6;">' + sanitize(release.notes) + '</td></tr>';
    }

    var typeLabel = release.type ? release.type.charAt(0).toUpperCase() + release.type.slice(1) : '';
    var typeBg = release.type === 'major' ? '#FEE2E2' : release.type === 'minor' ? '#DBEAFE' : '#F3F4F6';
    var typeColor = release.type === 'major' ? '#991B1B' : release.type === 'minor' ? '#1E40AF' : '#4B5563';
    var dateStr = release.date ? new Date(release.date).toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' }) : '';

    var inlineImages = {};
    var screenshotBlock = '';
    if (screenshotUrl) {
      var driveFileId = extractDriveFileId_(screenshotUrl);
      if (driveFileId) {
        try {
          var file = DriveApp.getFileById(driveFileId);
          var blob = file.getBlob();
          var mimeType = blob.getContentType();
          if (mimeType && mimeType.indexOf('image') === 0) {
            inlineImages.screenshot = blob;
            screenshotBlock = '<tr><td style="padding: 16px 0;">' +
              '<img src="cid:screenshot" alt="Release preview" style="max-width: 100%; border-radius: 8px; border: 1px solid #E5E7EB;" />' +
              '</td></tr>';
          }
        } catch (e) {
          console.error('sendReleaseNotesEmail: Drive file fetch failed:', e);
        }
      }
      if (!screenshotBlock) {
        screenshotBlock = '<tr><td style="padding: 16px 0;">' +
          '<img src="' + screenshotUrl + '" alt="Release preview" style="max-width: 100%; border-radius: 8px; border: 1px solid #E5E7EB;" />' +
          '</td></tr>';
      }
    }

    var introBlock = '';
    if (introText) {
      introBlock = '<tr><td style="padding: 0 0 16px 0; font-size: 14px; color: #374151; line-height: 1.6;">' + sanitize(introText).replace(/\n/g, '<br>') + '</td></tr>';
    }

    var htmlBody = '<div style="font-family: -apple-system, BlinkMacSystemFont, \'Segoe UI\', sans-serif; max-width: 600px; margin: 0 auto; background: #ffffff;">' +
      '<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse: collapse;">' +
      '<tr><td style="background: #1e293b; padding: 24px 28px; border-radius: 8px 8px 0 0;">' +
      '<h1 style="margin: 0; font-size: 18px; font-weight: 600; color: #ffffff;">COLONY</h1>' +
      '<p style="margin: 4px 0 0 0; font-size: 13px; color: #94a3b8;">Release Notes</p></td></tr>' +
      '<tr><td style="padding: 24px 28px; border: 1px solid #E5E7EB; border-top: none;">' +
      '<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse: collapse;">' +
      introBlock +
      '<tr><td style="padding: 0 0 16px 0;">' +
      '<table cellpadding="0" cellspacing="0" style="border-collapse: collapse;"><tr>' +
      '<td style="background: #1e293b; color: #ffffff; font-size: 13px; font-weight: 600; padding: 4px 12px; border-radius: 20px; font-family: monospace;">' + sanitize(release.version || '') + '</td>' +
      (typeLabel ? '<td style="padding-left: 8px;"><span style="background: ' + typeBg + '; color: ' + typeColor + '; font-size: 12px; padding: 4px 10px; border-radius: 20px; font-weight: 500;">' + typeLabel + '</span></td>' : '') +
      '<td style="padding-left: 12px; font-size: 12px; color: #9CA3AF;">' + dateStr + '</td>' +
      '</tr></table></td></tr>' +
      (release.summary ? '<tr><td style="padding: 0 0 12px 0; font-size: 15px; font-weight: 500; color: #111827;">' + sanitize(release.summary) + '</td></tr>' : '') +
      '<tr><td style="border-top: 1px solid #F3F4F6;"></td></tr>' +
      sectionsHtml +
      screenshotBlock +
      '</table></td></tr>' +
      '<tr><td style="padding: 16px 28px; background: #F9FAFB; border: 1px solid #E5E7EB; border-top: none; border-radius: 0 0 8px 8px; text-align: center;">' +
      '<p style="margin: 0; font-size: 12px; color: #9CA3AF;">' + sanitize(project.name || '') + ' &mdash; Sent via COLONY</p>' +
      '</td></tr></table></div>';

    var textBody = (project.name || 'Project') + ' - Release ' + (release.version || '') + '\n\n';
    if (introText) textBody += introText + '\n\n';
    if (release.summary) textBody += release.summary + '\n\n';
    if (release.sections) {
      Object.keys(sectionLabels).forEach(function(key) {
        if (release.sections[key] && release.sections[key].length) {
          textBody += sectionLabels[key].toUpperCase() + '\n';
          release.sections[key].forEach(function(item) { textBody += '  - ' + item + '\n'; });
          textBody += '\n';
        }
      });
    }

    var subject = (project.name || 'Project') + ' — Release ' + (release.version || '') + (release.summary ? ': ' + release.summary : '');
    if (subject.length > 120) subject = subject.substring(0, 117) + '...';

    var sent = [];
    var failed = [];
    var emailOptions = { htmlBody: htmlBody, name: senderName };
    if (Object.keys(inlineImages).length) emailOptions.inlineImages = inlineImages;

    recipients.forEach(function(email) {
      try {
        GmailApp.sendEmail(email, subject, textBody, emailOptions);
        sent.push(email);
      } catch (e) {
        console.error('sendReleaseNotesEmail failed for ' + email + ':', e);
        failed.push({ email: email, error: e.message });
      }
    });

    return { success: true, sent: sent, failed: failed, version: release.version };
  } catch (error) {
    console.error('sendReleaseNotesEmail failed:', error);
    return { success: false, error: error.message };
  }
}

function extractDriveFileId_(url) {
  if (!url) return null;
  var patterns = [
    /\/file\/d\/([a-zA-Z0-9_-]+)/,
    /[?&]id=([a-zA-Z0-9_-]+)/,
    /\/open\?id=([a-zA-Z0-9_-]+)/,
    /^([a-zA-Z0-9_-]{25,})$/
  ];
  for (var i = 0; i < patterns.length; i++) {
    var match = url.match(patterns[i]);
    if (match) return match[1];
  }
  return null;
}

function migrateProjectTaxonomyToColumns() {
  PermissionGuard.requirePermission('admin:settings');
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getProjectsSheet();
    var data = sheet.getDataRange().getValues();
    var columns = data[0].map(function(h) { return String(h).trim(); });
    var settingsIdx = columns.indexOf('settings');
    var fields = CONFIG.TAXONOMY_FIELDS.concat(['linkedProjectId']);
    var fieldIndices = {};
    fields.forEach(function(f) { fieldIndices[f] = columns.indexOf(f); });
    var updated = 0;
    for (var i = 1; i < data.length; i++) {
      var raw = data[i][settingsIdx];
      if (!raw) continue;
      var settings;
      try { settings = typeof raw === 'string' ? JSON.parse(raw) : raw; } catch (e) { continue; }
      var changed = false;
      fields.forEach(function(f) {
        var idx = fieldIndices[f];
        if (idx !== -1 && settings[f] && !data[i][idx]) {
          data[i][idx] = String(settings[f]);
          changed = true;
        }
      });
      if (changed) updated++;
    }
    if (updated > 0) {
      sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
      SpreadsheetApp.flush();
    }
    return { migrated: updated, total: data.length - 1 };
  } finally {
    lock.releaseLock();
  }
}

function exportProjectsAndDataAssetsCsv() {
  try {
    PermissionGuard.requirePermission('admin:settings');
    var projects = getAllProjectsOptimized();
    var dataAssets = [];
    try { dataAssets = getAllDataAssetsOptimized(); } catch (e) { dataAssets = []; }

    var projectSettingsKeys = {};
    projects.forEach(function(p) {
      try {
        var s = typeof p.settings === 'string' ? JSON.parse(p.settings) : (p.settings || {});
        Object.keys(s || {}).forEach(function(k) { projectSettingsKeys[k] = true; });
      } catch (e) {}
    });
    var settingsCols = Object.keys(projectSettingsKeys).sort();

    var projCols = ['id', 'name', 'description', 'status', 'ownerId', 'startDate', 'endDate',
      'createdAt', 'updatedAt', 'version', 'repoUrl', 'tags',
      'workstream', 'projectCategory', 'projectType', 'developmentPriority',
      'developmentPhase', 'techStack', 'sdmSupported', 'deploymentLocation',
      'linkedProjectId', 'lastUpdatedBy'];

    var projectHeaders = projCols.concat(settingsCols.map(function(k) { return 'settings.' + k; }));
    var projectRows = [projectHeaders];
    projects.forEach(function(p) {
      var settings = {};
      try { settings = typeof p.settings === 'string' ? JSON.parse(p.settings) : (p.settings || {}); } catch (e) {}
      var row = projCols.map(function(c) { return csvSafe_(p[c]); });
      settingsCols.forEach(function(k) { row.push(csvSafe_(settings[k])); });
      projectRows.push(row);
    });

    var daCols = CONFIG.DATA_ASSET_COLUMNS.filter(function(c) { return c !== 'jsonData'; });
    var daRows = [daCols];
    dataAssets.forEach(function(da) {
      daRows.push(daCols.map(function(c) { return csvSafe_(da[c]); }));
    });

    var projectCsv = projectRows.map(function(r) { return r.join(','); }).join('\n');
    var dataAssetCsv = daRows.map(function(r) { return r.join(','); }).join('\n');

    return {
      success: true,
      generatedAt: new Date().toISOString(),
      projectCount: projects.length,
      dataAssetCount: dataAssets.length,
      projectCsv: projectCsv,
      dataAssetCsv: dataAssetCsv
    };
  } catch (error) {
    console.error('exportProjectsAndDataAssetsCsv failed:', error);
    return { success: false, error: error.message };
  }
}

function csvSafe_(val) {
  if (val === null || val === undefined) return '';
  var str;
  if (typeof val === 'object') {
    try { str = JSON.stringify(val); } catch (e) { str = String(val); }
  } else {
    str = String(val);
  }
  if (str.indexOf(',') !== -1 || str.indexOf('"') !== -1 || str.indexOf('\n') !== -1 || str.indexOf('\r') !== -1) {
    return '"' + str.replace(/"/g, '""') + '"';
  }
  return str;
}

function repairProjectColumnAlignment() {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sheet = getProjectsSheet();
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h).trim(); });
    var wsIdx = headers.indexOf('workstream');
    var settingsIdx = headers.indexOf('settings');
    var jsonDataIdx = headers.indexOf('jsonData');

    if (wsIdx === -1 || settingsIdx === -1) return { success: false, error: 'Required columns not found' };

    var repaired = 0;
    for (var i = 1; i < data.length; i++) {
      var wsVal = String(data[i][wsIdx] || '');
      if (!wsVal || wsVal.charAt(0) !== '{') continue;
      try {
        var parsed = JSON.parse(wsVal);
        if (!parsed.id || !parsed.name) continue;
      } catch (e) { continue; }

      var settingsRaw = data[i][settingsIdx];
      var settings = {};
      try { settings = typeof settingsRaw === 'string' ? JSON.parse(settingsRaw) : (settingsRaw || {}); } catch (e) {}

      CONFIG.TAXONOMY_FIELDS.concat(['linkedProjectId']).forEach(function(f) {
        var idx = headers.indexOf(f);
        if (idx !== -1) {
          data[i][idx] = settings[f] ? String(settings[f]) : '';
        }
      });

      if (jsonDataIdx !== -1) {
        var obj = rowToObject(data[i], headers);
        var cleanObj = {};
        headers.forEach(function(col) { if (col !== 'jsonData') cleanObj[col] = obj[col] !== undefined ? obj[col] : ''; });
        data[i][jsonDataIdx] = JSON.stringify(cleanObj);
      }

      repaired++;
    }

    if (repaired > 0) {
      sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
      SpreadsheetApp.flush();
      try { invalidateProjectCache(); } catch (e) {}
    }

    return { success: true, repaired: repaired, total: data.length - 1 };
  } finally {
    lock.releaseLock();
  }
}

