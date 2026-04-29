function saveNewProject(projectData) {
  try {
    const project = createProject(projectData);
    invalidateProjectCache();
    return { success: true, project: project };
  } catch (error) {
    console.error('saveNewProject failed:', error);
    return {
      success: false,
      error: error.message,
      code: error.code || null,
      suggestedType: error.suggestedType || null
    };
  }
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
    var assetCountMap = {};
    try {
      if (typeof getAllDataAssetsOptimized === 'function') {
        var assets = getAllDataAssetsOptimized();
        for (var a = 0; a < assets.length; a++) {
          var ids = String(assets[a].relatedProjects || '').split(/[,;]/).map(function(s) { return s.trim(); }).filter(Boolean);
          ids.forEach(function(id) { assetCountMap[id] = (assetCountMap[id] || 0) + 1; });
        }
      }
    } catch (e) {}
    for (var j = 0; j < projects.length; j++) {
      var p = projects[j];
      p.taskCount = countMap[p.id] || 0;
      var count = assetCountMap[p.id] || 0;
      var parsedSettings = _parseProjectSettings_(p);
      var compliance = computeProjectCompliance(p, count, parsedSettings);
      p.complianceRating = compliance.rating;
      p.complianceScore = compliance.score;
      p.complianceFilled = compliance.filled;
      p.complianceTotal = compliance.total;
      p.origin = _resolveProjectOrigin_(p, parsedSettings);
      p.originSource = parsedSettings.originSource || 'auto';
    }
    var userRole = (typeof getCurrentUserRole === 'function') ? getCurrentUserRole() : null;
    if (userRole === 'client') {
      var userEmail = (typeof getCurrentUserEmailOptimized === 'function') ? getCurrentUserEmailOptimized() : null;
      projects = filterProjectsByUserRole(projects, [], userEmail, userRole);
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
    var userRole = (typeof getCurrentUserRole === 'function') ? getCurrentUserRole() : null;
    if (userRole === 'client') {
      var userEmail = (typeof getCurrentUserEmailOptimized === 'function') ? getCurrentUserEmailOptimized() : null;
      var allowed = getClientAccessibleProjectIds(userEmail, [project]);
      if (!allowed[project.id]) {
        return { success: false, error: 'Project not found' };
      }
    }
    var count = _countRelatedDataAssets_(projectId);
    var compliance = computeProjectCompliance(project, count);
    project.complianceRating = compliance.rating;
    project.complianceScore = compliance.score;
    project.complianceFilled = compliance.filled;
    project.complianceTotal = compliance.total;
    _decorateProjectOrigin_(project);
    return { success: true, project: project, compliance: compliance };
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
    var cached = null;
    try { cached = VersionedCache.get('PROJECT_TAXONOMY_CACHE'); } catch (e) {}
    if (cached) return cached;
    var projects = getAllProjectsOptimized();
    var fields = CONFIG.TAXONOMY_FIELDS;
    var MULTIVALUE_FIELDS = { techStack: true, deploymentLocation: true };
    var result = {};
    fields.forEach(function(f) { result[f] = {}; });
    projects.forEach(function(p) {
      var parsedSettings = null;
      fields.forEach(function(f) {
        var val = p[f];
        if (!val) {
          if (parsedSettings === null) {
            try {
              parsedSettings = typeof p.settings === 'string' ? JSON.parse(p.settings || '{}') : (p.settings || {});
            } catch (e) { parsedSettings = {}; }
          }
          val = parsedSettings[f];
        }
        if (val) {
          var arr = val;
          if (typeof arr === 'string') {
            try { var parsed = JSON.parse(arr); if (Array.isArray(parsed)) arr = parsed; } catch(e) {}
          }
          if (typeof arr === 'string' && MULTIVALUE_FIELDS[f] && arr.indexOf(',') !== -1) {
            arr = arr.split(',');
          }
          if (Array.isArray(arr)) {
            arr.forEach(function(v) {
              var t = String(v || '').trim();
              if (t && t !== 'N/A' && t !== 'FALSE') result[f][t] = true;
            });
          } else if (String(val).trim() && val !== 'N/A' && val !== 'FALSE') {
            result[f][String(val).trim()] = true;
          }
        }
      });
    });
    var hardcoded = CONFIG.TAXONOMY_OPTIONS || {};
    var assetOnlyTypes = CONFIG.DATA_ASSET_PROJECT_TYPES || [];
    var sorted = {};
    fields.forEach(function(f) {
      var merged = Object.assign({}, result[f]);
      if (hardcoded[f]) {
        hardcoded[f].forEach(function(opt) { merged[opt] = true; });
      }
      var keys = Object.keys(merged).sort();
      if (f === 'projectType' && assetOnlyTypes.length) {
        keys = keys.filter(function(v) { return assetOnlyTypes.indexOf(v) === -1; });
      }
      sorted[f] = keys;
    });
    sorted.biSmes = CONFIG.BI_SMES || [];
    sorted.aiInvolved = (CONFIG.TAXONOMY_OPTIONS && CONFIG.TAXONOMY_OPTIONS.aiInvolved) || [];
    sorted.intendedUsers = (CONFIG.TAXONOMY_OPTIONS && CONFIG.TAXONOMY_OPTIONS.intendedUsers) || [];
    sorted.contractOptions = (CONFIG.TAXONOMY_OPTIONS && CONFIG.TAXONOMY_OPTIONS.contractOptions) || [];
    sorted.dataCadence = (CONFIG.TAXONOMY_OPTIONS && CONFIG.TAXONOMY_OPTIONS.dataCadence) || [];
    sorted.projectTypeCode = CONFIG.PROJECT_TYPE_OPTIONS || [];
    try { VersionedCache.put('PROJECT_TAXONOMY_CACHE', sorted, 900); } catch (e) {}
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

function _isEmailLike_(value) {
  var s = String(value || '').trim();
  if (!s) return false;
  return s.indexOf('@') !== -1 && s.lastIndexOf('.') > s.indexOf('@');
}

function _normalizeOwnerKey_(value) {
  return String(value || '').trim().toLowerCase().replace(/\./g, ' ').replace(/\s+/g, ' ');
}

function _buildUserOwnerIndex_(users) {
  var byKey = {};
  (users || []).forEach(function(u) {
    if (!u || !u.email) return;
    var keys = [];
    if (u.name) keys.push(_normalizeOwnerKey_(u.name));
    var prefix = String(u.email).split('@')[0];
    if (prefix) keys.push(_normalizeOwnerKey_(prefix));
    keys.forEach(function(k) {
      if (!k) return;
      if (!byKey[k]) byKey[k] = [];
      if (byKey[k].indexOf(u.email) === -1) byKey[k].push(u.email);
    });
  });
  return byKey;
}

function _resolveOwnerToEmail_(rawValue, ownerIndex) {
  if (_isEmailLike_(rawValue)) return { email: String(rawValue).trim().toLowerCase(), match: 'email' };
  var key = _normalizeOwnerKey_(rawValue);
  if (!key) return { email: '', match: 'empty' };
  var matches = ownerIndex && ownerIndex[key];
  if (!matches || matches.length === 0) return { email: '', match: 'unmatched' };
  if (matches.length === 1) return { email: matches[0], match: 'matched' };
  return { email: '', match: 'ambiguous', candidates: matches.slice() };
}

function _userIdentityKeys_(email, name) {
  var keys = {};
  var emailLc = String(email || '').trim().toLowerCase();
  if (emailLc) keys[emailLc] = true;
  var nameKey = _normalizeOwnerKey_(name);
  if (nameKey) keys[nameKey] = true;
  if (emailLc.indexOf('@') !== -1) {
    var prefixKey = _normalizeOwnerKey_(emailLc.split('@')[0]);
    if (prefixKey) keys[prefixKey] = true;
  }
  return keys;
}

function _safeClaimForUser_(rawValue, emailLc, identity, ownerIndex) {
  if (!rawValue) return false;
  var s = String(rawValue).trim();
  if (!s) return false;
  if (s.toLowerCase() === emailLc) return false;
  var k = _normalizeOwnerKey_(s);
  if (!k || !identity[k]) return false;
  var owners = ownerIndex && ownerIndex[k];
  if (!owners || owners.length === 0) return true;
  if (owners.length === 1 && owners[0] === emailLc) return true;
  return false;
}

function backfillProjectOwnersForUser(email, name) {
  var emailLc = String(email || '').trim().toLowerCase();
  if (!emailLc) return { scanned: 0, updated: [] };
  var callerEmail = String(getCurrentUserEmail() || '').toLowerCase();
  var callerUser = callerEmail ? getUserByEmail(callerEmail) : null;
  var isAdmin = callerUser && callerUser.role === 'admin';
  var isSelf = callerEmail && callerEmail === emailLc;
  if (!isAdmin && !isSelf) {
    throw new Error('Permission denied: backfillProjectOwnersForUser');
  }
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return { scanned: 0, updated: [], error: 'Could not acquire lock for owner backfill' };
  }
  try {
    var identity = _userIdentityKeys_(emailLc, name);
    var users = (typeof getAllUsers === 'function') ? getAllUsers() : [];
    var ownerIndex = _buildUserOwnerIndex_(users);
    var projects = getAllProjectsOptimized();
    var updated = [];
    var skippedAmbiguous = [];

    projects.forEach(function(p) {
      if (!p) return;
      var settings = {};
      try { settings = typeof p.settings === 'string' ? JSON.parse(p.settings) : (p.settings || {}); }
      catch (e) { settings = {}; }

      var changes = {};
      var settingsChanged = false;
      var changeLog = [];
      var ambiguous = [];

      var rawOwner = String(p.ownerId || '').trim();
      var ownerIsEmail = _isEmailLike_(rawOwner);
      if (!ownerIsEmail && rawOwner) {
        if (_safeClaimForUser_(rawOwner, emailLc, identity, ownerIndex)) {
          changes.ownerId = emailLc;
          changeLog.push({ field: 'ownerId', from: rawOwner, to: emailLc });
        } else if (identity[_normalizeOwnerKey_(rawOwner)]) {
          ambiguous.push({ field: 'ownerId', value: rawOwner });
        }
      }
      if (!changes.ownerId && settings.pendingOwnerName) {
        if (_safeClaimForUser_(settings.pendingOwnerName, emailLc, identity, ownerIndex)) {
          changes.ownerId = emailLc;
          changeLog.push({ field: 'ownerId', from: 'pending:' + settings.pendingOwnerName, to: emailLc });
          delete settings.pendingOwnerName;
          settingsChanged = true;
        } else if (identity[_normalizeOwnerKey_(settings.pendingOwnerName)]) {
          ambiguous.push({ field: 'pendingOwnerName', value: settings.pendingOwnerName });
        }
      }

      var secondary = Array.isArray(settings.secondaryUsers) ? settings.secondaryUsers.slice() : [];
      var pendingSecondary = Array.isArray(settings.pendingSecondaryUsers) ? settings.pendingSecondaryUsers.slice() : [];
      var alreadyHasEmail = secondary.some(function(v) { return String(v || '').trim().toLowerCase() === emailLc; });
      var nextSecondary = secondary.filter(function(v) {
        var sv = String(v || '').trim();
        if (!sv) return false;
        if (_isEmailLike_(sv)) return true;
        if (_safeClaimForUser_(sv, emailLc, identity, ownerIndex)) {
          changeLog.push({ field: 'secondaryUsers', from: sv, to: emailLc });
          return false;
        }
        if (identity[_normalizeOwnerKey_(sv)]) ambiguous.push({ field: 'secondaryUsers', value: sv });
        return true;
      });
      var nextPending = pendingSecondary.filter(function(v) {
        if (_safeClaimForUser_(v, emailLc, identity, ownerIndex)) {
          changeLog.push({ field: 'pendingSecondaryUsers', from: v, to: emailLc });
          return false;
        }
        if (identity[_normalizeOwnerKey_(v)]) ambiguous.push({ field: 'pendingSecondaryUsers', value: v });
        return true;
      });
      var matchedInSecondary = nextSecondary.length !== secondary.length;
      var matchedInPending = nextPending.length !== pendingSecondary.length;
      if ((matchedInSecondary || matchedInPending) && !alreadyHasEmail) {
        nextSecondary.push(emailLc);
      }
      if (matchedInSecondary || matchedInPending) {
        settings.secondaryUsers = nextSecondary;
        if (Array.isArray(settings.pendingSecondaryUsers)) {
          if (nextPending.length === 0) delete settings.pendingSecondaryUsers;
          else settings.pendingSecondaryUsers = nextPending;
        }
        settingsChanged = true;
      }

      if (ambiguous.length) {
        skippedAmbiguous.push({ projectId: p.id, projectName: p.name || '', conflicts: ambiguous });
      }

      if (settingsChanged) changes.settings = JSON.stringify(settings);
      if (Object.keys(changes).length === 0) return;
      try {
        updateProject(p.id, changes);
        updated.push({ projectId: p.id, projectName: p.name || '', changes: changeLog });
      } catch (e) {
        console.error('backfillProjectOwnersForUser failed for ' + p.id + ':', e);
      }
    });

    if (updated.length) {
      invalidateProjectCache();
      try { logActivity(emailLc, 'owner_backfilled', 'admin', '', { count: updated.length, ambiguous: skippedAmbiguous.length }); } catch (e) {}
    }
    return { scanned: projects.length, updated: updated, skippedAmbiguous: skippedAmbiguous };
  } finally {
    lock.releaseLock();
  }
}

function linkProjectOwner(projectId, email) {
  try {
    PermissionGuard.requirePermission('admin:settings');
    if (!projectId) return { success: false, error: 'Project ID is required' };
    var project = getProjectById(projectId);
    if (!project) return { success: false, error: 'Project not found' };
    var emailLc = String(email || '').trim().toLowerCase();
    var settings = {};
    try { settings = typeof project.settings === 'string' ? JSON.parse(project.settings) : (project.settings || {}); } catch (e) { settings = {}; }
    var updates = {};
    if (emailLc) {
      if (!_isEmailLike_(emailLc)) return { success: false, error: 'A valid email is required' };
      var user = getUserByEmail(emailLc);
      if (!user) return { success: false, error: 'No user found for ' + emailLc };
      updates.ownerId = emailLc;
      if (settings.pendingOwnerName) {
        delete settings.pendingOwnerName;
        updates.settings = JSON.stringify(settings);
      }
    } else {
      updates.ownerId = '';
    }
    updateProject(projectId, updates);
    invalidateProjectCache(projectId);
    try { logActivity(getCurrentUserEmail(), 'owner_linked', 'project', projectId, { to: emailLc || '' }); } catch (e) {}
    return { success: true, projectId: projectId, ownerId: emailLc };
  } catch (error) {
    console.error('linkProjectOwner failed:', error);
    return { success: false, error: error.message };
  }
}

function normalizeAllProjectOwners(options) {
  try {
    PermissionGuard.requirePermission('admin:settings');
    options = options || {};
    var apply = options.apply === true;
    var lock = LockService.getScriptLock();
    if (apply && !lock.tryLock(60000)) {
      return { success: false, error: 'Could not acquire lock for normalization' };
    }
    try {
      var users = (typeof getAllUsers === 'function') ? getAllUsers() : [];
      var ownerIndex = _buildUserOwnerIndex_(users);
      var emailLookup = {};
      users.forEach(function(u) { if (u && u.email) emailLookup[String(u.email).toLowerCase()] = true; });
      var projects = getAllProjectsOptimized();
      var report = {
        scanned: projects.length,
        rewrittenOwner: 0,
        rewrittenSecondary: 0,
        movedToPendingOwner: 0,
        movedToPendingSecondary: 0,
        ambiguousOwner: [],
        unresolvedOwner: [],
        applied: [],
        apply: apply,
        timestamp: now()
      };
      var pendingChanges = [];

      projects.forEach(function(p) {
        if (!p) return;
        var settings = {};
        try { settings = typeof p.settings === 'string' ? JSON.parse(p.settings) : (p.settings || {}); }
        catch (e) { settings = {}; }
        var changes = {};
        var settingsChanged = false;
        var detail = { projectId: p.id, projectName: p.name || '', changes: [] };

        var rawOwner = String(p.ownerId || '').trim();
        if (rawOwner && !_isEmailLike_(rawOwner)) {
          var resolved = _resolveOwnerToEmail_(rawOwner, ownerIndex);
          if (resolved.match === 'matched') {
            changes.ownerId = resolved.email;
            detail.changes.push({ field: 'ownerId', from: rawOwner, to: resolved.email });
            report.rewrittenOwner++;
          } else if (resolved.match === 'ambiguous') {
            report.ambiguousOwner.push({ projectId: p.id, projectName: p.name || '', currentOwnerId: rawOwner, candidates: resolved.candidates || [] });
            if (settings.pendingOwnerName !== rawOwner) {
              settings.pendingOwnerName = rawOwner;
              changes.ownerId = '';
              detail.changes.push({ field: 'ownerId', from: rawOwner, to: 'pending:' + rawOwner });
              report.movedToPendingOwner++;
              settingsChanged = true;
            }
          } else {
            report.unresolvedOwner.push({ projectId: p.id, projectName: p.name || '', currentOwnerId: rawOwner });
            if (settings.pendingOwnerName !== rawOwner) {
              settings.pendingOwnerName = rawOwner;
              changes.ownerId = '';
              detail.changes.push({ field: 'ownerId', from: rawOwner, to: 'pending:' + rawOwner });
              report.movedToPendingOwner++;
              settingsChanged = true;
            }
          }
        } else if (rawOwner && _isEmailLike_(rawOwner)) {
          var lc = rawOwner.toLowerCase();
          if (rawOwner !== lc) {
            changes.ownerId = lc;
            detail.changes.push({ field: 'ownerId', from: rawOwner, to: lc });
          }
        }

        var secondary = Array.isArray(settings.secondaryUsers) ? settings.secondaryUsers.slice() : [];
        var pendingSecondary = Array.isArray(settings.pendingSecondaryUsers) ? settings.pendingSecondaryUsers.slice() : [];
        var nextSecondary = [];
        var nextPending = pendingSecondary.slice();
        var seen = {};
        var secondaryChanged = false;
        secondary.forEach(function(v) {
          var sv = String(v || '').trim();
          if (!sv) { secondaryChanged = true; return; }
          if (_isEmailLike_(sv)) {
            var lc = sv.toLowerCase();
            if (!seen[lc]) { nextSecondary.push(lc); seen[lc] = true; }
            if (lc !== sv) secondaryChanged = true;
            return;
          }
          var r = _resolveOwnerToEmail_(sv, ownerIndex);
          if (r.match === 'matched') {
            if (!seen[r.email]) { nextSecondary.push(r.email); seen[r.email] = true; }
            detail.changes.push({ field: 'secondaryUsers', from: sv, to: r.email });
            report.rewrittenSecondary++;
            secondaryChanged = true;
          } else {
            if (nextPending.indexOf(sv) === -1) {
              nextPending.push(sv);
              report.movedToPendingSecondary++;
            }
            detail.changes.push({ field: 'secondaryUsers', from: sv, to: 'pending:' + sv });
            secondaryChanged = true;
          }
        });
        if (secondaryChanged) {
          settings.secondaryUsers = nextSecondary;
          if (nextPending.length) settings.pendingSecondaryUsers = nextPending;
          else delete settings.pendingSecondaryUsers;
          settingsChanged = true;
        }

        if (settingsChanged) changes.settings = JSON.stringify(settings);
        if (Object.keys(changes).length === 0) return;
        pendingChanges.push({ projectId: p.id, changes: changes, detail: detail });
      });

      if (apply && pendingChanges.length) {
        pendingChanges.forEach(function(pc) {
          try {
            updateProject(pc.projectId, pc.changes);
            report.applied.push(pc.detail);
          } catch (e) {
            console.error('normalizeAllProjectOwners apply failed for ' + pc.projectId + ':', e);
            report.unresolvedOwner.push({ projectId: pc.projectId, error: e.message });
          }
        });
        invalidateProjectCache();
        try { logActivity(getCurrentUserEmail(), 'normalized_project_owners', 'admin', '', { applied: report.applied.length, scanned: report.scanned }); } catch (e) {}
      } else if (!apply) {
        report.preview = pendingChanges.map(function(pc) { return pc.detail; });
      }

      return { success: true, report: report };
    } finally {
      if (apply) lock.releaseLock();
    }
  } catch (error) {
    console.error('normalizeAllProjectOwners failed:', error);
    return { success: false, error: error.message };
  }
}

function repairProjectOwnerIds(options) {
  try {
    PermissionGuard.requirePermission('admin:settings');
    options = options || {};
    var apply = options.apply === true;
    var users = (typeof getAllUsers === 'function') ? getAllUsers() : [];
    var ownerIndex = _buildUserOwnerIndex_(users);
    var projects = getAllProjectsOptimized();
    var report = {
      scanned: projects.length,
      alreadyEmail: 0,
      empty: 0,
      matched: [],
      ambiguous: [],
      unmatched: [],
      applied: [],
      apply: apply,
      timestamp: now()
    };
    var pending = [];
    projects.forEach(function(p) {
      if (!p) return;
      var raw = p.ownerId || '';
      if (_isEmailLike_(raw)) { report.alreadyEmail++; return; }
      if (!String(raw).trim()) { report.empty++; return; }
      var resolved = _resolveOwnerToEmail_(raw, ownerIndex);
      var entry = { projectId: p.id, projectName: p.name || '', currentOwnerId: raw };
      if (resolved.match === 'matched') {
        entry.suggestedEmail = resolved.email;
        report.matched.push(entry);
        pending.push(entry);
      } else if (resolved.match === 'ambiguous') {
        entry.candidates = resolved.candidates || [];
        report.ambiguous.push(entry);
      } else {
        report.unmatched.push(entry);
      }
    });

    if (apply && pending.length) {
      pending.forEach(function(entry) {
        try {
          updateProject(entry.projectId, { ownerId: entry.suggestedEmail });
          report.applied.push({
            projectId: entry.projectId,
            projectName: entry.projectName,
            from: entry.currentOwnerId,
            to: entry.suggestedEmail
          });
        } catch (e) {
          console.error('repairProjectOwnerIds apply failed for ' + entry.projectId + ':', e);
          report.unmatched.push({
            projectId: entry.projectId,
            projectName: entry.projectName,
            currentOwnerId: entry.currentOwnerId,
            error: e.message
          });
        }
      });
      invalidateProjectCache();
      logActivity(getCurrentUserEmail(), 'repaired_owner_ids', 'admin', '', { applied: report.applied.length, scanned: report.scanned });
    }

    return { success: true, report: report };
  } catch (error) {
    console.error('repairProjectOwnerIds failed:', error);
    return { success: false, error: error.message };
  }
}

function previewProjectMigration(projectId) {
  try {
    PermissionGuard.requirePermission('project:read', { projectId: projectId });
    var project = getProjectById(projectId);
    if (!project) return { success: false, error: 'Project not found' };
    var settings = _parseProjectSettings_(project);
    var projectType = String(settings.projectType || project.projectType || '').trim();
    if ((CONFIG.DATA_ASSET_PROJECT_TYPES || []).indexOf(projectType) === -1) {
      return { success: false, error: 'Project type "' + (projectType || 'unknown') + '" is not a Data Asset type.' };
    }
    var tasks = getAllTasksOptimized().filter(function(t) { return t.projectId === projectId; });
    var openTasks = tasks.filter(function(t) { return String(t.status || '').toLowerCase() !== 'done'; });
    var inboundLinks = getAllProjectsOptimized().filter(function(p) {
      if (p.id === projectId) return false;
      var s = _parseProjectSettings_(p);
      var linked = s.linkedProjectId || s.linkedProjectIds;
      if (!linked) return false;
      if (Array.isArray(linked)) return linked.indexOf(projectId) !== -1;
      try { var parsed = JSON.parse(linked); if (Array.isArray(parsed)) return parsed.indexOf(projectId) !== -1; } catch (e) {}
      return String(linked).split(/[,;]/).map(function(s){return s.trim();}).indexOf(projectId) !== -1;
    }).map(function(p) { return { id: p.id, name: p.name }; });
    return {
      success: true,
      project: { id: project.id, name: project.name, projectType: projectType },
      taskCount: tasks.length,
      openTaskCount: openTasks.length,
      inboundLinkCount: inboundLinks.length,
      inboundLinks: inboundLinks
    };
  } catch (error) {
    console.error('previewProjectMigration failed:', error);
    return { success: false, error: error.message };
  }
}

function migrateProjectToDataAsset(projectId) {
  try {
    PermissionGuard.requirePermission('project:update', { projectId: projectId });
    PermissionGuard.requirePermission('dataasset:create');
    var project = getProjectById(projectId);
    if (!project) return { success: false, error: 'Project not found' };
    var settings = _parseProjectSettings_(project);
    var projectType = String(settings.projectType || project.projectType || '').trim();
    if ((CONFIG.DATA_ASSET_PROJECT_TYPES || []).indexOf(projectType) === -1) {
      return { success: false, error: 'Only Database and Data Pipeline projects can migrate to Data Assets.' };
    }
    var secondaryUsers = settings.secondaryUsers;
    if (typeof secondaryUsers === 'string') {
      try { var parsed = JSON.parse(secondaryUsers); if (Array.isArray(parsed)) secondaryUsers = parsed; } catch (e) {}
    }
    if (Array.isArray(secondaryUsers)) secondaryUsers = secondaryUsers.join(', ');
    var deploymentLoc = settings.deploymentLocation;
    if (Array.isArray(deploymentLoc)) deploymentLoc = deploymentLoc.join(', ');
    else if (typeof deploymentLoc === 'string') {
      try { var dParsed = JSON.parse(deploymentLoc); if (Array.isArray(dParsed)) deploymentLoc = dParsed.join(', '); } catch (e) {}
    }
    var techStack = settings.techStack;
    if (Array.isArray(techStack)) techStack = techStack.join(', ');
    else if (typeof techStack === 'string') {
      try { var tParsed = JSON.parse(techStack); if (Array.isArray(tParsed)) techStack = tParsed.join(', '); } catch (e) {}
    }
    var env = deploymentLoc || techStack || '';
    var assetPayload = {
      status: 'Active',
      assetOwner: project.ownerId || '',
      backupOwner: secondaryUsers || '',
      assetName: project.name || '',
      dataSource: settings.dataSourceExplain || '',
      targetFiles: settings.dataSourceFiles || '',
      relatedProjects: settings.linkedProjectId || settings.linkedProjectIds || '',
      primaryStakeholder: settings.clientOwner || '',
      updateFrequency: settings.dataCadence || '',
      updateSchedule: '',
      automatedSchedule: settings.automatedSchedule || '',
      currentEnvironment: env,
      githubLink: settings.githubLinks || project.repoUrl || '',
      dataSharingDocLink: settings.dataSharingDocLink || settings.googleDriveFolder || '',
      jsonData: JSON.stringify({
        migratedFromProjectId: project.id,
        migratedAt: new Date().toISOString(),
        originalDescription: project.description || '',
        originalProjectType: projectType
      })
    };
    var asset = createDataAsset(assetPayload);
    settings.migratedToAssetId = asset.id;
    settings.migratedAt = new Date().toISOString();
    updateProject(projectId, { status: 'archived', settings: JSON.stringify(settings) });
    invalidateProjectCache();
    invalidateDataAssetCache();
    return { success: true, asset: asset, projectId: projectId };
  } catch (error) {
    console.error('migrateProjectToDataAsset failed:', error);
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

function _maskWorkbookId_(id) {
  if (!id) return '';
  var s = String(id);
  if (s.length <= 10) return s.slice(0, 3) + '…';
  return s.slice(0, 6) + '…' + s.slice(-4);
}

function getStoredFedStaffingConfig() {
  try {
    PermissionGuard.requirePermission('admin:settings');
    var wbId = getFedStaffingWorkbookId();
    var roles = Object.keys(FED_STAFFING_ROLES).map(function(role) {
      return {
        role: role,
        label: FED_STAFFING_ROLES[role].label,
        sheetName: getFedStaffingSheetMap()[role] || '',
        defaultSheet: FED_STAFFING_ROLES[role].defaultSheet
      };
    });
    return {
      success: true,
      workbookId: wbId,
      workbookIdMasked: _maskWorkbookId_(wbId),
      configured: !!wbId,
      roles: roles
    };
  } catch (error) {
    console.error('getStoredFedStaffingConfig failed:', error);
    return { success: false, error: error.message, roles: [] };
  }
}

function saveFedStaffingWorkbookId(workbookId) {
  try {
    PermissionGuard.requirePermission('admin:settings');
    setFedStaffingWorkbookId(workbookId);
    return { success: true };
  } catch (error) {
    console.error('saveFedStaffingWorkbookId failed:', error);
    return { success: false, error: error.message };
  }
}

function saveFedStaffingSheetName(role, sheetName) {
  try {
    PermissionGuard.requirePermission('admin:settings');
    setFedStaffingSheet(role, sheetName);
    return { success: true };
  } catch (error) {
    console.error('saveFedStaffingSheetName failed:', error);
    return { success: false, error: error.message };
  }
}

function loadFedStaffingPersonnel(role) {
  try {
    PermissionGuard.requirePermission('project:read');
    return { success: true, role: role || 'personnel', personnel: getFedStaffingPersonnel(role) };
  } catch (error) {
    console.error('loadFedStaffingPersonnel failed:', error);
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

function _parseProjectSettings_(project) {
  if (!project) return {};
  try {
    return typeof project.settings === 'string' ? (project.settings ? JSON.parse(project.settings) : {}) : (project.settings || {});
  } catch (e) {
    return {};
  }
}

function _resolveProjectOrigin_(project, preparsedSettings) {
  if (!project) return 'internal';
  var settings = preparsedSettings || _parseProjectSettings_(project);
  if (settings && settings.origin) return settings.origin;
  var creator = settings.originCreatorEmail || project.lastUpdatedBy || '';
  if (typeof isInternalEmail === 'function') {
    return isInternalEmail(creator) ? 'internal' : 'client';
  }
  return 'internal';
}

function _decorateProjectOrigin_(project) {
  if (!project) return project;
  var settings = _parseProjectSettings_(project);
  project.origin = _resolveProjectOrigin_(project, settings);
  project.originSource = settings.originSource || 'auto';
  return project;
}

function setProjectOrigin(projectId, origin, reason) {
  try {
    PermissionGuard.requirePermission('project:overrideOrigin');
    if (!projectId) return { success: false, error: 'Project ID required' };
    var allowed = ['internal', 'client'];
    var nextOrigin = String(origin || '').toLowerCase().trim();
    if (allowed.indexOf(nextOrigin) === -1) return { success: false, error: 'Invalid origin: ' + origin };
    var project = getProjectById(projectId);
    if (!project) return { success: false, error: 'Project not found' };
    var settings = _parseProjectSettings_(project);
    var prevOrigin = settings.origin || _resolveProjectOrigin_(project, settings);
    settings.origin = nextOrigin;
    settings.originSource = 'manual';
    settings.originReason = reason ? String(reason).slice(0, 500) : '';
    var updated = updateProject(projectId, { settings: JSON.stringify(settings) });
    invalidateProjectCache(projectId);
    try {
      if (typeof logAuditEvent === 'function') {
        logAuditEvent('project.origin_changed', 'project', projectId, project.name || '', {
          before: prevOrigin,
          after: nextOrigin,
          reason: reason || ''
        });
      }
    } catch (e) {}
    _decorateProjectOrigin_(updated);
    return { success: true, project: updated, origin: nextOrigin };
  } catch (error) {
    console.error('setProjectOrigin failed:', error);
    return { success: false, error: error.message };
  }
}

function bulkReclassifyProjectOrigins() {
  try {
    PermissionGuard.requirePermission('project:overrideOrigin');
    var projects = getAllProjects();
    var changed = 0;
    for (var i = 0; i < projects.length; i++) {
      var p = projects[i];
      if (!p || !p.id) continue;
      var settings = _parseProjectSettings_(p);
      if (settings.originSource === 'manual') continue;
      var creator = settings.originCreatorEmail || p.lastUpdatedBy || '';
      var derived = (typeof isInternalEmail === 'function' && !isInternalEmail(creator)) ? 'client' : 'internal';
      if (settings.origin === derived && settings.originCreatorEmail) continue;
      settings.origin = derived;
      settings.originSource = 'auto';
      if (!settings.originCreatorEmail) settings.originCreatorEmail = creator;
      try {
        updateProject(p.id, { settings: JSON.stringify(settings) });
        changed++;
      } catch (e) {
        console.error('bulkReclassifyProjectOrigins update failed for ' + p.id + ':', e);
      }
    }
    invalidateProjectCache();
    try {
      if (typeof logAuditEvent === 'function') {
        logAuditEvent('project.origin_bulk_reclassified', 'project', '*', '', { count: changed });
      }
    } catch (e) {}
    return { success: true, changed: changed };
  } catch (error) {
    console.error('bulkReclassifyProjectOrigins failed:', error);
    return { success: false, error: error.message };
  }
}

var COMPLIANCE_EXEMPTABLE_KEYS = ['githubLinks', 'relatedDataAssets', 'userGuides'];

function _hasValue_(v) {
  if (v === null || v === undefined) return false;
  if (typeof v === 'boolean') return v === true;
  if (Array.isArray(v)) return v.length > 0;
  var s = String(v).trim();
  if (!s || s === '[]' || s === '{}' || s === 'null' || s === 'undefined') return false;
  if (s === CONFIG.TBD_DATE_SENTINEL) return false;
  return true;
}

function _hasMultiValue_(v) {
  if (!v) return false;
  if (Array.isArray(v)) return v.filter(function(x) { return String(x || '').trim(); }).length > 0;
  if (typeof v === 'string') {
    try { var parsed = JSON.parse(v); if (Array.isArray(parsed)) return parsed.filter(function(x) { return String(x || '').trim(); }).length > 0; } catch (e) {}
    return v.split(/[,;]/).map(function(s) { return s.trim(); }).filter(Boolean).length > 0;
  }
  return false;
}

function _boolDefined_(v) {
  return v === true || v === false || v === 'true' || v === 'false' || v === 'TRUE' || v === 'FALSE';
}

function _boolTrue_(v) {
  return v === true || v === 'true' || v === 'TRUE';
}

function buildComplianceRequirements(project, settings, relatedAssetCount) {
  var type = String(settings.projectType || project.projectType || '').trim();
  var isDeploymentType = type === 'Web Application' || type === 'Data Visualization';
  var deployedPublicly = _boolTrue_(settings.deployedPublicly);
  var activeGas = _boolTrue_(settings.activeGas);
  var activeDmp = _boolTrue_(settings.activeDataMgmt);
  var locationKeys = ['standaloneAppLink', 'd2dLink', 'aasIntranetLink', 'gsaLink', 'tableauLink'];
  var hasAnyLocation = locationKeys.some(function(k) { return _hasValue_(settings[k]); });

  var reqs = [
    { key: 'googleDriveFolder', label: 'Google Drive Folder', filled: _hasValue_(settings.googleDriveFolder) },
    { key: 'githubLinks', label: 'Git Repository', filled: _hasValue_(settings.githubLinks), exemptable: true },
    { key: 'ownerId', label: 'Primary Tech Owner', filled: _hasValue_(project.ownerId) },
    { key: 'secondaryUsers', label: 'Backup Owner(s)', filled: _hasMultiValue_(settings.secondaryUsers) },
    { key: 'projectCode', label: 'Project ID', filled: _hasValue_(settings.projectCode) },
    { key: 'startDate', label: 'Project Start Date', filled: _hasValue_(project.startDate) },
    { key: 'contractCurrent', label: 'Current Contract Ownership', filled: _hasValue_(settings.contractCurrent) },
    { key: 'description', label: 'Project Description', filled: _hasValue_(project.description) },
    { key: 'devDocs', label: 'Development Documentation', filled: _hasValue_(settings.devDocs) },
    { key: 'techStack', label: 'Tech Stack', filled: _hasMultiValue_(settings.techStack) },
    { key: 'activeGas', label: 'Active GAS Project', filled: _boolDefined_(settings.activeGas) },
    { key: 'relatedDataAssets', label: 'Related Data Assets', filled: (relatedAssetCount || 0) > 0, exemptable: true },
    { key: 'activeDataMgmt', label: 'Active Data Management Process', filled: _boolDefined_(settings.activeDataMgmt) },
    { key: 'dataDictionaries', label: 'Data Dictionary', filled: _hasValue_(settings.dataDictionaries) },
    { key: 'userGuides', label: 'User Guide', filled: _hasValue_(settings.userGuides), exemptable: true },
    { key: 'supportRunbook', label: 'Support Runbook', filled: _hasValue_(settings.supportRunbook) }
  ];

  if (isDeploymentType) {
    reqs.splice(2, 0, { key: 'deployedPublicly', label: 'Deployed Publicly', filled: _boolDefined_(settings.deployedPublicly) });
    if (deployedPublicly) {
      reqs.splice(3, 0,
        { key: 'deploymentLocation', label: 'Deployment Location', filled: _hasMultiValue_(settings.deploymentLocation) },
        { key: 'deploymentLink', label: 'Deployment Link (any of Standalone/D2D/AAS/GSA/Tableau)', filled: hasAnyLocation }
      );
    }
  }

  if (activeGas) {
    var gasIdx = reqs.findIndex(function(r) { return r.key === 'activeGas'; }) + 1;
    reqs.splice(gasIdx, 0,
      { key: 'gasProjectUid', label: 'GAS Project Link', filled: _hasValue_(settings.gasProjectUid) },
      { key: 'gasOwner', label: 'GAS Project Owner', filled: _hasValue_(settings.gasOwner) }
    );
  }

  if (activeDmp) {
    var dmpIdx = reqs.findIndex(function(r) { return r.key === 'activeDataMgmt'; }) + 1;
    reqs.splice(dmpIdx, 0,
      { key: 'dataSourceFiles', label: 'Data Source File(s)', filled: _hasValue_(settings.dataSourceFiles) },
      { key: 'dataCadence', label: 'Data Refresh Cadence', filled: _hasValue_(settings.dataCadence) },
      { key: 'dataSourceExplain', label: 'How is Data Sourced', filled: _hasValue_(settings.dataSourceExplain) }
    );
  }

  return reqs;
}

function computeProjectCompliance(project, relatedAssetCount, preparsedSettings) {
  if (!project) return { rating: 'Non-Compliant', score: 0, total: 0, filled: 0, missing: [], exempted: [], requirements: [] };
  var settings = preparsedSettings || _parseProjectSettings_(project);
  var exemptions = settings.docExemptions || {};
  var reqs = buildComplianceRequirements(project, settings, relatedAssetCount || 0);
  var total = 0, filled = 0;
  var missing = [], exempted = [];
  var enriched = reqs.map(function(r) {
    var exemption = r.exemptable && exemptions[r.key] ? exemptions[r.key] : null;
    if (exemption) {
      exempted.push({ key: r.key, label: r.label, reason: exemption.reason || '' });
      return Object.assign({}, r, { status: 'exempted', exemption: exemption });
    }
    total++;
    if (r.filled) { filled++; return Object.assign({}, r, { status: 'filled' }); }
    missing.push({ key: r.key, label: r.label, exemptable: !!r.exemptable });
    return Object.assign({}, r, { status: 'missing' });
  });
  var score = total === 0 ? 100 : Math.round((filled / total) * 100);
  var rating = score >= 100 ? 'Compliant' : 'Non-Compliant';
  return {
    rating: rating,
    score: score,
    total: total,
    filled: filled,
    missing: missing,
    exempted: exempted,
    requirements: enriched
  };
}

function _countRelatedDataAssets_(projectId) {
  try {
    if (typeof getAllDataAssetsOptimized !== 'function') return 0;
    var assets = getAllDataAssetsOptimized();
    var count = 0;
    for (var i = 0; i < assets.length; i++) {
      var ids = String(assets[i].relatedProjects || '').split(/[,;]/).map(function(s) { return s.trim(); }).filter(Boolean);
      if (ids.indexOf(projectId) !== -1) count++;
    }
    return count;
  } catch (e) { return 0; }
}

function getProjectCompliance(projectId) {
  try {
    var project = getProjectById(projectId);
    if (!project) return { success: false, error: 'Project not found' };
    var count = _countRelatedDataAssets_(projectId);
    var result = computeProjectCompliance(project, count);
    return { success: true, compliance: result };
  } catch (error) {
    console.error('getProjectCompliance failed:', error);
    return { success: false, error: error.message };
  }
}

function setProjectDocExemption(projectId, fieldKey, reason) {
  try {
    PermissionGuard.requirePermission('project:update', { projectId: projectId });
    if (COMPLIANCE_EXEMPTABLE_KEYS.indexOf(fieldKey) === -1) {
      return { success: false, error: 'Field is not exemptable: ' + fieldKey };
    }
    var trimmedReason = String(reason || '').trim();
    if (!trimmedReason) return { success: false, error: 'Exemption reason is required' };
    var project = getProjectById(projectId);
    if (!project) return { success: false, error: 'Project not found' };
    var settings = _parseProjectSettings_(project);
    if (!settings.docExemptions) settings.docExemptions = {};
    settings.docExemptions[fieldKey] = {
      reason: trimmedReason,
      exemptedAt: new Date().toISOString(),
      exemptedBy: getCurrentUserEmail()
    };
    updateProject(projectId, { settings: JSON.stringify(settings) });
    invalidateProjectCache(projectId);
    return { success: true, exemption: settings.docExemptions[fieldKey] };
  } catch (error) {
    console.error('setProjectDocExemption failed:', error);
    return { success: false, error: error.message };
  }
}

function clearProjectDocExemption(projectId, fieldKey) {
  try {
    PermissionGuard.requirePermission('project:update', { projectId: projectId });
    var project = getProjectById(projectId);
    if (!project) return { success: false, error: 'Project not found' };
    var settings = _parseProjectSettings_(project);
    if (settings.docExemptions && settings.docExemptions[fieldKey]) {
      delete settings.docExemptions[fieldKey];
      updateProject(projectId, { settings: JSON.stringify(settings) });
      invalidateProjectCache(projectId);
    }
    return { success: true };
  } catch (error) {
    console.error('clearProjectDocExemption failed:', error);
    return { success: false, error: error.message };
  }
}

var DOC_TEMPLATE_KEYS = ['designDoc', 'devDoc', 'dataDictionary', 'userGuide', 'supportRunbook'];

function getDocTemplates() {
  try {
    var props = PropertiesService.getScriptProperties();
    var raw = props.getProperty('DOC_TEMPLATES_JSON');
    var stored = {};
    if (raw) { try { stored = JSON.parse(raw); } catch (e) { stored = {}; } }
    var result = {};
    DOC_TEMPLATE_KEYS.forEach(function(k) {
      result[k] = stored[k] && stored[k].url ? stored[k] : null;
    });
    return { success: true, templates: result };
  } catch (error) {
    console.error('getDocTemplates failed:', error);
    return { success: false, error: error.message, templates: {} };
  }
}

function setDocTemplate(key, url, name) {
  try {
    PermissionGuard.requirePermission('admin:settings');
    if (DOC_TEMPLATE_KEYS.indexOf(key) === -1) return { success: false, error: 'Invalid template key: ' + key };
    var props = PropertiesService.getScriptProperties();
    var raw = props.getProperty('DOC_TEMPLATES_JSON');
    var stored = {};
    if (raw) { try { stored = JSON.parse(raw); } catch (e) { stored = {}; } }
    if (!url || !String(url).trim()) {
      delete stored[key];
    } else {
      stored[key] = {
        url: String(url).trim(),
        name: String(name || '').trim() || key,
        updatedAt: new Date().toISOString(),
        updatedBy: getCurrentUserEmail()
      };
    }
    props.setProperty('DOC_TEMPLATES_JSON', JSON.stringify(stored));
    return { success: true, templates: stored };
  } catch (error) {
    console.error('setDocTemplate failed:', error);
    return { success: false, error: error.message };
  }
}

function getProjectSubProjects(projectId) {
  try {
    var project = getProjectById(projectId);
    if (!project) return { success: false, error: 'Project not found', subProjects: [] };
    var settings = _parseProjectSettings_(project);
    var subs = Array.isArray(settings.subProjects) ? settings.subProjects : [];
    return { success: true, isModular: settings.isModular === true, subProjects: subs };
  } catch (error) {
    console.error('getProjectSubProjects failed:', error);
    return { success: false, error: error.message, subProjects: [] };
  }
}

function setProjectModular(projectId, isModular) {
  try {
    PermissionGuard.requirePermission('project:update', { projectId: projectId });
    var project = getProjectById(projectId);
    if (!project) return { success: false, error: 'Project not found' };
    var settings = _parseProjectSettings_(project);
    settings.isModular = isModular === true;
    if (!Array.isArray(settings.subProjects)) settings.subProjects = [];
    updateProject(projectId, { settings: JSON.stringify(settings) });
    invalidateProjectCache(projectId);
    return { success: true, isModular: settings.isModular, subProjects: settings.subProjects };
  } catch (error) {
    console.error('setProjectModular failed:', error);
    return { success: false, error: error.message };
  }
}

function addSubProject(projectId, subProjectData) {
  try {
    PermissionGuard.requirePermission('project:update', { projectId: projectId });
    var project = getProjectById(projectId);
    if (!project) return { success: false, error: 'Project not found' };
    var settings = _parseProjectSettings_(project);
    if (!settings.isModular) settings.isModular = true;
    if (!Array.isArray(settings.subProjects)) settings.subProjects = [];

    var name = (subProjectData && subProjectData.name ? String(subProjectData.name) : '').trim();
    if (!name) return { success: false, error: 'Sub-project name is required' };

    var parentCode = settings.projectCode || '';
    if (!parentCode) {
      return { success: false, error: 'Generate a Project ID for the parent before adding sub-projects' };
    }

    var maxSuffix = 0;
    settings.subProjects.forEach(function(sub) {
      if (!sub || !sub.id) return;
      var idx = sub.id.lastIndexOf('-');
      if (idx === -1) return;
      var tail = sub.id.substring(idx + 1);
      var n = parseInt(tail, 10);
      if (!isNaN(n) && n > maxSuffix) maxSuffix = n;
    });
    var nextSuffix = maxSuffix + 1;
    var newId = parentCode + '-' + nextSuffix;

    var sub = {
      id: newId,
      name: name,
      description: subProjectData.description ? String(subProjectData.description).trim() : '',
      createdAt: now()
    };
    settings.subProjects.push(sub);
    updateProject(projectId, { settings: JSON.stringify(settings) });
    invalidateProjectCache(projectId);
    return { success: true, subProject: sub, subProjects: settings.subProjects };
  } catch (error) {
    console.error('addSubProject failed:', error);
    return { success: false, error: error.message };
  }
}

function updateSubProject(projectId, subProjectId, updates) {
  try {
    PermissionGuard.requirePermission('project:update', { projectId: projectId });
    var project = getProjectById(projectId);
    if (!project) return { success: false, error: 'Project not found' };
    var settings = _parseProjectSettings_(project);
    var subs = Array.isArray(settings.subProjects) ? settings.subProjects : [];
    var idx = subs.findIndex(function(s) { return s && s.id === subProjectId; });
    if (idx === -1) return { success: false, error: 'Sub-project not found' };
    if (updates && typeof updates.name === 'string') {
      var newName = updates.name.trim();
      if (!newName) return { success: false, error: 'Sub-project name is required' };
      subs[idx].name = newName;
    }
    if (updates && typeof updates.description === 'string') {
      subs[idx].description = updates.description.trim();
    }
    subs[idx].updatedAt = now();
    settings.subProjects = subs;
    updateProject(projectId, { settings: JSON.stringify(settings) });
    invalidateProjectCache(projectId);
    return { success: true, subProject: subs[idx] };
  } catch (error) {
    console.error('updateSubProject failed:', error);
    return { success: false, error: error.message };
  }
}

function removeSubProject(projectId, subProjectId) {
  try {
    PermissionGuard.requirePermission('project:update', { projectId: projectId });
    var project = getProjectById(projectId);
    if (!project) return { success: false, error: 'Project not found' };
    var settings = _parseProjectSettings_(project);
    var subs = Array.isArray(settings.subProjects) ? settings.subProjects : [];
    var tasksUsingSub = 0;
    try {
      var tasks = getAllTasksOptimized();
      tasksUsingSub = tasks.filter(function(t) { return t.projectId === projectId && t.subProjectId === subProjectId; }).length;
    } catch (e) {}
    if (tasksUsingSub > 0) {
      return { success: false, error: 'Cannot remove: ' + tasksUsingSub + ' task(s) reference this sub-project' };
    }
    settings.subProjects = subs.filter(function(s) { return s && s.id !== subProjectId; });
    updateProject(projectId, { settings: JSON.stringify(settings) });
    invalidateProjectCache(projectId);
    return { success: true, subProjects: settings.subProjects };
  } catch (error) {
    console.error('removeSubProject failed:', error);
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

