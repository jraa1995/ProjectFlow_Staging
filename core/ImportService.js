function getProjectsWorkbookId() {
  return PropertiesService.getScriptProperties().getProperty('PROJECTS_WORKBOOK_ID') || '';
}

function setProjectsWorkbookId(workbookId) {
  PropertiesService.getScriptProperties().setProperty('PROJECTS_WORKBOOK_ID', workbookId);
  return true;
}

function getDataAssetsWorkbookId() {
  return PropertiesService.getScriptProperties().getProperty('DATA_ASSETS_WORKBOOK_ID') || '';
}

function setDataAssetsWorkbookId(workbookId) {
  PropertiesService.getScriptProperties().setProperty('DATA_ASSETS_WORKBOOK_ID', workbookId);
  return true;
}

var CPMD_DEFAULT_WORKBOOK_ID = '1hOnry0rHLp3-fQq2gMqaVbPzEJmmCjQdkrjfufCIBRI';

function getCpmdWorkbookId() {
  return PropertiesService.getScriptProperties().getProperty('CPMD_WORKBOOK_ID') || CPMD_DEFAULT_WORKBOOK_ID;
}

function setCpmdWorkbookId(workbookId) {
  var prev = PropertiesService.getScriptProperties().getProperty('CPMD_WORKBOOK_ID');
  PropertiesService.getScriptProperties().setProperty('CPMD_WORKBOOK_ID', workbookId);
  try {
    var cache = CacheService.getScriptCache();
    if (prev) cache.remove('CPMD_PERSONNEL_CACHE_' + prev);
    if (workbookId) cache.remove('CPMD_PERSONNEL_CACHE_' + workbookId);
  } catch (e) {}
  return true;
}

var FED_STAFFING_ROLES = {
  personnel:    { label: 'General Personnel Roster',     defaultSheet: 'Personnel'   },
  govOwners:    { label: 'Government Project Owners',    defaultSheet: 'GovOwners'   },
  stakeholders: { label: 'Primary Gov Stakeholders',     defaultSheet: 'Stakeholders' }
};

function _fedStaffingAssertRole_(role) {
  if (!role || !FED_STAFFING_ROLES.hasOwnProperty(role)) {
    throw new Error('Unknown Fed Staffing role: ' + role);
  }
  return role;
}

function getFedStaffingWorkbookId() {
  return PropertiesService.getScriptProperties().getProperty('FED_STAFFING_WORKBOOK_ID') || '';
}

function _fedStaffingClearAllCaches_() {
  try {
    var wbId = PropertiesService.getScriptProperties().getProperty('FED_STAFFING_WORKBOOK_ID');
    if (!wbId) return;
    var cache = CacheService.getScriptCache();
    var keys = Object.keys(FED_STAFFING_ROLES).map(function(r) { return 'FED_STAFFING_CACHE_' + wbId + '_' + r; });
    cache.removeAll(keys);
  } catch (e) {}
}

function setFedStaffingWorkbookId(workbookId) {
  var prev = PropertiesService.getScriptProperties().getProperty('FED_STAFFING_WORKBOOK_ID');
  if (prev) {
    try {
      var cache = CacheService.getScriptCache();
      var prevKeys = Object.keys(FED_STAFFING_ROLES).map(function(r) { return 'FED_STAFFING_CACHE_' + prev + '_' + r; });
      cache.removeAll(prevKeys);
    } catch (e) {}
  }
  if (workbookId) {
    PropertiesService.getScriptProperties().setProperty('FED_STAFFING_WORKBOOK_ID', workbookId);
  } else {
    PropertiesService.getScriptProperties().deleteProperty('FED_STAFFING_WORKBOOK_ID');
  }
  _fedStaffingClearAllCaches_();
  return true;
}

function getFedStaffingSheetMap() {
  var raw = PropertiesService.getScriptProperties().getProperty('FED_STAFFING_SHEETS');
  var map = {};
  if (raw) {
    try { map = JSON.parse(raw) || {}; } catch (e) { map = {}; }
  }
  Object.keys(FED_STAFFING_ROLES).forEach(function(role) {
    if (!map[role]) map[role] = FED_STAFFING_ROLES[role].defaultSheet;
  });
  return map;
}

function getFedStaffingSheetName(role) {
  _fedStaffingAssertRole_(role || 'personnel');
  return getFedStaffingSheetMap()[role || 'personnel'];
}

function setFedStaffingSheet(role, sheetName) {
  _fedStaffingAssertRole_(role);
  var raw = PropertiesService.getScriptProperties().getProperty('FED_STAFFING_SHEETS');
  var map = {};
  if (raw) {
    try { map = JSON.parse(raw) || {}; } catch (e) { map = {}; }
  }
  if (sheetName) map[role] = String(sheetName);
  else delete map[role];
  PropertiesService.getScriptProperties().setProperty('FED_STAFFING_SHEETS', JSON.stringify(map));
  try {
    var wbId = getFedStaffingWorkbookId();
    if (wbId) CacheService.getScriptCache().remove('FED_STAFFING_CACHE_' + wbId + '_' + role);
  } catch (e) {}
  return true;
}

function _toTitleCase_(s) {
  if (!s) return '';
  return String(s).trim().toLowerCase().replace(/\b\w/g, function(m) { return m.toUpperCase(); });
}

function _fedStaffingReadRoster_(wbId, sheetName) {
  var ss = SpreadsheetApp.openById(wbId);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0].map(function(h) { return String(h || '').trim().toLowerCase(); });
  var findCol = function(names) {
    for (var n = 0; n < names.length; n++) {
      var idx = headers.indexOf(names[n]);
      if (idx !== -1) return idx;
    }
    return -1;
  };
  var firstIdx = findCol(['first name', 'firstname', 'first']);
  var lastIdx = findCol(['last name', 'lastname', 'last']);
  var displayIdx = findCol(['display name', 'displayname']);
  var nameIdx = findCol(['full name', 'name', 'stakeholder', 'owner']);
  var emailIdx = findCol(['email', 'email address', 'e-mail']);
  var statusIdx = findCol(['status', 'active', 'employment status']);
  var typeIdx = findCol(['type', 'employment type', 'category']);
  var titleIdx = findCol(['title', 'role', 'position']);
  var orgIdx = findCol(['org', 'organization', 'org code', 'ipt', 'division', 'bu', 'contract task', 'contract']);
  var results = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (statusIdx !== -1) {
      var status = String(row[statusIdx] || '').trim().toLowerCase();
      if (status && status !== 'active' && status !== 'yes' && status !== 'true' && status !== '1') continue;
    }
    if (typeIdx !== -1) {
      var t = String(row[typeIdx] || '').trim().toLowerCase();
      if (t.indexOf('employee') !== 0) continue;
    }
    var name = '';
    if (firstIdx !== -1 || lastIdx !== -1) {
      var first = firstIdx !== -1 ? _toTitleCase_(row[firstIdx]) : '';
      var last = lastIdx !== -1 ? _toTitleCase_(row[lastIdx]) : '';
      name = (first + ' ' + last).trim();
    }
    if (!name && displayIdx !== -1) name = _toTitleCase_(row[displayIdx]);
    if (!name && nameIdx !== -1) name = _toTitleCase_(row[nameIdx]);
    if (!name) continue;
    var email = emailIdx !== -1 ? String(row[emailIdx] || '').trim().toLowerCase() : '';
    var title = titleIdx !== -1 ? String(row[titleIdx] || '').trim() : '';
    var org = orgIdx !== -1 ? String(row[orgIdx] || '').trim() : '';
    var rowType = typeIdx !== -1 ? String(row[typeIdx] || '').trim() : '';
    var entry = { name: name, email: email, title: title, org: org };
    if (rowType) entry.type = rowType;
    entry.label = email ? (name + ' (' + email + ')') : name;
    results.push(entry);
  }
  results.sort(function(a, b) { return a.name.localeCompare(b.name); });
  return results;
}

function getFedStaffingPersonnel(role) {
  role = _fedStaffingAssertRole_(role || 'personnel');
  var wbId = getFedStaffingWorkbookId();
  if (!wbId) return [];
  var sheetName = getFedStaffingSheetMap()[role];
  if (!sheetName) return [];
  var cacheKey = 'FED_STAFFING_CACHE_' + wbId + '_' + role;
  try {
    var cached = CacheService.getScriptCache().get(cacheKey);
    if (cached) return JSON.parse(cached);
  } catch (e) {}
  try {
    var results = _fedStaffingReadRoster_(wbId, sheetName);
    try { CacheService.getScriptCache().put(cacheKey, JSON.stringify(results), 1800); } catch (e) {}
    return results;
  } catch (error) {
    console.error('getFedStaffingPersonnel failed for role ' + role + ':', error);
    return [];
  }
}

function getCpmdPersonnel() {
  var wbId = getCpmdWorkbookId();
  if (!wbId) return [];
  var cacheKey = 'CPMD_PERSONNEL_CACHE_' + wbId;
  try {
    var cached = CacheService.getScriptCache().get(cacheKey);
    if (cached) return JSON.parse(cached);
  } catch (e) {}
  try {
    var ss = SpreadsheetApp.openById(wbId);
    var sheet = ss.getSheetByName('Personnel');
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    var results = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var status = String(row[7] || '').trim();
      if (status.toLowerCase() !== 'active') continue;
      var first = _toTitleCase_(row[2]);
      var last = _toTitleCase_(row[3]);
      var email = String(row[5] || '').trim().toLowerCase();
      if (!first && !last) continue;
      var name = (first + ' ' + last).trim();
      var entry = { name: name, email: email };
      entry.label = email ? (name + ' (' + email + ')') : name;
      results.push(entry);
    }
    results.sort(function(a, b) { return a.name.localeCompare(b.name); });
    try { CacheService.getScriptCache().put(cacheKey, JSON.stringify(results), 1800); } catch (e) {}
    return results;
  } catch (error) {
    console.error('getCpmdPersonnel failed:', error);
    return [];
  }
}

var WORKLOG_STATUS_MAP = {
  '00-Ideation': 'active',
  '01-Requirements Gathering': 'active',
  '02-In Development': 'active',
  '03-Deployed (Actively Maintained)': 'active',
  '04-Deployed (Inactive-Not Maintained)': 'archived',
  '05-Dev Complete/Pending Deployment': 'active',
  '06-Retired/Sunset (Permanent)': 'archived',
  '07-Backlog': 'active',
  '08-Stopped': 'archived'
};

var WORKLOG_PHASE_MAP = {
  '02-': 'Development',
  '03-': 'Deployed',
  '04-': 'Maintenance',
  '05-': 'Pre-Deploy',
  '06-': 'Retired',
  '07-': 'Backlog',
  '08-': 'Stopped'
};

function normalizeImportField(value) {
  if (!value) return '';
  var trimmed = String(value).trim();
  if (trimmed === 'N/A' || trimmed === 'FALSE' || trimmed === 'false') return '';
  return trimmed;
}

var WORKLOG_COL = {
  docCompleteness: 0,
  projectStatus: 1,
  workstream: 2,
  projectCategory: 3,
  projectType: 4,
  developmentPriority: 5,
  developmentPhase: 6,
  govRequestDate: 7,
  startDate: 8,
  firstDeployDate: 9,
  retiredDate: 10,
  externalProjectId: 11,
  linkedProjectId: 12,
  name: 13,
  description: 14,
  isInternalHive: 15,
  inProjectFlow: 16,
  legacyWbs: 17,
  techStack: 18,
  sdmSupported: 19,
  deployedPublicly: 20,
  deploymentLocation: 21,
  tableauLink: 22,
  standaloneAppLink: 23,
  d2dLink: 24,
  aasIntranetLink: 25,
  gsaLink: 26,
  isAiInvolved: 27,
  contractInitial: 28,
  contractCurrent: 29,
  contractTask: 30,
  projectOwner: 31,
  backupOwners: 32,
  projectPoc: 33,
  clientOwner: 34,
  clientStakeholders: 35,
  futureContractHome: 36,
  futureOwner: 37,
  transitionPriority: 38,
  transitionMeetingDate: 39,
  transitionComplete: 40,
  activeDataMgmt: 41,
  dataMgmtOwner: 42,
  dataTaskOwnership: 43,
  dataSourceFiles: 44,
  dataSourceExplain: 45,
  dataCadence: 46,
  activeGas: 47,
  gasOwner: 48,
  gasProjectUid: 49,
  googleDriveFolder: 50,
  githubLinks: 51,
  devDocs: 52,
  sops: 53,
  userGuides: 54,
  dataDictionaries: 55,
  additionalFiles: 56,
  notes: 57
};

function convertWorklogDate(val) {
  if (!val) return '';
  var str = String(val).trim();
  var match = str.match(/^(\d{1,2})\/(\d{4})$/);
  if (match) return match[2] + '-' + match[1].padStart(2, '0') + '-01';
  var match2 = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (match2) return match2[3] + '-' + match2[1].padStart(2, '0') + '-' + match2[2].padStart(2, '0');
  return '';
}

function isDeployedStatus(statusStr) {
  return statusStr && (statusStr.indexOf('03-') === 0 || statusStr.indexOf('04-') === 0 || statusStr.indexOf('05-') === 0 || statusStr.indexOf('06-') === 0);
}

function buildWorklogTags(row) {
  var C = WORKLOG_COL;
  var tags = [];
  if (row[C.workstream]) tags.push(String(row[C.workstream]).trim());
  if (row[C.projectCategory]) tags.push(String(row[C.projectCategory]).trim());
  if (row[C.projectType]) tags.push(String(row[C.projectType]).trim());
  if (row[C.developmentPriority]) tags.push(String(row[C.developmentPriority]).trim());
  if (row[C.contractCurrent]) tags.push(String(row[C.contractCurrent]).trim());
  if (row[C.developmentPhase]) tags.push(String(row[C.developmentPhase]).trim());
  var seen = {};
  return tags.filter(function(t) {
    if (!t || t === 'N/A' || t === 'FALSE') return false;
    var key = t.toLowerCase().trim();
    if (seen[key]) return false;
    seen[key] = true;
    return true;
  });
}

function buildWorklogSettings(row, ownerIndex) {
  var C = WORKLOG_COL;
  var s = {};
  var backups = String(row[C.backupOwners] || '').split(',').map(function(b) { return b.trim(); }).filter(Boolean);
  if (backups.length > 0) {
    var resolved = [];
    var pending = [];
    var seenEmails = {};
    backups.forEach(function(b) {
      if (typeof _resolveOwnerToEmail_ === 'function' && ownerIndex) {
        var r = _resolveOwnerToEmail_(b, ownerIndex);
        if (r.email) {
          var e = r.email.toLowerCase();
          if (!seenEmails[e]) { resolved.push(e); seenEmails[e] = true; }
        } else {
          pending.push(b);
        }
      } else {
        pending.push(b);
      }
    });
    if (resolved.length > 0) s.secondaryUsers = resolved;
    if (pending.length > 0) s.pendingSecondaryUsers = pending;
  }
  if (row[C.workstream]) s.workstream = String(row[C.workstream]).trim();
  if (row[C.projectCategory]) s.projectCategory = String(row[C.projectCategory]).trim();
  if (row[C.projectType]) s.projectType = String(row[C.projectType]).trim();
  if (row[C.developmentPriority]) s.developmentPriority = String(row[C.developmentPriority]).trim();
  if (row[C.developmentPhase]) s.developmentPhase = String(row[C.developmentPhase]).trim();
  if (row[C.techStack]) s.techStack = String(row[C.techStack]).trim();
  if (row[C.contractInitial]) s.contractInitial = String(row[C.contractInitial]).trim();
  if (row[C.contractCurrent]) s.contractCurrent = String(row[C.contractCurrent]).trim();
  if (row[C.contractTask]) s.contractTask = String(row[C.contractTask]).trim();
  if (row[C.deploymentLocation]) s.deploymentLocation = String(row[C.deploymentLocation]).trim();
  if (row[C.standaloneAppLink]) s.standaloneAppLink = String(row[C.standaloneAppLink]).trim();
  if (row[C.d2dLink]) s.d2dLink = String(row[C.d2dLink]).trim();
  if (row[C.aasIntranetLink]) s.aasIntranetLink = String(row[C.aasIntranetLink]).trim();
  if (row[C.gsaLink]) s.gsaLink = String(row[C.gsaLink]).trim();
  if (row[C.tableauLink]) s.tableauLink = String(row[C.tableauLink]).trim();
  if (row[C.gasProjectUid]) s.gasProjectUid = String(row[C.gasProjectUid]).trim();
  if (row[C.googleDriveFolder]) s.googleDriveFolder = String(row[C.googleDriveFolder]).trim();
  if (row[C.externalProjectId]) s.externalProjectId = String(row[C.externalProjectId]).trim();
  if (row[C.linkedProjectId]) s.linkedProjectId = String(row[C.linkedProjectId]).trim();
  if (row[C.sdmSupported]) s.sdmSupported = String(row[C.sdmSupported]).trim();
  if (row[C.projectPoc]) s.projectPoc = String(row[C.projectPoc]).trim();
  if (row[C.clientOwner]) s.clientOwner = String(row[C.clientOwner]).trim();
  if (row[C.clientStakeholders]) s.clientStakeholders = String(row[C.clientStakeholders]).trim();
  if (row[C.transitionPriority]) s.transitionPriority = String(row[C.transitionPriority]).trim();
  if (row[C.futureContractHome]) s.futureContractHome = String(row[C.futureContractHome]).trim();
  if (row[C.futureOwner]) s.futureOwner = String(row[C.futureOwner]).trim();
  if (row[C.dataCadence]) s.dataCadence = String(row[C.dataCadence]).trim();
  if (row[C.dataSourceFiles]) s.dataSourceFiles = String(row[C.dataSourceFiles]).trim();
  if (row[C.dataSourceExplain]) s.dataSourceExplain = String(row[C.dataSourceExplain]).trim();
  if (row[C.devDocs]) s.devDocs = String(row[C.devDocs]).trim();
  if (row[C.sops]) s.sops = String(row[C.sops]).trim();
  if (row[C.userGuides]) s.userGuides = String(row[C.userGuides]).trim();
  if (row[C.dataDictionaries]) s.dataDictionaries = String(row[C.dataDictionaries]).trim();
  if (row[C.notes]) s.notes = String(row[C.notes]).trim();
  var isAi = String(row[C.isAiInvolved] || '').trim();
  if (isAi === 'AI-Assisted' || isAi === 'TRUE' || isAi === 'true') {
    s.isAiInvolved = 'Yes - AI used in Development';
  } else if (isAi && isAi !== 'FALSE' && isAi !== 'false') {
    s.isAiInvolved = isAi;
  } else {
    s.isAiInvolved = 'No - AI Not Involved in Solution';
  }
  var isHive = String(row[C.isInternalHive] || '').trim();
  s.isInternalHive = isHive === 'TRUE' || isHive === 'true';
  var isPublic = String(row[C.deployedPublicly] || '').trim();
  s.deployedPublicly = isPublic === 'TRUE' || isPublic === 'true';
  var hasGas = String(row[C.activeGas] || '').trim();
  s.activeGas = hasGas === 'TRUE' || hasGas === 'true';
  if (row[C.gasOwner]) s.gasOwner = String(row[C.gasOwner]).trim();
  var hasDataMgmt = String(row[C.activeDataMgmt] || '').trim();
  s.activeDataMgmt = hasDataMgmt === 'TRUE' || hasDataMgmt === 'true';
  if (row[C.dataMgmtOwner]) s.dataMgmtOwner = String(row[C.dataMgmtOwner]).trim();
  if (row[C.docCompleteness]) s.docCompleteness = String(row[C.docCompleteness]).trim();
  if (row[C.govRequestDate]) s.govRequestDate = convertWorklogDate(row[C.govRequestDate]);
  if (row[C.firstDeployDate]) s.firstDeployDate = convertWorklogDate(row[C.firstDeployDate]);
  var inPF = String(row[C.inProjectFlow] || '').trim();
  s.inProjectFlow = inPF === 'TRUE' || inPF === 'true' || inPF === 'Public' || inPF === 'Private';
  if (row[C.legacyWbs]) s.legacyWbs = String(row[C.legacyWbs]).trim();
  if (row[C.transitionMeetingDate]) s.transitionMeetingDate = convertWorklogDate(row[C.transitionMeetingDate]);
  var transComplete = String(row[C.transitionComplete] || '').trim();
  s.transitionComplete = transComplete === 'TRUE' || transComplete === 'true';
  if (row[C.dataTaskOwnership]) s.dataTaskOwnership = String(row[C.dataTaskOwnership]).trim();
  if (row[C.additionalFiles]) s.additionalFiles = String(row[C.additionalFiles]).trim();
  if (row[C.githubLinks]) {
    var ghVal = String(row[C.githubLinks]).trim();
    var ghLinks = ghVal.split(/[,;\n]/).map(function(l) { return l.trim(); }).filter(Boolean);
    if (ghLinks.length > 1) s.githubLinks = ghVal;
  }
  return s;
}

function importProjectsFromWorkLog(workbookId) {
  PermissionGuard.requirePermission('project:create');
  var storedId = workbookId || getProjectsWorkbookId();
  if (!storedId) throw new Error('No workbook ID provided');
  var sourceSpreadsheet = SpreadsheetApp.openById(storedId);
  var sourceSheet = sourceSpreadsheet.getSheets()[0];
  var sourceData = sourceSheet.getDataRange().getValues();

  var C = WORKLOG_COL;
  var existingProjects = getAllProjects();
  var existingNames = {};
  existingProjects.forEach(function(p) { existingNames[p.name.toLowerCase().trim()] = true; });
  var generatedIds = existingProjects.slice();

  var projectSheet = getProjectsSheet();
  var sheetHeaders = projectSheet.getRange(1, 1, 1, projectSheet.getLastColumn()).getValues()[0].map(function(h) { return String(h).trim(); });
  var columns = sheetHeaders;
  var currentUser = getCurrentUserEmail();
  var timestamp = now();
  var ownerIndex = {};
  try {
    if (typeof _buildUserOwnerIndex_ === 'function' && typeof getAllUsers === 'function') {
      ownerIndex = _buildUserOwnerIndex_(getAllUsers());
    }
  } catch (ownerLookupErr) {
    console.error('Owner index build failed; falling back to raw values:', ownerLookupErr);
  }

  var rows = [];
  var imported = [];
  var skipped = [];

  for (var i = 1; i < sourceData.length; i++) {
    var row = sourceData[i];
    var name = String(row[C.name] || '').trim();
    if (!name) continue;

    if (existingNames[name.toLowerCase()]) {
      skipped.push(name);
      continue;
    }
    existingNames[name.toLowerCase()] = true;

    var projectId = generateProjectAcronym(name, generatedIds);
    var statusStr = String(row[C.projectStatus] || '').trim();
    var tags = buildWorklogTags(row);
    var settings = buildWorklogSettings(row, ownerIndex);
    if (!settings.developmentPhase && statusStr) {
      var phasePrefix = statusStr.substring(0, 3);
      if (WORKLOG_PHASE_MAP[phasePrefix]) settings.developmentPhase = WORKLOG_PHASE_MAP[phasePrefix];
    }
    if (statusStr) settings.projectStatus = statusStr;

    var rawOwner = String(row[C.projectOwner] || '').trim();
    var resolvedOwner = '';
    if (rawOwner && typeof _resolveOwnerToEmail_ === 'function') {
      var r = _resolveOwnerToEmail_(rawOwner, ownerIndex);
      if (r.email) {
        resolvedOwner = r.email;
      } else {
        settings.pendingOwnerName = rawOwner;
      }
    } else if (rawOwner) {
      settings.pendingOwnerName = rawOwner;
    } else {
      resolvedOwner = currentUser;
    }

    var project = {
      id: projectId,
      name: sanitize(name),
      description: sanitize(String(row[C.description] || '')),
      status: WORKLOG_STATUS_MAP[statusStr] || 'active',
      ownerId: resolvedOwner,
      startDate: convertWorklogDate(row[C.startDate]),
      endDate: convertWorklogDate(row[C.retiredDate]),
      createdAt: timestamp,
      updatedAt: timestamp,
      version: isDeployedStatus(statusStr) ? '1.0.0' : '0.1.0',
      repoUrl: String(row[C.githubLinks] || '').trim(),
      releaseNotes: '[]',
      changelog: '[]',
      tags: JSON.stringify(tags),
      settings: JSON.stringify(settings),
      lastUpdatedBy: currentUser,
      workstream: settings.workstream || '',
      projectCategory: settings.projectCategory || '',
      projectType: settings.projectType || '',
      developmentPriority: settings.developmentPriority || '',
      developmentPhase: settings.developmentPhase || '',
      techStack: settings.techStack || '',
      linkedProjectId: settings.linkedProjectId || ''
    };

    generatedIds.push(project);
    rows.push(objectToRow(project, columns));
    imported.push(name);
  }

  if (rows.length > 0) {
    var lastRow = projectSheet.getLastRow();
    projectSheet.getRange(lastRow + 1, 1, rows.length, columns.length).setValues(rows);
    SpreadsheetApp.flush();
    logActivity(currentUser, 'bulk_imported', 'project', 'batch', { count: rows.length });
  }

  invalidateProjectCache();

  return {
    success: true,
    imported: imported.length,
    skipped: skipped.length,
    importedNames: imported,
    skippedNames: skipped
  };
}

function syncWorklogSettings(workbookId) {
  PermissionGuard.requirePermission('project:update');
  var storedId = workbookId || getProjectsWorkbookId();
  if (!storedId) throw new Error('No workbook ID provided');
  var sourceSpreadsheet = SpreadsheetApp.openById(storedId);
  var sourceSheet = sourceSpreadsheet.getSheets()[0];
  var sourceData = sourceSheet.getDataRange().getValues();

  var C = WORKLOG_COL;
  var existingProjects = getAllProjects();
  var nameMap = {};
  existingProjects.forEach(function(p) { nameMap[p.name.toLowerCase().trim()] = p; });
  var ownerIndex = {};
  try {
    if (typeof _buildUserOwnerIndex_ === 'function' && typeof getAllUsers === 'function') {
      ownerIndex = _buildUserOwnerIndex_(getAllUsers());
    }
  } catch (ownerLookupErr) {
    console.error('Owner index build failed; falling back to raw values:', ownerLookupErr);
  }

  var updated = [];
  var skipped = [];

  for (var i = 1; i < sourceData.length; i++) {
    var row = sourceData[i];
    var name = String(row[C.name] || '').trim();
    if (!name) continue;

    var existing = nameMap[name.toLowerCase()];
    if (!existing) { skipped.push(name); continue; }

    var newSettings = buildWorklogSettings(row, ownerIndex);
    var currentSettings = {};
    try { currentSettings = typeof existing.settings === 'string' ? JSON.parse(existing.settings) : (existing.settings || {}); } catch (e) { currentSettings = {}; }

    Object.keys(newSettings).forEach(function(key) {
      if (newSettings[key] !== undefined && newSettings[key] !== '' && newSettings[key] !== false) {
        currentSettings[key] = newSettings[key];
      }
    });
    if (!newSettings.pendingSecondaryUsers && currentSettings.pendingSecondaryUsers) {
      delete currentSettings.pendingSecondaryUsers;
    }

    var newTags = buildWorklogTags(row);
    var statusStr = String(row[C.projectStatus] || '').trim();
    if (!currentSettings.developmentPhase && statusStr) {
      var phasePrefix = statusStr.substring(0, 3);
      if (WORKLOG_PHASE_MAP[phasePrefix]) currentSettings.developmentPhase = WORKLOG_PHASE_MAP[phasePrefix];
    }
    if (statusStr) currentSettings.projectStatus = statusStr;
    var rawOwnerUpdate = String(row[C.projectOwner] || '').trim();
    var resolvedOwnerUpdate = existing.ownerId || '';
    if (rawOwnerUpdate && typeof _resolveOwnerToEmail_ === 'function') {
      var ru = _resolveOwnerToEmail_(rawOwnerUpdate, ownerIndex);
      if (ru.email) {
        resolvedOwnerUpdate = ru.email;
        delete currentSettings.pendingOwnerName;
      } else {
        if (!_isEmailLike_(existing.ownerId)) resolvedOwnerUpdate = '';
        currentSettings.pendingOwnerName = rawOwnerUpdate;
      }
    } else if (rawOwnerUpdate) {
      if (!_isEmailLike_(existing.ownerId)) resolvedOwnerUpdate = '';
      currentSettings.pendingOwnerName = rawOwnerUpdate;
    }
    var updates = {
      settings: JSON.stringify(currentSettings),
      tags: JSON.stringify(newTags),
      description: sanitize(String(row[C.description] || existing.description || '')),
      ownerId: resolvedOwnerUpdate,
      repoUrl: String(row[C.githubLinks] || existing.repoUrl || '').trim(),
      startDate: convertWorklogDate(row[C.startDate]) || existing.startDate,
      endDate: convertWorklogDate(row[C.retiredDate]) || existing.endDate
    };
    if (statusStr && WORKLOG_STATUS_MAP[statusStr]) updates.status = WORKLOG_STATUS_MAP[statusStr];

    updateProject(existing.id, updates);
    updated.push(name);
  }

  invalidateProjectCache();

  return { success: true, updated: updated.length, skipped: skipped.length, updatedNames: updated, skippedNames: skipped };
}

var DATA_ASSET_COL = {
  status: 0,
  assetOwner: 1,
  backupOwner: 2,
  assetName: 3,
  dataSource: 4,
  targetFiles: 5,
  relatedProjects: 6,
  primaryStakeholder: 7,
  updateSchedule: 8,
  automatedSchedule: 9,
  currentEnvironment: 10,
  githubLink: 11,
  dataSharingDocLink: 12
};

function importDataAssetsFromWorkLog(workbookId) {
  PermissionGuard.requirePermission('dataasset:create');
  var storedId = workbookId || getDataAssetsWorkbookId() || getRequestsWorkbookId();
  if (!storedId) throw new Error('No workbook ID provided');

  var sourceSpreadsheet = SpreadsheetApp.openById(storedId);
  var sourceSheet = sourceSpreadsheet.getSheetByName('Data Assets');
  if (!sourceSheet) throw new Error('Sheet "Data Assets" not found in workbook');

  var sourceData = sourceSheet.getDataRange().getValues();
  if (sourceData.length <= 2) return { success: true, imported: 0, skipped: 0, importedNames: [], skippedNames: [] };

  var C = DATA_ASSET_COL;
  var existingAssets = getAllDataAssets();
  var existingNames = {};
  existingAssets.forEach(function(a) { existingNames[(a.assetName || '').toLowerCase().trim()] = true; });

  var sheet = getDataAssetsSheet();
  var columns = CONFIG.DATA_ASSET_COLUMNS;
  var currentUser = getCurrentUserEmail();
  var timestamp = now();

  var rows = [];
  var imported = [];
  var skipped = [];

  for (var i = 2; i < sourceData.length; i++) {
    var row = sourceData[i];
    var name = String(row[C.assetName] || '').trim();
    if (!name) continue;

    if (existingNames[name.toLowerCase()]) {
      skipped.push(name);
      continue;
    }
    existingNames[name.toLowerCase()] = true;

    var asset = {
      id: generateId('DA'),
      status: sanitize(String(row[C.status] || 'Active')),
      assetOwner: sanitize(String(row[C.assetOwner] || '')),
      backupOwner: sanitize(String(row[C.backupOwner] || '')),
      assetName: sanitize(name),
      dataSource: sanitize(String(row[C.dataSource] || '')),
      targetFiles: sanitize(String(row[C.targetFiles] || '')),
      relatedProjects: sanitize(String(row[C.relatedProjects] || '')),
      primaryStakeholder: sanitize(String(row[C.primaryStakeholder] || '')),
      updateFrequency: '',
      updateSchedule: sanitize(String(row[C.updateSchedule] || '')),
      automatedSchedule: sanitize(String(row[C.automatedSchedule] || '')),
      currentEnvironment: sanitize(String(row[C.currentEnvironment] || '')),
      githubLink: String(row[C.githubLink] || '').trim(),
      dataSharingDocLink: String(row[C.dataSharingDocLink] || '').trim(),
      createdAt: timestamp,
      updatedAt: timestamp,
      lastUpdatedBy: currentUser,
      assetType: '',
      bucketId: ''
    };

    rows.push(objectToRow(asset, columns));
    imported.push(name);
  }

  if (rows.length > 0) {
    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, rows.length, columns.length).setValues(rows);
    SpreadsheetApp.flush();
    logActivity(currentUser, 'bulk_imported', 'dataasset', 'batch', { count: rows.length });
  }

  invalidateDataAssetCache();

  return {
    success: true,
    imported: imported.length,
    skipped: skipped.length,
    importedNames: imported,
    skippedNames: skipped
  };
}

function previewImportFromWorkLog(workbookId) {
  var storedId = workbookId || getProjectsWorkbookId();
  if (!storedId) return { success: false, error: 'No workbook ID provided' };
  try {
    var sourceSpreadsheet = SpreadsheetApp.openById(storedId);
    var sourceSheet = sourceSpreadsheet.getSheets()[0];
    var sourceData = sourceSheet.getDataRange().getValues();
    var C = WORKLOG_COL;
    var existingProjects = getAllProjects();
    var existingNames = {};
    existingProjects.forEach(function(p) { existingNames[p.name.toLowerCase().trim()] = true; });
    var projects = [];
    var duplicates = [];
    var warnings = [];
    for (var i = 1; i < sourceData.length; i++) {
      var row = sourceData[i];
      var name = String(row[C.name] || '').trim();
      if (!name) continue;
      if (existingNames[name.toLowerCase()]) {
        duplicates.push(name);
        continue;
      }
      var statusStr = String(row[C.projectStatus] || '').trim();
      var settings = buildWorklogSettings(row);
      if (!settings.developmentPhase && statusStr) {
        var phasePrefix = statusStr.substring(0, 3);
        if (WORKLOG_PHASE_MAP[phasePrefix]) settings.developmentPhase = WORKLOG_PHASE_MAP[phasePrefix];
      }
      var ghVal = String(row[C.githubLinks] || '').trim();
      var repoUrl = ghVal;
      if (ghVal && ghVal.split(/[,;\n]/).filter(Boolean).length > 1) {
        repoUrl = ghVal.split(/[,;\n]/)[0].trim();
      }
      if (!statusStr) warnings.push(name + ': missing status');
      if (!row[C.projectOwner]) warnings.push(name + ': missing owner');
      projects.push({
        name: name,
        status: WORKLOG_STATUS_MAP[statusStr] || 'active',
        phase: settings.developmentPhase || '',
        owner: String(row[C.projectOwner] || '').trim(),
        repoUrl: repoUrl,
        settingsKeys: Object.keys(settings).length
      });
    }
    return { success: true, projects: projects, duplicates: duplicates, warnings: warnings };
  } catch (error) {
    console.error('previewImportFromWorkLog failed:', error);
    return { success: false, error: error.message };
  }
}
