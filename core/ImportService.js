var WORKLOG_STATUS_MAP = {
  '02-In Development': 'active',
  '03-Deployed (Actively Maintained)': 'active',
  '04-Deployed (Inactive-Not Maintained)': 'archived',
  '05-Dev Complete/Pending Deployment': 'active',
  '06-Retired/Sunset (Permanent)': 'archived',
  '07-Backlog': 'active',
  '08-Stopped': 'archived'
};

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
  return tags.filter(function(t) { return t && t !== 'N/A' && t !== 'FALSE'; });
}

function buildWorklogSettings(row) {
  var C = WORKLOG_COL;
  var s = {};
  var backups = String(row[C.backupOwners] || '').split(',').map(function(b) { return b.trim(); }).filter(Boolean);
  if (backups.length > 0) s.secondaryUsers = backups;
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
  s.isAiInvolved = isAi === 'AI-Assisted' || isAi === 'TRUE' || isAi === 'true';
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
  if (row[C.githubLinks]) s.githubLinks = String(row[C.githubLinks]).trim();
  return s;
}

function importProjectsFromWorkLog(workbookId) {
  PermissionGuard.requirePermission('project:create');
  var sourceSpreadsheet = SpreadsheetApp.openById(workbookId);
  var sourceSheet = sourceSpreadsheet.getSheets()[0];
  var sourceData = sourceSheet.getDataRange().getValues();

  var C = WORKLOG_COL;
  var existingProjects = getAllProjects();
  var existingNames = {};
  existingProjects.forEach(function(p) { existingNames[p.name.toLowerCase().trim()] = true; });
  var generatedIds = existingProjects.slice();

  var projectSheet = getProjectsSheet();
  var columns = CONFIG.PROJECT_COLUMNS;
  var currentUser = getCurrentUserEmail();
  var timestamp = now();

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
    var settings = buildWorklogSettings(row);

    var project = {
      id: projectId,
      name: sanitize(name),
      description: sanitize(String(row[C.description] || '')),
      status: WORKLOG_STATUS_MAP[statusStr] || 'active',
      ownerId: String(row[C.projectOwner] || currentUser).trim(),
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
      lastUpdatedBy: currentUser
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
  try { CacheService.getScriptCache().remove('BATCH_DATA_CACHE'); } catch (e) {}

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
  var sourceSpreadsheet = SpreadsheetApp.openById(workbookId);
  var sourceSheet = sourceSpreadsheet.getSheets()[0];
  var sourceData = sourceSheet.getDataRange().getValues();

  var C = WORKLOG_COL;
  var existingProjects = getAllProjects();
  var nameMap = {};
  existingProjects.forEach(function(p) { nameMap[p.name.toLowerCase().trim()] = p; });

  var updated = [];
  var skipped = [];

  for (var i = 1; i < sourceData.length; i++) {
    var row = sourceData[i];
    var name = String(row[C.name] || '').trim();
    if (!name) continue;

    var existing = nameMap[name.toLowerCase()];
    if (!existing) { skipped.push(name); continue; }

    var newSettings = buildWorklogSettings(row);
    var currentSettings = {};
    try { currentSettings = typeof existing.settings === 'string' ? JSON.parse(existing.settings) : (existing.settings || {}); } catch (e) { currentSettings = {}; }

    Object.keys(newSettings).forEach(function(key) {
      if (newSettings[key] !== undefined && newSettings[key] !== '' && newSettings[key] !== false) {
        currentSettings[key] = newSettings[key];
      }
    });

    var newTags = buildWorklogTags(row);
    var statusStr = String(row[C.projectStatus] || '').trim();
    var updates = {
      settings: JSON.stringify(currentSettings),
      tags: JSON.stringify(newTags),
      description: sanitize(String(row[C.description] || existing.description || '')),
      ownerId: String(row[C.projectOwner] || existing.ownerId || '').trim(),
      repoUrl: String(row[C.githubLinks] || existing.repoUrl || '').trim(),
      startDate: convertWorklogDate(row[C.startDate]) || existing.startDate,
      endDate: convertWorklogDate(row[C.retiredDate]) || existing.endDate
    };
    if (statusStr && WORKLOG_STATUS_MAP[statusStr]) updates.status = WORKLOG_STATUS_MAP[statusStr];

    updateProject(existing.id, updates);
    updated.push(name);
  }

  invalidateProjectCache();
  try { CacheService.getScriptCache().remove('BATCH_DATA_CACHE'); } catch (e) {}

  return { success: true, updated: updated.length, skipped: skipped.length, updatedNames: updated, skippedNames: skipped };
}
