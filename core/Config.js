const CONFIG = {
  SHEETS: {
    TASKS: 'Tasks',
    USERS: 'Users',
    PROJECTS: 'Projects',
    COMMENTS: 'Comments',
    ACTIVITY: 'Activity',
    MENTIONS: 'Mentions',
    NOTIFICATIONS: 'Notifications',
    ANALYTICS_CACHE: 'Analytics_Cache',
    TASK_DEPENDENCIES: 'Task_Dependencies',
    FUNNEL_STAGING: 'Funnel_Staging',
    ROLES: 'Roles',
    PROJECT_MEMBERS: 'Project_Members',
    ORGANIZATIONS: 'Organizations',
    ORGANIZATION_MEMBERS: 'Organization_Members',
    TEAMS: 'Teams',
    TEAM_MEMBERS: 'Team_Members',
    USER_PREFERENCES: 'User_Preferences',
    AUTOMATION_RULES: 'Automation_Rules',
    TEAM_CAPACITY: 'Team_Capacity',
    WEBHOOK_SUBSCRIPTIONS: 'Webhook_Subscriptions',
    SLA_CONFIG: 'SLA_Config',
    TRIAGE_QUEUE: 'Triage_Queue',
    JSON_CACHE: 'JSON_Cache',
    DATA_ASSETS: 'Data_Assets',
    ACCESS_REQUESTS: 'Access_Requests',
    USER_BADGES: 'User_Badges',
    USER_METRICS: 'User_Metrics'
  },

  JSON_CACHE_COLUMNS: ['key', 'data', 'updatedAt', 'version'],

  STATUSES: ['Backlog', 'To Do', 'In Progress', 'Review', 'Testing', 'Done'],

  PRIORITIES: ['Lowest', 'Low', 'Medium', 'High', 'Highest', 'Critical'],

  TYPES: ['Task', 'Bug', 'Feature', 'Enhancement', 'Story', 'Epic', 'Spike'],

  COLORS: {
    'Backlog': '#6B7280',
    'To Do': '#3B82F6',
    'In Progress': '#525252',
    'Review': '#8B5CF6',
    'Testing': '#06B6D4',
    'Done': '#10B981'
  },

  TASK_COLUMNS: [
    'id',
    'projectId',
    'title',
    'description',
    'status',
    'priority',
    'type',
    'assignee',
    'watchers',
    'reporter',
    'dueDate',
    'startDate',
    'sprint',
    'storyPoints',
    'estimatedHrs',
    'actualHrs',
    'labels',
    'parentId',
    'position',
    'createdAt',
    'updatedAt',
    'completedAt',
    'dependencies',
    'timeEntries',
    'customFields',
    'templateId',
    'recurringConfig',
    'isMilestone',
    'milestoneDate',
    'milestoneType',
    'version',
    'lastModifiedBy',
    'isDeleted',
    'deletedAt',
    'taskUid',
    'jsonData',
    'requestedBy'
  ],

  USER_COLUMNS: [
    'email',
    'name',
    'role',
    'active',
    'workbookId',
    'createdAt',
    'passwordHash',
    'passwordSalt',
    'lastPasswordChange',
    'failedLoginAttempts',
    'lockoutUntil',
    'mfaEnabled',
    'mfaSecret',
    'lastLogin',
    'organizationId',
    'teamId',
    'jsonData',
    'avatar',
    'title',
    'department',
    'updatedAt'
  ],

  PROJECT_COLUMNS: [
    'id',
    'name',
    'description',
    'status',
    'ownerId',
    'startDate',
    'endDate',
    'createdAt',
    'updatedAt',
    'version',
    'repoUrl',
    'releaseNotes',
    'changelog',
    'tags',
    'settings',
    'lastUpdatedBy',
    'workstream',
    'projectCategory',
    'projectType',
    'developmentPriority',
    'developmentPhase',
    'techStack',
    'sdmSupported',
    'deploymentLocation',
    'linkedProjectId',
    'jsonData'
  ],

  COMMENT_COLUMNS: [
    'id',
    'taskId',
    'userId',
    'content',
    'createdAt',
    'updatedAt',
    'mentionedUsers',
    'isEdited',
    'editHistory'
  ],

  ACTIVITY_COLUMNS: [
    'id',
    'userId',
    'action',
    'entityType',
    'entityId',
    'details',
    'timestamp'
  ],

  MENTION_COLUMNS: [
    'id',
    'commentId',
    'mentionedUserId',
    'mentionedByUserId',
    'taskId',
    'createdAt',
    'notificationSent',
    'acknowledged'
  ],

  NOTIFICATION_COLUMNS: [
    'id',
    'userId',
    'type',
    'title',
    'message',
    'entityType',
    'entityId',
    'read',
    'createdAt',
    'scheduledFor',
    'channels'
  ],

  ANALYTICS_CACHE_COLUMNS: [
    'id',
    'cacheKey',
    'data',
    'expiresAt',
    'createdAt'
  ],

  TASK_DEPENDENCY_COLUMNS: [
    'id',
    'predecessorId',
    'successorId',
    'dependencyType',
    'lag',
    'createdAt'
  ],

  FUNNEL_STAGING_COLUMNS: [
    'id',
    'sourceWorkbookId',
    'sourceSheetName',
    'rawData',
    'mappedData',
    'importStatus',
    'assignedTo',
    'notes',
    'importedTaskId',
    'createdAt',
    'importedAt'
  ],


  ROLE_COLUMNS: [
    'id',
    'name',
    'description',
    'permissions',
    'isSystemRole',
    'createdAt'
  ],

  PROJECT_MEMBER_COLUMNS: [
    'id',
    'projectId',
    'userId',
    'projectRole',
    'permissions',
    'addedAt',
    'addedBy'
  ],

  ORGANIZATION_COLUMNS: [
    'id',
    'name',
    'slug',
    'description',
    'settings',
    'plan',
    'createdAt',
    'ownerId'
  ],

  ORGANIZATION_MEMBER_COLUMNS: [
    'id',
    'organizationId',
    'userId',
    'orgRole',
    'joinedAt',
    'invitedBy'
  ],

  TEAM_COLUMNS: [
    'id',
    'organizationId',
    'name',
    'description',
    'leaderId',
    'parentTeamId',
    'createdAt'
  ],

  TEAM_MEMBER_COLUMNS: [
    'id',
    'teamId',
    'userId',
    'teamRole',
    'joinedAt'
  ],

  USER_PREFERENCE_COLUMNS: [
    'id',
    'userId',
    'category',
    'settings',
    'updatedAt'
  ],

  AUTOMATION_RULE_COLUMNS: [
    'id',
    'name',
    'description',
    'trigger',
    'conditions',
    'actions',
    'enabled',
    'projectId',
    'createdBy',
    'createdAt',
    'lastTriggeredAt'
  ],

  TEAM_CAPACITY_COLUMNS: [
    'id',
    'userId',
    'hoursPerDay',
    'workDays',
    'vacationDays',
    'effectiveFrom',
    'effectiveTo'
  ],

  WEBHOOK_SUBSCRIPTION_COLUMNS: [
    'id',
    'url',
    'events',
    'secret',
    'enabled',
    'projectId',
    'createdBy',
    'createdAt',
    'lastDeliveredAt',
    'failureCount'
  ],

  SLA_CONFIG_COLUMNS: [
    'id',
    'name',
    'priority',
    'responseTime',
    'resolutionTime',
    'escalationRules',
    'enabled',
    'createdAt'
  ],

  TRIAGE_QUEUE_COLUMNS: [
    'id',
    'sourceType',
    'sourceId',
    'rawContent',
    'suggestedAssignee',
    'suggestedPriority',
    'suggestedProject',
    'status',
    'assignedBy',
    'assignedAt',
    'taskId',
    'createdAt'
  ],

  NOTIFICATION_TYPES: [
    'mention',
    'task_assigned',
    'task_updated',
    'deadline_approaching',
    'project_update',
    'project_assigned',
    'project_ping',
    'comment_added',
    'badge_earned'
  ],

  NOTIFICATION_CHANNELS: [
    'email',
    'in_app',
    'push'
  ],

  DEPENDENCY_TYPES: [
    'finish_to_start',
    'start_to_start',
    'finish_to_finish',
    'start_to_finish'
  ],

  FUNNEL_STATUSES: ['pending', 'reviewed', 'assigned', 'imported', 'rejected'],

  REQUEST_TYPE_ROUTING: {
    'Ad Hoc Request':                   { kind: 'task',    projectId: 'ADHOC' },
    'New Project':                      { kind: 'project' },
    'Bug fix for existing project':     { kind: 'task',    taskType: 'Bug' },
    'Enhancement to existing project':  { kind: 'task',    taskType: 'Enhancement' }
  },

  TEMP_STAKEHOLDERS: ['Stakeholder 1', 'Stakeholder 2', 'Stakeholder 3'],

  REQ_GATHERING_TEMPLATES: [
    {
      id: 'business-req-v1', type: 'business', name: 'Business Requirements',
      description: 'Capture business objectives, success criteria, and scope.',
      fields: [
        { key: 'businessObjectives', label: 'Business Objectives', type: 'textarea', required: true },
        { key: 'successCriteria', label: 'Success Criteria', type: 'textarea', required: true },
        { key: 'stakeholderNeeds', label: 'Key Stakeholder Needs', type: 'textarea' },
        { key: 'scopeConstraints', label: 'Scope & Constraints', type: 'textarea' },
        { key: 'intendedUsers', label: 'Intended Users', type: 'static_select', optionsKey: 'intendedUsers', target: 'settings.intendedUsers' },
        { key: 'workstream', label: 'Workstream', type: 'taxonomy_select', target: 'project.workstream' }
      ]
    },
    {
      id: 'stakeholder-req-v1', type: 'stakeholder', name: 'Stakeholder Requirements',
      description: 'Identify stakeholders, expectations, and decision-makers.',
      fields: [
        { key: 'primaryStakeholder', label: 'Primary Stakeholder', type: 'text', target: 'settings.clientOwner' },
        { key: 'secondaryStakeholders', label: 'Other Gov Stakeholders', type: 'textarea', target: 'settings.clientStakeholders' },
        { key: 'stakeholderExpectations', label: 'Stakeholder Expectations', type: 'textarea' },
        { key: 'communicationCadence', label: 'Communication Cadence', type: 'text' },
        { key: 'decisionMakers', label: 'Decision Makers', type: 'textarea' }
      ]
    },
    {
      id: 'technical-req-v1', type: 'technical', name: 'Technical Requirements',
      description: 'Document tech stack, deployment, integrations, and non-functional needs.',
      fields: [
        { key: 'techStack', label: 'Tech Stack', type: 'taxonomy_multiselect', target: 'project.techStack' },
        { key: 'deploymentLocation', label: 'Deployment Location', type: 'taxonomy_select', target: 'project.deploymentLocation' },
        { key: 'dataCadence', label: 'Data Refresh Cadence', type: 'static_select', optionsKey: 'dataCadence', target: 'settings.dataCadence' },
        { key: 'dataSourceExplain', label: 'Data Sourcing', type: 'textarea', target: 'settings.dataSourceExplain' },
        { key: 'integrations', label: 'Integrations Required', type: 'textarea' },
        { key: 'performanceRequirements', label: 'Performance Requirements', type: 'textarea' },
        { key: 'securityRequirements', label: 'Security & Access Requirements', type: 'textarea' }
      ]
    },
    {
      id: 'data-req-v1', type: 'data', name: 'Data Requirements',
      description: 'Define data sources, quality expectations, and update cadence.',
      fields: [
        { key: 'dataSourceFiles', label: 'Data Source Files', type: 'drive_picker_multi', target: 'settings.dataSourceFiles' },
        { key: 'dataDictionaries', label: 'Data Dictionaries', type: 'drive_picker', target: 'settings.dataDictionaries' },
        { key: 'updateFrequency', label: 'Data Update Frequency', type: 'text' },
        { key: 'dataQualityExpectations', label: 'Data Quality Expectations', type: 'textarea' }
      ]
    }
  ],

  DATA_ASSET_COLUMNS: [
    'id', 'status', 'assetOwner', 'backupOwner', 'assetName',
    'dataSource', 'targetFiles', 'relatedProjects', 'primaryStakeholder',
    'updateSchedule', 'automatedSchedule', 'currentEnvironment',
    'githubLink', 'dataSharingDocLink',
    'createdAt', 'updatedAt', 'lastUpdatedBy', 'jsonData', 'updateFrequency'
  ],

  DATA_ASSET_STATUSES: ['Active', 'Inactive', 'Deprecated', 'In Development'],

  ACCESS_REQUEST_COLUMNS: [
    'id', 'email', 'name', 'requestedAt', 'status', 'reviewedBy', 'reviewedAt', 'reason'
  ],

  USER_BADGE_COLUMNS: [
    'id',
    'userEmail',
    'badgeId',
    'badgeName',
    'category',
    'rarity',
    'earnedAt',
    'awardedBy',
    'triggerMetric',
    'triggerValue',
    'notes'
  ],

  USER_METRIC_COLUMNS: [
    'userEmail',
    'tasksCreated',
    'tasksCompleted',
    'tasksCompletedWithinHour',
    'tasksAssignedToOthers',
    'uniqueAssigneesList',
    'uniqueAssigneesCount',
    'unblockerCount',
    'inProgressPeak',
    'commentsPosted',
    'projectsCreated',
    'dataAssetsCreated',
    'orgsOwnedCount',
    'teamsJoinedCount',
    'completionDays',
    'currentCompletionStreak',
    'longestCompletionStreak',
    'lastCompletedAt',
    'loginDays',
    'currentLoginStreak',
    'longestLoginStreak',
    'lastLoginAt',
    'totalBadges',
    'lastEvaluatedAt',
    'updatedAt'
  ],

  TAXONOMY_FIELDS: ['workstream', 'projectCategory', 'projectType', 'developmentPriority', 'developmentPhase', 'techStack', 'sdmSupported', 'deploymentLocation'],

  TAXONOMY_OPTIONS: {
    workstream: ['Business Intelligence', 'Business Development', 'Financial Management', 'Workforce Management', 'Customer Experience', 'Quality Management', 'Contractor Use'],
    sdmSupported: ['FEDSIM', 'FLEX', 'INNOVATE', 'ALL', 'N/A'],
    deploymentLocation: ['D2D', 'AAS Intranet (Google Sites)', 'External GSA Site', 'Standalone Web-App', 'Tableau Server', 'AWS', 'Google App Script Library'],
    developmentPriority: ['High (P1)', 'Medium (P2)', 'Low (P3)', 'Backlog', 'N/A'],
    developmentPhase: ['Coding', 'Complete', 'Deployment', 'Design', 'Enhancements', 'Maintenance', 'Not Started', 'Planning', 'Testing', 'Stopped'],
    aiInvolved: ['Yes - AI used in Development', 'Yes - AI incorporated in solution', 'No - AI Not Involved in Solution'],
    intendedUsers: ['All-AAS', 'OSO', 'BU Leadership', 'BU1', 'BU2', 'BU3', 'BU5', 'Multiple IPTs', 'Single IPT', 'All-Contractors', 'Contractor Leadership', 'BI Team'],
    contractOptions: ['AMPS', 'Forward', 'SQuAT'],
    dataCadence: ['Daily', 'Weekly (Mon-Fri)', 'Biweekly', 'Monthly', 'Quarterly', 'Annually', 'Ad Hoc', 'Real-Time', 'N/A']
  },

  BI_SMES: ['Andra Velea', 'Dan Kain', 'Dan Russell', 'Joe Boozer', 'Justin Aguila', 'Michael Gallahan', 'Michael Thoennes', 'Michelle McAllister'],

  PROJECT_ID_WORKSTREAM_CODES: {
    'Business Intelligence': 'BI',
    'Business Development': 'BD',
    'Financial Management': 'FM',
    'Workforce Management': 'WM',
    'Customer Experience': 'CE',
    'Quality Management': 'QM',
    'Contractor Use': 'CU'
  },

  PROJECT_ID_TYPE_CODES: {
    'Web Application': '01',
    'Data Pipeline': '02',
    'Database': '03',
    'Data Visualization': '04'
  },

  PROJECT_ID_CONTRACT_CODES: {
    'SQuAT': 'S',
    'Forward': 'F',
    'AMPS': 'A'
  },

  PROJECT_TYPE_OPTIONS: ['Web Application', 'Data Pipeline', 'Database', 'Data Visualization'],

  DEPRECATED_SETTINGS_KEYS: ['futureOwner', 'futureContractHome', 'transitionPriority', 'transitionMeetingDate', 'transitionComplete'],

  TBD_DATE_SENTINEL: 'TBD'

};

function getColumnIndex(sheetType, columnName) {
  const columns = CONFIG[sheetType + '_COLUMNS'];
  if (!columns) return -1;
  return columns.indexOf(columnName);
}

function generateId(prefix) {
  const timestamp = Date.now().toString(36);
  const random = Math.random().toString(36).substring(2, 6);
  return `${prefix}_${timestamp}${random}`;
}

function generateTaskId(projectId) {
  const prefix = projectId || 'TASK';
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    return generateTaskIdUnderLock_(prefix);
  } finally {
    lock.releaseLock();
  }
}

function generateTaskUid() {
  return Utilities.getUuid();
}

function isValidTaskUid(uid) {
  return typeof uid === 'string' && /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(uid);
}

function isValidTaskId(taskId) {
  return typeof taskId === 'string' && /^[A-Z0-9]{1,10}-\d+$/.test(taskId);
}

function generateTaskIdUnderLock_(projectId) {
  var prefix = projectId || 'TASK';
  var key = 'TASK_SEQ_' + prefix;
  var props = PropertiesService.getScriptProperties();
  var storedSeq = parseInt(props.getProperty(key) || '0');
  var sheetMax = getMaxTaskSeqFromSheet_(prefix);
  var base = Math.max(storedSeq, sheetMax);
  var next = base + 1;
  props.setProperty(key, String(next));
  return prefix + '-' + next;
}

function getMaxTaskSeqFromSheet_(prefix) {
  try {
    var sheet = getTasksSheet();
    var idCol = CONFIG.TASK_COLUMNS.indexOf('id');
    if (idCol === -1) return 0;
    var data = sheet.getDataRange().getValues();
    var max = 0;
    var pattern = prefix + '-';
    for (var i = 1; i < data.length; i++) {
      var id = String(data[i][idCol] || '');
      if (id.indexOf(pattern) === 0) {
        var seq = parseInt(id.substring(pattern.length));
        if (seq > max) max = seq;
      }
    }
    return max;
  } catch (e) {
    console.error('getMaxTaskSeqFromSheet_ failed:', e);
    return 0;
  }
}

function generateProjectAcronym(name, existingProjects, parentAcronym) {
  if (!name) return 'PROJ';

  const words = name.replace(/[^a-zA-Z0-9\s]/g, '').trim().split(/\s+/);
  let acronym;

  if (words.length >= 2) {
    acronym = words.slice(0, 4).map(w => w[0]).join('').toUpperCase();
  } else {
    acronym = words[0].substring(0, 4).toUpperCase();
  }

  if (parentAcronym) {
    acronym = parentAcronym + '-' + acronym;
  }

  let finalAcronym = acronym;
  let counter = 1;
  const existingIds = new Set(existingProjects.map(p => p.id));

  while (existingIds.has(finalAcronym)) {
    finalAcronym = acronym + counter;
    counter++;
  }

  return finalAcronym;
}

function now() {
  return new Date().toISOString();
}

function sanitize(input) {
  if (typeof input !== 'string') return input;
  if (/^[=+\-@]/.test(input)) {
    input = "'" + input;
  }
  return input
    .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '')
    .replace(/javascript\s*:/gi, '')
    .replace(/data\s*:[^,]*,/gi, '')
    .replace(/on\w+\s*=/gi, '')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}
