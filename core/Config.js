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
    ACCESS_REQUESTS: 'Access_Requests'
  },

  JSON_CACHE_COLUMNS: ['key', 'data', 'updatedAt', 'version'],

  STATUSES: ['Backlog', 'To Do', 'In Progress', 'Review', 'Testing', 'Done'],

  PRIORITIES: ['Lowest', 'Low', 'Medium', 'High', 'Highest', 'Critical'],

  TYPES: ['Task', 'Bug', 'Feature', 'Story', 'Epic', 'Spike'],

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
    'jsonData'
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
    'jsonData'
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
    'comment_added'
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

  FUNNEL_STATUSES: ['pending', 'reviewed', 'imported', 'rejected'],

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

  TAXONOMY_FIELDS: ['workstream', 'projectCategory', 'projectType', 'developmentPriority', 'developmentPhase', 'techStack'],

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
  const key = 'TASK_SEQ_' + prefix;
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const props = PropertiesService.getScriptProperties();
    const next = parseInt(props.getProperty(key) || '0') + 1;
    props.setProperty(key, String(next));
    return prefix + '-' + next;
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
  var next = parseInt(props.getProperty(key) || '0') + 1;
  props.setProperty(key, String(next));
  return prefix + '-' + next;
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
