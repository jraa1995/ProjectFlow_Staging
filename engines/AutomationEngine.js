const AutomationEngine = {
  TRIGGERS: {
    ON_CREATE: 'on_create',
    ON_STATUS_CHANGE: 'on_status_change',
    ON_ASSIGNMENT_CHANGE: 'on_assignment_change',
    ON_DUE_APPROACHING: 'on_due_approaching',
    ON_DEPENDENCY_COMPLETE: 'on_dependency_complete',
    ON_PRIORITY_CHANGE: 'on_priority_change',
    ON_COMMENT_ADDED: 'on_comment_added',
    SCHEDULED: 'scheduled'
  },

  ACTIONS: {
    SET_STATUS: 'set_status',
    SET_PRIORITY: 'set_priority',
    ASSIGN_TASK: 'assign_task',
    ADD_LABEL: 'add_label',
    REMOVE_LABEL: 'remove_label',
    SEND_NOTIFICATION: 'send_notification',
    CREATE_SUBTASK: 'create_subtask',
    SET_DUE_DATE: 'set_due_date',
    ADD_COMMENT: 'add_comment'
  },

  getRules(projectId) {
    const sheet = getAutomationRulesSheet();
    const data = sheet.getDataRange().getValues();
    const columns = CONFIG.AUTOMATION_RULE_COLUMNS;
    const rules = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      const rule = rowToObject(row, columns);
      if (typeof rule.conditions === 'string' && rule.conditions) {
        try {
          rule.conditions = JSON.parse(rule.conditions);
        } catch (e) {
          rule.conditions = [];
        }
      } else {
        rule.conditions = [];
      }
      if (typeof rule.actions === 'string' && rule.actions) {
        try {
          rule.actions = JSON.parse(rule.actions);
        } catch (e) {
          rule.actions = [];
        }
      } else {
        rule.actions = [];
      }
      rule.enabled = rule.enabled === true || rule.enabled === 'true';
      if (projectId && rule.projectId && rule.projectId !== projectId) {
        continue;
      }
      rules.push(rule);
    }
    return rules;
  },

  createRule(ruleData) {
    const sheet = getAutomationRulesSheet();
    const columns = CONFIG.AUTOMATION_RULE_COLUMNS;
    const currentUser = getCurrentUserEmail();
    const rule = {
      id: generateId('auto'),
      name: ruleData.name || 'New Rule',
      description: ruleData.description || '',
      trigger: ruleData.trigger,
      conditions: JSON.stringify(ruleData.conditions || []),
      actions: JSON.stringify(ruleData.actions || []),
      enabled: ruleData.enabled !== false,
      projectId: ruleData.projectId || '',
      createdBy: currentUser,
      createdAt: now(),
      lastTriggeredAt: ''
    };
    sheet.appendRow(objectToRow(rule, columns));
    return rule;
  },

  updateRule(ruleId, updates) {
    const sheet = getAutomationRulesSheet();
    const data = sheet.getDataRange().getValues();
    const columns = CONFIG.AUTOMATION_RULE_COLUMNS;
    const idIndex = 0;
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIndex] === ruleId) {
        const rule = rowToObject(data[i], columns);
        if (updates.conditions && typeof updates.conditions !== 'string') {
          updates.conditions = JSON.stringify(updates.conditions);
        }
        if (updates.actions && typeof updates.actions !== 'string') {
          updates.actions = JSON.stringify(updates.actions);
        }
        Object.assign(rule, updates);
        const newRow = objectToRow(rule, columns);
        sheet.getRange(i + 1, 1, 1, columns.length).setValues([newRow]);
        return rule;
      }
    }
    throw new Error('Rule not found: ' + ruleId);
  },

  deleteRule(ruleId) {
    const sheet = getAutomationRulesSheet();
    const data = sheet.getDataRange().getValues();
    const idIndex = 0;
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIndex] === ruleId) {
        sheet.deleteRow(i + 1);
        return true;
      }
    }
    return false;
  },

  evaluateTrigger(trigger, context) {
    const results = [];
    try {
      const rules = this.getRules(context.task?.projectId)
        .filter(rule => rule.enabled && rule.trigger === trigger);
      for (const rule of rules) {
        try {
          if (this.evaluateConditions(rule.conditions, context)) {
            const actionResults = this.executeActions(rule.actions, context);
            results.push({
              ruleId: rule.id,
              ruleName: rule.name,
              success: true,
              actions: actionResults
            });
            this.updateRule(rule.id, { lastTriggeredAt: now() });
          }
        } catch (error) {
          results.push({
            ruleId: rule.id,
            ruleName: rule.name,
            success: false,
            error: error.message
          });
        }
      }
    } catch (error) {
    }
    return results;
  },

  evaluateConditions(conditions, context) {
    if (!conditions || conditions.length === 0) return true;
    return conditions.every(condition => {
      const value = this.getContextValue(condition.field, context);
      return this.evaluateCondition(value, condition.operator, condition.value);
    });
  },

  getContextValue(field, context) {
    const parts = field.split('.');
    let value = context;
    for (const part of parts) {
      if (value === undefined || value === null) return undefined;
      value = value[part];
    }
    return value;
  },

  evaluateCondition(value, operator, expected) {
    switch (operator) {
      case 'equals':
        return value === expected;
      case 'not_equals':
        return value !== expected;
      case 'contains':
        return String(value).toLowerCase().includes(String(expected).toLowerCase());
      case 'not_contains':
        return !String(value).toLowerCase().includes(String(expected).toLowerCase());
      case 'starts_with':
        return String(value).toLowerCase().startsWith(String(expected).toLowerCase());
      case 'ends_with':
        return String(value).toLowerCase().endsWith(String(expected).toLowerCase());
      case 'is_empty':
        return !value || value === '' || (Array.isArray(value) && value.length === 0);
      case 'is_not_empty':
        return value && value !== '' && !(Array.isArray(value) && value.length === 0);
      case 'greater_than':
        return Number(value) > Number(expected);
      case 'less_than':
        return Number(value) < Number(expected);
      case 'in_list':
        const list = Array.isArray(expected) ? expected : String(expected).split(',').map(s => s.trim());
        return list.includes(String(value));
      case 'not_in_list':
        const notList = Array.isArray(expected) ? expected : String(expected).split(',').map(s => s.trim());
        return !notList.includes(String(value));
      default:
        return false;
    }
  },

  executeActions(actions, context) {
    const results = [];
    for (const action of actions) {
      try {
        const result = this.executeAction(action, context);
        results.push({
          action: action.type,
          success: true,
          result: result
        });
      } catch (error) {
        results.push({
          action: action.type,
          success: false,
          error: error.message
        });
      }
    }
    return results;
  },

  executeAction(action, context) {
    const task = context.task;
    switch (action.type) {
      case this.ACTIONS.SET_STATUS:
        return updateTask(task.id, { status: action.value });
      case this.ACTIONS.SET_PRIORITY:
        return updateTask(task.id, { priority: action.value });
      case this.ACTIONS.ASSIGN_TASK:
        const assignee = this.resolveAssignee(action.value, context);
        return updateTask(task.id, { assignee: assignee });
      case this.ACTIONS.ADD_LABEL:
        const currentLabels = task.labels || [];
        if (!currentLabels.includes(action.value)) {
          return updateTask(task.id, { labels: [...currentLabels, action.value] });
        }
        return task;
      case this.ACTIONS.REMOVE_LABEL:
        const labels = (task.labels || []).filter(l => l !== action.value);
        return updateTask(task.id, { labels: labels });
      case this.ACTIONS.SEND_NOTIFICATION:
        const recipient = action.recipient || task.assignee || task.reporter;
        if (recipient) {
          NotificationEngine.createNotification({
            userId: recipient,
            type: action.notificationType || 'task_updated',
            title: action.title || 'Automation Notification',
            message: this.interpolateMessage(action.message || '', context),
            entityType: 'task',
            entityId: task.id,
            channels: ['in_app', 'email']
          });
        }
        return { notified: recipient };
      case this.ACTIONS.CREATE_SUBTASK:
        const subtask = createTask({
          title: this.interpolateMessage(action.title || 'Subtask', context),
          description: action.description || '',
          projectId: task.projectId,
          parentId: task.id,
          assignee: action.assignee || task.assignee,
          priority: action.priority || task.priority,
          status: 'To Do'
        });
        return { subtaskId: subtask.id };
      case this.ACTIONS.SET_DUE_DATE:
        const dueDate = this.calculateDueDate(action.value, context);
        return updateTask(task.id, { dueDate: dueDate });
      case this.ACTIONS.ADD_COMMENT:
        return addComment(task.id, this.interpolateMessage(action.message || '', context));
      default:
        throw new Error('Unknown action type: ' + action.type);
    }
  },

  resolveAssignee(value, context) {
    if (value === 'reporter') {
      return context.task?.reporter;
    }
    if (value === 'round_robin') {
      return this.getNextRoundRobinAssignee(context.task?.projectId);
    }
    if (value === 'least_loaded') {
      return this.getLeastLoadedUser(context.task?.projectId);
    }
    return value;
  },

  getNextRoundRobinAssignee(projectId) {
    const users = getActiveUsers();
    if (users.length === 0) return null;
    const cache = CacheService.getScriptCache();
    const key = 'ROUND_ROBIN_' + (projectId || 'global');
    let index = parseInt(cache.get(key)) || 0;
    const user = users[index % users.length];
    cache.put(key, String((index + 1) % users.length), 86400);
    return user.email;
  },

  getLeastLoadedUser(projectId) {
    const users = getActiveUsers();
    const tasks = getAllTasks({ projectId: projectId });
    const taskCounts = {};
    users.forEach(u => taskCounts[u.email] = 0);
    tasks.forEach(t => {
      if (t.assignee && t.status !== 'Done') {
        taskCounts[t.assignee] = (taskCounts[t.assignee] || 0) + 1;
      }
    });
    let minUser = users[0];
    let minCount = Infinity;
    users.forEach(user => {
      const count = taskCounts[user.email] || 0;
      if (count < minCount) {
        minCount = count;
        minUser = user;
      }
    });
    return minUser?.email;
  },

  calculateDueDate(value, context) {
    if (value.startsWith('+')) {
      const amount = parseInt(value.substring(1));
      const unit = value.slice(-1);
      const date = new Date();
      switch (unit) {
        case 'd':
          date.setDate(date.getDate() + amount);
          break;
        case 'w':
          date.setDate(date.getDate() + (amount * 7));
          break;
        case 'm':
          date.setMonth(date.getMonth() + amount);
          break;
        default:
          date.setDate(date.getDate() + amount);
      }
      return date.toISOString().split('T')[0];
    }
    return value;
  },

  interpolateMessage(message, context) {
    return message.replace(/\{(\w+(?:\.\w+)*)\}/g, (match, path) => {
      const value = this.getContextValue(path, context);
      return value !== undefined ? String(value) : match;
    });
  },

  onTaskCreated(task) {
    this.evaluateTrigger(this.TRIGGERS.ON_CREATE, { task: task });
  },

  onTaskUpdated(task, oldValues, newValues) {
    const context = { task, oldValues, newValues };
    if (oldValues.status !== newValues.status) {
      this.evaluateTrigger(this.TRIGGERS.ON_STATUS_CHANGE, context);
    }
    if (oldValues.assignee !== newValues.assignee) {
      this.evaluateTrigger(this.TRIGGERS.ON_ASSIGNMENT_CHANGE, context);
    }
    if (oldValues.priority !== newValues.priority) {
      this.evaluateTrigger(this.TRIGGERS.ON_PRIORITY_CHANGE, context);
    }
  },

  onDependencyComplete(task, completedDependency) {
    this.evaluateTrigger(this.TRIGGERS.ON_DEPENDENCY_COMPLETE, {
      task: task,
      completedDependency: completedDependency
    });
  },

  processScheduledRules() {
    const rules = this.getRules().filter(r => r.enabled && r.trigger === this.TRIGGERS.SCHEDULED);
    for (const rule of rules) {
      try {
        const tasks = getAllTasks();
        const matchingTasks = tasks.filter(task =>
          this.evaluateConditions(rule.conditions, { task })
        );
        for (const task of matchingTasks) {
          this.executeActions(rule.actions, { task });
        }
        this.updateRule(rule.id, { lastTriggeredAt: now() });
      } catch (error) {
      }
    }
  },

  checkDueApproaching() {
    const tasks = getAllTasks();
    const now = new Date();
    const dayFromNow = new Date(now.getTime() + 24 * 60 * 60 * 1000);
    tasks.forEach(task => {
      if (!task.dueDate || task.status === 'Done') return;
      const dueDate = new Date(task.dueDate);
      if (dueDate > now && dueDate <= dayFromNow) {
        this.evaluateTrigger(this.TRIGGERS.ON_DUE_APPROACHING, {
          task: task,
          dueIn: 'day'
        });
      }
    });
  }
};
