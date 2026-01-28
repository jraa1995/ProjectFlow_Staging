const TriageEngine = {
  ASSIGNMENT_STRATEGIES: {
    KEYWORD: 'keyword',
    ROUND_ROBIN: 'round_robin',
    WORKLOAD: 'workload',
    PROJECT_DEFAULT: 'project_default'
  },

  submitTicket(ticketData) {
    const sheet = getTriageQueueSheet();
    const columns = CONFIG.TRIAGE_QUEUE_COLUMNS;
    const suggestions = this.autoSuggest(ticketData);
    const queueItem = {
      id: generateId('triage'),
      sourceType: ticketData.sourceType || 'manual',
      sourceId: ticketData.sourceId || '',
      rawContent: JSON.stringify({
        title: ticketData.title,
        description: ticketData.description,
        email: ticketData.submitterEmail,
        metadata: ticketData.metadata
      }),
      suggestedAssignee: suggestions.assignee || '',
      suggestedPriority: suggestions.priority || 'Medium',
      suggestedProject: suggestions.project || '',
      status: 'pending',
      assignedBy: '',
      assignedAt: '',
      taskId: '',
      createdAt: now()
    };
    sheet.appendRow(objectToRow(queueItem, columns));
    this.notifyTriagers(queueItem);
    return queueItem;
  },

  submitFromEmail(emailData) {
    return this.submitTicket({
      sourceType: 'email',
      sourceId: emailData.messageId || emailData.from,
      title: emailData.subject,
      description: emailData.body,
      submitterEmail: emailData.from,
      metadata: {
        receivedAt: emailData.receivedAt || now(),
        hasAttachments: (emailData.attachments || []).length > 0
      }
    });
  },

  submitFromForm(formData) {
    return this.submitTicket({
      sourceType: 'form',
      sourceId: formData.formId || 'external_form',
      title: formData.title || formData.subject,
      description: formData.description || formData.message,
      submitterEmail: formData.email,
      metadata: formData.metadata || {}
    });
  },

  autoSuggest(ticketData) {
    const suggestions = {
      assignee: null,
      priority: 'Medium',
      project: null,
      confidence: {
        assignee: 0,
        priority: 0,
        project: 0
      }
    };
    const content = (ticketData.title + ' ' + ticketData.description).toLowerCase();
    suggestions.priority = this.suggestPriority(content);
    suggestions.confidence.priority = 70;
    const projectSuggestion = this.suggestProject(content);
    if (projectSuggestion) {
      suggestions.project = projectSuggestion.id;
      suggestions.confidence.project = projectSuggestion.confidence;
    }
    const assigneeSuggestion = this.suggestAssignee(content, suggestions.project);
    if (assigneeSuggestion) {
      suggestions.assignee = assigneeSuggestion.email;
      suggestions.confidence.assignee = assigneeSuggestion.confidence;
    }
    return suggestions;
  },

  suggestPriority(content) {
    const priorityKeywords = {
      'Critical': ['critical', 'emergency', 'urgent', 'production down', 'outage', 'security breach'],
      'Highest': ['asap', 'immediately', 'blocking', 'severe', 'major issue'],
      'High': ['important', 'high priority', 'deadline', 'customer impact'],
      'Low': ['when possible', 'nice to have', 'minor', 'cosmetic', 'low priority'],
      'Lowest': ['someday', 'backlog', 'future', 'maybe']
    };
    for (const [priority, keywords] of Object.entries(priorityKeywords)) {
      if (keywords.some(kw => content.includes(kw))) {
        return priority;
      }
    }
    return 'Medium';
  },

  suggestProject(content) {
    const projects = getAllProjects();
    for (const project of projects) {
      const projectName = project.name.toLowerCase();
      const projectId = project.id.toLowerCase();
      if (content.includes(projectName) || content.includes(projectId)) {
        return { id: project.id, confidence: 90 };
      }
      const words = projectName.split(' ').filter(w => w.length > 3);
      const matchCount = words.filter(w => content.includes(w)).length;
      if (matchCount > 0 && matchCount >= words.length / 2) {
        return { id: project.id, confidence: 60 };
      }
    }
    return null;
  },

  suggestAssignee(content, projectId) {
    if (projectId) {
      const project = getProjectById(projectId);
      if (project?.ownerId) {
        return { email: project.ownerId, confidence: 60 };
      }
    }
    if (typeof CapacityEngine !== 'undefined') {
      const available = CapacityEngine.findAvailableUsers({});
      if (available.length > 0 && available[0].canFit) {
        return { email: available[0].userId, confidence: 40 };
      }
    }
    return null;
  },

  getTriageQueue(filters = {}) {
    const sheet = getTriageQueueSheet();
    const data = sheet.getDataRange().getValues();
    const columns = CONFIG.TRIAGE_QUEUE_COLUMNS;
    const items = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      const item = rowToObject(row, columns);
      if (typeof item.rawContent === 'string' && item.rawContent) {
        try {
          item.parsedContent = JSON.parse(item.rawContent);
        } catch (e) {
          item.parsedContent = { raw: item.rawContent };
        }
      }
      if (filters.status && item.status !== filters.status) continue;
      if (filters.sourceType && item.sourceType !== filters.sourceType) continue;
      items.push(item);
    }
    items.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
    return items;
  },

  getTriageItem(itemId) {
    const items = this.getTriageQueue();
    return items.find(i => i.id === itemId) || null;
  },

  assignTriageItem(itemId, taskData) {
    const item = this.getTriageItem(itemId);
    if (!item) {
      throw new Error('Triage item not found: ' + itemId);
    }
    if (item.status !== 'pending') {
      throw new Error('Triage item already processed');
    }
    let content = {};
    try {
      content = JSON.parse(item.rawContent);
    } catch (e) {
      content = { raw: item.rawContent };
    }
    const task = createTask({
      title: taskData.title || content.title || 'Imported Ticket',
      description: taskData.description || content.description || '',
      projectId: taskData.projectId || item.suggestedProject || '',
      assignee: taskData.assignee || item.suggestedAssignee || '',
      priority: taskData.priority || item.suggestedPriority || 'Medium',
      status: taskData.status || 'To Do',
      type: taskData.type || 'Task',
      labels: taskData.labels || ['imported', item.sourceType]
    });
    this.updateTriageItem(itemId, {
      status: 'assigned',
      taskId: task.id,
      assignedBy: getCurrentUserEmail(),
      assignedAt: now()
    });
    if (content.email) {
      this.notifySubmitter(content.email, task, 'assigned');
    }
    return task;
  },

  rejectTriageItem(itemId, reason) {
    const item = this.getTriageItem(itemId);
    if (!item) {
      throw new Error('Triage item not found: ' + itemId);
    }
    const updated = this.updateTriageItem(itemId, {
      status: 'rejected',
      notes: reason,
      assignedBy: getCurrentUserEmail(),
      assignedAt: now()
    });
    let content = {};
    try {
      content = JSON.parse(item.rawContent);
    } catch (e) {}
    if (content.email) {
      this.notifySubmitter(content.email, null, 'rejected', reason);
    }
    return updated;
  },

  updateTriageItem(itemId, updates) {
    const sheet = getTriageQueueSheet();
    const data = sheet.getDataRange().getValues();
    const columns = CONFIG.TRIAGE_QUEUE_COLUMNS;
    const idIndex = 0;
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIndex] === itemId) {
        const item = rowToObject(data[i], columns);
        Object.assign(item, updates);
        const newRow = objectToRow(item, columns);
        sheet.getRange(i + 1, 1, 1, columns.length).setValues([newRow]);
        return item;
      }
    }
    throw new Error('Triage item not found: ' + itemId);
  },

  notifyTriagers(queueItem) {
    const users = getActiveUsers().filter(u => ['admin', 'manager'].includes(u.role));
    for (const user of users) {
      try {
        NotificationEngine.createNotification({
          userId: user.email,
          type: 'project_update',
          title: 'New Ticket in Triage Queue',
          message: `A new ${queueItem.sourceType} ticket has been submitted and needs review`,
          entityType: 'triage',
          entityId: queueItem.id,
          channels: ['in_app']
        });
      } catch (error) {
      }
    }
  },

  notifySubmitter(email, task, status, reason) {
    try {
      let subject, body;
      if (status === 'assigned' && task) {
        subject = `Your request has been received: ${task.title}`;
        body = `
        <p>Thank you for your submission. Your request has been received and assigned to our team.</p>
        <p><strong>Reference:</strong> ${task.id}</p>
        <p><strong>Status:</strong> ${task.status}</p>
        <p>We will keep you updated on the progress.</p>
        `;
      } else if (status === 'rejected') {
        subject = 'Update on your request';
        body = `
        <p>Thank you for your submission. After review, we are unable to process your request at this time.</p>
        ${reason ? `<p><strong>Reason:</strong> ${reason}</p>` : ''}
        <p>If you have questions, please contact support.</p>
        `;
      }
      if (subject && body) {
        GmailApp.sendEmail(email, subject, '', {
          htmlBody: `
          <div style="font-family: sans-serif; max-width: 600px; margin: 0 auto;">
          <div style="background: #f8fafc; padding: 20px; border-radius: 8px; margin-bottom: 20px;">
          <h2 style="color: #f59e0b; margin: 0;">ProjectFlow</h2>
          </div>
          <div style="padding: 20px;">
          ${body}
          </div>
          </div>
          `,
          name: 'ProjectFlow'
        });
      }
    } catch (error) {
    }
  },

  importFromGmail(config = {}) {
    const labelName = config.labelName || 'ProjectFlow/Inbox';
    const maxMessages = config.maxMessages || 50;
    const imported = [];
    try {
      const label = GmailApp.getUserLabelByName(labelName);
      if (!label) {
        return imported;
      }
      const threads = label.getThreads(0, maxMessages);
      for (const thread of threads) {
        const messages = thread.getMessages();
        for (const message of messages) {
          if (message.getThread().getLabels().some(l => l.getName().includes('Processed'))) {
            continue;
          }
          const emailData = {
            messageId: message.getId(),
            from: message.getFrom(),
            subject: message.getSubject(),
            body: message.getPlainBody(),
            receivedAt: message.getDate().toISOString(),
            attachments: message.getAttachments().map(a => ({
              name: a.getName(),
              type: a.getContentType(),
              size: a.getSize()
            }))
          };
          const queueItem = this.submitFromEmail(emailData);
          imported.push(queueItem);
          if (config.markAsRead) {
            message.markRead();
          }
          const processedLabel = GmailApp.getUserLabelByName(labelName + '/Processed') ||
            GmailApp.createLabel(labelName + '/Processed');
          thread.addLabel(processedLabel);
        }
      }
    } catch (error) {
    }
    return imported;
  }
};
