class NotificationEngine {
  static createNotification(notificationData) {
    if (!notificationData.userId) {
      throw new Error('userId is required for notification creation');
    }
    if (!notificationData.type) {
      throw new Error('type is required for notification creation');
    }
    const userPreferences = this.getUserNotificationPreferences(notificationData.userId);
    const typePreferences = userPreferences[notificationData.type] || {};
    let channels = notificationData.channels || ['in_app'];
    if (Array.isArray(channels)) {
      channels = channels.filter(channel => {
        return typePreferences[channel] !== false;
      });
    }
    if (channels.length === 0) {
      return null;
    }
    const notification = createNotification({
      userId: notificationData.userId,
      type: notificationData.type,
      title: notificationData.title || this.getDefaultTitle(notificationData.type),
      message: notificationData.message || '',
      entityType: notificationData.entityType || '',
      entityId: notificationData.entityId || '',
      channels: channels,
      scheduledFor: notificationData.scheduledFor || now()
    });
    if (notificationData.reason) notification.reason = notificationData.reason;
    if (notificationData.pingReason) notification.pingReason = notificationData.pingReason;
    if (notificationData.customMessage) notification.customMessage = notificationData.customMessage;
    if (Array.isArray(notificationData.ccRecipients) && notificationData.ccRecipients.length > 0) {
      notification.ccRecipients = notificationData.ccRecipients.slice();
    }
    const scheduledTime = new Date(notification.scheduledFor);
    const currentTime = new Date();
    if (scheduledTime <= currentTime) {
      this.queueNotificationForDelivery(notification);
    }
    return notification;
  }

  static createBatchNotifications(notifications) {
    if (!Array.isArray(notifications)) {
      throw new Error('notifications must be an array');
    }
    const createdNotifications = [];
    const deliveryQueue = [];
    notifications.forEach(notificationData => {
      try {
        const notification = this.createNotification(notificationData);
        if (notification) {
          createdNotifications.push(notification);
          const scheduledTime = new Date(notification.scheduledFor);
          const currentTime = new Date();
          if (scheduledTime <= currentTime) {
            deliveryQueue.push(notification);
          }
        }
      } catch (error) {
      }
    });
    if (deliveryQueue.length > 0) {
      this.processBatchDelivery(deliveryQueue);
    }
    return createdNotifications;
  }

  static getUserNotificationPreferences(userId) {
    const defaultPreferences = {
      mention: {
        email: true,
        in_app: true,
        push: false
      },
      task_assigned: {
        email: true,
        in_app: true,
        push: false
      },
      task_updated: {
        email: true,
        in_app: true,
        push: false
      },
      deadline_approaching: {
        email: true,
        in_app: true,
        push: true
      },
      project_update: {
        email: false,
        in_app: true,
        push: false
      },
      project_assigned: {
        email: true,
        in_app: true,
        push: false
      },
      project_ping: {
        email: true,
        in_app: true,
        push: false
      },
      comment_added: {
        email: true,
        in_app: true,
        push: false
      }
    };
    return defaultPreferences;
  }

  static updateUserNotificationPreferences(userId, preferences) {
    const validTypes = CONFIG.NOTIFICATION_TYPES;
    const validChannels = CONFIG.NOTIFICATION_CHANNELS;
    Object.keys(preferences).forEach(type => {
      if (!validTypes.includes(type)) {
        throw new Error(`Invalid notification type: ${type}`);
      }
      Object.keys(preferences[type]).forEach(channel => {
        if (!validChannels.includes(channel)) {
          throw new Error(`Invalid notification channel: ${channel}`);
        }
      });
    });
    return preferences;
  }

  static queueNotificationForDelivery(notification) {
    try {
      var channels = notification.channels;
      if (typeof channels === 'string') {
        channels = channels.split(',').map(function(c) { return c.trim(); }).filter(Boolean);
      } else if (!Array.isArray(channels)) {
        channels = ['in_app'];
      }
      channels.forEach(channel => {
        switch (channel) {
          case 'email':
            this.queueEmailNotification(notification);
            break;
          case 'in_app':
            break;
          case 'push':
            this.queuePushNotification(notification);
            break;
          default:
        }
      });
    } catch (error) {
      console.error('queueNotificationForDelivery failed:', error);
    }
  }

  static queueEmailNotification(notification) {
    try {
      EmailNotificationService.sendEmailNotification(notification);
    } catch (error) {
      console.error('queueEmailNotification failed:', error);
      try {
        logActivity('system', 'email_failed', 'notification', notification.id, {
          recipient: notification.userId,
          error: error.message,
          retryable: true
        });
      } catch (e) {}
    }
  }

  static queuePushNotification(notification) {
  }

  static processBatchDelivery(notifications) {
    const channelGroups = {
      email: [],
      in_app: [],
      push: []
    };
    notifications.forEach(notification => {
      var channels = notification.channels;
      if (typeof channels === 'string') {
        channels = channels.split(',').map(function(c) { return c.trim(); }).filter(Boolean);
      } else if (!Array.isArray(channels)) {
        channels = ['in_app'];
      }
      channels.forEach(channel => {
        if (channelGroups[channel]) {
          channelGroups[channel].push(notification);
        }
      });
    });
    if (channelGroups.email.length > 0) {
      try {
        EmailNotificationService.sendBatchEmailNotifications(channelGroups.email);
      } catch (error) {
        console.error('sendBatchEmailNotifications failed:', error);
      }
    }
    ['in_app', 'push'].forEach(channel => {
      const channelNotifications = channelGroups[channel];
      if (channelNotifications.length > 0) {
        channelNotifications.forEach(notification => {
          if (channel === 'in_app') {
          } else if (channel === 'push') {
            this.queuePushNotification(notification);
          }
        });
      }
    });
  }

  static sendEmailWithRetry(email, subject, body, maxRetries = 3) {
    let attempt = 1;
    while (attempt <= maxRetries) {
      try {
        GmailApp.sendEmail(email, subject, '', {
          htmlBody: body,
          name: 'COLONY Notifications'
        });
        return;
      } catch (error) {
        if (attempt === maxRetries) {
          throw new Error(`Failed to send email after ${maxRetries} attempts: ${error.message}`);
        }
        const delay = Math.pow(2, attempt) * 1000;
        Utilities.sleep(delay);
        attempt++;
      }
    }
  }

  static formatEmailNotification(notification) {
    const templates = this.getEmailTemplates();
    const template = templates[notification.type] || templates.default;
    const contextData = this.getNotificationContext(notification);
    const subject = this.replaceTemplateVariables(template.subject, notification, contextData);
    const body = this.replaceTemplateVariables(template.body, notification, contextData);
    return {
      subject: subject,
      body: body
    };
  }

  static getEmailTemplates() {
    return {
      mention: {
        subject: 'You were mentioned in COLONY',
        body: `
        <div style="font-family: -apple-system, BlinkMacSystemFont, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: #f8fafc; padding: 20px; border-radius: 8px; margin-bottom: 20px;">
        <h2 style="color: #1e293b; margin: 0;">You were mentioned in a comment</h2>
        </div>
        <div style="background: white; padding: 20px; border: 1px solid #e2e8f0; border-radius: 8px;">
        <p><strong>{{mentionedByName}}</strong> mentioned you in a comment on task:</p>
        <h3 style="color: #3b82f6; margin: 10px 0;">{{taskTitle}}</h3>
        <div style="background: #f1f5f9; padding: 15px; border-radius: 6px; margin: 15px 0;">
        <p style="margin: 0; font-style: italic;">"{{commentPreview}}"</p>
        </div>
        <p>
        <a href="{{taskUrl}}" style="background: #3b82f6; color: white; padding: 10px 20px; text-decoration: none; border-radius: 6px; display: inline-block;">
        View Task
        </a>
        </p>
        </div>
        <div style="margin-top: 20px; padding: 15px; background: #f8fafc; border-radius: 6px; font-size: 12px; color: #64748b;">
        <p>This notification was sent by COLONY. To manage your notification preferences, visit your account settings.</p>
        </div>
        </div>
        `
      },
      task_assigned: {
        subject: 'New task assigned to you in COLONY',
        body: `
        <div style="font-family: -apple-system, BlinkMacSystemFont, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: #f8fafc; padding: 20px; border-radius: 8px; margin-bottom: 20px;">
        <h2 style="color: #1e293b; margin: 0;">New task assigned to you</h2>
        </div>
        <div style="background: white; padding: 20px; border: 1px solid #e2e8f0; border-radius: 8px;">
        <h3 style="color: #3b82f6; margin: 0 0 10px 0;">{{taskTitle}}</h3>
        <p><strong>Priority:</strong> {{taskPriority}}</p>
        <p><strong>Due Date:</strong> {{taskDueDate}}</p>
        {{#taskDescription}}
        <div style="background: #f1f5f9; padding: 15px; border-radius: 6px; margin: 15px 0;">
        <p style="margin: 0;">{{taskDescription}}</p>
        </div>
        {{/taskDescription}}
        <p>
        <a href="{{taskUrl}}" style="background: #3b82f6; color: white; padding: 10px 20px; text-decoration: none; border-radius: 6px; display: inline-block;">
        View Task
        </a>
        </p>
        </div>
        <div style="margin-top: 20px; padding: 15px; background: #f8fafc; border-radius: 6px; font-size: 12px; color: #64748b;">
        <p>This notification was sent by COLONY. To manage your notification preferences, visit your account settings.</p>
        </div>
        </div>
        `
      },
      deadline_approaching: {
        subject: 'Task deadline approaching in COLONY',
        body: `
        <div style="font-family: -apple-system, BlinkMacSystemFont, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: #f5f5f5; padding: 20px; border-radius: 8px; margin-bottom: 20px;">
        <h2 style="color: #262626; margin: 0;">Task deadline approaching</h2>
        </div>
        <div style="background: white; padding: 20px; border: 1px solid #e2e8f0; border-radius: 8px;">
        <h3 style="color: #525252; margin: 0 0 10px 0;">{{taskTitle}}</h3>
        <p><strong>Due Date:</strong> <span style="color: #dc2626;">{{taskDueDate}}</span></p>
        <p><strong>Priority:</strong> {{taskPriority}}</p>
        <p><strong>Status:</strong> {{taskStatus}}</p>
        <p>
        <a href="{{taskUrl}}" style="background: #525252; color: white; padding: 10px 20px; text-decoration: none; border-radius: 6px; display: inline-block;">
        View Task
        </a>
        </p>
        </div>
        <div style="margin-top: 20px; padding: 15px; background: #f8fafc; border-radius: 6px; font-size: 12px; color: #64748b;">
        <p>This notification was sent by COLONY. To manage your notification preferences, visit your account settings.</p>
        </div>
        </div>
        `
      },
      task_updated: {
        subject: 'Task updated in COLONY: {{taskTitle}}',
        body: `
        <div style="font-family: -apple-system, BlinkMacSystemFont, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: #f8fafc; padding: 20px; border-radius: 8px; margin-bottom: 20px;">
        <h2 style="color: #1e293b; margin: 0;">Task Updated</h2>
        </div>
        <div style="background: white; padding: 20px; border: 1px solid #e2e8f0; border-radius: 8px;">
        <h3 style="color: #3b82f6; margin: 0 0 10px 0;">{{taskTitle}}</h3>
        <p>{{message}}</p>
        <p><strong>Status:</strong> {{taskStatus}}</p>
        <p><strong>Priority:</strong> {{taskPriority}}</p>
        {{#reason}}<p style="font-size: 12px; color: #64748b;">{{reason}}</p>{{/reason}}
        <p>
        <a href="{{taskUrl}}" style="background: #3b82f6; color: white; padding: 10px 20px; text-decoration: none; border-radius: 6px; display: inline-block;">
        View Task
        </a>
        </p>
        </div>
        <div style="margin-top: 20px; padding: 15px; background: #f8fafc; border-radius: 6px; font-size: 12px; color: #64748b;">
        <p>This notification was sent by COLONY. To manage your notification preferences, visit your account settings.</p>
        </div>
        </div>
        `
      },
      project_assigned: {
        subject: 'New project assigned to you in COLONY',
        body: `
        <div style="font-family: -apple-system, BlinkMacSystemFont, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: #f8fafc; padding: 20px; border-radius: 8px; margin-bottom: 20px;">
        <h2 style="color: #1e293b; margin: 0;">New project assigned to you</h2>
        </div>
        <div style="background: white; padding: 20px; border: 1px solid #e2e8f0; border-radius: 8px;">
        <h3 style="color: #3b82f6; margin: 0 0 10px 0;">{{projectName}}</h3>
        <p>{{message}}</p>
        {{#projectDescription}}
        <div style="background: #f1f5f9; padding: 15px; border-radius: 6px; margin: 15px 0;">
        <p style="margin: 0;">{{projectDescription}}</p>
        </div>
        {{/projectDescription}}
        {{#reason}}<p style="font-size: 12px; color: #64748b;">{{reason}}</p>{{/reason}}
        <p>
        <a href="{{projectUrl}}" style="background: #3b82f6; color: white; padding: 10px 20px; text-decoration: none; border-radius: 6px; display: inline-block;">
        View Project
        </a>
        </p>
        </div>
        <div style="margin-top: 20px; padding: 15px; background: #f8fafc; border-radius: 6px; font-size: 12px; color: #64748b;">
        <p>This notification was sent by COLONY. To manage your notification preferences, visit your account settings.</p>
        </div>
        </div>
        `
      },
      project_ping: {
        subject: 'Action Required: {{pingReason}} for {{projectName}} - COLONY',
        body: `
        <div style="font-family: -apple-system, BlinkMacSystemFont, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: #fef3c7; padding: 20px; border-radius: 8px; margin-bottom: 20px;">
        <h2 style="color: #92400e; margin: 0;">Action requested: {{pingReason}}</h2>
        </div>
        <div style="background: white; padding: 20px; border: 1px solid #e2e8f0; border-radius: 8px;">
        <h3 style="color: #3b82f6; margin: 0 0 10px 0;">{{projectName}}</h3>
        <p>{{message}}</p>
        {{#customMessage}}
        <div style="background: #f1f5f9; padding: 15px; border-radius: 6px; margin: 15px 0; border-left: 3px solid #3b82f6;">
        <p style="margin: 0;">{{customMessage}}</p>
        </div>
        {{/customMessage}}
        <p>
        <a href="{{projectUrl}}" style="background: #3b82f6; color: white; padding: 10px 20px; text-decoration: none; border-radius: 6px; display: inline-block;">
        Open Project
        </a>
        </p>
        </div>
        <div style="margin-top: 20px; padding: 15px; background: #f8fafc; border-radius: 6px; font-size: 12px; color: #64748b;">
        <p>This notification was sent by COLONY. To manage your notification preferences, visit your account settings.</p>
        </div>
        </div>
        `
      },
      comment_added: {
        subject: 'New comment on task in COLONY: {{taskTitle}}',
        body: `
        <div style="font-family: -apple-system, BlinkMacSystemFont, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: #f8fafc; padding: 20px; border-radius: 8px; margin-bottom: 20px;">
        <h2 style="color: #1e293b; margin: 0;">New Comment</h2>
        </div>
        <div style="background: white; padding: 20px; border: 1px solid #e2e8f0; border-radius: 8px;">
        <h3 style="color: #3b82f6; margin: 0 0 10px 0;">{{taskTitle}}</h3>
        <p>{{message}}</p>
        {{#reason}}<p style="font-size: 12px; color: #64748b;">{{reason}}</p>{{/reason}}
        <p>
        <a href="{{taskUrl}}" style="background: #3b82f6; color: white; padding: 10px 20px; text-decoration: none; border-radius: 6px; display: inline-block;">
        View Task
        </a>
        </p>
        </div>
        <div style="margin-top: 20px; padding: 15px; background: #f8fafc; border-radius: 6px; font-size: 12px; color: #64748b;">
        <p>This notification was sent by COLONY. To manage your notification preferences, visit your account settings.</p>
        </div>
        </div>
        `
      },
      default: {
        subject: 'COLONY Notification',
        body: `
        <div style="font-family: -apple-system, BlinkMacSystemFont, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: #f8fafc; padding: 20px; border-radius: 8px; margin-bottom: 20px;">
        <h2 style="color: #1e293b; margin: 0;">{{title}}</h2>
        </div>
        <div style="background: white; padding: 20px; border: 1px solid #e2e8f0; border-radius: 8px;">
        <p>{{message}}</p>
        {{#entityUrl}}
        <p>
        <a href="{{entityUrl}}" style="background: #3b82f6; color: white; padding: 10px 20px; text-decoration: none; border-radius: 6px; display: inline-block;">
        View Details
        </a>
        </p>
        {{/entityUrl}}
        </div>
        <div style="margin-top: 20px; padding: 15px; background: #f8fafc; border-radius: 6px; font-size: 12px; color: #64748b;">
        <p>This notification was sent by COLONY. To manage your notification preferences, visit your account settings.</p>
        </div>
        </div>
        `
      }
    };
  }

  static getNotificationContext(notification) {
    const context = {};
    if (notification.entityType === 'task' && notification.entityId) {
      const task = getTaskById(notification.entityId);
      if (task) {
        context.taskTitle = task.title;
        context.taskDescription = task.description;
        context.taskPriority = task.priority;
        context.taskStatus = task.status;
        context.taskDueDate = task.dueDate || 'Not set';
        context.taskUrl = this.getTaskUrl(task.id);
        if (task.projectId) {
          try {
            const linkedProject = getProjectById(task.projectId);
            if (linkedProject) {
              context.projectName = linkedProject.name;
              context.projectUrl = this.getProjectUrl(linkedProject.id);
            }
          } catch (e) {}
        }
      }
    }
    if (notification.entityType === 'project' && notification.entityId) {
      const project = getProjectById(notification.entityId);
      if (project) {
        context.projectName = project.name;
        context.projectDescription = project.description;
        context.projectUrl = this.getProjectUrl(project.id);
      }
    }
    const user = getUserByEmail(notification.userId);
    if (user) {
      context.userName = user.name;
      context.userEmail = user.email;
    }
    context.reason = notification.reason || '';
    context.pingReason = notification.pingReason || notification.reason || '';
    context.customMessage = notification.customMessage || '';
    return context;
  }

  static replaceTemplateVariables(template, notification, context = {}) {
    let result = template;
    const data = {
      ...notification,
      ...context
    };
    Object.keys(data).forEach(key => {
      const pattern = new RegExp(`{{${key}}}`, 'g');
      result = result.replace(pattern, data[key] || '');
    });
    result = result.replace(/{{#(\w+)}}(.*?){{\/\1}}/gs, (match, variable, content) => {
      return data[variable] ? content : '';
    });
    result = result.replace(/{{[^}]+}}/g, '');
    return result;
  }

  static getDefaultTitle(type) {
    const titles = {
      mention: 'You were mentioned',
      task_assigned: 'New task assigned',
      task_updated: 'Task updated',
      deadline_approaching: 'Deadline approaching',
      project_update: 'Project updated',
      project_assigned: 'Project assigned',
      project_ping: 'Action requested on project',
      comment_added: 'New comment added'
    };
    return titles[type] || 'COLONY Notification';
  }

  static getTaskUrl(taskId) {
    const webAppUrl = ScriptApp.getService().getUrl();
    return `${webAppUrl}?taskId=${encodeURIComponent(taskId)}`;
  }

  static getProjectUrl(projectId, options) {
    const webAppUrl = ScriptApp.getService().getUrl();
    if (typeof projectId === 'string' && projectId.indexOf('DA_') === 0) {
      return `${webAppUrl}?dataAssetId=${encodeURIComponent(projectId)}`;
    }
    var url = `${webAppUrl}?projectId=${encodeURIComponent(projectId)}`;
    if (options && options.tab) url += `&tab=${encodeURIComponent(options.tab)}`;
    return url;
  }

  static processPendingNotifications(batchSize = 50) {
    const pendingNotifications = getPendingNotifications(batchSize);
    if (pendingNotifications.length === 0) {
      return 0;
    }
    pendingNotifications.forEach(notification => {
      try {
        this.queueNotificationForDelivery(notification);
      } catch (error) {
      }
    });
    return pendingNotifications.length;
  }

  static getNotificationStatistics(userId) {
    const notifications = getNotificationsForUser(userId, 1000);
    const stats = {
      total: notifications.length,
      unread: notifications.filter(n => !n.read).length,
      byType: {},
      byChannel: {},
      recent: notifications.slice(0, 10)
    };
    notifications.forEach(notification => {
      stats.byType[notification.type] = (stats.byType[notification.type] || 0) + 1;
      notification.channels.forEach(channel => {
        stats.byChannel[channel] = (stats.byChannel[channel] || 0) + 1;
      });
    });
    return stats;
  }
}
