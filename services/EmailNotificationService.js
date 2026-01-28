/**
 * EmailNotificationService - Enhanced Email Delivery System
 * Handles email template management, delivery, and error handling
 */

class EmailNotificationService {

  static sendEmailNotification(notification) {
    try {
      const user = getUserByEmail(notification.userId);
      if (!user || !user.active) {
        throw new Error(`Cannot send email to inactive user: ${notification.userId}`);
      }

      const emailContent = this.formatEmailContent(notification);

      const deliveryResult = this.sendEmailWithRetry(
        user.email,
        emailContent.subject,
        emailContent.htmlBody,
        emailContent.textBody,
        3
      );

      this.logEmailDelivery(notification, user.email, 'success', deliveryResult);

      return {
        success: true,
        recipient: user.email,
        messageId: deliveryResult.messageId,
        deliveredAt: new Date().toISOString()
      };

    } catch (error) {
      this.logEmailDelivery(notification, notification.userId, 'failed', { error: error.message });
      console.error('Failed to send email notification:', error);
      throw error;
    }
  }

  static sendBatchEmailNotifications(notifications) {
    const results = {
      total: notifications.length,
      successful: 0,
      failed: 0,
      results: []
    };

    notifications.forEach((notification, index) => {
      try {
        const result = this.sendEmailNotification(notification);
        results.successful++;
        results.results.push({
          notificationId: notification.id,
          success: true,
          ...result
        });

        if (index < notifications.length - 1) {
          Utilities.sleep(100);
        }

      } catch (error) {
        results.failed++;
        results.results.push({
          notificationId: notification.id,
          success: false,
          error: error.message,
          recipient: notification.userId
        });
      }
    });

    return results;
  }

  static formatEmailContent(notification) {
    const template = this.getEmailTemplate(notification.type);
    const context = this.buildEmailContext(notification);

    const subject = this.processTemplate(template.subject, context);
    const htmlBody = this.processTemplate(template.htmlBody, context);
    const textBody = this.processTemplate(template.textBody, context);

    return {
      subject: subject,
      htmlBody: htmlBody,
      textBody: textBody
    };
  }

  static getEmailTemplate(notificationType) {
    const templates = {
      mention: {
        subject: 'You were mentioned in {{taskTitle}} - ProjectFlow',
        htmlBody: this.getMentionEmailTemplate(),
        textBody: this.getMentionTextTemplate()
      },

      task_assigned: {
        subject: 'New task assigned: {{taskTitle}} - ProjectFlow',
        htmlBody: this.getTaskAssignedEmailTemplate(),
        textBody: this.getTaskAssignedTextTemplate()
      },

      task_updated: {
        subject: 'Task updated: {{taskTitle}} - ProjectFlow',
        htmlBody: this.getTaskUpdatedEmailTemplate(),
        textBody: this.getTaskUpdatedTextTemplate()
      },

      deadline_approaching: {
        subject: '⚠️ Deadline approaching: {{taskTitle}} - ProjectFlow',
        htmlBody: this.getDeadlineEmailTemplate(),
        textBody: this.getDeadlineTextTemplate()
      },

      project_update: {
        subject: 'Project update: {{projectName}} - ProjectFlow',
        htmlBody: this.getProjectUpdateEmailTemplate(),
        textBody: this.getProjectUpdateTextTemplate()
      },

      comment_added: {
        subject: 'New comment on {{taskTitle}} - ProjectFlow',
        htmlBody: this.getCommentEmailTemplate(),
        textBody: this.getCommentTextTemplate()
      }
    };

    return templates[notificationType] || templates.mention;
  }

  static sendMentionNotificationEmail(notification, task, comment, mentionedByUser) {
    try {
      const mentionedUser = getUserByEmail(notification.userId);
      if (!mentionedUser || !mentionedUser.active) {
        throw new Error(`Cannot send email to inactive user: ${notification.userId}`);
      }

      const context = this.buildMentionEmailContext(notification, task, comment, mentionedByUser);
      const template = this.getEmailTemplate('mention');

      const subject = this.processTemplate(template.subject, context);
      const htmlBody = this.processTemplate(template.htmlBody, context);
      const textBody = this.processTemplate(template.textBody, context);

      const deliveryResult = this.sendEmailWithRetry(
        mentionedUser.email,
        subject,
        htmlBody,
        textBody,
        3
      );

      this.logEmailDelivery(notification, mentionedUser.email, 'success', deliveryResult);

      return {
        success: true,
        recipient: mentionedUser.email,
        messageId: deliveryResult.messageId,
        deliveredAt: new Date().toISOString()
      };

    } catch (error) {
      this.logEmailDelivery(notification, notification.userId, 'failed', { error: error.message });
      console.error('Failed to send mention email:', error);

      return {
        success: false,
        recipient: notification.userId,
        error: error.message
      };
    }
  }

  static buildMentionEmailContext(notification, task, comment, mentionedByUser) {
    const mentionedUser = getUserByEmail(notification.userId);
    const project = task.projectId ? getProjectById(task.projectId) : null;

    const context = {
      userName: mentionedUser ? mentionedUser.name : notification.userId,
      userEmail: notification.userId,

      mentionedByName: mentionedByUser.name || mentionedByUser.email,
      mentionedByEmail: mentionedByUser.email,

      message: comment.content || '',
      commentId: comment.id || '',

      taskTitle: task.title || 'Untitled Task',
      taskDescription: task.description || '',
      taskStatus: task.status || 'To Do',
      taskPriority: task.priority || 'Medium',
      taskDueDate: task.dueDate ? this.formatDate(task.dueDate) : 'Not set',
      taskUrl: this.getTaskUrl(task.id),

      projectName: project ? project.name : 'No Project',
      projectDescription: project ? (project.description || '') : '',
      projectUrl: project ? this.getProjectUrl(project.id) : this.getSystemUrl(),

      systemName: 'ProjectFlow',
      systemUrl: this.getSystemUrl(),
      unsubscribeUrl: this.getUnsubscribeUrl(notification.userId),

      formattedDate: this.formatDate(comment.createdAt || notification.createdAt),
      priorityColor: this.getPriorityColor(task.priority),
      statusColor: this.getStatusColor(task.status)
    };

    return context;
  }

  static buildEmailContext(notification) {
    const context = {
      title: notification.title || '',
      message: notification.message || '',
      type: notification.type || '',
      createdAt: notification.createdAt || '',

      userName: '',
      userEmail: notification.userId || '',

      taskTitle: '',
      taskDescription: '',
      taskStatus: '',
      taskPriority: '',
      taskDueDate: '',
      taskUrl: '',

      projectName: '',
      projectDescription: '',
      projectUrl: '',

      systemName: 'ProjectFlow',
      systemUrl: this.getSystemUrl(),
      unsubscribeUrl: this.getUnsubscribeUrl(notification.userId),

      formattedDate: this.formatDate(notification.createdAt),
      priorityColor: '#3b82f6',
      statusColor: '#10b981'
    };

    const user = getUserByEmail(notification.userId);
    if (user) {
      context.userName = user.name;
      context.userEmail = user.email;
    }

    if (notification.entityType === 'task' && notification.entityId) {
      const task = getTaskById(notification.entityId);
      if (task) {
        context.taskTitle = task.title;
        context.taskDescription = task.description || '';
        context.taskStatus = task.status || '';
        context.taskPriority = task.priority || '';
        context.taskDueDate = task.dueDate || 'Not set';
        context.taskUrl = this.getTaskUrl(task.id);

        context.priorityColor = this.getPriorityColor(task.priority);
        context.statusColor = this.getStatusColor(task.status);

        if (task.projectId) {
          const project = getProjectById(task.projectId);
          if (project) {
            context.projectName = project.name;
            context.projectDescription = project.description || '';
            context.projectUrl = this.getProjectUrl(project.id);
          }
        }
      }
    }

    if (notification.entityType === 'project' && notification.entityId) {
      const project = getProjectById(notification.entityId);
      if (project) {
        context.projectName = project.name;
        context.projectDescription = project.description || '';
        context.projectUrl = this.getProjectUrl(project.id);
      }
    }

    return context;
  }

  static processTemplate(template, context) {
    let result = template;

    Object.keys(context).forEach(key => {
      const pattern = new RegExp(`{{${key}}}`, 'g');
      const value = context[key] || '';
      result = result.replace(pattern, value);
    });

    result = result.replace(/{{#(\w+)}}(.*?){{\/\1}}/gs, (match, variable, content) => {
      const value = context[variable];
      return (value && value !== '' && value !== 'Not set') ? content : '';
    });

    result = result.replace(/{{\\^(\w+)}}(.*?){{\/\1}}/gs, (match, variable, content) => {
      const value = context[variable];
      return (!value || value === '' || value === 'Not set') ? content : '';
    });

    result = result.replace(/{{[^}]+}}/g, '');

    return result;
  }

  static sendEmailWithRetry(email, subject, htmlBody, textBody, maxRetries = 3) {
    let attempt = 1;
    let lastError = null;

    while (attempt <= maxRetries) {
      try {
        const result = GmailApp.sendEmail(email, subject, textBody, {
          htmlBody: htmlBody,
          name: 'ProjectFlow Notifications',
          replyTo: 'noreply@projectflow.com'
        });

        return {
          success: true,
          attempt: attempt,
          messageId: `email_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
          sentAt: new Date().toISOString()
        };

      } catch (error) {
        lastError = error;
        console.error(`Email send attempt ${attempt} failed for ${email}:`, error.message);

        if (attempt === maxRetries) {
          throw new Error(`Failed to send email after ${maxRetries} attempts: ${error.message}`);
        }

        const delay = Math.pow(2, attempt) * 1000;
        Utilities.sleep(delay);
        attempt++;
      }
    }

    throw lastError;
  }

  static logEmailDelivery(notification, recipient, status, details) {
    try {
      logActivity('system', 'email_delivery', 'notification', notification.id, {
        recipient: recipient,
        status: status,
        type: notification.type,
        attempt: details.attempt || 1,
        messageId: details.messageId || null,
        error: details.error || null,
        sentAt: details.sentAt || new Date().toISOString()
      });
    } catch (error) {
      console.error('Failed to log email delivery:', error);
    }
  }

  static getSystemUrl() {
    try {
      return ScriptApp.getService().getUrl();
    } catch (error) {
      return 'https://your-projectflow-instance.com';
    }
  }

  static getTaskUrl(taskId) {
    const baseUrl = this.getSystemUrl();
    return `${baseUrl}#/task/${taskId}`;
  }

  static getProjectUrl(projectId) {
    const baseUrl = this.getSystemUrl();
    return `${baseUrl}#/project/${projectId}`;
  }

  static getUnsubscribeUrl(userId) {
    const baseUrl = this.getSystemUrl();
    return `${baseUrl}#/settings/notifications?unsubscribe=${encodeURIComponent(userId)}`;
  }

  static formatDate(dateString) {
    if (!dateString) return '';

    try {
      const date = new Date(dateString);
      return date.toLocaleDateString('en-US', {
        weekday: 'long',
        year: 'numeric',
        month: 'long',
        day: 'numeric',
        hour: '2-digit',
        minute: '2-digit'
      });
    } catch (error) {
      return dateString;
    }
  }

  static getPriorityColor(priority) {
    const colors = {
      'Critical': '#dc2626',
      'Highest': '#ea580c',
      'High': '#f59e0b',
      'Medium': '#3b82f6',
      'Low': '#10b981',
      'Lowest': '#6b7280'
    };

    return colors[priority] || '#3b82f6';
  }

  static getStatusColor(status) {
    const colors = {
      'Backlog': '#6b7280',
      'To Do': '#3b82f6',
      'In Progress': '#f59e0b',
      'Review': '#8b5cf6',
      'Testing': '#06b6d4',
      'Done': '#10b981'
    };

    return colors[status] || '#3b82f6';
  }

  static getMentionEmailTemplate() {
    return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>You were mentioned - ProjectFlow</title>
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; margin: 0; padding: 0; background-color: #f8fafc; }
    .container { max-width: 600px; margin: 0 auto; background-color: white; }
    .header { background: linear-gradient(135deg, #f59e0b 0%, #d97706 100%); color: white; padding: 30px; text-align: center; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1); }
    .header h1 { margin: 0; font-size: 24px; font-weight: 600; }
    .content { padding: 30px; }
    .mention-box { background: #fef3c7; border-left: 4px solid #f59e0b; padding: 20px; margin: 20px 0; border-radius: 0 8px 8px 0; }
    .mention-box strong { color: #92400e; }
    .task-info { background: #f8fafc; padding: 20px; border-radius: 8px; margin: 20px 0; border: 1px solid #e2e8f0; }
    .task-title { font-size: 18px; font-weight: 600; color: #1e293b; margin: 0 0 10px 0; }
    .task-meta { display: flex; gap: 20px; margin: 10px 0; flex-wrap: wrap; }
    .meta-item { font-size: 14px; color: #64748b; }
    .priority { padding: 4px 8px; border-radius: 4px; font-size: 12px; font-weight: 500; color: white; }
    .status { padding: 4px 8px; border-radius: 4px; font-size: 12px; font-weight: 500; color: white; }
    .button { display: inline-block; background: linear-gradient(135deg, #f59e0b 0%, #d97706 100%); color: white; padding: 12px 24px; text-decoration: none; border-radius: 8px; font-weight: 500; margin: 20px 0; box-shadow: 0 2px 4px rgba(245, 158, 11, 0.3); }
    .button:hover { background: linear-gradient(135deg, #d97706 0%, #b45309 100%); }
    .footer { background: #f1f5f9; padding: 20px; text-align: center; font-size: 12px; color: #64748b; }
    .footer a { color: #f59e0b; text-decoration: none; font-weight: 500; }
    .footer a:hover { text-decoration: underline; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>You were mentioned</h1>
      <p style="margin: 10px 0 0 0; opacity: 0.95; font-size: 14px;">Someone mentioned you in a comment</p>
    </div>

    <div class="content">
      <div class="mention-box">
        <p style="margin: 0 0 10px 0;"><strong>{{mentionedByName}}</strong> mentioned you in a comment:</p>
        <p style="font-style: italic; margin: 10px 0; color: #78350f; font-size: 15px;">"{{message}}"</p>
      </div>

      <div class="task-info">
        <div class="task-title">{{taskTitle}}</div>
        {{#taskDescription}}
        <p style="color: #64748b; margin: 10px 0; line-height: 1.5;">{{taskDescription}}</p>
        {{/taskDescription}}

        <div class="task-meta">
          <div class="meta-item">
            <span class="priority" style="background-color: {{priorityColor}};">{{taskPriority}}</span>
          </div>
          <div class="meta-item">
            <span class="status" style="background-color: {{statusColor}};">{{taskStatus}}</span>
          </div>
          {{#taskDueDate}}
          <div class="meta-item">Due: {{taskDueDate}}</div>
          {{/taskDueDate}}
        </div>
      </div>

      <center>
        <a href="{{taskUrl}}" class="button">View Task & Reply</a>
      </center>

      <p style="color: #64748b; font-size: 14px; margin-top: 30px; text-align: center;">
        This mention was posted on {{formattedDate}}<br>
        {{#projectName}}in the <strong>{{projectName}}</strong> project{{/projectName}}
      </p>
    </div>

    <div class="footer">
      <p style="margin: 5px 0;">You're receiving this because you were mentioned in <strong>ProjectFlow</strong>.</p>
      <p style="margin: 10px 0;"><a href="{{unsubscribeUrl}}">Manage notification preferences</a> | <a href="{{systemUrl}}">Open ProjectFlow</a></p>
    </div>
  </div>
</body>
</html>`;
  }

  static getMentionTextTemplate() {
    return `You were mentioned in ProjectFlow

{{mentionedByName}} mentioned you in a comment on "{{taskTitle}}":

"{{message}}"

Task Details:
- Title: {{taskTitle}}
- Status: {{taskStatus}}
- Priority: {{taskPriority}}
{{#taskDueDate}}
- Due Date: {{taskDueDate}}
{{/taskDueDate}}
{{#taskDescription}}
- Description: {{taskDescription}}
{{/taskDescription}}

View and reply to this task: {{taskUrl}}

This mention was posted on {{formattedDate}} in the {{projectName}} project.

---
You're receiving this because you were mentioned in ProjectFlow.
Manage your notification preferences: {{unsubscribeUrl}}`;
  }

  static getTaskAssignedEmailTemplate() {
    return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>New Task Assigned - ProjectFlow</title>
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; margin: 0; padding: 0; background-color: #f8fafc; }
    .container { max-width: 600px; margin: 0 auto; background-color: white; }
    .header { background: linear-gradient(135deg, #10b981 0%, #059669 100%); color: white; padding: 30px; text-align: center; }
    .header h1 { margin: 0; font-size: 24px; font-weight: 600; }
    .content { padding: 30px; }
    .task-card { background: #f8fafc; border: 1px solid #e2e8f0; padding: 25px; border-radius: 12px; margin: 20px 0; }
    .task-title { font-size: 20px; font-weight: 600; color: #1e293b; margin: 0 0 15px 0; }
    .task-meta { display: flex; gap: 15px; margin: 15px 0; flex-wrap: wrap; }
    .meta-badge { padding: 6px 12px; border-radius: 6px; font-size: 12px; font-weight: 500; color: white; }
    .button { display: inline-block; background: #10b981; color: white; padding: 14px 28px; text-decoration: none; border-radius: 8px; font-weight: 500; margin: 25px 0; }
    .footer { background: #f1f5f9; padding: 20px; text-align: center; font-size: 12px; color: #64748b; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>New Task Assigned</h1>
      <p style="margin: 10px 0 0 0; opacity: 0.9;">A new task has been assigned to you</p>
    </div>

    <div class="content">
      <div class="task-card">
        <div class="task-title">{{taskTitle}}</div>

        {{#taskDescription}}
        <p style="color: #475569; line-height: 1.6; margin: 15px 0;">{{taskDescription}}</p>
        {{/taskDescription}}

        <div class="task-meta">
          <span class="meta-badge" style="background-color: {{priorityColor}};">{{taskPriority}} Priority</span>
          <span class="meta-badge" style="background-color: {{statusColor}};">{{taskStatus}}</span>
          {{#taskDueDate}}
          <span class="meta-badge" style="background-color: #f59e0b;">Due {{taskDueDate}}</span>
          {{/taskDueDate}}
        </div>
      </div>

      <a href="{{taskUrl}}" class="button">View Task Details</a>

      <p style="color: #64748b; font-size: 14px; margin-top: 30px;">
        This task was assigned on {{formattedDate}} in the {{projectName}} project.
      </p>
    </div>

    <div class="footer">
      <p>You're receiving this because a task was assigned to you in ProjectFlow.</p>
      <p><a href="{{unsubscribeUrl}}">Manage notification preferences</a> | <a href="{{systemUrl}}">Open ProjectFlow</a></p>
    </div>
  </div>
</body>
</html>`;
  }

  static getTaskAssignedTextTemplate() {
    return `New Task Assigned - ProjectFlow

A new task has been assigned to you:

{{taskTitle}}

{{#taskDescription}}
Description: {{taskDescription}}
{{/taskDescription}}

Task Details:
- Status: {{taskStatus}}
- Priority: {{taskPriority}}
{{#taskDueDate}}
- Due Date: {{taskDueDate}}
{{/taskDueDate}}
- Project: {{projectName}}

View task details: {{taskUrl}}

This task was assigned on {{formattedDate}}.

---
You're receiving this because a task was assigned to you in ProjectFlow.
Manage your notification preferences: {{unsubscribeUrl}}`;
  }

  static getDeadlineEmailTemplate() {
    return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Deadline Approaching - ProjectFlow</title>
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; margin: 0; padding: 0; background-color: #f8fafc; }
    .container { max-width: 600px; margin: 0 auto; background-color: white; }
    .header { background: linear-gradient(135deg, #f59e0b 0%, #d97706 100%); color: white; padding: 30px; text-align: center; }
    .header h1 { margin: 0; font-size: 24px; font-weight: 600; }
    .content { padding: 30px; }
    .warning-box { background: #fef3c7; border: 2px solid #f59e0b; padding: 20px; border-radius: 8px; margin: 20px 0; }
    .task-card { background: #f8fafc; border: 1px solid #e2e8f0; padding: 25px; border-radius: 12px; margin: 20px 0; }
    .deadline { font-size: 18px; font-weight: 600; color: #dc2626; }
    .button { display: inline-block; background: #f59e0b; color: white; padding: 14px 28px; text-decoration: none; border-radius: 8px; font-weight: 500; margin: 25px 0; }
    .footer { background: #f1f5f9; padding: 20px; text-align: center; font-size: 12px; color: #64748b; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>Deadline Approaching</h1>
      <p style="margin: 10px 0 0 0; opacity: 0.9;">A task deadline is coming up soon</p>
    </div>

    <div class="content">
      <div class="warning-box">
        <p style="margin: 0; font-weight: 500;">Your task deadline is approaching!</p>
      </div>

      <div class="task-card">
        <div style="font-size: 20px; font-weight: 600; color: #1e293b; margin: 0 0 15px 0;">{{taskTitle}}</div>

        <div class="deadline">Due: {{taskDueDate}}</div>

        {{#taskDescription}}
        <p style="color: #475569; line-height: 1.6; margin: 15px 0;">{{taskDescription}}</p>
        {{/taskDescription}}

        <div style="margin: 15px 0;">
          <span style="padding: 6px 12px; border-radius: 6px; font-size: 12px; font-weight: 500; color: white; background-color: {{statusColor}};">{{taskStatus}}</span>
          <span style="padding: 6px 12px; border-radius: 6px; font-size: 12px; font-weight: 500; color: white; background-color: {{priorityColor}}; margin-left: 10px;">{{taskPriority}}</span>
        </div>
      </div>

      <a href="{{taskUrl}}" class="button">Update Task Status</a>

      <p style="color: #64748b; font-size: 14px; margin-top: 30px;">
        Don't let deadlines slip by. Update your task status to keep the team informed.
      </p>
    </div>

    <div class="footer">
      <p>You're receiving this because you have a task with an approaching deadline.</p>
      <p><a href="{{unsubscribeUrl}}">Manage notification preferences</a> | <a href="{{systemUrl}}">Open ProjectFlow</a></p>
    </div>
  </div>
</body>
</html>`;
  }

  static getDeadlineTextTemplate() {
    return `Deadline Approaching - ProjectFlow

Your task deadline is approaching:

{{taskTitle}}

Due Date: {{taskDueDate}}
Status: {{taskStatus}}
Priority: {{taskPriority}}

{{#taskDescription}}
Description: {{taskDescription}}
{{/taskDescription}}

Don't let deadlines slip by. Update your task status to keep the team informed.

Update task: {{taskUrl}}

---
You're receiving this because you have a task with an approaching deadline.
Manage your notification preferences: {{unsubscribeUrl}}`;
  }

  static getTaskUpdatedEmailTemplate() {
    return this.getMentionEmailTemplate()
      .replace('You were mentioned', 'Task Updated')
      .replace('Someone mentioned you', 'A task you\'re involved with was updated');
  }

  static getTaskUpdatedTextTemplate() {
    return this.getMentionTextTemplate()
      .replace('You were mentioned', 'Task Updated')
      .replace('mentioned you in a comment', 'updated a task you\'re involved with');
  }

  static getProjectUpdateEmailTemplate() {
    return this.getMentionEmailTemplate()
      .replace('You were mentioned', 'Project Update')
      .replace('Someone mentioned you', 'A project you\'re involved with was updated');
  }

  static getProjectUpdateTextTemplate() {
    return this.getMentionTextTemplate()
      .replace('You were mentioned', 'Project Update')
      .replace('mentioned you in a comment', 'updated a project you\'re involved with');
  }

  static getCommentEmailTemplate() {
    return this.getMentionEmailTemplate()
      .replace('You were mentioned', 'New Comment')
      .replace('Someone mentioned you', 'A new comment was added');
  }

  static getCommentTextTemplate() {
    return this.getMentionTextTemplate()
      .replace('You were mentioned', 'New Comment')
      .replace('mentioned you in a comment', 'added a comment');
  }
}
