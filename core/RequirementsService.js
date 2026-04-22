function getRequirementsTemplates() {
  try {
    return { success: true, templates: (CONFIG && CONFIG.REQ_GATHERING_TEMPLATES) || [] };
  } catch (error) {
    console.error('getRequirementsTemplates failed:', error);
    return { success: false, error: error.message, templates: [] };
  }
}

function getRequirementsForProject(projectId) {
  try {
    var project = getProjectById(projectId);
    if (!project) throw new Error('Project not found: ' + projectId);
    var settings = {};
    try { settings = typeof project.settings === 'string' ? JSON.parse(project.settings) : (project.settings || {}); } catch (e) {}
    return { success: true, requirements: settings.requirements || { templates: [] } };
  } catch (error) {
    console.error('getRequirementsForProject failed:', error);
    return { success: false, error: error.message, requirements: { templates: [] } };
  }
}

function saveRequirementsTemplate(projectId, templateId, data, statusOverride) {
  try {
    PermissionGuard.requirePermission('project:update', { projectId: projectId });
    var project = getProjectById(projectId);
    if (!project) throw new Error('Project not found: ' + projectId);

    var templates = (CONFIG && CONFIG.REQ_GATHERING_TEMPLATES) || [];
    var templateDef = templates.find(function(t) { return t.id === templateId; });
    if (!templateDef) throw new Error('Unknown template: ' + templateId);

    var settings = {};
    try { settings = typeof project.settings === 'string' ? JSON.parse(project.settings) : (project.settings || {}); } catch (e) {}
    if (!settings.requirements || typeof settings.requirements !== 'object') settings.requirements = { templates: [] };
    if (!Array.isArray(settings.requirements.templates)) settings.requirements.templates = [];

    var currentUser = getCurrentUserEmail();
    var existing = settings.requirements.templates.find(function(t) { return t.templateId === templateId; });
    var timestamp = now();
    var cleanData = data && typeof data === 'object' ? data : {};

    if (existing) {
      existing.data = Object.assign({}, existing.data || {}, cleanData);
      existing.updatedAt = timestamp;
      if (statusOverride) existing.status = statusOverride;
      if (existing.status === 'complete' && !existing.completedAt) existing.completedAt = timestamp;
    } else {
      settings.requirements.templates.push({
        templateId: templateId,
        type: templateDef.type,
        name: templateDef.name,
        status: statusOverride || 'draft',
        startedBy: currentUser,
        startedAt: timestamp,
        updatedAt: timestamp,
        completedAt: statusOverride === 'complete' ? timestamp : '',
        data: cleanData
      });
    }

    var projectLevelUpdates = {};
    templateDef.fields.forEach(function(field) {
      if (!field.target) return;
      var value = cleanData[field.key];
      if (value === undefined || value === null || value === '') return;
      if (field.target.indexOf('project.') === 0) {
        var key = field.target.substring('project.'.length);
        projectLevelUpdates[key] = Array.isArray(value) ? value.join(', ') : String(value);
      } else if (field.target.indexOf('settings.') === 0) {
        var sKey = field.target.substring('settings.'.length);
        settings[sKey] = value;
      }
    });

    var updates = Object.assign({}, projectLevelUpdates, { settings: JSON.stringify(settings) });
    var updated = updateProject(projectId, updates);
    try { invalidateProjectCache(projectId); } catch (e) { console.error('invalidateProjectCache failed:', e); }

    return { success: true, project: updated, requirements: settings.requirements };
  } catch (error) {
    console.error('saveRequirementsTemplate failed:', error);
    return { success: false, error: error.message };
  }
}

function getPingOwnerContext(projectId, reason) {
  try {
    var project = getProjectById(projectId);
    if (!project) throw new Error('Project not found: ' + projectId);
    var normalizedReason = reason || 'Update';
    var defaultMessages = {
      'Requirements Gathering': 'Please complete the Requirements Gathering workflow for this project. Open the project link below and click "Start Requirements Gathering" to begin.',
      'Incomplete Documentation': 'Please fill out the Documentation & Links tab for this project. The completeness score is visible at the top of that tab.',
      'Update': 'Please provide a status update on this project at your earliest convenience.'
    };
    var message = defaultMessages[normalizedReason] || 'Action requested on this project.';
    var users = [];
    try {
      users = (getActiveUsersOptimized() || []).map(function(u) {
        return { email: u.email, name: u.name || u.email };
      });
    } catch (e) { console.error('getPingOwnerContext users lookup failed:', e); }
    return {
      success: true,
      project: { id: project.id, name: project.name, ownerId: project.ownerId || '' },
      reason: normalizedReason,
      message: message,
      users: users
    };
  } catch (error) {
    console.error('getPingOwnerContext failed:', error);
    return { success: false, error: error.message };
  }
}

function pingProjectOwner(projectId, reason, payload) {
  try {
    PermissionGuard.requirePermission('project:update', { projectId: projectId });
    var currentUser = getCurrentUserEmail();
    var userRole = getCurrentUserRole ? getCurrentUserRole() : null;
    if (userRole && userRole !== 'admin' && userRole !== 'manager') {
      throw new Error('Only admin or manager can ping project owners.');
    }
    var project = getProjectById(projectId);
    if (!project) throw new Error('Project not found: ' + projectId);
    if (!project.ownerId) throw new Error('Project has no owner to ping.');

    payload = payload || {};
    if (typeof payload === 'string') payload = { message: payload };

    var normalizedReason = reason || 'Update';
    var message = (typeof payload.message === 'string' && payload.message.trim()) ? payload.message.trim() : '';
    if (!message) {
      var defaultMessages = {
        'Requirements Gathering': 'Please complete the Requirements Gathering workflow for this project.',
        'Incomplete Documentation': 'Please fill out the Documentation & Links fields for this project.',
        'Update': 'Please provide a status update on this project.'
      };
      message = defaultMessages[normalizedReason] || 'Action requested on this project.';
    }

    var ownerEmail = String(project.ownerId).trim();
    var ccList = [];
    if (Array.isArray(payload.cc)) {
      var seen = {};
      seen[ownerEmail.toLowerCase()] = true;
      payload.cc.forEach(function(e) {
        var email = (e || '').toString().trim();
        var key = email.toLowerCase();
        if (!email || !key || seen[key]) return;
        seen[key] = true;
        ccList.push(email);
      });
    }

    try {
      NotificationEngine.createNotification({
        userId: ownerEmail,
        type: 'project_ping',
        title: 'Action requested: ' + normalizedReason,
        message: message,
        entityType: 'project',
        entityId: projectId,
        channels: ['in_app'],
        pingReason: normalizedReason,
        reason: 'Requested by ' + (currentUser || 'admin')
      });
    } catch (e) {
      console.error('pingProjectOwner: in-app notification for owner failed:', e);
    }

    ccList.forEach(function(ccEmail) {
      try {
        NotificationEngine.createNotification({
          userId: ccEmail,
          type: 'project_ping',
          title: 'CC: Action requested on ' + (project.name || project.id),
          message: message,
          entityType: 'project',
          entityId: projectId,
          channels: ['in_app'],
          pingReason: normalizedReason,
          reason: 'CC from ' + (currentUser || 'admin')
        });
      } catch (e) {
        console.error('pingProjectOwner: in-app notification for CC ' + ccEmail + ' failed:', e);
      }
    });

    var emailResult = sendPingOwnerEmail_(project, ownerEmail, ccList, normalizedReason, message, currentUser);
    if (!emailResult.success) {
      return {
        success: false,
        error: 'In-app notifications sent, but email delivery failed: ' + (emailResult.error || 'unknown'),
        to: ownerEmail,
        cc: ccList
      };
    }

    try {
      logActivity(currentUser, 'pinged', 'project', projectId, {
        reason: normalizedReason,
        cc: ccList,
        messagePreview: message.substring(0, 140)
      });
    } catch (e) {}

    return { success: true, to: ownerEmail, cc: ccList };
  } catch (error) {
    console.error('pingProjectOwner failed:', error);
    return { success: false, error: error.message };
  }
}

function sendPingOwnerEmail_(project, ownerEmail, ccList, reason, message, currentUser) {
  try {
    var projectUrl = '';
    var tabForReason = '';
    if (reason === 'Requirements Gathering') tabForReason = 'requirements';
    else if (reason === 'Incomplete Documentation') tabForReason = 'documentation';
    try {
      projectUrl = NotificationEngine.getProjectUrl(project.id, tabForReason ? { tab: tabForReason } : null);
    } catch (e) {}
    if (!projectUrl) {
      try {
        projectUrl = ScriptApp.getService().getUrl() + '?projectId=' + encodeURIComponent(project.id);
        if (tabForReason) projectUrl += '&tab=' + encodeURIComponent(tabForReason);
      } catch (e) {}
    }

    var projectName = project.name || project.id || '';
    var subject = 'Action Required: ' + reason + ' for ' + projectName + ' - COLONY';
    var plainMessage = sanitizePlainText_(message);
    var senderLine = currentUser ? ('Requested by ' + sanitizePlainText_(currentUser)) : '';

    var htmlBody =
      '<div style="font-family: -apple-system, BlinkMacSystemFont, sans-serif; max-width: 600px; margin: 0 auto; background: #f8fafc; padding: 20px;">' +
      '<div style="background: white; padding: 24px; border-radius: 8px; border: 1px solid #e2e8f0;">' +
      '<div style="background: #fef3c7; color: #92400e; padding: 10px 14px; border-radius: 6px; margin: 0 0 16px 0; font-weight: 600;">' +
        'Action requested: ' + escapeHtml_(reason) +
      '</div>' +
      '<h3 style="color: #3b82f6; margin: 0 0 12px 0;">' + escapeHtml_(projectName) + '</h3>' +
      '<div style="color: #334155; white-space: pre-wrap; line-height: 1.5;">' + escapeHtml_(plainMessage) + '</div>' +
      (senderLine ? '<p style="font-size: 12px; color: #64748b; margin-top: 16px;">' + escapeHtml_(senderLine) + '</p>' : '') +
      '<p style="margin: 20px 0 0 0;">' +
        '<a href="' + escapeHtml_(projectUrl) + '" style="background: #3b82f6; color: white; padding: 10px 20px; text-decoration: none; border-radius: 6px; display: inline-block; font-weight: 500;">' +
          'Open Project' +
        '</a>' +
      '</p>' +
      '</div>' +
      '<div style="margin-top: 20px; padding: 12px 16px; font-size: 12px; color: #64748b; text-align: center;">' +
        '<p style="margin: 0;">COLONY &middot; Action requested on project ' + escapeHtml_(projectName) + '</p>' +
      '</div>' +
      '</div>';

    var textBody =
      'COLONY - Action requested: ' + reason + '\n\n' +
      'Project: ' + projectName + '\n\n' +
      plainMessage + '\n\n' +
      (senderLine ? senderLine + '\n\n' : '') +
      'Open the project: ' + projectUrl + '\n';

    var sendOptions = {
      htmlBody: htmlBody,
      name: 'COLONY.SYSTEM',
      replyTo: currentUser || 'noreply@projectflow.com'
    };
    if (ccList && ccList.length > 0) sendOptions.cc = ccList.join(',');

    var quotaRemaining = -1;
    try { quotaRemaining = MailApp.getRemainingDailyQuota(); } catch (e) {}
    if (quotaRemaining === 0) {
      return { success: false, error: 'Gmail daily quota exhausted; try again tomorrow.' };
    }

    GmailApp.sendEmail(ownerEmail, subject, textBody, sendOptions);
    return { success: true };
  } catch (error) {
    console.error('sendPingOwnerEmail_ failed:', error);
    return { success: false, error: error.message };
  }
}

function sanitizePlainText_(str) {
  if (typeof str !== 'string') return '';
  return str.replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '');
}
