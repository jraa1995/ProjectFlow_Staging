var MaintenanceService = {
  PROP_KEY: 'COLONY_MAINTENANCE_STATUS',
  CACHE_KEY: 'maintenance_status_v1',
  CACHE_TTL: 5,
  ALLOWED_STATUSES: ['none', 'maintenance', 'degraded', 'down', 'resolved'],
  RECIPIENT_CHUNK: 90,

  getStatus: function() {
    try {
      var cache = CacheService.getScriptCache();
      var cached = cache.get(this.CACHE_KEY);
      if (cached) {
        try { return JSON.parse(cached); } catch (e) {}
      }
      var props = PropertiesService.getScriptProperties();
      var raw = props.getProperty(this.PROP_KEY);
      var status = this._defaultStatus();
      if (raw) {
        try {
          var parsed = JSON.parse(raw);
          if (parsed && typeof parsed === 'object') status = parsed;
        } catch (e) {
          console.error('MaintenanceService.getStatus parse failed:', e);
        }
      }
      cache.put(this.CACHE_KEY, JSON.stringify(status), this.CACHE_TTL);
      return status;
    } catch (error) {
      console.error('MaintenanceService.getStatus failed:', error);
      return this._defaultStatus();
    }
  },

  setStatus: function(payload) {
    PermissionGuard.requirePermission('admin:settings');
    payload = payload || {};
    var status = String(payload.status || '').toLowerCase().trim();
    if (this.ALLOWED_STATUSES.indexOf(status) === -1) {
      throw new Error('Invalid status: ' + status);
    }
    var message = String(payload.message || '').trim().slice(0, 1000);
    var notifyEmail = payload.notifyEmail === true;

    var actor = '';
    var actorName = '';
    try {
      actor = getCurrentUserEmailOptimized() || '';
      var u = actor ? getUserByEmail(actor) : null;
      actorName = u ? (u.name || u.email) : actor;
    } catch (e) {}

    var record = {
      status: status,
      message: message,
      setAt: new Date().toISOString(),
      setBy: actor,
      setByName: actorName,
      versionId: 'M-' + Date.now() + '-' + Math.random().toString(36).slice(2, 8)
    };

    var props = PropertiesService.getScriptProperties();
    if (status === 'none') {
      props.deleteProperty(this.PROP_KEY);
    } else {
      props.setProperty(this.PROP_KEY, JSON.stringify(record));
    }
    try { CacheService.getScriptCache().remove(this.CACHE_KEY); } catch (e) {}

    var emailResult = { sent: 0, skipped: 0, attempted: 0, recipients: 0 };
    if (notifyEmail && status !== 'none') {
      try { emailResult = this._broadcastEmail(record); }
      catch (e) {
        console.error('MaintenanceService broadcast failed:', e);
        emailResult.error = e.message;
      }
    }

    try {
      logAuditEvent(
        status === 'none' ? 'maintenance.clear' : 'maintenance.set',
        'system',
        'maintenance',
        status,
        { status: status, notifyEmail: notifyEmail, recipients: emailResult.recipients, sent: emailResult.sent, message: message }
      );
    } catch (e) {}

    return { success: true, status: record, email: emailResult };
  },

  clearStatus: function(notifyEmail) {
    return this.setStatus({ status: 'none', message: '', notifyEmail: notifyEmail === true });
  },

  _defaultStatus: function() {
    return { status: 'none', message: '', setAt: '', setBy: '', setByName: '', versionId: '' };
  },

  _broadcastEmail: function(record) {
    var users;
    try { users = getActiveUsers(); }
    catch (e) {
      console.error('MaintenanceService getActiveUsers failed:', e);
      users = [];
    }
    var recipients = users
      .map(function(u) { return (u && u.email) ? String(u.email).trim() : ''; })
      .filter(function(e) { return e && e.indexOf('@') > 0; });

    var result = { sent: 0, skipped: 0, attempted: 0, recipients: recipients.length, chunks: 0 };
    if (recipients.length === 0) return result;

    var subject = this._subjectFor(record.status);
    var html = this._htmlFor(record);
    var text = this._textFor(record);
    var fromAddr = '';
    try { fromAddr = getCurrentUserEmailOptimized() || ''; } catch (e) {}
    if (!fromAddr) {
      try { fromAddr = Session.getActiveUser().getEmail() || ''; } catch (e) {}
    }
    if (!fromAddr && recipients.length > 0) fromAddr = recipients[0];

    for (var i = 0; i < recipients.length; i += this.RECIPIENT_CHUNK) {
      var chunk = recipients.slice(i, i + this.RECIPIENT_CHUNK);
      result.attempted += chunk.length;
      try {
        GmailApp.sendEmail(fromAddr, subject, text, {
          bcc: chunk.join(','),
          htmlBody: html,
          name: 'COLONY.SYSTEM',
          replyTo: 'noreply@projectflow.com'
        });
        result.sent += chunk.length;
        result.chunks += 1;
      } catch (e) {
        console.error('MaintenanceService chunk send failed:', e);
        result.skipped += chunk.length;
      }
    }
    return result;
  },

  _subjectFor: function(status) {
    switch (status) {
      case 'maintenance': return 'COLONY — Scheduled Maintenance';
      case 'degraded':    return 'COLONY — Degraded Service';
      case 'down':        return 'COLONY — Service Outage';
      case 'resolved':    return 'COLONY — Service Restored';
      default:            return 'COLONY — System Notice';
    }
  },

  _accentFor: function(status) {
    switch (status) {
      case 'maintenance': return { bg: '#fef3c7', border: '#f59e0b', ink: '#92400e', label: 'Scheduled Maintenance' };
      case 'degraded':    return { bg: '#ffedd5', border: '#f97316', ink: '#9a3412', label: 'Degraded Service' };
      case 'down':        return { bg: '#fee2e2', border: '#ef4444', ink: '#991b1b', label: 'Service Outage' };
      case 'resolved':    return { bg: '#d1fae5', border: '#10b981', ink: '#065f46', label: 'Service Restored' };
      default:            return { bg: '#f1f5f9', border: '#64748b', ink: '#1e293b', label: 'System Notice' };
    }
  },

  _htmlFor: function(record) {
    var accent = this._accentFor(record.status);
    var _esc = function(v) {
      return String(v == null ? '' : v)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#x27;');
    };
    var safeMsg = _esc(record.message || '').replace(/\n/g, '<br>');
    var url = '';
    try { url = ScriptApp.getService().getUrl() || ''; } catch (e) {}
    var setByLine = record.setByName ? 'Posted by ' + _esc(record.setByName) : '';
    var when = '';
    try { when = new Date(record.setAt).toLocaleString('en-US', { dateStyle: 'medium', timeStyle: 'short' }); } catch (e) { when = record.setAt || ''; }

    return '<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f8fafc;font-family:-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,sans-serif;">' +
      '<div style="max-width:560px;margin:0 auto;padding:24px 16px;">' +
      '<div style="background:#fff;border-radius:12px;border:1px solid #e2e8f0;overflow:hidden;">' +
      '<div style="background:' + accent.bg + ';border-left:4px solid ' + accent.border + ';padding:16px 20px;">' +
      '<div style="font-size:11px;letter-spacing:0.08em;text-transform:uppercase;color:' + accent.ink + ';font-weight:600;">' + accent.label + '</div>' +
      '<div style="font-size:18px;color:' + accent.ink + ';font-weight:600;margin-top:4px;">COLONY ' +
      (record.status === 'resolved' ? 'is back online' : (record.status === 'down' ? 'is currently unavailable' : (record.status === 'maintenance' ? 'is undergoing maintenance' : 'is experiencing issues'))) +
      '</div></div>' +
      '<div style="padding:20px;color:#1e293b;font-size:14px;line-height:1.55;">' +
      (safeMsg ? '<div style="white-space:pre-wrap;">' + safeMsg + '</div>' : '<div style="color:#64748b;">No additional details were provided.</div>') +
      '<div style="margin-top:18px;font-size:12px;color:#64748b;">' + (when ? when : '') + (setByLine ? ' &middot; ' + setByLine : '') + '</div>' +
      (url ? '<div style="margin-top:20px;"><a href="' + url + '" style="display:inline-block;background:#1c1b18;color:#fff;text-decoration:none;padding:10px 18px;border-radius:8px;font-weight:500;font-size:13px;">Open COLONY</a></div>' : '') +
      '</div></div>' +
      '<div style="text-align:center;color:#94a3b8;font-size:11px;margin-top:14px;">This is a one-time system notice from COLONY administrators.</div>' +
      '</div></body></html>';
  },

  _textFor: function(record) {
    var accent = this._accentFor(record.status);
    var lines = [
      'COLONY — ' + accent.label,
      '',
      record.message ? record.message : '(no additional details provided)',
      ''
    ];
    if (record.setByName) lines.push('Posted by: ' + record.setByName);
    if (record.setAt) lines.push('When: ' + record.setAt);
    try {
      var url = ScriptApp.getService().getUrl();
      if (url) lines.push('', 'Open COLONY: ' + url);
    } catch (e) {}
    return lines.join('\n');
  }
};

function getMaintenanceStatus() {
  try {
    return MaintenanceService.getStatus();
  } catch (error) {
    console.error('getMaintenanceStatus failed:', error);
    return { status: 'none', message: '', setAt: '', setBy: '', setByName: '', versionId: '' };
  }
}

function setMaintenanceStatus(payload) {
  try {
    return MaintenanceService.setStatus(payload || {});
  } catch (error) {
    console.error('setMaintenanceStatus failed:', error);
    return { success: false, error: error.message || 'Failed to set maintenance status' };
  }
}

function clearMaintenanceStatus(notifyEmail) {
  try {
    return MaintenanceService.clearStatus(notifyEmail === true);
  } catch (error) {
    console.error('clearMaintenanceStatus failed:', error);
    return { success: false, error: error.message || 'Failed to clear maintenance status' };
  }
}
