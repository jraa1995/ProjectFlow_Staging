var UPLOAD_MAX_BYTES = 10 * 1024 * 1024;
var AVATAR_MAX_BASE64_LEN = 45000;
var AVATAR_ALLOWED_MIMES = {
  'image/png': 'png',
  'image/jpeg': 'jpg',
  'image/jpg': 'jpg',
  'image/webp': 'webp',
  'image/gif': 'gif'
};

var AVATAR_PALETTE = [
  '#ef4444', '#f97316', '#f59e0b', '#eab308',
  '#84cc16', '#10b981', '#14b8a6', '#06b6d4',
  '#0ea5e9', '#3b82f6', '#6366f1', '#8b5cf6',
  '#a855f7', '#d946ef', '#ec4899', '#f43f5e'
];

function uploadAvatar(mimeType, base64Data, fileName) {
  try {
    var email = getCurrentUserEmailOptimized();
    if (!email) throw new Error('Not signed in');
    if (!base64Data) throw new Error('No file data received');
    if (!AVATAR_ALLOWED_MIMES[mimeType]) {
      throw new Error('Unsupported image format. Use PNG, JPG, WEBP, or GIF.');
    }

    var normalized = String(base64Data).replace(/\s+/g, '');
    if (normalized.length > AVATAR_MAX_BASE64_LEN) {
      throw new Error('Image is too large after compression. Try a smaller image.');
    }

    try { Utilities.base64Decode(normalized); }
    catch (e) { throw new Error('Failed to decode image payload'); }

    var dataUrl = 'data:' + mimeType + ';base64,' + normalized;

    var previousAvatar = '';
    try {
      var existing = getUserByEmail(email);
      if (existing) previousAvatar = existing.avatar || '';
    } catch (e) {}

    updateUser(email, { avatar: dataUrl });

    try {
      var jd = {};
      try {
        var fresh = getUserByEmail(email);
        if (fresh && typeof fresh.jsonData === 'string' && fresh.jsonData) {
          jd = JSON.parse(fresh.jsonData);
        }
      } catch (e) {}
      if (jd && jd.avatarFileId) {
        try { DriveApp.getFileById(jd.avatarFileId).setTrashed(true); } catch (e) {}
      }
      _setUserAvatarMeta_(email, { avatarFileId: '', avatarUploadedAt: now() });
    } catch (e) { console.error('Avatar metadata update failed:', e); }

    if (previousAvatar && /^https?:\/\/drive\.google\.com\/(thumbnail|uc)\?/.test(previousAvatar)) {
      var match = previousAvatar.match(/id=([a-zA-Z0-9_-]+)/);
      if (match && match[1]) {
        try { DriveApp.getFileById(match[1]).setTrashed(true); } catch (e) {}
      }
    }

    UserCache.invalidate();
    try {
      if (typeof CacheService !== 'undefined') {
        CacheService.getScriptCache().removeAll(['ACTIVE_USERS_CACHE', 'ALL_USERS_MENTIONS_CACHE', 'BATCH_DATA_CACHE']);
      }
    } catch (e) {}

    return {
      success: true,
      avatar: dataUrl
    };
  } catch (error) {
    console.error('uploadAvatar failed:', error);
    return { success: false, error: error.message };
  }
}

function clearAvatar() {
  try {
    var email = getCurrentUserEmailOptimized();
    if (!email) throw new Error('Not signed in');
    var user = getUserByEmail(email);
    if (user) {
      var jd = {};
      try { jd = typeof user.jsonData === 'string' ? JSON.parse(user.jsonData) : (user.jsonData || {}); } catch (e) {}
      if (jd.avatarFileId) {
        try { DriveApp.getFileById(jd.avatarFileId).setTrashed(true); } catch (e) {}
      }
      if (user.avatar && /^https?:\/\/drive\.google\.com\/(thumbnail|uc)\?/.test(user.avatar)) {
        var match = user.avatar.match(/id=([a-zA-Z0-9_-]+)/);
        if (match && match[1]) {
          try { DriveApp.getFileById(match[1]).setTrashed(true); } catch (e) {}
        }
      }
    }
    updateUser(email, { avatar: '' });
    _setUserAvatarMeta_(email, { avatarFileId: '', avatarUploadedAt: '' });
    UserCache.invalidate();
    try {
      if (typeof CacheService !== 'undefined') {
        CacheService.getScriptCache().removeAll(['ACTIVE_USERS_CACHE', 'ALL_USERS_MENTIONS_CACHE', 'BATCH_DATA_CACHE']);
      }
    } catch (e) {}
    return { success: true, avatar: '' };
  } catch (error) {
    console.error('clearAvatar failed:', error);
    return { success: false, error: error.message };
  }
}

function getMyAvatar() {
  var email = getCurrentUserEmailOptimized();
  if (!email) return { email: '', avatar: '', generated: _generateAvatarSpec_('') };
  var user = getUserByEmail(email);
  var avatarUrl = user && user.avatar ? user.avatar : '';
  return {
    email: email,
    avatar: avatarUrl,
    generated: _generateAvatarSpec_(user && user.name ? user.name : email)
  };
}

function getAvatarSpec(seed) {
  return _generateAvatarSpec_(seed || '');
}

function _generateAvatarSpec_(seed) {
  var s = String(seed || '').trim().toLowerCase();
  var hash = 0;
  for (var i = 0; i < s.length; i++) {
    hash = ((hash << 5) - hash) + s.charCodeAt(i);
    hash |= 0;
  }
  var idx = Math.abs(hash) % AVATAR_PALETTE.length;
  var color = AVATAR_PALETTE[idx];
  var initials = _computeInitials_(seed);
  return { color: color, initials: initials, seed: s };
}

function _computeInitials_(seed) {
  if (!seed) return '?';
  var parts = String(seed).split(/[\s@._-]+/).filter(Boolean);
  if (!parts.length) return '?';
  var out = '';
  for (var i = 0; i < parts.length && out.length < 2; i++) {
    if (parts[i][0]) out += parts[i][0];
  }
  return out.toUpperCase();
}

function resolveAvatarsFolder_() {
  var props = PropertiesService.getScriptProperties();
  var folderId = props.getProperty('AVATARS_FOLDER_ID');
  if (folderId) {
    try { return DriveApp.getFolderById(folderId); }
    catch (e) { console.error('Avatars folder missing, recreating:', e); }
  }
  var folder = DriveApp.createFolder('COLONY - User Avatars');
  props.setProperty('AVATARS_FOLDER_ID', folder.getId());
  return folder;
}

function _setUserAvatarMeta_(email, meta) {
  var lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    var sheet = getUsersSheet();
    var data = sheet.getDataRange().getValues();
    var columns = CONFIG.USER_COLUMNS;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && String(data[i][0]).toLowerCase() === email.toLowerCase()) {
        var user = rowToObject(data[i], columns);
        var jd = {};
        try { jd = typeof user.jsonData === 'string' && user.jsonData ? JSON.parse(user.jsonData) : (user.jsonData || {}); } catch (e) {}
        Object.keys(meta).forEach(function(k) { jd[k] = meta[k]; });
        user.jsonData = JSON.stringify(jd);
        var newRow = objectToRow(user, columns);
        sheet.getRange(i + 1, 1, 1, columns.length).setValues([newRow]);
        return;
      }
    }
  } finally {
    lock.releaseLock();
  }
}

function uploadFileToProject(projectId, fileName, mimeType, base64Data) {
  try {
    PermissionGuard.requirePermission('project:update', { projectId: projectId });
    if (!projectId) throw new Error('projectId is required');
    if (!fileName) throw new Error('fileName is required');
    if (!base64Data) throw new Error('No file data received');

    var project = getProjectById(projectId);
    if (!project) throw new Error('Project not found: ' + projectId);

    var bytes;
    try {
      bytes = Utilities.base64Decode(base64Data);
    } catch (e) {
      throw new Error('Failed to decode file payload');
    }
    if (bytes.length > UPLOAD_MAX_BYTES) {
      throw new Error('File exceeds the ' + Math.round(UPLOAD_MAX_BYTES / 1048576) + ' MB upload limit.');
    }

    var folder = resolveProjectFolder_(project);
    var blob = Utilities.newBlob(bytes, mimeType || 'application/octet-stream', fileName);
    var file = folder.createFile(blob);

    return {
      success: true,
      file: {
        id: file.getId(),
        name: file.getName(),
        url: file.getUrl(),
        mimeType: file.getMimeType()
      }
    };
  } catch (error) {
    console.error('uploadFileToProject failed:', error);
    return { success: false, error: error.message };
  }
}

function resolveProjectFolder_(project) {
  var settings = {};
  try { settings = typeof project.settings === 'string' ? JSON.parse(project.settings) : (project.settings || {}); } catch (e) {}

  if (settings.googleDriveFolderId) {
    try {
      return DriveApp.getFolderById(settings.googleDriveFolderId);
    } catch (e) {
      console.error('resolveProjectFolder_: existing folderId invalid, creating new:', e);
    }
  }

  var folder = DriveApp.createFolder('COLONY - ' + (project.name || project.id));
  settings.googleDriveFolderId = folder.getId();
  settings.googleDriveFolder = folder.getUrl();
  try {
    updateProject(project.id, { settings: JSON.stringify(settings) });
  } catch (e) {
    console.error('resolveProjectFolder_: failed to persist folder ID on project:', e);
  }
  return folder;
}
