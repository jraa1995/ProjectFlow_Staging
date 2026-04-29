var RecipientGroupsService = (function() {

  function _normalizeEmail(value) {
    return String(value || '').trim().toLowerCase();
  }

  function _isValidEmail(email) {
    if (!email) return false;
    var s = String(email).trim();
    return s.indexOf('@') !== -1 && s.indexOf('.') !== -1 && s.length <= 254;
  }

  function _requireProjectAccess(projectId, action) {
    if (!projectId) throw new Error('projectId is required');
    var permission = action === 'write' ? 'project:update' : 'project:read';
    PermissionGuard.requirePermission(permission, { projectId: projectId });
  }

  function _readGroups(projectId) {
    var sheet = getEmailGroupsSheet();
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    var columns = CONFIG.EMAIL_GROUP_COLUMNS;
    var groups = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var g = rowToObject(data[i], columns);
      if (projectId && g.projectId !== projectId) continue;
      groups.push(g);
    }
    return groups;
  }

  function _readMembersForGroup(groupId) {
    if (!groupId) return [];
    var sheet = getEmailGroupMembersSheet();
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    var columns = CONFIG.EMAIL_GROUP_MEMBER_COLUMNS;
    var members = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var m = rowToObject(data[i], columns);
      if (m.groupId !== groupId) continue;
      members.push(m);
    }
    return members;
  }

  function _readAllMembersForProject(projectId) {
    var sheet = getEmailGroupMembersSheet();
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return {};
    var columns = CONFIG.EMAIL_GROUP_MEMBER_COLUMNS;
    var groups = _readGroups(projectId);
    var allowed = {};
    groups.forEach(function(g) { allowed[g.id] = true; });
    var byGroup = {};
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var m = rowToObject(data[i], columns);
      if (!allowed[m.groupId]) continue;
      if (!byGroup[m.groupId]) byGroup[m.groupId] = [];
      byGroup[m.groupId].push(m);
    }
    return byGroup;
  }

  function _writeMembers(groupId, members) {
    var sheet = getEmailGroupMembersSheet();
    var data = sheet.getDataRange().getValues();
    var columns = CONFIG.EMAIL_GROUP_MEMBER_COLUMNS;
    for (var i = data.length - 1; i >= 1; i--) {
      if (data[i][1] === groupId) {
        sheet.deleteRow(i + 1);
      }
    }
    SpreadsheetApp.flush();
    if (!members || !members.length) return [];
    var seen = {};
    var rows = [];
    var saved = [];
    var timestamp = now();
    members.forEach(function(raw) {
      var email = _normalizeEmail(raw && (raw.email || raw));
      if (!email || !_isValidEmail(email)) return;
      if (seen[email]) return;
      seen[email] = true;
      var member = {
        id: generateId('EGM'),
        groupId: groupId,
        email: email,
        displayName: sanitize((raw && raw.displayName) ? String(raw.displayName) : ''),
        memberType: sanitize((raw && raw.memberType) ? String(raw.memberType) : 'user'),
        addedAt: timestamp
      };
      rows.push(objectToRow(member, columns));
      saved.push(member);
    });
    if (rows.length) {
      var startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, rows.length, columns.length).setValues(rows);
      SpreadsheetApp.flush();
    }
    return saved;
  }

  function listProjectGroups(projectId) {
    _requireProjectAccess(projectId, 'read');
    var groups = _readGroups(projectId);
    var members = _readAllMembersForProject(projectId);
    return groups.map(function(g) {
      var ms = members[g.id] || [];
      return {
        id: g.id,
        projectId: g.projectId,
        name: g.name,
        description: g.description || '',
        createdBy: g.createdBy,
        createdAt: g.createdAt,
        updatedAt: g.updatedAt,
        members: ms.map(function(m) {
          return { email: m.email, displayName: m.displayName || '', memberType: m.memberType || 'user' };
        })
      };
    });
  }

  function createGroup(payload) {
    if (!payload || !payload.projectId) throw new Error('projectId is required');
    if (!payload.name || !String(payload.name).trim()) throw new Error('Group name is required');
    _requireProjectAccess(payload.projectId, 'write');
    var existing = _readGroups(payload.projectId);
    var trimmedName = sanitize(String(payload.name).trim());
    var dup = existing.find(function(g) { return (g.name || '').toLowerCase() === trimmedName.toLowerCase(); });
    if (dup) throw new Error('A group with this name already exists in this project.');
    var sheet = getEmailGroupsSheet();
    var columns = CONFIG.EMAIL_GROUP_COLUMNS;
    var currentUser = getCurrentUserEmail();
    var group = {
      id: generateId('EG'),
      projectId: payload.projectId,
      name: trimmedName,
      description: sanitize(payload.description || ''),
      createdBy: currentUser,
      createdAt: now(),
      updatedAt: now()
    };
    sheet.appendRow(objectToRow(group, columns));
    SpreadsheetApp.flush();
    var savedMembers = _writeMembers(group.id, payload.members || []);
    logActivity(currentUser, 'created', 'email_group', group.id, { projectId: payload.projectId, name: group.name, memberCount: savedMembers.length });
    return Object.assign({}, group, {
      members: savedMembers.map(function(m) {
        return { email: m.email, displayName: m.displayName || '', memberType: m.memberType || 'user' };
      })
    });
  }

  function updateGroup(payload) {
    if (!payload || !payload.id) throw new Error('Group id is required');
    var sheet = getEmailGroupsSheet();
    var data = sheet.getDataRange().getValues();
    var columns = CONFIG.EMAIL_GROUP_COLUMNS;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === payload.id) {
        var existing = rowToObject(data[i], columns);
        _requireProjectAccess(existing.projectId, 'write');
        if (payload.name && String(payload.name).trim()) {
          existing.name = sanitize(String(payload.name).trim());
        }
        if (Object.prototype.hasOwnProperty.call(payload, 'description')) {
          existing.description = sanitize(payload.description || '');
        }
        existing.updatedAt = now();
        sheet.getRange(i + 1, 1, 1, columns.length).setValues([objectToRow(existing, columns)]);
        SpreadsheetApp.flush();
        var savedMembers = Array.isArray(payload.members)
          ? _writeMembers(existing.id, payload.members)
          : _readMembersForGroup(existing.id);
        var currentUser = getCurrentUserEmail();
        logActivity(currentUser, 'updated', 'email_group', existing.id, { projectId: existing.projectId, name: existing.name, memberCount: savedMembers.length });
        return Object.assign({}, existing, {
          members: savedMembers.map(function(m) {
            return { email: m.email, displayName: m.displayName || '', memberType: m.memberType || 'user' };
          })
        });
      }
    }
    throw new Error('Email group not found: ' + payload.id);
  }

  function deleteGroup(groupId) {
    if (!groupId) throw new Error('Group id is required');
    var sheet = getEmailGroupsSheet();
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === groupId) {
        var projectId = data[i][1];
        _requireProjectAccess(projectId, 'write');
        sheet.deleteRow(i + 1);
        SpreadsheetApp.flush();
        var memberSheet = getEmailGroupMembersSheet();
        var memberData = memberSheet.getDataRange().getValues();
        for (var j = memberData.length - 1; j >= 1; j--) {
          if (memberData[j][1] === groupId) memberSheet.deleteRow(j + 1);
        }
        SpreadsheetApp.flush();
        logActivity(getCurrentUserEmail(), 'deleted', 'email_group', groupId, { projectId: projectId });
        return { id: groupId, projectId: projectId };
      }
    }
    throw new Error('Email group not found: ' + groupId);
  }

  return {
    listProjectGroups: listProjectGroups,
    createGroup: createGroup,
    updateGroup: updateGroup,
    deleteGroup: deleteGroup
  };
})();

function getRecipientGroups(projectId) {
  try {
    return { success: true, groups: RecipientGroupsService.listProjectGroups(projectId) };
  } catch (error) {
    console.error('getRecipientGroups failed:', error);
    return { success: false, error: error.message, groups: [] };
  }
}

function saveRecipientGroup(payload) {
  try {
    return { success: true, group: RecipientGroupsService.createGroup(payload) };
  } catch (error) {
    console.error('saveRecipientGroup failed:', error);
    return { success: false, error: error.message };
  }
}

function updateRecipientGroup(payload) {
  try {
    return { success: true, group: RecipientGroupsService.updateGroup(payload) };
  } catch (error) {
    console.error('updateRecipientGroup failed:', error);
    return { success: false, error: error.message };
  }
}

function deleteRecipientGroup(groupId) {
  try {
    var result = RecipientGroupsService.deleteGroup(groupId);
    return { success: true, id: result.id };
  } catch (error) {
    console.error('deleteRecipientGroup failed:', error);
    return { success: false, error: error.message };
  }
}
