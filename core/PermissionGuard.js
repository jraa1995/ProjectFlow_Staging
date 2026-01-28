const PermissionGuard = {
  PERMISSIONS: {
    'task:create': 'Create new tasks',
    'task:read:own': 'Read own tasks',
    'task:read:all': 'Read all tasks',
    'task:update:own': 'Update own tasks',
    'task:update:all': 'Update all tasks',
    'task:delete:own': 'Delete own tasks',
    'task:delete:all': 'Delete all tasks',
    'task:assign': 'Assign tasks to others',
    'project:create': 'Create new projects',
    'project:read': 'View projects',
    'project:update': 'Update projects',
    'project:delete': 'Delete projects',
    'project:manage_members': 'Manage project members',
    'user:create': 'Create new users',
    'user:read': 'View users',
    'user:update': 'Update users',
    'user:delete': 'Delete users',
    'user:manage_roles': 'Manage user roles',
    'analytics:view:own': 'View own analytics',
    'analytics:view:team': 'View team analytics',
    'analytics:view:all': 'View all analytics',
    'admin:settings': 'Manage system settings',
    'admin:audit_log': 'View audit logs',
    'admin:automation': 'Manage automation rules',
    'admin:webhooks': 'Manage webhooks'
  },

  DEFAULT_ROLES: {
    admin: {
      name: 'admin',
      description: 'Full system access',
      permissions: Object.keys(this.PERMISSIONS || {}),
      isSystemRole: true
    },
    manager: {
      name: 'manager',
      description: 'Task and project management with team analytics',
      permissions: [
        'task:create', 'task:read:all', 'task:update:all', 'task:delete:own', 'task:assign',
        'project:create', 'project:read', 'project:update', 'project:manage_members',
        'user:read',
        'analytics:view:own', 'analytics:view:team'
      ],
      isSystemRole: true
    },
    member: {
      name: 'member',
      description: 'Standard team member',
      permissions: [
        'task:create', 'task:read:own', 'task:read:all', 'task:update:own', 'task:delete:own',
        'project:read',
        'user:read',
        'analytics:view:own'
      ],
      isSystemRole: true
    },
    viewer: {
      name: 'viewer',
      description: 'Read-only access',
      permissions: [
        'task:read:all',
        'project:read',
        'user:read',
        'analytics:view:own'
      ],
      isSystemRole: true
    }
  },

  _roleCache: null,
  _roleCacheExpiry: null,
  CACHE_TTL_MS: 5 * 60 * 1000,

  can(permission, context = {}) {
    try {
      const userEmail = context.userEmail || getCurrentUserEmail();
      if (!userEmail) return false;

      const user = getUserByEmail(userEmail);
      if (!user) return false;

      if (user.role === 'admin') return true;

      const rolePermissions = this.getRolePermissions(user.role);

      if (rolePermissions.includes(permission)) return true;

      const [resource, action, scope] = permission.split(':');

      if (rolePermissions.includes(`${resource}:*`)) return true;

      if (scope === 'own') {
        if (rolePermissions.includes(`${resource}:${action}:all`)) return true;
      }

      if (context.projectId) {
        const projectPermissions = this.getProjectPermissions(userEmail, context.projectId);
        if (projectPermissions.includes(permission)) return true;
      }

      if (scope === 'own' && context.ownerId) {
        if (context.ownerId === userEmail) {
          if (rolePermissions.includes(`${resource}:${action}:own`)) return true;
        }
      }

      return false;
    } catch (error) {
      console.error('Permission check failed:', error);
      return false;
    }
  },

  requirePermission(permission, context = {}) {
    if (!this.can(permission, context)) {
      const errorMessage = `Permission denied: ${permission}`;
      console.error(errorMessage, { permission, context });
      throw new Error(errorMessage);
    }
  },

  canAll(permissions, context = {}) {
    return permissions.every(p => this.can(p, context));
  },

  canAny(permissions, context = {}) {
    return permissions.some(p => this.can(p, context));
  },

  getRolePermissions(roleName) {
    if (this._roleCache && this._roleCacheExpiry > Date.now()) {
      const cachedRole = this._roleCache[roleName];
      if (cachedRole) return cachedRole.permissions || [];
    }

    this.loadRoles();
    const role = this._roleCache[roleName];
    return role ? (role.permissions || []) : [];
  },

  loadRoles() {
    try {
      const sheet = getRolesSheet();
      const data = sheet.getDataRange().getValues();
      const columns = CONFIG.ROLE_COLUMNS;

      this._roleCache = {};

      Object.keys(this.DEFAULT_ROLES).forEach(roleName => {
        this._roleCache[roleName] = {
          ...this.DEFAULT_ROLES[roleName],
          permissions: this.DEFAULT_ROLES[roleName].permissions || this.getDefaultPermissionsForRole(roleName)
        };
      });

      for (let i = 1; i < data.length; i++) {
        const role = rowToObject(data[i], columns);
        if (role.name) {
          let permissions = role.permissions;
          if (typeof permissions === 'string') {
            try {
              permissions = JSON.parse(permissions);
            } catch (e) {
              permissions = permissions.split(',').map(p => p.trim());
            }
          }
          this._roleCache[role.name] = {
            ...role,
            permissions: permissions || []
          };
        }
      }

      this._roleCacheExpiry = Date.now() + this.CACHE_TTL_MS;
    } catch (error) {
      console.error('Failed to load roles:', error);
      this._roleCache = {};
      Object.keys(this.DEFAULT_ROLES).forEach(roleName => {
        this._roleCache[roleName] = {
          ...this.DEFAULT_ROLES[roleName],
          permissions: this.getDefaultPermissionsForRole(roleName)
        };
      });
      this._roleCacheExpiry = Date.now() + this.CACHE_TTL_MS;
    }
  },

  getDefaultPermissionsForRole(roleName) {
    const allPermissions = Object.keys(this.PERMISSIONS);

    switch (roleName) {
      case 'admin':
        return allPermissions;
      case 'manager':
        return [
          'task:create', 'task:read:all', 'task:update:all', 'task:delete:own', 'task:assign',
          'project:create', 'project:read', 'project:update', 'project:manage_members',
          'user:read',
          'analytics:view:own', 'analytics:view:team'
        ];
      case 'member':
        return [
          'task:create', 'task:read:own', 'task:read:all', 'task:update:own', 'task:delete:own',
          'project:read',
          'user:read',
          'analytics:view:own'
        ];
      case 'viewer':
        return [
          'task:read:all',
          'project:read',
          'user:read',
          'analytics:view:own'
        ];
      default:
        return ['task:read:own', 'project:read'];
    }
  },

  createOrUpdateRole(roleData) {
    this.requirePermission('user:manage_roles');

    const sheet = getRolesSheet();
    const data = sheet.getDataRange().getValues();
    const columns = CONFIG.ROLE_COLUMNS;

    const permissions = Array.isArray(roleData.permissions)
      ? roleData.permissions
      : (roleData.permissions || '').split(',').map(p => p.trim());

    const role = {
      id: roleData.id || generateId('role'),
      name: roleData.name,
      description: roleData.description || '',
      permissions: JSON.stringify(permissions),
      isSystemRole: roleData.isSystemRole || false,
      createdAt: roleData.createdAt || now()
    };

    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][columns.indexOf('name')] === role.name) {
        const newRow = objectToRow(role, columns);
        sheet.getRange(i + 1, 1, 1, columns.length).setValues([newRow]);
        found = true;
        break;
      }
    }

    if (!found) {
      sheet.appendRow(objectToRow(role, columns));
    }

    this._roleCache = null;
    return role;
  },

  initializeDefaultRoles() {
    const sheet = getRolesSheet();
    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      Object.keys(this.DEFAULT_ROLES).forEach(roleName => {
        const role = this.DEFAULT_ROLES[roleName];
        const roleData = {
          id: generateId('role'),
          name: roleName,
          description: role.description,
          permissions: JSON.stringify(this.getDefaultPermissionsForRole(roleName)),
          isSystemRole: true,
          createdAt: now()
        };
        sheet.appendRow(objectToRow(roleData, CONFIG.ROLE_COLUMNS));
      });
    }
  },

  getProjectPermissions(userEmail, projectId) {
    try {
      const sheet = getProjectMembersSheet();
      const data = sheet.getDataRange().getValues();
      const columns = CONFIG.PROJECT_MEMBER_COLUMNS;

      for (let i = 1; i < data.length; i++) {
        const member = rowToObject(data[i], columns);
        if (member.projectId === projectId && member.userId === userEmail) {
          let permissions = member.permissions;
          if (typeof permissions === 'string' && permissions) {
            try {
              permissions = JSON.parse(permissions);
            } catch (e) {
              permissions = [];
            }
          }
          return permissions || [];
        }
      }

      return [];
    } catch (error) {
      console.error('Failed to get project permissions:', error);
      return [];
    }
  },

  addProjectMember(projectId, userEmail, projectRole, permissions) {
    this.requirePermission('project:manage_members', { projectId });

    const sheet = getProjectMembersSheet();
    const columns = CONFIG.PROJECT_MEMBER_COLUMNS;
    const currentUser = getCurrentUserEmail();

    const member = {
      id: generateId('pm'),
      projectId: projectId,
      userId: userEmail,
      projectRole: projectRole || 'member',
      permissions: permissions ? JSON.stringify(permissions) : '',
      addedAt: now(),
      addedBy: currentUser
    };

    sheet.appendRow(objectToRow(member, columns));
    return member;
  },

  removeProjectMember(projectId, userEmail) {
    this.requirePermission('project:manage_members', { projectId });

    const sheet = getProjectMembersSheet();
    const data = sheet.getDataRange().getValues();
    const columns = CONFIG.PROJECT_MEMBER_COLUMNS;

    for (let i = 1; i < data.length; i++) {
      const member = rowToObject(data[i], columns);
      if (member.projectId === projectId && member.userId === userEmail) {
        sheet.deleteRow(i + 1);
        return true;
      }
    }

    return false;
  },

  getProjectMembers(projectId) {
    const sheet = getProjectMembersSheet();
    const data = sheet.getDataRange().getValues();
    const columns = CONFIG.PROJECT_MEMBER_COLUMNS;
    const members = [];

    for (let i = 1; i < data.length; i++) {
      const member = rowToObject(data[i], columns);
      if (member.projectId === projectId) {
        members.push(member);
      }
    }

    return members;
  },

  canReadTask(task, context = {}) {
    const userEmail = context.userEmail || getCurrentUserEmail();

    if (this.can('task:read:all', context)) return true;

    if (task.assignee === userEmail || task.reporter === userEmail) {
      return this.can('task:read:own', { ...context, ownerId: userEmail });
    }

    if (task.projectId) {
      const projectMembers = this.getProjectMembers(task.projectId);
      if (projectMembers.some(m => m.userId === userEmail)) {
        return true;
      }
    }

    return false;
  },

  canUpdateTask(task, context = {}) {
    const userEmail = context.userEmail || getCurrentUserEmail();

    if (this.can('task:update:all', context)) return true;

    if (task.assignee === userEmail || task.reporter === userEmail) {
      return this.can('task:update:own', { ...context, ownerId: userEmail });
    }

    return false;
  },

  canDeleteTask(task, context = {}) {
    const userEmail = context.userEmail || getCurrentUserEmail();

    if (this.can('task:delete:all', context)) return true;

    if (task.reporter === userEmail) {
      return this.can('task:delete:own', { ...context, ownerId: userEmail });
    }

    return false;
  },

  filterTasksByPermission(tasks, context = {}) {
    const userEmail = context.userEmail || getCurrentUserEmail();

    if (this.can('task:read:all', context)) return tasks;

    return tasks.filter(task => this.canReadTask(task, context));
  }
};

if (typeof module !== 'undefined' && module.exports) {
  module.exports = PermissionGuard;
}
