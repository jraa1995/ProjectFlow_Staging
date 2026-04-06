function getCurrentUserEmailOptimized() {
  try {
    // Primary: use GAS built-in user identification (per-user, no crossover)
    const activeUser = Session.getActiveUser();
    if (activeUser) {
      const email = activeUser.getEmail();
      if (email) {
        return email.toLowerCase().trim();
      }
    }

    // Fallback: check user-specific session in cache (keyed by session token)
    const cache = CacheService.getScriptCache();
    const sessionEmail = cache.get('ACTIVE_SESSION_EMAIL_' + Session.getTemporaryActiveUserKey());
    if (sessionEmail) {
      return sessionEmail;
    }

    return null;
  } catch (error) {
    console.error('getCurrentUserEmailOptimized failed:', error);
    return null;
  }
}

function setCurrentUserEmailOptimized(email) {
  try {
    // Store session keyed by temporary user key (per-user, no crossover)
    const userKey = Session.getTemporaryActiveUserKey();
    if (userKey) {
      const cache = CacheService.getScriptCache();
      const sessionDuration = 24 * 60 * 60; // 24 hours in seconds
      cache.put('ACTIVE_SESSION_EMAIL_' + userKey, email.toLowerCase().trim(), sessionDuration);
    }
  } catch (error) {
    console.error('setCurrentUserEmailOptimized failed:', error);
  }
}

function sanitizeUserForClient(user) {
  if (!user) return null;
  return {
    email: user.email,
    name: user.name,
    role: user.role,
    active: user.active,
    avatar: user.avatar,
    organizationId: user.organizationId,
    teamId: user.teamId,
    mfaEnabled: user.mfaEnabled || false
  };
}

function filterTasksByUserRole(tasks, userEmail, userRole) {
  if (!userRole || userRole === 'admin' || userRole === 'manager') {
    return tasks;
  }
  // Members and viewers only see tasks assigned to them or created by them
  return tasks.filter(function(task) {
    return (task.assignee && task.assignee.toLowerCase() === userEmail.toLowerCase()) ||
           (task.reporter && task.reporter.toLowerCase() === userEmail.toLowerCase());
  });
}

function filterProjectsByUserRole(projects, tasks, userEmail, userRole) {
  if (!userRole || userRole === 'admin' || userRole === 'manager') {
    return projects;
  }
  var userProjectIds = {};
  tasks.forEach(function(task) {
    if (task.projectId) {
      userProjectIds[task.projectId] = true;
    }
  });
  return projects.filter(function(project) {
    return userProjectIds[project.id];
  });
}

function getCurrentUserRole() {
  try {
    var email = getCurrentUserEmailOptimized();
    if (!email) return null;
    var user = getUserByEmailOptimized(email);
    return user ? (user.role || 'member') : null;
  } catch (e) {
    console.error('getCurrentUserRole failed:', e);
    return null;
  }
}

function getUserByEmailOptimized(email) {
  try {
    const sheet = getUsersSheet();
    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) return null;

    const headers = data[0];
    const emailIndex = headers.indexOf('email');

    if (emailIndex === -1) return null;

    for (let i = 1; i < data.length; i++) {
      if (data[i][emailIndex] && data[i][emailIndex].toLowerCase() === email.toLowerCase()) {
        const user = {};
        headers.forEach((header, index) => {
          user[header] = data[i][index];
        });
        return user;
      }
    }

    return null;
  } catch (error) {
    console.error('getUserByEmailOptimized failed:', error);
    return null;
  }
}

function getUserOrganizations() {
  try {
    const userEmail = getCurrentUserEmailOptimized();

    if (!userEmail) {
      return { success: false, error: 'Not logged in', organizations: [] };
    }

    const user = getUserByEmailOptimized(userEmail);
    const orgs = getAllOrganizations();
    const orgMembers = [];

    orgs.forEach(org => {
      const members = getOrganizationMembers(org.id);
      if (members.some(m => m.userId === userEmail)) {
        orgMembers.push({
          ...org,
          memberRole: members.find(m => m.userId === userEmail)?.orgRole
        });
      }
    });

    return {
      success: true,
      organizations: orgMembers,
      currentOrgId: user?.organizationId
    };
  } catch (error) {
    console.error('getUserOrganizations failed:', error);
    return { success: false, error: error.message, organizations: [] };
  }
}

function getUserTeams() {
  try {
    const userEmail = getCurrentUserEmailOptimized();

    if (!userEmail) {
      return { success: false, error: 'Not logged in', teams: [] };
    }

    const user = getUserByEmailOptimized(userEmail);
    const teams = getTeamsForUser(userEmail);

    return {
      success: true,
      teams: teams.map(team => ({
        ...team,
        memberCount: getTeamMembers(team.id).length
      })),
      currentTeamId: user?.teamId
    };
  } catch (error) {
    console.error('getUserTeams failed:', error);
    return { success: false, error: error.message, teams: [] };
  }
}

function createNewOrganization(orgData) {
  try {
    const org = createOrganization(orgData);
    return { success: true, organization: org };
  } catch (error) {
    console.error('createNewOrganization failed:', error);
    return { success: false, error: error.message };
  }
}

function createNewTeam(teamData) {
  try {
    const team = createTeam(teamData);
    return { success: true, team: team };
  } catch (error) {
    console.error('createNewTeam failed:', error);
    return { success: false, error: error.message };
  }
}

function getOrganizationDetails(orgId) {
  try {
    const org = getOrganizationById(orgId);

    if (!org) {
      return { success: false, error: 'Organization not found' };
    }

    const members = getOrganizationMembers(orgId);
    const teams = getAllTeams(orgId);

    const enrichedMembers = members.map(m => {
      const user = getUserByEmail(m.userId);
      return {
        ...m,
        name: user?.name || m.userId,
        email: m.userId,
        role: user?.role,
        active: user?.active
      };
    });

    return {
      success: true,
      organization: org,
      members: enrichedMembers,
      teams: teams
    };
  } catch (error) {
    console.error('getOrganizationDetails failed:', error);
    return { success: false, error: error.message };
  }
}

function getTeamDetails(teamId) {
  try {
    const team = getTeamById(teamId);

    if (!team) {
      return { success: false, error: 'Team not found' };
    }

    const members = getTeamMembers(teamId);

    const enrichedMembers = members.map(m => {
      const user = getUserByEmail(m.userId);
      return {
        ...m,
        name: user?.name || m.userId,
        email: m.userId,
        role: user?.role,
        active: user?.active
      };
    });

    return {
      success: true,
      team: team,
      members: enrichedMembers
    };
  } catch (error) {
    console.error('getTeamDetails failed:', error);
    return { success: false, error: error.message };
  }
}

function addOrgMember(orgId, userEmail, role) {
  try {
    const member = addOrganizationMember(orgId, userEmail, role);
    return { success: true, member: member };
  } catch (error) {
    console.error('addOrgMember failed:', error);
    return { success: false, error: error.message };
  }
}

function addTeamMemberToTeam(teamId, userEmail, role) {
  try {
    const member = addTeamMember(teamId, userEmail, role);
    return { success: true, member: member };
  } catch (error) {
    console.error('addTeamMemberToTeam failed:', error);
    return { success: false, error: error.message };
  }
}
