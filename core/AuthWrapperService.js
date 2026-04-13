function buildLoginPayload_(user, email) {
  try {
    var fastData = getInitialDataFast();
    return {
      user: sanitizeUserForClient(user),
      board: fastData.board,
      config: fastData.config
    };
  } catch (e) {
    console.error('buildLoginPayload_ fast path failed:', e);
    return {
      user: sanitizeUserForClient(user),
      board: {
        columns: [],
        projects: [],
        users: [],
        stats: { total: 0, completed: 0, progress: 0, dueSoon: 0, overdue: 0 },
        taskCount: 0
      },
      config: {
        statuses: CONFIG.STATUSES,
        priorities: CONFIG.PRIORITIES,
        types: CONFIG.TYPES,
        colors: CONFIG.COLORS
      }
    };
  }
}

function loginWithEmail() {
  return { success: false, error: 'Authentication is handled by Google. Please reload the page.' };
}

function loginWithPassword() {
  return { success: false, error: 'Authentication is handled by Google. Please reload the page.' };
}

function logout() {
  try {
    clearAllCaches();
    return { success: true };
  } catch (error) {
    console.error('logout failed:', error);
    return { success: true };
  }
}

function requestPasswordReset() {
  return { success: false, error: 'Password reset is no longer available. Authentication is handled by Google.' };
}

function resetPasswordWithCode() {
  return { success: false, error: 'Password reset is no longer available. Authentication is handled by Google.' };
}

function setUserPassword() {
  return { success: false, error: 'Password management is no longer available. Authentication is handled by Google.' };
}

function validateSessionToken() {
  return { valid: true, reason: 'Google native auth' };
}

function getLoginRateLimitStatus() {
  return { allowed: true, remainingAttempts: 999 };
}

function logoutAllDevices() {
  try {
    clearAllCaches();
    return { success: true };
  } catch (error) {
    console.error('logoutAllDevices failed:', error);
    return { success: true };
  }
}

function requestMfaCode() {
  return { success: false, error: 'MFA is no longer used. Authentication is handled by Google.' };
}

function setMfaEnabled() {
  return { success: false, error: 'MFA is no longer used. Authentication is handled by Google.' };
}
