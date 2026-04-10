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

function loginWithEmail(email) {
  try {
    if (!email || !email.includes('@')) {
      throw new Error('Please enter a valid email address');
    }

    email = email.toLowerCase().trim();
    var sessionEmail = Session.getActiveUser().getEmail();
    if (sessionEmail && sessionEmail.toLowerCase() !== email) {
      throw new Error('Email mismatch. You can only log in as yourself.');
    }
    var user = getUserByEmailOptimized(email);

    if (!user) {
      throw new Error('User not found. Please contact your administrator for access.');
    }

    if (typeof AuthService !== 'undefined') {
      var userWithPw = AuthService.getUserWithPassword(email);
      if (userWithPw && userWithPw.passwordHash && userWithPw.passwordSalt) {
        throw new Error('Password required. Please use the password login.');
      }
    }

    setCurrentUserEmailOptimized(email);
    return buildLoginPayload_(user, email);
  } catch (error) {
    console.error('loginWithEmail failed:', error);
    throw error;
  }
}

function logout() {
  try {
    const userEmail = getCurrentUserEmailOptimized();

    if (typeof AuthService !== 'undefined' && userEmail) {
      AuthService.invalidateAllUserSessions(userEmail);
    }

    // Clear per-user session cache
    const userKey = Session.getTemporaryActiveUserKey();
    if (userKey) {
      const cache = CacheService.getScriptCache();
      cache.remove('ACTIVE_SESSION_EMAIL_' + userKey);
    }

    clearAllCaches();

    return { success: true };
  } catch (error) {
    console.error('logout failed:', error);
    throw error;
  }
}

function loginWithPassword(email, password, options) {
  try {
    if (typeof AuthService === 'undefined') {
      throw new Error('AuthService not available');
    }

    var result = AuthService.authenticate(email, password, options || {});

    if (!result.success) {
      return result;
    }

    setCurrentUserEmailOptimized(email);
    var payload = buildLoginPayload_(result.user, email);
    payload.success = true;
    payload.session = result.session;
    return payload;
  } catch (error) {
    console.error('loginWithPassword failed:', error);
    return {
      success: false,
      error: error.message || 'Authentication failed'
    };
  }
}

function requestPasswordReset(email) {
  try {
    if (typeof AuthService === 'undefined') throw new Error('AuthService not available');
    return AuthService.requestPasswordReset(email);
  } catch (error) {
    console.error('requestPasswordReset failed:', error);
    return { success: false, error: error.message };
  }
}

function resetPasswordWithCode(email, code, newPassword) {
  try {
    if (typeof AuthService === 'undefined') throw new Error('AuthService not available');
    return AuthService.resetPasswordWithCode(email, code, newPassword);
  } catch (error) {
    console.error('resetPasswordWithCode failed:', error);
    return { success: false, error: error.message };
  }
}

function setUserPassword(email, newPassword, currentPassword) {
  try {
    if (typeof AuthService === 'undefined') {
      throw new Error('AuthService not available');
    }

    if (!email) {
      email = getCurrentUserEmailOptimized();
    }

    const currentUserEmail = getCurrentUserEmailOptimized();

    if (currentUserEmail !== email) {
      const currentUser = getUserByEmailOptimized(currentUserEmail);
      if (!currentUser || currentUser.role !== 'admin') {
        return { success: false, error: 'Permission denied' };
      }
    }

    return AuthService.setPassword(email, newPassword, currentPassword);
  } catch (error) {
    console.error('setUserPassword failed:', error);
    return { success: false, error: error.message };
  }
}

function validateSessionToken(token, userId) {
  try {
    if (typeof AuthService === 'undefined') {
      const email = getCurrentUserEmailOptimized();
      return {
        valid: email === userId,
        reason: email ? 'Legacy session' : 'No session'
      };
    }

    return AuthService.validateSession(token, userId);
  } catch (error) {
    console.error('validateSessionToken failed:', error);
    return { valid: false, reason: error.message };
  }
}

function getLoginRateLimitStatus(email) {
  try {
    if (typeof AuthService === 'undefined') {
      return { allowed: true, remainingAttempts: 5 };
    }

    return AuthService.checkRateLimit(email);
  } catch (error) {
    console.error('getLoginRateLimitStatus failed:', error);
    return { allowed: true, remainingAttempts: 5 };
  }
}

function logoutAllDevices() {
  try {
    const userEmail = getCurrentUserEmailOptimized();

    if (!userEmail) {
      return { success: false, error: 'Not logged in' };
    }

    if (typeof AuthService !== 'undefined') {
      AuthService.invalidateAllUserSessions(userEmail);
    }

    // Clear per-user session cache
    const userKey = Session.getTemporaryActiveUserKey();
    if (userKey) {
      const cache = CacheService.getScriptCache();
      cache.remove('ACTIVE_SESSION_EMAIL_' + userKey);
    }

    clearAllCaches();

    return { success: true };
  } catch (error) {
    console.error('logoutAllDevices failed:', error);
    return { success: false, error: error.message };
  }
}

function requestMfaCode(email) {
  try {
    if (typeof AuthService === 'undefined') {
      return { success: false, error: 'MFA not available' };
    }

    const user = getUserByEmailOptimized(email);

    if (!user) {
      return { success: false, error: 'User not found' };
    }

    AuthService.sendMfaCode(user);
    return { success: true, message: 'MFA code sent to your email' };
  } catch (error) {
    console.error('requestMfaCode failed:', error);
    return { success: false, error: error.message };
  }
}

function setMfaEnabled(enabled) {
  try {
    const userEmail = getCurrentUserEmailOptimized();

    if (!userEmail) {
      return { success: false, error: 'Not logged in' };
    }

    updateUser(userEmail, { mfaEnabled: enabled });
    return { success: true };
  } catch (error) {
    console.error('setMfaEnabled failed:', error);
    return { success: false, error: error.message };
  }
}
