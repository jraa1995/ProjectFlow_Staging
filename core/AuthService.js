const AuthService = {
  SESSION_DURATION_MS: 24 * 60 * 60 * 1000,
  SESSION_REFRESH_THRESHOLD_MS: 60 * 60 * 1000,
  RATE_LIMIT_ATTEMPTS: 5,
  RATE_LIMIT_WINDOW_MS: 15 * 60 * 1000,
  LOCKOUT_DURATION_MS: 30 * 60 * 1000,
  MFA_CODE_LENGTH: 6,
  MFA_CODE_EXPIRY_MS: 5 * 60 * 1000,

  createSession(userId, ipFingerprint) {
    const token = this.generateSessionToken();
    const now = new Date();
    const expiresAt = new Date(now.getTime() + this.SESSION_DURATION_MS);

    const session = {
      id: generateId('sess'),
      userId: userId,
      token: token,
      createdAt: now.toISOString(),
      expiresAt: expiresAt.toISOString(),
      lastActivityAt: now.toISOString(),
      ipFingerprint: ipFingerprint || '',
      isValid: true
    };

    this.saveSession(session);
    const userCache = CacheService.getUserCache();
    userCache.put('SESSION_' + userId, JSON.stringify(session), this.SESSION_DURATION_MS / 1000);

    return session;
  },

  generateSessionToken() {
    const uuid1 = Utilities.getUuid();
    const uuid2 = Utilities.getUuid();
    const timestamp = Date.now().toString(36);
    const combined = uuid1 + uuid2 + timestamp;
    const hashBytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, combined);
    return Utilities.base64Encode(hashBytes);
  },

  validateSession(token, userId) {
    if (!token || !userId) {
      return { valid: false, session: null, reason: 'Missing token or userId' };
    }

    const userCache = CacheService.getUserCache();
    const cachedSession = userCache.get('SESSION_' + userId);

    if (cachedSession) {
      try {
        const session = JSON.parse(cachedSession);
        if (session.token === token && session.isValid) {
          const expiresAt = new Date(session.expiresAt);
          const now = new Date();
          if (expiresAt > now) {
            this.updateSessionActivity(session.id, userId);
            return { valid: true, session: session, reason: 'Valid' };
          }
        }
      } catch (e) {
        console.error('Cache parse error:', e);
      }
    }

    const session = this.getSessionByToken(token, userId);

    if (!session) {
      return { valid: false, session: null, reason: 'Session not found' };
    }

    if (!session.isValid) {
      return { valid: false, session: null, reason: 'Session invalidated' };
    }

    const expiresAt = new Date(session.expiresAt);
    const now = new Date();

    if (expiresAt <= now) {
      this.invalidateSession(session.id);
      return { valid: false, session: null, reason: 'Session expired' };
    }

    this.updateSessionActivity(session.id, userId);
    userCache.put('SESSION_' + userId, JSON.stringify(session), this.SESSION_DURATION_MS / 1000);

    return { valid: true, session: session, reason: 'Valid' };
  },

  updateSessionActivity(sessionId, userId) {
    try {
      const sheet = this.getSessionTokensSheet();
      const data = sheet.getDataRange().getValues();
      const columns = CONFIG.SESSION_TOKEN_COLUMNS;
      const idIndex = 0;
      const lastActivityIndex = columns.indexOf('lastActivityAt');

      for (let i = 1; i < data.length; i++) {
        if (data[i][idIndex] === sessionId) {
          const now = new Date().toISOString();
          sheet.getRange(i + 1, lastActivityIndex + 1).setValue(now);

          const expiresAt = new Date(data[i][columns.indexOf('expiresAt')]);
          const remaining = expiresAt.getTime() - Date.now();

          if (remaining < this.SESSION_REFRESH_THRESHOLD_MS) {
            const newExpiry = new Date(Date.now() + this.SESSION_DURATION_MS);
            sheet.getRange(i + 1, columns.indexOf('expiresAt') + 1).setValue(newExpiry.toISOString());
          }
          break;
        }
      }
    } catch (error) {
      console.error('Failed to update session activity:', error);
    }
  },

  invalidateSession(sessionId) {
    try {
      const sheet = this.getSessionTokensSheet();
      const data = sheet.getDataRange().getValues();
      const columns = CONFIG.SESSION_TOKEN_COLUMNS;
      const idIndex = 0;
      const isValidIndex = columns.indexOf('isValid');
      const userIdIndex = columns.indexOf('userId');

      for (let i = 1; i < data.length; i++) {
        if (data[i][idIndex] === sessionId) {
          sheet.getRange(i + 1, isValidIndex + 1).setValue(false);
          const userId = data[i][userIdIndex];
          if (userId) {
            const userCache = CacheService.getUserCache();
            userCache.remove('SESSION_' + userId);
          }
          break;
        }
      }
    } catch (error) {
      console.error('Failed to invalidate session:', error);
    }
  },

  invalidateAllUserSessions(userId) {
    try {
      const sheet = this.getSessionTokensSheet();
      const data = sheet.getDataRange().getValues();
      const columns = CONFIG.SESSION_TOKEN_COLUMNS;
      const userIdIndex = columns.indexOf('userId');
      const isValidIndex = columns.indexOf('isValid');

      for (let i = 1; i < data.length; i++) {
        if (data[i][userIdIndex] === userId && data[i][isValidIndex] !== false) {
          sheet.getRange(i + 1, isValidIndex + 1).setValue(false);
        }
      }

      const userCache = CacheService.getUserCache();
      userCache.remove('SESSION_' + userId);
    } catch (error) {
      console.error('Failed to invalidate all user sessions:', error);
    }
  },

  checkRateLimit(email) {
    const cache = CacheService.getScriptCache();
    const key = 'LOGIN_ATTEMPTS_' + email.toLowerCase();
    const lockoutKey = 'LOGIN_LOCKOUT_' + email.toLowerCase();

    const lockoutUntil = cache.get(lockoutKey);
    if (lockoutUntil) {
      const lockoutTime = new Date(lockoutUntil);
      if (lockoutTime > new Date()) {
        return {
          allowed: false,
          remainingAttempts: 0,
          lockoutUntil: lockoutTime,
          message: 'Account locked. Try again at ' + lockoutTime.toLocaleTimeString()
        };
      } else {
        cache.remove(lockoutKey);
        cache.remove(key);
      }
    }

    const attemptsData = cache.get(key);
    let attempts = 0;

    if (attemptsData) {
      try {
        const data = JSON.parse(attemptsData);
        attempts = data.count || 0;
      } catch (e) {
        attempts = 0;
      }
    }

    const remainingAttempts = this.RATE_LIMIT_ATTEMPTS - attempts;

    return {
      allowed: remainingAttempts > 0,
      remainingAttempts: Math.max(0, remainingAttempts),
      lockoutUntil: null,
      message: remainingAttempts > 0 ? null : 'Too many login attempts'
    };
  },

  recordFailedAttempt(email) {
    const cache = CacheService.getScriptCache();
    const key = 'LOGIN_ATTEMPTS_' + email.toLowerCase();
    const lockoutKey = 'LOGIN_LOCKOUT_' + email.toLowerCase();

    const attemptsData = cache.get(key);
    let attempts = 0;
    let firstAttemptTime = Date.now();

    if (attemptsData) {
      try {
        const data = JSON.parse(attemptsData);
        attempts = data.count || 0;
        firstAttemptTime = data.firstAttempt || Date.now();
      } catch (e) {
        attempts = 0;
      }
    }

    attempts++;

    cache.put(key, JSON.stringify({
      count: attempts,
      firstAttempt: firstAttemptTime
    }), this.RATE_LIMIT_WINDOW_MS / 1000);

    if (attempts >= this.RATE_LIMIT_ATTEMPTS) {
      const lockoutUntil = new Date(Date.now() + this.LOCKOUT_DURATION_MS);
      cache.put(lockoutKey, lockoutUntil.toISOString(), this.LOCKOUT_DURATION_MS / 1000);
      this.updateUserLockout(email, lockoutUntil, attempts);

      return {
        allowed: false,
        remainingAttempts: 0,
        lockoutUntil: lockoutUntil,
        message: 'Account locked for 30 minutes due to too many failed attempts'
      };
    }

    return {
      allowed: true,
      remainingAttempts: this.RATE_LIMIT_ATTEMPTS - attempts,
      lockoutUntil: null,
      message: `${this.RATE_LIMIT_ATTEMPTS - attempts} attempts remaining`
    };
  },

  clearRateLimit(email) {
    const cache = CacheService.getScriptCache();
    const key = 'LOGIN_ATTEMPTS_' + email.toLowerCase();
    const lockoutKey = 'LOGIN_LOCKOUT_' + email.toLowerCase();
    cache.remove(key);
    cache.remove(lockoutKey);
    this.clearUserLockout(email);
  },

  authenticate(email, password, options = {}) {
    email = (email || '').toLowerCase().trim();

    if (!email || !email.includes('@')) {
      return { success: false, error: 'Invalid email format' };
    }

    if (!password) {
      return { success: false, error: 'Password is required' };
    }

    const rateLimit = this.checkRateLimit(email);
    if (!rateLimit.allowed) {
      return {
        success: false,
        error: rateLimit.message,
        lockoutUntil: rateLimit.lockoutUntil
      };
    }

    const user = this.getUserWithPassword(email);

    if (!user) {
      this.recordFailedAttempt(email);
      return {
        success: false,
        error: 'Invalid email or password',
        remainingAttempts: rateLimit.remainingAttempts - 1
      };
    }

    if (user.lockoutUntil) {
      const lockoutTime = new Date(user.lockoutUntil);
      if (lockoutTime > new Date()) {
        return {
          success: false,
          error: 'Account is locked. Try again later.',
          lockoutUntil: lockoutTime
        };
      }
    }

    if (!user.passwordHash || !user.passwordSalt) {
      return {
        success: false,
        error: 'Password not set. Please contact administrator.',
        passwordNotSet: true
      };
    }

    const passwordValid = PasswordService.verifyPassword(
      password,
      user.passwordHash,
      user.passwordSalt
    );

    if (!passwordValid) {
      const failResult = this.recordFailedAttempt(email);
      return {
        success: false,
        error: 'Invalid email or password',
        remainingAttempts: failResult.remainingAttempts,
        lockoutUntil: failResult.lockoutUntil
      };
    }

    if (user.mfaEnabled) {
      if (!options.mfaCode) {
        this.sendMfaCode(user);
        return {
          success: false,
          mfaRequired: true,
          error: 'MFA code required. Check your email.'
        };
      }

      if (!this.verifyMfaCode(email, options.mfaCode)) {
        return {
          success: false,
          error: 'Invalid or expired MFA code',
          mfaRequired: true
        };
      }
    }

    this.clearRateLimit(email);
    const session = this.createSession(email, options.ipFingerprint);
    this.updateLastLogin(email);

    const safeUser = {
      email: user.email,
      name: user.name,
      role: user.role,
      active: user.active
    };

    return {
      success: true,
      user: safeUser,
      session: {
        token: session.token,
        expiresAt: session.expiresAt
      }
    };
  },

  setPassword(email, newPassword, currentPassword) {
    email = (email || '').toLowerCase().trim();
    const user = this.getUserWithPassword(email);

    if (!user) {
      return { success: false, error: 'User not found' };
    }

    if (user.passwordHash && user.passwordSalt) {
      if (!currentPassword) {
        return { success: false, error: 'Current password is required' };
      }

      const currentValid = PasswordService.verifyPassword(
        currentPassword,
        user.passwordHash,
        user.passwordSalt
      );

      if (!currentValid) {
        return { success: false, error: 'Current password is incorrect' };
      }
    }

    try {
      const { hash, salt } = PasswordService.createPasswordHash(newPassword);
      this.updateUserPassword(email, hash, salt);
      this.invalidateAllUserSessions(email);
      return { success: true };
    } catch (error) {
      return { success: false, error: error.message };
    }
  },

  sendMfaCode(user) {
    try {
      const code = this.generateMfaCode();
      const expiresAt = new Date(Date.now() + this.MFA_CODE_EXPIRY_MS);

      this.saveMfaCode({
        userId: user.email,
        code: code,
        createdAt: new Date().toISOString(),
        expiresAt: expiresAt.toISOString(),
        used: false,
        channel: 'email'
      });

      if (typeof EmailNotificationService !== 'undefined') {
        EmailNotificationService.sendMfaCode(user.email, code, user.name);
      } else {
        GmailApp.sendEmail(
          user.email,
          'ProjectFlow Login Code',
          `Your login verification code is: ${code}\n\nThis code expires in 5 minutes.`,
          {
            name: 'ProjectFlow',
            htmlBody: `
            <div style="font-family: sans-serif; padding: 20px;">
            <h2 style="color: #f59e0b;">ProjectFlow Login Code</h2>
            <p>Your verification code is:</p>
            <p style="font-size: 32px; font-weight: bold; letter-spacing: 8px; color: #1f2937;">${code}</p>
            <p style="color: #6b7280;">This code expires in 5 minutes.</p>
            </div>
            `
          }
        );
      }

      return true;
    } catch (error) {
      console.error('Failed to send MFA code:', error);
      return false;
    }
  },

  generateMfaCode() {
    let code = '';
    for (let i = 0; i < this.MFA_CODE_LENGTH; i++) {
      code += Math.floor(Math.random() * 10);
    }
    return code;
  },

  verifyMfaCode(email, code) {
    const sheet = this.getMfaCodesSheet();
    const data = sheet.getDataRange().getValues();
    const columns = CONFIG.MFA_CODE_COLUMNS;
    const userIdIndex = columns.indexOf('userId');
    const codeIndex = columns.indexOf('code');
    const expiresAtIndex = columns.indexOf('expiresAt');
    const usedIndex = columns.indexOf('used');
    const now = new Date();

    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      if (row[userIdIndex] === email && row[codeIndex] === code) {
        const expiresAt = new Date(row[expiresAtIndex]);
        if (expiresAt > now && !row[usedIndex]) {
          sheet.getRange(i + 1, usedIndex + 1).setValue(true);
          return true;
        }
      }
    }

    return false;
  },

  getSessionTokensSheet() {
    return getSheet(CONFIG.SHEETS.SESSION_TOKENS, CONFIG.SESSION_TOKEN_COLUMNS);
  },

  getMfaCodesSheet() {
    return getSheet(CONFIG.SHEETS.MFA_CODES, CONFIG.MFA_CODE_COLUMNS);
  },

  saveSession(session) {
    const sheet = this.getSessionTokensSheet();
    sheet.appendRow(objectToRow(session, CONFIG.SESSION_TOKEN_COLUMNS));
  },

  getSessionByToken(token, userId) {
    const sheet = this.getSessionTokensSheet();
    const data = sheet.getDataRange().getValues();
    const columns = CONFIG.SESSION_TOKEN_COLUMNS;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const session = rowToObject(row, columns);
      if (session.token === token && session.userId === userId) {
        return session;
      }
    }

    return null;
  },

  saveMfaCode(mfaData) {
    const sheet = this.getMfaCodesSheet();
    const data = {
      id: generateId('mfa'),
      ...mfaData
    };
    sheet.appendRow(objectToRow(data, CONFIG.MFA_CODE_COLUMNS));
  },

  getUserWithPassword(email) {
    const sheet = getUsersSheet();
    const data = sheet.getDataRange().getValues();
    const columns = CONFIG.USER_COLUMNS;
    const emailIndex = 0;

    for (let i = 1; i < data.length; i++) {
      if (data[i][emailIndex]?.toLowerCase() === email.toLowerCase()) {
        return rowToObject(data[i], columns);
      }
    }

    return null;
  },

  updateUserPassword(email, hash, salt) {
    const sheet = getUsersSheet();
    const data = sheet.getDataRange().getValues();
    const columns = CONFIG.USER_COLUMNS;
    const emailIndex = 0;
    const hashIndex = columns.indexOf('passwordHash');
    const saltIndex = columns.indexOf('passwordSalt');
    const lastChangeIndex = columns.indexOf('lastPasswordChange');

    for (let i = 1; i < data.length; i++) {
      if (data[i][emailIndex]?.toLowerCase() === email.toLowerCase()) {
        if (hashIndex !== -1) sheet.getRange(i + 1, hashIndex + 1).setValue(hash);
        if (saltIndex !== -1) sheet.getRange(i + 1, saltIndex + 1).setValue(salt);
        if (lastChangeIndex !== -1) sheet.getRange(i + 1, lastChangeIndex + 1).setValue(new Date().toISOString());
        return;
      }
    }
  },

  updateUserLockout(email, lockoutUntil, failedAttempts) {
    const sheet = getUsersSheet();
    const data = sheet.getDataRange().getValues();
    const columns = CONFIG.USER_COLUMNS;
    const emailIndex = 0;
    const failedIndex = columns.indexOf('failedLoginAttempts');
    const lockoutIndex = columns.indexOf('lockoutUntil');

    for (let i = 1; i < data.length; i++) {
      if (data[i][emailIndex]?.toLowerCase() === email.toLowerCase()) {
        if (failedIndex !== -1) sheet.getRange(i + 1, failedIndex + 1).setValue(failedAttempts);
        if (lockoutIndex !== -1) sheet.getRange(i + 1, lockoutIndex + 1).setValue(lockoutUntil.toISOString());
        return;
      }
    }
  },

  clearUserLockout(email) {
    const sheet = getUsersSheet();
    const data = sheet.getDataRange().getValues();
    const columns = CONFIG.USER_COLUMNS;
    const emailIndex = 0;
    const failedIndex = columns.indexOf('failedLoginAttempts');
    const lockoutIndex = columns.indexOf('lockoutUntil');

    for (let i = 1; i < data.length; i++) {
      if (data[i][emailIndex]?.toLowerCase() === email.toLowerCase()) {
        if (failedIndex !== -1) sheet.getRange(i + 1, failedIndex + 1).setValue(0);
        if (lockoutIndex !== -1) sheet.getRange(i + 1, lockoutIndex + 1).setValue('');
        return;
      }
    }
  },

  updateLastLogin(email) {
    try {
      const sheet = getUsersSheet();
      const data = sheet.getDataRange().getValues();
      const columns = CONFIG.USER_COLUMNS;
      const emailIndex = 0;
      const lastLoginIndex = columns.indexOf('lastLogin');

      if (lastLoginIndex === -1) return;

      for (let i = 1; i < data.length; i++) {
        if (data[i][emailIndex]?.toLowerCase() === email.toLowerCase()) {
          sheet.getRange(i + 1, lastLoginIndex + 1).setValue(new Date().toISOString());
          return;
        }
      }
    } catch (error) {
      console.error('Failed to update last login:', error);
    }
  }
};

if (typeof module !== 'undefined' && module.exports) {
  module.exports = AuthService;
}
