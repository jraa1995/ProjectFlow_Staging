const PasswordService = {
  ITERATIONS: 10000,
  SALT_LENGTH: 32,
  HASH_ALGORITHM: Utilities.DigestAlgorithm.SHA_256,

  generateSalt() {
    const uuid = Utilities.getUuid();
    const timestamp = Date.now().toString();
    const combined = uuid + timestamp + Math.random().toString();
    const saltBytes = Utilities.computeDigest(this.HASH_ALGORITHM, combined);
    return Utilities.base64Encode(saltBytes);
  },

  hashPassword(password, salt) {
    if (!password || typeof password !== 'string') {
      throw new Error('Invalid password');
    }

    if (!salt || typeof salt !== 'string') {
      throw new Error('Invalid salt');
    }

    let hash = password + salt;

    for (let i = 0; i < this.ITERATIONS; i++) {
      const hashBytes = Utilities.computeDigest(this.HASH_ALGORITHM, hash);
      hash = Utilities.base64Encode(hashBytes);
    }

    return hash;
  },

  createPasswordHash(password) {
    this.validatePasswordStrength(password);
    const salt = this.generateSalt();
    const hash = this.hashPassword(password, salt);

    return {
      hash: hash,
      salt: salt
    };
  },

  verifyPassword(password, storedHash, salt) {
    if (!password || !storedHash || !salt) {
      return false;
    }

    try {
      const computedHash = this.hashPassword(password, salt);
      return this.constantTimeCompare(computedHash, storedHash);
    } catch (error) {
      console.error('Password verification error:', error);
      return false;
    }
  },

  constantTimeCompare(a, b) {
    if (typeof a !== 'string' || typeof b !== 'string') {
      return false;
    }

    let result = 0;
    const maxLength = Math.max(a.length, b.length);

    for (let i = 0; i < maxLength; i++) {
      const charA = i < a.length ? a.charCodeAt(i) : 0;
      const charB = i < b.length ? b.charCodeAt(i) : 0;
      result |= charA ^ charB;
    }

    result |= a.length ^ b.length;
    return result === 0;
  },

  validatePasswordStrength(password) {
    if (!password || typeof password !== 'string') {
      throw new Error('Password is required');
    }

    const minLength = 8;
    const errors = [];

    if (password.length < minLength) {
      errors.push(`Password must be at least ${minLength} characters`);
    }

    if (!/[A-Z]/.test(password)) {
      errors.push('Password must contain at least one uppercase letter');
    }

    if (!/[a-z]/.test(password)) {
      errors.push('Password must contain at least one lowercase letter');
    }

    if (!/[0-9]/.test(password)) {
      errors.push('Password must contain at least one number');
    }

    if (errors.length > 0) {
      throw new Error(errors.join('. '));
    }

    return true;
  },

  needsRehash(metadata) {
    return false;
  }
};

if (typeof module !== 'undefined' && module.exports) {
  module.exports = PasswordService;
}
