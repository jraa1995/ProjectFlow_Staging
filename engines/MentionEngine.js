class MentionEngine {
  static parseCommentForMentions(commentText) {
    if (!commentText || typeof commentText !== 'string') {
      return {
        originalText: commentText || '',
        mentionedUsers: [],
        formattedText: commentText || '',
        isValid: true
      };
    }

    const mentionedUsers = [];
    const invalidMentions = [];
    let formattedText = commentText;
    let match;

    const namedMentionPattern = /@\[([^\]]+)\]\(([^)]+)\)/g;
    while ((match = namedMentionPattern.exec(commentText)) !== null) {
      const fullMatch = match[0];
      const displayName = match[1];
      const emailAddress = match[2];
      if (!mentionedUsers.find(m => m.email === emailAddress)) {
        const user = getUserByEmail(emailAddress);
        if (user && user.active) {
          mentionedUsers.push({
            email: emailAddress,
            name: user.name || displayName,
            fullMatch: fullMatch,
            isValid: true
          });
        } else {
          invalidMentions.push({
            email: emailAddress,
            fullMatch: fullMatch,
            reason: user ? 'User is inactive' : 'User not found'
          });
        }
      }
    }

    const emailMentionPattern = /@([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/g;
    while ((match = emailMentionPattern.exec(commentText)) !== null) {
      const fullMatch = match[0];
      const emailAddress = match[1];
      if (!mentionedUsers.find(m => m.email === emailAddress) &&
          !invalidMentions.find(m => m.email === emailAddress)) {
        const user = getUserByEmail(emailAddress);
        if (user && user.active) {
          mentionedUsers.push({
            email: emailAddress,
            name: user.name,
            fullMatch: fullMatch,
            isValid: true
          });
        } else {
          invalidMentions.push({
            email: emailAddress,
            fullMatch: fullMatch,
            reason: user ? 'User is inactive' : 'User not found'
          });
        }
      }
    }

    const nameMentionPattern = /@([A-Za-z][A-Za-z0-9 ]*[A-Za-z0-9])(?=\s|$|[.,!?])/g;
    const allUsers = getActiveUsers();
    while ((match = nameMentionPattern.exec(commentText)) !== null) {
      const fullMatch = match[0];
      const mentionedName = match[1].trim();

      if (mentionedName.includes('@')) continue;
      if (mentionedUsers.find(m => m.fullMatch === fullMatch)) continue;

      const user = allUsers.find(u =>
        u.name && u.name.toLowerCase() === mentionedName.toLowerCase()
      );

      if (user && user.active) {
        if (!mentionedUsers.find(m => m.email === user.email)) {
          mentionedUsers.push({
            email: user.email,
            name: user.name,
            fullMatch: fullMatch,
            isValid: true
          });
        }
      } else {
        if (mentionedName[0] === mentionedName[0].toUpperCase()) {
          invalidMentions.push({
            name: mentionedName,
            fullMatch: fullMatch,
            reason: 'User not found by name'
          });
        }
      }
    }

    mentionedUsers.forEach(mention => {
      const highlightedMention = '<span class="mention" data-user="' + mention.email + '">@' + mention.name + '</span>';
      formattedText = formattedText.replace(
        new RegExp(escapeRegExp(mention.fullMatch), 'g'),
        highlightedMention
      );
    });

    return {
      originalText: commentText,
      mentionedUsers: mentionedUsers,
      invalidMentions: invalidMentions,
      formattedText: formattedText,
      isValid: invalidMentions.length === 0
    };
  }

  static getUserSuggestions(query, excludeUsers = []) {
    if (!query || query.length < 1) {
      return [];
    }

    const activeUsers = getActiveUsers();
    const queryLower = query.toLowerCase();
    const excludeSet = new Set(excludeUsers.map(email => email.toLowerCase()));

    const suggestions = activeUsers
      .filter(user => {
        if (excludeSet.has(user.email.toLowerCase())) {
          return false;
        }
        const emailMatch = user.email.toLowerCase().includes(queryLower);
        const nameMatch = user.name.toLowerCase().includes(queryLower);
        return emailMatch || nameMatch;
      })
      .map(user => ({
        email: user.email,
        name: user.name,
        displayText: `${user.name} (${user.email})`,
        matchScore: this.calculateMatchScore(query, user)
      }))
      .sort((a, b) => b.matchScore - a.matchScore)
      .slice(0, 10);

    return suggestions;
  }

  static calculateMatchScore(query, user) {
    const queryLower = query.toLowerCase();
    const emailLower = user.email.toLowerCase();
    const nameLower = user.name.toLowerCase();
    let score = 0;

    if (emailLower === queryLower || nameLower === queryLower) {
      score += 100;
    }
    if (emailLower.startsWith(queryLower) || nameLower.startsWith(queryLower)) {
      score += 50;
    }
    if (emailLower.includes(queryLower)) {
      score += 20;
    }
    if (nameLower.includes(queryLower)) {
      score += 15;
    }

    const avgLength = (user.email.length + user.name.length) / 2;
    score += Math.max(0, 50 - avgLength);
    return score;
  }

  static validateMentions(mentionedUsers) {
    if (!Array.isArray(mentionedUsers)) {
      return {
        isValid: false,
        validUsers: [],
        invalidUsers: [],
        errors: ['mentionedUsers must be an array']
      };
    }

    const validUsers = [];
    const invalidUsers = [];
    const errors = [];

    mentionedUsers.forEach(userRef => {
      const email = typeof userRef === 'string' ? userRef : userRef.email;
      if (!email || typeof email !== 'string') {
        invalidUsers.push({
          input: userRef,
          reason: 'Invalid email format'
        });
        return;
      }

      const user = getUserByEmail(email);
      if (!user) {
        invalidUsers.push({
          email: email,
          reason: 'User not found in directory'
        });
      } else if (!user.active) {
        invalidUsers.push({
          email: email,
          name: user.name,
          reason: 'User is inactive'
        });
      } else {
        validUsers.push({
          email: user.email,
          name: user.name,
          role: user.role
        });
      }
    });

    if (invalidUsers.length > 0) {
      errors.push(`${invalidUsers.length} invalid user(s) found`);
    }

    return {
      isValid: invalidUsers.length === 0,
      validUsers: validUsers,
      invalidUsers: invalidUsers,
      errors: errors
    };
  }

  static formatMentionsForDisplay(text, mentionedUsers = []) {
    if (!text) {
      return '';
    }

    let formattedText = text;

    formattedText = formattedText.replace(/@\[([^\]]+)\]\(([^)]+)\)/g, function(match, name, email) {
      return '<span class="mention" data-user="' + email + '" title="' + email + '">@' + name + '</span>';
    });

    if (Array.isArray(mentionedUsers) && mentionedUsers.length > 0) {
      mentionedUsers.forEach(function(user) {
        var email = typeof user === 'string' ? user : user.email;
        var name = typeof user === 'string' ? email : (user.name || email);
        var mentionPattern = new RegExp('@' + escapeRegExp(email) + '(?![^<]*>)', 'g');
        var highlightedMention = '<span class="mention" data-user="' + email + '" title="' + name + '">@' + name + '</span>';
        formattedText = formattedText.replace(mentionPattern, highlightedMention);
      });
    }

    formattedText = formattedText.replace(/@([A-Za-z][A-Za-z0-9 ]*[A-Za-z0-9])(?![^<]*<\/span>)(?=\s|$|[.,!?])/g, function(match, name) {
      return '<span class="mention">@' + name + '</span>';
    });

    return formattedText;
  }

  static extractPlainTextMentions(formattedText) {
    if (!formattedText) return '';
    return formattedText.replace(
      /<span class="mention"[^>]*>@([^<]+)<\/span>/g,
      '@$1'
    );
  }

  static getMentionStatistics(commentText) {
    const parseResult = this.parseCommentForMentions(commentText);
    return {
      totalMentions: parseResult.mentionedUsers.length + parseResult.invalidMentions.length,
      validMentions: parseResult.mentionedUsers.length,
      invalidMentions: parseResult.invalidMentions.length,
      uniqueUsers: parseResult.mentionedUsers.length,
      hasInvalidMentions: parseResult.invalidMentions.length > 0,
      mentionedUserEmails: parseResult.mentionedUsers.map(u => u.email)
    };
  }

  static createMentionNotifications(taskId, mentionedUsers, comment) {
    if (!Array.isArray(mentionedUsers) || mentionedUsers.length === 0) {
      return [];
    }

    const notifications = [];
    const task = getTaskById(taskId);
    if (!task) {
      console.error('createMentionNotifications: Task not found: ' + taskId);
      return [];
    }

    const taskTitle = task.title || 'Task ' + taskId;
    const mentionedBy = comment.userId || getCurrentUserEmail();
    const mentionedByUser = getUserByEmail(mentionedBy);
    const mentionedByName = mentionedByUser ? mentionedByUser.name : mentionedBy;

    mentionedUsers.forEach(user => {
      const email = typeof user === 'string' ? user : user.email;

      if (email.toLowerCase() === mentionedBy.toLowerCase()) {
        return;
      }

      try {
        const mention = createMention({
          commentId: comment.id,
          mentionedUserId: email,
          mentionedByUserId: mentionedBy,
          taskId: taskId
        });

        const notification = createNotification({
          userId: email,
          type: 'mention',
          title: 'You were mentioned in a comment',
          message: mentionedByName + ' mentioned you in a comment on "' + taskTitle + '"',
          entityType: 'task',
          entityId: taskId,
          channels: ['in_app', 'email']
        });

        notifications.push(notification);

        try {
          if (typeof EmailNotificationService !== 'undefined') {
            EmailNotificationService.sendMentionNotificationEmail(
              notification,
              task,
              comment,
              mentionedByUser
            );
          }
        } catch (emailError) {
          console.error('createMentionNotifications: Email error for ' + email + ': ' + emailError.message);
        }
      } catch (error) {
        console.error('createMentionNotifications: Error processing mention for ' + email + ': ' + error.message);
      }
    });

    return notifications;
  }
}

function escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function addCommentWithMentions(taskId, content, userId = null) {
  const currentUser = userId || getCurrentUserEmail();
  const timestamp = now();

  const mentionData = MentionEngine.parseCommentForMentions(content);
  const validation = MentionEngine.validateMentions(mentionData.mentionedUsers);

  const comment = {
    id: generateId('cmt'),
    taskId: taskId,
    userId: currentUser,
    content: sanitize(content),
    createdAt: timestamp,
    updatedAt: timestamp,
    mentionedUsers: validation.validUsers.map(u => u.email).join(','),
    isEdited: false,
    editHistory: JSON.stringify([])
  };

  const sheet = getCommentsSheet();
  const row = objectToRow(comment, CONFIG.COMMENT_COLUMNS);
  sheet.appendRow(row);
  SpreadsheetApp.flush();

  const notifications = MentionEngine.createMentionNotifications(
    taskId,
    validation.validUsers,
    comment
  );

  logActivity(currentUser, 'commented', 'task', taskId, {
    preview: content.substring(0, 50),
    mentions: validation.validUsers.length,
    notifications: notifications.length
  });

  comment.mentionedUsers = validation.validUsers.map(u => u.email);
  comment.mentionData = mentionData;
  comment.notifications = notifications;
  comment.formattedContent = MentionEngine.formatMentionsForDisplay(
    comment.content,
    validation.validUsers
  );

  return comment;
}

function getFormattedCommentContent(comment) {
  if (!comment || !comment.content) {
    return '';
  }

  const mentionedEmails = comment.mentionedUsers ?
    (typeof comment.mentionedUsers === 'string' ?
      comment.mentionedUsers.split(',').map(e => e.trim()) :
      comment.mentionedUsers) : [];

  const mentionedUsers = mentionedEmails.map(email => {
    const user = getUserByEmail(email);
    return user ? { email: user.email, name: user.name } : { email: email, name: email };
  });

  return MentionEngine.formatMentionsForDisplay(comment.content, mentionedUsers);
}
