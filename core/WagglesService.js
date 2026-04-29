function _wgGetQuestionsSheet_() {
  return getSheet(CONFIG.SHEETS.WAGGLES_QUESTIONS, CONFIG.WAGGLES_QUESTION_COLUMNS);
}

function _wgGetAnswersSheet_() {
  return getSheet(CONFIG.SHEETS.WAGGLES_ANSWERS, CONFIG.WAGGLES_ANSWER_COLUMNS);
}

function _wgGetVotesSheet_() {
  return getSheet(CONFIG.SHEETS.WAGGLES_VOTES, CONFIG.WAGGLES_VOTE_COLUMNS);
}

function _wgGetCommentsSheet_() {
  return getSheet(CONFIG.SHEETS.WAGGLES_COMMENTS, CONFIG.WAGGLES_COMMENT_COLUMNS);
}

function _wgRequireAuthor_(targetEmail) {
  var current = getCurrentUserEmailOptimized();
  if (!current) throw new Error('Not authenticated');
  if (String(current).toLowerCase() !== String(targetEmail || '').toLowerCase()) {
    if (!PermissionGuard.hasPermission('admin:settings', { userEmail: current })) {
      throw new Error('Only the author or an admin can perform this action');
    }
  }
  return current;
}

function _wgRequireWrite_() {
  var email = getCurrentUserEmailOptimized();
  if (!email) throw new Error('Not authenticated');
  return email;
}

function _wgResolveAuthorName_(email) {
  try {
    var user = getUserByEmailOptimized(email);
    if (user && user.name) return user.name;
  } catch (e) {}
  if (!email) return '';
  var at = String(email).indexOf('@');
  return at > 0 ? String(email).substring(0, at) : String(email);
}

function _wgNormalizeTags_(tags) {
  if (!tags) return [];
  var arr = Array.isArray(tags) ? tags : String(tags).split(',');
  var seen = {};
  var out = [];
  for (var i = 0; i < arr.length; i++) {
    var t = String(arr[i] || '').trim().toLowerCase().replace(/\s+/g, '-').replace(/[^a-z0-9-]/g, '');
    if (!t || t.length > 30 || seen[t]) continue;
    seen[t] = true;
    out.push(t);
    if (out.length >= 5) break;
  }
  return out;
}

function _wgReadAllQuestions_() {
  var sheet = _wgGetQuestionsSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 1, lastRow - 1, CONFIG.WAGGLES_QUESTION_COLUMNS.length).getValues();
  var out = [];
  for (var i = 0; i < data.length; i++) {
    if (!data[i][0]) continue;
    var q = rowToObject(data[i], CONFIG.WAGGLES_QUESTION_COLUMNS);
    if (q.status === 'deleted') continue;
    q.tags = q.tags ? String(q.tags).split(',').filter(Boolean) : [];
    q.viewCount = parseInt(q.viewCount, 10) || 0;
    out.push(q);
  }
  return out;
}

function _wgReadAllAnswers_() {
  var sheet = _wgGetAnswersSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 1, lastRow - 1, CONFIG.WAGGLES_ANSWER_COLUMNS.length).getValues();
  var out = [];
  for (var i = 0; i < data.length; i++) {
    if (!data[i][0]) continue;
    var a = rowToObject(data[i], CONFIG.WAGGLES_ANSWER_COLUMNS);
    a.isAccepted = a.isAccepted === true || a.isAccepted === 'true';
    out.push(a);
  }
  return out;
}

function _wgReadAllVotes_() {
  var sheet = _wgGetVotesSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 1, lastRow - 1, CONFIG.WAGGLES_VOTE_COLUMNS.length).getValues();
  var out = [];
  for (var i = 0; i < data.length; i++) {
    if (!data[i][0]) continue;
    var v = rowToObject(data[i], CONFIG.WAGGLES_VOTE_COLUMNS);
    v.value = parseInt(v.value, 10) || 0;
    out.push(v);
  }
  return out;
}

function _wgReadCommentsForTargets_(targetIds) {
  var sheet = _wgGetCommentsSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 1, lastRow - 1, CONFIG.WAGGLES_COMMENT_COLUMNS.length).getValues();
  var lookup = {};
  if (targetIds && targetIds.length) {
    for (var k = 0; k < targetIds.length; k++) lookup[targetIds[k]] = true;
  }
  var out = [];
  for (var i = 0; i < data.length; i++) {
    if (!data[i][0]) continue;
    var c = rowToObject(data[i], CONFIG.WAGGLES_COMMENT_COLUMNS);
    if (targetIds && targetIds.length && !lookup[c.targetId]) continue;
    out.push(c);
  }
  return out;
}

function _wgScoresByTarget_(votes) {
  var latest = {};
  for (var i = 0; i < votes.length; i++) {
    var v = votes[i];
    if (!v.targetId) continue;
    var key = String(v.voterEmail || '').toLowerCase() + '|' + v.targetId;
    var ts = new Date(v.updatedAt || v.createdAt || 0).getTime();
    if (!latest[key] || ts >= latest[key]._ts) {
      latest[key] = { targetId: v.targetId, value: v.value, voter: String(v.voterEmail || '').toLowerCase(), _ts: ts };
    }
  }
  var scores = {};
  var voted = {};
  Object.keys(latest).forEach(function(k) {
    var row = latest[k];
    scores[row.targetId] = (scores[row.targetId] || 0) + row.value;
    voted[row.voter + '|' + row.targetId] = row.value;
  });
  return { scores: scores, byVoter: voted };
}

function _wgUpdateRowById_(sheet, columns, idValue, updates) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  var idCol = columns.indexOf('id');
  if (idCol === -1) return false;
  var ids = sheet.getRange(2, idCol + 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0] === idValue) {
      var rowIdx = i + 2;
      var range = sheet.getRange(rowIdx, 1, 1, columns.length);
      var current = range.getValues()[0];
      var merged = rowToObject(current, columns);
      Object.keys(updates).forEach(function(k) { merged[k] = updates[k]; });
      range.setValues([objectToRow(merged, columns)]);
      return true;
    }
  }
  return false;
}

function _wgFindRowById_(sheet, columns, idValue) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  var data = sheet.getRange(2, 1, lastRow - 1, columns.length).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === idValue) return rowToObject(data[i], columns);
  }
  return null;
}

function _wgTokenize_(text) {
  if (!text) return [];
  return String(text)
    .toLowerCase()
    .replace(/[^a-z0-9\s-]/g, ' ')
    .split(/\s+/)
    .filter(function(t) { return t.length >= 2 && t.length <= 30; });
}

function searchWagglesQuestionIds(query) {
  var tokens = _wgTokenize_(query);
  if (!tokens.length) return null;
  var questions = _wgReadAllQuestions_();
  var matches = [];
  for (var i = 0; i < questions.length; i++) {
    var q = questions[i];
    var hay = (q.title + ' ' + q.body + ' ' + (q.tags || []).join(' ')).toLowerCase();
    var hits = 0;
    for (var t = 0; t < tokens.length; t++) {
      if (hay.indexOf(tokens[t]) !== -1) hits++;
    }
    if (hits > 0) matches.push({ id: q.id, score: hits });
  }
  matches.sort(function(a, b) { return b.score - a.score; });
  return matches.map(function(m) { return m.id; });
}

function getWagglesList(filters) {
  try {
    _wgRequireWrite_();
    filters = filters || {};
    var questions = _wgReadAllQuestions_();
    var votes = _wgReadAllVotes_();
    var answers = _wgReadAllAnswers_();
    var scoreMap = _wgScoresByTarget_(votes).scores;

    var answerCounts = {};
    for (var i = 0; i < answers.length; i++) {
      var a = answers[i];
      answerCounts[a.questionId] = (answerCounts[a.questionId] || 0) + 1;
    }

    if (filters.tag) {
      var t = String(filters.tag).toLowerCase();
      questions = questions.filter(function(q) { return (q.tags || []).indexOf(t) !== -1; });
    }
    if (filters.projectId) {
      questions = questions.filter(function(q) { return q.projectId === filters.projectId; });
    }
    if (filters.unanswered === true) {
      questions = questions.filter(function(q) { return !answerCounts[q.id]; });
    }
    if (filters.search) {
      var matchIds = searchWagglesQuestionIds(filters.search);
      if (matchIds) {
        var allow = {};
        for (var s = 0; s < matchIds.length; s++) allow[matchIds[s]] = true;
        questions = questions.filter(function(q) { return allow[q.id]; });
      }
    }

    var enriched = questions.map(function(q) {
      return {
        id: q.id,
        title: q.title,
        bodyPreview: String(q.body || '').substring(0, 240),
        authorEmail: q.authorEmail,
        authorName: q.authorName,
        tags: q.tags,
        projectId: q.projectId,
        taskId: q.taskId,
        dataAssetId: q.dataAssetId,
        status: q.status,
        acceptedAnswerId: q.acceptedAnswerId,
        viewCount: q.viewCount,
        answerCount: answerCounts[q.id] || 0,
        score: scoreMap[q.id] || 0,
        createdAt: q.createdAt,
        lastActivityAt: q.lastActivityAt || q.createdAt
      };
    });

    enriched.sort(function(a, b) {
      var ta = new Date(a.lastActivityAt || a.createdAt).getTime();
      var tb = new Date(b.lastActivityAt || b.createdAt).getTime();
      return tb - ta;
    });

    var tagCounts = {};
    for (var qi = 0; qi < enriched.length; qi++) {
      (enriched[qi].tags || []).forEach(function(tag) {
        tagCounts[tag] = (tagCounts[tag] || 0) + 1;
      });
    }
    var topTags = Object.keys(tagCounts).map(function(name) {
      return { name: name, count: tagCounts[name] };
    }).sort(function(a, b) { return b.count - a.count; }).slice(0, 24);

    return { success: true, questions: enriched, topTags: topTags };
  } catch (error) {
    console.error('getWagglesList failed:', error);
    return { success: false, error: error.message, questions: [], topTags: [] };
  }
}

function getWagglesQuestionDetails(questionId) {
  try {
    var viewerEmail = _wgRequireWrite_();
    var qSheet = _wgGetQuestionsSheet_();
    var question = _wgFindRowById_(qSheet, CONFIG.WAGGLES_QUESTION_COLUMNS, questionId);
    if (!question || question.status === 'deleted') {
      return { success: false, error: 'Question not found' };
    }
    question.tags = question.tags ? String(question.tags).split(',').filter(Boolean) : [];

    var answers = _wgReadAllAnswers_().filter(function(a) { return a.questionId === questionId; });
    var votes = _wgReadAllVotes_();
    var voteData = _wgScoresByTarget_(votes);
    var scoreMap = voteData.scores;
    var byVoter = voteData.byVoter;

    var targetIds = [questionId];
    answers.forEach(function(a) { targetIds.push(a.id); });
    var comments = _wgReadCommentsForTargets_(targetIds);
    var commentsByTarget = {};
    for (var i = 0; i < comments.length; i++) {
      var c = comments[i];
      if (!commentsByTarget[c.targetId]) commentsByTarget[c.targetId] = [];
      commentsByTarget[c.targetId].push(c);
    }

    answers.sort(function(a, b) {
      if (a.isAccepted && !b.isAccepted) return -1;
      if (!a.isAccepted && b.isAccepted) return 1;
      var sb = (scoreMap[b.id] || 0) - (scoreMap[a.id] || 0);
      if (sb !== 0) return sb;
      return new Date(a.createdAt).getTime() - new Date(b.createdAt).getTime();
    });

    var enrichedAnswers = answers.map(function(a) {
      return {
        id: a.id,
        body: a.body,
        authorEmail: a.authorEmail,
        authorName: a.authorName,
        isAccepted: a.isAccepted,
        createdAt: a.createdAt,
        updatedAt: a.updatedAt,
        score: scoreMap[a.id] || 0,
        myVote: byVoter[viewerEmail + '|' + a.id] || 0,
        comments: commentsByTarget[a.id] || []
      };
    });

    try {
      _wgUpdateRowById_(qSheet, CONFIG.WAGGLES_QUESTION_COLUMNS, questionId, {
        viewCount: (parseInt(question.viewCount, 10) || 0) + 1
      });
    } catch (e) {}

    return {
      success: true,
      question: {
        id: question.id,
        title: question.title,
        body: question.body,
        authorEmail: question.authorEmail,
        authorName: question.authorName,
        tags: question.tags,
        projectId: question.projectId,
        taskId: question.taskId,
        dataAssetId: question.dataAssetId,
        status: question.status,
        acceptedAnswerId: question.acceptedAnswerId,
        viewCount: (parseInt(question.viewCount, 10) || 0) + 1,
        score: scoreMap[question.id] || 0,
        myVote: byVoter[viewerEmail + '|' + question.id] || 0,
        createdAt: question.createdAt,
        updatedAt: question.updatedAt,
        lastActivityAt: question.lastActivityAt,
        comments: commentsByTarget[question.id] || []
      },
      answers: enrichedAnswers
    };
  } catch (error) {
    console.error('getWagglesQuestionDetails failed:', error);
    return { success: false, error: error.message };
  }
}

function createWagglesQuestion(payload) {
  try {
    var email = _wgRequireWrite_();
    payload = payload || {};
    var title = sanitize(String(payload.title || '').trim());
    var body = sanitize(String(payload.body || '').trim());
    if (title.length < 8) return { success: false, error: 'Title must be at least 8 characters' };
    if (body.length < 15) return { success: false, error: 'Body must be at least 15 characters' };
    if (title.length > 200) return { success: false, error: 'Title is too long (max 200)' };

    var question = {
      id: generateId('wgq'),
      title: title,
      body: body,
      authorEmail: email,
      authorName: _wgResolveAuthorName_(email),
      tags: _wgNormalizeTags_(payload.tags).join(','),
      projectId: payload.projectId || '',
      taskId: payload.taskId || '',
      dataAssetId: payload.dataAssetId || '',
      status: 'open',
      acceptedAnswerId: '',
      viewCount: 0,
      createdAt: now(),
      updatedAt: now(),
      lastActivityAt: now()
    };

    var sheet = _wgGetQuestionsSheet_();
    sheet.appendRow(objectToRow(question, CONFIG.WAGGLES_QUESTION_COLUMNS));
    SpreadsheetApp.flush();

    try {
      logAuditEvent('waggles.question.create', 'waggles_question', question.id, question.title, {
        tagCount: (question.tags ? question.tags.split(',').length : 0),
        projectId: question.projectId
      });
    } catch (e) {}

    return { success: true, questionId: question.id };
  } catch (error) {
    console.error('createWagglesQuestion failed:', error);
    return { success: false, error: error.message };
  }
}

function editWagglesQuestion(questionId, payload) {
  try {
    var sheet = _wgGetQuestionsSheet_();
    var existing = _wgFindRowById_(sheet, CONFIG.WAGGLES_QUESTION_COLUMNS, questionId);
    if (!existing) return { success: false, error: 'Question not found' };
    _wgRequireAuthor_(existing.authorEmail);

    var updates = { updatedAt: now(), lastActivityAt: now() };
    if (typeof payload.title === 'string') {
      var t = sanitize(payload.title.trim());
      if (t.length < 8 || t.length > 200) return { success: false, error: 'Invalid title length' };
      updates.title = t;
    }
    if (typeof payload.body === 'string') {
      var b = sanitize(payload.body.trim());
      if (b.length < 15) return { success: false, error: 'Body must be at least 15 characters' };
      updates.body = b;
    }
    if (payload.tags !== undefined) {
      updates.tags = _wgNormalizeTags_(payload.tags).join(',');
    }

    _wgUpdateRowById_(sheet, CONFIG.WAGGLES_QUESTION_COLUMNS, questionId, updates);
    SpreadsheetApp.flush();
    return { success: true };
  } catch (error) {
    console.error('editWagglesQuestion failed:', error);
    return { success: false, error: error.message };
  }
}

function deleteWagglesQuestion(questionId) {
  try {
    var sheet = _wgGetQuestionsSheet_();
    var existing = _wgFindRowById_(sheet, CONFIG.WAGGLES_QUESTION_COLUMNS, questionId);
    if (!existing) return { success: false, error: 'Question not found' };
    _wgRequireAuthor_(existing.authorEmail);
    _wgUpdateRowById_(sheet, CONFIG.WAGGLES_QUESTION_COLUMNS, questionId, {
      status: 'deleted',
      updatedAt: now()
    });
    try {
      logAuditEvent('waggles.question.delete', 'waggles_question', questionId, existing.title, {});
    } catch (e) {}
    return { success: true };
  } catch (error) {
    console.error('deleteWagglesQuestion failed:', error);
    return { success: false, error: error.message };
  }
}

function createWagglesAnswer(questionId, body) {
  try {
    var email = _wgRequireWrite_();
    var clean = sanitize(String(body || '').trim());
    if (clean.length < 15) return { success: false, error: 'Answer must be at least 15 characters' };

    var qSheet = _wgGetQuestionsSheet_();
    var question = _wgFindRowById_(qSheet, CONFIG.WAGGLES_QUESTION_COLUMNS, questionId);
    if (!question || question.status === 'deleted') return { success: false, error: 'Question not found' };
    if (question.status === 'closed') return { success: false, error: 'Question is closed' };

    var answer = {
      id: generateId('wga'),
      questionId: questionId,
      body: clean,
      authorEmail: email,
      authorName: _wgResolveAuthorName_(email),
      isAccepted: false,
      createdAt: now(),
      updatedAt: now()
    };

    var aSheet = _wgGetAnswersSheet_();
    aSheet.appendRow(objectToRow(answer, CONFIG.WAGGLES_ANSWER_COLUMNS));
    _wgUpdateRowById_(qSheet, CONFIG.WAGGLES_QUESTION_COLUMNS, questionId, { lastActivityAt: now() });
    SpreadsheetApp.flush();

    if (question.authorEmail && String(question.authorEmail).toLowerCase() !== String(email).toLowerCase()) {
      try {
        NotificationEngine.createNotification({
          userId: question.authorEmail,
          type: 'waggles_answer',
          title: 'New answer to your question',
          message: answer.authorName + ' answered: ' + question.title,
          entityType: 'waggles_question',
          entityId: questionId,
          channels: ['in_app', 'email']
        });
      } catch (e) {
        console.error('Waggles notify on answer failed:', e);
      }
    }

    return { success: true, answerId: answer.id };
  } catch (error) {
    console.error('createWagglesAnswer failed:', error);
    return { success: false, error: error.message };
  }
}

function editWagglesAnswer(answerId, body) {
  try {
    var sheet = _wgGetAnswersSheet_();
    var existing = _wgFindRowById_(sheet, CONFIG.WAGGLES_ANSWER_COLUMNS, answerId);
    if (!existing) return { success: false, error: 'Answer not found' };
    _wgRequireAuthor_(existing.authorEmail);
    var clean = sanitize(String(body || '').trim());
    if (clean.length < 15) return { success: false, error: 'Answer must be at least 15 characters' };
    _wgUpdateRowById_(sheet, CONFIG.WAGGLES_ANSWER_COLUMNS, answerId, {
      body: clean,
      updatedAt: now()
    });
    SpreadsheetApp.flush();
    return { success: true };
  } catch (error) {
    console.error('editWagglesAnswer failed:', error);
    return { success: false, error: error.message };
  }
}

function deleteWagglesAnswer(answerId) {
  try {
    var sheet = _wgGetAnswersSheet_();
    var existing = _wgFindRowById_(sheet, CONFIG.WAGGLES_ANSWER_COLUMNS, answerId);
    if (!existing) return { success: false, error: 'Answer not found' };
    _wgRequireAuthor_(existing.authorEmail);
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (ids[i][0] === answerId) {
          sheet.deleteRow(i + 2);
          break;
        }
      }
    }
    if (existing.questionId) {
      var qSheet = _wgGetQuestionsSheet_();
      var question = _wgFindRowById_(qSheet, CONFIG.WAGGLES_QUESTION_COLUMNS, existing.questionId);
      if (question && question.acceptedAnswerId === answerId) {
        _wgUpdateRowById_(qSheet, CONFIG.WAGGLES_QUESTION_COLUMNS, existing.questionId, {
          acceptedAnswerId: '',
          updatedAt: now()
        });
      }
    }
    SpreadsheetApp.flush();
    return { success: true };
  } catch (error) {
    console.error('deleteWagglesAnswer failed:', error);
    return { success: false, error: error.message };
  }
}

function acceptWagglesAnswer(questionId, answerId) {
  try {
    var qSheet = _wgGetQuestionsSheet_();
    var question = _wgFindRowById_(qSheet, CONFIG.WAGGLES_QUESTION_COLUMNS, questionId);
    if (!question) return { success: false, error: 'Question not found' };
    _wgRequireAuthor_(question.authorEmail);

    var aSheet = _wgGetAnswersSheet_();
    var answer = _wgFindRowById_(aSheet, CONFIG.WAGGLES_ANSWER_COLUMNS, answerId);
    if (!answer || answer.questionId !== questionId) return { success: false, error: 'Answer not found' };

    var lastRow = aSheet.getLastRow();
    if (lastRow >= 2) {
      var data = aSheet.getRange(2, 1, lastRow - 1, CONFIG.WAGGLES_ANSWER_COLUMNS.length).getValues();
      var qIdCol = CONFIG.WAGGLES_ANSWER_COLUMNS.indexOf('questionId');
      var acceptedCol = CONFIG.WAGGLES_ANSWER_COLUMNS.indexOf('isAccepted');
      for (var i = 0; i < data.length; i++) {
        if (data[i][qIdCol] === questionId && data[i][acceptedCol] === true) {
          aSheet.getRange(i + 2, acceptedCol + 1).setValue(false);
        }
      }
    }
    _wgUpdateRowById_(aSheet, CONFIG.WAGGLES_ANSWER_COLUMNS, answerId, {
      isAccepted: true,
      updatedAt: now()
    });
    _wgUpdateRowById_(qSheet, CONFIG.WAGGLES_QUESTION_COLUMNS, questionId, {
      acceptedAnswerId: answerId,
      updatedAt: now(),
      lastActivityAt: now()
    });
    SpreadsheetApp.flush();
    return { success: true };
  } catch (error) {
    console.error('acceptWagglesAnswer failed:', error);
    return { success: false, error: error.message };
  }
}

function voteWaggles(targetType, targetId, value) {
  try {
    var email = _wgRequireWrite_();
    if (targetType !== 'question' && targetType !== 'answer') {
      return { success: false, error: 'Invalid target type' };
    }
    var v = parseInt(value, 10);
    if (v !== -1 && v !== 0 && v !== 1) return { success: false, error: 'Invalid vote value' };

    var sheet = _wgGetVotesSheet_();
    var lastRow = sheet.getLastRow();
    var found = false;
    if (lastRow >= 2) {
      var data = sheet.getRange(2, 1, lastRow - 1, CONFIG.WAGGLES_VOTE_COLUMNS.length).getValues();
      var emailCol = CONFIG.WAGGLES_VOTE_COLUMNS.indexOf('voterEmail');
      var targetCol = CONFIG.WAGGLES_VOTE_COLUMNS.indexOf('targetId');
      for (var i = 0; i < data.length; i++) {
        if (data[i][targetCol] === targetId &&
            String(data[i][emailCol]).toLowerCase() === String(email).toLowerCase()) {
          var existing = rowToObject(data[i], CONFIG.WAGGLES_VOTE_COLUMNS);
          existing.value = v;
          existing.updatedAt = now();
          sheet.getRange(i + 2, 1, 1, CONFIG.WAGGLES_VOTE_COLUMNS.length)
            .setValues([objectToRow(existing, CONFIG.WAGGLES_VOTE_COLUMNS)]);
          found = true;
          break;
        }
      }
    }
    if (!found) {
      sheet.appendRow(objectToRow({
        id: generateId('wgv'),
        targetType: targetType,
        targetId: targetId,
        voterEmail: email,
        value: v,
        createdAt: now(),
        updatedAt: now()
      }, CONFIG.WAGGLES_VOTE_COLUMNS));
    }

    var votes = _wgReadAllVotes_();
    var score = 0;
    for (var k = 0; k < votes.length; k++) {
      if (votes[k].targetId === targetId) score += votes[k].value;
    }
    SpreadsheetApp.flush();
    return { success: true, score: score, myVote: v };
  } catch (error) {
    console.error('voteWaggles failed:', error);
    return { success: false, error: error.message };
  }
}

function addWagglesComment(targetType, targetId, body) {
  try {
    var email = _wgRequireWrite_();
    if (targetType !== 'question' && targetType !== 'answer') {
      return { success: false, error: 'Invalid target type' };
    }
    var clean = sanitize(String(body || '').trim());
    if (clean.length < 2) return { success: false, error: 'Comment is too short' };
    if (clean.length > 600) return { success: false, error: 'Comment is too long (max 600)' };

    var comment = {
      id: generateId('wgc'),
      targetType: targetType,
      targetId: targetId,
      body: clean,
      authorEmail: email,
      authorName: _wgResolveAuthorName_(email),
      createdAt: now()
    };
    _wgGetCommentsSheet_().appendRow(objectToRow(comment, CONFIG.WAGGLES_COMMENT_COLUMNS));

    if (targetType === 'question') {
      var qSheet = _wgGetQuestionsSheet_();
      _wgUpdateRowById_(qSheet, CONFIG.WAGGLES_QUESTION_COLUMNS, targetId, { lastActivityAt: now() });
    }
    SpreadsheetApp.flush();
    return { success: true, comment: comment };
  } catch (error) {
    console.error('addWagglesComment failed:', error);
    return { success: false, error: error.message };
  }
}

function getWagglesFormOptions() {
  try {
    _wgRequireWrite_();
    var projects = [];
    try {
      projects = getAllProjectsOptimized().map(function(p) {
        return { id: p.id, name: p.name || '' };
      });
    } catch (e) {}
    return { success: true, projects: projects };
  } catch (error) {
    console.error('getWagglesFormOptions failed:', error);
    return { success: false, error: error.message, projects: [] };
  }
}
