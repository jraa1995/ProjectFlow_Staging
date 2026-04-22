function pushTaskToGoogleCalendar(taskData, options) {
  options = options || {};
  var result = { success: false, eventId: '', taskId: '', errors: [] };
  if (!taskData) {
    result.errors.push('No task data provided');
    return result;
  }
  var title = taskData.title || 'Task';
  var description = taskData.description || '';
  var dueDate = taskData.dueDate && taskData.dueDate !== 'TBD' ? taskData.dueDate : '';
  var startDate = taskData.startDate && taskData.startDate !== 'TBD' ? taskData.startDate : dueDate;
  var assignee = taskData.assignee || '';
  var watchers = taskData.watchers || '';
  var taskIdText = taskData.id ? (' [' + taskData.id + ']') : '';

  if (options.event) {
    try {
      if (!dueDate && !startDate) {
        result.errors.push('Event requires a Start or Due Date');
      } else {
        var cal = CalendarApp.getDefaultCalendar();
        var s = startDate ? new Date(startDate) : new Date(dueDate);
        var e = dueDate ? new Date(dueDate) : new Date(startDate);
        if (!isNaN(e.getTime())) e.setDate(e.getDate() + 1);
        var guests = [];
        if (assignee) guests.push(assignee);
        if (watchers) {
          watchers.split(',').map(function(w) { return w.trim(); }).filter(Boolean).forEach(function(w) {
            if (guests.indexOf(w) === -1) guests.push(w);
          });
        }
        var desc = description + (taskIdText ? '\n\nCOLONY Task: ' + taskIdText : '');
        var customFields = {};
        if (taskData.customFields) {
          try { customFields = typeof taskData.customFields === 'string' ? JSON.parse(taskData.customFields) : taskData.customFields; } catch (e2) { customFields = {}; }
        }
        var existingId = customFields.calendarEventId || '';
        var event = null;
        if (existingId) {
          try {
            var existing = cal.getEventById(existingId);
            if (existing) {
              existing.setTitle(title + taskIdText);
              existing.setDescription(desc);
              existing.setAllDayDates(s, e);
              guests.forEach(function(g) { try { existing.addGuest(g); } catch (_) {} });
              event = existing;
            }
          } catch (lookupErr) {
            console.error('pushTaskToGoogleCalendar: event lookup failed, creating new:', lookupErr);
          }
        }
        if (!event) {
          var eventOpts = { description: desc };
          if (guests.length) eventOpts.guests = guests.join(',');
          eventOpts.sendInvites = true;
          event = cal.createAllDayEvent(title + taskIdText, s, e, eventOpts);
          customFields.calendarEventId = event.getId();
          if (taskData.id) {
            try { updateTask(taskData.id, { customFields: JSON.stringify(customFields) }); } catch (persistErr) { console.error('pushTaskToGoogleCalendar: persist customFields failed:', persistErr); }
          }
        }
        result.eventId = event.getId();
        result.success = true;
      }
    } catch (err) {
      console.error('pushTaskToGoogleCalendar (event) failed:', err);
      result.errors.push('Event: ' + err.message);
    }
  }

  if (options.task) {
    try {
      if (typeof Tasks === 'undefined' || !Tasks.Tasks) {
        result.errors.push('Google Tasks API is not enabled. Enable it in Apps Script Advanced Services.');
      } else {
        var taskLists = Tasks.Tasklists.list();
        var taskListId = '@default';
        if (taskLists && taskLists.items && taskLists.items.length) {
          taskListId = taskLists.items[0].id;
        }
        var tItem = { title: title + taskIdText, notes: description };
        if (dueDate) {
          var d = new Date(dueDate);
          tItem.due = d.toISOString();
        }
        var created = Tasks.Tasks.insert(tItem, taskListId);
        result.taskId = created.id;
        result.success = true;
      }
    } catch (err) {
      console.error('pushTaskToGoogleCalendar (task) failed:', err);
      result.errors.push('Task: ' + err.message);
    }
  }

  if (!options.event && !options.task) {
    result.errors.push('No calendar target selected');
  }
  return result;
}

function pushExistingTaskToCalendar(taskId, options) {
  try {
    var task = loadTask(taskId);
    if (!task) return { success: false, errors: ['Task not found'] };
    return pushTaskToGoogleCalendar(task, options);
  } catch (err) {
    console.error('pushExistingTaskToCalendar failed:', err);
    return { success: false, errors: [err.message] };
  }
}
