const REQUESTS_STATUS_MAP = {
  'AH-ST-001': 'To Do',
  'AH-ST-002': 'To Do',
  'AH-ST-003': 'In Progress',
  'AH-ST-004': 'Backlog',
  'AH-ST-005': 'Backlog',
  'AH-ST-006': 'Review',
  'AH-ST-007': 'Done',
  'AH-ST-008': 'Done',
  'AH-ST-009': 'Done',
  'AH-ST-010': 'Done',
  'AH-ST-011': 'Done'
};

const REQUESTS_COLUMN_MAPPING = {
  title: 'Request Name',
  description: 'Description Of Request',
  status: 'Status',
  priority: 'Priority',
  type: 'Request Type Category',
  assignee: 'Assigned To',
  dueDate: 'Target Completion Date',
  startDate: 'Start Date Time',
  estimatedHrs: 'Time To Complete Estimated Hours',
  actualHrs: 'Time To Complete Actual Hours'
};

function getRequestsWorkbookId() {
  return PropertiesService.getScriptProperties().getProperty('REQUESTS_WORKBOOK_ID') || '';
}

function setRequestsWorkbookId(workbookId) {
  PropertiesService.getScriptProperties().setProperty('REQUESTS_WORKBOOK_ID', workbookId);
  return true;
}

function getExistingRequestIds() {
  try {
    const sheet = getFunnelStagingSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return new Set();

    const headers = data[0];
    const rawDataCol = headers.indexOf('rawData');
    const sourceSheetCol = headers.indexOf('sourceSheetName');
    const ids = new Set();

    for (let i = 1; i < data.length; i++) {
      if (data[i][sourceSheetCol] !== 'Requests') continue;
      try {
        const raw = JSON.parse(data[i][rawDataCol]);
        if (raw['Request ID']) ids.add(String(raw['Request ID']));
      } catch (e) {}
    }
    return ids;
  } catch (error) {
    console.error('getExistingRequestIds error:', error);
    return new Set();
  }
}

function normalizeRequestStatus(statusCode) {
  return REQUESTS_STATUS_MAP[statusCode] || 'To Do';
}

function importFromRequestsWorkbook(workbookId) {
  try {
    if (workbookId) {
      setRequestsWorkbookId(workbookId);
    }

    const storedId = workbookId || getRequestsWorkbookId();
    if (!storedId) {
      return { success: false, error: 'No workbook ID configured', stagedCount: 0, skippedCount: 0 };
    }

    const result = importTicketsFromWorkbook(storedId, 'Requests');
    if (!result.success) {
      return { success: false, error: result.error, stagedCount: 0, skippedCount: 0 };
    }

    const existingIds = getExistingRequestIds();
    const newTickets = result.tickets.filter(t => !existingIds.has(String(t['Request ID'] || '')));
    const skippedCount = result.tickets.length - newTickets.length;

    newTickets.forEach(ticket => {
      if (ticket['Status']) {
        ticket['Status'] = normalizeRequestStatus(ticket['Status']);
      }
    });

    if (newTickets.length === 0) {
      return { success: true, stagedCount: 0, skippedCount: skippedCount };
    }

    const stageResult = stageTicketsInFunnel(newTickets, REQUESTS_COLUMN_MAPPING, storedId, 'Requests');

    return {
      success: stageResult.success,
      stagedCount: stageResult.stagedCount || 0,
      skippedCount: skippedCount,
      funnelIds: stageResult.funnelIds || [],
      error: stageResult.error
    };
  } catch (error) {
    console.error('importFromRequestsWorkbook error:', error);
    return { success: false, error: error.message, stagedCount: 0, skippedCount: 0 };
  }
}

function getFunnelStagingSheet() {
  const ss = getColonySpreadsheet_();
  let sheet = ss.getSheetByName('Funnel_Staging');

  if (!sheet) {
    sheet = ss.insertSheet('Funnel_Staging');
    const headers = [
      'id',
      'sourceWorkbookId',
      'sourceSheetName',
      'rawData',
      'mappedData',
      'importStatus',
      'assignedTo',
      'notes',
      'importedTaskId',
      'createdAt',
      'importedAt'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#525252')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
  }

  return sheet;
}

function importTicketsFromWorkbook(workbookId, sheetName) {
  try {
    const sourceSheet = SpreadsheetApp.openById(workbookId).getSheetByName(sheetName);

    if (!sourceSheet) {
      throw new Error(`Sheet "${sheetName}" not found in workbook`);
    }

    const data = sourceSheet.getDataRange().getValues();

    if (data.length < 2) {
      throw new Error('No data rows found (need at least headers + 1 row)');
    }

    const headers = data[0];
    const rows = data.slice(1);

    const tickets = rows.map((row, index) => {
      const ticket = {
        _rowIndex: index + 2,
        _hasData: row.some(cell => cell !== null && cell !== '')
      };
      headers.forEach((header, colIndex) => {
        if (header) {
          ticket[header] = row[colIndex];
        }
      });
      return ticket;
    }).filter(ticket => ticket._hasData);

    return {
      success: true,
      workbookId: workbookId,
      sheetName: sheetName,
      headers: headers.filter(h => h),
      tickets: tickets,
      count: tickets.length
    };
  } catch (error) {
    console.error('importTicketsFromWorkbook error:', error);
    return {
      success: false,
      error: error.message,
      tickets: [],
      count: 0
    };
  }
}

function stageTicketsInFunnel(tickets, mapping, workbookId, sheetName) {
  try {
    const sheet = getFunnelStagingSheet();
    const timestamp = now();

    const stagedTickets = tickets.map(ticket => {
      const funnelId = generateId('FNL');
      const mappedData = {};

      Object.keys(mapping).forEach(taskField => {
        const sourceColumn = mapping[taskField];
        if (sourceColumn && ticket[sourceColumn] !== undefined) {
          mappedData[taskField] = ticket[sourceColumn];
        }
      });

      const criticalUnmappedFields = [
        'businessJustification',
        'whatContractDoesItApplyTo',
        'assignedTeam',
        'requestor',
        'contractorPoc',
        'contractorPocEmail',
        'internalNotes',
        'topic',
        'requestId',
        'slaBreached',
        'escalated'
      ];

      let additionalContext = '';
      const metadata = {};
      const usedColumns = Object.values(mapping);

      criticalUnmappedFields.forEach(fieldName => {
        Object.keys(ticket).forEach(ticketKey => {
          const normalizedKey = ticketKey.toLowerCase().replace(/[_\s-]/g, '');
          const normalizedField = fieldName.toLowerCase().replace(/[_\s-]/g, '');
          if (normalizedKey === normalizedField ||
            normalizedKey.includes(normalizedField) ||
            normalizedField.includes(normalizedKey)) {
            if (!usedColumns.includes(ticketKey) && ticket[ticketKey]) {
              const displayName = fieldName.replace(/([A-Z])/g, ' $1').trim();
              additionalContext += `\n**${displayName}:** ${ticket[ticketKey]}`;
              metadata[fieldName] = ticket[ticketKey];
            }
          }
        });
      });

      if (additionalContext) {
        const originalDesc = mappedData.description || '';
        mappedData.description = originalDesc +
          '\n\n---\n**Additional Request Details:**' +
          additionalContext;
      }

      if (Object.keys(metadata).length > 0) {
        mappedData.customFields = JSON.stringify(metadata);
      }

      if (!mappedData.status) mappedData.status = 'To Do';
      if (!mappedData.priority) mappedData.priority = 'Medium';
      if (!mappedData.type) mappedData.type = 'Task';

      return {
        id: funnelId,
        sourceWorkbookId: workbookId,
        sourceSheetName: sheetName,
        rawData: JSON.stringify(ticket),
        mappedData: JSON.stringify(mappedData),
        importStatus: 'pending',
        assignedTo: mappedData.assignee || '',
        notes: '',
        importedTaskId: '',
        createdAt: timestamp,
        importedAt: ''
      };
    });

    const rows = stagedTickets.map(ticket => [
      ticket.id,
      ticket.sourceWorkbookId,
      ticket.sourceSheetName,
      ticket.rawData,
      ticket.mappedData,
      ticket.importStatus,
      ticket.assignedTo,
      ticket.notes,
      ticket.importedTaskId,
      ticket.createdAt,
      ticket.importedAt
    ]);

    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }

    return {
      success: true,
      stagedCount: stagedTickets.length,
      funnelIds: stagedTickets.map(t => t.id)
    };
  } catch (error) {
    console.error('stageTicketsInFunnel error:', error);
    return {
      success: false,
      error: error.message,
      stagedCount: 0
    };
  }
}

function getFunnelTickets(status) {
  try {
    const sheet = getFunnelStagingSheet();
    const data = sheet.getDataRange().getValues();

    if (data.length < 2) {
      return [];
    }

    const headers = data[0];
    const rows = data.slice(1);

    const tickets = rows.map(row => {
      const ticket = {};
      headers.forEach((header, index) => {
        ticket[header] = row[index];
      });

      try {
        ticket.rawData = ticket.rawData ? JSON.parse(ticket.rawData) : {};
        ticket.mappedData = ticket.mappedData ? JSON.parse(ticket.mappedData) : {};
      } catch (e) {
        ticket.rawData = {};
        ticket.mappedData = {};
      }

      return ticket;
    });

    if (status) {
      return tickets.filter(t => t.importStatus === status);
    }

    return tickets;
  } catch (error) {
    console.error('getFunnelTickets error:', error);
    return [];
  }
}

function updateFunnelTicket(funnelId, updates) {
  try {
    const sheet = getFunnelStagingSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    let rowIndex = -1;

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === funnelId) {
        rowIndex = i;
        break;
      }
    }

    if (rowIndex === -1) {
      throw new Error(`Funnel ticket not found: ${funnelId}`);
    }

    Object.keys(updates).forEach(field => {
      const colIndex = headers.indexOf(field);
      if (colIndex !== -1) {
        let value = updates[field];
        if ((field === 'mappedData' || field === 'rawData') && typeof value === 'object') {
          value = JSON.stringify(value);
        }
        sheet.getRange(rowIndex + 1, colIndex + 1).setValue(value);
      }
    });

    return true;
  } catch (error) {
    console.error('updateFunnelTicket error:', error);
    return false;
  }
}

function importFunnelTicketsToTasks(funnelIds, options) {
  try {
    options = options || {};
    const funnelTickets = getFunnelTickets().filter(t => funnelIds.includes(t.id));

    if (funnelTickets.length === 0) {
      throw new Error('No funnel tickets found with provided IDs');
    }

    const importedTasks = [];
    const errors = [];

    funnelTickets.forEach(funnelTicket => {
      try {
        const taskData = { ...funnelTicket.mappedData };

        if (!taskData.title) {
          throw new Error('Title is required');
        }

        if (!taskData.projectId) {
          taskData.projectId = '';
        }

        if (!taskData.assignee) {
          if (options.surgeUnassigned) {
            taskData.assignee = '';
            const existingLabels = taskData.labels ? String(taskData.labels).split(',').map(l => l.trim()).filter(l => l) : [];
            existingLabels.push('surge');
            taskData.labels = existingLabels;
          } else if (options.assignTo) {
            taskData.assignee = options.assignTo;
          }
        }

        const task = createTask(taskData);
        importedTasks.push(task);

        updateFunnelTicket(funnelTicket.id, {
          importStatus: 'imported',
          importedTaskId: task.id,
          importedAt: now()
        });
      } catch (error) {
        console.error(`Failed to import funnel ticket ${funnelTicket.id}:`, error);
        errors.push({
          funnelId: funnelTicket.id,
          error: error.message
        });
      }
    });

    return {
      success: errors.length === 0,
      importedCount: importedTasks.length,
      importedTasks: importedTasks,
      errors: errors
    };
  } catch (error) {
    console.error('importFunnelTicketsToTasks error:', error);
    return {
      success: false,
      importedCount: 0,
      error: error.message,
      errors: []
    };
  }
}

function deleteFunnelTickets(funnelIds) {
  try {
    const sheet = getFunnelStagingSheet();
    const data = sheet.getDataRange().getValues();
    const rowsToDelete = [];

    for (let i = data.length - 1; i >= 1; i--) {
      if (funnelIds.includes(data[i][0])) {
        rowsToDelete.push(i + 1);
      }
    }

    rowsToDelete.forEach(rowNum => {
      sheet.deleteRow(rowNum);
    });

    return true;
  } catch (error) {
    console.error('deleteFunnelTickets error:', error);
    return false;
  }
}

function saveColumnMapping(mappingConfig) {
  try {
    const { workbookId, sheetName, mapping } = mappingConfig;
    const cache = CacheService.getScriptCache();
    const key = `FUNNEL_MAPPING_${workbookId}_${sheetName}`;

    cache.put(key, JSON.stringify(mapping), 3600);
    PropertiesService.getUserProperties().setProperty(key, JSON.stringify(mapping));

    return {
      success: true,
      mapping: mapping
    };
  } catch (error) {
    console.error('saveColumnMapping error:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

function getColumnMapping(workbookId, sheetName) {
  try {
    const cache = CacheService.getScriptCache();
    const key = `FUNNEL_MAPPING_${workbookId}_${sheetName}`;

    let mappingJson = cache.get(key);

    if (!mappingJson) {
      mappingJson = PropertiesService.getUserProperties().getProperty(key);
      if (mappingJson) {
        cache.put(key, mappingJson, 3600);
      }
    }

    return mappingJson ? JSON.parse(mappingJson) : null;
  } catch (error) {
    console.error('getColumnMapping error:', error);
    return null;
  }
}

function getFunnelStats() {
  try {
    const tickets = getFunnelTickets();

    const stats = {
      total: tickets.length,
      pending: 0,
      reviewed: 0,
      imported: 0,
      rejected: 0
    };

    tickets.forEach(ticket => {
      const status = ticket.importStatus || 'pending';
      if (stats.hasOwnProperty(status)) {
        stats[status]++;
      }
    });

    return stats;
  } catch (error) {
    console.error('getFunnelStats error:', error);
    return {
      total: 0,
      pending: 0,
      reviewed: 0,
      imported: 0,
      rejected: 0
    };
  }
}
