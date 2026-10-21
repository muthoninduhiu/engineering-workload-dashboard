function ensureAuditSheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.AUDIT_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.AUDIT_SHEET);
  }

  const headers = [
    'Timestamp',
    'User Email',
    'User Name',
    'Role',
    'Action',
    'Entity Type',
    'Task ID',
    'Task Name',
    'Product',
    'Source Sheet',
    'Week Start',
    'Field',
    'Old Value',
    'New Value',
    'Details',
    'Request ID',
    'Origin'
  ];

  const range = sheet.getRange(1, 1, 1, headers.length);
  const existing = range.getValues()[0];

  let mismatch = false;

  for (let i = 0; i < headers.length; i++) {
    if (clean_(existing[i]) !== headers[i]) {
      mismatch = true;
      break;
    }
  }

  if (mismatch) {
    range
      .setValues([headers])
      .setFontWeight('bold')
      .setBackground('#f1f5f9');
  }

  sheet.setFrozenRows(1);

  return sheet;
}

function auditValue_(value) {
  if (value === null || value === undefined) return '';

  if (value instanceof Date) {
    return value.toISOString();
  }

  if (typeof value === 'object') {
    try {
      return JSON.stringify(value);
    } catch (error) {
      return String(value);
    }
  }

  return String(value);
}

function logActivities_(entries, actor) {
  entries = Array.isArray(entries)
    ? entries.filter(Boolean)
    : [];

  if (!entries.length) return;

  actor = actor || getAccessContext_();

  const ss = getDashboardSpreadsheet_();
  const sheet = ensureAuditSheet_(ss);
  const now = new Date();
  const requestId = Utilities.getUuid();

  const rows = entries.map(function(entry) {
    try {
      console.log(
        '[AUDIT] ' +
        JSON.stringify({
          timestamp: now.toISOString(),
          userEmail: actor.email || '',
          userName: actor.name || '',
          role: actor.role || 'Viewer',
          action: clean_(entry.action),
          entityType: clean_(entry.entityType) || 'Task',
          taskId: clean_(entry.taskId),
          taskName: clean_(entry.taskName),
          product: clean_(entry.product),
          sourceSheet: clean_(entry.sourceSheet),
          weekStart: clean_(entry.weekStart),
          field: clean_(entry.field),
          oldValue: auditValue_(entry.oldValue),
          newValue: auditValue_(entry.newValue),
          details: clean_(entry.details),
          requestId: requestId,
          origin: clean_(entry.origin) || 'App'
        })
      );
    } catch (error) {
      console.warn('Could not write audit event to execution log.', error);
    }

    return [
      now,
      actor.email || '',
      actor.name || '',
      actor.role || 'Viewer',
      clean_(entry.action),
      clean_(entry.entityType) || 'Task',
      clean_(entry.taskId),
      clean_(entry.taskName),
      clean_(entry.product),
      clean_(entry.sourceSheet),
      clean_(entry.weekStart),
      clean_(entry.field),
      auditValue_(entry.oldValue),
      auditValue_(entry.newValue),
      clean_(entry.details),
      requestId,
      clean_(entry.origin) || 'App'
    ];
  });

  sheet
    .getRange(
      sheet.getLastRow() + 1,
      1,
      rows.length,
      17
    )
    .setValues(rows);
}

function getDirectAuditSources_() {
  const dashboard = getDashboardSpreadsheet_();
  const byId = {};

  function addSource(id, label) {
    id = clean_(id);
    if (!id) return;

    if (!byId[id]) {
      byId[id] = {
        spreadsheetId: id,
        label: label || id
      };
      return;
    }

    if (
      label &&
      byId[id].label.indexOf(label) === -1
    ) {
      byId[id].label +=
        ' / ' + label;
    }
  }

  addSource(
    dashboard.getId(),
    'Engineering Workload dashboard'
  );

  Object.keys(TASK_ROUTING).forEach(function(product) {
    TASK_ROUTING[product].forEach(function(route) {
      addSource(
        route.spreadsheetId,
        displayProductName_(product) + ' · ' + route.sourceSheet
      );
    });
  });

  return Object.keys(byId).map(function(id) {
    return byId[id];
  });
}

function getDirectAuditTriggerStatus_() {
  const sources = getDirectAuditSources_();
  const triggers = ScriptApp.getProjectTriggers();

  return sources.map(function(source) {
    let editInstalled = false;
    let changeInstalled = false;

    triggers.forEach(function(trigger) {
      const sourceId = clean_(
        trigger.getTriggerSourceId &&
        trigger.getTriggerSourceId()
      );

      if (sourceId !== source.spreadsheetId) {
        return;
      }

      const handler = clean_(
        trigger.getHandlerFunction()
      );

      if (
        handler ===
        CONFIG.DIRECT_AUDIT_EDIT_HANDLER
      ) {
        editInstalled = true;
      }

      if (
        handler ===
        CONFIG.DIRECT_AUDIT_CHANGE_HANDLER
      ) {
        changeInstalled = true;
      }
    });

    return {
      spreadsheetId: source.spreadsheetId,
      label: source.label,
      editInstalled: editInstalled,
      changeInstalled: changeInstalled,
      installed:
        editInstalled &&
        changeInstalled
    };
  });
}

function getDirectAuditTriggerStatus() {
  requireAdmin_();
  return getDirectAuditTriggerStatus_();
}

function installDirectEditAuditTriggers() {
  const actor = requireAdmin_();

  // Ensure the main workbook ID is remembered before external
  // spreadsheet triggers start firing.
  rememberDashboardSpreadsheet_();
  ensureAuditSheet_(
    getDashboardSpreadsheet_()
  );

  const currentTriggers =
    ScriptApp.getProjectTriggers();

  currentTriggers.forEach(function(trigger) {
    const handler = clean_(
      trigger.getHandlerFunction()
    );

    if (
      handler === CONFIG.DIRECT_AUDIT_EDIT_HANDLER ||
      handler === CONFIG.DIRECT_AUDIT_CHANGE_HANDLER
    ) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  const sources = getDirectAuditSources_();
  const failures = [];

  sources.forEach(function(source) {
    try {
      // Opening first gives a useful permission failure instead
      // of creating a trigger that cannot read the workbook.
      openSpreadsheetWithRetry_(
        source.spreadsheetId
      );

      ScriptApp
        .newTrigger(
          CONFIG.DIRECT_AUDIT_EDIT_HANDLER
        )
        .forSpreadsheet(
          source.spreadsheetId
        )
        .onEdit()
        .create();

      ScriptApp
        .newTrigger(
          CONFIG.DIRECT_AUDIT_CHANGE_HANDLER
        )
        .forSpreadsheet(
          source.spreadsheetId
        )
        .onChange()
        .create();

    } catch (error) {
      failures.push({
        spreadsheetId:
          source.spreadsheetId,
        label:
          source.label,
        error:
          error && error.message
            ? error.message
            : String(error)
      });
    }
  });

  const status =
    getDirectAuditTriggerStatus_();

  logActivities_([
    {
      action: 'Configured direct sheet auditing',
      entityType: 'System',
      field: 'Audit triggers',
      oldValue: '',
      newValue:
        status.filter(function(item) {
          return item.installed;
        }).length +
        ' of ' +
        status.length +
        ' workbooks',
      details:
        failures.length
          ? ('Failures: ' + JSON.stringify(failures))
          : 'All configured workbooks are monitored.',
      origin: 'System'
    }
  ], actor);

  return {
    success: failures.length === 0,
    status: status,
    failures: failures
  };
}

function removeDirectEditAuditTriggers() {
  const actor = requireAdmin_();
  let removed = 0;

  ScriptApp
    .getProjectTriggers()
    .forEach(function(trigger) {
      const handler = clean_(
        trigger.getHandlerFunction()
      );

      if (
        handler === CONFIG.DIRECT_AUDIT_EDIT_HANDLER ||
        handler === CONFIG.DIRECT_AUDIT_CHANGE_HANDLER
      ) {
        ScriptApp.deleteTrigger(trigger);
        removed++;
      }
    });

  logActivities_([
    {
      action: 'Removed direct sheet auditing',
      entityType: 'System',
      field: 'Audit triggers',
      oldValue: 'Installed',
      newValue: 'Removed',
      details: removed + ' trigger(s) removed.',
      origin: 'System'
    }
  ], actor);

  return {
    success: true,
    removed: removed,
    status: getDirectAuditTriggerStatus_()
  };
}

function getDirectEditActor_(eventObject) {
  let email = '';

  try {
    if (
      eventObject &&
      eventObject.user &&
      typeof eventObject.user.getEmail === 'function'
    ) {
      email = clean_(
        eventObject.user.getEmail()
      ).toLowerCase();
    }
  } catch (error) {
    email = '';
  }

  if (!email) {
    try {
      email = clean_(
        Session.getActiveUser().getEmail()
      ).toLowerCase();
    } catch (error) {
      email = '';
    }
  }

  const registry = getAccessRegistry_();
  const entry = email
    ? (registry[email] || null)
    : null;

  return {
    identified: !!email,
    email: email,
    name:
      entry && clean_(entry.name)
        ? clean_(entry.name)
        : (email || 'Unidentified sheet editor'),
    role:
      entry
        ? normalizeRole_(entry.role)
        : (email ? 'Unlisted' : 'Unknown'),
    active:
      entry
        ? entry.active !== false
        : true,
    canEdit: false,
    canAdmin: false
  };
}

function getRoutesForSource_(spreadsheetId, sheetName) {
  const matches = [];
  const id = clean_(spreadsheetId);
  const name = clean_(sheetName).toLowerCase();

  Object.keys(TASK_ROUTING).forEach(function(product) {
    TASK_ROUTING[product].forEach(function(route) {
      if (
        clean_(route.spreadsheetId) === id &&
        clean_(route.sourceSheet).toLowerCase() === name
      ) {
        matches.push({
          product: product,
          workstream: route.workstream,
          sourceSheet: route.sourceSheet,
          spreadsheetId: route.spreadsheetId
        });
      }
    });
  });

  return matches;
}

function getSourceEditContext_(sheet, rowNumber, routes) {
  const context = {
    entityType: 'Task',
    taskId: '',
    taskName: '',
    product: '',
    sourceSheet: sheet.getName(),
    weekStart: ''
  };

  const structure = getSourceSheetMap_(sheet);

  if (rowNumber <= structure.headerRow) {
    return context;
  }
  const row = sheet
    .getRange(
      rowNumber,
      1,
      1,
      structure.lastColumn
    )
    .getValues()[0];

  const idColumn = findColumn_(
    structure.map,
    [
      'Task ID',
      'ID',
      'Issue ID',
      'Issue Key',
      'Key'
    ]
  );

  const titleColumn = findColumn_(
    structure.map,
    [
      'Task Title',
      'Task Name',
      'Requirement / Task',
      'Task',
      'Title',
      'Deliverable',
      'Requirement'
    ]
  );

  const productColumn = findColumn_(
    structure.map,
    [
      'Product',
      'Source Product'
    ]
  );

  context.taskId = clean_(
    getCell_(row, idColumn)
  );

  context.taskName = clean_(
    getCell_(row, titleColumn)
  );

  const rowProduct = canonicalProduct_(
    getCell_(row, productColumn)
  );

  if (rowProduct) {
    context.product = rowProduct;
  }

  if (!context.product && routes.length === 1) {
    context.product = routes[0].product;
  }

  if (context.product && routes.length > 1) {
    const matched = routes.find(function(route) {
      return (
        canonicalProduct_(route.product) ===
        canonicalProduct_(context.product)
      );
    });

    if (matched) {
      context.product = matched.product;
    }
  }

  return context;
}

function getWeeklyEditContext_(sheet, rowNumber) {
  const context = {
    entityType: 'Weekly Plan',
    taskId: '',
    taskName: '',
    product: '',
    sourceSheet: '',
    weekStart: ''
  };

  if (rowNumber < 2) {
    return context;
  }

  const row = sheet
    .getRange(
      rowNumber,
      1,
      1,
      16
    )
    .getValues()[0];

  context.taskId = clean_(row[1]);
  context.taskName = clean_(row[2]);
  context.product = clean_(row[3]);
  context.weekStart = formatDate_(row[0]);
  context.sourceSheet = clean_(row[14]);

  return context;
}

function getDirectEditFieldLabel_(sheet, column) {
  if (!column || column < 1) {
    return '';
  }

  let headerRow = 1;

  try {
    headerRow = getSourceSheetMap_(sheet).headerRow || 1;
  } catch (error) {
    headerRow = 1;
  }

  const header = clean_(
    sheet.getRange(
      headerRow,
      column
    ).getDisplayValue()
  );

  return header || ('Column ' + column);
}

function summarizeEditedRangeValues_(values) {
  let text = '';

  try {
    text = JSON.stringify(values);
  } catch (error) {
    text = String(values);
  }

  if (text.length > 2000) {
    return text.substring(0, 2000) + '…';
  }

  return text;
}

function shouldAuditDirectEdit_(spreadsheetId, sheetName) {
  const dashboard = getDashboardSpreadsheet_();

  if (
    spreadsheetId === dashboard.getId()
  ) {
    return (
      sheetName === CONFIG.WEEKLY_SHEET
    );
  }

  return (
    getRoutesForSource_(
      spreadsheetId,
      sheetName
    ).length > 0
  );
}

function handleDirectSheetEdit(eventObject) {
  try {
    if (
      !eventObject ||
      !eventObject.range
    ) {
      return;
    }

    const range = eventObject.range;
    const sheet = range.getSheet();
    const spreadsheet = sheet.getParent();
    const spreadsheetId = spreadsheet.getId();
    const sheetName = sheet.getName();

    if (
      !shouldAuditDirectEdit_(
        spreadsheetId,
        sheetName
      )
    ) {
      return;
    }

    const actor =
      getDirectEditActor_(
        eventObject
      );

    const dashboardId =
      getDashboardSpreadsheet_().getId();

    const isWeekly =
      spreadsheetId === dashboardId &&
      sheetName === CONFIG.WEEKLY_SHEET;

    if (!isWeekly) {
      clearSourceTaskCache_();
    }

    const routes = isWeekly
      ? []
      : getRoutesForSource_(
          spreadsheetId,
          sheetName
        );

    const rowCount = range.getNumRows();
    const columnCount = range.getNumColumns();
    const startRow = range.getRow();
    const startColumn = range.getColumn();
    const entries = [];
    let auditHeaderRow = 1;

    if (!isWeekly) {
      try {
        auditHeaderRow = getSourceSheetMap_(sheet).headerRow || 1;
      } catch (error) {
        auditHeaderRow = 1;
      }
    }

    if (
      columnCount === 1 &&
      rowCount <= 100
    ) {
      const values =
        range.getDisplayValues();

      const field =
        getDirectEditFieldLabel_(
          sheet,
          startColumn
        );

      for (
        let rowOffset = 0;
        rowOffset < rowCount;
        rowOffset++
      ) {
        const rowNumber =
          startRow + rowOffset;

        const context = isWeekly
          ? getWeeklyEditContext_(
              sheet,
              rowNumber
            )
          : getSourceEditContext_(
              sheet,
              rowNumber,
              routes
            );

        const oldValue =
          rowCount === 1
            ? auditValue_(
                eventObject.oldValue
              )
            : '[Previous value unavailable for multi-cell edit]';

        const newValue =
          values[rowOffset][0];

        entries.push({
          action:
            rowNumber === auditHeaderRow
              ? 'Direct sheet header edit'
              : 'Direct sheet edit',
          entityType: context.entityType,
          taskId: context.taskId,
          taskName: context.taskName,
          product: context.product,
          sourceSheet:
            context.sourceSheet || sheetName,
          weekStart: context.weekStart,
          field: field,
          oldValue: oldValue,
          newValue: newValue,
          details:
            'Manual edit in ' +
            spreadsheet.getName() +
            ' / ' +
            sheetName +
            ' at ' +
            sheet.getRange(
              rowNumber,
              startColumn
            ).getA1Notation(),
          origin: 'Direct Sheet'
        });

        if (
          isWeekly &&
          context.weekStart
        ) {
          clearWeeklyPlanCache_(
            context.weekStart
          );
        }

        if (
          isWeekly &&
          field.toLowerCase() === 'week start' &&
          rowCount === 1 &&
          eventObject.oldValue
        ) {
          clearWeeklyPlanCache_(
            eventObject.oldValue
          );
        }
      }

    } else {
      const firstContext = isWeekly
        ? getWeeklyEditContext_(
            sheet,
            startRow
          )
        : getSourceEditContext_(
            sheet,
            startRow,
            routes
          );

      entries.push({
        action: 'Direct sheet range edit',
        entityType: firstContext.entityType,
        taskId: firstContext.taskId,
        taskName: firstContext.taskName,
        product: firstContext.product,
        sourceSheet:
          firstContext.sourceSheet || sheetName,
        weekStart: firstContext.weekStart,
        field: range.getA1Notation(),
        oldValue:
          '[Previous values unavailable for range edit]',
        newValue:
          summarizeEditedRangeValues_(
            range.getDisplayValues()
          ),
        details:
          'Manual ' +
          rowCount +
          '×' +
          columnCount +
          ' range edit in ' +
          spreadsheet.getName() +
          ' / ' +
          sheetName +
          '.',
        origin: 'Direct Sheet'
      });

      if (isWeekly) {
        clearAllWeeklyPlanCache_();
      }
    }

    if (entries.length) {
      logActivities_(
        entries,
        actor
      );
    }

  } catch (error) {
    console.error(
      'Direct edit audit failed:',
      error
    );
  }
}

function handleDirectSheetChange(eventObject) {
  try {
    if (
      !eventObject ||
      !eventObject.source
    ) {
      return;
    }

    const changeType = clean_(
      eventObject.changeType
    );

    // Cell edits are already captured by onEdit. Formatting is
    // intentionally ignored to keep the activity feed useful.
    if (
      !changeType ||
      changeType === 'EDIT' ||
      changeType === 'FORMAT'
    ) {
      return;
    }

    const spreadsheet =
      eventObject.source;

    const spreadsheetId =
      spreadsheet.getId();

    const monitoredIds =
      getDirectAuditSources_()
        .map(function(source) {
          return source.spreadsheetId;
        });

    if (
      monitoredIds.indexOf(
        spreadsheetId
      ) === -1
    ) {
      return;
    }

    let sheetName = '';

    try {
      const activeSheet =
        spreadsheet.getActiveSheet();

      sheetName = activeSheet
        ? activeSheet.getName()
        : '';
    } catch (error) {
      sheetName = '';
    }

    const actor =
      getDirectEditActor_(
        eventObject
      );

    logActivities_([
      {
        action: 'Direct sheet structural change',
        entityType: 'Sheet',
        sourceSheet: sheetName,
        field: 'Structure',
        oldValue: '',
        newValue: changeType,
        details:
          'Manual structural change in ' +
          spreadsheet.getName() +
          (sheetName ? (' / ' + sheetName) : '') +
          '.',
        origin: 'Direct Sheet'
      }
    ], actor);

    if (
      spreadsheetId ===
      getDashboardSpreadsheet_().getId()
    ) {
      clearAllWeeklyPlanCache_();
    } else {
      clearSourceTaskCache_();
    }

  } catch (error) {
    console.error(
      'Direct change audit failed:',
      error
    );
  }
}

/* ============================================================
   GOOGLE SHEETS DATA VALIDATION
   - Adds dropdown / type validation to editable raw source columns.
   - Does not rewrite existing cell values.
============================================================ */

function readActivityRows_(limit) {
  const ss = getDashboardSpreadsheet_();
  const appSheet = ensureAuditSheet_(ss);
  const sheetNames = [
    CONFIG.LEGACY_AUDIT_SHEET,
    CONFIG.AUDIT_SHEET
  ];
  const rows = [];

  limit = Math.max(1, Math.min(Number(limit) || 500, 2000));

  sheetNames.forEach(function(sheetName) {
    const sheet =
      sheetName === CONFIG.AUDIT_SHEET
        ? appSheet
        : ss.getSheetByName(sheetName);

    if (!sheet || sheet.getLastRow() < 2) return;

    const count =
      Math.min(limit, sheet.getLastRow() - 1);
    const start =
      sheet.getLastRow() - count + 1;

    sheet
      .getRange(start, 1, count, 17)
      .getValues()
      .forEach(function(row) {
        rows.push({
          timestamp:
            row[0] instanceof Date
              ? row[0].toISOString()
              : clean_(row[0]),
          userEmail: clean_(row[1]),
          userName: clean_(row[2]),
          role: clean_(row[3]),
          action: clean_(row[4]),
          entityType: clean_(row[5]),
          taskId: clean_(row[6]),
          taskName: clean_(row[7]),
          product: clean_(row[8]),
          sourceSheet: clean_(row[9]),
          weekStart: formatDate_(row[10]),
          field: clean_(row[11]),
          oldValue: clean_(row[12]),
          newValue: clean_(row[13]),
          details: clean_(row[14]),
          requestId: clean_(row[15]),
          origin: clean_(row[16]) || 'App'
        });
      });
  });

  rows.sort(function(a, b) {
    return b.timestamp.localeCompare(a.timestamp);
  });

  return rows.slice(0, limit);
}

function getActivityLog(limit) {
  requireAdmin_();
  return readActivityRows_(limit);
}

function getTaskHistory(taskId,limit) {
  const access=requireViewer_();
  const task=getTaskById(taskId);
  if(!task)return[];

  requireProductAccess_(access,task.product,false);

  const key=clean_(taskId).toLowerCase();
  return readActivityRows_(limit||500)
    .filter(function(item){
      return clean_(item.taskId).toLowerCase()===key;
    })
    .slice(0,100);
}
