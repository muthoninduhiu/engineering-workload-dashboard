function ensureWeeklySheet_(ss) {
  let sheet =
    ss.getSheetByName(
      CONFIG.WEEKLY_SHEET
    );

  const access =
    getAccessContext_();

  const headers = [
    'Week Start',
    'Task ID',
    'Weekly Commitment',
    'Product',
    'Workstream',
    'Owner',
    'Priority',
    'Weekly Status',
    'Due Date',
    'Planning Note',
    'Source Row',
    'Added At',
    'Commitment Level',
    'Points',
    'Source Sheet',
    'Estimated Hours'
  ];

  if (!sheet) {
    if (!access.canEdit) {
      throw new Error(
        'Weekly Workplan is not available for read-only users yet. Ask an editor to open the app once.'
      );
    }

    sheet =
      ss.insertSheet(
        CONFIG.WEEKLY_SHEET
      );

    sheet
      .getRange(
        1,
        1,
        1,
        headers.length
      )
      .setValues(
        [headers]
      )
      .setFontWeight(
        'bold'
      )
      .setBackground(
        '#f1f5f9'
      );
  } else if (access.canEdit) {
    // Preserve all existing Weekly Workplan column names.
    // Only populate the new Estimated Hours header when column P is blank.
    const estimatedHoursHeader =
      clean_(
        sheet
          .getRange(
            1,
            16
          )
          .getDisplayValue()
      );

    if (!estimatedHoursHeader) {
      sheet
        .getRange(
          1,
          16
        )
        .setValue(
          'Estimated Hours'
        );
    }
  }

  if (
    access.canEdit &&
    sheet.getFrozenRows() !== 1
  ) {
    sheet.setFrozenRows(1);
  }

  return sheet;
}

/* ============================================================
   READ WEEKLY PLAN FOR SELECTED WEEK
============================================================ */

function serializeWeeklyPlanRow_(row, index) {
  const taskId = clean_(row[1]);

  if (!taskId) return null;

  let points = '';

  if (
    row[13] !== '' &&
    !isNaN(
      Number(row[13])
    )
  ) {
    points = Number(row[13]);
  }

  return {
    weeklyRow: index + 2,
    weekStart: formatDate_(row[0]),
    taskId: taskId,
    title: clean_(row[2]),
    product: resolveTaskProduct_('',  clean_(row[3]), clean_(row[1])),
    workstream: clean_(row[4]),
    owner: normalizeOwner_(row[5]),
    priority: clean_(row[6]) || 'P2',
    weeklyStatus: clean_(row[7]) || 'Todo',
    dueDate: formatDate_(row[8]),
    planningNote: clean_(row[9]),
    sourceRow: Number(row[10]) || '',
    addedAt:
      row[11] instanceof Date
        ? row[11].toISOString()
        : clean_(row[11]),
    commitmentLevel: normalizeCommitmentLevel_(row[12]),
    points: points,
    sourceSheet: clean_(row[14]),
    estimatedHours:
      row[15] !== '' &&
      row[15] !== null &&
      row[15] !== undefined &&
      !isNaN(Number(row[15]))
        ? Number(row[15])
        : ''
  };
}

function getWeeklyCacheVersion_() {
  const properties =
    PropertiesService.getScriptProperties();

  let version = clean_(
    properties.getProperty(
      CONFIG.WEEK_CACHE_VERSION_PROPERTY_KEY
    )
  );

  if (!version) {
    version = '1';
    properties.setProperty(
      CONFIG.WEEK_CACHE_VERSION_PROPERTY_KEY,
      version
    );
  }

  return version;
}

function weeklyPlanCacheKey_(weekStart) {
  return (
    'EW_WEEK_' +
    getWeeklyCacheVersion_() +
    '_' +
    getWeekFromStart_(weekStart).start
  );
}

function clearWeeklyPlanCache_(weekStart) {
  if (!weekStart) return;

  try {
    CacheService
      .getScriptCache()
      .remove(
        weeklyPlanCacheKey_(
          weekStart
        )
      );
  } catch (error) {
    console.warn(
      'Could not clear week cache.',
      error
    );
  }
}

function clearAllWeeklyPlanCache_() {
  const properties =
    PropertiesService.getScriptProperties();

  const current =
    Number(
      getWeeklyCacheVersion_()
    ) || 1;

  properties.setProperty(
    CONFIG.WEEK_CACHE_VERSION_PROPERTY_KEY,
    String(current + 1)
  );
}

function getWeeklyPlanRowsForWeeks_(ss, selectedWeekStarts) {
  const sheet = ensureWeeklySheet_(ss);

  selectedWeekStarts = Array.isArray(selectedWeekStarts)
    ? selectedWeekStarts
    : [selectedWeekStarts];

  const starts = [];

  selectedWeekStarts.forEach(function(value) {
    const start = getWeekFromStart_(value).start;

    if (starts.indexOf(start) === -1) {
      starts.push(start);
    }
  });

  const output = {};
  const missing = [];
  const cache = CacheService.getScriptCache();

  starts.forEach(function(start) {
    const cached = cache.get(
      weeklyPlanCacheKey_(start)
    );

    if (cached !== null) {
      try {
        output[start] = JSON.parse(cached) || [];
        return;
      } catch (error) {
        // Fall through and rebuild this week from the sheet.
      }
    }

    output[start] = [];
    missing.push(start);
  });

  if (!missing.length) {
    return output;
  }

  if (sheet.getLastRow() >= 2) {
    const wanted = new Set(missing);

    const rows = sheet
      .getRange(
        2,
        1,
        sheet.getLastRow() - 1,
        16
      )
      .getValues();

    rows.forEach(function(row, index) {
      const item = serializeWeeklyPlanRow_(
        row,
        index
      );

      if (
        !item ||
        !wanted.has(item.weekStart)
      ) {
        return;
      }

      output[item.weekStart].push(item);
    });
  }

  missing.forEach(function(start) {
    try {
      cache.put(
        weeklyPlanCacheKey_(start),
        JSON.stringify(output[start] || []),
        CONFIG.WEEK_CACHE_SECONDS
      );
    } catch (error) {
      // Cache limits should never break the dashboard.
      console.warn(
        'Could not cache week ' + start,
        error
      );
    }
  });

  return output;
}

function getWeeklyPlanRows_(ss, selectedWeekStart) {
  const week = getWeekFromStart_(selectedWeekStart);

  const grouped = getWeeklyPlanRowsForWeeks_(
    ss,
    [week.start]
  );

  return grouped[week.start] || [];
}

function getWeeklyPlan_(ss, tasks, selectedWeekStart) {
  tasks = tasks || getSourceTasks_();
  const idMap = new Map();

  tasks.forEach(function(task) {
    idMap.set(clean_(task.taskId).toLowerCase(), task);
  });

  return getWeeklyPlanRows_(ss, selectedWeekStart).map(function(item) {
    const masterTask =
      idMap.get(clean_(item.taskId).toLowerCase()) ||
      null;

  return Object.assign({}, item, {
    product: masterTask
      ? resolveTaskProduct_(
          masterTask.product,
          item.product,
          item.taskId
        )
      : item.product,

      points:
        masterTask && masterTask.points !== ''
          ? masterTask.points
          : item.points,

      sourceSheet:
        masterTask
          ? masterTask.sourceSheet
          : item.sourceSheet,

      registerStatus:
        masterTask
          ? masterTask.status
          : item.weeklyStatus,

      registerStatusGroup:
        masterTask
          ? masterTask.statusGroup
          : normalizeStatus_(item.weeklyStatus),

      registerBlocker:
        masterTask
          ? masterTask.dependency
          : '',

      registerPriority:
        masterTask
          ? masterTask.normalizedPriority
          : normalizePriority_(item.priority)
    });
  });
}


/* ============================================================
   ADD EXISTING TASK(S) TO SELECTED WEEK
============================================================ */

function addTasksToWeekCore_(taskIds, commitmentLevel, selectedWeekStart, estimatedHours) {
  const actor = requireEditor_();

  if (
    !Array.isArray(taskIds) ||
    !taskIds.length
  ) {
    throw new Error(
      'No tasks were selected.'
    );
  }

  commitmentLevel = canonicalCommitmentInput_(commitmentLevel);

  const lock =
    LockService
      .getDocumentLock();

  if (
    !lock.tryLock(10000)
  ) {
    throw new Error(
      'The weekly plan is currently being updated.'
    );
  }

  try {
    const ss =
      getDashboardSpreadsheet_();

    const tasks =
      getSourceTasks_();

    const weeklySheet =
      ensureWeeklySheet_(ss);

    const week =
      getWeekFromStart_(
        selectedWeekStart
      );

    const plan =
      getWeeklyPlan_(
        ss,
        tasks,
        week.start
      );

    const existingIds =
      new Set(
        plan.map(function(item) {
          return clean_(
            item.taskId
          ).toLowerCase();
        })
      );

    const appendRows = [];
    const skipped = [];
    const auditEntries = [];

    function estimatedHoursForTask_(task) {
      let raw = estimatedHours;

      if (
        raw &&
        typeof raw === 'object' &&
        !Array.isArray(raw)
      ) {
        raw =
          raw[task.taskId];
      }

      return statusUsesEstimatedHours_(task.status)
        ? validateEstimatedHoursInput_(raw, true)
        : '';
    }

    taskIds
      .map(function(taskId) { return clean_(taskId).toLowerCase(); })
      .forEach(function(taskId) {
        const task =
          tasks.find(function(item) {
            return clean_(item.taskId).toLowerCase() === taskId;
          });

        if (!task) {
          skipped.push(
            taskId
          );
          return;
        }

        if (task.closed) {
          skipped.push(
            task.taskId + ' (done / closed)'
          );
          return;
        }

        const resolvedEstimatedHours =
          estimatedHoursForTask_(task);

        const taskIdKey =
          clean_(
            task.taskId
          ).toLowerCase();

        if (
          existingIds.has(taskIdKey)
        ) {
          skipped.push(
            task.taskId
          );
          return;
        }

        appendRows.push([
          week.start,
          task.taskId,
          task.title,
          task.product,
          task.workstream,
          task.owner,
          task.normalizedPriority,
          task.status,
          task.dueDate,
          '',
          '',
          new Date(),
          commitmentLevel,
          task.points,
          task.sourceSheet,
          resolvedEstimatedHours
        ]);

        auditEntries.push({
          action: 'Added to weekly plan',
          entityType: 'Weekly Plan',
          taskId: task.taskId,
          taskName: task.title,
          product: task.product,
          sourceSheet: task.sourceSheet,
          weekStart: week.start,
          field: 'Commitment',
          oldValue: '',
          newValue:
            commitmentLevel +
            ' · ' +
            resolvedEstimatedHours +
            'h'
        });

        existingIds.add(
          taskIdKey
        );
      });

    if (appendRows.length) {
      weeklySheet
        .getRange(
          weeklySheet.getLastRow() + 1,
          1,
          appendRows.length,
          16
        )
        .setValues(
          appendRows
        );

      clearWeeklyPlanCache_(
        week.start
      );

      logActivities_(
        auditEntries,
        actor
      );
    }

    SpreadsheetApp.flush();

    return {
      success: true,
      added: appendRows.length,
      addedCount: appendRows.length,
      requested: taskIds.length,
      skipped: skipped,
      weekStart: week.start
    };
  } finally {
    lock.releaseLock();
  }
}

/* ============================================================
   ADD NEWLY CREATED TASK DIRECTLY TO SELECTED WEEK
============================================================ */

function addCreatedTaskToWeek_(
  ss,
  task,
  route,
  commitmentLevel,
  selectedWeekStart,
  estimatedHours
) {
  if (isClosedStatus_(task.status)) {
    throw new Error('Done or closed tasks cannot be added to a weekly plan.');
  }

  const sheet = ensureWeeklySheet_(ss);
  const week = getWeekFromStart_(selectedWeekStart);
  estimatedHours = statusUsesEstimatedHours_(task.status)
    ? validateEstimatedHoursInput_(estimatedHours, true)
    : '';
  const tasks = getSourceTasks_();
  const plan = getWeeklyPlan_(ss, tasks, week.start);

  if (
    plan.some(function(item) {
      return (
        clean_(item.taskId).toLowerCase() ===
        clean_(task.taskId).toLowerCase()
      );
    })
  ) {
    return false;
  }

  sheet.appendRow([
    week.start,
    task.taskId,
    task.title,
    route.product,
    route.workstream,
    task.owner,
    task.priority,
    task.status,
    task.dueDate,
    '',
    '',
    new Date(),
    canonicalCommitmentInput_(commitmentLevel),
    task.points,
    route.sourceSheet,
    estimatedHours
  ]);

  clearWeeklyPlanCache_(
    week.start
  );

  return true;
}

/* ============================================================
   UPDATE WEEKLY
============================================================ */

function updateWeeklyTaskFieldCore_(weeklyRow, field, value) {
  const actor = requireEditor_();
  const ss = getDashboardSpreadsheet_();
  const sheet = ensureWeeklySheet_(ss);

  weeklyRow = Number(weeklyRow);

  if (!weeklyRow || weeklyRow < 2) {
    throw new Error(
      'Invalid weekly task row.'
    );
  }

  const columns = {
    title: 3,
    owner: 6,
    priority: 7,
    weeklyStatus: 8,
    dueDate: 9,
    planningNote: 10,
    commitmentLevel: 13,
    points: 14,
    estimatedHours: 16
  };

  const column =
    columns[field];

  if (!column) {
    throw new Error(
      'Unsupported weekly field: ' +
      field
    );
  }

  const row = sheet
    .getRange(
      weeklyRow,
      1,
      1,
      16
    )
    .getValues()[0];

  const oldValue =
    row[column - 1];

  if (field === 'owner') {
    value = normalizeOwner_(value);
  } else if (field === 'title') {
    value = clean_(value);
    if (!value) throw new Error('Task title is required.');
  } else if (field === 'priority') {
    value = canonicalPriorityInput_(value);
  } else if (field === 'weeklyStatus') {
    value = canonicalStatusInput_(value);
  } else if (field === 'dueDate') {
    value = validateIsoDateInput_(value, 'Due date');
  } else if (field === 'commitmentLevel') {
    value = canonicalCommitmentInput_(value);
  } else if (field === 'points') {
    value = validatePointsInput_(value);
  } else if (field === 'estimatedHours') {
    const currentStatus = clean_(row[7]);
    value = statusUsesEstimatedHours_(currentStatus)
      ? validateEstimatedHoursInput_(value, true)
      : '';
  }

  sheet
    .getRange(
      weeklyRow,
      column
    )
    .setValue(
      value
    );

  if (
    field === 'weeklyStatus' &&
    !statusUsesEstimatedHours_(value)
  ) {
    sheet
      .getRange(
        weeklyRow,
        16
      )
      .clearContent();
  }

  clearWeeklyPlanCache_(
    formatDate_(row[0])
  );

  SpreadsheetApp.flush();

  if (
    auditValue_(oldValue) !==
    auditValue_(value)
  ) {
    logActivities_([
      {
        action: 'Updated weekly task',
        entityType: 'Weekly Plan',
        taskId: clean_(row[1]),
        taskName: clean_(row[2]),
        product: clean_(row[3]),
        sourceSheet: clean_(row[14]),
        weekStart: formatDate_(row[0]),
        field: field,
        oldValue: oldValue,
        newValue: value
      }
    ], actor);
  }

  return {
    success: true
  };
}

function updateWeeklyTaskFieldsCore_(weeklyRow, payload) {
  const actor = requireEditor_();
  const ss = getDashboardSpreadsheet_();
  const sheet = ensureWeeklySheet_(ss);

  weeklyRow = Number(weeklyRow);

  if (!weeklyRow || weeklyRow < 2) {
    throw new Error(
      'Invalid weekly task row.'
    );
  }

  payload = payload || {};

  const oldRow = sheet
    .getRange(
      weeklyRow,
      1,
      1,
      16
    )
    .getValues()[0];

  const title = clean_(payload.title);
  if (!title) throw new Error('Task title is required.');

  const points = validatePointsInput_(payload.points);
  const weeklyStatus =
    canonicalStatusInput_(payload.weeklyStatus || 'Todo');
  const estimatedHours =
    statusUsesEstimatedHours_(weeklyStatus)
      ? validateEstimatedHoursInput_(payload.estimatedHours, true)
      : '';

  const updates = {
    3: title,
    6: normalizeOwner_(payload.owner),
    7: canonicalPriorityInput_(payload.priority || 'P2'),
    8: weeklyStatus,
    9: validateIsoDateInput_(payload.dueDate, 'Due date'),
    10: clean_(payload.planningNote),
    13: canonicalCommitmentInput_(payload.commitmentLevel),
    14: points,
    16: estimatedHours
  };

  const fieldNames = {
    3: 'title',
    6: 'owner',
    7: 'priority',
    8: 'weeklyStatus',
    9: 'dueDate',
    10: 'planningNote',
    13: 'commitmentLevel',
    14: 'points',
    16: 'estimatedHours'
  };

  const auditEntries = [];

  Object.keys(updates)
    .forEach(function(column) {
      const numericColumn =
        Number(column);

      const oldValue =
        oldRow[
          numericColumn - 1
        ];

      const newValue =
        updates[column];

      sheet
        .getRange(
          weeklyRow,
          numericColumn
        )
        .setValue(
          newValue
        );

      if (
        auditValue_(oldValue) !==
        auditValue_(newValue)
      ) {
        auditEntries.push({
          action: 'Updated weekly task',
          entityType: 'Weekly Plan',
          taskId: clean_(oldRow[1]),
          taskName: clean_(oldRow[2]),
          product: clean_(oldRow[3]),
          sourceSheet: clean_(oldRow[14]),
          weekStart: formatDate_(oldRow[0]),
          field: fieldNames[numericColumn] || ('Column ' + numericColumn),
          oldValue: oldValue,
          newValue: newValue
        });
      }
    });

  clearWeeklyPlanCache_(
    formatDate_(oldRow[0])
  );

  SpreadsheetApp.flush();

  if (auditEntries.length) {
    logActivities_(
      auditEntries,
      actor
    );
  }

  return {
    success: true
  };
}

/* ============================================================
   REMOVE WEEKLY
============================================================ */

function removeWeeklyTaskCore_(weeklyRow) {
  const actor = requireEditor_();
  const ss = getDashboardSpreadsheet_();
  const sheet = ensureWeeklySheet_(ss);

  weeklyRow = Number(weeklyRow);

  if (
    !weeklyRow ||
    weeklyRow < 2 ||
    weeklyRow > sheet.getLastRow()
  ) {
    throw new Error(
      'Invalid weekly task row.'
    );
  }

  const row = sheet
    .getRange(
      weeklyRow,
      1,
      1,
      15
    )
    .getValues()[0];

  sheet.deleteRow(
    weeklyRow
  );

  clearWeeklyPlanCache_(
    formatDate_(row[0])
  );

  SpreadsheetApp.flush();

  logActivities_([
    {
      action: 'Removed from weekly plan',
      entityType: 'Weekly Plan',
      taskId: clean_(row[1]),
      taskName: clean_(row[2]),
      product: clean_(row[3]),
      sourceSheet: clean_(row[14]),
      weekStart: formatDate_(row[0]),
      field: 'Commitment',
      oldValue: normalizeCommitmentLevel_(row[12]),
      newValue: 'Removed'
    }
  ], actor);

  return {
    success: true
  };
}

/* ============================================================
   CREATE TASK IN REAL SOURCE SHEET
============================================================ */

function addTasksToWeek(taskIds,commitmentLevel,selectedWeekStart,estimatedHours) {
  const access=requireEditor_();
  const tasks=getSourceTasks_();
  (taskIds||[]).forEach(function(taskId){const task=tasks.find(function(t){return clean_(t.taskId).toLowerCase()===clean_(taskId).toLowerCase();});if(!task)throw new Error('Task '+taskId+' was not found in its configured source sheet.');requireProductAccess_(access,task.product,true);});
  return addTasksToWeekCore_(taskIds,commitmentLevel,selectedWeekStart,estimatedHours);
}

function weeklyRowProduct_(weeklyRow) {
  const sheet = ensureWeeklySheet_(getDashboardSpreadsheet_());

  weeklyRow = Number(weeklyRow);

  if (
    !weeklyRow ||
    weeklyRow < 2 ||
    weeklyRow > sheet.getLastRow()
  ) {
    throw new Error('Invalid weekly task row.');
  }

  const row = sheet.getRange(weeklyRow, 2, 1, 3).getValues()[0];

  return resolveTaskProduct_(
    '',
    clean_(row[2]),
    clean_(row[0])
  );
}
function updateWeeklyTaskField(weeklyRow,field,value) {
  const access=requireEditor_();
  requireProductAccess_(access,weeklyRowProduct_(weeklyRow),true);
  return updateWeeklyTaskFieldCore_(weeklyRow,field,value);
}
function updateWeeklyTaskFields(weeklyRow,payload) {
  const access=requireEditor_();
  requireProductAccess_(access,weeklyRowProduct_(weeklyRow),true);
  return updateWeeklyTaskFieldsCore_(weeklyRow,payload);
}
function removeWeeklyTask(weeklyRow) {
  const access=requireEditor_();
  requireProductAccess_(access,weeklyRowProduct_(weeklyRow),true);
  return removeWeeklyTaskCore_(weeklyRow);
}
function addTaskToWeek(
  taskId,
  commitmentLevel,
  selectedWeekStart,
  estimatedHours
) {
  return addTasksToWeek(
    [taskId],
    commitmentLevel,
    selectedWeekStart,
    estimatedHours
  );
}
function runClearWeeklyPlanCache() {
  clearAllWeeklyPlanCache_();
  Logger.log('Weekly Plan cache cleared successfully.');
}