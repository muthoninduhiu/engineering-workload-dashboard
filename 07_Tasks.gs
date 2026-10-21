function getSourceTasks_() {
  return getConfiguredSourceTasks_();
}

/* ============================================================
   GENERIC SOURCE SHEET HELPERS
============================================================ */

function getSourceSchemaProfile_(sheet) {
  const spreadsheetId = clean_(sheet.getParent().getId());
  const key =
    spreadsheetId +
    '::' +
    clean_(sheet.getName()).toLowerCase();

  const override =
    SOURCE_SCHEMA_OVERRIDES[key] || {};

  return {
    headerRow:
      Number(override.headerRow) || null,

    idAliases:
      override.idAliases || [
        'Task ID', 'ID', 'Identifier', 'Issue Identifier', 'Linear ID',
        'Issue ID', 'Issue Key', 'Key', 'Task #',
        'Ticket ID', 'Requirement ID', 'Reference ID'
      ],

    titleAliases:
      override.titleAliases || [
        'Task Title', 'Task Name', 'Requirement / Task',
        'Feature / Issue', 'Repository',
        'Task', 'Title', 'Summary', 'Name', 'Deliverable', 'Requirement'
      ],

    descriptionAliases:
      override.descriptionAliases || [
        'Description',
        'Requirement / Description',
        'Purpose / Description',
        'Requirement Description',
        'Details'
      ],

    ownerAliases:
      override.ownerAliases || [
        'Owner', 'Primary Owner', 'Raw Owner',
        'Normalized Owner', 'Assignee'
      ],

    statusAliases:
      override.statusAliases || [
        'Status', 'Raw Status', 'Task Status', 'Normalized Status'
      ],

    priorityAliases:
      override.priorityAliases || ['Priority'],

    workstreamAliases:
      override.workstreamAliases || [
        'Workstream', 'Category', 'Track', 'Module'
      ],

    productAliases:
      override.productAliases || [
        'Product', 'Source Product'
      ],

    dependencyAliases:
      override.dependencyAliases || [
        'Dependency', 'Dependencies', 'Dependency / Related ID',
        'Blocker', 'Blocker / Dependency'
      ],

    dueDateAliases:
      override.dueDateAliases || [
        'Due Date', 'Due'
      ],

    pointsAliases:
      override.pointsAliases || [
        'Points', 'Story Points'
      ]
  };
}

function sourceSchemaCandidate_(row, rowNumber, profile) {
  const map = buildHeaderMap_(row);

  const idColumn =
    findColumn_(
      map,
      profile.idAliases
    );

  const titleColumn =
    findColumn_(
      map,
      profile.titleAliases
    );

  const descriptionColumn =
    findColumn_(
      map,
      profile.descriptionAliases
    );

  const ownerColumn =
    findColumn_(
      map,
      profile.ownerAliases
    );

  const statusColumn =
    findColumn_(
      map,
      profile.statusAliases
    );

  let score = 0;

  if (idColumn !== -1) score += 20;
  if (titleColumn !== -1) score += 20;
  if (idColumn !== -1 && titleColumn !== -1) score += 50;

  [
    profile.statusAliases,
    profile.ownerAliases,
    profile.priorityAliases,
    profile.workstreamAliases,
    profile.productAliases,
    profile.dependencyAliases,
    profile.dueDateAliases,
    profile.descriptionAliases
  ].forEach(function(group) {
    if (findColumn_(map, group) !== -1) score += 2;
  });

  return {
    rowNumber: rowNumber,
    headers: row,
    map: map,
    idColumn: idColumn,
    titleColumn: titleColumn,
    descriptionColumn: descriptionColumn,
    ownerColumn: ownerColumn,
    statusColumn: statusColumn,
    score: score
  };
}

function getSourceSheetMap_(sheet) {
  const lastColumn = sheet.getLastColumn();

  if (lastColumn < 1) {
    throw new Error(
      '"' + sheet.getName() +
      '" does not contain any columns.'
    );
  }

  const profile =
    getSourceSchemaProfile_(sheet);

  let best = null;

  if (profile.headerRow) {
    if (profile.headerRow > sheet.getMaxRows()) {
      throw new Error(
        'Configured header row ' +
        profile.headerRow +
        ' is outside "' +
        sheet.getName() +
        '".'
      );
    }

    const row =
      sheet
        .getRange(
          profile.headerRow,
          1,
          1,
          lastColumn
        )
        .getDisplayValues()[0];

    best =
      sourceSchemaCandidate_(
        row,
        profile.headerRow,
        profile
      );

    if (
      best.idColumn === -1 ||
      best.titleColumn === -1
    ) {
      throw new Error(
        'Configured header row ' +
        profile.headerRow +
        ' in "' +
        sheet.getName() +
        '" does not contain the expected ID/title aliases. ' +
        'Headers seen: ' +
        row.map(clean_).filter(Boolean).join(' | ')
      );
    }
  } else {
    const lastRow =
      Math.max(
        1,
        sheet.getLastRow()
      );

    const scanRows =
      Math.min(
        lastRow,
      100
      );

    const rows =
      sheet
        .getRange(
          1,
          1,
          scanRows,
          lastColumn
        )
        .getDisplayValues();

    rows.forEach(function(row, index) {
      const candidate =
        sourceSchemaCandidate_(
          row,
          index + 1,
          profile
        );

      if (
        !best ||
        candidate.score > best.score ||
        (
          candidate.score === best.score &&
          candidate.rowNumber < best.rowNumber
        )
      ) {
        best = candidate;
      }
    });
  }

  if (
    !best ||
    best.idColumn === -1 ||
    best.titleColumn === -1
  ) {
    throw new Error(
      'Could not detect the task header row in "' +
      sheet.getName() +
      '". Expected a recognized task ID column and task title column. ' +
      'Best candidate row ' +
      (best ? best.rowNumber : 'unknown') +
      ' headers: ' +
      (best ? best.headers.map(clean_).filter(Boolean).join(' | ') : 'none')
    );
  }

  return {
    headerRow: best.rowNumber,
    dataStartRow: best.rowNumber + 1,
    lastColumn: lastColumn,
    headers: best.headers,
    map: best.map,
    idColumn: best.idColumn,
    titleColumn: best.titleColumn,
    descriptionColumn: best.descriptionColumn,
    ownerColumn: best.ownerColumn,
    statusColumn: best.statusColumn,
    aliases: profile
  };
}

function setArrayValue_(row, map, aliases, value, allMatches) {
  let written = false;

  for (
    let i = 0;
    i < aliases.length;
    i++
  ) {
    const key =
      aliases[i]
        .toLowerCase();

    if (
      !Object.prototype
        .hasOwnProperty
        .call(
          map,
          key
        )
    ) {
      continue;
    }

    row[
      map[key]
    ] = value;

    written = true;

    if (!allMatches) {
      break;
    }
  }

  return written;
}

function findSourceTaskRow_(sheet, taskId) {
  const structure = getSourceSheetMap_(sheet);
  const idColumn =
    findColumn_(
      structure.map,
      structure.aliases.idAliases
    );

  const startRow = structure.dataStartRow;
  const lastRow = sheet.getLastRow();

  if (idColumn === -1 || lastRow < startRow) return -1;

  const values = sheet
    .getRange(startRow, idColumn + 1, lastRow - startRow + 1, 1)
    .getValues();

  const target = clean_(taskId).toLowerCase();

  for (let i = 0; i < values.length; i++) {
    if (clean_(values[i][0]).toLowerCase() === target) {
      return i + startRow;
    }
  }

  return -1;
}

function getTaskIdsFromSourceSheet_(sheet) {
  const structure = getSourceSheetMap_(sheet);
  const idColumn =
    findColumn_(
      structure.map,
      structure.aliases.idAliases
    );

  const startRow = structure.dataStartRow;
  const lastRow = sheet.getLastRow();

  if (idColumn === -1 || lastRow < startRow) return [];

  return sheet
    .getRange(startRow, idColumn + 1, lastRow - startRow + 1, 1)
    .getValues()
    .flat()
    .map(clean_)
    .filter(Boolean);
}

/* ============================================================
   APPEND TO REAL SOURCE SHEET
============================================================ */

function setSourceCellValue_(sheet, rowNumber, map, aliases, value, allMatches, overwriteFormula) {
  let written = false;

  for (let i = 0; i < aliases.length; i++) {
    const key = aliases[i].toLowerCase();
    if (!Object.prototype.hasOwnProperty.call(map, key)) continue;

    const cell = sheet.getRange(rowNumber, map[key] + 1);
    if (!overwriteFormula && cell.getFormula()) continue;

    cell.setValue(value);
    written = true;
    if (!allMatches) break;
  }

  return written;
}

function getLastSourceTaskRow_(sheet, structure) {
  const candidates = [
    findColumn_(
      structure.map,
      structure.aliases.idAliases
    ),
    findColumn_(
      structure.map,
      structure.aliases.titleAliases
    )
  ].filter(function(index) { return index !== -1; });

  const sheetLast = Math.max(structure.headerRow, sheet.getLastRow());
  const startRow = structure.dataStartRow;
  let last = structure.headerRow;

  candidates.forEach(function(index) {
    if (sheetLast < startRow) return;
    const values = sheet
      .getRange(startRow, index + 1, sheetLast - startRow + 1, 1)
      .getDisplayValues();

    for (let i = values.length - 1; i >= 0; i--) {
      if (clean_(values[i][0])) {
        last = Math.max(last, i + startRow);
        break;
      }
    }
  });

  return last;
}

function appendTaskToSourceSheet_(ss, route, task) {
  const sheet = getTargetSheet_(route);

  const existingRow = findSourceTaskRow_(sheet, task.taskId);
  if (existingRow !== -1) {
    throw new Error('Task ID "' + task.taskId + '" already exists in "' + route.sourceSheet + '".');
  }

  const structure = getSourceSheetMap_(sheet);
  const idColumn = structure.idColumn;
  const titleColumn = structure.titleColumn;

  if (idColumn === -1 || titleColumn === -1) {
    throw new Error(
      '"' + route.sourceSheet + '" header detection failed unexpectedly. ' +
      'Detected header row: ' + structure.headerRow + '. Headers: ' + structure.headers.join(' | ')
    );
  }

  const lastDataRow = getLastSourceTaskRow_(sheet, structure);
  let newRow = lastDataRow + 1;

  // Never overwrite a non-empty footer/summary row.
  if (newRow <= sheet.getLastRow()) {
    const occupied = sheet.getRange(newRow, 1, 1, structure.lastColumn)
      .getDisplayValues()[0]
      .some(function(value) { return clean_(value) !== ''; });
    if (occupied) newRow = sheet.getLastRow() + 1;
  }

  if (newRow > sheet.getMaxRows()) {
    sheet.insertRowsAfter(sheet.getMaxRows(), newRow - sheet.getMaxRows());
  }

  if (lastDataRow > structure.headerRow) {
    const sourceRange = sheet.getRange(lastDataRow, 1, 1, structure.lastColumn);
    const targetRange = sheet.getRange(newRow, 1, 1, structure.lastColumn);
    sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
  }

  // Write only actual task-input fields. Formula-derived columns stay intact.
  setSourceCellValue_(sheet,newRow,structure.map,structure.aliases.idAliases,task.taskId,false,true);
  setSourceCellValue_(sheet,newRow,structure.map,structure.aliases.titleAliases,task.title,false,true);
  setSourceCellValue_(sheet,newRow,structure.map,structure.aliases.productAliases,route.product,true,false);
  setSourceCellValue_(sheet,newRow,structure.map,structure.aliases.workstreamAliases,route.workstream,true,false);
  setSourceCellValue_(sheet,newRow,structure.map,structure.aliases.ownerAliases,task.owner,true,false);
  setSourceCellValue_(sheet,newRow,structure.map,structure.aliases.priorityAliases,task.priority,true,false);
  setSourceCellValue_(sheet,newRow,structure.map,structure.aliases.statusAliases,task.status,true,false);
  setSourceCellValue_(sheet,newRow,structure.map,structure.aliases.dependencyAliases,task.dependency,true,false);
  setSourceCellValue_(sheet,newRow,structure.map,structure.aliases.dueDateAliases,task.dueDate,true,false);
  setSourceCellValue_(sheet,newRow,structure.map,structure.aliases.pointsAliases,task.points,true,false);

  SpreadsheetApp.flush();
  return newRow;
}

/* ============================================================
   UPDATE EXISTING SOURCE TASK
============================================================ */

function updateSourceTaskFieldById_(
  product,
  sourceSheetName,
  taskId,
  field,
  value
) {
  if (!product || !sourceSheetName) return false;

  let route;

  try {
    route = resolveRoute_(product, '', sourceSheetName);
  } catch (error) {
    console.warn(
      'Could not resolve source route:',
      product,
      sourceSheetName,
      error
    );
    return false;
  }

  const sheet = getTargetSheet_(route);
  const sourceRow = findSourceTaskRow_(sheet, taskId);

  if (sourceRow === -1) return false;

  const structure = getSourceSheetMap_(sheet);
  let aliases = [];

  switch (field) {
    case 'owner':
      aliases = structure.aliases.ownerAliases;
      value = normalizeOwner_(value);
      break;

    case 'priority':
      aliases = structure.aliases.priorityAliases;
      break;

    case 'status':
      aliases = structure.aliases.statusAliases;
      break;

    case 'title':
      aliases = structure.aliases.titleAliases;
      break;

    case 'dependency':
      aliases = structure.aliases.dependencyAliases;
      break;

    case 'dueDate':
      aliases = structure.aliases.dueDateAliases;
      break;

    case 'points':
      aliases = structure.aliases.pointsAliases;
      break;

    default:
      return false;
  }

  let updated = false;

  aliases.forEach(function(alias) {
    const key = normalizeHeaderKey_(alias);

    if (Object.prototype.hasOwnProperty.call(structure.map, key)) {
      sheet
        .getRange(sourceRow, structure.map[key] + 1)
        .setValue(value);
      updated = true;
    }
  });

  SpreadsheetApp.flush();
  if (updated) clearSourceTaskCache_();
  return updated;
}

// Resolves a task source exactly once, then verifies every requested column
// before changing anything. This prevents a drawer save from succeeding for
// early fields and failing halfway through on a later field.
function updateSourceTaskFieldsById_(product, sourceSheetName, taskId, fields) {
  if (!product || !sourceSheetName || !fields || !Object.keys(fields).length) {
    return false;
  }

  let route;
  try {
    route = resolveRoute_(product, '', sourceSheetName);
  } catch (error) {
    throw new Error('Could not resolve the configured source for ' + taskId + '.');
  }

  const sheet = getTargetSheet_(route);
  const sourceRow = findSourceTaskRow_(sheet, taskId);
  if (sourceRow === -1) {
    throw new Error(
      'Could not find ' + taskId + ' in its configured source sheet "' +
      sourceSheetName + '". Refresh the workspace and verify the task ID exists there.'
    );
  }

  const structure = getSourceSheetMap_(sheet);
  const aliasByField = {
    owner: structure.aliases.ownerAliases,
    priority: structure.aliases.priorityAliases,
    status: structure.aliases.statusAliases,
    title: structure.aliases.titleAliases,
    dependency: structure.aliases.dependencyAliases,
    dueDate: structure.aliases.dueDateAliases,
    points: structure.aliases.pointsAliases
  };
  const writes = [];

  Object.keys(fields).forEach(function(field) {
    const aliases = aliasByField[field];
    if (!aliases) throw new Error('Unsupported task field: ' + field);
    const columns = aliases
      .map(function(alias) { return structure.map[normalizeHeaderKey_(alias)]; })
      .filter(function(column) { return column !== undefined; });
    if (!columns.length) {
      throw new Error(
        'The configured source sheet "' + sourceSheetName +
        '" has no column for ' + field + '.'
      );
    }
    columns.forEach(function(column) {
      writes.push({ column: column + 1, value: fields[field] });
    });
  });

  writes.forEach(function(write) {
    sheet.getRange(sourceRow, write.column).setValue(write.value);
  });
  SpreadsheetApp.flush();
  patchCachedSourceTask_(taskId, fields);
  return true;
}

// Keep an already-complete catalogue hot after a normal task edit. A cache
// miss remains safe: the next read rebuilds the complete snapshot as before.
function patchCachedSourceTask_(taskId, fields) {
  const cached = readCachedSourceTasks_();
  if (!cached) return;
  const key = clean_(taskId).toLowerCase();
  let changed = false;
  const updated = cached.map(function(task) {
    if (clean_(task.taskId).toLowerCase() !== key) return task;
    changed = true;
    const next = Object.assign({}, task, fields);
    if (Object.prototype.hasOwnProperty.call(fields, 'owner')) {
      next.owner = normalizeOwner_(fields.owner);
    }
    if (Object.prototype.hasOwnProperty.call(fields, 'priority')) {
      next.normalizedPriority = normalizePriority_(fields.priority);
    }
    return next;
  });
  if (changed) cacheSourceTasks_(updated);
}

/* ============================================================
   WEEKLY SHEET
============================================================ */

function createMasterTaskCore_(payload) {
  const actor = requireEditor_();

  if (!payload) {
    throw new Error(
      'No task data was provided.'
    );
  }

  const lock =
    LockService
      .getScriptLock();

  if (
    !lock.tryLock(15000)
  ) {
    throw new Error(
      'Another task is being created right now. Please try again in a moment.'
    );
  }

  try {
    const ss =
      getDashboardSpreadsheet_();

    const title =
      clean_(
        payload.title
      );

    if (!title) {
      throw new Error(
        'Task name is required.'
      );
    }

    const route =
      resolveRoute_(
        payload.product,
        payload.workstream,
        payload.sourceSheet
      );

    const owner =
      normalizeOwner_(
        payload.owner
      );

    const priority = canonicalPriorityInput_(payload.priority || 'P2');

    const status = canonicalStatusInput_(payload.status || 'Backlog');

    if (payload.addToWeek === true && isClosedStatus_(status)) {
      throw new Error('Done or closed tasks cannot be added to a weekly plan.');
    }

    if (
      payload.addToWeek === true &&
      statusUsesEstimatedHours_(status)
    ) {
      validateEstimatedHoursInput_(payload.estimatedHours, true);
    }

    if (
      payload.addToWeek === true &&
      !statusUsesEstimatedHours_(status)
    ) {
      payload.estimatedHours = '';
    }

    const dependency =
      clean_(
        payload.dependency
      );

    const dueDate = validateIsoDateInput_(payload.dueDate, 'Due date');

    const points = validatePointsInput_(payload.points);

    if (payload.addToWeek === true) {
      canonicalCommitmentInput_(payload.commitmentLevel);
    }

    const sourceTasks =
      getSourceTasks_();

    const targetSheet =
      getTargetSheet_(
        route
      );

    const productIds =
      sourceTasks
        .filter(function(task) {
          return (
            canonicalProduct_(
              task.product
            ) ===
            route.product
          );
        })
        .map(function(task) {
          return task.taskId;
        });

    const existingIds =
      productIds.concat(
        getTaskIdsFromSourceSheet_(
          targetSheet
        )
      );

    const taskId =
      generateNextTaskIdFromIds_(
        route,
        existingIds
      );

    const task = {
      taskId: taskId,
      title: title,
      product: route.product,
      workstream: route.workstream,
      sourceSheet: route.sourceSheet,
      owner: owner,
      priority: priority,
      status: status,
      dependency: dependency,
      dueDate: dueDate,
      points: points
    };

    const createdSourceRow =
      appendTaskToSourceSheet_(
        ss,
        route,
        task
      );

    SpreadsheetApp.flush();
    clearSourceTaskCache_();

    let addedToWeek = false;

    if (
      payload.addToWeek ===
      true
    ) {
      addedToWeek =
        addCreatedTaskToWeek_(
          ss,
          task,
          route,
          payload.commitmentLevel,
          payload.weekStart,
          payload.estimatedHours
        );
    }

    SpreadsheetApp.flush();

    const auditEntries = [
      {
        action: 'Created task',
        entityType: 'Task',
        taskId: taskId,
        taskName: title,
        product: route.product,
        sourceSheet: route.sourceSheet,
        field: '',
        oldValue: '',
        newValue: status,
        details:
          'Owner: ' +
          owner +
          ' · Priority: ' +
          priority
      }
    ];

    if (addedToWeek) {
      auditEntries.push({
        action: 'Added to weekly plan',
        entityType: 'Weekly Plan',
        taskId: taskId,
        taskName: title,
        product: route.product,
        sourceSheet: route.sourceSheet,
        weekStart:
          getWeekFromStart_(
            payload.weekStart
          ).start,
        field: 'Commitment',
        oldValue: '',
        newValue:
          normalizeCommitmentLevel_(
            payload.commitmentLevel
          ) +
          (
            statusUsesEstimatedHours_(status)
              ? ' · ' +
                validateEstimatedHoursInput_(
                  payload.estimatedHours,
                  true
                ) +
                'h'
              : ' · no hours'
          )
      });
    }

    logActivities_(
      auditEntries,
      actor
    );

    const createdTask = {
      rowIndex: null,
      sourceRow: createdSourceRow,
      taskId: taskId,
      title: title,
      owner: owner,
      product: route.product,
      rawProduct: route.product,
      workstream: route.workstream,
      sourceSheet: route.sourceSheet,
      priority: priority,
      normalizedPriority: normalizePriority_(priority),
      status: status,
      statusGroup: normalizeStatus_(status),
      dependency: dependency,
      dueDate: dueDate,
      points: points,
      closed: isClosedStatus_(status),
      sourceBacked: true,
      pendingSync: false,
      createdAt: new Date().toISOString()
    };

    return {
      success: true,
      taskId: taskId,
      taskIdAutoGenerated: true,
      targetSheet: route.sourceSheet,
      workstream: route.workstream,
      product: route.product,
      sourceRow: createdSourceRow,
      task: createdTask,
      addedToWeek: addedToWeek,
      weekStart:
        getWeekFromStart_(
          payload.weekStart
        ).start
    };
  } finally {
    lock.releaseLock();
  }
}

function preferredTaskPrefixForRoute_(route) {
  const product = canonicalProduct_(
    route && typeof route === 'object' ? route.product : route
  );
  const sourceSheet = clean_(
    route && typeof route === 'object' ? route.sourceSheet : ''
  ).toLowerCase();

  if (product === 'MediaPal') return 'MED';
  if (product === 'WePlay') return 'WP';

  if (product === 'Nano') {
    if (sourceSheet === 'fe backlog') return 'NANO-FE';
    if (sourceSheet === 'be backlog') return 'NANO-BE';
    if (sourceSheet === 'master backlog') return 'NANO';
  }

  return '';
}

function generateNextTaskIdFromIds_(routeOrProduct, ids) {
  const stats = {};

  (ids || []).forEach(function(id) {
    const cleanId = clean_(id).toUpperCase();
    const match = cleanId.match(/^(.+?)-(\d+)$/);
    if (!match) return;

    const prefix = match[1]
      .replace(/[^A-Z0-9-]/g, '')
      .replace(/-+/g, '-')
      .replace(/^-|-$/g, '');
    const number = Number(match[2]);

    if (!prefix || !Number.isFinite(number)) return;

    if (!stats[prefix]) {
      stats[prefix] = { count: 0, max: 0 };
    }

    stats[prefix].count++;
    stats[prefix].max = Math.max(stats[prefix].max, number);
  });

  const preferred = preferredTaskPrefixForRoute_(routeOrProduct);
  if (preferred) {
    const next = (stats[preferred] ? stats[preferred].max : 0) + 1;
    return preferred + '-' + String(next).padStart(3, '0');
  }

  const prefixes = Object.keys(stats);
  let prefix = 'TASK';
  let next = 1;

  if (prefixes.length) {
    prefixes.sort(function(a, b) {
      const countDifference = stats[b].count - stats[a].count;
      if (countDifference !== 0) return countDifference;
      return stats[b].max - stats[a].max;
    });

    prefix = prefixes[0];
    next = stats[prefix].max + 1;
  } else {
    const product = canonicalProduct_(
      routeOrProduct && typeof routeOrProduct === 'object'
        ? routeOrProduct.product
        : routeOrProduct
    );
    const words = clean_(product)
      .replace(/[^A-Za-z0-9 ]/g, ' ')
      .split(/\s+/)
      .filter(Boolean);

    if (words.length > 1) {
      prefix = words.map(function(word) { return word.charAt(0); }).join('').toUpperCase();
    } else if (words.length === 1) {
      prefix = words[0].slice(0, 4).toUpperCase();
    }
  }

  return prefix + '-' + String(next).padStart(3, '0');
}

/* ============================================================
   V6 — OWNER-EXECUTED APP AUTH + PRODUCT RBAC
   ------------------------------------------------------------
   Deployment model:
     Execute as: Me (app owner)
     Who has access: Anyone with Google account

   Why:
   - Viewers/editors do NOT need direct access to the raw workbooks.
   - The app owner reads/writes Sheets.
   - App identity is established separately:
       1) Google account email from Apps Script when Google exposes it.
       2) Otherwise a one-time email verification code.
   - Every API call carries a short-lived app session token.
   ============================================================ */


/* ============================================================
   SOURCE TASK REGISTER
============================================================ */
function sourceValueByAliases_(row, map, aliases) {
  const index = findColumn_(map, aliases);
  return index === -1 ? '' : getCell_(row, index);
}
function sourceDetailsFromRow_(structure, row) {
  const excluded = new Set([
    'task id','id','issue id','issue key','key','task #','ticket id','requirement id','reference id',
    'task title','task name','requirement / task','feature / issue','repository','task','title','deliverable','requirement',
    'description','requirement / description','purpose / description','requirement description','details',
    'product','source product','workstream','category','track','module','owner','primary owner','raw owner',
    'normalized owner','assignee','priority','status','raw status','task status','normalized status',
    'dependency','dependencies','dependency / related id','blocker','blocker / dependency','due date','due',
    'points','story points'
  ]);
  const details = [];
  structure.headers.forEach(function(header, index) {
    const label = clean_(header);
    const key = normalizeHeaderKey_(label);
    const value = clean_(row[index]);
    if (!label || !value || excluded.has(key)) return;
    details.push({label: label, value: value});
  });
  return details;
}
function sourceTaskFromRow_(
  route,
  sheet,
  structure,
  row,
  rowNumber,
  includeDetails
) {
  const aliases = structure.aliases;

  const taskId =
    clean_(
      sourceValueByAliases_(
        row,
        structure.map,
        aliases.idAliases
      )
    );

  const title =
    clean_(
      sourceValueByAliases_(
        row,
        structure.map,
        aliases.titleAliases
      )
    );

  if (!taskId || !title) return null;

  const description =
    clean_(
      sourceValueByAliases_(
        row,
        structure.map,
        aliases.descriptionAliases
      )
    );

  const owner =
    normalizeOwner_(
      sourceValueByAliases_(
        row,
        structure.map,
        aliases.ownerAliases
      )
    );

  const status =
    clean_(
      sourceValueByAliases_(
        row,
        structure.map,
        aliases.statusAliases
      )
    ) || 'Backlog';

  const priorityRaw =
    clean_(
      sourceValueByAliases_(
        row,
        structure.map,
        aliases.priorityAliases
      )
    );

  const priority =
    priorityRaw || 'No priority';

  const dependency =
    clean_(
      sourceValueByAliases_(
        row,
        structure.map,
        aliases.dependencyAliases
      )
    );

  const dueDate =
    formatDate_(
      sourceValueByAliases_(
        row,
        structure.map,
        aliases.dueDateAliases
      )
    );

  const workstreamRaw =
    clean_(
      sourceValueByAliases_(
        row,
        structure.map,
        aliases.workstreamAliases
      )
    );

  const sourceProduct =
    clean_(
      sourceValueByAliases_(
        row,
        structure.map,
        aliases.productAliases
      )
    );

  const rawPoints =
    sourceValueByAliases_(
      row,
      structure.map,
      aliases.pointsAliases
    );

  let points = '';

  if (
    rawPoints !== '' &&
    rawPoints !== null &&
    rawPoints !== undefined &&
    !isNaN(Number(rawPoints))
  ) {
    points = Number(rawPoints);
  }

  const routedProduct = canonicalProduct_(route.product);
  const product =
    routedProduct === 'MediaPal' || routedProduct === 'WePlay'
      ? resolveTaskProduct_(sourceProduct, routedProduct, taskId)
      : routedProduct;

  return {
    rowIndex: null,
    sourceRow: rowNumber,
    taskId: taskId,
    title: title,
    description: description,
    owner: owner,
    product: product,
    rawProduct: product,
    workstream:
      classifyWorkstream_(
        product,
        route.sourceSheet,
        workstreamRaw || route.workstream
      ),
    sourceSheet: route.sourceSheet,
    sourceWorkbook: sheet.getParent().getName(),
    priority: priority,
    normalizedPriority: normalizePriority_(priority),
    status: status,
    statusGroup: normalizeStatus_(status),
    dependency: dependency,
    dueDate: dueDate,
    points: points,
    closed: isClosedStatus_(status),
    sourceBacked: true,
    pendingSync: false,
    sourceDetails:
      includeDetails
        ? sourceDetailsFromRow_(
            structure,
            row
          )
        : []
  };
}

function readTasksFromSourceRoute_(route, includeDetails) {
  const sheet=getTargetSheet_(route); const structure=getSourceSheetMap_(sheet); const lastRow=sheet.getLastRow();
  if(lastRow<structure.dataStartRow)return [];
  const rows=sheet.getRange(structure.dataStartRow,1,lastRow-structure.dataStartRow+1,structure.lastColumn).getValues();
  return rows.map(function(row,index){return sourceTaskFromRow_(route,sheet,structure,row,structure.dataStartRow+index,includeDetails===true);}).filter(Boolean);
}

function getSourceTaskCacheVersion_() {
  const properties = PropertiesService.getScriptProperties();
  let version = clean_(properties.getProperty(CONFIG.SOURCE_TASK_CACHE_VERSION_PROPERTY_KEY));
  if (!version) {
    version = '1';
    properties.setProperty(CONFIG.SOURCE_TASK_CACHE_VERSION_PROPERTY_KEY, version);
  }
  return version;
}

function sourceTaskCacheKey_(suffix, version) {
  // V2 deliberately discards the earlier cache format, which could contain
  // a snapshot assembled before complete-source verification existed.
  return 'EW_SOURCE_TASKS_V2_' + (version || getSourceTaskCacheVersion_()) + '_' + suffix;
}

function clearSourceTaskCache_() {
  const properties = PropertiesService.getScriptProperties();
  const current = Number(getSourceTaskCacheVersion_()) || 1;
  properties.setProperty(
    CONFIG.SOURCE_TASK_CACHE_PREVIOUS_VERSION_PROPERTY_KEY,
    String(current)
  );
  properties.setProperty(CONFIG.SOURCE_TASK_CACHE_VERSION_PROPERTY_KEY, String(current + 1));
  clearDashboardDataCache_();
}

function getPreviousSourceTaskCacheVersion_() {
  return clean_(
    PropertiesService.getScriptProperties().getProperty(
      CONFIG.SOURCE_TASK_CACHE_PREVIOUS_VERSION_PROPERTY_KEY
    )
  );
}

function readCachedSourceTasks_(requestedVersion) {
  const cache = CacheService.getScriptCache();
  const version = requestedVersion || getSourceTaskCacheVersion_();
  const manifest = cache.get(sourceTaskCacheKey_('manifest', version));
  if (!manifest) return null;

  try {
    const parsedManifest = JSON.parse(manifest);
    // Ignore cache entries created before complete-source verification.
    if (parsedManifest.complete !== true) return null;
    const count = Number(parsedManifest.chunks);
    if (!Number.isInteger(count) || count < 1 || count > 20) return null;
    const parts = [];
    for (let i = 0; i < count; i += 1) {
      const part = cache.get(sourceTaskCacheKey_('chunk_' + i, version));
      if (part === null) return null;
      parts.push(part);
    }
    const tasks = JSON.parse(parts.join(''));
    return Array.isArray(tasks) ? tasks : null;
  } catch (error) {
    return null;
  }
}

function cacheSourceTasks_(tasks) {
  try {
    const payload = JSON.stringify(tasks || []);
    const chunkSize = 75000;
    const chunks = [];
    for (let offset = 0; offset < payload.length; offset += chunkSize) {
      chunks.push(payload.slice(offset, offset + chunkSize));
    }
    if (!chunks.length || chunks.length > 20) return;

    const cache = CacheService.getScriptCache();
    const version = getSourceTaskCacheVersion_();
    chunks.forEach(function(chunk, index) {
      cache.put(sourceTaskCacheKey_('chunk_' + index, version), chunk, CONFIG.SOURCE_TASK_CACHE_SECONDS);
    });
    cache.put(sourceTaskCacheKey_('manifest', version), JSON.stringify({chunks: chunks.length, complete: true}), CONFIG.SOURCE_TASK_CACHE_SECONDS);
  } catch (error) {
    console.warn('Could not cache source tasks.', error);
  }
}

function getConfiguredSourceTasks_() {
  const cached = readCachedSourceTasks_();
  if (cached) return cached;

  const previousVersion = getPreviousSourceTaskCacheVersion_();
  const lastComplete = previousVersion
    ? readCachedSourceTasks_(previousVersion)
    : null;

  const tasks = [];
  const seenTaskIds = {};
  const seenSourceSheets = {};
  const failures = [];

  Object.keys(TASK_ROUTING).forEach(function(product) {
    (TASK_ROUTING[product] || []).forEach(function(routeConfig) {
      const route = {
        product: product,
        workstream: routeConfig.workstream,
        sourceSheet: routeConfig.sourceSheet,
        spreadsheetId: routeConfig.spreadsheetId || ''
      };

      const sourceKey =
        clean_(route.spreadsheetId) +
        '::' +
        clean_(route.sourceSheet).toLowerCase();
      if (seenSourceSheets[sourceKey]) return;
      seenSourceSheets[sourceKey] = true;

      const routeTaskIds = {};
      let routeTasks = [];

      try {
        routeTasks = readTasksFromSourceRoute_(route, false);
      } catch (error) {
        // Never treat partial source data as a successful dashboard snapshot.
        console.warn(
          'Could not load ' + route.product + ' / ' + route.sourceSheet + '.',
          error
        );
        failures.push(route.product + ' / ' + route.sourceSheet);
        return;
      }

      routeTasks.forEach(function(task) {
        const taskKey =
          canonicalProduct_(task.product) +
          '::' +
          clean_(task.taskId).toLowerCase();

        // Prefer the dedicated Nano backlogs over duplicate Master rows.
        if (seenTaskIds[taskKey] && !routeTaskIds[taskKey]) return;

        seenTaskIds[taskKey] = true;
        routeTaskIds[taskKey] = true;
        tasks.push(task);
      });
    });
  });

  if (failures.length) {
    if (lastComplete) {
      console.warn('Using the last complete source-task cache after a source load failure.');
      return lastComplete;
    }

    throw new Error(
      'Could not load all source workbooks (' + failures.join(', ') +
      '). Please try again shortly; no partial dashboard data was applied.'
    );
  }

  cacheSourceTasks_(tasks);
  return tasks;
}

function findSourceTaskById_(taskId, productHint, sourceSheetHint, includeDetails) {
  const key=clean_(taskId).toLowerCase(); if(!key)return null;
  const product=canonicalProduct_(productHint); const sourceSheet=clean_(sourceSheetHint).toLowerCase(); const candidates=[]; const seen={};
  function addRoute(productName,routeConfig){
    const routeKey=clean_(routeConfig.spreadsheetId)+'::'+clean_(routeConfig.sourceSheet).toLowerCase(); if(seen[routeKey])return; seen[routeKey]=true;
    candidates.push({product:canonicalProduct_(productName),workstream:routeConfig.workstream,sourceSheet:routeConfig.sourceSheet,spreadsheetId:routeConfig.spreadsheetId||''});
  }
  if(product&&TASK_ROUTING[product])TASK_ROUTING[product].forEach(function(route){if(!sourceSheet||clean_(route.sourceSheet).toLowerCase()===sourceSheet)addRoute(product,route);});
  Object.keys(TASK_ROUTING).forEach(function(productName){(TASK_ROUTING[productName]||[]).forEach(function(route){addRoute(productName,route);});});
  for(let i=0;i<candidates.length;i++){
    const route=candidates[i];
    try{
      const sheet=getTargetSheet_(route); const rowNumber=findSourceTaskRow_(sheet,taskId); if(rowNumber===-1)continue;
      const structure=getSourceSheetMap_(sheet); const row=sheet.getRange(rowNumber,1,1,structure.lastColumn).getValues()[0];
      return sourceTaskFromRow_(route,sheet,structure,row,rowNumber,includeDetails===true);
    }catch(error){console.warn('Could not hydrate source task '+taskId+' from '+route.sourceSheet+':',error);}
  }
  return null;
}
function hydrateTaskFromSource_(task, includeDetails) {
  if(!task||!task.taskId)return task;
  const sourceTask=findSourceTaskById_(task.taskId,task.product,task.sourceSheet,includeDetails); if(!sourceTask)return task;
  const hydrated=Object.assign({},task,sourceTask); if(Number(task.rowIndex))hydrated.rowIndex=task.rowIndex;
  hydrated.sourceBacked=true; hydrated.pendingSync=false; return hydrated;
}
function normalizeTaskUpdateFields_(fields) {
  const allowed = ['title', 'owner', 'priority', 'status', 'dependency', 'dueDate'];
  const normalized = {};
  Object.keys(fields || {}).forEach(function(field) {
    if (allowed.indexOf(field) === -1) {
      throw new Error('Unsupported task field: ' + field);
    }
    let value = fields[field];
    if (field === 'title') {
      value = clean_(value);
      if (!value) throw new Error('Task title is required.');
    } else if (field === 'owner') {
      value = normalizeOwner_(value);
    } else if (field === 'priority') {
      value = canonicalPriorityInput_(value);
    } else if (field === 'status') {
      value = canonicalStatusInput_(value);
    } else if (field === 'dueDate') {
      value = validateIsoDateInput_(value, 'Due date');
    } else if (field === 'dependency') {
      value = clean_(value);
    }
    normalized[field] = value;
  });
  if (!Object.keys(normalized).length) throw new Error('No task changes were supplied.');
  return normalized;
}

function updateTaskFieldsById(taskId, fields) {
  const access = requireEditor_();
  const task = getTaskById(taskId);
  if (!task) throw new Error('Task was not found.');
  requireProductAccess_(access, task.product, true);

  const updates = normalizeTaskUpdateFields_(fields);
  updateSourceTaskFieldsById_(task.product, task.sourceSheet, task.taskId, updates);
  clearDashboardDataCache_();

  const activities = Object.keys(updates).map(function(field) {
    return {
      action: 'Updated task', entityType: 'Task', taskId: task.taskId,
      taskName: Object.prototype.hasOwnProperty.call(updates, 'title') ? updates.title : task.title,
      product: task.product, sourceSheet: task.sourceSheet, field: field,
      oldValue: task[field], newValue: updates[field]
    };
  });
  logActivities_(activities, access);

  const updatedTask = Object.assign({}, task, updates, {
    normalizedPriority: Object.prototype.hasOwnProperty.call(updates, 'priority')
      ? normalizePriority_(updates.priority)
      : task.normalizedPriority
  });
  return {success: true, task: updatedTask, updatedSource: true, updatedFields: Object.keys(updates)};
}

function updateTaskById(taskId, field, value) {
  return updateTaskFieldsById(taskId, (function() {
    const fields = {};
    fields[field] = value;
    return fields;
  })());
}
function addTaskIdToWeek(taskId, commitmentLevel, selectedWeekStart, estimatedHours) {
  const access=requireEditor_(); const task=getTaskById(taskId); if(!task)throw new Error('Task was not found.'); requireProductAccess_(access,task.product,true);
  if(task.closed)throw new Error('Done or closed tasks cannot be added to a weekly plan.');
  const route=resolveRoute_(task.product,task.workstream,task.sourceSheet);
  const added=addCreatedTaskToWeek_(getDashboardSpreadsheet_(),task,route,commitmentLevel,selectedWeekStart,estimatedHours);
  if(added)logActivities_([{action:'Added to weekly plan',entityType:'Weekly Plan',taskId:task.taskId,taskName:task.title,product:task.product,sourceSheet:task.sourceSheet,weekStart:getWeekFromStart_(selectedWeekStart).start,field:'Commitment',oldValue:'',newValue:canonicalCommitmentInput_(commitmentLevel)+(statusUsesEstimatedHours_(task.status)?' · '+validateEstimatedHoursInput_(estimatedHours,true)+'h':' · no hours')}],access);
  return {success:true,added:added?1:0,skipped:added?[]:[task.taskId]};
}


function getPendingCreatedTasks_() {
  const raw = PropertiesService.getScriptProperties()
    .getProperty(CONFIG.PENDING_CREATED_TASKS_PROPERTY_KEY);

  if (!raw) return {};

  try {
    return JSON.parse(raw) || {};
  } catch (error) {
    console.warn('Could not parse pending created tasks registry.', error);
    return {};
  }
}

function savePendingCreatedTasks_(registry) {
  PropertiesService.getScriptProperties()
    .setProperty(
      CONFIG.PENDING_CREATED_TASKS_PROPERTY_KEY,
      JSON.stringify(registry || {})
    );
}

function registerPendingCreatedTask_(task) {
  if (!task || !clean_(task.taskId)) return;

  const key = clean_(task.taskId).toLowerCase();
  const registry = getPendingCreatedTasks_();

  registry[key] = Object.assign({}, task, {
    taskId: clean_(task.taskId),
    pendingSync: true,
    createdAt: task.createdAt || new Date().toISOString()
  });

  savePendingCreatedTasks_(registry);
}

function removePendingCreatedTask_(taskId) {
  const key = clean_(taskId).toLowerCase();
  if (!key) return;

  const registry = getPendingCreatedTasks_();

  if (registry[key]) {
    delete registry[key];
    savePendingCreatedTasks_(registry);
  }
}

function mergePendingCreatedTasks_(normalizedTasks) {
  const live = Array.isArray(normalizedTasks) ? normalizedTasks.slice() : [];
  const registry = getPendingCreatedTasks_();
  const byId = {};

  live.forEach(function(task) {
    const key = clean_(task.taskId).toLowerCase();
    if (key) byId[key] = true;
  });

  let changed = false;

  Object.keys(registry).forEach(function(key) {
    if (byId[key]) {
      delete registry[key];
      changed = true;
      return;
    }

    const task = registry[key];
    if (!task || !clean_(task.taskId)) {
      delete registry[key];
      changed = true;
      return;
    }

    live.unshift(Object.assign({}, task, {
      rowIndex: null,
      pendingSync: true,
      normalizedPriority: normalizePriority_(task.priority),
      statusGroup: normalizeStatus_(task.status),
      closed: isClosedStatus_(task.status)
    }));
  });

  if (changed) savePendingCreatedTasks_(registry);

  return live;
}

function weeklyTaskAsPendingTask_(weekly) {
  return {
    rowIndex: null,
    taskId: clean_(weekly.taskId),
    title: clean_(weekly.title),
    owner: normalizeOwner_(weekly.owner),
    product: canonicalProduct_(weekly.product),
    rawProduct: clean_(weekly.product),
    workstream: clean_(weekly.workstream),
    sourceSheet: clean_(weekly.sourceSheet),
    sourceRow: Number(weekly.sourceRow) || '',
    priority: clean_(weekly.priority) || 'P2',
    normalizedPriority: normalizePriority_(weekly.priority || 'P2'),
    status: clean_(weekly.weeklyStatus) || 'Backlog',
    statusGroup: normalizeStatus_(weekly.weeklyStatus || 'Backlog'),
    dependency: '',
    dueDate: clean_(weekly.dueDate),
    points: weekly.points === '' ? '' : Number(weekly.points || 0),
    closed: isClosedStatus_(weekly.weeklyStatus),
    pendingSync: true,
    recoveredFromWeeklyPlan: true
  };
}

function mergeWeeklyOnlyTasks_(tasks, weeklyRows) {
  const output = Array.isArray(tasks) ? tasks.slice() : [];
  const ids = {};

  output.forEach(function(task) {
    const key = clean_(task.taskId).toLowerCase();
    if (key) ids[key] = true;
  });

  (weeklyRows || []).forEach(function(weekly) {
    const key = clean_(weekly.taskId).toLowerCase();
    if (!key || ids[key]) return;

    output.unshift(weeklyTaskAsPendingTask_(weekly));
    ids[key] = true;
  });

  return output;
}

function findPendingTask_(taskId) {
  const key = clean_(taskId).toLowerCase();
  if (!key) return null;

  const registry = getPendingCreatedTasks_();
  return registry[key] ? Object.assign({}, registry[key], {
    rowIndex: null,
    pendingSync: true,
    normalizedPriority: normalizePriority_(registry[key].priority),
    statusGroup: normalizeStatus_(registry[key].status),
    closed: isClosedStatus_(registry[key].status)
  }) : null;
}

function findWeeklyTaskById_(ss, taskId) {
  const key = clean_(taskId).toLowerCase();
  if (!key) return null;

  const sheet = ensureWeeklySheet_(ss);
  if (sheet.getLastRow() < 2) return null;

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 16).getValues();

  for (let i = 0; i < rows.length; i++) {
    if (clean_(rows[i][1]).toLowerCase() === key) {
      return serializeWeeklyPlanRow_(rows[i], i);
    }
  }

  return null;
}

function findSourceTaskLocationById_(taskId, productHint, sourceSheetHint) {
  const key = clean_(taskId).toLowerCase();
  if (!key) return null;

  const preferredProduct = canonicalProduct_(productHint);
  const preferredSheet = clean_(sourceSheetHint).toLowerCase();
  const candidates = [];
  const seen = {};

  function addRoute(product, route) {
    const routeKey = clean_(route.spreadsheetId) + '::' + clean_(route.sourceSheet).toLowerCase();
    if (seen[routeKey]) return;
    seen[routeKey] = true;

    candidates.push({
      product: canonicalProduct_(product),
      workstream: route.workstream,
      sourceSheet: route.sourceSheet,
      spreadsheetId: route.spreadsheetId
    });
  }

  if (preferredProduct && TASK_ROUTING[preferredProduct]) {
    TASK_ROUTING[preferredProduct].forEach(function(route) {
      if (!preferredSheet || clean_(route.sourceSheet).toLowerCase() === preferredSheet) {
        addRoute(preferredProduct, route);
      }
    });
  }

  Object.keys(TASK_ROUTING).forEach(function(product) {
    TASK_ROUTING[product].forEach(function(route) {
      addRoute(product, route);
    });
  });

  for (let i = 0; i < candidates.length; i++) {
    const route = candidates[i];

    try {
      const sheet = getTargetSheet_(route);
      const row = findSourceTaskRow_(sheet, taskId);

      if (row) {
        return {
          route: route,
          sheet: sheet,
          row: row
        };
      }
    } catch (error) {
      console.warn(
        'Could not inspect source while locating task ' + taskId + '.',
        error
      );
    }
  }

  return null;
}


function getTaskById(taskId) {
  const access = requireViewer_();
  if (!clean_(taskId)) return null;
  const ss = getDashboardSpreadsheet_();
  const key = clean_(taskId).toLowerCase();
  // The verified catalogue already carries the configured route. Prefer it to
  // avoid scanning every workbook again for each drawer open or save.
  let task = getSourceTasks_().find(function(item) {
    return clean_(item.taskId).toLowerCase() === key;
  }) || null;
  if (!task) task = findSourceTaskById_(taskId, '', '', true);
  if (task) removePendingCreatedTask_(task.taskId);
  if (!task) task = findPendingTask_(taskId);
  if (!task) {
    const weekly = findWeeklyTaskById_(ss, taskId);
    if (weekly) task = weeklyTaskAsPendingTask_(weekly);
  }
  if (!task) return null;
  requireProductAccess_(access, task.product, false);
  return task;
}

function updateMasterTaskField(taskId,field,value) {
  return updateTaskById(taskId,field,value);
}

function createMasterTask(payload) {
  const access=requireEditor_();
  payload=payload||{};
  const route=resolveRoute_(payload.product,payload.workstream,payload.sourceSheet);
  requireProductAccess_(access,route.product,true);
  return createMasterTaskCore_(payload);
}

function deleteTask(taskId) {
  const actor = requireAdmin_();
  const key = clean_(taskId).toLowerCase();

  if (!key) throw new Error('Task ID is required.');

  const ss = getDashboardSpreadsheet_();

  let task = getSourceTasks_().find(function(item) {
    return clean_(item.taskId).toLowerCase() === key;
  }) || null;

  if (!task) task = findPendingTask_(taskId);

  const weeklyMatch = findWeeklyTaskById_(ss, taskId);

  if (!task && weeklyMatch) {
    task = weeklyTaskAsPendingTask_(weeklyMatch);
  }

  let sourceLocation = null;

  if (task) {
    requireProductAccess_(actor, task.product, true);

    sourceLocation = findSourceTaskLocationById_(
      taskId,
      task.product,
      task.sourceSheet
    );
  } else {
    sourceLocation = findSourceTaskLocationById_(
      taskId,
      weeklyMatch && weeklyMatch.product,
      weeklyMatch && weeklyMatch.sourceSheet
    );

    if (sourceLocation) {
      task = {
        taskId: taskId,
        title: weeklyMatch ? weeklyMatch.title : taskId,
        owner: weeklyMatch ? weeklyMatch.owner : '',
        product: sourceLocation.route.product,
        workstream: sourceLocation.route.workstream,
        sourceSheet: sourceLocation.route.sourceSheet,
        status: weeklyMatch ? weeklyMatch.weeklyStatus : '',
        priority: weeklyMatch ? weeklyMatch.priority : '',
        dueDate: weeklyMatch ? weeklyMatch.dueDate : '',
        points: weeklyMatch ? weeklyMatch.points : ''
      };

      requireProductAccess_(actor, task.product, true);
    }
  }

  if (!task && !sourceLocation && !weeklyMatch) {
    throw new Error(
      'Task ' + taskId +
      ' was not found in its configured source sheets, pending task registry, or weekly plan.'
    );
  }

  if (!sourceLocation) {
    sourceLocation = findSourceTaskLocationById_(
      taskId,
      task && task.product,
      task && task.sourceSheet
    );
  }

  let deletedFromSource = false;

  if (sourceLocation && sourceLocation.row) {
    sourceLocation.sheet.deleteRow(sourceLocation.row);
    deletedFromSource = true;
  }

  const weekly = ensureWeeklySheet_(ss);
  let removedWeeklyRows = 0;

  if (weekly.getLastRow() > 1) {
    const ids = weekly
      .getRange(2, 2, weekly.getLastRow() - 1, 1)
      .getDisplayValues();

    for (let i = ids.length - 1; i >= 0; i--) {
      if (clean_(ids[i][0]).toLowerCase() === key) {
        weekly.deleteRow(i + 2);
        removedWeeklyRows++;
      }
    }
  }

  removePendingCreatedTask_(taskId);
  if (deletedFromSource) clearSourceTaskCache_();
  clearAllWeeklyPlanCache_();
  SpreadsheetApp.flush();

  logActivities_([{
    action: 'Deleted task',
    entityType: 'Task',
    taskId: taskId,
    taskName: task && task.title ? task.title : taskId,
    product: task && task.product ? task.product : '',
    sourceSheet:
      task && task.sourceSheet
        ? task.sourceSheet
        : sourceLocation
          ? sourceLocation.route.sourceSheet
          : '',
    field: 'Task',
    oldValue: JSON.stringify({
      owner: task && task.owner ? task.owner : '',
      status: task && task.status ? task.status : '',
      priority: task && task.priority ? task.priority : '',
      dueDate: task && task.dueDate ? task.dueDate : '',
      points: task && task.points !== undefined ? task.points : ''
    }),
    newValue: 'Deleted',
    details:
      'Permanent deletion. Source row removed: ' +
      (deletedFromSource ? 'yes' : 'not found') +
      ' · Weekly plan rows removed: ' +
      removedWeeklyRows
  }], actor);

  return {
    success: true,
    taskId: taskId,
    deletedFromSource: deletedFromSource,
    removedWeeklyRows: removedWeeklyRows
  };
}



/* ============================================================
   V7 PUBLIC WEEKLY PLANNING WRAPPER
============================================================ */
