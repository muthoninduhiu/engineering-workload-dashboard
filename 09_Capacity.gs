function ensureCapacitySheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.CAPACITY_SHEET);

  if (!sheet) {
    const access = getAccessContext_();

    if (!access.canEdit) {
      return null;
    }

    sheet = ss.insertSheet(CONFIG.CAPACITY_SHEET);
  }

  const headers = [
    'Week Start',
    'Owner',
    'Available Hours',
    'Note',
    'Updated At',
    'Updated By'
  ];

  const range = sheet.getRange(1, 1, 1, headers.length);
  const existing = range.getDisplayValues()[0];

  let mismatch = false;
  for (let i = 0; i < headers.length; i++) {
    if (clean_(existing[i]) !== headers[i]) {
      mismatch = true;
      break;
    }
  }

  if (mismatch) {
    const access = getAccessContext_();
    if (access.canEdit) {
      range
        .setValues([headers])
        .setFontWeight('bold')
        .setBackground('#f1f5f9');
      sheet.setFrozenRows(1);
    }
  }

  return sheet;
}

function readCapacityOverrides_(ss, weekStart) {
  const sheet = ensureCapacitySheet_(ss);
  const output = {};

  if (!sheet || sheet.getLastRow() < 2) {
    return output;
  }

  const week = getWeekFromStart_(weekStart).start;
  const rows = sheet
    .getRange(2, 1, sheet.getLastRow() - 1, 6)
    .getValues();

  rows.forEach(function(row) {
    if (formatDate_(row[0]) !== week) return;

    const owner = normalizeOwner_(row[1]);
    if (!owner) return;

    const hours = Number(row[2]);
    if (!Number.isFinite(hours) || hours < 0) return;

    output[owner] = {
      owner: owner,
      availableHours: Math.round(hours * 4) / 4,
      note: clean_(row[3]),
      isOverride: true
    };
  });

  return output;
}

function readCapacityOverridesForWeeks_(ss, weekStarts) {
  const sheet = ensureCapacitySheet_(ss);
  const output = {};
  const wanted = {};

  (weekStarts || []).forEach(function(value) {
    wanted[getWeekFromStart_(value).start] = true;
  });
  Object.keys(wanted).forEach(function(week) { output[week] = {}; });

  if (!sheet || sheet.getLastRow() < 2 || !Object.keys(wanted).length) return output;

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  rows.forEach(function(row) {
    const week = formatDate_(row[0]);
    if (!wanted[week]) return;
    const owner = normalizeOwner_(row[1]);
    const hours = Number(row[2]);
    if (!owner || !Number.isFinite(hours) || hours < 0) return;
    output[week][owner] = {owner: owner, availableHours: Math.round(hours * 4) / 4, note: clean_(row[3]), isOverride: true};
  });

  return output;
}

function getCapacityForWeek_(ss, weekStart, owners, preloadedOverrides) {
  const week = getWeekFromStart_(weekStart).start;
  const overrides = preloadedOverrides || readCapacityOverrides_(ss, week);
  const output = {};

  (owners || []).forEach(function(ownerValue) {
    const owner = normalizeOwner_(ownerValue);
    if (!owner) return;

    output[owner] =
      overrides[owner] ||
      {
        owner: owner,
        availableHours: CONFIG.DEFAULT_WEEKLY_CAPACITY_HOURS,
        note: '',
        isOverride: false
      };
  });

  return {
    weekStart: week,
    defaultHours: CONFIG.DEFAULT_WEEKLY_CAPACITY_HOURS,
    owners: output
  };
}

function setOwnerCapacity(owner, availableHours, selectedWeekStart, note) {
  const actor = requireAdmin_();
  owner = normalizeOwner_(owner);

  if (!owner || owner === 'Unassigned') {
    throw new Error('Choose a named owner for a capacity override.');
  }

  const hours = Number(availableHours);
  if (!Number.isFinite(hours) || hours < 0 || hours > 168) {
    throw new Error('Available hours must be between 0 and 168.');
  }

  const roundedHours = Math.round(hours * 4) / 4;
  const ss = getDashboardSpreadsheet_();
  const sheet = ensureCapacitySheet_(ss);
  const week = getWeekFromStart_(selectedWeekStart).start;
  const lastRow = sheet.getLastRow();

  let targetRow = 0;
  if (lastRow >= 2) {
    const rows = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (let i = 0; i < rows.length; i++) {
      if (
        formatDate_(rows[i][0]) === week &&
        normalizeOwner_(rows[i][1]).toLowerCase() === owner.toLowerCase()
      ) {
        targetRow = i + 2;
        break;
      }
    }
  }

  const oldHours =
    targetRow
      ? sheet.getRange(targetRow, 3).getValue()
      : CONFIG.DEFAULT_WEEKLY_CAPACITY_HOURS;

  const row = [
    week,
    owner,
    roundedHours,
    clean_(note),
    new Date(),
    actor.email
  ];

  if (targetRow) {
    sheet.getRange(targetRow, 1, 1, 6).setValues([row]);
  } else {
    sheet.appendRow(row);
  }

  logActivities_([{
    action: 'Updated weekly capacity',
    entityType: 'Capacity',
    taskName: owner,
    weekStart: week,
    field: 'Available Hours',
    oldValue: oldHours,
    newValue: roundedHours,
    details: clean_(note)
  }], actor);

  return {
    success: true,
    weekStart: week,
    owner: owner,
    availableHours: roundedHours,
    note: clean_(note),
    isOverride: true
  };
}

function resetOwnerCapacity(owner, selectedWeekStart) {
  const actor = requireAdmin_();
  owner = normalizeOwner_(owner);

  const ss = getDashboardSpreadsheet_();
  const sheet = ensureCapacitySheet_(ss);
  const week = getWeekFromStart_(selectedWeekStart).start;

  if (sheet && sheet.getLastRow() >= 2) {
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();

    for (let i = rows.length - 1; i >= 0; i--) {
      if (
        formatDate_(rows[i][0]) === week &&
        normalizeOwner_(rows[i][1]).toLowerCase() === owner.toLowerCase()
      ) {
        sheet.deleteRow(i + 2);
      }
    }
  }

  logActivities_([{
    action: 'Reset weekly capacity',
    entityType: 'Capacity',
    taskName: owner,
    weekStart: week,
    field: 'Available Hours',
    newValue: CONFIG.DEFAULT_WEEKLY_CAPACITY_HOURS
  }], actor);

  return {
    success: true,
    weekStart: week,
    owner: owner,
    availableHours: CONFIG.DEFAULT_WEEKLY_CAPACITY_HOURS,
    note: '',
    isOverride: false
  };
}


/* ============================================================
   PENDING CREATED TASK REGISTRY
   - New tasks are written to their real source sheet first.
   - Source reads can briefly lag after a new task is written.
   - Keep a tiny persistent overlay so the task remains visible
     until it appears in the source-backed task list.
============================================================ */
