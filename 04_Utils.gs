function canonicalPriorityInput_(value) {
  const normalized = normalizePriority_(value);

  if (!VALID_TASK_PRIORITIES.includes(normalized)) {
    throw new Error(
      'Priority must be one of: ' + VALID_TASK_PRIORITIES.join(', ') + '.'
    );
  }

  return normalized;
}

function canonicalStatusInput_(value) {
  const raw = clean_(value);
  if (!raw) return 'Backlog';

  const match = VALID_TASK_STATUSES.find(function(status) {
    return status.toLowerCase() === raw.toLowerCase();
  });

  if (!match) {
    throw new Error(
      'Status must be one of: ' + VALID_TASK_STATUSES.join(', ') + '.'
    );
  }

  return match;
}

function canonicalCommitmentInput_(value) {
  const raw = clean_(value);
  const normalized = normalizeCommitmentLevel_(raw);

  if (
    raw &&
    !VALID_WEEKLY_COMMITMENTS.some(function(item) {
      return item.toLowerCase() === raw.toLowerCase();
    })
  ) {
    throw new Error(
      'Weekly commitment must be Commit, Stretch, or Watch / Unblock.'
    );
  }

  return normalized;
}

function validateIsoDateInput_(value, label) {
  const raw = clean_(value);
  if (!raw) return '';

  if (!/^\d{4}-\d{2}-\d{2}$/.test(raw)) {
    throw new Error((label || 'Date') + ' must use YYYY-MM-DD format.');
  }

  const parts = raw.split('-').map(Number);
  const date = new Date(parts[0], parts[1] - 1, parts[2]);

  if (
    date.getFullYear() !== parts[0] ||
    date.getMonth() !== parts[1] - 1 ||
    date.getDate() !== parts[2]
  ) {
    throw new Error((label || 'Date') + ' is not a valid calendar date.');
  }

  return raw;
}

function validatePointsInput_(value) {
  if (value === '' || value === null || value === undefined) return '';

  const number = Number(value);
  if (!Number.isFinite(number) || number < 0) {
    throw new Error('Points must be a non-negative number.');
  }

  return number;
}

function statusUsesEstimatedHours_(status) {
  const group = normalizeStatus_(status);
  return (
    group !== 'QA' &&
    group !== 'In Review' &&
    group !== 'Done'
  );
}

function validateEstimatedHoursInput_(value, required) {
  const blank =
    value === '' ||
    value === null ||
    value === undefined;

  if (blank) {
    if (required) {
      throw new Error('Estimated hours are required when planning a task.');
    }
    return '';
  }

  const hours = Number(value);

  if (
    !Number.isFinite(hours) ||
    hours <= 0 ||
    hours > 168
  ) {
    throw new Error('Estimated hours must be greater than 0 and no more than 168.');
  }

  return Math.round(hours * 4) / 4;
}

function getCurrentWeek_() {
  return getWeekFromDate_(new Date());
}

function getWeekFromStart_(weekStart) {
  if (!weekStart) {
    return getCurrentWeek_();
  }

  const parts = String(weekStart).split('-').map(Number);

  if (
    parts.length !== 3 ||
    !parts[0] ||
    !parts[1] ||
    !parts[2]
  ) {
    return getCurrentWeek_();
  }

  const date = new Date(parts[0], parts[1] - 1, parts[2]);
  return getWeekFromDate_(date);
}

function getWeekFromDate_(date) {
  const timezone = Session.getScriptTimeZone();
  const base = new Date(date);
  const day = base.getDay();
  const mondayOffset = day === 0 ? -6 : 1 - day;

  const monday = new Date(base);
  monday.setDate(base.getDate() + mondayOffset);
  monday.setHours(0, 0, 0, 0);

  const friday = new Date(monday);
  friday.setDate(monday.getDate() + 4);

  return {
    start: Utilities.formatDate(monday, timezone, 'yyyy-MM-dd'),
    end: Utilities.formatDate(friday, timezone, 'yyyy-MM-dd')
  };
}

function shiftWeekStart_(weekStart, offsetWeeks) {
  const week = getWeekFromStart_(weekStart);
  const parts = week.start.split('-').map(Number);
  const date = new Date(parts[0], parts[1] - 1, parts[2]);

  date.setDate(
    date.getDate() +
    (Number(offsetWeeks) || 0) * 7
  );

  return getWeekFromDate_(date).start;
}

function getWeekWindowStarts_(centerWeekStart, radius) {
  radius = Math.max(
    0,
    Number(radius) || 0
  );

  const starts = [];

  for (let offset = -radius; offset <= radius; offset++) {
    starts.push(
      shiftWeekStart_(
        centerWeekStart,
        offset
      )
    );
  }

  return starts;
}

/* ============================================================
   BASIC HELPERS
============================================================ */

function clean_(value) {
  return String(
    value === null || value === undefined ? '' : value
  ).trim();
}

function firstNonEmpty_(values) {
  for (let i = 0; i < values.length; i++) {
    const value = clean_(values[i]);
    if (value) return value;
  }
  return '';
}

function formatDate_(value) {
  if (!value) return '';

  if (value instanceof Date) {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    );
  }

  return clean_(value);
}

function normalizeOwner_(value) {
  const owner = clean_(value);
  if (!owner) return 'Unassigned';

  const lower = owner.toLowerCase();
  const aliases = {
    'abhishek': 'Abhishek', 'abhishek@nostacklabs.com': 'Abhishek', 'abhishek.verma@mediapal.net': 'Abhishek',
    'ambrose': 'Ambrose', 'ambrose.ahawo@mediapal.net': 'Ambrose',
    'antony': 'Antony', 'anthony': 'Antony', 'tulla': 'Antony', 'tula': 'Antony',
    'antony tulla': 'Antony', 'anthony tulla': 'Antony', 'anthony.tulla@mediapal.net': 'Antony',
    'james': 'James', 'james.odera@mediapal.net': 'James',
    'maryanne': 'Maryanne', 'maryanne.nduhiu@mediapal.net': 'Maryanne',
    'pthapli@nostacklabs.com': 'PT', 'pratyush thapliyal': 'PT', 'pt': 'PT',
    'rohit': 'Rohit', 'rohit tatu': 'Rohit', 'rtatu@nostacklabs.com': 'Rohit', 'rtatu@weplayarcade.com': 'Rohit',
    'yash': 'Yash', 'yash@weplayarcade.com': 'Yash'
  };

  if (aliases[lower]) return aliases[lower];

  return owner;
}

/* ============================================================
   STATUS
============================================================ */

function normalizeStatus_(status) {
  const s = clean_(status).toLowerCase();

  if (
    ['done', 'complete', 'completed', 'resolved', 'passed', 'closed'].includes(s)
  ) {
    return 'Done';
  }
  if (
  String(status || '').trim().toLowerCase() === 'future backlog'
) {
  return 'Backlog';
}

  if (
    
    s === 'backlog' ||
    s === 'none' ||
    s === 'not started' ||
    s === 'not-started' ||
    s === 'not_started'
  ) {
    return 'Backlog';
  }

  if (s === 'todo' || s === 'to do') return 'Todo';

  if (
    s.includes('progress') ||
    s === 'active' ||
    s === 'doing'
  ) {
    return 'In Progress';
  }

  if (s.includes('review')) return 'In Review';

  if (
    s === 'qa' ||
    s.includes('testing') ||
    s === 'test' ||
    s.includes('retest')
  ) {
    return 'QA';
  }

  if (
    s === 'blocked' ||
    s.includes('awaiting tech fix') ||
    s === 'failed'
  ) {
    return 'Blocked';
  }

  return 'Todo';
}

function isClosedStatus_(status) {
  const s = clean_(status).toLowerCase();

  return [
    'done',
    'complete',
    'completed',
    'resolved',
    'passed',
    'closed',
    'legacy',
    'replaced',
    'cancelled',
    'canceled'
  ].includes(s);
}

/* ============================================================
   PRIORITY / COMMITMENT
============================================================ */

function normalizePriority_(priority) {
  const raw = clean_(priority);
  const p = raw.toLowerCase();

  if (p === 'p0' || p === 'critical' || p === 'urgent') return 'P0';
  if (p === 'p1' || p === 'high') return 'P1';
  if (p === 'p2' || p === 'medium') return 'P2';
  if (p === 'p3' || p === 'low') return 'P3';

  return raw || 'P2';
}

function normalizeCommitmentLevel_(value) {
  const c = clean_(value).toLowerCase();

  if (c === 'stretch') return 'Stretch';
  if (c.includes('watch') || c.includes('unblock')) return 'Watch / Unblock';

  return 'Commit';
}

/* ============================================================
   HEADER HELPERS
============================================================ */

function normalizeHeaderKey_(value) {
  return clean_(value)
    .toLowerCase()
    .replace(/[\u2013\u2014]/g, '-')
    .replace(/\s+/g, ' ')
    .trim();
}

function buildHeaderMap_(headers) {
  const map = {};

  headers.forEach(function(header, index) {
    const key = normalizeHeaderKey_(header);
    if (key && !Object.prototype.hasOwnProperty.call(map, key)) {
      map[key] = index;
    }
  });

  return map;
}

function findColumn_(map, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const key = normalizeHeaderKey_(aliases[i]);

    if (Object.prototype.hasOwnProperty.call(map, key)) {
      return map[key];
    }
  }

  return -1;
}

function getCell_(row, index) {
  if (index === -1 || index === null || index === undefined) return '';
  return row[index];
}

/* ============================================================
   MASTER COLUMN MAP
============================================================ */
