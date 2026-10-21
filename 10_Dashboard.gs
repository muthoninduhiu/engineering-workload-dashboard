function scopeTasksForAccess_(tasks, access) {
  return (tasks || []).filter(function(task) { return canAccessProduct_(access, task.product); });
}

function scopeWeeklyForAccess_(rows, access) {
  return (rows || []).filter(function(item) { return canAccessProduct_(access, item.product); });
}

function getDashboardData(selectedWeekStart) {
  const access = requireViewer_();
  const ss = rememberDashboardSpreadsheet_();
  ensureWeeklySheet_(ss);

  const week = getWeekFromStart_(selectedWeekStart);
  const todayWeek = getCurrentWeek_();

  const weekStarts =
    getWeekWindowStarts_(
      week.start,
      CONFIG.WEEK_WINDOW_RADIUS
    );

  const rawWindow =
    getWeeklyPlanRowsForWeeks_(
      ss,
      weekStarts
    );

  const weekWindow = {};
  const capacityWindow = {};
  const capacityOverrides = readCapacityOverridesForWeeks_(ss, weekStarts);

  weekStarts.forEach(function(start) {
    const scoped =
      scopeWeeklyForAccess_(
        rawWindow[start] || [],
        access
      );

    weekWindow[start] = scoped;

    const owners =
      Array.from(
        new Set(
          scoped
            .map(function(item) {
              return normalizeOwner_(item.owner);
            })
            .filter(Boolean)
        )
      );

    capacityWindow[start] =
      getCapacityForWeek_(
        ss,
        start,
        owners,
        capacityOverrides[start]
      );
  });

  const sourceTasks = getSourceTasks_();
  const withPending = mergePendingCreatedTasks_(sourceTasks);
  const withWeeklyOnly = mergeWeeklyOnlyTasks_(
    withPending,
    rawWindow[week.start] || []
  );

  const tasks =
    scopeTasksForAccess_(
      withWeeklyOnly,
      access
    );

  return {
    tasks: tasks,
    weeklyPlan: weekWindow[week.start] || [],
    weekWindow: weekWindow,
    capacity: capacityWindow[week.start] || {
      weekStart: week.start,
      defaultHours: CONFIG.DEFAULT_WEEKLY_CAPACITY_HOURS,
      owners: {}
    },
    capacityWindow: capacityWindow,
    week: week,
    todayWeek: todayWeek,
    routing: getRoutingConfig_(),
    access: access,
    validation: buildDataValidationReport_(
      tasks,
      weekWindow[week.start] || []
    )
  };
}

function getInitialDashboardData() {
  return getDashboardData(null);
}

// Keeps app startup to one Apps Script request for authorized users.
function getStartupDashboardData() {
  const access = getAccessContext_();
  const bootstrap = {
    access: access,
    deployment: {
      expectedExecuteAs: 'Me (app owner)',
      expectedWhoHasAccess: 'Anyone with Google account'
    },
    dashboard: {
      label: CONFIG.DASHBOARD_LABEL,
      spreadsheetId: getConfiguredDashboardId_()
    }
  };

  return {
    bootstrap: bootstrap,
    data: access.canView ? getDashboardData(null) : null
  };
}

// A user-initiated refresh must bypass the short-lived source-task cache.
// Normal navigation and reloads remain fast by using the cached catalogue.
function refreshDashboardData(selectedWeekStart) {
  requireViewer_();
  clearSourceTaskCache_();
  return getDashboardData(selectedWeekStart);
}

function getWeekDataBatch(selectedWeekStarts) {
  const access = requireViewer_();
  const ss = getDashboardSpreadsheet_();

  selectedWeekStarts =
    Array.isArray(selectedWeekStarts)
      ? selectedWeekStarts
      : [selectedWeekStarts];

  const normalizedStarts = [];

  selectedWeekStarts.forEach(function(value) {
    const start = getWeekFromStart_(value).start;
    if (normalizedStarts.indexOf(start) === -1) {
      normalizedStarts.push(start);
    }
  });

  const grouped =
    getWeeklyPlanRowsForWeeks_(
      ss,
      normalizedStarts
    );

  const todayWeek = getCurrentWeek_();

  return normalizedStarts.map(function(start) {
    const weeklyPlan =
      scopeWeeklyForAccess_(
        grouped[start] || [],
        access
      );

    const owners =
      Array.from(
        new Set(
          weeklyPlan
            .map(function(item) {
              return normalizeOwner_(item.owner);
            })
            .filter(Boolean)
        )
      );

    return {
      week: getWeekFromStart_(start),
      todayWeek: todayWeek,
      weeklyPlan: weeklyPlan,
      capacity: getCapacityForWeek_(
        ss,
        start,
        owners
      )
    };
  });
}

function getWeekData(selectedWeekStart) {
  const batch =
    getWeekDataBatch(
      [selectedWeekStart]
    );

  return batch.length
    ? batch[0]
    : {
        week: getWeekFromStart_(selectedWeekStart),
        todayWeek: getCurrentWeek_(),
        weeklyPlan: [],
        capacity: {
          weekStart: getWeekFromStart_(selectedWeekStart).start,
          defaultHours: CONFIG.DEFAULT_WEEKLY_CAPACITY_HOURS,
          owners: {}
        }
      };
}
