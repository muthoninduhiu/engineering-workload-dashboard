function scopeTasksForAccess_(tasks, access) {
  return (tasks || []).filter(function(task) { return canAccessProduct_(access, task.product); });
}

function scopeWeeklyForAccess_(rows, access) {
  return (rows || []).filter(function(item) { return canAccessProduct_(access, item.product); });
}

function getDashboardCacheVersion_() {
  const properties = PropertiesService.getScriptProperties();
  let version = clean_(properties.getProperty(CONFIG.DASHBOARD_CACHE_VERSION_PROPERTY_KEY));
  if (!version) {
    version = '1';
    properties.setProperty(CONFIG.DASHBOARD_CACHE_VERSION_PROPERTY_KEY, version);
  }
  return version;
}

function dashboardCacheScope_(access, weekStart) {
  const products = access && access.allProducts
    ? 'all'
    : ((access && access.products) || []).slice().sort().join(',');
  return [
    clean_(access && access.email).toLowerCase().replace(/[^a-z0-9]/g, '_'),
    clean_(access && access.role).toLowerCase(),
    products.replace(/[^a-zA-Z0-9,]/g, '_'),
    getWeekFromStart_(weekStart).start
  ].join('_');
}

function dashboardCacheKey_(suffix, access, weekStart, version) {
  return 'EW_DASHBOARD_V1_' + (version || getDashboardCacheVersion_()) + '_' +
    dashboardCacheScope_(access, weekStart) + '_' + suffix;
}

function readCachedDashboardData_(access, weekStart) {
  try {
    const cache = CacheService.getScriptCache();
    const manifest = cache.get(dashboardCacheKey_('manifest', access, weekStart));
    if (!manifest) return null;
    const parsed = JSON.parse(manifest);
    const count = Number(parsed.chunks);
    if (!parsed.complete || !Number.isInteger(count) || count < 1 || count > 20) return null;
    const chunks = [];
    for (let index = 0; index < count; index += 1) {
      const chunk = cache.get(dashboardCacheKey_('chunk_' + index, access, weekStart));
      if (chunk === null) return null;
      chunks.push(chunk);
    }
    return JSON.parse(chunks.join(''));
  } catch (error) {
    return null;
  }
}

function cacheDashboardData_(access, weekStart, data) {
  try {
    const payload = JSON.stringify(data);
    const chunkSize = 75000;
    const chunks = [];
    for (let offset = 0; offset < payload.length; offset += chunkSize) {
      chunks.push(payload.slice(offset, offset + chunkSize));
    }
    if (!chunks.length || chunks.length > 20) return;
    const cache = CacheService.getScriptCache();
    chunks.forEach(function(chunk, index) {
      cache.put(
        dashboardCacheKey_('chunk_' + index, access, weekStart),
        chunk,
        CONFIG.DASHBOARD_CACHE_SECONDS
      );
    });
    cache.put(
      dashboardCacheKey_('manifest', access, weekStart),
      JSON.stringify({chunks: chunks.length, complete: true}),
      CONFIG.DASHBOARD_CACHE_SECONDS
    );
  } catch (error) {
    console.warn('Could not cache dashboard data.', error);
  }
}

function clearDashboardDataCache_() {
  const properties = PropertiesService.getScriptProperties();
  const current = Number(getDashboardCacheVersion_()) || 1;
  properties.setProperty(CONFIG.DASHBOARD_CACHE_VERSION_PROPERTY_KEY, String(current + 1));
}

function getDashboardData(selectedWeekStart) {
  const access = requireViewer_();
  const cached = readCachedDashboardData_(access, selectedWeekStart);
  if (cached) return cached;
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

  const data = {
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
  cacheDashboardData_(access, week.start, data);
  return data;
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
