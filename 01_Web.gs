function include(filename) {
  return HtmlService
    .createHtmlOutputFromFile(filename)
    .getContent();
}

function doGet() {
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Engineering Workload')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function isDashboardSpreadsheet_(spreadsheet) {
  return !!(
    spreadsheet &&
    spreadsheet.getSheetByName(CONFIG.WEEKLY_SHEET)
  );
}

/* ============================================================
   DASHBOARD WORKBOOK ID
   - Needed because direct-edit triggers run from the source
     workbook that was edited. The audit log must always be
     written back to the dashboard workbook.
============================================================ */

function rememberDashboardSpreadsheet_() {
  let active = null;

  try {
    active = SpreadsheetApp.getActiveSpreadsheet();

    if (isDashboardSpreadsheet_(active)) {
      PropertiesService
        .getScriptProperties()
        .setProperty(
          CONFIG.DASHBOARD_ID_PROPERTY_KEY,
          active.getId()
        );

      return active;
    }
  } catch (error) {
    console.warn(
      'Could not inspect the active dashboard spreadsheet.',
      error
    );
  }

  return getDashboardSpreadsheet_();
}

function getConfiguredDashboardId_() {
  return clean_(
    PropertiesService
      .getScriptProperties()
      .getProperty(
        CONFIG.DASHBOARD_ID_PROPERTY_KEY
      )
  );
}

function buildAccessDeniedError_(
  scope,
  label,
  spreadsheetId,
  requiredPermission,
  originalError
) {
  const email =
    getCurrentUserEmail_() ||
    'unidentified Google account';

  const safeScope =
    clean_(scope) ||
    'WORKBOOK';

  const safeLabel =
    clean_(label) ||
    'requested workbook';

  const safePermission =
    clean_(requiredPermission) ||
    'access';

  const details =
    originalError &&
    originalError.message
      ? clean_(originalError.message)
      : '';

  return new Error(
    [
      'APP_ACCESS_DENIED',
      safeScope,
      safeLabel,
      clean_(spreadsheetId),
      email,
      safePermission,
      details
    ].join('||')
  );
}

function getDashboardSpreadsheet_() {
  const properties =
    PropertiesService.getScriptProperties();

  const configuredId =
    getConfiguredDashboardId_();

  if (configuredId) {
    try {
      return openSpreadsheetWithRetry_(configuredId);
    } catch (error) {
      try {
        const active =
          SpreadsheetApp.getActiveSpreadsheet();

        if (isDashboardSpreadsheet_(active)) {
          properties.setProperty(
            CONFIG.DASHBOARD_ID_PROPERTY_KEY,
            active.getId()
          );

          return active;
        }
      } catch (activeError) {
        console.warn(
          'Could not use active spreadsheet as a dashboard fallback.',
          activeError
        );
      }

      throw buildAccessDeniedError_(
        'DASHBOARD',
        CONFIG.DASHBOARD_LABEL,
        configuredId,
        'Viewer access or higher',
        error
      );
    }
  }

  try {
    const active =
      SpreadsheetApp.getActiveSpreadsheet();

    if (isDashboardSpreadsheet_(active)) {
      properties.setProperty(
        CONFIG.DASHBOARD_ID_PROPERTY_KEY,
        active.getId()
      );

      return active;
    }
  } catch (error) {
    console.warn(
      'Could not inspect active spreadsheet while resolving the dashboard.',
      error
    );
  }

  throw new Error(
    'APP_CONFIGURATION_ERROR||Dashboard spreadsheet ID is not configured. Open the dashboard-bound Apps Script project as an admin and run installDirectEditAuditTriggers() once.'
  );
}

/* ============================================================
   ACCESS CONTROL
   - Role source of truth is Script Properties, NOT a shared sheet.
   - Unknown / unidentified users are Viewer-only.
   - Mutation functions enforce permissions server-side.
============================================================ */
