/* ============================================================
   SCHEMA-AWARE GOOGLE SHEETS VALIDATION
   ------------------------------------------------------------
   Safety rules:
   - Never rename headers.
   - Never convert/remove Google Sheets tables.
   - Never overwrite a typed table column.
   - Never overwrite a different existing cell validation rule.
   - If the requested validation is already present, report it as skipped.
============================================================ */

function validationValueForSignature_(value) {
  if (value instanceof Date) {
    return value.toISOString();
  }

  if (Array.isArray(value)) {
    return value.map(validationValueForSignature_);
  }

  if (
    value &&
    typeof value === 'object' &&
    typeof value.getA1Notation === 'function'
  ) {
    return {
      a1: value.getA1Notation(),
      sheet:
        value.getSheet
          ? value.getSheet().getName()
          : ''
    };
  }

  if (
    value === null ||
    value === undefined
  ) {
    return value;
  }

  return String(value);
}

function validationRuleSignature_(rule) {
  if (!rule) return null;

  return JSON.stringify({
    criteriaType:
      String(
        rule.getCriteriaType()
      ),

    criteriaValues:
      (rule.getCriteriaValues() || [])
        .map(validationValueForSignature_),

    allowInvalid:
      rule.getAllowInvalid()
  });
}

function validationRulesEquivalent_(a, b) {
  if (!a || !b) return false;

  try {
    return (
      validationRuleSignature_(a) ===
      validationRuleSignature_(b)
    );
  } catch (error) {
    return false;
  }
}

function validationA1Range_(sheet, startRow, column, rowCount) {
  if (!rowCount) return '';

  return sheet
    .getRange(
      startRow,
      column,
      rowCount,
      1
    )
    .getA1Notation();
}

function getSheetsTableCatalog_(spreadsheetId) {
  const fields =
    'sheets(' +
      'properties(sheetId,title),' +
      'tables(' +
        'tableId,' +
        'name,' +
        'range(startRowIndex,endRowIndex,startColumnIndex,endColumnIndex),' +
        'columnProperties(' +
          'columnIndex,' +
          'columnName,' +
          'columnType,' +
          'dataValidationRule' +
        ')' +
      ')' +
    ')';

  const url =
    'https://sheets.googleapis.com/v4/spreadsheets/' +
    encodeURIComponent(spreadsheetId) +
    '?fields=' +
    encodeURIComponent(fields);

  try {
    const response =
      UrlFetchApp.fetch(
        url,
        {
          method: 'get',
          headers: {
            Authorization:
              'Bearer ' +
              ScriptApp.getOAuthToken()
          },
          muteHttpExceptions: true
        }
      );

    const status =
      response.getResponseCode();

    if (
      status < 200 ||
      status >= 300
    ) {
      return {
        available: false,
        error:
          'Sheets table metadata request returned HTTP ' +
          status +
          '.'
      };
    }

    const payload =
      JSON.parse(
        response.getContentText()
      );

    const sheetsById = {};

    (payload.sheets || []).forEach(function(item) {
      if (!item.properties) return;

      sheetsById[
        String(item.properties.sheetId)
      ] = {
        title:
          clean_(
            item.properties.title
          ),
        tables:
          item.tables || []
      };
    });

    return {
      available: true,
      sheetsById: sheetsById
    };
  } catch (error) {
    return {
      available: false,
      error:
        error &&
        error.message
          ? error.message
          : String(error)
    };
  }
}

function getTableCatalogForSpreadsheet_(spreadsheetId, cache) {
  const key =
    clean_(
      spreadsheetId
    );

  if (!cache[key]) {
    cache[key] =
      getSheetsTableCatalog_(key);
  }

  return cache[key];
}

function typedTableSegmentsForColumn_(catalog, sheet, zeroBasedColumn) {
  if (
    !catalog ||
    !catalog.available
  ) {
    return [];
  }

  const sheetMeta =
    catalog.sheetsById[
      String(
        sheet.getSheetId()
      )
    ];

  if (!sheetMeta) return [];

  const segments = [];

  (sheetMeta.tables || []).forEach(function(table) {
    const range =
      table.range || {};

    const startColumn =
      Number(
        range.startColumnIndex || 0
      );

    const endColumn =
      range.endColumnIndex === undefined ||
      range.endColumnIndex === null
        ? startColumn +
          (table.columnProperties || []).length
        : Number(range.endColumnIndex);

    if (
      zeroBasedColumn < startColumn ||
      zeroBasedColumn >= endColumn
    ) {
      return;
    }

    const relativeColumn =
      zeroBasedColumn -
      startColumn;

    const columnProperty =
      (table.columnProperties || [])
        .find(function(item) {
          return (
            Number(item.columnIndex) ===
            relativeColumn
          );
        }) || null;

    const columnType =
      columnProperty &&
      clean_(
        columnProperty.columnType
      );

    if (
      !columnType ||
      columnType ===
        'COLUMN_TYPE_UNSPECIFIED'
    ) {
      return;
    }

    segments.push({
      tableId:
        clean_(
          table.tableId
        ),

      tableName:
        clean_(
          table.name
        ) || 'Google Sheets table',

      columnName:
        clean_(
          columnProperty.columnName
        ),

      columnType:
        columnType,

      startRow:
        Number(
          range.startRowIndex || 0
        ) + 1,

      endRow:
        range.endRowIndex === undefined ||
        range.endRowIndex === null
          ? sheet.getLastRow()
          : Number(range.endRowIndex)
    });
  });

  return segments;
}

function rowTypedTableSegment_(segments, rowNumber) {
  for (
    let i = 0;
    i < segments.length;
    i++
  ) {
    const segment =
      segments[i];

    if (
      rowNumber >= segment.startRow &&
      rowNumber <= segment.endRow
    ) {
      return segment;
    }
  }

  return null;
}

function distinctExistingValues_(sheet, startRow, column, rowCount) {
  if (rowCount <= 0) return [];

  const seen = {};
  const output = [];

  sheet
    .getRange(
      startRow,
      column,
      rowCount,
      1
    )
    .getDisplayValues()
    .forEach(function(row) {
      const value =
        clean_(
          row[0]
        );

      if (!value) return;

      const key =
        value.toLowerCase();

      if (seen[key]) return;

      seen[key] = true;
      output.push(value);
    });

  return output;
}

function mergeValidationListValues_(baseValues, existingValues) {
  const seen = {};
  const output = [];

  (baseValues || [])
    .concat(
      existingValues || []
    )
    .forEach(function(value) {
      const cleaned =
        clean_(
          value
        );

      if (!cleaned) return;

      const key =
        cleaned.toLowerCase();

      if (seen[key]) return;

      seen[key] = true;
      output.push(cleaned);
    });

  return output;
}

function buildListValidationRuleForColumn_(
  sheet,
  structure,
  aliases,
  baseValues,
  helpText
) {
  const column =
    findColumn_(
      structure.map,
      aliases
    );

  if (column === -1) {
    return null;
  }

  const startRow =
    structure.dataStartRow;

  const lastRow =
    Math.max(
      startRow,
      sheet.getLastRow()
    );

  const rowCount =
    Math.max(
      1,
      lastRow - startRow + 1
    );

  const existingValues =
    distinctExistingValues_(
      sheet,
      startRow,
      column + 1,
      rowCount
    );

  const allowed =
    mergeValidationListValues_(
      baseValues,
      existingValues
    );

  return SpreadsheetApp
    .newDataValidation()
    .requireValueInList(
      allowed,
      true
    )
    .setAllowInvalid(
      false
    )
    .setHelpText(
      helpText
    )
    .build();
}

function applyValidationRuleToAliases_(
  sheet,
  structure,
  aliases,
  rule,
  fieldName,
  tableCatalog,
  tableInspectionRequired
) {
  const column =
    findColumn_(
      structure.map,
      aliases
    );

  if (column === -1) {
    return {
      field: fieldName,
      status: 'skipped',
      applied: false,
      skipped: true,
      failed: false,
      detail:
        'Column not present. Accepted aliases: ' +
        aliases.join(', ')
    };
  }

  const columnName =
    structure.headers[column] ||
    aliases[0];

  if (
    tableInspectionRequired &&
    (
      !tableCatalog ||
      !tableCatalog.available
    )
  ) {
    return {
      field: fieldName,
      column: columnName,
      status: 'skipped',
      applied: false,
      skipped: true,
      failed: false,
      detail:
        'Skipped to preserve Google Sheets table column types because table metadata could not be inspected' +
        (
          tableCatalog &&
          tableCatalog.error
            ? ': ' +
              tableCatalog.error
            : '.'
        )
    };
  }

  const startRow =
    structure.dataStartRow;

  const lastRow =
    Math.max(
      startRow,
      sheet.getLastRow()
    );

  const rowCount =
    Math.max(
      1,
      lastRow - startRow + 1
    );

  const typedSegments =
    typedTableSegmentsForColumn_(
      tableCatalog,
      sheet,
      column
    );

  const validations =
    sheet
      .getRange(
        startRow,
        column + 1,
        rowCount,
        1
      )
      .getDataValidations();

  const targetSignature =
    validationRuleSignature_(
      rule
    );

  const states = [];
  const typedNames = {};
  let alreadyValidRows = 0;
  let conflictingRows = 0;
  let typedRows = 0;

  for (
    let i = 0;
    i < rowCount;
    i++
  ) {
    const rowNumber =
      startRow + i;

    const typed =
      rowTypedTableSegment_(
        typedSegments,
        rowNumber
      );

    if (typed) {
      states.push('skip-typed');
      typedRows++;

      const typedKey =
        typed.tableName +
        ' / ' +
        (
          typed.columnName ||
          columnName
        ) +
        ' (' +
        typed.columnType +
        ')';

      typedNames[typedKey] = true;
      continue;
    }

    const existing =
      validations[i][0];

    if (!existing) {
      states.push('apply');
      continue;
    }

    if (
      validationRuleSignature_(
        existing
      ) ===
      targetSignature
    ) {
      states.push('already');
      alreadyValidRows++;
      continue;
    }

    states.push('skip-existing');
    conflictingRows++;
  }

  let appliedRows = 0;
  let failedRows = 0;
  const appliedRanges = [];
  const failedRanges = [];

  let runStart = -1;

  function flushRun(endExclusive) {
    if (runStart === -1) return;

    const count =
      endExclusive -
      runStart;

    const sheetRow =
      startRow +
      runStart;

    const rangeA1 =
      validationA1Range_(
        sheet,
        sheetRow,
        column + 1,
        count
      );

    try {
      sheet
        .getRange(
          sheetRow,
          column + 1,
          count,
          1
        )
        .setDataValidation(
          rule
        );

      appliedRows += count;
      appliedRanges.push(rangeA1);
    } catch (error) {
      failedRows += count;

      failedRanges.push(
        rangeA1 +
        ': ' +
        (
          error &&
          error.message
            ? error.message
            : String(error)
        )
      );
    }

    runStart = -1;
  }

  for (
    let i = 0;
    i <= states.length;
    i++
  ) {
    const shouldApply =
      i < states.length &&
      states[i] === 'apply';

    if (
      shouldApply &&
      runStart === -1
    ) {
      runStart = i;
      continue;
    }

    if (
      !shouldApply &&
      runStart !== -1
    ) {
      flushRun(i);
    }
  }

  const operationStatus =
    failedRows > 0
      ? (
          appliedRows > 0
            ? 'partial'
            : 'failed'
        )
      : (
          appliedRows > 0
            ? 'applied'
            : 'skipped'
        );

  const details = [];

  if (appliedRows) {
    details.push(
      appliedRows +
      ' row(s) validated' +
      (
        appliedRanges.length
          ? ' [' +
            appliedRanges.join(', ') +
            ']'
          : ''
      )
    );
  }

  if (alreadyValidRows) {
    details.push(
      alreadyValidRows +
      ' row(s) already had the same validation'
    );
  }

  if (typedRows) {
    details.push(
      typedRows +
      ' typed-table row(s) preserved: ' +
      Object.keys(
        typedNames
      ).join(', ')
    );
  }

  if (conflictingRows) {
    details.push(
      conflictingRows +
      ' row(s) kept an existing different validation'
    );
  }

  if (failedRows) {
    details.push(
      failedRows +
      ' row(s) failed: ' +
      failedRanges.join(' · ')
    );
  }

  if (!details.length) {
    details.push(
      'No validation change was required.'
    );
  }

  return {
    field: fieldName,
    column: columnName,
    status: operationStatus,
    applied:
      appliedRows > 0,
    skipped:
      operationStatus === 'skipped',
    failed:
      operationStatus === 'failed',
    appliedRows: appliedRows,
    alreadyValidRows: alreadyValidRows,
    typedRows: typedRows,
    conflictingRows: conflictingRows,
    failedRows: failedRows,
    detail: details.join(' · ')
  };
}

function applyTaskSheetDataValidation_(
  sheet,
  product,
  tableCatalog
) {
  const structure =
    getSourceSheetMap_(sheet);

  const tableInspectionRequired =
    [
      'Nano',
      'Klima'
    ].includes(
      canonicalProduct_(
        product
      )
    );

  const statusRule =
    buildListValidationRuleForColumn_(
      sheet,
      structure,
      structure.aliases.statusAliases,
      VALID_SOURCE_STATUSES,
      'Choose a valid task status.'
    );

  const priorityRule =
    buildListValidationRuleForColumn_(
      sheet,
      structure,
      structure.aliases.priorityAliases,
      VALID_SOURCE_PRIORITIES,
      'Choose a valid task priority.'
    );

  const dateRule =
    SpreadsheetApp
      .newDataValidation()
      .requireDate()
      .setAllowInvalid(false)
      .setHelpText(
        'Enter a valid date.'
      )
      .build();

  const pointsRule =
    SpreadsheetApp
      .newDataValidation()
      .requireNumberGreaterThanOrEqualTo(0)
      .setAllowInvalid(false)
      .setHelpText(
        'Points must be zero or greater.'
      )
      .build();

  return {
    status:
      statusRule
        ? applyValidationRuleToAliases_(
            sheet,
            structure,
            structure.aliases.statusAliases,
            statusRule,
            'Status',
            tableCatalog,
            tableInspectionRequired
          )
        : {
            field: 'Status',
            status: 'skipped',
            skipped: true,
            detail: 'Status column not present.'
          },

    priority:
      priorityRule
        ? applyValidationRuleToAliases_(
            sheet,
            structure,
            structure.aliases.priorityAliases,
            priorityRule,
            'Priority',
            tableCatalog,
            tableInspectionRequired
          )
        : {
            field: 'Priority',
            status: 'skipped',
            skipped: true,
            detail: 'Priority column not present.'
          },

    dueDate:
      applyValidationRuleToAliases_(
        sheet,
        structure,
        structure.aliases.dueDateAliases,
        dateRule,
        'Due Date',
        tableCatalog,
        tableInspectionRequired
      ),

    points:
      applyValidationRuleToAliases_(
        sheet,
        structure,
        structure.aliases.pointsAliases,
        pointsRule,
        'Points',
        tableCatalog,
        tableInspectionRequired
      )
  };
}

function summarizeValidationApplication_(applied) {
  const values =
    Object.keys(
      applied || {}
    )
    .map(function(key) {
      return applied[key];
    })
    .filter(Boolean);

  return {
    successful:
      values.filter(function(item) {
        return (
          item.status === 'applied' ||
          item.status === 'partial'
        ) &&
        Number(item.appliedRows || 0) > 0;
      }).length,

    skipped:
      values.filter(function(item) {
        return (
          item.status === 'skipped'
        );
      }).length,

    failed:
      values.filter(function(item) {
        return (
          item.status === 'failed'
        );
      }).length,

    partial:
      values.filter(function(item) {
        return (
          item.status === 'partial'
        );
      }).length,

    operations: values
  };
}

function getWeeklyValidationStructure_(sheet) {
  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        Math.max(
          16,
          sheet.getLastColumn()
        )
      )
      .getDisplayValues()[0];

  return {
    headerRow: 1,
    dataStartRow: 2,
    lastColumn: headers.length,
    headers: headers,
    map:
      buildHeaderMap_(
        headers
      )
  };
}

function applyWeeklySheetDataValidation_(sheet) {
  const structure =
    getWeeklyValidationStructure_(sheet);

  const statusRule =
    buildListValidationRuleForColumn_(
      sheet,
      structure,
      ['Weekly Status'],
      VALID_TASK_STATUSES,
      'Choose a valid weekly status.'
    );

  const priorityRule =
    buildListValidationRuleForColumn_(
      sheet,
      structure,
      ['Priority'],
      VALID_TASK_PRIORITIES,
      'Choose P0, P1, P2, P3, or No priority.'
    );

  const commitmentRule =
    buildListValidationRuleForColumn_(
      sheet,
      structure,
      ['Commitment Level'],
      VALID_WEEKLY_COMMITMENTS,
      'Choose Commit, Stretch, or Watch / Unblock.'
    );

  const dateRule =
    SpreadsheetApp
      .newDataValidation()
      .requireDate()
      .setAllowInvalid(false)
      .setHelpText(
        'Enter a valid date.'
      )
      .build();

  const numberRule =
    SpreadsheetApp
      .newDataValidation()
      .requireNumberGreaterThanOrEqualTo(0)
      .setAllowInvalid(false)
      .setHelpText(
        'Enter zero or a positive number.'
      )
      .build();

  return {
    priority:
      priorityRule
        ? applyValidationRuleToAliases_(
            sheet,
            structure,
            ['Priority'],
            priorityRule,
            'Priority',
            null,
            false
          )
        : {
            field: 'Priority',
            status: 'skipped',
            skipped: true,
            detail: 'Priority column not present.'
          },

    status:
      statusRule
        ? applyValidationRuleToAliases_(
            sheet,
            structure,
            ['Weekly Status'],
            statusRule,
            'Weekly Status',
            null,
            false
          )
        : {
            field: 'Weekly Status',
            status: 'skipped',
            skipped: true,
            detail: 'Weekly Status column not present.'
          },

    dueDate:
      applyValidationRuleToAliases_(
        sheet,
        structure,
        ['Due Date'],
        dateRule,
        'Due Date',
        null,
        false
      ),

    commitment:
      commitmentRule
        ? applyValidationRuleToAliases_(
            sheet,
            structure,
            ['Commitment Level'],
            commitmentRule,
            'Commitment Level',
            null,
            false
          )
        : {
            field: 'Commitment Level',
            status: 'skipped',
            skipped: true,
            detail: 'Commitment Level column not present.'
          },

    points:
      applyValidationRuleToAliases_(
        sheet,
        structure,
        ['Points'],
        numberRule,
        'Points',
        null,
        false
      ),

    estimatedHours:
      applyValidationRuleToAliases_(
        sheet,
        structure,
        ['Estimated Hours'],
        numberRule,
        'Estimated Hours',
        null,
        false
      )
  };
}

function validationSheetStatus_(summary) {
  if (
    summary.failed > 0 &&
    summary.successful === 0
  ) {
    return 'failed';
  }

  if (
    summary.failed > 0 ||
    summary.partial > 0
  ) {
    return 'partial';
  }

  if (
    summary.successful > 0
  ) {
    return 'applied';
  }

  return 'skipped';
}

function validationSheetDetail_(summary) {
  const pieces = [];

  if (summary.successful) {
    pieces.push(
      summary.successful +
      ' successful operation(s)'
    );
  }

  if (summary.skipped) {
    pieces.push(
      summary.skipped +
      ' skipped operation(s)'
    );
  }

  if (summary.failed) {
    pieces.push(
      summary.failed +
      ' failed operation(s)'
    );
  }

  if (!pieces.length) {
    pieces.push(
      'No validation operation was required.'
    );
  }

  return pieces.join(' · ');
}

function applySourceDataValidation() {
  const actor =
    requireAdmin_();

  const dashboard =
    getDashboardSpreadsheet_();

  const results = [];
  const seen = {};
  const tableCatalogCache = {};

  Object.keys(TASK_ROUTING)
    .forEach(function(product) {
      TASK_ROUTING[product]
        .forEach(function(route) {
          const key =
            clean_(
              route.spreadsheetId
            ) +
            '::' +
            clean_(
              route.sourceSheet
            ).toLowerCase();

          if (seen[key]) return;

          seen[key] = true;

          try {
            const book =
              route.spreadsheetId
                ? openSpreadsheetWithRetry_(
                    route.spreadsheetId
                  )
                : dashboard;

            const sheet =
              book.getSheetByName(
                route.sourceSheet
              );

            if (!sheet) {
              results.push({
                product:
                  displayProductName_(
                    product
                  ),
                workbook:
                  book.getName(),
                sheet:
                  route.sourceSheet,
                status: 'failed',
                success: false,
                operations: [],
                detail:
                  'Configured source sheet was not found.'
              });

              return;
            }

            const spreadsheetId =
              book.getId();

            const needsTypedTableInspection =
              [
                'Nano',
                'Klima'
              ].includes(
                canonicalProduct_(
                  product
                )
              );

            const tableCatalog =
              needsTypedTableInspection
                ? getTableCatalogForSpreadsheet_(
                    spreadsheetId,
                    tableCatalogCache
                  )
                : {
                    available: true,
                    sheetsById: {}
                  };

            const applied =
              applyTaskSheetDataValidation_(
                sheet,
                product,
                tableCatalog
              );

            const summary =
              summarizeValidationApplication_(
                applied
              );

            const status =
              validationSheetStatus_(
                summary
              );

            results.push({
              product:
                displayProductName_(
                  product
                ),
              workbook:
                book.getName(),
              sheet:
                route.sourceSheet,
              status: status,
              success:
                status !== 'failed',
              operations:
                summary.operations,
              applied: applied,
              detail:
                validationSheetDetail_(
                  summary
                )
            });
          } catch (error) {
            results.push({
              product:
                displayProductName_(
                  product
                ),
              workbook:
                clean_(
                  route.spreadsheetId
                ),
              sheet:
                route.sourceSheet,
              status: 'failed',
              success: false,
              operations: [],
              detail:
                error &&
                error.message
                  ? error.message
                  : String(error)
            });
          }
        });
    });

  try {
    const weeklySheet =
      ensureWeeklySheet_(
        dashboard
      );

    const applied =
      applyWeeklySheetDataValidation_(
        weeklySheet
      );

    const summary =
      summarizeValidationApplication_(
        applied
      );

    const status =
      validationSheetStatus_(
        summary
      );

    results.push({
      product: 'Weekly Plan',
      workbook:
        dashboard.getName(),
      sheet:
        CONFIG.WEEKLY_SHEET,
      status: status,
      success:
        status !== 'failed',
      operations:
        summary.operations,
      applied: applied,
      detail:
        validationSheetDetail_(
          summary
        )
    });
  } catch (error) {
    results.push({
      product: 'Weekly Plan',
      workbook:
        dashboard.getName(),
      sheet:
        CONFIG.WEEKLY_SHEET,
      status: 'failed',
      success: false,
      operations: [],
      detail:
        error &&
        error.message
          ? error.message
          : String(error)
    });
  }

  const successfulOperations = [];
  const skippedOperations = [];
  const failedOperations = [];

  results.forEach(function(result) {
    (result.operations || [])
      .forEach(function(operation) {
        const item = {
          product:
            result.product,
          workbook:
            result.workbook,
          sheet:
            result.sheet,
          field:
            operation.field,
          column:
            operation.column || '',
          status:
            operation.status,
          detail:
            operation.detail
        };

        if (
          operation.status === 'failed'
        ) {
          failedOperations.push(item);
          return;
        }

        if (
          operation.status === 'partial'
        ) {
          successfulOperations.push(item);
          failedOperations.push(
            Object.assign(
              {},
              item,
              {
                detail:
                  'Partial: ' +
                  item.detail
              }
            )
          );
          return;
        }

        if (
          operation.status === 'applied'
        ) {
          successfulOperations.push(item);
          return;
        }

        skippedOperations.push(item);
      });
  });

  const summary = {
    applied:
      results.filter(function(item) {
        return item.status === 'applied';
      }).length,

    partial:
      results.filter(function(item) {
        return item.status === 'partial';
      }).length,

    skipped:
      results.filter(function(item) {
        return item.status === 'skipped';
      }).length,

    failed:
      results.filter(function(item) {
        return item.status === 'failed';
      }).length,

    total:
      results.length,

    successfulOperations:
      successfulOperations.length,

    skippedOperations:
      skippedOperations.length,

    failedOperations:
      failedOperations.length
  };

  logActivities_([
    {
      action:
        'Applied data validation',
      entityType:
        'System',
      details:
        summary.successfulOperations +
        ' successful · ' +
        summary.skippedOperations +
        ' skipped · ' +
        summary.failedOperations +
        ' failed validation operation(s)'
    }
  ], actor);

  return {
    success:
      failedOperations.length === 0,

    summary:
      summary,

    operations: {
      successful:
        successfulOperations,
      skipped:
        skippedOperations,
      failed:
        failedOperations
    },

    results:
      results
  };
}


/* ============================================================
   DASHBOARD
============================================================ */

function isRecognizedStatus_(value) {
  const s = clean_(value).toLowerCase();
  if (!s) return false;

  if ([
    'done','complete','completed','resolved','passed','closed',
    'backlog','none','not started','not-started','not_started','todo','to do','active','doing','qa','test',
    'blocked','failed'
  ].includes(s)) return true;

  return (
    s.includes('progress') ||
    s.includes('review') ||
    s.includes('testing') ||
    s.includes('retest') ||
    s.includes('awaiting tech fix')
  );
}

function buildDataValidationReport_(tasks, weeklyPlan) {
  tasks = Array.isArray(tasks) ? tasks : [];
  weeklyPlan = Array.isArray(weeklyPlan) ? weeklyPlan : [];

  const issues = [];
  const seenIds = {};
  const masterIds = new Set();

  function add(severity, type, task, message) {
    issues.push({
      severity: severity,
      type: type,
      taskId: task && task.taskId ? task.taskId : '',
      title: task && task.title ? task.title : '',
      product: task && task.product ? task.product : '',
      message: message
    });
  }

  tasks.forEach(function(task) {
    const idKey = clean_(task.taskId).toLowerCase();
    if (idKey) {
      masterIds.add(idKey);
      seenIds[idKey] = (seenIds[idKey] || 0) + 1;
    }

    if (!clean_(task.title)) {
      add('error', 'Missing title', task, 'Task title is blank.');
    }

    if (!isKnownProduct_(task.product)) {
      add(
        'error',
        'Unmapped product',
        task,
        'Product "' + clean_(task.rawProduct || task.product || 'blank') + '" is not mapped to a configured product.'
      );
    } else {
      const routes = getProductRoutes_(task.product);
      const source = clean_(task.sourceSheet).toLowerCase();
      if (
        source &&
        !routes.some(function(route) {
          return clean_(route.sourceSheet).toLowerCase() === source;
        })
      ) {
        add(
          'error',
          'Routing mismatch',
          task,
          'Source sheet "' + task.sourceSheet + '" does not match the configured route for ' + task.product + '.'
        );
      }
    }

    if (!clean_(task.owner) || clean_(task.owner).toLowerCase() === 'unassigned') {
      add('warning', 'Missing owner', task, 'Task has no assigned owner.');
    }

    const priorityRaw = clean_(task.priority);
    const priorityNeutral = [
      '',
      'no priority',
      'none',
      'n/a',
      'na',
      'not set',
      'unset'
    ].includes(priorityRaw.toLowerCase());

    // Completed work does not need priority cleanup.
    // For active work, a neutral/missing priority is a hygiene warning,
    // but it is not an "invalid priority" error.
    if (!isClosedStatus_(task.status)) {
      if (priorityNeutral) {
        add('warning', 'Missing priority', task, 'Active task has no priority set.');
      } else if (!VALID_TASK_PRIORITIES.includes(normalizePriority_(task.priority))) {
        add('warning', 'Invalid priority', task, 'Priority "' + priorityRaw + '" is not recognized.');
      }
    }

    if (!isRecognizedStatus_(task.status)) {
      add('warning', 'Invalid status', task, 'Status "' + clean_(task.status) + '" is not recognized.');
    }

    const due = clean_(task.dueDate);
    if (due && !/^\d{4}-\d{2}-\d{2}$/.test(due)) {
      add('warning', 'Invalid due date', task, 'Due date "' + due + '" is not in YYYY-MM-DD format.');
    }

    if (task.points !== '' && task.points !== null && task.points !== undefined) {
      const points = Number(task.points);
      if (!Number.isFinite(points) || points < 0) {
        add('warning', 'Invalid points', task, 'Points must be a non-negative number.');
      }
    }
  });

  Object.keys(seenIds).forEach(function(idKey) {
    if (seenIds[idKey] > 1) {
      const task = tasks.find(function(item) {
        return clean_(item.taskId).toLowerCase() === idKey;
      });
      add('error', 'Duplicate task ID', task, 'Task ID appears ' + seenIds[idKey] + ' times in the configured source tasks.');
    }
  });

  const taskStatusById = {};
  tasks.forEach(function(task) {
    const key = clean_(task.taskId).toLowerCase();
    if (key) taskStatusById[key] = task.status;
  });

  const weeklySeen = {};
  weeklyPlan.forEach(function(item) {
    const idKey = clean_(item.taskId).toLowerCase();
    if (!idKey || !masterIds.has(idKey)) {
      add('error', 'Orphan weekly task', item, 'Weekly plan task is not present in the configured source tasks.');
    }

    if (idKey) {
      weeklySeen[idKey] = (weeklySeen[idKey] || 0) + 1;
    }

    const estimateStatus =
      taskStatusById[idKey] ||
      item.weeklyStatus;

    if (
      statusUsesEstimatedHours_(estimateStatus) &&
      (
        item.estimatedHours === '' ||
        item.estimatedHours === null ||
        item.estimatedHours === undefined ||
        !Number.isFinite(Number(item.estimatedHours)) ||
        Number(item.estimatedHours) <= 0
      )
    ) {
      add(
        'warning',
        'Missing estimate',
        item,
        'Active development task has no estimated hours.'
      );
    }

    const rawCommitment = clean_(item.commitmentLevel);
    if (
      rawCommitment &&
      !VALID_WEEKLY_COMMITMENTS.some(function(value) {
        return value.toLowerCase() === rawCommitment.toLowerCase();
      })
    ) {
      add('warning', 'Invalid commitment', item, 'Weekly commitment "' + rawCommitment + '" is not recognized.');
    }
  });

  Object.keys(weeklySeen).forEach(function(idKey) {
    if (weeklySeen[idKey] > 1) {
      const item = weeklyPlan.find(function(row) {
        return clean_(row.taskId).toLowerCase() === idKey;
      });
      add('error', 'Duplicate weekly task', item, 'Task appears more than once in the selected week.');
    }
  });

  const byType = {};
  issues.forEach(function(issue) {
    byType[issue.type] = (byType[issue.type] || 0) + 1;
  });

  return {
    totalIssues: issues.length,
    errorCount: issues.filter(function(issue) { return issue.severity === 'error'; }).length,
    warningCount: issues.filter(function(issue) { return issue.severity === 'warning'; }).length,
    byType: byType,
    issues: issues.slice(0, 100)
  };
}

function getSourceSchemaDiagnostics() {
  requireAdmin_();

  const rows = [];
  const seen = {};

  Object.keys(TASK_ROUTING).forEach(function(product) {
    TASK_ROUTING[product].forEach(function(route) {
      const key = clean_(route.spreadsheetId) + '::' + clean_(route.sourceSheet);
      if (seen[key]) return;
      seen[key] = true;

      try {
        const resolved = {
          product: product,
          workstream: route.workstream,
          sourceSheet: route.sourceSheet,
          spreadsheetId: route.spreadsheetId || ''
        };
        const sheet = getTargetSheet_(resolved);
        const structure = getSourceSheetMap_(sheet);

        rows.push({
          product: product,
          sourceSheet: route.sourceSheet,
          spreadsheetId: route.spreadsheetId || '',
          ok: true,
          headerRow: structure.headerRow,
          idHeader: structure.headers[structure.idColumn] || '',
          titleHeader: structure.headers[structure.titleColumn] || '',
          headers: structure.headers.filter(Boolean)
        });
      } catch (error) {
        rows.push({
          product: product,
          sourceSheet: route.sourceSheet,
          spreadsheetId: route.spreadsheetId || '',
          ok: false,
          error: error && error.message ? error.message : String(error)
        });
      }
    });
  });

  return rows;
}

/* ============================================================
   GENERATE TASK ID
============================================================ */
