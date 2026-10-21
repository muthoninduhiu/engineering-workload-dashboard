function displayProductLabelForServer_(product) {
  const canonical =
    canonicalProduct_(product);

  if (canonical === 'MediaPal') return 'Mediapal';
  if (canonical === 'WePlay') return 'Weplay';

  return canonical || 'Unknown';
}

function canonicalProduct_(value) {
  const raw = clean_(value);
  const key = raw.toLowerCase().replace(/[^a-z0-9]/g, '');

  if (!key) return '';

  // Keep stable internal keys for routing. Display capitalization is handled
  // separately in Index.html (Mediapal / Weplay).
  if (key === 'mediapal' || key === 'mediapalgroup' || key === 'mediapaldsp') {
    return 'MediaPal';
  }

  if (key === 'weplay' || key === 'weplayarcade' || key === 'weplaygames') {
    return 'WePlay';
  }

  if (key === 'nano' || key === 'nanodigital') return 'Nano';
  if (key === 'showapp' || key === 'showappke') return 'ShowApp';
  if (key === 'klima' || key === 'klimapp') return 'Klima';
  if (key === 'governance' || key === 'documentationgovernance') return 'Governance';

  return raw;
}

function isKnownProduct_(value) {
  const product = canonicalProduct_(value);
  return Object.prototype.hasOwnProperty.call(TASK_ROUTING, product);
}

function displayProductName_(value) {
  const product = canonicalProduct_(value);
  if (product === 'MediaPal') return 'Mediapal';
  if (product === 'WePlay') return 'Weplay';
  return product || clean_(value);
}

function inferSharedLinearProductFromTaskId_(taskId) {
  const id = clean_(taskId).toUpperCase();

  if (/^(WP|WEP)-\d+$/i.test(id)) return 'WePlay';
  if (/^MED-\d+$/i.test(id)) return 'MediaPal';

  return '';
}
function resolveTaskProduct_(sourceProduct, product, taskId) {
  const source = canonicalProduct_(sourceProduct);
  const primary = canonicalProduct_(product);
  const idProduct = inferSharedLinearProductFromTaskId_(taskId);

  // MediaPal and WePlay use the same Consolidated Linear Tasks source.
  // WP-* / MED-* IDs disambiguate tasks in the shared Linear source sheet.
  if (idProduct && (
    !source ||
    source === 'MediaPal' ||
    source === 'WePlay'
  )) {
    return idProduct;
  }

  if (isKnownProduct_(source)) return source;
  if (isKnownProduct_(primary)) return primary;
  if (idProduct) return idProduct;

  return source || primary || 'General';
}

function getProductRoutes_(product) {
  const canonical = canonicalProduct_(product);
  return TASK_ROUTING[canonical] || [];
}

/**
 * Google occasionally returns a transient "Service Drive timed out" error
 * while opening a spreadsheet. Retrying only that class of failure avoids
 * turning a momentary Drive delay into a failed app request, without hiding
 * genuine sharing or configuration errors.
 */
function openSpreadsheetWithRetry_(spreadsheetId) {
  const id = clean_(spreadsheetId);
  const attempts = 3;
  let lastError;

  for (let attempt = 0; attempt < attempts; attempt += 1) {
    try {
      return SpreadsheetApp.openById(id);
    } catch (error) {
      lastError = error;
      const message = String(
        error && error.message ? error.message : error
      ).toLowerCase();
      const isTransient =
        message.indexOf('service drive timed out') !== -1 ||
        message.indexOf('service unavailable') !== -1 ||
        message.indexOf('internal error') !== -1;

      if (!isTransient || attempt === attempts - 1) break;

      // Short exponential backoff: 500 ms, then 1 s.
      Utilities.sleep(500 * Math.pow(2, attempt));
    }
  }

  throw lastError;
}

function resolveRoute_(product, workstream, sourceSheet) {
  const canonicalProduct = canonicalProduct_(product);
  const routes = getProductRoutes_(canonicalProduct);

  if (!routes.length) {
    throw new Error(
      'No task routing has been configured for "' + product + '".'
    );
  }

  const requestedSheet = clean_(sourceSheet).toLowerCase();
  const requestedWorkstream = clean_(workstream).toLowerCase();

  let route = null;

  if (requestedSheet) {
    route = routes.find(function(item) {
      return clean_(item.sourceSheet).toLowerCase() === requestedSheet;
    });
  }

  if (!route && requestedWorkstream) {
    route = routes.find(function(item) {
      return (
        clean_(item.workstream).toLowerCase() === requestedWorkstream ||
        clean_(item.sourceSheet).toLowerCase() === requestedWorkstream
      );
    });
  }

  if (!route && routes.length === 1) {
    route = routes[0];
  }

  if (!route) {
    throw new Error(
      'Choose a valid target sheet for ' + canonicalProduct + '.'
    );
  }

  return {
    product: canonicalProduct,
    workstream: route.workstream,
    sourceSheet: route.sourceSheet,
    spreadsheetId: route.spreadsheetId || ''
  };
}

function getTargetSheet_(route) {
  if (!route) {
    throw new Error('No target route was provided.');
  }

  const spreadsheetId = clean_(route.spreadsheetId);

  let spreadsheet;

  try {
    spreadsheet = spreadsheetId
      ? openSpreadsheetWithRetry_(spreadsheetId)
      : getDashboardSpreadsheet_();
  } catch (error) {
    throw buildAccessDeniedError_(
      'SOURCE',
      displayProductLabelForServer_(
        route.product || 'this product'
      ) + ' source workbook',
      spreadsheetId,
      'Editor access to modify tasks',
      error
    );
  }

  const sheet = spreadsheet.getSheetByName(route.sourceSheet);

  if (!sheet) {
    throw new Error(
      'Could not find "' +
      route.sourceSheet +
      '" for ' +
      (route.product || 'the selected product') +
      ' / ' +
      (route.workstream || 'the selected workstream') +
      '.'
    );
  }

  return sheet;
}

function classifyWorkstream_(product, sourceSheet, existingWorkstream) {
  const routes = getProductRoutes_(product);
  const source = clean_(sourceSheet);
  const existing = clean_(existingWorkstream);

  if (routes.length) {
    const sourceMatch = routes.find(function(route) {
      return (
        clean_(route.sourceSheet).toLowerCase() === source.toLowerCase()
      );
    });

    if (sourceMatch) {
      return sourceMatch.workstream;
    }

    const workstreamMatch = routes.find(function(route) {
      return (
        clean_(route.workstream).toLowerCase() === existing.toLowerCase()
      );
    });

    if (workstreamMatch) {
      return workstreamMatch.workstream;
    }
  }

  // Existing Handover Requirements etc. can still be read.
  return existing || source || canonicalProduct_(product);
}

/* ============================================================
   WEEK HELPERS
============================================================ */

function getRoutingConfig_() {
  const access = getAccessContext_();
  const order = V6_PRODUCTS.filter(function(product) { return canAccessProduct_(access, product); });
  return order.map(function(product) {
    return {product:product, workstreams:(TASK_ROUTING[product]||[]).map(function(route){return {label:route.workstream,sourceSheet:route.sourceSheet};})};
  });
}
