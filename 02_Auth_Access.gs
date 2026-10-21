function saveAccessRegistry_(registry) {
  PropertiesService
    .getScriptProperties()
    .setProperty(
      CONFIG.ACCESS_PROPERTY_KEY,
      JSON.stringify(registry || {})
    );
}

function getAccessDiagnostics() {
  const access = getAccessContext_();
  const resources = [];
  const seen = {};

  const dashboardId =
    getConfiguredDashboardId_();

  resources.push({
    key: 'dashboard',
    label: CONFIG.DASHBOARD_LABEL,
    spreadsheetId: dashboardId,
    requiredPermission: 'Viewer access or higher'
  });

  Object.keys(TASK_ROUTING).forEach(function(product) {
    (TASK_ROUTING[product] || []).forEach(function(route) {
      const id = clean_(route.spreadsheetId);

      if (!id) return;

      if (!seen[id]) {
        seen[id] = {
          key: 'source-' + id,
          label: displayProductLabelForServer_(product) + ' source workbook',
          spreadsheetId: id,
          requiredPermission:
            access.canEdit
              ? 'Editor access to modify tasks'
              : 'Viewer access to inspect tasks'
        };

        resources.push(
          seen[id]
        );
      } else {
        const existing = seen[id];
        const productLabel =
          displayProductLabelForServer_(product);

        if (
          existing.label.indexOf(
            productLabel
          ) === -1
        ) {
          existing.label =
            existing.label.replace(
              ' source workbook',
              ''
            ) +
            ' / ' +
            productLabel +
            ' source workbook';
        }
      }
    });
  });

  const results = resources.map(function(resource) {
    if (!resource.spreadsheetId) {
      return {
        key: resource.key,
        label: resource.label,
        spreadsheetId: '',
        accessible: false,
        workbookName: '',
        requiredPermission: resource.requiredPermission,
        error: 'Spreadsheet ID is not configured.'
      };
    }

    try {
      const book =
        openSpreadsheetWithRetry_(
          resource.spreadsheetId
        );

      return {
        key: resource.key,
        label: resource.label,
        spreadsheetId: resource.spreadsheetId,
        accessible: true,
        workbookName: book.getName(),
        requiredPermission: resource.requiredPermission,
        error: ''
      };
    } catch (error) {
      return {
        key: resource.key,
        label: resource.label,
        spreadsheetId: resource.spreadsheetId,
        accessible: false,
        workbookName: '',
        requiredPermission: resource.requiredPermission,
        error:
          error && error.message
            ? clean_(error.message)
            : String(error)
      };
    }
  });

  return {
    access: access,
    deployment: {
      expectedExecuteAs: 'User accessing the web app',
      expectedWhoHasAccess: 'Anyone with Google account'
    },
    resources: results
  };
}

function normalizeRole_(role) {
  const value = clean_(role).toLowerCase();
  if (value === 'admin') return 'Admin';
  if (value === 'executive') return 'Executive';
  if (value === 'editor') return 'Editor';
  if (value === 'viewer') return 'Viewer';
  return 'No access';
}

function normalizeProductList_(products, role) {
  role = normalizeRole_(role);
  if (role === 'Admin' || role === 'Executive') return V6_PRODUCTS.slice();
  const input = Array.isArray(products) ? products : [];
  const seen = {};
  return input.map(canonicalProduct_).filter(function(p) {
    if (V6_PRODUCTS.indexOf(p) === -1 || seen[p]) return false;
    seen[p] = true;
    return true;
  });
}

function getAccessRegistry_() {
  const properties = PropertiesService.getScriptProperties();
  const raw = properties.getProperty(CONFIG.ACCESS_PROPERTY_KEY);
  let registry = {};
  let changed = false;
  if (raw) {
    try { registry = JSON.parse(raw) || {}; }
    catch (error) { registry = {}; changed = true; }
  }

  const adminEmail = clean_(CONFIG.DEFAULT_ADMIN_EMAIL).toLowerCase();
  if (!registry[adminEmail]) {
    registry[adminEmail] = {name:'Maryanne', role:'Admin', active:true, products:V6_PRODUCTS.slice()};
    changed = true;
  }

  Object.keys(registry).forEach(function(email) {
    const entry = registry[email] || {};
    const role = normalizeRole_(entry.role);
    const normalized = {
      name: clean_(entry.name) || email,
      role: role === 'No access' ? 'Viewer' : role,
      active: entry.active !== false,
      products: normalizeProductList_(entry.products, role === 'No access' ? 'Viewer' : role)
    };
    if (JSON.stringify(entry) !== JSON.stringify(normalized)) changed = true;
    registry[email] = normalized;
  });

  const forcedAdmin = {name:clean_(registry[adminEmail].name)||'Maryanne',role:'Admin',active:true,products:V6_PRODUCTS.slice()};
  if (JSON.stringify(registry[adminEmail]) !== JSON.stringify(forcedAdmin)) changed = true;
  registry[adminEmail] = forcedAdmin;
  if (changed || !raw) saveAccessRegistry_(registry);
  return registry;
}

function buildAccessContextForEmail_(email, fallbackName) {
  email = clean_(email).toLowerCase();
  const registry = getAccessRegistry_();
  const entry = email ? (registry[email] || null) : null;

  if (!email) {
    return {identified:false, verified:false, email:'', name:'Guest', role:'No access', active:false,
      canView:false, canEdit:false, canAdmin:false, canManageUsers:false, products:[], allProducts:false,
      needsApproval:false};
  }

  if (!entry || entry.active === false) {
    return {identified:true, verified:true, email:email, name:clean_(fallbackName)||email, role:'No access', active:false,
      canView:false, canEdit:false, canAdmin:false, canManageUsers:false, products:[], allProducts:false,
      needsApproval:true};
  }

  const role = normalizeRole_(entry.role);
  const products = normalizeProductList_(entry.products, role);
  const allProducts = role === 'Admin' || role === 'Executive';
  return {
    identified:true,
    verified:true,
    email:email,
    name:clean_(entry.name)||clean_(fallbackName)||email,
    role:role,
    active:true,
    canView:true,
    canEdit:role === 'Admin' || role === 'Editor',
    canAdmin:role === 'Admin',
    canManageUsers:role === 'Admin',
    products:allProducts ? V6_PRODUCTS.slice() : products,
    allProducts:allProducts,
    needsApproval:false
  };
}

function getCurrentUserEmail_() {
  if (REQUEST_AUTH_CONTEXT_ && REQUEST_AUTH_CONTEXT_.email) return REQUEST_AUTH_CONTEXT_.email;
  try {
    const active = clean_(Session.getActiveUser().getEmail()).toLowerCase();
    if (active) return active;
  } catch (error) {}
  try {
    return clean_(Session.getEffectiveUser().getEmail()).toLowerCase();
  } catch (error) { return ''; }
}

function getAccessContext_() {
  if (REQUEST_AUTH_CONTEXT_) return REQUEST_AUTH_CONTEXT_;
  return buildAccessContextForEmail_(getCurrentUserEmail_(), '');
}

function canAccessProduct_(access, product) {
  access = access || getAccessContext_();
  const canonical = canonicalProduct_(product);
  if (!access || !access.canView) return false;
  if (access.allProducts || access.role === 'Admin' || access.role === 'Executive') return true;
  return (access.products || []).indexOf(canonical) !== -1;
}

function requireViewer_() {
  const access = getAccessContext_();
  if (!access.verified || !access.canView) {
    throw new Error('APP_NOT_AUTHORIZED||Your verified account is not approved for this workspace.');
  }
  return access;
}

function requireEditor_() {
  const access = requireViewer_();
  if (!access.canEdit) throw new Error('You have view-only access to this workspace.');
  return access;
}

function requireAdmin_() {
  const access = requireViewer_();
  if (!access.canAdmin) throw new Error('Admin access is required for this action.');
  return access;
}

function requireProductAccess_(access, product, editRequired) {
  const canonical = canonicalProduct_(product);
  if (!canAccessProduct_(access, canonical)) {
    throw new Error('You do not have access to ' + displayProductLabelForServer_(canonical) + '.');
  }
  if (editRequired && !access.canEdit) throw new Error('You have view-only access to this workspace.');
  return canonical;
}

function sessionPropertyKey_(token) {
  return V6_AUTH.SESSION_PROPERTY_PREFIX + clean_(token);
}

function persistAppSession_(token, payload) {
  if (!token || !payload) return;

  try {
    PropertiesService.getScriptProperties()
      .setProperty(sessionPropertyKey_(token), JSON.stringify(payload));
  } catch (error) {
    console.warn('Could not persist app session.', error);
  }

  try {
    CacheService.getScriptCache().put(
      V6_AUTH.SESSION_PREFIX + token,
      JSON.stringify(payload),
      V6_AUTH.SESSION_TTL_SECONDS
    );
  } catch (error) {
    console.warn('Could not cache app session.', error);
  }
}

function issueAppSession_(email, name) {
  const access = buildAccessContextForEmail_(email, name);
  const token =
    Utilities.getUuid().replace(/-/g,'') +
    Utilities.getUuid().replace(/-/g,'');

  const now = Date.now();
  const payload = {
    email: access.email,
    name: access.name,
    issuedAt: now,
    lastSeenAt: now,
    expiresAt: now + V6_AUTH.SESSION_PERSIST_DAYS * 86400000
  };

  persistAppSession_(token, payload);

  return {
    sessionToken: token,
    access: access,
    expiresIn: V6_AUTH.SESSION_PERSIST_DAYS * 86400
  };
}

function readAppSession_(token) {
  token = clean_(token);
  if (!token) {
    throw new Error('APP_SESSION_REQUIRED||Sign in to continue.');
  }

  let raw = '';

  try {
    raw = CacheService.getScriptCache()
      .get(V6_AUTH.SESSION_PREFIX + token) || '';
  } catch (error) {}

  if (!raw) {
    try {
      raw = PropertiesService.getScriptProperties()
        .getProperty(sessionPropertyKey_(token)) || '';
    } catch (error) {}
  }

  if (!raw) {
    throw new Error('APP_SESSION_EXPIRED||Your saved app session is no longer available. Verify your email once to continue.');
  }

  let payload;
  try {
    payload = JSON.parse(raw);
  } catch (error) {
    PropertiesService.getScriptProperties().deleteProperty(sessionPropertyKey_(token));
    throw new Error('APP_SESSION_EXPIRED||Your saved app session is invalid. Verify your email again.');
  }

  const now = Date.now();
  if (!payload.expiresAt) {
    // Migrate an existing V6 cache-only session into persistent storage.
    payload.expiresAt = now + V6_AUTH.SESSION_PERSIST_DAYS * 86400000;
  }

  if (Number(payload.expiresAt) <= now) {
    try {
      CacheService.getScriptCache().remove(V6_AUTH.SESSION_PREFIX + token);
    } catch (error) {}
    PropertiesService.getScriptProperties().deleteProperty(sessionPropertyKey_(token));
    throw new Error('APP_SESSION_EXPIRED||Your 30-day app session expired. Verify your email again.');
  }

  // Rolling 30-day session while the user remains active.
  payload.lastSeenAt = now;
  payload.expiresAt = now + V6_AUTH.SESSION_PERSIST_DAYS * 86400000;
  persistAppSession_(token, payload);

  return buildAccessContextForEmail_(payload.email, payload.name);
}

function getPublicBootstrap() {
  let platformEmail = '';
  try { platformEmail = clean_(Session.getActiveUser().getEmail()).toLowerCase(); } catch (error) {}
  let automaticSession = null;
  if (platformEmail) automaticSession = issueAppSession_(platformEmail, '');
  return {
    appName:'Engineering Workload',
    platformEmail:platformEmail,
    automaticSession:automaticSession,
    deployment:{expectedExecuteAs:'Me (app owner)', expectedWhoHasAccess:'Anyone with Google account'},
    identityMode:platformEmail ? 'google-account' : 'email-verification'
  };
}

function startEmailVerification(email) {
  email = clean_(email).toLowerCase();
  if (!email || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('Enter a valid email address.');
  const cache = CacheService.getScriptCache();
  const rateKey = V6_AUTH.OTP_RATE_PREFIX + Utilities.base64EncodeWebSafe(email).slice(0,80);
  if (cache.get(rateKey)) throw new Error('A verification code was sent recently. Wait a minute before requesting another.');
  cache.put(rateKey, '1', 60);

  const code = String(Math.floor(100000 + Math.random() * 900000));
  const verificationId = Utilities.getUuid();
  cache.put(V6_AUTH.OTP_PREFIX + verificationId, JSON.stringify({email:email, code:code, attempts:0}), V6_AUTH.OTP_TTL_SECONDS);

  MailApp.sendEmail({
    to:email,
    subject:'Engineering Workload verification code',
    htmlBody:'<p>Your Engineering Workload verification code is:</p><p style="font-size:28px;font-weight:700;letter-spacing:6px">' + code + '</p><p>This code expires in 10 minutes.</p>',
    name:'Engineering Workload'
  });

  return {verificationId:verificationId, email:email, expiresIn:V6_AUTH.OTP_TTL_SECONDS};
}

function completeEmailVerification(verificationId, code) {
  verificationId = clean_(verificationId);
  code = clean_(code);
  const cache = CacheService.getScriptCache();
  const key = V6_AUTH.OTP_PREFIX + verificationId;
  const raw = cache.get(key);
  if (!raw) throw new Error('That verification code expired. Request a new one.');
  const payload = JSON.parse(raw);
  payload.attempts = Number(payload.attempts || 0) + 1;
  if (payload.attempts > 6) { cache.remove(key); throw new Error('Too many incorrect attempts. Request a new code.'); }
  if (code !== payload.code) { cache.put(key, JSON.stringify(payload), V6_AUTH.OTP_TTL_SECONDS); throw new Error('Incorrect verification code.'); }
  cache.remove(key);
  return issueAppSession_(payload.email, '');
}

function logoutAppSession(token) {
  token = clean_(token);

  if (token) {
    try {
      CacheService.getScriptCache().remove(V6_AUTH.SESSION_PREFIX + token);
    } catch (error) {}

    try {
      PropertiesService.getScriptProperties()
        .deleteProperty(sessionPropertyKey_(token));
    } catch (error) {}
  }

  return {success:true};
}

function getBootstrapInfo() {
  const access = getAccessContext_();
  return {
    access:access,
    deployment:{expectedExecuteAs:'Me (app owner)', expectedWhoHasAccess:'Anyone with Google account'},
    dashboard:{label:CONFIG.DASHBOARD_LABEL, spreadsheetId:getConfiguredDashboardId_()}
  };
}

function getAccessRequests_() {
  const raw = PropertiesService.getScriptProperties().getProperty(V6_AUTH.ACCESS_REQUESTS_PROPERTY_KEY);
  if (!raw) return {};
  try { return JSON.parse(raw) || {}; } catch (error) { return {}; }
}

function saveAccessRequests_(requests) {
  PropertiesService.getScriptProperties().setProperty(V6_AUTH.ACCESS_REQUESTS_PROPERTY_KEY, JSON.stringify(requests || {}));
}

function requestAccess(note) {
  const access = getAccessContext_();
  if (!access.verified || !access.email) throw new Error('Verify your email first.');
  if (access.canView) return {success:true, alreadyApproved:true};
  const requests = getAccessRequests_();
  requests[access.email] = {email:access.email, name:access.name, note:clean_(note), requestedAt:new Date().toISOString(), status:'Pending'};
  saveAccessRequests_(requests);
  try {
    logActivities_([{action:'Requested access', entityType:'Access', taskName:access.email, details:clean_(note), origin:'App'}], access);
  } catch (error) {}
  return {success:true};
}

function getAccessUsers() {
  requireAdmin_();
  const registry = getAccessRegistry_();
  return Object.keys(registry).map(function(email) {
    const entry = registry[email] || {};
    const role = normalizeRole_(entry.role);
    return {email:email, name:clean_(entry.name)||email, role:role, active:entry.active!==false, products:normalizeProductList_(entry.products,role)};
  }).sort(function(a,b) {
    const rank={Admin:0,Executive:1,Editor:2,Viewer:3};
    const d=(rank[a.role]??9)-(rank[b.role]??9);
    return d || a.email.localeCompare(b.email);
  });
}

function getAccessAdminData() {
  requireAdmin_();
  const requests = getAccessRequests_();
  return {users:getAccessUsers(), requests:Object.keys(requests).map(function(k){return requests[k];}).sort(function(a,b){return String(b.requestedAt).localeCompare(String(a.requestedAt));})};
}

function upsertAccessUser(payload) {
  const actor = requireAdmin_();
  payload = payload || {};
  const email = clean_(payload.email).toLowerCase();
  if (!email || email.indexOf('@') === -1) throw new Error('Enter a valid email address.');
  const adminEmail = clean_(CONFIG.DEFAULT_ADMIN_EMAIL).toLowerCase();
  let role = normalizeRole_(payload.role);
  if (role === 'No access') role = 'Viewer';
  if (email === adminEmail) role = 'Admin';
  const registry = getAccessRegistry_();
  const previous = registry[email] || null;
  registry[email] = {
    name:clean_(payload.name)||(previous&&previous.name)||email,
    role:role,
    active:email===adminEmail ? true : payload.active !== false,
    products:normalizeProductList_(payload.products, role)
  };
  saveAccessRegistry_(registry);

  const requests = getAccessRequests_();
  if (requests[email]) { delete requests[email]; saveAccessRequests_(requests); }

  logActivities_([{action:previous?'Updated access':'Granted access',entityType:'Access',taskName:email,field:'Role / Products',oldValue:previous?normalizeRole_(previous.role)+' · '+normalizeProductList_(previous.products,previous.role).map(displayProductLabelForServer_).join(', '):'',newValue:role+' · '+registry[email].products.map(displayProductLabelForServer_).join(', '),details:registry[email].name}],actor);
  return {success:true, users:getAccessUsers()};
}

function deactivateAccessUser(email) {
  const actor = requireAdmin_();
  const key = clean_(email).toLowerCase();
  if (key === clean_(CONFIG.DEFAULT_ADMIN_EMAIL).toLowerCase()) throw new Error('The default admin account cannot be deactivated.');
  const registry = getAccessRegistry_();
  if (!registry[key]) throw new Error('That access record was not found.');
  registry[key].active = false;
  saveAccessRegistry_(registry);
  logActivities_([{action:'Deactivated access',entityType:'Access',taskName:key,field:'Active',oldValue:'Active',newValue:'Inactive',details:registry[key].name||key}],actor);
  return {success:true, users:getAccessUsers()};
}



/* ============================================================
   WEEKLY CAPACITY
   - Default 40 hours per owner per week.
   - Admin can override an owner's availability for a specific week.
============================================================ */
