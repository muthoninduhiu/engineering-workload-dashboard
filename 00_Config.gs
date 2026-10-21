/**
 * ENGINEERING WORKLOAD V7
 * Modular portfolio + weekly capacity planning.
 * Generated from the stable V6.7 build and refactored by domain.
 */

/**
 * ============================================================
 * ENGINEERING WORKLOAD & WEEKLY PLANNING
 * FULL Code.gs
 * ============================================================
 *
 * TASK REGISTER:
 *   Configured source sheets in TASK_ROUTING
 *
 * SOURCE TASK SHEETS:
 *   MediaPal / WePlay -> Consolidated Linear Tasks
 *   Nano Backend      -> BE Backlog
 *   Nano Frontend     -> FE Backlog
 *   Nano Master       -> Master Backlog
 *   ShowApp           -> Master Backlog (ShowApp workbook)
 *   Klima             -> Master Backlog (Klima workbook)
 *   Governance        -> Repository Register
 *
 * WEEKLY PLAN:
 *   Weekly Workplan
 *
 * IMPORTANT:
 * - New tasks are created in the REAL source workbook/sheet.
 * - Task data is read from and written to its configured source sheet.
 * - Weekly Workplan is dynamic by selected week.
 * ============================================================
 */

const CONFIG = {
  WEEKLY_SHEET: 'Weekly Workplan',
  CAPACITY_SHEET: 'Weekly Capacity',
  DEFAULT_WEEKLY_CAPACITY_HOURS: 40,
  AUDIT_SHEET: 'Activity Log',
  ACCESS_PROPERTY_KEY: 'ENGINEERING_WORKLOAD_ACCESS_V1',
  DASHBOARD_ID_PROPERTY_KEY: 'ENGINEERING_WORKLOAD_DASHBOARD_SPREADSHEET_ID',
  DASHBOARD_LABEL: 'Engineering Workload dashboard',
  DEFAULT_ADMIN_EMAIL: 'maryanne.nduhiu@mediapal.net',
  WEEK_WINDOW_RADIUS: 2,
  WEEK_CACHE_SECONDS: 120,
  WEEK_CACHE_VERSION_PROPERTY_KEY: 'ENGINEERING_WORKLOAD_WEEK_CACHE_VERSION',
  SOURCE_TASK_CACHE_SECONDS: 900,
  SOURCE_TASK_CACHE_VERSION_PROPERTY_KEY: 'ENGINEERING_WORKLOAD_SOURCE_TASK_CACHE_VERSION',
  SOURCE_TASK_CACHE_PREVIOUS_VERSION_PROPERTY_KEY: 'ENGINEERING_WORKLOAD_SOURCE_TASK_CACHE_PREVIOUS_VERSION',
  PENDING_CREATED_TASKS_PROPERTY_KEY: 'ENGINEERING_WORKLOAD_PENDING_CREATED_TASKS_V1',
  DIRECT_AUDIT_EDIT_HANDLER: 'handleDirectSheetEdit',
  DIRECT_AUDIT_CHANGE_HANDLER: 'handleDirectSheetChange'
};

const DEFAULT_ACCESS_USERS = {
  'maryanne.nduhiu@mediapal.net': {
    name: 'Maryanne',
    role: 'Admin',
    active: true
  }
};

const VALID_TASK_PRIORITIES = ['P0', 'P1', 'P2', 'P3', 'No priority'];
const VALID_TASK_STATUSES = [
  'Future Backlog',
  'Backlog',
  'Not Started',
  'Todo',
  'In Progress',
  'In Review',
  'QA',
  'Testing',
  'Retest',
  'Awaiting Tech Fix',
  'Blocked',
  'Failed',
  'Done',
  'Passed',
  'Closed'
];
const VALID_WEEKLY_COMMITMENTS = ['Commit', 'Stretch', 'Watch / Unblock'];
const VALID_SOURCE_STATUSES = [
  'Backlog', 'None', 'Not Started', 'Todo', 'To Do', 'In Progress', 'Active', 'Doing',
  'In Review', 'QA', 'Testing', 'Test', 'Retest', 'Awaiting Tech Fix',
  'Blocked', 'Failed', 'Done', 'Complete', 'Completed', 'Resolved',
  'Passed', 'Closed', 'Legacy', 'Replaced', 'Cancelled', 'Canceled'
];
const VALID_SOURCE_PRIORITIES = [
  'P0', 'P1', 'P2', 'P3', 'Critical', 'Urgent', 'High', 'Medium', 'Low', 'No priority', 'None'
];

/* ============================================================
   ROUTING
============================================================ */

const TASK_ROUTING = {
  MediaPal: [
    {
      workstream: 'Linear / Engineering',
      sourceSheet: 'Consolidated Linear Tasks',
      spreadsheetId: '1eT8tDGRmLn3Y7xLf6-nVfgnEF7_r785jllNIEGsXjFI'
    }
  ],

  WePlay: [
    {
      workstream: 'Product / Engineering',
      sourceSheet: 'Consolidated Linear Tasks',
      spreadsheetId: '1eT8tDGRmLn3Y7xLf6-nVfgnEF7_r785jllNIEGsXjFI'
    }
  ],

  Nano: [
    {
      workstream: 'Backend',
      sourceSheet: 'BE Backlog',
      spreadsheetId: '1pTpLKCxhqaXxYLXrjeWHkXFWIhyMLFu3b0YlPrMsHRI'
    },
    {
      workstream: 'Frontend',
      sourceSheet: 'FE Backlog',
      spreadsheetId: '1pTpLKCxhqaXxYLXrjeWHkXFWIhyMLFu3b0YlPrMsHRI'
    },
    {
      workstream: 'Master Backlog',
      sourceSheet: 'Master Backlog',
      spreadsheetId: '1PEQeiQvE6Tg7S1t8sz6-x-Z_7JWxuhaS6nv8c6IIJeM'
    }
  ],

  ShowApp: [
    {
      workstream: 'Product Backlog',
      sourceSheet: 'Master Backlog',
      spreadsheetId: '1Mg7MIfqXsIbilpoVOE9681h_wNRtnhkN4KJ1eWJ1PBo'
    }
  ],

  Klima: [
    {
      workstream: 'Product Backlog',
      sourceSheet: 'Master Backlog',
      spreadsheetId: '1anoMgoOVb3s4p3SAVZwzKY9PqjC3PlJFWcgq6XkudvI'
    }
  ],

  Governance: [
    {
      workstream: 'Documentation & Governance',
      sourceSheet: 'Repository Register',
      spreadsheetId: '1Ypn2koQnt5HiBDZAYMwed68l7SJM1SJm1nv-Sz6FdsU'
    }
  ]
};


/* ============================================================
   SOURCE SCHEMA PROFILES
   ------------------------------------------------------------
   These are aliases only. They NEVER rename source columns.
============================================================ */

const SOURCE_SCHEMA_OVERRIDES = {
  '1gFCDw6wCN56woJ949olHmSOqpN653SOKeEMe7jNflD4::consolidated linear tasks': {
    idAliases: ['Identifier', 'Issue Identifier', 'Linear ID', 'Issue ID', 'Issue Key', 'Task ID', 'ID', 'Key'],
    titleAliases: ['Title', 'Summary', 'Name', 'Task Title', 'Task Name'],
    descriptionAliases: ['Description'],
    ownerAliases: ['Assignee', 'Owner'],
    statusAliases: ['Status'],
    priorityAliases: ['Priority'],
    dueDateAliases: ['Due Date', 'Due']
  },

  '1aC6svnsItgbtGzz7C-_dmjtNGaF3804n7oNJxpDKabs::master backlog': {
    idAliases: ['ID'],
    titleAliases: ['Feature / Issue'],
    descriptionAliases: ['Requirement / Description'],
    ownerAliases: ['Primary Owner'],
    statusAliases: ['Status']
  },

  
'1Uo7IeD-QdARCIx_IlCK5-wK-cMpA1-nouL57VdIdrkI::repository register': {
  headerRow: 4,

  idAliases: [
    'ID',
    'Task ID',
    'Reference ID'
  ],

  titleAliases: [
    'Repository',
    'Task Title',
    'Task',
    'Title'
  ],

  descriptionAliases: [
    'Purpose / Description',
    'Description',
    'Notes'
  ],

  ownerAliases: [
    'Primary Owner',
    'Owner'
  ],

  statusAliases: [
    'Status'
  ]
}

};

/* ============================================================
   WEB APP
============================================================ */

const V6_AUTH = {
  SESSION_PREFIX: 'EW_V6_SESSION_',
  SESSION_PROPERTY_PREFIX: 'EW_V6_SESSION_PERSIST_',
  SESSION_TTL_SECONDS: 21600,
  SESSION_PERSIST_DAYS: 30,
  OTP_PREFIX: 'EW_V6_OTP_',
  OTP_RATE_PREFIX: 'EW_V6_OTP_RATE_',
  OTP_TTL_SECONDS: 600,
  ACCESS_REQUESTS_PROPERTY_KEY: 'ENGINEERING_WORKLOAD_ACCESS_REQUESTS_V1'
};

const V6_PRODUCTS = ['MediaPal', 'WePlay', 'Nano', 'ShowApp', 'Klima', 'Governance'];

var REQUEST_AUTH_CONTEXT_ = null;
