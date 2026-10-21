# Engineering Workload V7 — Modular + Capacity Planning

## What changed

- Code is split by domain instead of one giant `Code.gs` / `Index.html`.
- Weekly planning now uses **Estimated Hours** and a default **40h weekly capacity per owner**.
- Admin can override capacity for an owner for a specific week (leave, training, holidays, reduced availability).
- Source `Points` are preserved as legacy metadata; they are no longer used for workload capacity.
- Done / closed tasks still cannot be planned.
- Existing Weekly Workplan rows are preserved. A new `Estimated Hours` column is appended; old rows remain blank until estimated.
- A `Weekly Capacity` sheet is created automatically when an editor/admin first needs it.
- Summary and Weekly Plan show planned hours, available capacity, remaining capacity, utilization, and missing estimates.
- Tasks pagination, collapsible weekly owners, persistent verification sessions, Admin-only Changelog/Users, pending-task sync, and robust delete remain included.

## Apps Script files

### Server
- `00_Config.gs` — constants, source workbook routing
- `01_Web.gs` — web app/template entry point + dashboard workbook resolution
- `02_Auth_Access.gs` — persistent 30-day sessions, email verification, RBAC/users
- `03_Routing.gs` — product/source routing helpers
- `04_Utils.gs` — dates, weeks, status, priorities, validation helpers
- `05_Audit.gs` — app audit + direct Google Sheet audit triggers
- `06_Validation.gs` — data quality + source-sheet validation
- `07_Tasks.gs` — source-backed task CRUD, IDs, pending sync, delete
- `08_WeeklyPlan.gs` — weekly workplan CRUD and estimated hours
- `09_Capacity.gs` — 40h defaults and per-owner weekly overrides
- `10_Dashboard.gs` — scoped dashboard/week payloads
- `11_Api.gs` — authenticated API dispatcher

### Client
- `Index.html` — shell only
- `Styles.html` — all CSS
- `00_State.html` — state, constants, shared helpers, API client
- `10_Auth.html` — sign-in / verification / startup
- `20_Shell.html` — navigation and app chrome
- `30_SummaryView.html` — executive summary
- `40_TasksView.html` — Tasks table/board/pagination/drawer
- `50_WeeklyView.html` — weekly plan, collapsible owners, hours/capacity
- `60_AdminView.html` — Changelog, Users, audit/validation admin
- `70_TaskForms.html` — create-task and shared modal/form behavior
- `99_App.html` — startup event bindings

## Deploy

1. In the existing Apps Script project, create/replace the files above.
2. Do **not** keep the old monolithic `Code.gs` functions alongside these modules — duplicate function names will make maintenance confusing.
3. Save.
4. Deploy → Manage deployments → Edit → New version → Deploy.
5. Keep **Execute as: Me** and the same existing deployment URL.

## Capacity behavior

- Default: `40h` per owner per selected week.
- Planned hours = sum of `Estimated Hours` for active weekly tasks.
- Remaining = available capacity − planned hours.
- Utilization = planned / available × 100.
- Completed weekly rows are excluded from active utilization.
- Admin can set a weekly override from the owner header.
- Estimated Hours are required when adding a task to the weekly plan.
- Old weekly rows with blank estimates are shown as unestimated instead of being guessed from story points.


## V7.1 — Suggested Remaining Effort + UI fix

- Fixes the malformed **Add to week** checkbox row. The problem was that the generic `.field input { width:100% }` rule was also stretching checkbox inputs.
- Adds a dedicated `15_Estimation.html` module.
- Every non-Done task gets a suggested **remaining-effort** estimate:
  - Tiny config/content/data change → 1h
  - Very small fix / QA fix → 2h
  - Small UI/backend change → 4h
  - Normal engineering task → 8h
  - Larger feature/integration → 12–16h
  - Complex multi-part work → 24h + split recommendation
- In Progress work is reduced by one effort tier.
- QA / Review is capped at 4h remaining effort.
- Blocked work is marked low-confidence.
- Tasks table and board show the suggestion.
- The Plan modal pre-fills the suggested hours, but the user can override them.
- Manual weekly estimates continue to be the persisted value and are never overwritten by the suggestion engine.


## V7.2 — Nano/ShowApp source visibility + form cleanup

- Nano now supports **Master Backlog** alongside **BE Backlog** and **FE Backlog**.
- Nano Master Backlog task IDs use `NANO-###`.
- All product tasks are read directly from their configured source sheets; no normalized task register sheet is required.
- Opening a task rehydrates the drawer from the real source row.
- Source-specific metadata is shown in a read-only **Source details** section when available.
- Tasks are edited and planned directly from their configured source rows.
- Removed **Source points (optional / legacy)** from the New Task form and Task drawer. Existing source Points data is left untouched.
- Fixed the grid stretching that made Weekly Commitment much taller than Estimated Remaining Hours.
- Simplified estimate guidance in forms to compact inline text while keeping detailed reasoning in the task drawer.


## V7.3 — Manual estimated hours only

- Removed all automatic/suggested hour estimates from the UI.
- Engineers now enter **Estimated Hours** manually for development work.
- **QA / Review / In Review** tasks do not display, require, or contribute estimated hours.
- **Done / Complete / Completed / Resolved / Passed / Closed** tasks do not display or contribute estimated hours.
- Missing-estimate warnings apply only to capacity-relevant development tasks.
- Moving a Weekly Plan row into QA / Review or Done clears its weekly estimated-hours value.
- Backlog, Todo, In Progress, and Blocked work can still use manual hours for 40-hour capacity planning.


## V7.4 — Schema-aware, table-safe validation

Validation setup is now non-destructive and source-schema aware.

- **Nano** (`BE Backlog`, `FE Backlog`, `Master Backlog`) and **Klima** (`Master Backlog`) inspect Google Sheets table metadata before any validation write. Typed table columns are preserved and skipped rather than overwritten.
- If table metadata cannot be inspected for Nano/Klima, validation is skipped for safety instead of risking a table-type change.
- **ShowApp / Master Backlog** maps:
  - `ID` → Task ID
  - `Feature / Issue` → Task Title
  - `Requirement / Description` → Description
- **Governance / Repository Register** explicitly uses header row 4 and maps:
  - `ID` → Task ID
  - `Repository` → Task Title
  - `Purpose / Description` → Description
  - `Primary Owner` → Owner
  - `Status` → Status
- Existing source column names are never renamed.
- MediaPal and Weekly Workplan validation is idempotent: existing equivalent rules are left in place rather than rewritten.
- Existing *different* validation rules are also preserved and reported as skipped instead of being overwritten.
- Status/Priority dropdown rules include current source values so existing product-specific values remain valid.
- Validation results now report **successful**, **skipped**, and **failed** field operations separately.
- Weekly Workplan headers are preserved; only a blank `Estimated Hours` header is populated when needed.
- Source-backed reads now also support Klima and Governance; ShowApp/Governance descriptions are read without renaming their source columns.


## V7.5 — Unified right drawer + clearer changelog

- **Add Task uses the same right-side drawer as existing tasks** instead of a centered modal.
- The right drawer is reused for existing task details, new task creation, and audit-event details.
- Changelog adds compact metrics, an Action filter, cleaner source badges and a denser table.
- The **Change** column is now clickable and opens the complete audit event in the right drawer.
- Change detail shows Before, After, Context, user, timestamp, source, source sheet, product, field, entity type and request ID.
- Task IDs remain clickable and open the existing task drawer.
- V7.4 schema-aware validation and source mappings remain unchanged.


## V7.6 — Cleaner individual task drawer

The task drawer now shows only information needed for normal task management.

- Removed the always-visible **Source details** block from every task.
- Raw source metadata such as Team, Project ID, Creator, Labels, Created, Updated and UUID remains in the original source system; it is not duplicated into a new app database.
- Admins can still inspect this metadata through a collapsed **Technical metadata** disclosure when troubleshooting.
- Removed the empty `History — No history loaded yet` section. Task activity only appears when history exists and is collapsed by default.
- Full audit history remains centralized in **Changelog**, which is the primary history store/view.
- Removed the permanent `Selected week` section from unplanned tasks. Weekly context only appears when that task is actually planned for the selected week.
- Eligible unplanned tasks get a compact **Plan this week** action in the drawer footer.
- Source sheet/workbook is reduced to one quiet source line rather than a large metadata panel.
- No underlying source-sheet fields or data were deleted or migrated.
