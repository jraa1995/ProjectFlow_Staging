# ProjectFlow — Copilot Instructions

## Platform
Google Apps Script (V8 runtime). Web app via `HtmlService`. Google Sheets = database (25 sheets). No npm, no ES modules, no `import`/`export` — all files share global scope. 6-min execution limit. Deploy as `USER_DEPLOYING`, `DOMAIN` access. Timezone: `America/New_York`.

## Architecture
```
UI (HTML/JS in ui/) → API (core/Code.js ~70 endpoints) → Data (core/Data.js CRUD) → Google Sheets
                                ↕
                    Engine Layer (engines/, core/DependencyEngine.js)
                    Service Layer (services/)
```

## File Map
- `core/Config.js` — `CONFIG` object: sheet names (`CONFIG.SHEETS.*`), column defs (`*_COLUMNS`), enums (`CONFIG.STATUSES/PRIORITIES/TYPES`), helpers (`generateId()`, `generateTaskId()`, `sanitize()`, `getColumnIndex()`)
- `core/Code.js` — `doGet(e)`, all `google.script.run` endpoints, `RequestCache`, `include(filename)`, cache invalidation (`invalidateTaskCache/ProjectCache/UserCache`), `patchTaskCache()`
- `core/Data.js` — `get*Sheet()` (auto-creates with headers), `rowToObject(row, columns)`, `objectToRow(obj, columns)`, `getAllTasks()`, `createTask()`, `updateTask()`, `deleteTask()`, `RowIndexCache`, `UserCache`
- `core/Board.js` — `buildBoardData()`, `calculateBoardStats()`, `searchTasks()`, `getFilteredTasks()`
- `core/AuthService.js` — Object literal. `authenticate()`, `createSession()`, `validateSession()`, rate limiting, MFA
- `core/PasswordService.js` — Object literal. SHA-256, 10000 iterations, `hashPassword()`, `verifyPassword()`, `constantTimeCompare()`
- `core/PermissionGuard.js` — Object literal. `can(userEmail, permission, projectId)`, `requirePermission(permission)`, 27 permissions, 4 roles (admin/manager/member/viewer)
- `core/Setup.js` — `quickSetup()`, `testSystem()`, `runQuickTests()`, `diagnose()`, `resetSystem()`
- `core/DependencyEngine.js` — Standalone functions. `addDependency()`, `detectCircularDependency()`, `calculateCriticalPath()`, `canStartTask()`
- `core/FunnelEngine.js` — Standalone functions. `importTicketsFromWorkbook()`, `stageTicketsInFunnel()`, `importFunnelTicketsToTasks()`
- `core/LockManager.js` — `updateTaskWithLocking()`, `acquireEditLock()`, `releaseEditLock()`, `withScriptLock(fn)` (uses GAS `LockService`)
- `engines/AnalyticsEngine.js` — Class, static methods. `calculateProductivityMetrics()`, `calculateVelocityTrend()`, `generatePredictions()`, `identifyBottlenecks()`, `calculateBurndownData()`, `detectAnomalies()`
- `engines/AutomationEngine.js` — Object literal. 8 trigger types, 9 action types, `evaluateConditions()`, `resolveAssignee()`
- `engines/CapacityEngine.js` — Object literal. `calculateUserWorkload()`, `getDailyWorkloadDistribution()`, `findAvailableUsers()`
- `engines/CriticalPathAnalysis.js` — Standalone functions. `generateCriticalPathAnalysis()`
- `engines/MentionEngine.js` — Class, static methods. `parseMentions()`, `getUserSuggestions()`
- `engines/NotificationEngine.js` — Class, static methods. `queueNotificationForDelivery()`, `replaceTemplateVariables()`
- `engines/SLAEngine.js` — Object literal. Priority-based SLA configs, `processEscalations()`
- `engines/TimelineEngine.js` — Class, static methods. `parseFlexibleDate()`, `generateProjectTimeline()`, `calculateCriticalPath()`
- `engines/TimelineFiltering.js` — Standalone functions. `searchTimelineTasks()`, `parseSearchQuery()`
- `engines/TriageEngine.js` — Object literal. `autoSuggest()`, `importFromGmail()`
- `services/EmailNotificationService.js` — Class, static methods. 6 email templates, retry (3 attempts, exponential backoff)

## UI Files
- `ui/Index.html` — Main shell. Uses `<?!= include('ui/FileName'); ?>` for composition. Contains layout, modals, view containers.
- `ui/Scripts.html` — 3385 lines. `State` (global state), `LocalCache` (localStorage, `pf_` prefix), `OptimisticUI` (snapshot/revert), `RequestQueue` (debounced batching), `SyncManager` (30s polling), `LoginModal`, `switchView()`, `renderBoard()`, `renderTaskCard()`, `saveTask()`, `toast()`, `escapeHtml()`, `$(id)`/`$$(selector)` shorthand, keyboard shortcuts (Esc/Ctrl+N/Ctrl+/)
- `ui/CacheService.html` — `CacheService` object: IndexedDB with memory fallback, TTL-based
- `ui/DesignSystem.html` — 60+ CSS variables (`--pf-primary-50` to `--pf-primary-900`, `--pf-gray-*`, semantic), component classes (`.pf-btn`, `.pf-card`, `.pf-badge`, `.pf-input`, `.pf-avatar`, `.pf-skeleton`, `.pf-spinner`), animations
- `ui/Styles.html` — Nav, modal, mention, comment, notification, board CSS
- `ui/TaskDetailPanel.html` — `TaskDetailPanel` object. Slide-in, tabs (details/comments/activity)
- `ui/TaskEditModal.html` — `TaskEditModal` object. Split-pane (form left, comments right)
- `ui/GanttChart.html` — `GanttState` object. Frappe Gantt integration
- `ui/ListView.html` — `ListViewPage` object. Sortable table, inline editing, bulk ops
- `ui/CalendarView.html` — `CalendarViewPage` object. FullCalendar integration
- `ui/MatrixView.html` — `MatrixViewPage` object. 2D grid (status × assignee/priority), drag-drop
- `ui/MilestoneView.html` — `MilestonePage` object. Milestone dashboard + timeline
- `ui/FunnelView.html` — `FunnelPage` object. 3-step import wizard
- `ui/DependencyPicker.html` — `DependencyPicker` object. Task search + type selector
- `ui/MentionAutocomplete.html` — `MentionAutocomplete` class. Input listener, dropdown, debounced
- `ui/NotificationCenter.html` — `NotificationCenter` object. Slide-in, tabs (all/unread/mentions)
- `ui/TimelineIntegration.html` — DEPRECATED. Removed from includes (caused `switchView` conflicts).

## Sheets (CONFIG.SHEETS)
Tasks, Users, Projects, Comments, Activity, Mentions, Notifications, Analytics_Cache, Task_Dependencies, Funnel_Staging, Session_Tokens, MFA_Codes, Roles, Project_Members, Organizations, Organization_Members, Teams, Team_Members, User_Preferences, Automation_Rules, Team_Capacity, Webhook_Subscriptions, SLA_Config, Triage_Queue

## Enums (CONFIG)
- STATUSES: `['Backlog','To Do','In Progress','Review','Testing','Done']`
- PRIORITIES: `['Lowest','Low','Medium','High','Highest','Critical']`
- TYPES: `['Task','Bug','Feature','Story','Epic','Spike']`
- MILESTONE_TYPES: `['phase-end','deliverable','review','launch','other']`
- DEPENDENCY_TYPES: `['finish_to_start','start_to_start','finish_to_finish','start_to_finish']`
- FUNNEL_STATUSES: `['pending','reviewed','imported','rejected']`

## Permissions (27)
`task:create`, `task:read:own`, `task:read:all`, `task:update:own`, `task:update:all`, `task:delete:own`, `task:delete:all`, `task:assign`, `task:move`, `comment:create`, `comment:read`, `comment:delete:own`, `comment:delete:all`, `project:create`, `project:read`, `project:update`, `project:delete`, `project:manage_members`, `user:read`, `user:manage`, `role:manage`, `analytics:view`, `automation:manage`, `sla:manage`, `funnel:manage`, `webhook:manage`, `admin:system`

## Coding Patterns

### API Return Format (MANDATORY)
Every `google.script.run`-exposed function in Code.js:
```javascript
function myEndpoint(params) {
  try {
    // PermissionGuard.requirePermission('relevant:permission') if needed
    // logic
    return { success: true, data: result };
  } catch (error) {
    console.error('myEndpoint error:', error);
    return { success: false, error: error.message };
  }
}
```

### Module Patterns (match existing file)
1. **Object literal singleton** (AuthService, PermissionGuard, AutomationEngine, SLAEngine, CapacityEngine, TriageEngine): `const MyModule = { method() {}, ... };`
2. **Class with static methods** (AnalyticsEngine, TimelineEngine, MentionEngine, NotificationEngine, EmailNotificationService): `class MyEngine { static method() {} }`
3. **Standalone functions** (CriticalPathAnalysis, TimelineFiltering, DependencyEngine, FunnelEngine, LockManager, Setup, Board, Data): `function myFunction() {}`
For new engines: prefer object literal.

### Naming
- Functions/methods: `camelCase`
- Constants/enums: `SCREAMING_SNAKE`
- Classes & singleton objects: `PascalCase`
- IDs: `generateId('prefix')` → `prefix_abc123`; task IDs: `generateTaskId(projectId, tasks)` → `PROJ-001`
- DOM element IDs: `camelCase`
- Sheet names: `Title_Case` with underscores
- CSS variables: `--pf-*`
- CSS classes: `.pf-*`
- localStorage keys: `pf_*` prefix

### Data Access
- Each sheet has `get*Sheet()` in Data.js that auto-creates with headers if missing
- Column definitions in Config.js as `*_COLUMNS` objects mapping field names to 1-based indices
- Read: `sheet.getDataRange().getValues()` → `rowToObject(row, columns)` per row (skip header)
- Write: `objectToRow(obj, columns)` → `sheet.appendRow()` or `sheet.getRange().setValues()`
- Always call appropriate `invalidate*Cache()` after mutations

### Frontend → Backend
```javascript
google.script.run
  .withSuccessHandler(result => { if (result.success) { /* use result.data */ } else { toast(result.error, 'error'); } })
  .withFailureHandler(err => handleError(err))
  .serverFunctionName(args);
```

### Frontend State
- `State` global object: `.user`, `.board`, `.projects`, `.users`, `.config`, `.currentView`, `.projectFilter`, `.tasks`
- `switchView(view)` — 9 views: `my`, `master`, `timeline`, `analytics`, `funnel`, `list`, `calendar`, `matrix`, `milestones`
- Each view has a page controller object with `init()` or `render()` method
- New views: register in `switchView()` in Scripts.html, add container div in Index.html, create `ui/NewView.html` with page controller, add `<?!= include('ui/NewView'); ?>` in Index.html

### Optimistic Updates
```javascript
OptimisticUI.apply(
  'uniqueKey',
  snapshotData,
  (data) => { /* immediate DOM update */ },
  (data) => { /* update State/cache */ },
  (data, resolve, reject) => { google.script.run.withSuccessHandler(resolve).withFailureHandler(reject).endpoint(data); }
);
```

### Boolean Fields from Sheets
Always check multiple truthy forms: `value === true || value === 'true' || value === 'TRUE' || value === 1 || value === '1'`

### Date Handling
Use `TimelineEngine.parseFlexibleDate(value)` — handles Date objects, Excel serial numbers, US format (MM/DD/YYYY), ISO strings.

### Error Toast
`toast(message, 'error'|'success'|'info'|'warning')` — auto-dismiss 3s.

### XSS Prevention
Always `escapeHtml(text)` for user-generated content rendered as HTML.

### Locking for Concurrent Mutations
Use `LockManager.updateTaskWithLocking(taskId, updateFn)` for task updates. `withScriptLock(fn)` for sheet-level atomicity. Version conflict: throws `{code:'VERSION_CONFLICT', serverVersion, clientVersion, taskId}`.

### Auth on New Endpoints
Call `PermissionGuard.requirePermission('permission:name')` at the start of any new Code.js endpoint that modifies data. Use `PermissionGuard.can(email, permission, projectId)` for conditional checks.

## Caching (5 layers — respect all on mutations)
1. `RequestCache` — in-memory per GAS execution (Code.js)
2. `UserCache`/`RowIndexCache` — in-memory cross-function (Data.js, 300-600s)
3. `CacheService.getScriptCache()` — GAS server cache (300-600s TTL, 100KB max per value)
4. `CacheService` (IndexedDB) — client persistent (ui/CacheService.html, 5-30min TTL)
5. `LocalCache` (localStorage) — client fast (ui/Scripts.html, `pf_` prefix)
After server mutations: call `invalidateTaskCache()`/`invalidateProjectCache()`/`invalidateUserCache()` in Code.js. Client: `State.invalidateCache()` + `CacheService.invalidate()`.

## External Libraries (CDN only, no npm)
Tailwind CSS (CDN script, primary amber `#f59e0b`), Font Awesome 6.4.0, Frappe Gantt 0.6.1, Chart.js 4.4.0, FullCalendar 6.1.10

## GAS Constraints
- No `import`/`export` — global scope shared across all .js files
- `include(filename)` in Code.js strips folder prefixes (`ui/`, `core/`, `engines/`, `services/`) because GAS flattens paths
- `Session.getActiveUser().getEmail()` can return `''` — `getCurrentUserEmail()` in Data.js has fallback
- 6-minute max execution — use `LockService` and batch operations
- 100KB `CacheService` value limit — check size before storing
- Dates from Sheets may be Excel serial numbers
- `module.exports` guard: `if (typeof module !== 'undefined') module.exports = ...` (AuthService, PasswordService only, for testing)

## Key Constants
| Constant | Value |
|---|---|
| Server task cache TTL | 300s |
| Server project/user cache TTL | 600s |
| Cache size limit | 100KB |
| Session duration | 24h |
| Session refresh threshold | 1h |
| Rate limit | 5 attempts / 15min, 30min lockout |
| MFA code expiry | 5min |
| Edit lock TTL | 300s |
| Password hash iterations | 10000 |
| Overload threshold | 120% |
| Default work hours | 8h/day |
| Sync poll interval | 30s |
| Inactivity threshold | 60s |
| Init timeout | 15s |
| Toast dismiss | 3s |
| IndexedDB TTL (tasks) | 5min |
| IndexedDB TTL (projects) | 10min |
| IndexedDB TTL (users) | 15min |
| IndexedDB TTL (config) | 30min |

## Roadmap Status (ENTERPRISE_ROADMAP.md)
- Phase 1 (Milestones): DONE
- Phase 2 (Dependencies/Critical Path): DONE
- Phase 3 (Advanced Views — List/Calendar/Matrix): DONE
- Phase 4 (Custom Fields): PLANNED
- Phase 5 (Workflow Automations/SLA): DONE
- Phase 6 (Resource Management): PARTIAL — CapacityEngine exists, Workload/Time views planned
- Phase 7 (Portfolio & Reporting): PLANNED
- Phase 8 (Templates, API): PARTIAL — RBAC done, rest planned
