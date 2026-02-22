# PermiX — Codebase Reference

## Overview

PermiX is a PowerShell HTTP server that serves a browser SPA. The server collects SharePoint data via PnP PowerShell / Microsoft Graph and exposes it through a REST API. The frontend is vanilla HTML/CSS/JS.

**Request flow:** Browser → `WebServer.ps1` (HttpListener) → `ApiHandlers.ps1` → SharePoint/Data functions → `SharePointDataManager.ps1` (in-memory store) → JSON response → Frontend JS

---

## Entry Points

### `Start-SPOTool-Web.ps1`
Dot-sources all modules in order (Core → Analysis → SharePoint → Demo → Server), calls `Initialize-SharePointDataManager`, then `Start-WebServer`. If `SPO_TENANT_URL` + `SPO_CLIENT_ID` env vars exist (container mode), auto-connects via device code before starting the server.

### `docker-entrypoint.ps1`
Container startup: sets `SPO_HEADLESS=1`, calls `Start-SPOTool-Web.ps1` with `-NoBrowser -ListenAddress +`.

---

## Backend — Functions/Server/

### `WebServer.ps1`

| Function | What it does |
|---|---|
| `Start-WebServer` | Creates `System.Net.HttpListener`, populates `$script:ServerState` (synchronized hashtable shared with background runspaces), starts the request loop. Uses `BeginGetContext`/`EndGetContext` with a 500 ms wait so the `Running` flag can be checked between requests. Opens browser via `Start-Process`. |
| `Invoke-RequestHandler` | Dispatches each request: paths starting with `/api/` go to `Invoke-ApiHandler`, everything else to `Send-StaticFile`. |
| `Send-StaticFile` | Resolves path under `Web/`, enforces path-traversal check, maps extensions to MIME types, sets cache headers (1 h for CSS/JS, no-cache for HTML), writes raw bytes. |
| `Send-JsonResponse` | Serializes `$Data` with `ConvertTo-Json -Depth 10`, writes UTF-8 bytes with `Content-Type: application/json`. Used by all API handlers. |
| `Read-RequestBody` | Reads `InputStream`, returns parsed JSON object via `ConvertFrom-Json`. |
| `Stop-WebServer` | Sets `$script:ServerState.Running = $false` and calls `$listener.Stop()`. |

**`$script:ServerState`** (synchronized hashtable) — key shared state:
- `OperationLog` — `[ArrayList]::Synchronized` written by background runspaces, read by `/api/progress`
- `OperationRunning` / `OperationComplete` / `OperationError` — progress flags
- `SharePointData` — reference to the same object as `$script:SharePointData` in the data manager
- `BackgroundJob` — `IAsyncResult` from `PowerShell.BeginInvoke()`

---

### `ApiHandlers.ps1`

All REST routes live here. `Invoke-ApiHandler` uses `switch -Wildcard` to dispatch.

| Route | Handler | What it does |
|---|---|---|
| `GET /api/status` | `Handle-GetStatus` | Returns `connected`, `demoMode`, `headless`, current metrics. |
| `POST /api/connect` | `Handle-PostConnect` | Saves settings, checks PnP module, connects via `-Interactive` (local) or `-DeviceLogin` (headless), calls `Test-UserCapabilities`, returns site info + capability flags. |
| `POST /api/demo` | `Handle-PostDemo` | Sets `$script:DemoMode = $true`, calls `New-DemoData`. |
| `POST /api/sites` | `Handle-PostSites` | Starts `Get-RealSites-DataDriven` in a background runspace via `Start-BackgroundOperation`. Returns immediately; frontend polls `/api/progress`. |
| `POST /api/prepare-analysis` | `Handle-PostPrepareAnalysis` | Checks if re-auth is needed (container mode, different site URL). Returns `needsAuth: true/false`. |
| `POST /api/permissions` | `Handle-PostPermissions` | Optionally re-authenticates (device code, synchronously), then starts `Get-RealPermissions-DataDriven` in background. |
| `GET /api/progress` | `Handle-GetProgress` | Returns `OperationLog[]`, `running`, `complete`, optional `error` and `enrichmentResult`. |
| `GET /api/data/:type` | `Handle-GetData` | Maps URL type via `$script:DataTypeMap`, calls `Get-SharePointData`. |
| `GET /api/metrics` | `Handle-GetMetrics` | Returns `OperationMetrics` from the data manager. |
| `POST /api/enrich` | `Handle-PostEnrich` | Demo: calls `Invoke-DemoEnrichment`. Live: starts `Invoke-ExternalUserEnrichment` in background. |
| `GET /api/enrichment` | `Handle-GetEnrichment` | Calls `Get-EnrichmentSummary` — returns counts of disabled/stale/guest accounts. |
| `GET /api/risk` | `Handle-GetRisk` | Calls `Get-RiskAssessment` — evaluates all risk rules and returns scored findings. |
| `GET /api/audit` | `Handle-GetAudit` | Returns current `$script:AuditSession` metadata. |
| `GET /api/export/:type` | `Handle-PostExport` | Converts data to CSV, sends as `Content-Disposition: attachment`. |
| `GET /api/export-json` | `Handle-PostExportJson` | Returns full `Build-GovernanceReport` object as JSON. |
| `GET /api/export-json/:type` | `Handle-PostExportJsonType` | Returns single data type as JSON with schema envelope. |
| `POST /api/build-permissions-matrix` | `Handle-PostBuildPermissionsMatrix` | Demo: `Get-DemoPermissionsMatrix`. Live: reconnects to the specific site, calls `Get-SitePermissionsMatrix`. |
| `POST /api/shutdown` | inline | Calls `Stop-WebServer`. |

---

### `BackgroundJobManager.ps1`

#### `Start-BackgroundOperation`
Spawns a new PowerShell runspace for long-running operations (site fetch, permissions analysis, enrichment) so the HTTP server stays responsive.

**Flow:**
1. Grabs current PnP access token + tenant/client settings from the main thread.
2. Creates a fresh `Runspace`, injects `$SharedState`, `$ScriptRoot`, `$AccessToken`, `$TenantUrl`, `$ClientId`.
3. The wrapper scriptblock dot-sources all modules into the new runspace, overrides `Write-ConsoleOutput` to append to `$SharedState.OperationLog`, and points `$script:SharePointData` at `$SharedState.SharePointData` (same object reference — changes are visible immediately in the main thread).
4. Re-establishes PnP connection inside the runspace:
   - **Container/headless:** tries access token (same site only), falls back to `-DeviceLogin`.
   - **Local/Windows:** uses `-Interactive`, falls back to access token.
5. Sets `OperationRunning = $true`, runs the caller's scriptblock, sets `OperationComplete = $true` on success or writes `OperationError` on failure.

---

## Backend — Functions/Core/

### `SharePointDataManager.ps1`
Central in-memory store. All collections are `[ArrayList]`, all metrics are in `OperationMetrics`.

| Function | What it does |
|---|---|
| `Initialize-SharePointDataManager` | Resets `$script:SharePointData` to empty structure. Called once at startup. |
| `Add-SharePointSite` | Validates required fields, auto-calculates `UsageLevel` from storage MB, appends to `Sites`, increments `TotalSites`. |
| `Add-SharePointUser` | Validates fields, appends to `Users`, increments `TotalUsers` / `ExternalUsers`. |
| `Add-SharePointGroup` | Validates fields, appends to `Groups`, increments `TotalGroups`. |
| `Add-SharePointRoleAssignment` | Appends to `RoleAssignments`, increments `TotalRoleAssignments`. |
| `Add-SharePointInheritanceItem` | Appends to `InheritanceItems`, increments `InheritanceBreaks` if `HasUniquePermissions`. |
| `Add-SharePointSharingLink` | Appends to `SharingLinks`, increments `TotalSharingLinks`. |
| `Clear-SharePointData` | Clears one or all collections in-place (does NOT replace the hashtable — preserves shared reference held by background runspaces). |
| `Get-SharePointData` | Returns the requested collection or `OperationMetrics`. Keys: `Sites`, `Users`, `Groups`, `Permissions`, `RoleAssignments`, `InheritanceItems`, `SharingLinks`, `Metrics`, `All`. |
| `Set-SharePointOperationContext` | Sets `LastOperation` string and `LastUpdateTime`. |

---

### `Settings.ps1`
- `$script:AppSettings` — nested hashtable (dot-path keys like `"SharePoint.TenantUrl"`).
- `$script:SPOConnected`, `$script:DemoMode` — global connection flags.
- `Get-AppSetting` / `Set-AppSetting` — read/write by dot-path (e.g. `"Logging.LogPath"`).

### `Logging.ps1`
- `Write-ActivityLog` — timestamped console output + appends to `Logs/activity_log.txt`. Colour-coded by level.
- `Write-ErrorLog` — writes to `Logs/error_log.txt` and prints red.

### `OutputAdapter.ps1`
- `Write-ConsoleOutput` — in the main thread, appends to `$script:ServerState.OperationLog`. In background runspaces this function is **overridden** by `BackgroundJobManager` to append to `$SharedState.OperationLog`.
- `Update-UIAndWait` — no-op (WPF remnant, kept for compatibility).

### `ThrottleProtection.ps1`
#### `Invoke-WithThrottleProtection`
Wraps any scriptblock. On `429`/`503`/throttle/timeout errors, retries up to `MaxRetries` (default 5) with exponential backoff + jitter (starts at 2 s, max 60 s). Respects `Retry-After` header. Rethrows on non-throttle errors or after max retries.

- `Get-ThrottleStats` / `Reset-ThrottleStats` — read/reset `$script:ThrottleState` counters.

### `AuditLog.ps1`
Compliance-grade session tracking.

| Function | What it does |
|---|---|
| `Start-AuditSession` | Creates `$script:AuditSession` with GUID, timestamps, tenant info, current user. |
| `Write-AuditEvent` | Appends typed event `{Timestamp, EventType, Detail, AffectedObject}` to session. Increments `ErrorCount` on Error events. Also calls `Write-ActivityLog`. |
| `Complete-AuditSession` | Calculates duration, captures final metrics, writes JSON file to `Logs/audit_<timestamp>_<shortId>.json`. |
| `Get-AuditSession` | Returns `$script:AuditSession`. |

### `Checkpoint.ps1`
Resume support for long-running analyses.

| Function | What it does |
|---|---|
| `Start-Checkpoint` | Creates `$script:CheckpointData`, writes to `Logs/checkpoint_<type>.json`. |
| `Update-Checkpoint` | Updates `Phase`, `ProcessedItems[key]`, `TotalItems[key]`, saves to disk. |
| `Complete-Checkpoint` | Marks status, saves, then **deletes** the file on success (so it won't resume a completed run). |
| `Get-Checkpoint` | Reads checkpoint file; returns data only if `Status == "InProgress"`. |

---

## Backend — Functions/SharePoint/

### `SPOConnection.ps1`

| Function | What it does |
|---|---|
| `Test-PnPModuleAvailable` | Three-method check: `Get-Module -ListAvailable`, `Import-Module`, then command availability. Returns `$true` if PnP 2.x+ found. |
| `Test-UserCapabilities` | Calls `Get-PnPTenantSite` and `Invoke-PnPGraphMethod` to detect four capability flags: `CanEnumerateSites`, `CanReadUsers`, `CanAccessStorageData`, `CanReadExternalUsers`. Each test is independent with a 500 ms delay between them. |

### `SiteCollector.ps1`

#### `Get-RealSites-DataDriven`
Runs inside a background runspace.

**Flow:**
1. `Clear-SharePointData("Sites")`, `Start-Checkpoint`, `Reset-ThrottleStats`, `Start-AuditSession`.
2. Tries `Get-PnPTenantSite` on current connection. If denied, reconnects to `-admin.sharepoint.com` and retries with `-Detailed`. If still failing, falls back to current site only via `Get-PnPWeb` + `Get-PnPSite`.
3. Processes up to 25 sites: extracts storage MB (from `StorageUsageCurrent`, `Usage.Storage`, or per-site connect), calls `Add-SharePointSite` for each.
4. Logs storage summary, calls `Complete-Checkpoint`, `Complete-AuditSession`.

### `PermissionsCollector.ps1`

#### `Get-RealPermissions-DataDriven -SiteUrl`
Runs inside a background runspace.

**Flow:**
1. `Clear-SharePointData("All")`, checkpoint, throttle reset, audit session.
2. Connects to `$SiteUrl` if not already connected there.
3. Gets site info + storage via `Get-PnPSite -Includes Usage`. Falls back to admin connection if storage unavailable.
4. Calls `Add-SharePointSite` for the analyzed site.
5. **Users:** `Get-PnPUser`, filters out system/app accounts, calls `Add-SharePointUser`.
6. **Groups:** `Get-PnPGroup`, filters out SharingLinks/LimitedAccess groups, gets member counts, calls `Add-SharePointGroup`.
7. **Role assignments (site level):** `Get-PnPProperty RoleAssignments` on the web object. For each, loads `Member` + `RoleDefinitionBindings`, skips "Limited Access", calls `Add-SharePointRoleAssignment`.
8. **Inheritance:** Adds site-level entry, then `Get-PnPList` for all visible lists. For lists with `HasUniqueRoleAssignments`, also captures list-level role assignments and adds them.
9. **Sharing links:** Finds `SharingLinks.*` groups via `Get-PnPGroup`, parses link type from group name (AnonymousView/AnonymousEdit/OrganizationView/OrganizationEdit/Flexible), calls `Add-SharePointSharingLink`.
10. Completes checkpoint and audit session.

### `PermissionsMatrix.ps1`

#### `Get-SitePermissionsMatrix -SiteUrl -ScanType`
Called synchronously from `Handle-PostBuildPermissionsMatrix` (not in background).

**Flow:**
1. Gets site web + role assignments → builds site node.
2. Iterates visible lists with `HasUniqueRoleAssignments` loaded.
3. For each list with unique perms, loads `RoleAssignments`.
4. **Quick scan:** only descends into items if list has unique perms, and only processes items where `HasUniqueRoleAssignments` is true.
5. **Full scan:** processes every item in every list.
6. Returns `{totalItems, uniquePermissions, totalPrincipals, tree, scanType, scannedAt}`.

---

## Backend — Functions/Analysis/

### `GraphEnrichment.ps1`

| Function | What it does |
|---|---|
| `Invoke-ExternalUserEnrichment` | Iterates external users from the data store. For each, calls `Get-GraphUserData`. Mutates user objects in-place: adds `GraphUserType`, `GraphAccountEnabled`, `GraphLastSignIn`, `GraphCreatedDate`, `GraphDisplayName`, `GraphEnriched`. 100 ms delay between users. |
| `Test-GraphAccess` | `GET v1.0/me` — returns `$true` if Graph is reachable. |
| `Get-GraphUserData` | Tries `GET v1.0/users?$filter=mail eq '...' or userPrincipalName eq '...'`, falls back to direct `GET v1.0/users/{email}`. Returns `{DisplayName, UserType, AccountEnabled, CreatedDateTime, LastSignIn}` or `$null`. |
| `Get-EnrichmentSummary` | Counts enriched users, disabled accounts (`GraphAccountEnabled = false`), guests (`GraphUserType = "Guest"`), stale accounts (no sign-in in 90 days or created 90+ days ago with no sign-in). |

### `RiskScoring.ps1`

#### `Get-RiskAssessment`
Loads all data types, evaluates rules, returns scored findings.

| Rule ID | Trigger | Severity |
|---|---|---|
| EXT-001 | External users with Edit/Contribute/Full Control | High |
| EXT-002 | External users who are site admins | Critical |
| EXT-003 | External users from >5 domains | Medium |
| SHARE-001 | Anonymous links with Edit access | Critical |
| SHARE-002 | Any anonymous links | High |
| SHARE-003 | >10 company-wide links | Medium |
| PERM-001 | >5 Full Control role assignments | High |
| PERM-002 | >10 direct-user role assignments | Medium |
| INH-001 | >50% of items have broken inheritance | High |
| INH-002 | >25% broken inheritance | Medium |
| GRP-001 | Empty groups | Low |

Overall score = average of top-5 finding scores, capped at 100. Risk level: ≥80 Critical, ≥60 High, ≥30 Medium, >0 Low, 0 None.

### `JsonExport.ps1`

| Function | What it does |
|---|---|
| `Export-GovernanceJson` | Calls `Build-GovernanceReport`, writes to `Reports/Generated/spo_governance_<timestamp>.json`. |
| `Build-GovernanceReport` | Assembles ordered hashtable: `schemaVersion`, `exportedAt`, `metadata` (tool info, tenant, scan time), `summary` (metrics), then arrays for `sites`, `users`, `groups`, `roleAssignments`, `inheritance`, `sharingLinks`. Schema version `1.0.0`. |

---

## Backend — Functions/Demo/

### `DemoDataGenerator.ps1`

| Function | What it does |
|---|---|
| `New-DemoData` | Clears all data, then populates 5 sites, 22 users (10 internal + 12 external across 9 domains), 9 groups (3 empty), 26 role assignments, 10 inheritance items, 20 sharing links. Data is crafted to trigger all risk rules. |
| `Invoke-DemoEnrichment` | Marks all external users as enriched with randomized last sign-in dates (1–120 days ago). |
| `Get-DemoPermissionsMatrix` | Returns a hardcoded realistic permissions tree (Site → Libraries → Folders/Files) for any given `$SiteUrl`. |

Also defines `$script:DataTypeMap` — URL slug to data manager key mapping used by data/export handlers.

---

## Frontend — Web/js/

### `app.js` — Entry point & tab router
`DOMContentLoaded` calls `initTabs()`, `initConnection()`, `initOperations()`, `initAnalytics()`, `initGlobalSearch()`, `initExportModal()`, `pollStatus()`.

- **`initTabs`** — wires `.tab-btn` clicks; switching to the analytics tab triggers `refreshAnalytics()`. Hides Operations/Analytics tabs until connected.
- **`pollStatus`** — one-shot on load: calls `API.getStatus()`, restores UI if server already has a session (page refresh recovery).
- **`updateTabVisibility`** — shows/hides Operations and Analytics tab buttons and enables/disables the global search input.

---

### `app-state.js` — Shared state & utilities
**`appState`** object: `connected`, `demoMode`, `dataLoaded`, `headless`, `capabilities`, `connectedSiteUrl`.

Utility functions available globally:
- `setText(id, value)` — safe `textContent` setter.
- `esc(str)` — HTML-escapes a string (via `div.textContent`).
- `formatStorage(mb)` — returns `"X MB"` or `"X.X GB"`.
- `toast(message, type)` — appends a self-removing toast div to `#toast-container` (4 s).
- `pollUntilComplete(consoleEl, intervalMs)` — polls `API.getProgress()` every `intervalMs` ms, writes messages to `consoleEl.textContent`, resolves when `complete && !running`.
- `showAuditSummary(consoleEl)` — appends audit trail info to the console element after an operation.

---

### `api.js` — HTTP client
`API` object with two base methods (`get`, `post`) and named wrappers for every endpoint:

| Method | Endpoint |
|---|---|
| `getStatus()` | `GET /api/status` |
| `connect(tenantUrl, clientId)` | `POST /api/connect` |
| `startDemo()` | `POST /api/demo` |
| `getSites()` | `POST /api/sites` |
| `prepareAnalysis(siteUrl)` | `POST /api/prepare-analysis` |
| `analyzePermissions(siteUrl)` | `POST /api/permissions` |
| `getProgress()` | `GET /api/progress` |
| `getData(type)` | `GET /api/data/:type` |
| `getMetrics()` | `GET /api/metrics` |
| `enrichExternal()` | `POST /api/enrich` |
| `getEnrichment()` | `GET /api/enrichment` |
| `getRisk()` | `GET /api/risk` |
| `getAudit()` | `GET /api/audit` |
| `exportData(type)` | Opens `GET /api/export/:type` (CSV download) |
| `exportDataJson(type)` | Opens `GET /api/export-json/:type` (JSON download) |
| `exportJson()` | `GET /api/export-json` |
| `buildPermissionsMatrix(siteUrl, scanType)` | `POST /api/build-permissions-matrix` |

---

### `connection.js` — Connection tab
- `initConnection` — wires `#btn-connect` and `#btn-demo`.
- `handleConnect` — reads tenant URL + client ID, calls `API.connect()`, on success updates `appState`, calls `displayCapabilities()`, `updateConnectionUI()`, `updateTabVisibility()`, `updateOperationsButtons()`.
- `handleDemo` — calls `API.startDemo()`, sets full capabilities in `appState`, calls `refreshAnalytics()`.
- `updateConnectionUI` — updates the status dot and text in the header.
- `displayCapabilities` — renders the four capability flags into `#capability-status`.
- `updateOperationsButtons` — disables `#btn-get-sites` with tooltip if `CanEnumerateSites` is false.

---

### `operations.js` — Operations tab
- `initOperations` — wires `#btn-get-sites`, `#btn-analyze`, `#btn-report`.
- `handleGetSites` — calls `API.getSites()`, polls until complete, then calls `API.getData('sites')` and renders site list into the console, then `refreshAnalytics()`.
- `handleAnalyze` — two-step: (1) `API.prepareAnalysis()` to check re-auth, shows device-login instructions if needed, (2) `API.analyzePermissions()`, polls until complete, updates `appState.dataLoaded`, shows metrics summary.
- `handleReport` — calls `showExportModal('all', true)`.

---

### `analytics.js` — Analytics tab & risk
- `initAnalytics` — makes metric cards clickable; clicks call `openDeepDive(card.dataset.deepdive)`.
- `refreshAnalytics` — fetches metrics, animates counters, fetches sites/users/groups, calls `renderStorageChart`, `renderPermissionChart`, `renderSitesTable`, `renderAlerts`, `refreshRiskBanner`.
- `refreshRiskBanner` — calls `API.getRisk()`, updates `#risk-banner` class + score/level/summary text, wires "View Details" button to `openRiskDeepDive`.
- `openRiskDeepDive` — opens the modal with finding cards, filter buttons per severity, expand-on-click detail toggle.
- `renderSitesTable` — fills `#sites-table-body` with colour-coded storage usage badges.
- `renderAlerts` — generates alert items in `#alerts-container` based on metric thresholds.

---

### `charts.js` — Chart.js wrappers
- `renderStorageChart(sites)` — destroys old chart, renders bar chart in `#chart-storage`. Top 10 sites by storage, gradient colours by tier (green/orange/red/purple). Clicking a bar opens that site's deep dive.
- `renderPermissionChart(users, groups)` — destroys old chart, renders doughnut in `#chart-permissions`. Aggregates `Permission`/`Role` field across users + groups. Clicking a segment opens the permissions deep dive filtered to that role.
- `renderDeepDiveChart(canvasId, type, data)` — generic chart renderer for deep-dive modals. Destroys existing chart on canvas, renders bar or doughnut.

**`COLORS`** — named colour constants for permission levels and chart datasets (matches design system).

---

### `deep-dives.js` — Modal deep dives
- `openDeepDive(type)` — opens the modal, fetches the relevant data type, delegates to one of the render functions below.

| Render Function | Data | What it shows |
|---|---|---|
| `renderSitesDeepDive` | Sites | Stats bar, search filter, table with "🔍 Matrix" button per row. |
| `renderUsersDeepDive` | Users | Stats bar, search + type filter, table of all users. |
| `renderGroupsDeepDive` | Groups | Stats bar, search filter, table with empty group count. |
| `renderExternalDeepDive` | Users (external only) | Domain analysis, "Enrich via Graph" button, account status + last sign-in columns. |
| `renderPermissionsDeepDive` | RoleAssignments | 3-tab view: table (with search + role filter), doughnut chart, security findings. |
| `renderInheritanceDeepDive` | InheritanceItems | 4-tab view: tree view (expandable site→list hierarchy), table, doughnut chart, findings. |
| `renderSharingDeepDive` | SharingLinks | 3-tab view: table (with link type filter), doughnut chart, findings. |

Helper functions:
- `buildInheritanceTree(data)` — transforms flat item list into `{site, children[]}` groups.
- `renderTreeView(treeData)` — builds HTML tree with expand/collapse via `toggleTreeNode`.
- `openSiteDetailDeepDive(siteName)` — opens sites deep dive with search pre-filled.
- `openFilteredPermissionsDeepDive(permissionLevel)` — opens permissions deep dive with role filter pre-selected.
- `showEnrichmentBanner` — fetches enrichment summary and renders disabled/stale account findings.

---

### `permissions-matrix.js` — Permissions matrix modal
- `openPermissionsMatrix(siteUrl, siteTitle)` — opens `#matrix-modal`, shows scan type chooser.
- `buildMatrix(siteUrl, scanType)` — calls `API.buildPermissionsMatrix()`, renders result.
- `renderPermissionsMatrix(data)` — shows stats bar + toolbar, calls `renderMatrixTree`.
- `renderMatrixTree(nodes, level)` — recursive HTML renderer; indents by `level * 20 px`, shows permission badges per node.
- `renderPermissionBadges(permissions)` — renders `principal: role` badges, or "Inherited" if empty.
- `exportMatrixToCSV(matrixData)` / `exportMatrixToJSON(matrixData)` — client-side file generation via `Blob` + `URL.createObjectURL`.

---

### `search.js` — Global search (Ctrl+K)
- `initGlobalSearch` — wires keyboard shortcut, input debounce (300 ms), keyboard navigation (arrow keys + enter), lazy data load on focus, click-outside close.
- `loadGlobalSearchData` — fetches sites, users, groups, roleassignments, inheritance in parallel; stores in `globalSearchData`.
- `performGlobalSearch(query)` — filters each data type (up to 5 results each), calls `renderSearchResults`.
- `renderSearchResults` — renders grouped dropdown with type sections (Sites/Users/Groups/Permissions/Inheritance).
- `navigateToSearchResult(type, item)` — switches to analytics tab, then opens the appropriate deep dive with the item pre-selected or filter pre-filled.

---

### `export.js` — Export modal
- `showExportModal(type, isFullReport)` — opens `#export-modal`, stores pending type.
- `initExportModal` — wires `.export-format-btn` clicks to `handleDataExport` or `handleReportExport`.
- `handleDataExport(type, format)` — for `permissions-matrix` calls client-side export functions; for other types calls `API.exportData` (CSV) or `API.exportDataJson` (JSON).
- `handleReportExport(format)` — JSON: fetches full governance report via `API.exportJson()`, creates blob download. CSV: triggers `API.exportData` for all 6 data types.

---

### `ui-helpers.js` — UI utilities (`UIHelpers` object)

| Method | What it does |
|---|---|
| `setButtonLoading(id, loading)` | Adds/removes `.btn-loading` class and disables button. |
| `showSkeleton(id, type)` | Injects skeleton placeholder HTML (text/card/metric/table). |
| `showMetricSkeletons` | Injects skeleton in each metric card element. |
| `showLoadingOverlay / hideLoadingOverlay` | Full-page spinner overlay. |
| `updateProgress(id, percent)` | Creates/updates a `.progress-bar` inside the container. |
| `showIndeterminateProgress(id)` | Shows CSS-animated indeterminate progress bar. |
| `setInputValidation(id, isValid, msg)` | Adds `.input-valid`/`.input-error` + feedback message. |
| `showEmptyState(id, title, msg, icon)` | Renders empty state placeholder HTML. |
| `createBadge(text, variant, withDot)` | Returns a badge `<span>`. |
| `makeSortable(tableId)` | Wires `th.sortable` click handlers for numeric/string sort. |
| `filterTable(tableId, term, cols)` | Shows/hides rows by search term. |
| `animateCounter(id, target, duration)` | Counts up a number over time using `setInterval`. |
| `highlightSearchTerm(text, term)` | Returns HTML with `<span class="search-highlight">` around matches. |
| `debounce(func, wait)` / `throttle(func, limit)` | Standard debounce/throttle wrappers. |
| `copyToClipboard(text)` | `navigator.clipboard.writeText` + success toast. |
| `formatRelativeDate(date)` | Returns "just now" / "Xm ago" / "Xh ago" / "Xd ago" / locale date. |
| `createRipple(element, event)` | Appends ripple `<span>` on click (auto-registered on `.btn`). |

---

## Data Flow Summary

```
User clicks "Analyze Permissions"
  → operations.js: handleAnalyze()
    → API.prepareAnalysis(siteUrl)                    [POST /api/prepare-analysis]
    → API.analyzePermissions(siteUrl)                 [POST /api/permissions]
      → ApiHandlers: Handle-PostPermissions
        → Start-BackgroundOperation { Get-RealPermissions-DataDriven }
      ← { started: true }
    → pollUntilComplete()                             [GET /api/progress × N]
      ← { complete: true }
    → API.getMetrics()                                [GET /api/metrics]
    → refreshAnalytics()
      → API.getMetrics(), getData(sites/users/groups) [GET /api/metrics, /api/data/*]
      → renderStorageChart(), renderPermissionChart()
      → refreshRiskBanner()
        → API.getRisk()                               [GET /api/risk]
          → Get-RiskAssessment() → evaluates rules → scored findings
```

---

## Key File Locations Quick Reference

| What you need | File |
|---|---|
| HTTP server loop | `Functions/Server/WebServer.ps1:7` |
| All API routes | `Functions/Server/ApiHandlers.ps1:21` |
| Background runspace logic | `Functions/Server/BackgroundJobManager.ps1:8` |
| In-memory data store | `Functions/Core/SharePointDataManager.ps1` |
| Connection state flags | `Functions/Core/Settings.ps1:15` |
| PnP module check / capability test | `Functions/SharePoint/SPOConnection.ps1` |
| Site enumeration | `Functions/SharePoint/SiteCollector.ps1:8` |
| Full permissions analysis | `Functions/SharePoint/PermissionsCollector.ps1:8` |
| Permissions matrix (tree) | `Functions/SharePoint/PermissionsMatrix.ps1:1` |
| Graph enrichment | `Functions/Analysis/GraphEnrichment.ps1` |
| Risk scoring rules | `Functions/Analysis/RiskScoring.ps1:21` |
| JSON export schema | `Functions/Analysis/JsonExport.ps1:41` |
| Demo data + DataTypeMap | `Functions/Demo/DemoDataGenerator.ps1` |
| Throttle retry wrapper | `Functions/Core/ThrottleProtection.ps1:13` |
| Audit session | `Functions/Core/AuditLog.ps1` |
| Checkpoint/resume | `Functions/Core/Checkpoint.ps1` |
| App state + polling helper | `Web/js/app-state.js` |
| All API fetch wrappers | `Web/js/api.js` |
| Tab routing + startup | `Web/js/app.js` |
| Connection UI + capabilities | `Web/js/connection.js` |
| Operations UI | `Web/js/operations.js` |
| Analytics + risk banner | `Web/js/analytics.js` |
| Chart rendering | `Web/js/charts.js` |
| Deep dive modals | `Web/js/deep-dives.js` |
| Permissions matrix modal | `Web/js/permissions-matrix.js` |
| Global search (Ctrl+K) | `Web/js/search.js` |
| Export modal | `Web/js/export.js` |
| UI component helpers | `Web/js/ui-helpers.js` |
