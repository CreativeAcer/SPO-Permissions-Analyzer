// ============================================
// app.js - Main application logic
// ============================================

// --- State ---
let appState = {
    connected: false,
    demoMode: false,
    dataLoaded: false,
    headless: false
};

// Permissions matrix state
let currentMatrixData = null;
let currentMatrixSiteUrl = null;

// --- Initialization ---
document.addEventListener('DOMContentLoaded', () => {
    initTabs();
    initConnection();
    initOperations();
    initAnalytics();
    initGlobalSearch();
    initExportModal();
    pollStatus();
});

// --- Tabs ---
function initTabs() {
    // Hide operations and analytics tabs initially
    updateTabVisibility(false);

    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.addEventListener('click', async () => {
            document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
            document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
            btn.classList.add('active');
            const target = document.getElementById('tab-' + btn.dataset.tab);
            if (target) target.classList.add('active');

            // Refresh analytics when switching to analytics tab
            if (btn.dataset.tab === 'analytics' && appState.dataLoaded) {
                await refreshAnalytics();
            }
        });
    });
}

// Show/hide tabs based on connection state
function updateTabVisibility(connected) {
    const operationsTab = document.querySelector('.tab-btn[data-tab="operations"]');
    const analyticsTab = document.querySelector('.tab-btn[data-tab="analytics"]');
    const searchInput = document.getElementById('global-search-input');

    if (connected) {
        if (operationsTab) operationsTab.style.display = 'block';
        if (analyticsTab) analyticsTab.style.display = 'block';
        if (searchInput) {
            searchInput.disabled = false;
            searchInput.placeholder = 'Search sites, users, groups... (Ctrl+K)';
        }
    } else {
        if (operationsTab) operationsTab.style.display = 'none';
        if (analyticsTab) analyticsTab.style.display = 'none';
        if (searchInput) {
            searchInput.disabled = true;
            searchInput.placeholder = 'Connect to SharePoint to enable search';
        }
    }
}

// --- Connection Tab ---
function initConnection() {
    document.getElementById('btn-connect').addEventListener('click', handleConnect);
    document.getElementById('btn-demo').addEventListener('click', handleDemo);
}

async function handleConnect() {
    const tenantUrl = document.getElementById('input-tenant-url').value.trim();
    const clientId = document.getElementById('input-client-id').value.trim();
    const results = document.getElementById('connection-results');

    if (!tenantUrl || !clientId) {
        results.textContent = 'Please enter both Tenant URL and Client ID.';
        return;
    }

    if (appState.headless) {
        results.textContent = 'Connecting via device code flow...\n\n'
            + 'Check the container terminal for the authentication code.\n'
            + '(Run "podman logs <container>" or check the terminal where compose is running)\n\n'
            + 'Open https://microsoft.com/devicelogin and enter the code shown there.\n\n'
            + 'Waiting for authentication...';
    } else {
        results.textContent = 'Connecting to SharePoint Online...\nPlease complete authentication in the popup window.';
    }
    setButtonLoading('btn-connect', true);

    try {
        const res = await API.connect(tenantUrl, clientId);
        if (res.success) {
            appState.connected = true;
            results.textContent = `Connected successfully!\n\nSite: ${res.siteTitle || 'N/A'}\nURL: ${res.siteUrl || 'N/A'}\nUser: ${res.user || 'N/A'}\n\nYou can now use SharePoint Operations.`;
            updateConnectionUI(true);
            updateTabVisibility(true);
            toast('Connected to SharePoint', 'success');
        } else {
            results.textContent = `Connection failed: ${res.message}`;
            toast(res.message, 'error');
        }
    } catch (e) {
        results.textContent = `Error: ${e.message}`;
        toast('Connection failed', 'error');
    } finally {
        setButtonLoading('btn-connect', false);
    }
}

async function handleDemo() {
    const results = document.getElementById('connection-results');
    results.textContent = 'Starting Demo Mode...\nGenerating sample SharePoint data...';
    setButtonLoading('btn-demo', true);

    try {
        const res = await API.startDemo();
        if (res.success) {
            appState.connected = true;
            appState.demoMode = true;
            appState.dataLoaded = true;
            results.textContent = 'Demo Mode activated!\n\nSample data has been generated.\nSwitch to Operations or Visual Analytics tab to explore.';
            updateConnectionUI(true);
            updateTabVisibility(true);
            toast('Demo mode activated', 'success');
            await refreshAnalytics();
        } else {
            results.textContent = `Demo mode failed: ${res.message}`;
        }
    } catch (e) {
        results.textContent = `Error: ${e.message}`;
    } finally {
        setButtonLoading('btn-demo', false);
    }
}

function updateConnectionUI(connected) {
    const indicator = document.getElementById('connection-indicator');
    const dot = indicator.querySelector('.status-dot');
    const text = indicator.querySelector('span:last-child');

    if (connected) {
        dot.classList.remove('disconnected');
        dot.classList.add('connected');
        text.textContent = appState.demoMode ? 'Demo Mode' : 'Connected';
        document.getElementById('operations-status').textContent = appState.demoMode ? 'Demo Mode Active' : 'Connected to SharePoint';
        document.getElementById('operations-status').style.color = '#28A745';
        // Enable operation buttons
        document.querySelectorAll('#tab-operations .btn-disabled').forEach(btn => {
            btn.classList.remove('btn-disabled');
        });
    } else {
        dot.classList.remove('connected');
        dot.classList.add('disconnected');
        text.textContent = 'Not connected';
    }
}

// --- Operations Tab ---
function initOperations() {
    document.getElementById('btn-get-sites').addEventListener('click', handleGetSites);
    document.getElementById('btn-analyze').addEventListener('click', handleAnalyze);
    document.getElementById('btn-report').addEventListener('click', handleReport);
}

async function handleGetSites() {
    const console_ = document.getElementById('operations-console');
    console_.textContent = 'Fetching sites...\n';
    setButtonLoading('btn-get-sites', true);

    try {
        const res = await API.getSites();
        if (res.started) {
            // Background operation — poll until complete
            const progress = await pollUntilComplete(console_);
            if (progress.error) {
                console_.textContent += `\nError: ${progress.error}`;
                toast('Site retrieval failed', 'error');
                return;
            }
        } else if (!res.success) {
            console_.textContent += `\nError: ${res.message}`;
            return;
        }

        const sites = await API.getData('sites');
        const siteList = sites.data || [];
        console_.textContent += `\nRetrieved ${siteList.length} sites:\n`;
        siteList.forEach((s, i) => {
            console_.textContent += `\n${i + 1}. ${s.Title || 'Unknown'}\n   URL: ${s.Url || 'N/A'}\n   Owner: ${s.Owner || 'N/A'}\n   Storage: ${s.Storage || '0'} MB\n`;
        });
        await refreshAnalytics();
        await showAuditSummary(console_);
        toast(`Retrieved ${siteList.length} sites`, 'success');
    } catch (e) {
        console_.textContent += `\nError: ${e.message}`;
    } finally {
        setButtonLoading('btn-get-sites', false);
    }
}

async function handleAnalyze() {
    const siteUrl = document.getElementById('input-site-url').value.trim();
    const console_ = document.getElementById('operations-console');
    console_.textContent = 'Starting permissions analysis...\n';
    setButtonLoading('btn-analyze', true);

    try {
        const res = await API.analyzePermissions(siteUrl);
        if (res.started) {
            // Background operation — poll until complete
            const progress = await pollUntilComplete(console_);
            if (progress.error) {
                console_.textContent += `\nError: ${progress.error}`;
                toast('Permissions analysis failed', 'error');
                return;
            }
        } else if (!res.success) {
            console_.textContent += `\nError: ${res.message}`;
            return;
        }

        appState.dataLoaded = true;
        // Append final metrics
        const metrics = await API.getMetrics();
        console_.textContent += '\n\n=== ANALYSIS COMPLETE ===';
        console_.textContent += `\nUsers: ${metrics.totalUsers} | Groups: ${metrics.totalGroups} | External: ${metrics.externalUsers}`;
        console_.textContent += `\nRole Assignments: ${metrics.totalRoleAssignments} | Inheritance Breaks: ${metrics.inheritanceBreaks} | Sharing Links: ${metrics.totalSharingLinks}`;
        console_.textContent += '\n\nSwitch to Visual Analytics tab for charts and deep dives.';
        await refreshAnalytics();
        await showAuditSummary(console_);
        toast('Permissions analysis complete', 'success');
    } catch (e) {
        console_.textContent += `\nError: ${e.message}`;
    } finally {
        setButtonLoading('btn-analyze', false);
    }
}

async function handleReport() {
    if (!appState.dataLoaded) {
        toast('Run an analysis first', 'info');
        return;
    }

    // Show export format modal for full report
    showExportModal('all', true);
}

// --- Analytics Tab ---
function initAnalytics() {
    // Make metric cards clickable
    document.querySelectorAll('.metric-card[data-deepdive]').forEach(card => {
        card.addEventListener('click', () => {
            if (!appState.dataLoaded) {
                toast('Run an analysis first to view deep dives', 'info');
                return;
            }
            openDeepDive(card.dataset.deepdive);
        });
    });
}

async function refreshAnalytics() {
    try {
        // Show skeleton loaders if UIHelpers available
        if (typeof UIHelpers !== 'undefined') {
            UIHelpers.showMetricSkeletons();
        }

        const metrics = await API.getMetrics();

        // Update metric cards with animation if available
        const metricMap = {
            'metric-sites': metrics.totalSites,
            'metric-users': metrics.totalUsers,
            'metric-groups': metrics.totalGroups,
            'metric-external': metrics.externalUsers,
            'metric-roles': metrics.totalRoleAssignments,
            'metric-inheritance': metrics.inheritanceBreaks,
            'metric-sharing': metrics.totalSharingLinks
        };

        // Animate counters if UIHelpers available, otherwise just set text
        if (typeof UIHelpers !== 'undefined') {
            Object.entries(metricMap).forEach(([id, value]) => {
                UIHelpers.animateCounter(id, value, 800);
            });
        } else {
            Object.entries(metricMap).forEach(([id, value]) => {
                setText(id, value);
            });
        }

        setText('analytics-subtitle', `Last updated: ${new Date().toLocaleString()}`);

        // Fetch data for charts
        const [sitesRes, usersRes, groupsRes] = await Promise.all([
            API.getData('sites'),
            API.getData('users'),
            API.getData('groups')
        ]);

        renderStorageChart(sitesRes.data);
        renderPermissionChart(usersRes.data, groupsRes.data);

        // Update sites table
        renderSitesTable(sitesRes.data);

        // Generate alerts
        renderAlerts(metrics, usersRes.data);

        // Risk assessment
        await refreshRiskBanner();

    } catch (e) {
        console.error('Failed to refresh analytics:', e);
    }
}

async function refreshRiskBanner() {
    console.log('[Risk] Refreshing risk banner...');
    const banner = document.getElementById('risk-banner');

    if (!banner) {
        console.error('[Risk] Risk banner element not found in DOM');
        return;
    }

    try {
        console.log('[Risk] Fetching risk assessment from API...');
        const risk = await API.getRisk();
        console.log('[Risk] Risk assessment received:', risk);

        // Handle both camelCase (overallScore) and PascalCase (OverallScore) from backend
        const overallScore = risk.overallScore ?? risk.OverallScore ?? 0;

        // Handle riskLevel - it might be a string or an array (PowerShell switch bug)
        let riskLevelRaw = risk.riskLevel ?? risk.RiskLevel ?? 'Unknown';
        const riskLevel = Array.isArray(riskLevelRaw) ? riskLevelRaw[0] : riskLevelRaw;

        const totalFindings = risk.totalFindings ?? risk.TotalFindings ?? 0;
        const criticalCount = risk.criticalCount ?? risk.CriticalCount ?? 0;
        const highCount = risk.highCount ?? risk.HighCount ?? 0;
        const mediumCount = risk.mediumCount ?? risk.MediumCount ?? 0;
        const lowCount = risk.lowCount ?? risk.LowCount ?? 0;

        console.log('[Risk] Parsed values:', { overallScore, riskLevel, totalFindings });

        // Always show the banner (remove hidden)
        banner.classList.remove('hidden', 'risk-critical', 'risk-high', 'risk-medium', 'risk-low', 'risk-none');
        banner.classList.add('risk-' + riskLevel.toLowerCase());

        setText('risk-score-value', overallScore);
        setText('risk-level', riskLevel);

        const parts = [];
        if (criticalCount > 0) parts.push(`${criticalCount} critical`);
        if (highCount > 0) parts.push(`${highCount} high`);
        if (mediumCount > 0) parts.push(`${mediumCount} medium`);
        if (lowCount > 0) parts.push(`${lowCount} low`);
        setText('risk-summary', parts.length > 0
            ? `${totalFindings} finding(s): ${parts.join(', ')}`
            : 'No security findings detected');

        // Wire up details button
        const btn = document.getElementById('btn-risk-details');
        if (btn) {
            btn.onclick = () => openRiskDeepDive(risk);
        }

        // Store risk data for reuse
        appState.riskData = risk;
        console.log('[Risk] Risk banner updated successfully');
    } catch (e) {
        console.error('[Risk] Failed to load risk assessment:', e);
        // Show error to user with toast
        if (typeof toast === 'function') {
            toast('Risk assessment unavailable: ' + e.message, 'error');
        }
        // Still show banner with zero state
        banner.classList.remove('hidden', 'risk-critical', 'risk-high', 'risk-medium', 'risk-low', 'risk-none');
        banner.classList.add('risk-none');
        setText('risk-score-value', '0');
        setText('risk-level', 'Unknown');
        setText('risk-summary', 'Risk assessment unavailable');
        console.log('[Risk] Risk banner set to error state');
    }
}

function openRiskDeepDive(risk) {
    const overlay = document.getElementById('modal-overlay');
    const title = document.getElementById('modal-title');
    const body = document.getElementById('modal-body');

    overlay.classList.remove('hidden');
    document.getElementById('modal-close').onclick = () => overlay.classList.add('hidden');
    overlay.onclick = (e) => { if (e.target === overlay) overlay.classList.add('hidden'); };
    const escHandler = (e) => { if (e.key === 'Escape') { overlay.classList.add('hidden'); document.removeEventListener('keydown', escHandler); } };
    document.addEventListener('keydown', escHandler);

    title.textContent = `Risk Assessment (Score: ${risk.overallScore}/100)`;

    const severityColors = { Critical: '#DC3545', High: '#E65100', Medium: '#FFC107', Low: '#28A745' };

    // Stats section
    let html = `<div class="dd-stats">
        <div class="dd-stat"><span class="dd-stat-value" style="color:${severityColors[risk.riskLevel] || '#6C757D'}">${risk.overallScore}</span><span class="dd-stat-label">Risk Score</span></div>
        <div class="dd-stat"><span class="dd-stat-value" style="color:#DC3545">${risk.criticalCount}</span><span class="dd-stat-label">Critical</span></div>
        <div class="dd-stat"><span class="dd-stat-value" style="color:#E65100">${risk.highCount}</span><span class="dd-stat-label">High</span></div>
        <div class="dd-stat"><span class="dd-stat-value" style="color:#FFC107">${risk.mediumCount}</span><span class="dd-stat-label">Medium</span></div>
        <div class="dd-stat"><span class="dd-stat-value" style="color:#28A745">${risk.lowCount}</span><span class="dd-stat-label">Low</span></div>
    </div>`;

    // Filter buttons
    html += `<div style="margin: 16px 0; display: flex; gap: 8px; flex-wrap: wrap;">
        <button class="btn btn-secondary" onclick="filterRiskFindings('all')" style="padding: 6px 12px; font-size: 13px;">All (${risk.totalFindings})</button>
        <button class="btn btn-secondary" onclick="filterRiskFindings('Critical')" style="padding: 6px 12px; font-size: 13px; background: #DC3545; color: white;">Critical (${risk.criticalCount})</button>
        <button class="btn btn-secondary" onclick="filterRiskFindings('High')" style="padding: 6px 12px; font-size: 13px; background: #E65100; color: white;">High (${risk.highCount})</button>
        <button class="btn btn-secondary" onclick="filterRiskFindings('Medium')" style="padding: 6px 12px; font-size: 13px; background: #FFC107; color: white;">Medium (${risk.mediumCount})</button>
        <button class="btn btn-secondary" onclick="filterRiskFindings('Low')" style="padding: 6px 12px; font-size: 13px; background: #28A745; color: white;">Low (${risk.lowCount})</button>
    </div>`;

    // Findings container
    html += '<div id="risk-findings-container">';
    if (risk.findings && risk.findings.length > 0) {
        html += risk.findings.map((f, idx) => {
            const sev = f.Severity || f.severity;
            const category = f.Category || f.category || 'General';
            const count = f.Count || f.count || 0;
            const color = severityColors[sev] || '#6C757D';
            return `<div class="finding finding-item" data-severity="${sev}" style="border-left: 4px solid ${color}; margin-bottom: 12px; padding: 12px; background: white; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); cursor: pointer; transition: all 0.2s;" onclick="toggleFindingDetails(${idx})">
                <div style="display: flex; justify-content: space-between; align-items: start;">
                    <h4 style="margin: 0 0 8px 0; font-size: 15px;">[${esc(f.RuleId || f.ruleId)}] ${esc(f.Title || f.title)}</h4>
                    <span style="background: ${color}; color: white; padding: 2px 8px; border-radius: 4px; font-size: 11px; font-weight: 600;">${sev}</span>
                </div>
                <div style="display: flex; gap: 16px; margin-bottom: 8px; font-size: 12px; color: #666;">
                    <span><strong>Category:</strong> ${esc(category)}</span>
                    ${count > 0 ? `<span><strong>Affected Items:</strong> ${count}</span>` : ''}
                    <span><strong>Score:</strong> ${f.Score || f.score}/100</span>
                </div>
                <p style="margin: 0; color: #333;">${esc(f.Description || f.description)}</p>
                <div id="finding-details-${idx}" style="display: none; margin-top: 12px; padding-top: 12px; border-top: 1px solid #E5E7EB;">
                    <p style="margin: 0; font-size: 13px; color: #666;"><strong>Recommendation:</strong> Review and remediate this finding to improve your security posture.</p>
                </div>
            </div>`;
        }).join('');
    } else {
        html += '<div class="finding low"><h4>No Issues Found</h4><p>No security findings detected. Your environment looks clean.</p></div>';
    }
    html += '</div>';

    body.innerHTML = html;

    // Store findings for filtering
    window.currentRiskFindings = risk.findings || [];
}

function filterRiskFindings(severity) {
    const findings = document.querySelectorAll('.finding-item');
    findings.forEach(f => {
        if (severity === 'all' || f.dataset.severity === severity) {
            f.style.display = 'block';
        } else {
            f.style.display = 'none';
        }
    });
}

function toggleFindingDetails(idx) {
    const details = document.getElementById(`finding-details-${idx}`);
    if (details) {
        details.style.display = details.style.display === 'none' ? 'block' : 'none';
    }
}

function renderSitesTable(sites) {
    const tbody = document.getElementById('sites-table-body');
    if (!tbody) return;

    if (!sites || sites.length === 0) {
        tbody.innerHTML = '<tr><td colspan="5" class="text-center text-muted">No site data available</td></tr>';
        return;
    }

    tbody.innerHTML = sites.map(s => {
        const storage = parseInt(s.Storage) || 0;
        let usageClass = 'usage-low', usageText = 'Low';
        if (storage >= 1500) { usageClass = 'usage-critical'; usageText = 'Critical'; }
        else if (storage >= 1000) { usageClass = 'usage-high'; usageText = 'High'; }
        else if (storage >= 500) { usageClass = 'usage-medium'; usageText = 'Medium'; }

        return `<tr>
            <td>${esc(s.Title)}</td>
            <td>${esc(s.Url)}</td>
            <td>${esc(s.Owner)}</td>
            <td>${storage} MB</td>
            <td><span class="usage-badge ${usageClass}">${usageText}</span></td>
        </tr>`;
    }).join('');
}

function renderAlerts(metrics, users) {
    const container = document.getElementById('alerts-container');
    if (!container) return;

    const alerts = [];

    if (metrics.externalUsers > 0) {
        alerts.push({ type: 'warning', title: 'External Users Detected', detail: `${metrics.externalUsers} external user(s) have access to your SharePoint environment.` });
    }
    if (metrics.totalSharingLinks > 0) {
        alerts.push({ type: 'danger', title: 'Sharing Links Found', detail: `${metrics.totalSharingLinks} sharing link(s) detected. Review for anonymous or overly broad access.` });
    }
    if (metrics.inheritanceBreaks > 0) {
        alerts.push({ type: 'warning', title: 'Broken Inheritance', detail: `${metrics.inheritanceBreaks} item(s) have unique permissions (broken inheritance).` });
    }
    if (metrics.totalSites > 0) {
        alerts.push({ type: 'info', title: 'Analysis Complete', detail: `Analyzed ${metrics.totalSites} sites, ${metrics.totalUsers} users, ${metrics.totalGroups} groups.` });
    }
    if (alerts.length === 0) {
        alerts.push({ type: 'info', title: 'No Data', detail: 'Run an analysis to see security alerts here.' });
    }

    container.innerHTML = alerts.map(a => `
        <div class="alert-item ${a.type}">
            <h4>${esc(a.title)}</h4>
            <p>${esc(a.detail)}</p>
        </div>
    `).join('');
}

// --- Deep Dives ---
async function openDeepDive(type) {
    const overlay = document.getElementById('modal-overlay');
    const title = document.getElementById('modal-title');
    const body = document.getElementById('modal-body');

    overlay.classList.remove('hidden');
    document.getElementById('modal-close').onclick = () => overlay.classList.add('hidden');
    overlay.onclick = (e) => { if (e.target === overlay) overlay.classList.add('hidden'); };

    // Escape key closes modal
    const escHandler = (e) => { if (e.key === 'Escape') { overlay.classList.add('hidden'); document.removeEventListener('keydown', escHandler); } };
    document.addEventListener('keydown', escHandler);

    body.innerHTML = '<p class="text-center text-muted">Loading...</p>';

    const deepDiveConfig = {
        sites:       { title: 'Sites Deep Dive', dataType: 'sites', render: renderSitesDeepDive },
        users:       { title: 'Users Deep Dive', dataType: 'users', render: renderUsersDeepDive },
        groups:      { title: 'Groups Deep Dive', dataType: 'groups', render: renderGroupsDeepDive },
        external:    { title: 'External Users Deep Dive', dataType: 'users', render: renderExternalDeepDive },
        permissions: { title: 'Role Assignment Mapping', dataType: 'roleassignments', render: renderPermissionsDeepDive },
        inheritance: { title: 'Permission Inheritance Analysis', dataType: 'inheritance', render: renderInheritanceDeepDive },
        sharing:     { title: 'Sharing Links Security Audit', dataType: 'sharinglinks', render: renderSharingDeepDive }
    };

    const config = deepDiveConfig[type];
    if (!config) { body.innerHTML = '<p>Unknown deep dive type</p>'; return; }

    title.textContent = config.title;

    try {
        const res = await API.getData(config.dataType);
        config.render(body, res.data || []);
    } catch (e) {
        body.innerHTML = `<p class="text-center" style="color:#DC3545">Error loading data: ${esc(e.message)}</p>`;
    }
}

// --- Chart Drill-Down Helper Functions ---

// Open sites deep dive filtered to a specific site
window.openSiteDetailDeepDive = async function(siteName) {
    if (!appState.dataLoaded) {
        toast('Run an analysis first', 'info');
        return;
    }

    const overlay = document.getElementById('modal-overlay');
    const title = document.getElementById('modal-title');
    const body = document.getElementById('modal-body');

    overlay.classList.remove('hidden');
    document.getElementById('modal-close').onclick = () => overlay.classList.add('hidden');
    overlay.onclick = (e) => { if (e.target === overlay) overlay.classList.add('hidden'); };

    const escHandler = (e) => { if (e.key === 'Escape') { overlay.classList.add('hidden'); document.removeEventListener('keydown', escHandler); } };
    document.addEventListener('keydown', escHandler);

    title.textContent = 'Sites Deep Dive';
    body.innerHTML = '<p class="text-center text-muted">Loading...</p>';

    try {
        const res = await API.getData('sites');
        renderSitesDeepDive(body, res.data || []);

        // Pre-fill search with site name after a short delay to ensure DOM is ready
        setTimeout(() => {
            const searchInput = document.getElementById('dd-search');
            if (searchInput) {
                searchInput.value = siteName;
                searchInput.dispatchEvent(new Event('input'));
                searchInput.focus();
            }
        }, 100);
    } catch (e) {
        body.innerHTML = `<p class="text-center" style="color:#DC3545">Error loading data: ${esc(e.message)}</p>`;
    }
};

// Open permissions deep dive filtered to a specific permission level
window.openFilteredPermissionsDeepDive = async function(permissionLevel) {
    if (!appState.dataLoaded) {
        toast('Run an analysis first', 'info');
        return;
    }

    const overlay = document.getElementById('modal-overlay');
    const title = document.getElementById('modal-title');
    const body = document.getElementById('modal-body');

    overlay.classList.remove('hidden');
    document.getElementById('modal-close').onclick = () => overlay.classList.add('hidden');
    overlay.onclick = (e) => { if (e.target === overlay) overlay.classList.add('hidden'); };

    const escHandler = (e) => { if (e.key === 'Escape') { overlay.classList.add('hidden'); document.removeEventListener('keydown', escHandler); } };
    document.addEventListener('keydown', escHandler);

    title.textContent = 'Role Assignment Mapping';
    body.innerHTML = '<p class="text-center text-muted">Loading...</p>';

    try {
        const res = await API.getData('roleassignments');
        renderPermissionsDeepDive(body, res.data || []);

        // Pre-select role filter after a short delay to ensure DOM is ready
        setTimeout(() => {
            const roleFilter = document.getElementById('dd-role-filter');
            if (roleFilter) {
                roleFilter.value = permissionLevel;
                roleFilter.dispatchEvent(new Event('change'));
            }
        }, 100);
    } catch (e) {
        body.innerHTML = `<p class="text-center" style="color:#DC3545">Error loading data: ${esc(e.message)}</p>`;
    }
};

// Generic handler for deep dive chart clicks (for future extensions)
window.onDeepDiveChartClick = function(canvasId, clickedData) {
    // This can be extended in the future to handle specific chart interactions
    // For now, we'll show a toast with the clicked data
    console.log('Deep dive chart clicked:', canvasId, clickedData);
};

// --- Deep Dive Renderers ---

function renderSitesDeepDive(container, data) {
    const total = data.length;
    const totalStorage = data.reduce((sum, s) => sum + (parseInt(s.Storage) || 0), 0);
    const avgStorage = total > 0 ? Math.round(totalStorage / total) : 0;

    container.innerHTML = `
        <div class="dd-stats">
            <div class="dd-stat"><span class="dd-stat-value">${total}</span><span class="dd-stat-label">Total Sites</span></div>
            <div class="dd-stat"><span class="dd-stat-value">${formatStorage(totalStorage)}</span><span class="dd-stat-label">Total Storage</span></div>
            <div class="dd-stat"><span class="dd-stat-value">${avgStorage} MB</span><span class="dd-stat-label">Avg Storage</span></div>
        </div>
        <div class="dd-filter-bar"><input type="text" placeholder="Search sites..." id="dd-search"><button class="btn btn-secondary" onclick="showExportModal('sites')">Export</button></div>
        <table><thead><tr><th>Title</th><th>URL</th><th>Owner</th><th>Storage (MB)</th><th>Template</th><th>Actions</th></tr></thead>
        <tbody id="dd-sites-body">${renderSitesRows(data)}</tbody></table>`;

    document.getElementById('dd-search').addEventListener('input', (e) => {
        const q = e.target.value.toLowerCase();
        const filtered = data.filter(s => (s.Title || '').toLowerCase().includes(q) || (s.Url || '').toLowerCase().includes(q));
        document.getElementById('dd-sites-body').innerHTML = renderSitesRows(filtered);
    });
}

function renderSitesRows(data) {
    return data.map(s => `<tr>
        <td>${esc(s.Title)}</td>
        <td>${esc(s.Url)}</td>
        <td>${esc(s.Owner)}</td>
        <td>${s.Storage || 0}</td>
        <td>${esc(s.Template || 'N/A')}</td>
        <td><button class="btn btn-sm btn-primary" onclick="openPermissionsMatrix('${esc(s.Url)}', '${esc(s.Title)}')" title="Build Permissions Matrix">🔍 Matrix</button></td>
    </tr>`).join('');
}

function renderUsersDeepDive(container, data) {
    const internal = data.filter(u => u.Type !== 'External' && !u.IsExternal);
    const external = data.filter(u => u.Type === 'External' || u.IsExternal);

    container.innerHTML = `
        <div class="dd-stats">
            <div class="dd-stat"><span class="dd-stat-value">${data.length}</span><span class="dd-stat-label">Total Users</span></div>
            <div class="dd-stat"><span class="dd-stat-value" style="color:#2E7D32">${internal.length}</span><span class="dd-stat-label">Internal</span></div>
            <div class="dd-stat"><span class="dd-stat-value" style="color:#C62828">${external.length}</span><span class="dd-stat-label">External</span></div>
        </div>
        <div class="dd-filter-bar"><input type="text" placeholder="Search users..." id="dd-search">
        <select id="dd-type-filter"><option value="">All Types</option><option value="Internal">Internal</option><option value="External">External</option></select>
        <button class="btn btn-secondary" onclick="showExportModal('users')">Export</button></div>
        <table><thead><tr><th>Name</th><th>Email</th><th>Type</th><th>Permission</th><th>Site Admin</th></tr></thead>
        <tbody id="dd-users-body">${renderUsersRows(data)}</tbody></table>`;

    const filterUsers = () => {
        const q = document.getElementById('dd-search').value.toLowerCase();
        const t = document.getElementById('dd-type-filter').value;
        const filtered = data.filter(u => {
            const matchSearch = !q || (u.Name || '').toLowerCase().includes(q) || (u.Email || '').toLowerCase().includes(q);
            const matchType = !t || (t === 'External' ? (u.Type === 'External' || u.IsExternal) : (u.Type !== 'External' && !u.IsExternal));
            return matchSearch && matchType;
        });
        document.getElementById('dd-users-body').innerHTML = renderUsersRows(filtered);
    };
    document.getElementById('dd-search').addEventListener('input', filterUsers);
    document.getElementById('dd-type-filter').addEventListener('change', filterUsers);
}

function renderUsersRows(data) {
    return data.map(u => `<tr><td>${esc(u.Name)}</td><td>${esc(u.Email)}</td><td>${esc(u.Type || (u.IsExternal ? 'External' : 'Internal'))}</td><td>${esc(u.Permission)}</td><td>${u.IsSiteAdmin ? 'Yes' : ''}</td></tr>`).join('');
}

function renderGroupsDeepDive(container, data) {
    const totalMembers = data.reduce((sum, g) => sum + (parseInt(g.MemberCount) || 0), 0);
    const empty = data.filter(g => (parseInt(g.MemberCount) || 0) === 0).length;

    container.innerHTML = `
        <div class="dd-stats">
            <div class="dd-stat"><span class="dd-stat-value">${data.length}</span><span class="dd-stat-label">Total Groups</span></div>
            <div class="dd-stat"><span class="dd-stat-value">${totalMembers}</span><span class="dd-stat-label">Total Members</span></div>
            <div class="dd-stat"><span class="dd-stat-value" style="color:#DC3545">${empty}</span><span class="dd-stat-label">Empty Groups</span></div>
        </div>
        <div class="dd-filter-bar"><input type="text" placeholder="Search groups..." id="dd-search"><button class="btn btn-secondary" onclick="showExportModal('groups')">Export</button></div>
        <table><thead><tr><th>Name</th><th>Members</th><th>Permission</th><th>Description</th></tr></thead>
        <tbody id="dd-groups-body">${renderGroupsRows(data)}</tbody></table>`;

    document.getElementById('dd-search').addEventListener('input', (e) => {
        const q = e.target.value.toLowerCase();
        const filtered = data.filter(g => (g.Name || '').toLowerCase().includes(q));
        document.getElementById('dd-groups-body').innerHTML = renderGroupsRows(filtered);
    });
}

function renderGroupsRows(data) {
    return data.map(g => `<tr><td>${esc(g.Name)}</td><td>${g.MemberCount || 0}</td><td>${esc(g.Permission)}</td><td>${esc(g.Description || '')}</td></tr>`).join('');
}

function renderExternalDeepDive(container, allUsers) {
    const data = allUsers.filter(u => u.Type === 'External' || u.IsExternal);
    // Domain analysis
    const domains = {};
    data.forEach(u => {
        const email = u.Email || '';
        const domain = email.includes('@') ? email.split('@')[1] : 'Unknown';
        domains[domain] = (domains[domain] || 0) + 1;
    });

    const editAccess = data.filter(u => ['Edit', 'Contribute', 'Full Control'].includes(u.Permission)).length;
    const enrichedCount = data.filter(u => u.GraphEnriched).length;

    container.innerHTML = `
        <div class="dd-stats">
            <div class="dd-stat"><span class="dd-stat-value" style="color:#C62828">${data.length}</span><span class="dd-stat-label">External Users</span></div>
            <div class="dd-stat"><span class="dd-stat-value">${Object.keys(domains).length}</span><span class="dd-stat-label">Domains</span></div>
            <div class="dd-stat"><span class="dd-stat-value" style="color:#DC3545">${editAccess}</span><span class="dd-stat-label">With Edit+</span></div>
            <div class="dd-stat"><span class="dd-stat-value" style="color:#0078D4">${enrichedCount}</span><span class="dd-stat-label">Enriched</span></div>
        </div>
        ${editAccess > 0 ? '<div class="finding high"><h4>External Users with Edit Access</h4><p>' + editAccess + ' external user(s) have edit or higher permissions. Review and restrict where possible.</p></div>' : ''}
        <div id="enrichment-banner" style="margin-bottom:12px"></div>
        <div class="dd-filter-bar">
            <input type="text" placeholder="Search external users..." id="dd-search">
            <button class="btn btn-primary" id="btn-enrich" style="margin-left:8px">Enrich via Graph</button>
            <button class="btn btn-secondary" onclick="showExportModal('users')">Export</button>
        </div>
        <table><thead><tr><th>Name</th><th>Email</th><th>Domain</th><th>Permission</th><th>Account Status</th><th>Last Sign-In</th></tr></thead>
        <tbody id="dd-ext-body">${renderExternalRows(data)}</tbody></table>`;

    document.getElementById('dd-search').addEventListener('input', (e) => {
        const q = e.target.value.toLowerCase();
        const filtered = data.filter(u => (u.Name || '').toLowerCase().includes(q) || (u.Email || '').toLowerCase().includes(q));
        document.getElementById('dd-ext-body').innerHTML = renderExternalRows(filtered);
    });

    document.getElementById('btn-enrich').addEventListener('click', async () => {
        const btn = document.getElementById('btn-enrich');
        btn.textContent = 'Enriching...';
        btn.classList.add('btn-disabled');
        try {
            const res = await API.enrichExternal();
            if (res.started) {
                // Background operation — poll until complete
                const dummyConsole = document.createElement('div');
                const progress = await pollUntilComplete(dummyConsole);
                if (progress.error) {
                    toast('Enrichment failed: ' + progress.error, 'error');
                    return;
                }
                const result = progress.enrichmentResult || {};
                toast(`Enriched ${result.Enriched || 0} of ${result.TotalExternal || 0} external users`, 'success');
            } else if (res.success) {
                toast(`Enriched ${res.enriched} of ${res.totalExternal} external users`, 'success');
            } else {
                toast(res.message || 'Enrichment failed', 'error');
                return;
            }
            // Reload external users data
            const usersRes = await API.getData('users');
            const newData = (usersRes.data || []).filter(u => u.Type === 'External' || u.IsExternal);

            // Recalculate statistics with updated data
            const domains = {};
            newData.forEach(u => {
                const email = u.Email || '';
                const domain = email.includes('@') ? email.split('@')[1] : 'Unknown';
                domains[domain] = (domains[domain] || 0) + 1;
            });
            const editAccess = newData.filter(u => ['Edit', 'Contribute', 'Full Control'].includes(u.Permission)).length;
            const enrichedCount = newData.filter(u => u.GraphEnriched).length;

            // Update the stats in the dd-stats section
            const statsSection = document.querySelector('.dd-stats');
            if (statsSection) {
                statsSection.innerHTML = `
                    <div class="dd-stat"><span class="dd-stat-value" style="color:#C62828">${newData.length}</span><span class="dd-stat-label">External Users</span></div>
                    <div class="dd-stat"><span class="dd-stat-value">${Object.keys(domains).length}</span><span class="dd-stat-label">Domains</span></div>
                    <div class="dd-stat"><span class="dd-stat-value" style="color:#DC3545">${editAccess}</span><span class="dd-stat-label">With Edit+</span></div>
                    <div class="dd-stat"><span class="dd-stat-value" style="color:#0078D4">${enrichedCount}</span><span class="dd-stat-label">Enriched</span></div>
                `;
            }

            // Update the table body
            document.getElementById('dd-ext-body').innerHTML = renderExternalRows(newData);

            // Show enrichment summary
            showEnrichmentBanner();
        } catch (e) {
            toast('Enrichment failed: ' + e.message, 'error');
        } finally {
            btn.textContent = 'Enrich via Graph';
            btn.classList.remove('btn-disabled');
        }
    });

    // Show enrichment banner if already enriched
    if (enrichedCount > 0) {
        showEnrichmentBanner();
    }
}

async function showEnrichmentBanner() {
    try {
        const summary = await API.getEnrichment();
        const banner = document.getElementById('enrichment-banner');
        if (!banner) return;

        const findings = [];
        if (summary.disabledAccounts > 0) {
            findings.push(`<div class="finding high"><h4>${summary.disabledAccounts} disabled account(s) with access</h4><p>These accounts are disabled in Azure AD but still have SharePoint permissions. Remove their access.</p></div>`);
        }
        if (summary.staleAccounts > 0) {
            findings.push(`<div class="finding medium"><h4>${summary.staleAccounts} stale external account(s)</h4><p>No sign-in activity in 90+ days. Consider removing access for inactive external users.</p></div>`);
        }
        if (summary.enrichedCount > 0 && findings.length === 0) {
            findings.push(`<div class="finding low"><h4>External access looks healthy</h4><p>${summary.enrichedCount} external user(s) enriched. No disabled or stale accounts detected.</p></div>`);
        }
        banner.innerHTML = findings.join('');
    } catch (e) {
        // Non-critical, ignore
    }
}

function renderExternalRows(data) {
    return data.map(u => {
        const email = u.Email || '';
        const domain = email.includes('@') ? email.split('@')[1] : 'Unknown';

        // Calculate if account is stale (90+ days since last sign-in)
        let isStale = false;
        if (u.GraphLastSignIn) {
            const daysSinceSignIn = (Date.now() - new Date(u.GraphLastSignIn)) / (1000 * 60 * 60 * 24);
            isStale = daysSinceSignIn > 90;
        } else if (u.GraphCreatedDate && u.GraphEnriched) {
            // Never signed in, check if created more than 90 days ago
            const daysSinceCreated = (Date.now() - new Date(u.GraphCreatedDate)) / (1000 * 60 * 60 * 24);
            isStale = daysSinceCreated > 90;
        }

        // Build account status display
        let accountStatus = '<span style="color:#999">-</span>';
        if (u.GraphEnriched) {
            if (!u.GraphAccountEnabled) {
                accountStatus = '<span style="color:#DC3545">Disabled</span>';
            } else if (isStale) {
                accountStatus = '<span style="color:#28A745">Active</span> <span style="color:#F59E0B;font-size:11px">⚠ Stale</span>';
            } else {
                accountStatus = '<span style="color:#28A745">Active</span>';
            }
        }

        const lastSignIn = u.GraphLastSignIn
            ? new Date(u.GraphLastSignIn).toLocaleDateString()
            : (u.GraphEnriched ? '<span style="color:#DC3545">Never</span>' : '-');

        return `<tr><td>${esc(u.Name)}</td><td>${esc(email)}</td><td>${esc(domain)}</td><td>${esc(u.Permission)}</td><td>${accountStatus}</td><td>${lastSignIn}</td></tr>`;
    }).join('');
}

function renderPermissionsDeepDive(container, data) {
    const fullControl = data.filter(r => r.Role === 'Full Control').length;
    const edit = data.filter(r => r.Role === 'Edit' || r.Role === 'Contribute').length;
    const read = data.filter(r => r.Role === 'Read' || r.Role === 'View Only').length;

    // Security findings
    const findings = [];
    if (fullControl > 5) findings.push({ severity: 'high', title: `${fullControl} Full Control assignments`, detail: 'Review and reduce Full Control to minimum necessary.' });
    const extEditors = data.filter(r => r.PrincipalType === 'User' && ['Full Control', 'Edit', 'Contribute'].includes(r.Role));
    if (extEditors.length > 0) findings.push({ severity: 'medium', title: `${extEditors.length} users with edit+ access`, detail: 'Verify each user needs this level of access.' });

    container.innerHTML = `
        <div class="dd-stats">
            <div class="dd-stat"><span class="dd-stat-value">${data.length}</span><span class="dd-stat-label">Total Assignments</span></div>
            <div class="dd-stat"><span class="dd-stat-value" style="color:#DC3545">${fullControl}</span><span class="dd-stat-label">Full Control</span></div>
            <div class="dd-stat"><span class="dd-stat-value" style="color:#FFC107">${edit}</span><span class="dd-stat-label">Edit/Contribute</span></div>
            <div class="dd-stat"><span class="dd-stat-value" style="color:#28A745">${read}</span><span class="dd-stat-label">Read/View</span></div>
        </div>
        <div class="dd-tabs">
            <button class="dd-tab-btn active" data-ddtab="dd-table">All Assignments</button>
            <button class="dd-tab-btn" data-ddtab="dd-chart">Distribution</button>
            <button class="dd-tab-btn" data-ddtab="dd-findings">Security Review</button>
        </div>
        <div id="dd-table" class="dd-tab-content active">
            <div class="dd-filter-bar"><input type="text" placeholder="Search..." id="dd-search">
            <select id="dd-role-filter"><option value="">All Roles</option><option>Full Control</option><option>Edit</option><option>Contribute</option><option>Read</option><option>View Only</option></select>
            <button class="btn btn-secondary" onclick="showExportModal('roleassignments')">Export</button></div>
            <table><thead><tr><th>Principal</th><th>Type</th><th>Role</th><th>Scope</th><th>Location</th></tr></thead>
            <tbody id="dd-perm-body">${renderPermRows(data)}</tbody></table>
        </div>
        <div id="dd-chart" class="dd-tab-content"><div style="height:250px"><canvas id="dd-perm-chart"></canvas></div></div>
        <div id="dd-findings" class="dd-tab-content">${renderFindings(findings)}</div>`;

    initDDTabs();

    // Render chart when tab is shown
    const roleCounts = {};
    data.forEach(r => { roleCounts[r.Role] = (roleCounts[r.Role] || 0) + 1; });
    const chartData = Object.entries(roleCounts).map(([label, value]) => ({
        label, value,
        color: { 'Full Control': COLORS.fullControl, 'Edit': COLORS.edit, 'Contribute': COLORS.contribute, 'Read': COLORS.read, 'View Only': COLORS.viewOnly }[label] || COLORS.custom
    }));
    setTimeout(() => renderDeepDiveChart('dd-perm-chart', 'doughnut', chartData), 100);

    // Filters
    const filterPerms = () => {
        const q = (document.getElementById('dd-search').value || '').toLowerCase();
        const r = document.getElementById('dd-role-filter').value;
        const filtered = data.filter(item => {
            const matchSearch = !q || (item.Principal || '').toLowerCase().includes(q) || (item.ScopeUrl || '').toLowerCase().includes(q);
            const matchRole = !r || item.Role === r;
            return matchSearch && matchRole;
        });
        document.getElementById('dd-perm-body').innerHTML = renderPermRows(filtered);
    };
    document.getElementById('dd-search').addEventListener('input', filterPerms);
    document.getElementById('dd-role-filter').addEventListener('change', filterPerms);
}

function renderPermRows(data) {
    return data.map(r => `<tr><td>${esc(r.Principal)}</td><td>${esc(r.PrincipalType)}</td><td>${esc(r.Role)}</td><td>${esc(r.Scope)}</td><td>${esc(r.ScopeUrl)}</td></tr>`).join('');
}

// --- Tree Visualization Helper Functions ---

// Transform flat inheritance data into hierarchical tree structure
function buildInheritanceTree(data) {
    const siteGroups = {};

    // First pass: identify all sites (roots of the tree)
    data.forEach(item => {
        if (item.Type === 'Site') {
            if (!siteGroups[item.Url]) {
                siteGroups[item.Url] = {
                    site: item,
                    children: []
                };
            }
        }
    });

    // Second pass: assign children to their parent sites
    data.forEach(item => {
        if (item.Type !== 'Site' && item.ParentUrl) {
            const parent = siteGroups[item.ParentUrl];
            if (parent) {
                parent.children.push(item);
            } else {
                // Orphaned item (parent site not in data) - create a placeholder
                if (!siteGroups[item.ParentUrl]) {
                    siteGroups[item.ParentUrl] = {
                        site: { Title: 'Unknown Site', Url: item.ParentUrl, Type: 'Site', HasUniquePermissions: false },
                        children: [item]
                    };
                }
            }
        }
    });

    return Object.values(siteGroups);
}

// Render tree view HTML
function renderTreeView(treeData) {
    let html = '<div class="tree-view">';

    treeData.forEach((siteGroup, index) => {
        const site = siteGroup.site;
        const siteId = `tree-site-${index}`;
        const isBroken = site.HasUniquePermissions === true || site.HasUniquePermissions === 'True';

        html += `
            <div class="tree-node tree-node-site ${isBroken ? 'tree-node-broken' : 'tree-node-inheriting'}">
                <div class="tree-node-header" onclick="toggleTreeNode('${siteId}')">
                    <span class="tree-expand-icon ${siteGroup.children.length === 0 ? 'tree-no-children' : ''}" id="${siteId}-icon">
                        ${siteGroup.children.length > 0 ? '▼' : ''}
                    </span>
                    <span class="tree-node-icon">🌐</span>
                    <span class="tree-node-title">${esc(site.Title)}</span>
                    <span class="tree-node-badge ${isBroken ? 'badge-broken' : 'badge-inheriting'}">
                        ${isBroken ? 'Unique Permissions' : 'Inherited'}
                    </span>
                    ${site.RoleAssignmentCount ? `<span class="tree-node-count">${site.RoleAssignmentCount} assignments</span>` : ''}
                </div>
                <div class="tree-node-children" id="${siteId}-children">`;

        // Render children (libraries/lists)
        siteGroup.children.forEach((child, childIndex) => {
            const childBroken = child.HasUniquePermissions === true || child.HasUniquePermissions === 'True';
            const icon = child.Type === 'Document Library' || child.Type === 'Library' ? '📁' : '📄';

            html += `
                <div class="tree-node tree-node-child ${childBroken ? 'tree-node-broken' : 'tree-node-inheriting'}">
                    <div class="tree-node-header">
                        <span class="tree-node-icon">${icon}</span>
                        <span class="tree-node-title">${esc(child.Title)}</span>
                        <span class="tree-node-type">${esc(child.Type)}</span>
                        <span class="tree-node-badge ${childBroken ? 'badge-broken' : 'badge-inheriting'}">
                            ${childBroken ? 'Unique Permissions' : 'Inherited'}
                        </span>
                        ${child.RoleAssignmentCount ? `<span class="tree-node-count">${child.RoleAssignmentCount} assignments</span>` : ''}
                    </div>
                </div>`;
        });

        html += `
                </div>
            </div>`;
    });

    html += '</div>';
    return html;
}

// Toggle tree node expansion
window.toggleTreeNode = function(nodeId) {
    const children = document.getElementById(`${nodeId}-children`);
    const icon = document.getElementById(`${nodeId}-icon`);

    if (children && icon) {
        if (children.style.display === 'none') {
            children.style.display = 'block';
            icon.textContent = '▼';
        } else {
            children.style.display = 'none';
            icon.textContent = '▶';
        }
    }
};

function renderInheritanceDeepDive(container, data) {
    const broken = data.filter(i => i.HasUniquePermissions === true || i.HasUniquePermissions === 'True').length;
    const inheriting = data.length - broken;
    const libraries = data.filter(i => i.Type === 'Document Library' || i.Type === 'Library').length;
    const lists = data.filter(i => i.Type === 'List').length;

    const findings = [];
    const breakPct = data.length > 0 ? Math.round((broken / data.length) * 100) : 0;
    if (breakPct > 50) findings.push({ severity: 'high', title: `${breakPct}% of items have broken inheritance`, detail: 'Excessive permission breaks. Consider consolidating at site level.' });
    else if (breakPct > 25) findings.push({ severity: 'medium', title: `${breakPct}% broken inheritance`, detail: 'Review broken items and consolidate where possible.' });

    const brokenLibs = data.filter(i => (i.HasUniquePermissions === true || i.HasUniquePermissions === 'True') && (i.Type === 'Document Library' || i.Type === 'Library'));
    if (brokenLibs.length > 0) findings.push({ severity: 'medium', title: `${brokenLibs.length} libraries with unique permissions`, detail: 'Verify access on these document libraries is intentional.' });

    // Build tree structure
    const treeData = buildInheritanceTree(data);

    container.innerHTML = `
        <div class="dd-stats">
            <div class="dd-stat"><span class="dd-stat-value">${data.length}</span><span class="dd-stat-label">Total Items</span></div>
            <div class="dd-stat"><span class="dd-stat-value" style="color:#28A745">${inheriting}</span><span class="dd-stat-label">Inheriting</span></div>
            <div class="dd-stat"><span class="dd-stat-value" style="color:#DC3545">${broken}</span><span class="dd-stat-label">Broken</span></div>
            <div class="dd-stat"><span class="dd-stat-value">${libraries}</span><span class="dd-stat-label">Libraries</span></div>
            <div class="dd-stat"><span class="dd-stat-value">${lists}</span><span class="dd-stat-label">Lists</span></div>
        </div>
        <div class="dd-tabs">
            <button class="dd-tab-btn active" data-ddtab="dd-tree">Tree View</button>
            <button class="dd-tab-btn" data-ddtab="dd-table">Table View</button>
            <button class="dd-tab-btn" data-ddtab="dd-chart">Overview</button>
            <button class="dd-tab-btn" data-ddtab="dd-findings">Findings</button>
        </div>
        <div id="dd-tree" class="dd-tab-content active">
            <div class="dd-filter-bar">
                <select id="dd-tree-filter"><option value="">All Items</option><option value="broken">Broken Inheritance Only</option><option value="inheriting">Inheriting Only</option></select>
                <button class="btn btn-secondary" onclick="showExportModal('inheritance')">Export</button>
            </div>
            <div id="dd-tree-container">${renderTreeView(treeData)}</div>
        </div>
        <div id="dd-table" class="dd-tab-content">
            <div class="dd-filter-bar"><input type="text" placeholder="Search..." id="dd-search">
            <select id="dd-inh-filter"><option value="">All Items</option><option value="broken">Broken Inheritance</option><option value="inheriting">Inheriting</option></select>
            <button class="btn btn-secondary" onclick="showExportModal('inheritance')">Export</button></div>
            <table><thead><tr><th>Title</th><th>Type</th><th>Unique Perms</th><th>Role Assignments</th><th>Site</th></tr></thead>
            <tbody id="dd-inh-body">${renderInhRows(data)}</tbody></table>
        </div>
        <div id="dd-chart" class="dd-tab-content"><div style="height:250px"><canvas id="dd-inh-chart"></canvas></div></div>
        <div id="dd-findings" class="dd-tab-content">${renderFindings(findings)}</div>`;

    initDDTabs();
    setTimeout(() => renderDeepDiveChart('dd-inh-chart', 'doughnut', [
        { label: 'Inheriting', value: inheriting, color: COLORS.green },
        { label: 'Broken', value: broken, color: COLORS.red }
    ]), 100);

    const filterInh = () => {
        const q = (document.getElementById('dd-search').value || '').toLowerCase();
        const f = document.getElementById('dd-inh-filter').value;
        const filtered = data.filter(i => {
            const matchSearch = !q || (i.Title || '').toLowerCase().includes(q);
            const isBroken = i.HasUniquePermissions === true || i.HasUniquePermissions === 'True';
            const matchFilter = !f || (f === 'broken' && isBroken) || (f === 'inheriting' && !isBroken);
            return matchSearch && matchFilter;
        });
        document.getElementById('dd-inh-body').innerHTML = renderInhRows(filtered);
    };
    document.getElementById('dd-search').addEventListener('input', filterInh);
    document.getElementById('dd-inh-filter').addEventListener('change', filterInh);

    // Tree view filter
    const filterTree = () => {
        const f = document.getElementById('dd-tree-filter').value;
        const filtered = data.filter(i => {
            const isBroken = i.HasUniquePermissions === true || i.HasUniquePermissions === 'True';
            return !f || (f === 'broken' && isBroken) || (f === 'inheriting' && !isBroken);
        });
        const filteredTree = buildInheritanceTree(filtered);
        document.getElementById('dd-tree-container').innerHTML = renderTreeView(filteredTree);
    };
    document.getElementById('dd-tree-filter').addEventListener('change', filterTree);
}

function renderInhRows(data) {
    return data.map(i => {
        const broken = i.HasUniquePermissions === true || i.HasUniquePermissions === 'True';
        return `<tr><td>${esc(i.Title)}</td><td>${esc(i.Type)}</td><td style="color:${broken ? '#DC3545' : '#28A745'}">${broken ? 'Yes' : 'No'}</td><td>${i.RoleAssignmentCount || 0}</td><td>${esc(i.SiteTitle)}</td></tr>`;
    }).join('');
}

function renderSharingDeepDive(container, data) {
    const anonymous = data.filter(l => l.LinkType === 'Anonymous').length;
    const org = data.filter(l => l.LinkType === 'Company-wide' || l.LinkType === 'Organization').length;
    const specific = data.filter(l => l.LinkType === 'Specific People').length;
    const totalRecipients = data.reduce((sum, l) => sum + (parseInt(l.MemberCount) || 0), 0);

    const findings = [];
    const anonEdit = data.filter(l => l.LinkType === 'Anonymous' && l.AccessLevel === 'Edit');
    if (anonEdit.length > 0) findings.push({ severity: 'high', title: `${anonEdit.length} anonymous edit link(s)`, detail: 'CRITICAL: Anyone with these links can modify content. Remove immediately.' });
    if (anonymous > 0) findings.push({ severity: 'high', title: `${anonymous} anonymous link(s)`, detail: 'Anonymous links allow access without authentication. Review all.' });
    if (org > 5) findings.push({ severity: 'medium', title: `${org} company-wide links`, detail: 'Exposes content to all employees. Use specific-people links for sensitive content.' });

    container.innerHTML = `
        <div class="dd-stats">
            <div class="dd-stat"><span class="dd-stat-value">${data.length}</span><span class="dd-stat-label">Total Links</span></div>
            <div class="dd-stat"><span class="dd-stat-value" style="color:#DC3545">${anonymous}</span><span class="dd-stat-label">Anonymous</span></div>
            <div class="dd-stat"><span class="dd-stat-value" style="color:#FFC107">${org}</span><span class="dd-stat-label">Company-wide</span></div>
            <div class="dd-stat"><span class="dd-stat-value" style="color:#28A745">${specific}</span><span class="dd-stat-label">Specific People</span></div>
            <div class="dd-stat"><span class="dd-stat-value">${totalRecipients}</span><span class="dd-stat-label">Recipients</span></div>
        </div>
        <div class="dd-tabs">
            <button class="dd-tab-btn active" data-ddtab="dd-table">All Links</button>
            <button class="dd-tab-btn" data-ddtab="dd-chart">Distribution</button>
            <button class="dd-tab-btn" data-ddtab="dd-findings">Security Findings</button>
        </div>
        <div id="dd-table" class="dd-tab-content active">
            <div class="dd-filter-bar"><input type="text" placeholder="Search..." id="dd-search">
            <select id="dd-link-filter"><option value="">All Types</option><option>Anonymous</option><option>Company-wide</option><option>Specific People</option></select>
            <button class="btn btn-secondary" onclick="showExportModal('sharinglinks')">Export</button></div>
            <table><thead><tr><th>Link Type</th><th>Access</th><th>Recipients</th><th>Site</th><th>Created</th></tr></thead>
            <tbody id="dd-share-body">${renderShareRows(data)}</tbody></table>
        </div>
        <div id="dd-chart" class="dd-tab-content"><div style="height:250px"><canvas id="dd-share-chart"></canvas></div></div>
        <div id="dd-findings" class="dd-tab-content">${renderFindings(findings)}</div>`;

    initDDTabs();
    setTimeout(() => renderDeepDiveChart('dd-share-chart', 'doughnut', [
        { label: 'Anonymous', value: anonymous, color: COLORS.red },
        { label: 'Company-wide', value: org, color: COLORS.amber },
        { label: 'Specific People', value: specific, color: COLORS.green }
    ]), 100);

    const filterShare = () => {
        const q = (document.getElementById('dd-search').value || '').toLowerCase();
        const f = document.getElementById('dd-link-filter').value;
        const filtered = data.filter(l => {
            const matchSearch = !q || (l.GroupName || '').toLowerCase().includes(q) || (l.SiteTitle || '').toLowerCase().includes(q);
            const matchType = !f || l.LinkType === f;
            return matchSearch && matchType;
        });
        document.getElementById('dd-share-body').innerHTML = renderShareRows(filtered);
    };
    document.getElementById('dd-search').addEventListener('input', filterShare);
    document.getElementById('dd-link-filter').addEventListener('change', filterShare);
}

function renderShareRows(data) {
    return data.map(l => {
        const typeColor = l.LinkType === 'Anonymous' ? '#DC3545' : l.LinkType === 'Company-wide' ? '#F57F17' : '#28A745';
        return `<tr><td style="color:${typeColor};font-weight:600">${esc(l.LinkType)}</td><td>${esc(l.AccessLevel)}</td><td>${l.MemberCount || 0}</td><td>${esc(l.SiteTitle)}</td><td>${esc(l.CreatedDate || 'N/A')}</td></tr>`;
    }).join('');
}

// --- Deep Dive Helpers ---
function initDDTabs() {
    document.querySelectorAll('.dd-tab-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const parent = btn.closest('.modal-body') || document.getElementById('modal-body');
            parent.querySelectorAll('.dd-tab-btn').forEach(b => b.classList.remove('active'));
            parent.querySelectorAll('.dd-tab-content').forEach(c => c.classList.remove('active'));
            btn.classList.add('active');
            const target = document.getElementById(btn.dataset.ddtab);
            if (target) target.classList.add('active');
        });
    });
}

function renderFindings(findings) {
    if (findings.length === 0) {
        return '<div class="finding low"><h4>No Issues Found</h4><p>No significant security findings detected.</p></div>';
    }
    return findings.map(f => `<div class="finding ${f.severity}"><h4>${esc(f.title)}</h4><p>${esc(f.detail)}</p></div>`).join('');
}

// --- Permissions Matrix Functions ---

// Open permissions matrix modal
window.openPermissionsMatrix = function(siteUrl, siteTitle) {
    const modal = document.getElementById('matrix-modal');
    const title = document.getElementById('matrix-title');
    const body = document.getElementById('matrix-body');

    title.textContent = `Permissions Matrix: ${siteTitle}`;
    currentMatrixSiteUrl = siteUrl;

    // Show scan type chooser
    body.innerHTML = `
        <div class="matrix-scan-chooser">
            <h3>Select Scan Type</h3>
            <p style="color: #64748B; margin-bottom: 24px;">
                Choose how to build the permissions matrix for this site:
            </p>
            <div class="scan-type-options">
                <button class="scan-type-btn" onclick="buildMatrix('${esc(siteUrl)}', 'quick')">
                    <span class="scan-icon">⚡</span>
                    <span class="scan-name">Quick Scan</span>
                    <span class="scan-desc">Only items with unique permissions (faster)</span>
                    <span class="scan-time">~30 seconds</span>
                </button>
                <button class="scan-type-btn" onclick="buildMatrix('${esc(siteUrl)}', 'full')">
                    <span class="scan-icon">🔍</span>
                    <span class="scan-name">Full Scan</span>
                    <span class="scan-desc">All files and folders (comprehensive)</span>
                    <span class="scan-time">~5-10 minutes for large sites</span>
                </button>
            </div>
        </div>
    `;

    modal.classList.remove('hidden');

    // Close handlers
    document.getElementById('matrix-close').onclick = () => modal.classList.add('hidden');
    modal.onclick = (e) => { if (e.target === modal) modal.classList.add('hidden'); };
};

// Build permissions matrix
window.buildMatrix = async function(siteUrl, scanType) {
    const body = document.getElementById('matrix-body');

    body.innerHTML = `
        <div class="matrix-loading">
            <div class="loading-spinner"></div>
            <p>Building permissions matrix...</p>
            <p style="font-size: 0.9rem; color: #64748B;">
                ${scanType === 'quick' ? 'Scanning items with unique permissions' : 'Scanning all files and folders'}
            </p>
        </div>
    `;

    try {
        const response = await API.buildPermissionsMatrix(siteUrl, scanType);
        currentMatrixData = response.data;
        renderPermissionsMatrix(response.data);
        toast(`Matrix built: ${response.data.totalItems} items scanned`, 'success');
    } catch (e) {
        body.innerHTML = `<p class="text-center" style="color:#DC3545">
            Failed to build matrix: ${esc(e.message)}
        </p>`;
        toast('Matrix build failed', 'error');
    }
};

// Render permissions matrix tree
function renderPermissionsMatrix(data) {
    const body = document.getElementById('matrix-body');

    body.innerHTML = `
        <div class="matrix-stats">
            <div class="dd-stat">
                <span class="dd-stat-value">${data.totalItems}</span>
                <span class="dd-stat-label">Total Items</span>
            </div>
            <div class="dd-stat">
                <span class="dd-stat-value">${data.uniquePermissions}</span>
                <span class="dd-stat-label">Unique Permissions</span>
            </div>
            <div class="dd-stat">
                <span class="dd-stat-value">${data.totalPrincipals}</span>
                <span class="dd-stat-label">Users/Groups</span>
            </div>
        </div>
        <div class="matrix-toolbar">
            <button class="btn btn-secondary" onclick="showExportModal('permissions-matrix', false)">
                Export Matrix
            </button>
        </div>
        <div class="matrix-tree-container">
            ${renderMatrixTree(data.tree)}
        </div>
    `;
}

// Render matrix tree hierarchy
function renderMatrixTree(nodes, level = 0) {
    let html = '<div class="matrix-tree">';

    nodes.forEach((node, index) => {
        const nodeId = `matrix-node-${level}-${index}`;
        const hasChildren = node.children && node.children.length > 0;
        const icon = getNodeIcon(node.type);

        html += `
            <div class="matrix-node matrix-node-${node.type.toLowerCase()}" style="margin-left: ${level * 20}px">
                <div class="matrix-node-header" ${hasChildren ? `onclick="toggleMatrixNode('${nodeId}')"` : ''}>
                    ${hasChildren ? `<span class="tree-expand-icon" id="${nodeId}-icon">▼</span>` : '<span class="tree-expand-icon"></span>'}
                    <span class="tree-node-icon">${icon}</span>
                    <span class="matrix-node-title">${esc(node.title)}</span>
                    <span class="matrix-node-type">${node.type}</span>
                    ${renderPermissionBadges(node.permissions)}
                </div>
                ${hasChildren ? `<div class="matrix-node-children" id="${nodeId}-children">
                    ${renderMatrixTree(node.children, level + 1)}
                </div>` : ''}
            </div>
        `;
    });

    html += '</div>';
    return html;
}

// Render permission badges for a node
function renderPermissionBadges(permissions) {
    if (!permissions || permissions.length === 0) {
        return '<span class="permission-badge inherited">Inherited</span>';
    }

    return permissions.map(p => `
        <span class="permission-badge" title="${esc(p.principal)} - ${esc(p.role)}">
            ${esc(p.principal)}: ${esc(p.role)}
        </span>
    `).join('');
}

// Get icon for node type
function getNodeIcon(type) {
    const icons = {
        'Site': '🌐',
        'Library': '📚',
        'List': '📋',
        'Folder': '📁',
        'File': '📄'
    };
    return icons[type] || '📦';
}

// Toggle matrix tree node
window.toggleMatrixNode = function(nodeId) {
    const children = document.getElementById(`${nodeId}-children`);
    const icon = document.getElementById(`${nodeId}-icon`);

    if (children && icon) {
        if (children.style.display === 'none') {
            children.style.display = 'block';
            icon.textContent = '▼';
        } else {
            children.style.display = 'none';
            icon.textContent = '▶';
        }
    }
};

// Export matrix to CSV
function exportMatrixToCSV(matrixData) {
    const rows = [];
    rows.push(['Type', 'Title', 'URL', 'Principal', 'Role']);

    function traverseNode(node) {
        if (node.permissions && node.permissions.length > 0) {
            node.permissions.forEach(p => {
                rows.push([node.type, node.title, node.url, p.principal, p.role]);
            });
        } else {
            rows.push([node.type, node.title, node.url, '', 'Inherited']);
        }

        if (node.children) {
            node.children.forEach(child => traverseNode(child));
        }
    }

    matrixData.tree.forEach(node => traverseNode(node));

    const csv = rows.map(row => row.map(cell => `"${cell}"`).join(',')).join('\n');
    const blob = new Blob([csv], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `permissions_matrix_${currentMatrixSiteUrl.replace(/[^a-z0-9]/gi, '_')}_${Date.now()}.csv`;
    a.click();
    URL.revokeObjectURL(url);
    toast('Matrix exported as CSV', 'success');
}

// Export matrix to JSON
function exportMatrixToJSON(matrixData) {
    const blob = new Blob([JSON.stringify(matrixData, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `permissions_matrix_${currentMatrixSiteUrl.replace(/[^a-z0-9]/gi, '_')}_${Date.now()}.json`;
    a.click();
    URL.revokeObjectURL(url);
    toast('Matrix exported as JSON', 'success');
}

// --- Utilities ---
function setText(id, value) {
    const el = document.getElementById(id);
    if (el) el.textContent = value;
}

function esc(str) {
    if (!str) return '';
    const div = document.createElement('div');
    div.textContent = String(str);
    return div.innerHTML;
}

function formatStorage(mb) {
    if (mb >= 1024) return (mb / 1024).toFixed(1) + ' GB';
    return mb + ' MB';
}

function setButtonLoading(id, loading) {
    // Use UIHelpers if available for modern loading spinner
    if (typeof UIHelpers !== 'undefined') {
        UIHelpers.setButtonLoading(id, loading);
    } else {
        // Fallback to text-based loading
        const btn = document.getElementById(id);
        if (!btn) return;
        if (loading) {
            btn.dataset.originalText = btn.textContent;
            btn.textContent = 'Loading...';
            btn.classList.add('btn-disabled');
        } else {
            btn.textContent = btn.dataset.originalText || btn.textContent;
            btn.classList.remove('btn-disabled');
        }
    }
}

function toast(message, type = 'info') {
    const container = document.getElementById('toast-container');
    if (!container) return;
    const t = document.createElement('div');
    t.className = `toast ${type}`;
    t.textContent = message;
    container.appendChild(t);
    setTimeout(() => t.remove(), 4000);
}

// --- Audit Summary ---
async function showAuditSummary(console_) {
    try {
        const audit = await API.getAudit();
        if (!audit.hasSession) return;

        console_.textContent += '\n\n=== AUDIT TRAIL ===';
        console_.textContent += `\nSession ID: ${audit.sessionId}`;
        console_.textContent += `\nOperation: ${audit.operationType}`;
        console_.textContent += `\nStatus: ${audit.status}`;
        console_.textContent += `\nUser: ${audit.userPrincipal}`;
        if (audit.duration) console_.textContent += `\nDuration: ${audit.duration}`;
        console_.textContent += `\nEvents: ${audit.eventCount} | Errors: ${audit.errorCount}`;
        if (audit.outputFiles && audit.outputFiles.length > 0) {
            console_.textContent += `\nOutput files: ${audit.outputFiles.length}`;
        }
        console_.textContent += '\n(Full audit log saved to ./Logs/)';
    } catch (e) {
        // Audit info is supplementary, don't fail the operation
    }
}

// --- Operation polling helper ---
// Polls /api/progress until the operation is complete, updating the console element.
async function pollUntilComplete(consoleEl, intervalMs = 1000) {
    while (true) {
        await new Promise(r => setTimeout(r, intervalMs));
        try {
            const progress = await API.getProgress();
            if (progress.messages) {
                consoleEl.textContent = progress.messages.join('\n');
            }
            if (!progress.running && progress.complete) {
                return progress;
            }
        } catch (e) {
            // Transient fetch failure — keep polling
        }
    }
}

// --- Export Format Selection ---

let pendingExportType = null;
let pendingExportIsFullReport = false;

// Show export format selection modal
window.showExportModal = function(type, isFullReport = false) {
    const modal = document.getElementById('export-modal');
    const closeBtn = document.getElementById('export-modal-close');

    pendingExportType = type;
    pendingExportIsFullReport = isFullReport;

    modal.classList.remove('hidden');

    // Close handlers
    closeBtn.onclick = () => modal.classList.add('hidden');
    modal.onclick = (e) => { if (e.target === modal) modal.classList.add('hidden'); };

    // Escape key
    const escHandler = (e) => {
        if (e.key === 'Escape') {
            modal.classList.add('hidden');
            document.removeEventListener('keydown', escHandler);
        }
    };
    document.addEventListener('keydown', escHandler);
};

// Handle format selection
function initExportModal() {
    document.querySelectorAll('.export-format-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const format = btn.dataset.format;
            const modal = document.getElementById('export-modal');

            modal.classList.add('hidden');

            // Perform export based on format
            if (pendingExportIsFullReport) {
                handleReportExport(format);
            } else {
                handleDataExport(pendingExportType, format);
            }
        });
    });
}

// Export data in selected format
function handleDataExport(type, format) {
    if (type === 'permissions-matrix') {
        if (format === 'csv') {
            exportMatrixToCSV(currentMatrixData);
        } else if (format === 'json') {
            exportMatrixToJSON(currentMatrixData);
        }
    } else {
        if (format === 'csv') {
            API.exportData(type);
            toast(`Exporting ${type} as CSV`, 'success');
        } else if (format === 'json') {
            API.exportDataJson(type);
            toast(`Exporting ${type} as JSON`, 'success');
        }
    }
}

// Export full report in selected format
async function handleReportExport(format) {
    if (!appState.dataLoaded) {
        toast('Run an analysis first', 'info');
        return;
    }

    try {
        if (format === 'json') {
            const report = await API.exportJson();
            const blob = new Blob([JSON.stringify(report, null, 2)], { type: 'application/json' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `spo_governance_${new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19)}.json`;
            a.click();
            URL.revokeObjectURL(url);
            toast('Governance JSON report downloaded', 'success');
        } else if (format === 'csv') {
            // Export all data types as separate CSV files
            const types = ['sites', 'users', 'groups', 'roleassignments', 'inheritance', 'sharinglinks'];
            types.forEach(type => API.exportData(type));
            toast('Exporting all data as CSV files', 'success');
        }
    } catch (e) {
        toast('Export failed: ' + e.message, 'error');
    }
}

// --- Global Search Functionality ---

let globalSearchData = {
    sites: [],
    users: [],
    groups: [],
    permissions: [],
    inheritance: [],
    loaded: false
};

let selectedResultIndex = -1;
let currentSearchResults = [];

// Initialize global search
function initGlobalSearch() {
    const searchInput = document.getElementById('global-search-input');
    const searchResults = document.getElementById('global-search-results');

    if (!searchInput || !searchResults) return;

    // Keyboard shortcut (Ctrl+K or Cmd+K)
    document.addEventListener('keydown', (e) => {
        if ((e.ctrlKey || e.metaKey) && e.key === 'k') {
            e.preventDefault();
            searchInput.focus();
        }
    });

    // Search input handler with debounce
    let debounceTimer;
    searchInput.addEventListener('input', (e) => {
        clearTimeout(debounceTimer);
        const query = e.target.value.trim();

        if (query.length < 2) {
            searchResults.classList.add('hidden');
            return;
        }

        if (!appState.dataLoaded) {
            searchResults.innerHTML = '<div class="search-no-results">Please connect to SharePoint or start Demo Mode first</div>';
            searchResults.classList.remove('hidden');
            return;
        }

        debounceTimer = setTimeout(() => {
            performGlobalSearch(query);
        }, 300);
    });

    // Keyboard navigation in search results
    searchInput.addEventListener('keydown', (e) => {
        if (!searchResults.classList.contains('hidden')) {
            if (e.key === 'ArrowDown') {
                e.preventDefault();
                selectedResultIndex = Math.min(selectedResultIndex + 1, currentSearchResults.length - 1);
                highlightSelectedResult();
            } else if (e.key === 'ArrowUp') {
                e.preventDefault();
                selectedResultIndex = Math.max(selectedResultIndex - 1, -1);
                highlightSelectedResult();
            } else if (e.key === 'Enter' && selectedResultIndex >= 0) {
                e.preventDefault();
                const result = currentSearchResults[selectedResultIndex];
                if (result) {
                    navigateToSearchResult(result.type, result.item);
                }
            } else if (e.key === 'Escape') {
                searchResults.classList.add('hidden');
                searchInput.blur();
            }
        }
    });

    // Focus handler - lazy load data
    searchInput.addEventListener('focus', async () => {
        if (!globalSearchData.loaded && appState.dataLoaded) {
            await loadGlobalSearchData();
        }
    });

    // Click outside to close
    document.addEventListener('click', (e) => {
        if (!searchInput.contains(e.target) && !searchResults.contains(e.target)) {
            searchResults.classList.add('hidden');
        }
    });
}

// Load all data for searching (lazy loading)
async function loadGlobalSearchData() {
    try {
        const [sites, users, groups, permissions, inheritance] = await Promise.all([
            API.getData('sites'),
            API.getData('users'),
            API.getData('groups'),
            API.getData('roleassignments'),
            API.getData('inheritance')
        ]);

        globalSearchData = {
            sites: sites.data || [],
            users: users.data || [],
            groups: groups.data || [],
            permissions: permissions.data || [],
            inheritance: inheritance.data || [],
            loaded: true
        };
    } catch (e) {
        console.error('Failed to load search data:', e);
    }
}

// Perform search across all data types
function performGlobalSearch(query) {
    const q = query.toLowerCase();
    const results = {
        sites: [],
        users: [],
        groups: [],
        permissions: [],
        inheritance: []
    };

    // Search sites
    results.sites = globalSearchData.sites.filter(s =>
        (s.Title || '').toLowerCase().includes(q) ||
        (s.Url || '').toLowerCase().includes(q) ||
        (s.Owner || '').toLowerCase().includes(q)
    ).slice(0, 5);

    // Search users
    results.users = globalSearchData.users.filter(u =>
        (u.Name || '').toLowerCase().includes(q) ||
        (u.Email || '').toLowerCase().includes(q)
    ).slice(0, 5);

    // Search groups
    results.groups = globalSearchData.groups.filter(g =>
        (g.Name || '').toLowerCase().includes(q) ||
        (g.Description || '').toLowerCase().includes(q)
    ).slice(0, 5);

    // Search permissions
    results.permissions = globalSearchData.permissions.filter(p =>
        (p.Principal || '').toLowerCase().includes(q) ||
        (p.Role || '').toLowerCase().includes(q)
    ).slice(0, 5);

    // Search inheritance
    results.inheritance = globalSearchData.inheritance.filter(i =>
        (i.Title || '').toLowerCase().includes(q) ||
        (i.SiteTitle || '').toLowerCase().includes(q)
    ).slice(0, 5);

    renderSearchResults(results, query);
}

// Render search results dropdown
function renderSearchResults(results, query) {
    const resultsContainer = document.getElementById('global-search-results');
    const totalResults = results.sites.length + results.users.length + results.groups.length +
                        results.permissions.length + results.inheritance.length;

    if (totalResults === 0) {
        resultsContainer.innerHTML = '<div class="search-no-results">No results found</div>';
        resultsContainer.classList.remove('hidden');
        currentSearchResults = [];
        selectedResultIndex = -1;
        return;
    }

    currentSearchResults = [];
    let html = '';

    // Helper to add result group
    const addGroup = (title, items, type, icon) => {
        if (items.length > 0) {
            html += `<div class="search-result-group">
                <div class="search-result-group-title">${icon} ${title} (${items.length})</div>`;

            items.forEach((item, index) => {
                const resultIndex = currentSearchResults.length;
                currentSearchResults.push({ type, item });

                let primaryText = '';
                let secondaryText = '';

                if (type === 'sites') {
                    primaryText = esc(item.Title);
                    secondaryText = esc(item.Url);
                } else if (type === 'users') {
                    primaryText = esc(item.Name);
                    secondaryText = esc(item.Email);
                } else if (type === 'groups') {
                    primaryText = esc(item.Name);
                    secondaryText = `${item.MemberCount || 0} members`;
                } else if (type === 'permissions') {
                    primaryText = esc(item.Principal);
                    secondaryText = `${item.Role} on ${item.Scope}`;
                } else if (type === 'inheritance') {
                    primaryText = esc(item.Title);
                    secondaryText = esc(item.SiteTitle);
                }

                html += `<div class="search-result-item" data-result-index="${resultIndex}" onclick="window.navigateToSearchResult('${type}', ${resultIndex})">
                    <div class="search-result-primary">${primaryText}</div>
                    <div class="search-result-secondary">${secondaryText}</div>
                </div>`;
            });

            html += '</div>';
        }
    };

    addGroup('Sites', results.sites, 'sites', '🌐');
    addGroup('Users', results.users, 'users', '👤');
    addGroup('Groups', results.groups, 'groups', '👥');
    addGroup('Permissions', results.permissions, 'permissions', '🔐');
    addGroup('Inheritance', results.inheritance, 'inheritance', '🔗');

    resultsContainer.innerHTML = html;
    resultsContainer.classList.remove('hidden');
    selectedResultIndex = -1;
}

// Highlight selected result
function highlightSelectedResult() {
    const items = document.querySelectorAll('.search-result-item');
    items.forEach((item, index) => {
        if (index === selectedResultIndex) {
            item.classList.add('selected');
            item.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
        } else {
            item.classList.remove('selected');
        }
    });
}

// Navigate to selected search result
window.navigateToSearchResult = function(type, indexOrItem) {
    let item;

    if (typeof indexOrItem === 'number') {
        item = currentSearchResults[indexOrItem].item;
    } else {
        item = indexOrItem;
    }

    // Hide search results
    document.getElementById('global-search-results').classList.add('hidden');
    document.getElementById('global-search-input').value = '';

    // Switch to analytics tab
    const analyticsTab = document.querySelector('.tab-btn[data-tab="analytics"]');
    if (analyticsTab) {
        analyticsTab.click();
    }

    // Open appropriate deep dive with item pre-selected
    setTimeout(() => {
        if (type === 'sites') {
            openSiteDetailDeepDive(item.Title);
        } else if (type === 'users') {
            openDeepDive('users');
            setTimeout(() => {
                const searchInput = document.getElementById('dd-search');
                if (searchInput) {
                    searchInput.value = item.Name || item.Email;
                    searchInput.dispatchEvent(new Event('input'));
                }
            }, 100);
        } else if (type === 'groups') {
            openDeepDive('groups');
            setTimeout(() => {
                const searchInput = document.getElementById('dd-search');
                if (searchInput) {
                    searchInput.value = item.Name;
                    searchInput.dispatchEvent(new Event('input'));
                }
            }, 100);
        } else if (type === 'permissions') {
            openFilteredPermissionsDeepDive(item.Role);
        } else if (type === 'inheritance') {
            openDeepDive('inheritance');
            setTimeout(() => {
                const searchInput = document.getElementById('dd-search');
                if (searchInput) {
                    searchInput.value = item.Title;
                    searchInput.dispatchEvent(new Event('input'));
                }
            }, 100);
        }
    }, 200);
};

// --- Status polling (lightweight, once on load) ---
async function pollStatus() {
    try {
        const status = await API.getStatus();
        appState.headless = !!status.headless;

        // Show container auth note if headless
        if (appState.headless) {
            const note = document.getElementById('headless-note');
            if (note) note.classList.remove('hidden');
        }

        if (status.connected) {
            appState.connected = true;
            appState.demoMode = status.demoMode;
            if (status.metrics && status.metrics.totalSites > 0) {
                appState.dataLoaded = true;
            }
            updateConnectionUI(true);
            updateTabVisibility(true);
            // Show pre-connected message
            const results = document.getElementById('connection-results');
            if (results && !appState.demoMode) {
                results.textContent = 'Already connected to SharePoint Online.\nYou can use SharePoint Operations.';
            }
            if (appState.dataLoaded) await refreshAnalytics();
        }
    } catch (e) {
        // Server not ready yet, ignore
    }
}
