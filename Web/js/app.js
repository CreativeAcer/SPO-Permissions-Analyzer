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

// --- Initialization ---
document.addEventListener('DOMContentLoaded', () => {
    initTabs();
    initConnection();
    initOperations();
    initAnalytics();
    pollStatus();
});

// --- Tabs ---
function initTabs() {
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
            document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
            btn.classList.add('active');
            const target = document.getElementById('tab-' + btn.dataset.tab);
            if (target) target.classList.add('active');
        });
    });
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
        if (res.success) {
            const sites = await API.getData('sites');
            const siteList = sites.data || [];
            console_.textContent += `\nRetrieved ${siteList.length} sites:\n`;
            siteList.forEach((s, i) => {
                console_.textContent += `\n${i + 1}. ${s.Title || 'Unknown'}\n   URL: ${s.Url || 'N/A'}\n   Owner: ${s.Owner || 'N/A'}\n   Storage: ${s.Storage || '0'} MB\n`;
            });
            await refreshAnalytics();
            toast(`Retrieved ${siteList.length} sites`, 'success');
        } else {
            console_.textContent += `\nError: ${res.message}`;
        }
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
        // Fetch progress log
        const progress = await API.getProgress();
        if (progress.messages) {
            console_.textContent = progress.messages.join('\n');
        }

        if (res.success) {
            appState.dataLoaded = true;
            // Append final metrics
            const metrics = await API.getMetrics();
            console_.textContent += '\n\n=== ANALYSIS COMPLETE ===';
            console_.textContent += `\nUsers: ${metrics.totalUsers} | Groups: ${metrics.totalGroups} | External: ${metrics.externalUsers}`;
            console_.textContent += `\nRole Assignments: ${metrics.totalRoleAssignments} | Inheritance Breaks: ${metrics.inheritanceBreaks} | Sharing Links: ${metrics.totalSharingLinks}`;
            console_.textContent += '\n\nSwitch to Visual Analytics tab for charts and deep dives.';
            await refreshAnalytics();
            toast('Permissions analysis complete', 'success');
        } else {
            console_.textContent += `\nError: ${res.message}`;
        }
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

    try {
        // Export full governance JSON report
        const report = await API.exportJson();
        const blob = new Blob([JSON.stringify(report, null, 2)], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `spo_governance_${new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19)}.json`;
        a.click();
        URL.revokeObjectURL(url);
        toast('Governance JSON report downloaded', 'success');
    } catch (e) {
        toast('Report generation failed: ' + e.message, 'error');
    }
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
        const metrics = await API.getMetrics();

        // Update metric cards
        setText('metric-sites', metrics.totalSites);
        setText('metric-users', metrics.totalUsers);
        setText('metric-groups', metrics.totalGroups);
        setText('metric-external', metrics.externalUsers);
        setText('metric-roles', metrics.totalRoleAssignments);
        setText('metric-inheritance', metrics.inheritanceBreaks);
        setText('metric-sharing', metrics.totalSharingLinks);

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

    } catch (e) {
        console.error('Failed to refresh analytics:', e);
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
        <div class="dd-filter-bar"><input type="text" placeholder="Search sites..." id="dd-search"><button class="btn btn-secondary" onclick="API.exportData('sites')">Export CSV</button></div>
        <table><thead><tr><th>Title</th><th>URL</th><th>Owner</th><th>Storage (MB)</th><th>Template</th></tr></thead>
        <tbody id="dd-sites-body">${renderSitesRows(data)}</tbody></table>`;

    document.getElementById('dd-search').addEventListener('input', (e) => {
        const q = e.target.value.toLowerCase();
        const filtered = data.filter(s => (s.Title || '').toLowerCase().includes(q) || (s.Url || '').toLowerCase().includes(q));
        document.getElementById('dd-sites-body').innerHTML = renderSitesRows(filtered);
    });
}

function renderSitesRows(data) {
    return data.map(s => `<tr><td>${esc(s.Title)}</td><td>${esc(s.Url)}</td><td>${esc(s.Owner)}</td><td>${s.Storage || 0}</td><td>${esc(s.Template || 'N/A')}</td></tr>`).join('');
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
        <button class="btn btn-secondary" onclick="API.exportData('users')">Export CSV</button></div>
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
        <div class="dd-filter-bar"><input type="text" placeholder="Search groups..." id="dd-search"><button class="btn btn-secondary" onclick="API.exportData('groups')">Export CSV</button></div>
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

    container.innerHTML = `
        <div class="dd-stats">
            <div class="dd-stat"><span class="dd-stat-value" style="color:#C62828">${data.length}</span><span class="dd-stat-label">External Users</span></div>
            <div class="dd-stat"><span class="dd-stat-value">${Object.keys(domains).length}</span><span class="dd-stat-label">Domains</span></div>
            <div class="dd-stat"><span class="dd-stat-value" style="color:#DC3545">${editAccess}</span><span class="dd-stat-label">With Edit+</span></div>
        </div>
        ${editAccess > 0 ? '<div class="finding high"><h4>External Users with Edit Access</h4><p>' + editAccess + ' external user(s) have edit or higher permissions. Review and restrict where possible.</p></div>' : ''}
        <div class="dd-filter-bar"><input type="text" placeholder="Search external users..." id="dd-search"><button class="btn btn-secondary" onclick="API.exportData('users')">Export CSV</button></div>
        <table><thead><tr><th>Name</th><th>Email</th><th>Domain</th><th>Permission</th></tr></thead>
        <tbody id="dd-ext-body">${renderExternalRows(data)}</tbody></table>`;

    document.getElementById('dd-search').addEventListener('input', (e) => {
        const q = e.target.value.toLowerCase();
        const filtered = data.filter(u => (u.Name || '').toLowerCase().includes(q) || (u.Email || '').toLowerCase().includes(q));
        document.getElementById('dd-ext-body').innerHTML = renderExternalRows(filtered);
    });
}

function renderExternalRows(data) {
    return data.map(u => {
        const email = u.Email || '';
        const domain = email.includes('@') ? email.split('@')[1] : 'Unknown';
        return `<tr><td>${esc(u.Name)}</td><td>${esc(email)}</td><td>${esc(domain)}</td><td>${esc(u.Permission)}</td></tr>`;
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
            <button class="btn btn-secondary" onclick="API.exportData('roleassignments')">Export CSV</button></div>
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

    container.innerHTML = `
        <div class="dd-stats">
            <div class="dd-stat"><span class="dd-stat-value">${data.length}</span><span class="dd-stat-label">Total Items</span></div>
            <div class="dd-stat"><span class="dd-stat-value" style="color:#28A745">${inheriting}</span><span class="dd-stat-label">Inheriting</span></div>
            <div class="dd-stat"><span class="dd-stat-value" style="color:#DC3545">${broken}</span><span class="dd-stat-label">Broken</span></div>
            <div class="dd-stat"><span class="dd-stat-value">${libraries}</span><span class="dd-stat-label">Libraries</span></div>
            <div class="dd-stat"><span class="dd-stat-value">${lists}</span><span class="dd-stat-label">Lists</span></div>
        </div>
        <div class="dd-tabs">
            <button class="dd-tab-btn active" data-ddtab="dd-table">Inheritance Tree</button>
            <button class="dd-tab-btn" data-ddtab="dd-chart">Overview</button>
            <button class="dd-tab-btn" data-ddtab="dd-findings">Findings</button>
        </div>
        <div id="dd-table" class="dd-tab-content active">
            <div class="dd-filter-bar"><input type="text" placeholder="Search..." id="dd-search">
            <select id="dd-inh-filter"><option value="">All Items</option><option value="broken">Broken Inheritance</option><option value="inheriting">Inheriting</option></select>
            <button class="btn btn-secondary" onclick="API.exportData('inheritance')">Export CSV</button></div>
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
            <button class="btn btn-secondary" onclick="API.exportData('sharinglinks')">Export CSV</button></div>
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

function toast(message, type = 'info') {
    const container = document.getElementById('toast-container');
    if (!container) return;
    const t = document.createElement('div');
    t.className = `toast ${type}`;
    t.textContent = message;
    container.appendChild(t);
    setTimeout(() => t.remove(), 4000);
}

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
