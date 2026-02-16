// ============================================
// deep-dives.js - Deep dive renderers and helpers
// ============================================

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
