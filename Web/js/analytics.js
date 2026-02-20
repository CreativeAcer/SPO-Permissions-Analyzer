// ============================================
// analytics.js - Analytics tab, risk banner, sites table, alerts
// ============================================

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
            // Map to existing CSS severity modifier so the dark-theme gradient background is applied
            const sevClass = { Critical: 'high', High: 'high', Medium: 'medium', Low: 'low' }[sev] || 'info';
            return `<div class="finding finding-item ${sevClass}" data-severity="${sev}" style="border-left-color: ${color}; cursor: pointer;" onclick="toggleFindingDetails(${idx})">
                <div style="display: flex; justify-content: space-between; align-items: start;">
                    <h4 style="margin: 0 0 8px 0; font-size: var(--font-size-md); color: var(--color-text-primary);">[${esc(f.RuleId || f.ruleId)}] ${esc(f.Title || f.title)}</h4>
                    <span style="background: ${color}; color: white; padding: 2px 8px; border-radius: var(--radius-sm); font-size: var(--font-size-xs); font-weight: var(--font-weight-semibold);">${sev}</span>
                </div>
                <div style="display: flex; gap: 16px; margin-bottom: 8px; font-size: var(--font-size-sm); color: var(--color-text-secondary);">
                    <span><strong>Category:</strong> ${esc(category)}</span>
                    ${count > 0 ? `<span><strong>Affected Items:</strong> ${count}</span>` : ''}
                    <span><strong>Score:</strong> ${f.Score || f.score}/100</span>
                </div>
                <p style="margin: 0; color: var(--color-text-primary); font-size: var(--font-size-base);">${esc(f.Description || f.description)}</p>
                <div id="finding-details-${idx}" style="display: none; margin-top: 12px; padding-top: 12px; border-top: 1px solid var(--color-border);">
                    <p style="margin: 0; font-size: var(--font-size-sm); color: var(--color-text-secondary);"><strong>Recommendation:</strong> Review and remediate this finding to improve your security posture.</p>
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
