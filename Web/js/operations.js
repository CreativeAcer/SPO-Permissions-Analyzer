// ============================================
// operations.js - Operations tab logic
// ============================================

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
