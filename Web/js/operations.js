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
    UIHelpers.setButtonLoading('btn-get-sites', true);

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
        UIHelpers.setButtonLoading('btn-get-sites', false);
    }
}

async function handleAnalyze() {
    console.log('handleAnalyze called');
    const inputSiteUrl = document.getElementById('input-site-url').value.trim();
    const siteUrl = inputSiteUrl || appState.connectedSiteUrl;
    const console_ = document.getElementById('operations-console');

    console.log('Site URL for analysis:', siteUrl);
    console.log('Connected site URL:', appState.connectedSiteUrl);

    if (!siteUrl) {
        console_.textContent = 'Error: No site URL available. Please enter a site URL or reconnect to SharePoint.';
        toast('No site URL available', 'error');
        return;
    }

    UIHelpers.setButtonLoading('btn-analyze', true);

    try {
        // Step 1: Prepare analysis - check if re-auth is needed
        console.log('Preparing analysis...');
        const prepareRes = await API.prepareAnalysis(siteUrl);

        if (!prepareRes.success && prepareRes.message) {
            console_.textContent = `Error: ${prepareRes.message}`;
            toast('Preparation failed', 'error');
            return;
        }

        // If re-auth is needed, show message immediately
        if (prepareRes.needsAuth) {
            console_.textContent = '';
            console_.textContent += '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n';
            console_.textContent += '📋 AUTHENTICATION REQUIRED\n';
            console_.textContent += '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n';
            console_.textContent += '\n';
            console_.textContent += 'Analyzing a different site requires re-authentication.\n';
            console_.textContent += '\n';
            console_.textContent += '👉 CHECK YOUR TERMINAL FOR DEVICE CODE\n';
            console_.textContent += '\n';
            console_.textContent += 'Run: podman logs <container>\n';
            console_.textContent += '\n';
            console_.textContent += '🌐 Opening device login page in new tab...\n';
            console_.textContent += '\n';
            console_.textContent += '⏳ Waiting for authentication...\n';

            // Automatically open device login page in new tab
            const loginWindow = window.open('https://microsoft.com/devicelogin', '_blank');
            if (!loginWindow) {
                console_.textContent += '\n⚠️  Popup blocked! Manually visit: https://microsoft.com/devicelogin\n';
            }

            // Give user a moment to see the message
            await new Promise(r => setTimeout(r, 500));
        } else {
            console_.textContent = `Starting permissions analysis for: ${siteUrl}\n`;
        }

        // Step 2: Execute analysis (will do auth if needed, then analyze)
        console.log('Calling API.analyzePermissions...');
        const res = await API.analyzePermissions(siteUrl);
        console.log('API.analyzePermissions response:', res);

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
        toast('Analysis failed', 'error');
    } finally {
        UIHelpers.setButtonLoading('btn-analyze', false);
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
