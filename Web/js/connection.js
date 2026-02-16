// ============================================
// connection.js - Connection tab logic
// ============================================

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
