// ============================================
// app.js - Application entry point and tab routing
// ============================================

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

// --- Startup status check (once on load, restores existing session state) ---
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

            // For pre-existing connections, assume full capabilities if demo mode
            // otherwise assume limited capabilities (user should re-connect for fresh check)
            if (status.demoMode) {
                appState.capabilities = {
                    CanEnumerateSites: true,
                    CanReadUsers: true,
                    CanAccessStorageData: true,
                    CanReadExternalUsers: true
                };
            } else {
                // Assume limited capabilities on reconnect
                // User can re-connect to get fresh capability check
                appState.capabilities = {
                    CanEnumerateSites: false,
                    CanReadUsers: false,
                    CanAccessStorageData: false,
                    CanReadExternalUsers: false
                };
            }

            updateConnectionUI(true);
            updateTabVisibility(true);
            updateOperationsButtons(appState.capabilities);

            // Show pre-connected message
            const results = document.getElementById('connection-results');
            if (results && !appState.demoMode) {
                results.textContent = 'Already connected to SharePoint Online.\nRe-connect to refresh capability status.';
            }
            if (appState.dataLoaded) await refreshAnalytics();
        }
    } catch (e) {
        // Server not ready yet, ignore
    }
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
