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
