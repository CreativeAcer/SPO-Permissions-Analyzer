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

            // Store capabilities in app state
            appState.capabilities = res.capabilities || {};

            results.textContent = `Connected successfully!\n\nSite: ${res.siteTitle || 'N/A'}\nURL: ${res.siteUrl || 'N/A'}\nUser: ${res.user || 'N/A'}\n\nYou can now use SharePoint Operations.`;

            // Display capabilities
            displayCapabilities(res.capabilities);

            updateConnectionUI(true);
            updateTabVisibility(true);

            // Conditionally enable operations based on capabilities
            updateOperationsButtons(res.capabilities);

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

            // Demo mode has all capabilities
            appState.capabilities = {
                CanEnumerateSites: true,
                CanReadUsers: true,
                CanAccessStorageData: true,
                CanReadExternalUsers: true,
                CheckedAt: new Date().toISOString()
            };

            results.textContent = 'Demo Mode activated!\n\nSample data has been generated.\nSwitch to Operations or Visual Analytics tab to explore.';

            // Display demo capabilities
            displayCapabilities(appState.capabilities);

            updateConnectionUI(true);
            updateTabVisibility(true);

            // Enable all buttons in demo mode
            updateOperationsButtons(appState.capabilities);

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

// --- Capability Display ---
function displayCapabilities(capabilities) {
    const container = document.getElementById('capability-status');
    if (!container || !capabilities) return;

    const items = [
        {
            key: 'CanEnumerateSites',
            label: 'Site enumeration',
            enabledNote: 'Can retrieve all tenant sites',
            disabledNote: 'Requires SharePoint Administrator role'
        },
        {
            key: 'CanReadUsers',
            label: 'User iteration',
            enabledNote: 'Can enumerate tenant users',
            disabledNote: 'Requires User.Read.All permission'
        },
        {
            key: 'CanAccessStorageData',
            label: 'Storage data',
            enabledNote: 'Can view site storage metrics',
            disabledNote: 'Limited storage information available'
        },
        {
            key: 'CanReadExternalUsers',
            label: 'External user data',
            enabledNote: 'Can query external/guest users',
            disabledNote: 'Requires User.Read.All permission'
        }
    ];

    const itemsHtml = items.map(item => {
        const enabled = capabilities[item.key] === true;
        const icon = enabled ? '✓' : '✗';
        const iconClass = enabled ? 'enabled' : 'disabled';
        const note = enabled ? item.enabledNote : item.disabledNote;

        return `
            <div class="capability-item">
                <span class="capability-icon ${iconClass}">${icon}</span>
                <div class="capability-text">
                    <span class="capability-label">${item.label}:</span>
                    <span class="${enabled ? 'text-success' : 'text-warning'}">
                        ${enabled ? 'Enabled' : 'Disabled'}
                    </span>
                    <span class="capability-note">${note}</span>
                </div>
            </div>
        `;
    }).join('');

    container.innerHTML = `
        <h4>Your Capabilities</h4>
        <div class="capability-list">
            ${itemsHtml}
        </div>
    `;

    container.classList.remove('hidden');
}

// --- Update Operations Buttons Based on Capabilities ---
function updateOperationsButtons(capabilities) {
    const getSitesBtn = document.getElementById('btn-get-sites');

    if (!capabilities || !capabilities.CanEnumerateSites) {
        // Keep button disabled and add tooltip
        getSitesBtn.classList.add('btn-disabled', 'btn-with-tooltip');

        // Wrap in tooltip if not already wrapped
        if (!getSitesBtn.parentElement.classList.contains('tooltip-wrapper')) {
            const wrapper = document.createElement('div');
            wrapper.className = 'tooltip-wrapper';
            getSitesBtn.parentNode.insertBefore(wrapper, getSitesBtn);
            wrapper.appendChild(getSitesBtn);

            const tooltip = document.createElement('span');
            tooltip.className = 'tooltip';
            tooltip.textContent = 'SharePoint Administrator role required to enumerate all tenant sites';
            wrapper.appendChild(tooltip);
        }
    } else {
        // User has admin rights - enable the button
        getSitesBtn.classList.remove('btn-disabled', 'btn-with-tooltip');

        // Remove tooltip wrapper if it exists
        if (getSitesBtn.parentElement.classList.contains('tooltip-wrapper')) {
            const wrapper = getSitesBtn.parentElement;
            const parent = wrapper.parentNode;
            parent.insertBefore(getSitesBtn, wrapper);
            wrapper.remove();
        }
    }
}
