// ============================================
// permissions-matrix.js - Permissions matrix functionality
// ============================================

// Permissions matrix state
let currentMatrixData = null;
let currentMatrixSiteUrl = null;

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
