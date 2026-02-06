// ============================================
// charts.js - Chart.js wrappers
// ============================================

let storageChart = null;
let permissionChart = null;

const COLORS = {
    blue: '#1976D2',
    green: '#2E7D32',
    orange: '#F57C00',
    red: '#C62828',
    purple: '#7B1FA2',
    teal: '#00897B',
    amber: '#F57F17',
    grey: '#6C757D',
    // Permission-specific
    fullControl: '#DC3545',
    edit: '#FFC107',
    contribute: '#FD7E14',
    read: '#28A745',
    viewOnly: '#17A2B8',
    limited: '#6C757D',
    custom: '#6F42C1'
};

function renderStorageChart(sites) {
    const canvas = document.getElementById('chart-storage');
    if (!canvas) return;

    if (storageChart) storageChart.destroy();

    if (!sites || sites.length === 0) {
        storageChart = new Chart(canvas, {
            type: 'bar',
            data: { labels: ['No data'], datasets: [{ data: [0], backgroundColor: '#E1DFDD' }] },
            options: { plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true } } }
        });
        return;
    }

    // Sort by storage descending, take top 10
    const sorted = [...sites]
        .map(s => ({ title: s.Title || 'Unknown', storage: parseInt(s.Storage) || 0 }))
        .sort((a, b) => b.storage - a.storage)
        .slice(0, 10);

    const colors = sorted.map(s => {
        if (s.storage >= 1500) return COLORS.purple;
        if (s.storage >= 1000) return COLORS.red;
        if (s.storage >= 500) return COLORS.orange;
        return COLORS.green;
    });

    storageChart = new Chart(canvas, {
        type: 'bar',
        data: {
            labels: sorted.map(s => s.title.length > 20 ? s.title.substring(0, 20) + '...' : s.title),
            datasets: [{
                label: 'Storage (MB)',
                data: sorted.map(s => s.storage),
                backgroundColor: colors,
                borderRadius: 4
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                tooltip: {
                    callbacks: {
                        label: ctx => `${ctx.parsed.y} MB`
                    }
                }
            },
            scales: {
                y: { beginAtZero: true, title: { display: true, text: 'MB' } },
                x: { ticks: { maxRotation: 45 } }
            }
        }
    });
}

function renderPermissionChart(users, groups) {
    const canvas = document.getElementById('chart-permissions');
    if (!canvas) return;

    if (permissionChart) permissionChart.destroy();

    // Aggregate permission counts
    const counts = {};
    const allItems = [...(users || []), ...(groups || [])];

    if (allItems.length === 0) {
        permissionChart = new Chart(canvas, {
            type: 'doughnut',
            data: { labels: ['No data'], datasets: [{ data: [1], backgroundColor: ['#E1DFDD'] }] },
            options: { plugins: { legend: { position: 'right' } } }
        });
        return;
    }

    allItems.forEach(item => {
        const perm = item.Permission || item.Role || 'Unknown';
        counts[perm] = (counts[perm] || 0) + 1;
    });

    const colorMap = {
        'Full Control': COLORS.fullControl,
        'Edit': COLORS.edit,
        'Contribute': COLORS.contribute,
        'Read': COLORS.read,
        'View Only': COLORS.viewOnly,
        'Limited Access': COLORS.limited,
        'Member': COLORS.blue,
        'Group Permission': COLORS.orange
    };

    const labels = Object.keys(counts);
    const data = Object.values(counts);
    const bgColors = labels.map(l => colorMap[l] || COLORS.custom);

    permissionChart = new Chart(canvas, {
        type: 'doughnut',
        data: {
            labels,
            datasets: [{
                data,
                backgroundColor: bgColors,
                borderWidth: 2,
                borderColor: '#fff'
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'right',
                    labels: { font: { size: 11 }, padding: 8 }
                }
            }
        }
    });
}

function renderDeepDiveChart(canvasId, type, data) {
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;

    // Destroy existing chart on this canvas
    const existing = Chart.getChart(canvas);
    if (existing) existing.destroy();

    if (!data || data.length === 0) return;

    if (type === 'bar') {
        new Chart(canvas, {
            type: 'bar',
            data: {
                labels: data.map(d => d.label),
                datasets: [{
                    data: data.map(d => d.value),
                    backgroundColor: data.map(d => d.color || COLORS.blue),
                    borderRadius: 4
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { legend: { display: false } },
                scales: { y: { beginAtZero: true } }
            }
        });
    } else if (type === 'doughnut') {
        new Chart(canvas, {
            type: 'doughnut',
            data: {
                labels: data.map(d => d.label),
                datasets: [{
                    data: data.map(d => d.value),
                    backgroundColor: data.map(d => d.color || COLORS.blue),
                    borderWidth: 2,
                    borderColor: '#fff'
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { legend: { position: 'right', labels: { font: { size: 11 } } } }
            }
        });
    }
}
