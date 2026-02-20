// ============================================
// charts.js - Chart.js wrappers (Enhanced)
// ============================================

let storageChart = null;
let permissionChart = null;

// Modern color palette with gradients
const COLORS = {
    blue: '#1976D2',
    green: '#10B981',
    orange: '#F97316',
    red: '#EF4444',
    purple: '#A855F7',
    teal: '#14B8A6',
    amber: '#F59E0B',
    grey: '#71717A',
    // Permission-specific (updated to modern palette)
    fullControl: '#EF4444',
    edit: '#F59E0B',
    contribute: '#F97316',
    read: '#10B981',
    viewOnly: '#3B82F6',
    limited: '#71717A',
    custom: '#A855F7'
};

// Gradient definitions for charts
const GRADIENTS = {
    blue: ['#3B82F6', '#2563EB'],
    green: ['#10B981', '#059669'],
    orange: ['#F97316', '#EA580C'],
    red: ['#EF4444', '#DC2626'],
    purple: ['#A855F7', '#9333EA'],
    amber: ['#F59E0B', '#D97706']
};

// Chart.js default configuration
Chart.defaults.font.family = "'Segoe UI', -apple-system, system-ui, 'Inter', sans-serif";
Chart.defaults.font.size = 13;
Chart.defaults.color = '#94A3B8';
Chart.defaults.plugins.legend.labels.usePointStyle = true;
Chart.defaults.plugins.legend.labels.padding = 12;
Chart.defaults.plugins.tooltip.backgroundColor = 'rgba(24, 24, 27, 0.95)';
Chart.defaults.plugins.tooltip.cornerRadius = 8;
Chart.defaults.plugins.tooltip.padding = 12;
Chart.defaults.plugins.tooltip.titleFont.weight = '600';
Chart.defaults.plugins.tooltip.bodyFont.size = 13;

function renderStorageChart(sites) {
    const canvas = document.getElementById('chart-storage');
    if (!canvas) return;

    if (storageChart) storageChart.destroy();

    if (!sites || sites.length === 0) {
        storageChart = new Chart(canvas, {
            type: 'bar',
            data: { labels: ['No data'], datasets: [{ data: [0], backgroundColor: '#2E2E4A' }] },
            options: {
                plugins: { legend: { display: false } },
                scales: {
                    y: {
                        beginAtZero: true,
                        grid: { color: '#2E2E4A' }
                    },
                    x: { grid: { display: false } }
                }
            }
        });
        return;
    }

    // Sort by storage descending, take top 10
    const sorted = [...sites]
        .map(s => ({ title: s.Title || 'Unknown', storage: parseInt(s.Storage) || 0 }))
        .sort((a, b) => b.storage - a.storage)
        .slice(0, 10);

    // Create gradient colors
    const ctx = canvas.getContext('2d');
    const colors = sorted.map(s => {
        const gradient = ctx.createLinearGradient(0, 0, 0, canvas.height);
        if (s.storage >= 1500) {
            gradient.addColorStop(0, GRADIENTS.purple[0]);
            gradient.addColorStop(1, GRADIENTS.purple[1]);
        } else if (s.storage >= 1000) {
            gradient.addColorStop(0, GRADIENTS.red[0]);
            gradient.addColorStop(1, GRADIENTS.red[1]);
        } else if (s.storage >= 500) {
            gradient.addColorStop(0, GRADIENTS.orange[0]);
            gradient.addColorStop(1, GRADIENTS.orange[1]);
        } else {
            gradient.addColorStop(0, GRADIENTS.green[0]);
            gradient.addColorStop(1, GRADIENTS.green[1]);
        }
        return gradient;
    });

    storageChart = new Chart(canvas, {
        type: 'bar',
        data: {
            labels: sorted.map(s => s.title.length > 20 ? s.title.substring(0, 20) + '...' : s.title),
            datasets: [{
                label: 'Storage (MB)',
                data: sorted.map(s => s.storage),
                backgroundColor: colors,
                borderRadius: 8,
                borderSkipped: false
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            animation: {
                duration: 800,
                easing: 'easeOutQuart'
            },
            onClick: (event, activeElements) => {
                if (activeElements && activeElements.length > 0) {
                    const index = activeElements[0].index;
                    const siteName = sorted[index].title;
                    if (window.openSiteDetailDeepDive) {
                        window.openSiteDetailDeepDive(siteName);
                    }
                }
            },
            plugins: {
                legend: { display: false },
                tooltip: {
                    callbacks: {
                        label: ctx => `Storage: ${ctx.parsed.y} MB`,
                        title: ctx => sorted[ctx[0].dataIndex].title
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    grid: {
                        color: '#2E2E4A',
                        drawBorder: false
                    },
                    ticks: {
                        callback: value => `${value} MB`,
                        padding: 8
                    }
                },
                x: {
                    grid: { display: false },
                    ticks: {
                        maxRotation: 45,
                        minRotation: 0,
                        padding: 8
                    }
                }
            },
            interaction: {
                intersect: false,
                mode: 'index'
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
            data: { labels: ['No data'], datasets: [{ data: [1], backgroundColor: ['#2E2E4A'] }] },
            options: {
                plugins: {
                    legend: {
                        position: 'right',
                        labels: { padding: 12 }
                    }
                }
            }
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
                borderWidth: 3,
                borderColor: '#1A1A2E',
                hoverBorderWidth: 4,
                hoverOffset: 8
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            animation: {
                duration: 1000,
                easing: 'easeOutQuart',
                animateRotate: true,
                animateScale: true
            },
            onClick: (event, activeElements) => {
                if (activeElements && activeElements.length > 0) {
                    const index = activeElements[0].index;
                    const permissionLevel = labels[index];
                    if (window.openFilteredPermissionsDeepDive) {
                        window.openFilteredPermissionsDeepDive(permissionLevel);
                    }
                }
            },
            plugins: {
                legend: {
                    position: 'right',
                    labels: {
                        padding: 12,
                        usePointStyle: true,
                        pointStyle: 'circle',
                        font: {
                            size: 12,
                            weight: '500'
                        },
                        generateLabels: (chart) => {
                            const data = chart.data;
                            if (data.labels.length && data.datasets.length) {
                                return data.labels.map((label, i) => {
                                    const value = data.datasets[0].data[i];
                                    const total = data.datasets[0].data.reduce((a, b) => a + b, 0);
                                    const percentage = ((value / total) * 100).toFixed(1);
                                    return {
                                        text: `${label} (${percentage}%)`,
                                        fillStyle: data.datasets[0].backgroundColor[i],
                                        hidden: false,
                                        index: i
                                    };
                                });
                            }
                            return [];
                        }
                    }
                },
                tooltip: {
                    callbacks: {
                        label: ctx => {
                            const label = ctx.label || '';
                            const value = ctx.parsed;
                            const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
                            const percentage = ((value / total) * 100).toFixed(1);
                            return `${label}: ${value} (${percentage}%)`;
                        }
                    }
                }
            },
            cutout: '60%'
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
        // Create gradients for bar chart
        const ctx = canvas.getContext('2d');
        const gradients = data.map(d => {
            const gradient = ctx.createLinearGradient(0, 0, 0, canvas.height);
            const color = d.color || COLORS.blue;
            gradient.addColorStop(0, color);
            gradient.addColorStop(1, color + 'CC'); // Add transparency
            return gradient;
        });

        new Chart(canvas, {
            type: 'bar',
            data: {
                labels: data.map(d => d.label),
                datasets: [{
                    data: data.map(d => d.value),
                    backgroundColor: gradients,
                    borderRadius: 8,
                    borderSkipped: false
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                animation: {
                    duration: 800,
                    easing: 'easeOutQuart'
                },
                onClick: (event, activeElements) => {
                    if (activeElements && activeElements.length > 0) {
                        const index = activeElements[0].index;
                        const clickedData = data[index];
                        if (window.onDeepDiveChartClick) {
                            window.onDeepDiveChartClick(canvasId, clickedData);
                        }
                    }
                },
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        callbacks: {
                            label: ctx => `${ctx.label}: ${ctx.parsed.y}`
                        }
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        grid: {
                            color: '#2E2E4A',
                            drawBorder: false
                        },
                        ticks: { padding: 8 }
                    },
                    x: {
                        grid: { display: false },
                        ticks: { padding: 8 }
                    }
                },
                interaction: {
                    intersect: false,
                    mode: 'index'
                }
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
                    borderWidth: 3,
                    borderColor: '#1A1A2E',
                    hoverBorderWidth: 4,
                    hoverOffset: 8
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                animation: {
                    duration: 1000,
                    easing: 'easeOutQuart',
                    animateRotate: true,
                    animateScale: true
                },
                onClick: (event, activeElements) => {
                    if (activeElements && activeElements.length > 0) {
                        const index = activeElements[0].index;
                        const clickedData = data[index];
                        if (window.onDeepDiveChartClick) {
                            window.onDeepDiveChartClick(canvasId, clickedData);
                        }
                    }
                },
                plugins: {
                    legend: {
                        position: 'right',
                        labels: {
                            padding: 12,
                            usePointStyle: true,
                            pointStyle: 'circle',
                            font: {
                                size: 12,
                                weight: '500'
                            },
                            generateLabels: (chart) => {
                                const data = chart.data;
                                if (data.labels.length && data.datasets.length) {
                                    return data.labels.map((label, i) => {
                                        const value = data.datasets[0].data[i];
                                        const total = data.datasets[0].data.reduce((a, b) => a + b, 0);
                                        const percentage = ((value / total) * 100).toFixed(1);
                                        return {
                                            text: `${label} (${percentage}%)`,
                                            fillStyle: data.datasets[0].backgroundColor[i],
                                            hidden: false,
                                            index: i
                                        };
                                    });
                                }
                                return [];
                            }
                        }
                    },
                    tooltip: {
                        callbacks: {
                            label: ctx => {
                                const label = ctx.label || '';
                                const value = ctx.parsed;
                                const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
                                const percentage = ((value / total) * 100).toFixed(1);
                                return `${label}: ${value} (${percentage}%)`;
                            }
                        }
                    }
                },
                cutout: '60%'
            }
        });
    }
}
