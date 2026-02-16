// ============================================
// export.js - Export format selection and handlers
// ============================================

let pendingExportType = null;
let pendingExportIsFullReport = false;

// Show export format selection modal
window.showExportModal = function(type, isFullReport = false) {
    const modal = document.getElementById('export-modal');
    const closeBtn = document.getElementById('export-modal-close');

    pendingExportType = type;
    pendingExportIsFullReport = isFullReport;

    modal.classList.remove('hidden');

    // Close handlers
    closeBtn.onclick = () => modal.classList.add('hidden');
    modal.onclick = (e) => { if (e.target === modal) modal.classList.add('hidden'); };

    // Escape key
    const escHandler = (e) => {
        if (e.key === 'Escape') {
            modal.classList.add('hidden');
            document.removeEventListener('keydown', escHandler);
        }
    };
    document.addEventListener('keydown', escHandler);
};

// Handle format selection
function initExportModal() {
    document.querySelectorAll('.export-format-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const format = btn.dataset.format;
            const modal = document.getElementById('export-modal');

            modal.classList.add('hidden');

            // Perform export based on format
            if (pendingExportIsFullReport) {
                handleReportExport(format);
            } else {
                handleDataExport(pendingExportType, format);
            }
        });
    });
}

// Export data in selected format
function handleDataExport(type, format) {
    if (type === 'permissions-matrix') {
        if (format === 'csv') {
            exportMatrixToCSV(currentMatrixData);
        } else if (format === 'json') {
            exportMatrixToJSON(currentMatrixData);
        }
    } else {
        if (format === 'csv') {
            API.exportData(type);
            toast(`Exporting ${type} as CSV`, 'success');
        } else if (format === 'json') {
            API.exportDataJson(type);
            toast(`Exporting ${type} as JSON`, 'success');
        }
    }
}

// Export full report in selected format
async function handleReportExport(format) {
    if (!appState.dataLoaded) {
        toast('Run an analysis first', 'info');
        return;
    }

    try {
        if (format === 'json') {
            const report = await API.exportJson();
            const blob = new Blob([JSON.stringify(report, null, 2)], { type: 'application/json' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `spo_governance_${new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19)}.json`;
            a.click();
            URL.revokeObjectURL(url);
            toast('Governance JSON report downloaded', 'success');
        } else if (format === 'csv') {
            // Export all data types as separate CSV files
            const types = ['sites', 'users', 'groups', 'roleassignments', 'inheritance', 'sharinglinks'];
            types.forEach(type => API.exportData(type));
            toast('Exporting all data as CSV files', 'success');
        }
    } catch (e) {
        toast('Export failed: ' + e.message, 'error');
    }
}
