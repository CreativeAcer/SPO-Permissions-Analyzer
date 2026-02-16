// ============================================
// app-state.js - Global state and shared utilities
// ============================================

// --- State ---
let appState = {
    connected: false,
    demoMode: false,
    dataLoaded: false,
    headless: false
};

// --- Utilities ---
function setText(id, value) {
    const el = document.getElementById(id);
    if (el) el.textContent = value;
}

function esc(str) {
    if (!str) return '';
    const div = document.createElement('div');
    div.textContent = String(str);
    return div.innerHTML;
}

function formatStorage(mb) {
    if (mb >= 1024) return (mb / 1024).toFixed(1) + ' GB';
    return mb + ' MB';
}

function setButtonLoading(id, loading) {
    // Use UIHelpers if available for modern loading spinner
    if (typeof UIHelpers !== 'undefined') {
        UIHelpers.setButtonLoading(id, loading);
    } else {
        // Fallback to text-based loading
        const btn = document.getElementById(id);
        if (!btn) return;
        if (loading) {
            btn.dataset.originalText = btn.textContent;
            btn.textContent = 'Loading...';
            btn.classList.add('btn-disabled');
        } else {
            btn.textContent = btn.dataset.originalText || btn.textContent;
            btn.classList.remove('btn-disabled');
        }
    }
}

function toast(message, type = 'info') {
    const container = document.getElementById('toast-container');
    if (!container) return;
    const t = document.createElement('div');
    t.className = `toast ${type}`;
    t.textContent = message;
    container.appendChild(t);
    setTimeout(() => t.remove(), 4000);
}

// --- Operation polling helper ---
// Polls /api/progress until the operation is complete, updating the console element.
async function pollUntilComplete(consoleEl, intervalMs = 1000) {
    while (true) {
        await new Promise(r => setTimeout(r, intervalMs));
        try {
            const progress = await API.getProgress();
            if (progress.messages) {
                consoleEl.textContent = progress.messages.join('\n');
            }
            if (!progress.running && progress.complete) {
                return progress;
            }
        } catch (e) {
            // Transient fetch failure — keep polling
        }
    }
}

// --- Audit Summary ---
async function showAuditSummary(console_) {
    try {
        const audit = await API.getAudit();
        if (!audit.hasSession) return;

        console_.textContent += '\n\n=== AUDIT TRAIL ===';
        console_.textContent += `\nSession ID: ${audit.sessionId}`;
        console_.textContent += `\nOperation: ${audit.operationType}`;
        console_.textContent += `\nStatus: ${audit.status}`;
        console_.textContent += `\nUser: ${audit.userPrincipal}`;
        if (audit.duration) console_.textContent += `\nDuration: ${audit.duration}`;
        console_.textContent += `\nEvents: ${audit.eventCount} | Errors: ${audit.errorCount}`;
        if (audit.outputFiles && audit.outputFiles.length > 0) {
            console_.textContent += `\nOutput files: ${audit.outputFiles.length}`;
        }
        console_.textContent += '\n(Full audit log saved to ./Logs/)';
    } catch (e) {
        // Audit info is supplementary, don't fail the operation
    }
}
