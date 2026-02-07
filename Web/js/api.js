// ============================================
// api.js - Backend communication layer
// ============================================
// All fetch calls to the PowerShell HTTP server.

const API = {
    async get(endpoint) {
        const res = await fetch(`/api/${endpoint}`);
        if (!res.ok) throw new Error(`API error: ${res.status}`);
        return res.json();
    },

    async post(endpoint, body = {}) {
        const res = await fetch(`/api/${endpoint}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(body)
        });
        if (!res.ok) throw new Error(`API error: ${res.status}`);
        return res.json();
    },

    // --- Specific endpoints ---

    getStatus() {
        return this.get('status');
    },

    connect(tenantUrl, clientId) {
        return this.post('connect', { tenantUrl, clientId });
    },

    startDemo() {
        return this.post('demo');
    },

    getSites() {
        return this.post('sites');
    },

    analyzePermissions(siteUrl = '') {
        return this.post('permissions', { siteUrl });
    },

    getProgress() {
        return this.get('progress');
    },

    getData(type) {
        return this.get(`data/${type}`);
    },

    getMetrics() {
        return this.get('metrics');
    },

    enrichExternal() {
        return this.post('enrich');
    },

    getEnrichment() {
        return this.get('enrichment');
    },

    getRisk() {
        return this.get('risk');
    },

    getAudit() {
        return this.get('audit');
    },

    exportData(type) {
        // Returns a download, not JSON
        window.open(`/api/export/${type}`, '_blank');
    },

    async exportJson() {
        // Returns full governance JSON report
        return this.get('export-json');
    },

    async exportJsonType(type) {
        // Returns typed JSON export
        return this.get(`export-json/${type}`);
    }
};
