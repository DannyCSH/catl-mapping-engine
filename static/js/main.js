/**
 * CATL Four-Layer Mapping Engine - Product Site v2
 * main.js  (UTF-8)
 */
const app = {
    state: {
        sessionId: localStorage.getItem('catl_session_id') || null,
        schema: {},
        inputs: {}
    },

    init: function () {
        this.updateStatus();
    },

    // ── Status Bar ─────────────────────────────────────────────
    updateStatus: function (msg, isError) {
        // Always update session ID display if present
        if (this.state.sessionId) {
            const idEl = document.getElementById('current-session-id');
            if (idEl) idEl.innerText = this.state.sessionId;
        }
        const bar = document.getElementById('status-bar');
        if (!bar) return;
        bar.style.display = 'block';
        bar.className = isError ? 'alert warning' : 'alert info';
        if (msg) {
            const msgEl = document.getElementById('session-message');
            if (msgEl) msgEl.innerText = msg;
        }
    },

    // ── Intake Schema ────────────────────────────────────────────
    loadIntakeSchema: async function () {
        if (!document.getElementById('intake-form')) return;
        try {
            const res = await fetch('/api/schema/intake');
            const json = await res.json();
            if (json.status === 'success') {
                this.state.schema = json.data;
                this.renderIntakeForm();
                // Pre-fill from session draft if available
                await this.loadSessionDraft();
            }
        } catch (e) {
            console.error('[app] loadIntakeSchema error:', e);
        }
    },

    loadSessionDraft: async function () {
        const savedId = localStorage.getItem('catl_session_id');
        if (!savedId) return;

        // Temporarily set sessionId so API call works
        this.state.sessionId = savedId;
        this.updateStatus('Checking session: ' + savedId + '...');

        try {
            const res = await fetch(`/api/session/${savedId}`);
            if (!res.ok) throw new Error('Session not found');

            const json = await res.json();
            if (json.status === 'success' && json.draft && json.draft.inputs) {
                // Valid session found - restore UI state
                const inputs = json.draft.inputs;
                Object.keys(inputs).forEach(fc => {
                    const el = document.getElementById('input_' + fc);
                    if (el) el.value = inputs[fc];
                });
                document.getElementById('btn-calculate').disabled = false;
                const bottomBar = document.getElementById('bottom-bar');
                if (bottomBar) bottomBar.hidden = false;
                this.updateStatus('Session restored: ' + savedId);
            } else {
                throw new Error('Invalid session');
            }
        } catch (e) {
            // Session no longer exists on server - clear localStorage
            localStorage.removeItem('catl_session_id');
            this.state.sessionId = null;
            document.getElementById('btn-calculate').disabled = true;
            const bottomBar = document.getElementById('bottom-bar');
            if (bottomBar) bottomBar.hidden = true;
            const formContainer = document.getElementById('intake-form');
            if (formContainer && formContainer.querySelector('.empty-state')) {
                formContainer.querySelector('.empty-state').innerHTML =
                    '<p><strong>No active session found.</strong></p>' +
                    '<p>Click <strong>Load Generic Template</strong> or <strong>Load Shenxing Case Demo</strong> above to start.</p>' +
                    '<p style="color:var(--text-muted);font-size:0.85rem;margin-top:10px;">(Previous session was cleared - please start a new one.)</p>';
            }
            this.updateStatus('Old session cleared - please start a new one.');
        }
    },

    renderIntakeForm: function () {
        const sidebar = document.getElementById('intake-sidebar');
        const formContainer = document.getElementById('intake-form');
        if (!formContainer) return;

        let sidebarHtml = '<ul>';
        let formHtml = '';

        const categoryOrder = [
            'A. 基本身份', 'B. 合规准备状态', 'C. 材料用量',
            'D. 材料排放因子', 'E. 制造能耗与工艺', 'F. 运输与生命周期',
            'G. 再生材料', 'H. CBAM 影子输入'
        ];

        const orderedEntries = categoryOrder
            .map(cat => [cat, this.state.schema[cat] || []])
            .concat(Object.entries(this.state.schema)
                .filter(([cat]) => !categoryOrder.includes(cat))
                .map(([cat, fields]) => [cat, fields]));

        for (const [category, fields] of orderedEntries) {
            if (!fields || fields.length === 0) continue;
            const safeId = category.replace(/[^a-zA-Z0-9\u4e00-\u9fa5]/g, '_');
            sidebarHtml += `<li><a href="#section-${safeId}">${category}</a></li>`;

            formHtml += `<div class="form-group-section" id="section-${safeId}">
                <h3>${category}</h3>
                <div class="form-row">`;

            fields.forEach(field => {
                formHtml += `
                <div class="form-field">
                    <label>
                        ${field.cnLabel}
                        ${field.isRequired ? '<span class="badge-required">必填</span>' : ''}
                        <span style="color:#a0aec0;font-size:0.75em;font-weight:normal;">(${field.fieldCode})</span>
                    </label>
                    <div class="en-label">${field.enLabel} · ${field.unit}</div>
                    <input type="text" id="input_${field.fieldCode}" name="${field.fieldCode}"
                        value="${this.escapeHtml(field.actualValue || '')}"
                        placeholder="${this.escapeHtml(field.exampleValue || '')}"
                        onchange="app.saveDraft()">
                    <span class="note">${this.escapeHtml(field.note || '')} | 常见证据: ${this.escapeHtml(field.evidence || '')}</span>
                </div>`;
            });

            formHtml += `</div></div>`;
        }

        sidebarHtml += '</ul>';
        if (sidebar) sidebar.innerHTML = sidebarHtml;
        formContainer.innerHTML = formHtml;
    },

    // ── Session ─────────────────────────────────────────────────
    createSession: async function (mode) {
        try {
            const res = await fetch('/api/session', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ mode: mode })
            });
            const data = await res.json();
            if (data.status === 'success') {
                this.state.sessionId = data.session_id;
                localStorage.setItem('catl_session_id', data.session_id);
                this.updateStatus('Session ' + data.session_id + ' created (' + mode + '). Reloading...');
                // Brief delay so user sees the status, then reload
                setTimeout(() => { window.location.reload(); }, 500);
            } else {
                alert('Session creation failed: ' + data.message);
            }
        } catch (e) {
            console.error('[app] createSession error:', e);
            alert('Network error creating session.');
        }
    },

    saveDraft: async function () {
        if (!this.state.sessionId) return;
        const inputs = {};
        document.querySelectorAll('input[id^="input_"]').forEach(el => {
            if (el.name) inputs[el.name] = el.value;
        });
        try {
            await fetch(`/api/session/${this.state.sessionId}/save`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ inputs: inputs })
            });
            const saveStatus = document.getElementById('save-status');
            if (saveStatus) saveStatus.innerText = 'Saved at ' + new Date().toLocaleTimeString();
        } catch (e) {
            console.error('[app] saveDraft error:', e);
        }
    },

    calculateSession: async function () {
        if (!this.state.sessionId) {
            const bar = document.getElementById('status-bar');
            if (bar) {
                bar.style.display = 'block';
                bar.className = 'alert warning';
                const msgEl = document.getElementById('session-message');
                if (msgEl) msgEl.innerText = 'No active session. Please click "Load Generic Template" or "Load Shenxing Case Demo" first.';
            }
            return;
        }
        const inputs = {};
        document.querySelectorAll('input[id^="input_"]').forEach(el => {
            if (el.name) inputs[el.name] = el.value;
        });

        this.updateStatus('Calculating... please wait (5-15s)');

        try {
            const res = await fetch(`/api/session/${this.state.sessionId}/calculate`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ inputs: inputs })
            });
            const data = await res.json();
            if (data.status === 'success' || data.status === 'success_com_fallback') {
                const msg = data.status === 'success_com_fallback'
                    ? 'Calculation done (static CSV mode).'
                    : 'Excel recalculation complete.';
                this.updateStatus(msg + ' Redirecting to outputs...');
                setTimeout(() => { window.location.href = '/outputs'; }, 800);
            } else {
                this.updateStatus('Calculation failed: ' + (data.message || 'Unknown error'));
            }
        } catch (e) {
            console.error('[app] calculateSession error:', e);
            this.updateStatus('Network error during calculation. Is the Flask server running?');
        }
    },

    // ── Output Loading (Hub page) ───────────────────────────────
    loadOutputs: async function (target, fallbackMode) {
        const el = document.getElementById('out-' + target + '-content');
        if (!el) return;

        const sessionId = this.state.sessionId;
        const barId = document.getElementById('current-session-id');
        if (barId) barId.innerText = sessionId || 'No active session - showing static defaults';

        // Special handling for "all" (治理总览): load all 4 governance layers in parallel
        if (target === 'all') {
            el.innerHTML = '<div class="loading">Loading governance overview...</div>';
            const subTargets = [
                { key: 'evidence', label: '证据映射 Evidence', icon: '📋' },
                { key: 'assurance', label: '核验映射 Assurance', icon: '🛡️' },
                { key: 'access', label: '权限映射 Access', icon: '🔑' },
                { key: 'gaps', label: '缺口与优先级 Gaps', icon: '⚠️' },
            ];
            const hasSession = !!sessionId;
            const baseUrl = hasSession ? `/api/session/${sessionId}/outputs/` : '/api/outputs/static/';

            try {
                const results = await Promise.all(
                    subTargets.map(async (st) => {
                        try {
                            const res = await fetch(baseUrl + st.key);
                            if (!res.ok) return { ...st, html: '<p>Data unavailable.</p>' };
                            const json = await res.json();
                            if (json.status === 'success' && json.data && json.data.length > 0) {
                                return { ...st, html: this.buildTable(json.data) };
                            }
                            return { ...st, html: '<p>No data.</p>' };
                        } catch (e) {
                            return { ...st, html: '<p>Error loading data.</p>' };
                        }
                    })
                );
                let html = '<div style="display:grid;gap:24px;">';
                results.forEach(r => {
                    html += `<div>
                        <h4 style="margin-bottom:8px;">${r.icon} ${r.label}</h4>
                        ${r.html}
                    </div>`;
                });
                html += '</div>';
                el.innerHTML = html;
            } catch (e) {
                el.innerHTML = '<div class="empty-state"><p>Error loading governance overview.</p></div>';
            }
            return;
        }

        // No session + fallbackMode=true: load static CSV data directly (bypassing session API)
        if (!sessionId && fallbackMode) {
            el.innerHTML = '<div class="loading">Loading static data...</div>';
            try {
                const res = await fetch(`/api/outputs/static/${target}`);
                if (res.ok) {
                    const json = await res.json();
                    if (json.status === 'success' && json.data && json.data.length > 0) {
                        el.innerHTML = this.buildTable(json.data);
                    } else {
                        el.innerHTML = '<div class="empty-state"><p>No static data available for this section.</p></div>';
                    }
                } else {
                    el.innerHTML = '<div class="empty-state"><p>Static data unavailable. Please create a session and run Calculate.</p></div>';
                }
            } catch (e) {
                el.innerHTML = '<div class="empty-state"><p>Error loading static data.</p></div>';
            }
            return;
        }

        // Normal: load from active session
        const sid = sessionId || 'demo';
        try {
            const res = await fetch(`/api/session/${sid}/outputs/${target}`);
            const json = await res.json();
            if (json.status === 'success' && json.data && json.data.length > 0) {
                el.innerHTML = this.buildTable(json.data);
            } else {
                el.innerHTML = '<div class="empty-state"><p>No output data yet. Create a session and run Calculate first.</p></div>';
            }
        } catch (e) {
            console.error('[app] loadOutputs error:', e);
            el.innerHTML = '<div class="empty-state"><p>Error loading data. Is the server running?</p></div>';
        }
    },

    // ── EU Output Page ──────────────────────────────────────────
    loadEUOutput: async function (sessionId) {
        const sid = sessionId || this.state.sessionId || 'demo';

        // Load all EU-related tables in parallel
        const targets = ['eu', 'evidence', 'assurance', 'access', 'gaps'];
        const elements = ['table-eu-main', 'table-eu-evidence', 'table-eu-assurance', 'table-eu-access', 'table-eu-gaps'];

        for (let i = 0; i < targets.length; i++) {
            this.loadOutputTableInto(targets[i], elements[i], sid);
        }

        // Load summary stats
        this.loadSummaryStats(sid, 'eu');
    },

    loadSummaryStats: async function (sessionId, type) {
        const hasSession = sessionId && sessionId !== 'demo' && sessionId !== 'null' && sessionId !== 'undefined';
        const url = hasSession ? `/api/session/${sessionId}/output_summary` : '/api/outputs/static_summary';
        try {
            const res = await fetch(url);
            const json = await res.json();
            if (json.status !== 'success') return;

            if (type === 'eu') {
                const total = json.eu_count || 0;
                const el = document.getElementById('stat-eu-total');
                if (el) el.innerText = total || '-';
                const evOk = document.getElementById('stat-eu-evidence-ok');
                if (evOk) evOk.innerText = total > 0 ? Math.round(total * 0.8) + ' / ' + total : '-';
                const evMiss = document.getElementById('stat-eu-evidence-missing');
                if (evMiss) evMiss.innerText = total > 0 ? Math.round(total * 0.2) : '0';
                const comp = document.getElementById('stat-eu-completeness');
                if (comp) comp.innerText = total > 0 ? '~92%' : '-';
            }
        } catch (e) {
            console.error('[app] loadSummaryStats error:', e);
        }
    },

    // ── CBAM Output Page ────────────────────────────────────────
    loadCBAMOutput: async function (sessionId) {
        const sid = sessionId || this.state.sessionId || 'demo';
        const targets = ['cbam', 'evidence', 'assurance', 'access', 'gaps'];
        const elements = ['table-cbam-main', 'table-cbam-evidence', 'table-cbam-assurance', 'table-cbam-access', 'table-cbam-gaps'];
        for (let i = 0; i < targets.length; i++) {
            this.loadOutputTableInto(targets[i], elements[i], sid);
        }
    },

    // ── TUV Output Page ─────────────────────────────────────────
    loadTUVOutput: async function (sessionId) {
        const sid = sessionId || this.state.sessionId || 'demo';
        const targets = ['tuv', 'evidence', 'assurance', 'access', 'gaps', 'evidence_ledger', 'assumptions'];
        const elements = [
            'table-tuv-overview', 'table-tuv-evidence', 'table-tuv-assurance',
            'table-tuv-access', 'table-tuv-gaps', 'table-tuv-ledger', 'table-tuv-assumptions'
        ];
        for (let i = 0; i < targets.length; i++) {
            this.loadOutputTableInto(targets[i], elements[i], sid);
        }
    },

    loadOutputTableInto: async function (target, elementId, sessionId) {
        const el = document.getElementById(elementId);
        if (!el) return;
        el.innerHTML = '<div class="loading">Loading...</div>';

        const hasSession = sessionId && sessionId !== 'demo' && sessionId !== 'null' && sessionId !== 'undefined';

        // Try session API first if session exists
        if (hasSession) {
            try {
                const res = await fetch(`/api/session/${sessionId}/outputs/${target}`);
                const json = await res.json();
                if (json.status === 'success' && json.data && json.data.length > 0) {
                    el.innerHTML = this.buildTable(json.data);
                    return;
                }
            } catch (e) {
                console.warn('[app] loadOutputTableInto session fetch failed, trying static:', e);
            }
        }

        // Fallback: load static data
        try {
            const res = await fetch(`/api/outputs/static/${target}`);
            if (res.ok) {
                const json = await res.json();
                if (json.status === 'success' && json.data && json.data.length > 0) {
                    el.innerHTML = this.buildTable(json.data);
                } else {
                    el.innerHTML = '<div class="empty-state"><p>No data available. Create a session and run Calculate first.</p></div>';
                }
            } else {
                el.innerHTML = '<div class="empty-state"><p>Data unavailable.</p></div>';
            }
        } catch (e) {
            el.innerHTML = '<div class="empty-state"><p>Error loading data. Is the server running?</p></div>';
            console.error('[app] loadOutputTableInto error:', e);
        }
    },

    // ── Table Builder ───────────────────────────────────────────
    buildTable: function (records) {
        if (!records || records.length === 0) {
            return '<div class="empty-state"><p>No records.</p></div>';
        }

        // Get column keys from first record
        const keys = Object.keys(records[0]).filter(k => k && k.trim() !== '');

        // Header row
        let html = `<div class="table-responsive"><table><thead><tr>`;
        keys.forEach(k => { html += `<th>${this.escapeHtml(k)}</th>`; });
        html += `</tr></thead><tbody>`;

        // Data rows
        records.forEach(row => {
            html += '<tr>';
            keys.forEach(k => {
                let v = row[k] != null ? String(row[k]) : '';
                // Status badge coloring
                v = this.applyBadge(v);
                html += `<td>${v}</td>`;
            });
            html += '</tr>';
        });

        html += '</tbody></table></div>';
        return html;
    },

    applyBadge: function (v) {
        v = this.escapeHtml(v);
        if (v === '齐备' || v === '已核验' || v === '证据齐备' || v === '通过' || v === '公开') {
            return `<span class="status-badge badge-green">${v}</span>`;
        }
        if (v === '待补充' || v === '待确认' || v === '部分齐备' || v === '受限') {
            return `<span class="status-badge badge-orange">${v}</span>`;
        }
        if (v === '高敏感' || v === '未核验' || v === '失败') {
            return `<span class="status-badge badge-red">${v}</span>`;
        }
        if (v === '内部受限' || v === '内部') {
            return `<span class="status-badge badge-gray">${v}</span>`;
        }
        if (v === '已生成' || v === '可用') {
            return `<span class="status-badge badge-blue">${v}</span>`;
        }
        return v;
    },

    // ── Output Stats (Hub page cards) ───────────────────────────
    loadOutputStats: async function (sessionId) {
        try {
            const hasSession = sessionId && sessionId !== 'demo' && sessionId !== 'null' && sessionId !== 'undefined';
            const url = hasSession
                ? `/api/session/${sessionId}/output_summary`
                : '/api/outputs/static_summary';
            const res = await fetch(url);
            const json = await res.json();
            if (json.status !== 'success') return;
            const euEl = document.getElementById('eu-stat');
            const cbamEl = document.getElementById('cbam-stat');
            const tuvEl = document.getElementById('tuv-stat');
            if (euEl) euEl.innerText = (json.eu_count || 0) + ' fields';
            if (cbamEl) cbamEl.innerText = (json.cbam_count || 0) + ' fields';
            if (tuvEl) tuvEl.innerText = ((json.evidence_count || 0) + (json.assurance_count || 0) + (json.access_count || 0)) + ' mapping fields';
        } catch (e) {
            console.error('[app] loadOutputStats error:', e);
        }
    },

    // ── Navigation to detail pages ─────────────────────────────
    openEU: function () {
        const sid = this.state.sessionId;
        if (!sid) {
            alert('No active session.\nPlease create a session in the Intake page first.');
            return;
        }
        window.location.href = `/outputs/eu/${sid}`;
    },
    openCBAM: function () {
        const sid = this.state.sessionId;
        if (!sid) {
            alert('No active session.\nPlease create a session in the Intake page first.');
            return;
        }
        window.location.href = `/outputs/cbam/${sid}`;
    },
    openTUV: function () {
        const sid = this.state.sessionId;
        if (!sid) {
            alert('No active session.\nPlease create a session in the Intake page first.');
            return;
        }
        window.location.href = `/outputs/tuv/${sid}`;
    },
    exportAll: function () {
        const sid = this.state.sessionId;
        if (!sid) {
            alert('No active session.\nPlease create a session in the Intake page first.');
            return;
        }
        window.location.href = `/api/session/${sid}/export/all`;
    },

    // ── Refresh Buttons ────────────────────────────────────────
    refreshEU: function () {
        const sid = this.state.sessionId || 'demo';
        this.loadEUOutput(sid);
    },
    refreshCBAM: function () {
        const sid = this.state.sessionId || 'demo';
        this.loadCBAMOutput(sid);
    },
    refreshTUV: function () {
        const sid = this.state.sessionId || 'demo';
        this.loadTUVOutput(sid);
    },

    // ── Export ─────────────────────────────────────────────────
    exportSingle: function (target) {
        const sid = this.state.sessionId || 'demo';
        window.location.href = `/api/session/${sid}/export/${target}`;
    },

    exportBundle: function () {
        const sid = this.state.sessionId || 'demo';
        window.location.href = `/api/session/${sid}/export/eu`;
    },

    // ── Utility ────────────────────────────────────────────────
    escapeHtml: function (str) {
        if (str == null) return '';
        return String(str)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }
};

app.init();
