(function () {
    'use strict';

    const GRAPH_SCOPES = [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/Group.Read.All'
    ];

    function getEl(id) {
        return document.getElementById(id);
    }

    function compareDe(a, b) {
        return String(a || '').localeCompare(String(b || ''), 'de', { sensitivity: 'base' });
    }

    async function getGraphToken(scopes) {
        if (typeof window.ms365AuthAcquireToken === 'function') {
            return await window.ms365AuthAcquireToken(scopes);
        }
        throw new Error('Bitte oben rechts anmelden (MSAL-Widget nicht verfügbar).');
    }

    function sleep(ms) {
        return new Promise((r) => setTimeout(r, ms));
    }

    async function graphRequest(method, pathOrUrl, token, body, extraHeaders) {
        const url = pathOrUrl.indexOf('http') === 0 ? pathOrUrl : 'https://graph.microsoft.com/v1.0' + pathOrUrl;
        let attempt = 0;
        while (true) {
            const headers = { Authorization: 'Bearer ' + token };
            if (extraHeaders && typeof extraHeaders === 'object') Object.assign(headers, extraHeaders);
            let payload = undefined;
            if (body !== undefined) {
                headers['Content-Type'] = 'application/json';
                payload = JSON.stringify(body);
            }
            const res = await fetch(url, { method, headers, body: payload });
            if (res.status === 429 && attempt < 8) {
                const ra = parseInt(res.headers.get('Retry-After') || '5', 10);
                await sleep((isNaN(ra) ? 5 : ra) * 1000);
                attempt++;
                continue;
            }
            return res;
        }
    }

    async function graphJson(method, pathOrUrl, token, body, extraHeaders) {
        const res = await graphRequest(method, pathOrUrl, token, body, extraHeaders);
        const text = await res.text();
        let data = null;
        if (text) {
            try {
                data = JSON.parse(text);
            } catch {
                data = text;
            }
        }
        if (!res.ok) {
            const msg =
                typeof data === 'object' && data && data.error ? JSON.stringify(data.error) : text || String(res.status);
            throw new Error(method + ' ' + pathOrUrl + ': ' + msg);
        }
        return data || {};
    }

    async function fetchAllPages(token, initialPath, onProgress, extraHeaders) {
        const out = [];
        let next = initialPath;
        let page = 0;
        while (next) {
            page++;
            const data = await graphJson('GET', next, token, undefined, extraHeaders);
            const vals = data.value;
            if (Array.isArray(vals)) for (let i = 0; i < vals.length; i++) out.push(vals[i]);
            next = data['@odata.nextLink'] || null;
            if (typeof onProgress === 'function') onProgress({ page, loaded: out.length, hasMore: !!next });
        }
        return out;
    }

    async function fetchCount(token, groupId, segment) {
        const path = '/groups/' + encodeURIComponent(groupId) + '/' + segment + '/$count';
        const res = await graphRequest('GET', path, token, undefined, { ConsistencyLevel: 'eventual' });
        const text = await res.text();
        if (!res.ok) return -1;
        const n = parseInt(String(text).trim(), 10);
        return isNaN(n) ? -1 : n;
    }

    function isUnified(g) {
        const gt = g && Array.isArray(g.groupTypes) ? g.groupTypes : [];
        return gt.indexOf('Unified') !== -1;
    }

    function isTeam(g) {
        const ro = g && Array.isArray(g.resourceProvisioningOptions) ? g.resourceProvisioningOptions : [];
        return ro.indexOf('Team') !== -1;
    }

    function groupKindLabel(g) {
        const parts = [];
        if (isUnified(g)) parts.push('M365');
        if (isTeam(g)) parts.push('Team');
        if (g && g.securityEnabled && !g.mailEnabled) parts.push('Sicherheit');
        if (g && g.mailEnabled && !isUnified(g)) parts.push('Mail');
        return parts.length ? parts.join(' · ') : 'Sonstige';
    }

    function csvEscape(cell) {
        const s = String(cell ?? '');
        if (/[",\n\r]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
        return s;
    }

    function rowsToCsv(rows, columns) {
        const header = columns.map((c) => c.label).join(';');
        const lines = rows.map((row) => columns.map((c) => csvEscape(c.value(row))).join(';'));
        return '\uFEFF' + header + '\n' + lines.join('\n');
    }

    async function runPool(tasks, concurrency) {
        const results = new Array(tasks.length);
        let i = 0;
        async function worker() {
            while (true) {
                const idx = i++;
                if (idx >= tasks.length) return;
                results[idx] = await tasks[idx]();
            }
        }
        const n = Math.max(1, Math.min(concurrency, tasks.length));
        const workers = [];
        for (let w = 0; w < n; w++) workers.push(worker());
        await Promise.all(workers);
        return results;
    }

    function bind() {
        const progressEl = getEl('lgrProgress');
        const tbody = getEl('lgrTbody');
        const btn = getEl('lgrBtnRun');
        const btnCsv = getEl('lgrBtnCsv');
        const filterProblem = getEl('lgrFilterProblem');

        /** @type {Array<{id:string,displayName:string,mail:string,kind:string,owners:number,members:number,isTeam:boolean}>} */
        let lastRows = [];

        function setProgress(on, text) {
            if (!progressEl) return;
            progressEl.style.display = on ? '' : 'none';
            if (text) progressEl.textContent = String(text);
        }

        function applyTableFilter() {
            if (!tbody) return;
            const onlyProb = filterProblem && filterProblem.checked;
            const frag = document.createDocumentFragment();
            for (const r of lastRows) {
                if (onlyProb && r.owners > 0 && r.members > 0) continue;
                const tr = document.createElement('tr');
                tr.innerHTML =
                    '<td>' +
                    escapeHtml(r.displayName) +
                    '</td><td>' +
                    escapeHtml(r.mail) +
                    '</td><td>' +
                    escapeHtml(r.kind) +
                    '</td><td style="text-align:right">' +
                    (r.owners < 0 ? '–' : r.owners) +
                    '</td><td style="text-align:right">' +
                    (r.members < 0 ? '–' : r.members) +
                    '</td><td>' +
                    escapeHtml(r.flags) +
                    '</td>';
                frag.appendChild(tr);
            }
            tbody.replaceChildren(frag);
        }

        function escapeHtml(s) {
            return String(s ?? '')
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;')
                .replace(/"/g, '&quot;');
        }

        btn?.addEventListener('click', async () => {
            const scopeMode = String(getEl('lgrScope')?.value || 'unified');
            btn.disabled = true;
            btnCsv.disabled = true;
            lastRows = [];
            if (tbody) tbody.replaceChildren();
            setProgress(true, 'Gruppen werden geladen …');

            try {
                const token = await getGraphToken(GRAPH_SCOPES);
                const select =
                    'id,displayName,mail,mailNickname,groupTypes,resourceProvisioningOptions,visibility,securityEnabled,mailEnabled';
                let initial;
                const headers = {};

                if (scopeMode === 'all') {
                    initial = '/groups?$select=' + encodeURIComponent(select) + '&$top=999';
                } else if (scopeMode === 'team') {
                    initial =
                        '/groups?$filter=' +
                        encodeURIComponent("resourceProvisioningOptions/Any(x:x eq 'Team')") +
                        '&$select=' +
                        encodeURIComponent(select) +
                        '&$count=true&$top=999';
                    headers.ConsistencyLevel = 'eventual';
                } else {
                    initial =
                        '/groups?$filter=' +
                        encodeURIComponent("groupTypes/any(c:c eq 'Unified')") +
                        '&$select=' +
                        encodeURIComponent(select) +
                        '&$count=true&$top=999';
                    headers.ConsistencyLevel = 'eventual';
                }

                const groups = await fetchAllPages(token, initial, (p) => {
                    setProgress(true, 'Gruppen laden … Seite ' + p.page + ', ' + p.loaded + ' Gruppen');
                }, headers);

                const tasks = groups.map((g) => async () => {
                    const id = String(g.id || '');
                    if (!id) return null;
                    const [owners, members] = await Promise.all([
                        fetchCount(token, id, 'owners'),
                        fetchCount(token, id, 'members')
                    ]);
                    const flags = [];
                    if (owners === 0) flags.push('ohne Besitzer');
                    if (members === 0) flags.push('ohne Mitglieder');
                    if (owners < 0 || members < 0) flags.push('Zählen fehlgeschlagen');
                    return {
                        id,
                        displayName: String(g.displayName || ''),
                        mail: String(g.mail || ''),
                        kind: groupKindLabel(g),
                        owners,
                        members,
                        isTeam: isTeam(g),
                        flags: flags.join(', ') || '–'
                    };
                });

                setProgress(true, 'Besitzer/Mitglieder zählen … 0 / ' + tasks.length);
                const concurrency = 4;
                let done = 0;
                const enriched = await runPool(
                    tasks.map((fn) => async () => {
                        const row = await fn();
                        done++;
                        if (done % 5 === 0 || done === tasks.length) {
                            setProgress(true, 'Besitzer/Mitglieder zählen … ' + done + ' / ' + tasks.length);
                        }
                        return row;
                    }),
                    concurrency
                );

                lastRows = enriched.filter(Boolean);
                lastRows.sort((a, b) => compareDe(a.displayName, b.displayName));
                setProgress(true, 'Fertig: ' + lastRows.length + ' Gruppen ausgewertet.');
                applyTableFilter();
                btnCsv.disabled = !lastRows.length;
                setTimeout(() => setProgress(false, ''), 2400);
            } catch (e) {
                setProgress(true, 'Fehler: ' + (e && e.message ? e.message : String(e)));
            } finally {
                btn.disabled = false;
            }
        });

        filterProblem?.addEventListener('change', () => applyTableFilter());

        btnCsv?.addEventListener('click', () => {
            if (!lastRows.length) return;
            const cols = [
                { label: 'Anzeigename', value: (r) => r.displayName },
                { label: 'E-Mail', value: (r) => r.mail },
                { label: 'Typ', value: (r) => r.kind },
                { label: 'Besitzer', value: (r) => r.owners },
                { label: 'Mitglieder', value: (r) => r.members },
                { label: 'Hinweise', value: (r) => r.flags }
            ];
            const csv = rowsToCsv(lastRows, cols);
            const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = 'leere-gruppen-report.csv';
            a.click();
            setTimeout(() => URL.revokeObjectURL(a.href), 4000);
        });
    }

    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', bind);
    else bind();
})();
