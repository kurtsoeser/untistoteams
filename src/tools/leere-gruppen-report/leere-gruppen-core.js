(function () {
    'use strict';

    function compareDe(a, b) {
        return String(a || '').localeCompare(String(b || ''), 'de', { sensitivity: 'base' });
    }

    function escapeHtml(s) {
        return String(s ?? '')
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
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

    function sleep(ms) {
        return new Promise((r) => setTimeout(r, ms));
    }

    async function getGraphToken(scopes) {
        if (typeof window.ms365AuthAcquireToken === 'function') {
            return await window.ms365AuthAcquireToken(scopes);
        }
        throw new Error('Bitte oben rechts anmelden (MSAL-Widget nicht verfügbar).');
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
                typeof data === 'object' && data && data.error
                    ? JSON.stringify(data.error)
                    : text || String(res.status);
            const err = new Error(method + ' ' + pathOrUrl + ': ' + msg);
            err.status = res.status;
            throw err;
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
        const n = Math.max(1, Math.min(concurrency, tasks.length || 1));
        const workers = [];
        for (let w = 0; w < n; w++) workers.push(worker());
        await Promise.all(workers);
        return results;
    }

    function isUnified(g) {
        const gt = g && Array.isArray(g.groupTypes) ? g.groupTypes : [];
        return gt.indexOf('Unified') !== -1;
    }

    function isTeam(g) {
        const ro = g && Array.isArray(g.resourceProvisioningOptions) ? g.resourceProvisioningOptions : [];
        return ro.indexOf('Team') !== -1;
    }

    function isSecurity(g) {
        return !!(g && g.securityEnabled && !g.mailEnabled && !isUnified(g));
    }

    function isMailGroup(g) {
        return !!(g && g.mailEnabled && !isUnified(g));
    }

    function groupKindLabel(g) {
        const parts = [];
        if (isUnified(g)) parts.push('M365');
        if (isTeam(g)) parts.push('Team');
        if (g && g.securityEnabled && !g.mailEnabled) parts.push('Sicherheit');
        if (g && g.mailEnabled && !isUnified(g)) parts.push('Mail');
        return parts.length ? parts.join(' · ') : 'Sonstige';
    }

    function kindBadgesHtml(row) {
        const out = [];
        if (row.isUnified) out.push('<span class="lgr-badge is-m365">M365</span>');
        if (row.isTeam) out.push('<span class="lgr-badge is-team">Team</span>');
        if (row.isSecurity) out.push('<span class="lgr-badge is-security">Sicherheit</span>');
        if (row.isMail) out.push('<span class="lgr-badge is-mail">Mail</span>');
        if (!out.length) out.push('<span class="lgr-badge">Sonstige</span>');
        return out.join('');
    }

    /**
     * Konvertiert Graph-Group-Objekte (mit zusätzlich gezählten owners/members) zu Zeilen-
     * Datensätzen, wie sie die Tabelle benötigt.
     * @param {object} g - Microsoft-Graph-Group-Objekt (id, displayName, mail, groupTypes, …)
     * @param {number} owners - Anzahl Besitzer (negativ = Zählung fehlgeschlagen)
     * @param {number} members - Anzahl Mitglieder (negativ = Zählung fehlgeschlagen)
     */
    function buildRow(g, owners, members) {
        const flags = [];
        if (owners === 0) flags.push('ohne Besitzer');
        if (members === 0) flags.push('ohne Mitglieder');
        if (owners < 0 || members < 0) flags.push('Zählen fehlgeschlagen');
        return {
            id: String(g.id || ''),
            displayName: String(g.displayName || ''),
            mail: String(g.mail || ''),
            kind: groupKindLabel(g),
            isUnified: isUnified(g),
            isTeam: isTeam(g),
            isSecurity: isSecurity(g),
            isMail: isMailGroup(g),
            owners,
            members,
            flags: flags.join(', ') || '–'
        };
    }

    function buildGroupsListInitialPath(scopeMode) {
        const select =
            'id,displayName,mail,mailNickname,groupTypes,resourceProvisioningOptions,visibility,securityEnabled,mailEnabled';
        if (scopeMode === 'all') {
            return {
                path: '/groups?$select=' + encodeURIComponent(select) + '&$top=999',
                headers: {}
            };
        }
        if (scopeMode === 'team') {
            return {
                path:
                    '/groups?$filter=' +
                    encodeURIComponent("resourceProvisioningOptions/Any(x:x eq 'Team')") +
                    '&$select=' +
                    encodeURIComponent(select) +
                    '&$count=true&$top=999',
                headers: { ConsistencyLevel: 'eventual' }
            };
        }
        return {
            path:
                '/groups?$filter=' +
                encodeURIComponent("groupTypes/any(c:c eq 'Unified')") +
                '&$select=' +
                encodeURIComponent(select) +
                '&$count=true&$top=999',
            headers: { ConsistencyLevel: 'eventual' }
        };
    }

    window.ms365LeereGruppenCore = {
        compareDe,
        escapeHtml,
        csvEscape,
        rowsToCsv,
        sleep,
        getGraphToken,
        graphRequest,
        graphJson,
        fetchAllPages,
        fetchCount,
        runPool,
        isUnified,
        isTeam,
        isSecurity,
        isMailGroup,
        groupKindLabel,
        kindBadgesHtml,
        buildRow,
        buildGroupsListInitialPath
    };
})();
