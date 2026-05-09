(function () {
    'use strict';

    const GRAPH_SCOPES = [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/Group.Read.All'
    ];

    /**
     * GET /invitations (Sammlung) gibt es in Graph v1.0 nicht → 404. Stattdessen Verzeichnis-Audit
     * (Microsoft-Doku: u. a. „Invite external user“).
     */
    const B2B_AUDIT_SCOPES = [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/Group.Read.All',
        'https://graph.microsoft.com/AuditLog.Read.All'
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

    async function fetchGuestUsersForGroup(token, groupId) {
        const select = 'id,displayName,userPrincipalName,mail,userType';
        let next =
            '/groups/' +
            encodeURIComponent(groupId) +
            '/members/microsoft.graph.user?$select=' +
            encodeURIComponent(select) +
            '&$top=999';
        const guests = [];
        let pages = 0;
        while (next && pages < 80) {
            pages++;
            const data = await graphJson('GET', next, token, undefined, undefined);
            const vals = data.value || [];
            for (let i = 0; i < vals.length; i++) {
                const u = vals[i];
                if (String(u.userType || '').toLowerCase() === 'guest') guests.push(u);
            }
            next = data['@odata.nextLink'] || null;
        }
        guests.sort((a, b) => compareDe(a.displayName || a.userPrincipalName, b.displayName || b.userPrincipalName));
        return guests;
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

    function csvEscape(cell) {
        const s = String(cell ?? '');
        if (/[",\n\r]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
        return s;
    }

    function escapeHtml(s) {
        return String(s ?? '')
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    /** Tab-Sitzung: übersteht Seitenreload/Redirect, kein localStorage (geringer DSGVO-Footprint). */
    const GZ_STORAGE_KEY = 'ms365-gast-zugaenge-snapshot-v1';

    function gzLoadSnapshot() {
        try {
            const raw = sessionStorage.getItem(GZ_STORAGE_KEY);
            if (!raw) return null;
            const o = JSON.parse(raw);
            if (!o || typeof o !== 'object') return null;
            if (!Array.isArray(o.teams) || !Array.isArray(o.invitations)) return null;
            return o;
        } catch {
            return null;
        }
    }

    function gzSaveSnapshot(teams, invitations, invNoteText) {
        try {
            sessionStorage.setItem(
                GZ_STORAGE_KEY,
                JSON.stringify({
                    savedAt: new Date().toISOString(),
                    teams,
                    invitations,
                    invNote: String(invNoteText || '')
                })
            );
        } catch {
            /* Quota, privates Fenster */
        }
    }

    function gzExtractGuestIdentifierFromAudit(a) {
        const tr = Array.isArray(a.targetResources) ? a.targetResources : [];
        for (let i = 0; i < tr.length; i++) {
            const upn = String(tr[i].userPrincipalName || '').trim();
            if (upn) return upn;
        }
        for (let i = 0; i < tr.length; i++) {
            const dn = String(tr[i].displayName || '').trim();
            if (dn.indexOf('@') !== -1) return dn;
        }
        const details = Array.isArray(a.additionalDetails) ? a.additionalDetails : [];
        for (let j = 0; j < details.length; j++) {
            const k = String(details[j].key || '').toLowerCase();
            if (k.indexOf('mail') !== -1 || k.indexOf('upn') !== -1) {
                const v = String(details[j].value || '').trim();
                if (v) return v;
            }
        }
        for (let i = 0; i < tr.length; i++) {
            const mods = tr[i].modifiedProperties;
            if (!Array.isArray(mods)) continue;
            for (let m = 0; m < mods.length; m++) {
                const name = String(mods[m].displayName || '').toLowerCase();
                if (name.includes('mail') || name.includes('userprincipalname')) {
                    const nv = String(mods[m].newValue || '')
                        .trim()
                        .replace(/^"+|"+$/g, '');
                    if (nv) return nv;
                }
            }
        }
        return '';
    }

    function gzExtractGuestDisplayNameFromAudit(a) {
        const tr = Array.isArray(a.targetResources) ? a.targetResources : [];
        for (let i = 0; i < tr.length; i++) {
            const t = tr[i];
            if (String(t.type || '').toLowerCase() === 'user') {
                const dn = String(t.displayName || '').trim();
                if (dn) return dn;
            }
        }
        for (let i = 0; i < tr.length; i++) {
            const dn = String(tr[i].displayName || '').trim();
            if (dn && dn.indexOf('@') === -1) return dn;
        }
        return '';
    }

    /** Mappt directoryAudit → gleiche Tabellenfelder wie zuvor bei /invitations. */
    function gzMapDirectoryAuditsToB2bRows(entries) {
        if (!Array.isArray(entries)) return [];
        const out = [];
        for (let i = 0; i < entries.length; i++) {
            const e = entries[i];
            const id = gzExtractGuestIdentifierFromAudit(e);
            const guestName = gzExtractGuestDisplayNameFromAudit(e);
            const initiator =
                e.initiatedBy && e.initiatedBy.user
                    ? String(e.initiatedBy.user.displayName || e.initiatedBy.user.userPrincipalName || '').trim()
                    : '';
            const when = e.activityDateTime ? String(e.activityDateTime).replace('T', ' ').slice(0, 19) + 'Z' : '';
            const parts = [];
            if (when) parts.push(when);
            if (e.result) parts.push(String(e.result));
            if (e.resultReason && String(e.resultReason) !== String(e.result || '')) parts.push(String(e.resultReason));
            out.push({
                invitedUserEmailAddress: id || '–',
                status: parts.join(' · '),
                invitedUserDisplayName: guestName || initiator || String(e.activityDisplayName || '')
            });
        }
        return out;
    }

    function gzBuildB2bAuditInitialPath() {
        const filter =
            "category eq 'UserManagement' and " +
            "(activityDisplayName eq 'Invite external user' " +
            "or activityDisplayName eq 'Invite external user with reset invitation status' " +
            "or activityDisplayName eq 'Invite internal user to B2B collaboration')";
        const select =
            'activityDateTime,activityDisplayName,result,resultReason,initiatedBy,targetResources,additionalDetails';
        return (
            '/auditLogs/directoryAudits?$filter=' +
            encodeURIComponent(filter) +
            '&$orderby=' +
            encodeURIComponent('activityDateTime desc') +
            '&$select=' +
            encodeURIComponent(select) +
            '&$top=999'
        );
    }

    function bind() {
        const progressEl = getEl('gzProgress');
        const tbodyTeams = getEl('gzTbodyTeams');
        const tbodyInv = getEl('gzTbodyInv');
        const btn = getEl('gzBtnRun');
        const btnCsvTeams = getEl('gzBtnCsvTeams');
        const btnCsvInv = getEl('gzBtnCsvInv');
        const invNote = getEl('gzInvNote');

        let lastTeams = [];
        let lastInv = [];

        function setProgress(on, text) {
            if (!progressEl) return;
            progressEl.style.display = on ? '' : 'none';
            if (text) progressEl.textContent = String(text);
        }

        function renderTeamsLoadingRow(message) {
            if (!tbodyTeams) return;
            const tr = document.createElement('tr');
            tr.innerHTML =
                '<td colspan="4" class="muted" style="padding:14px 10px;line-height:1.45;">' +
                escapeHtml(message) +
                '</td>';
            tbodyTeams.replaceChildren(tr);
        }

        function renderTeams() {
            if (!tbodyTeams) return;
            const frag = document.createDocumentFragment();
            for (const r of lastTeams) {
                const tr = document.createElement('tr');
                const guestLines = r.guests
                    .map((g) => String(g.displayName || '').trim() + ' <' + String(g.userPrincipalName || g.mail || '').trim() + '>')
                    .join('\n');
                const shortList =
                    r.guests.length <= 4
                        ? guestLines
                        : r.guests
                              .slice(0, 4)
                              .map((g) => String(g.displayName || '').trim() + ' <' + String(g.userPrincipalName || g.mail || '').trim() + '>')
                              .join('\n') +
                          '\n… +' +
                          (r.guests.length - 4) +
                          ' weitere';
                tr.innerHTML =
                    '<td>' +
                    escapeHtml(r.displayName) +
                    '</td><td>' +
                    escapeHtml(r.mail) +
                    '</td><td style="text-align:right">' +
                    r.guests.length +
                    '</td><td><pre style="margin:0;font-family:inherit;font-size:0.88em;white-space:pre-wrap;max-width:420px;">' +
                    escapeHtml(shortList) +
                    '</pre></td>';
                if (guestLines) tr.title = guestLines;
                frag.appendChild(tr);
            }
            tbodyTeams.replaceChildren(frag);
        }

        function renderInv() {
            if (!tbodyInv) return;
            const frag = document.createDocumentFragment();
            for (const inv of lastInv) {
                const tr = document.createElement('tr');
                tr.innerHTML =
                    '<td>' +
                    escapeHtml(inv.invitedUserEmailAddress || '') +
                    '</td><td>' +
                    escapeHtml(inv.status || '') +
                    '</td><td>' +
                    escapeHtml(inv.invitedUserDisplayName || '') +
                    '</td>';
                frag.appendChild(tr);
            }
            tbodyInv.replaceChildren(frag);
        }

        if (btn && btn.dataset.gzBound === '1') return;
        if (btn) btn.dataset.gzBound = '1';

        (function gzRestoreFromSession() {
            const snap = gzLoadSnapshot();
            if (!snap) return;
            lastTeams = snap.teams;
            lastInv = snap.invitations;
            let when = '';
            try {
                when = snap.savedAt ? new Date(snap.savedAt).toLocaleString('de-DE') : '';
            } catch {
                when = '';
            }
            renderTeams();
            if (lastInv.length) {
                renderInv();
            } else if (tbodyInv) {
                tbodyInv.replaceChildren();
            }
            btnCsvTeams.disabled = !lastTeams.length;
            btnCsvInv.disabled = !lastInv.length;
            if (invNote) {
                const hint =
                    (when ? 'Zwischengespeichert in dieser Browser-Sitzung (Stand ' + when + '). ' : '') +
                    (snap.invNote ? String(snap.invNote) : '');
                invNote.textContent = hint.trim() || '';
            }
        })();

        btn?.addEventListener('click', async () => {
            btn.disabled = true;
            btnCsvTeams.disabled = true;
            btnCsvInv.disabled = true;
            lastTeams = [];
            lastInv = [];
            renderTeamsLoadingRow('Teamliste wird geladen …');
            if (tbodyInv) tbodyInv.replaceChildren();
            if (invNote) invNote.textContent = '';

            setProgress(true, 'Teams-Gruppen werden geladen …');

            try {
                const token = await getGraphToken(GRAPH_SCOPES);
                const select = 'id,displayName,mail,description,visibility';
                const initial =
                    '/groups?$filter=' +
                    encodeURIComponent("resourceProvisioningOptions/Any(x:x eq 'Team')") +
                    '&$select=' +
                    encodeURIComponent(select) +
                    '&$count=true&$top=999';

                const teamGroups = await fetchAllPages(
                    token,
                    initial,
                    (p) => setProgress(true, 'Teams-Gruppen … Seite ' + p.page + ', ' + p.loaded),
                    { ConsistencyLevel: 'eventual' }
                );

                const tasks = teamGroups.map((g) => async () => {
                    const id = String(g.id || '');
                    if (!id) return null;
                    const guests = await fetchGuestUsersForGroup(token, id);
                    if (!guests.length) return null;
                    return {
                        id,
                        displayName: String(g.displayName || ''),
                        mail: String(g.mail || ''),
                        guests
                    };
                });

                if (tasks.length) {
                    renderTeamsLoadingRow(
                        'Gast-Mitglieder werden in ' +
                            tasks.length +
                            ' Team(s) geprüft … Treffer erscheinen fortlaufend in der Tabelle.'
                    );
                } else {
                    renderTeamsLoadingRow('Keine Teams im Mandant gefunden.');
                }

                const b2bAuditPromise = (async () => {
                    try {
                        const auditToken = await getGraphToken(B2B_AUDIT_SCOPES);
                        const auditPages = await fetchAllPages(
                            auditToken,
                            gzBuildB2bAuditInitialPath(),
                            undefined,
                            undefined
                        );
                        return { ok: true, data: gzMapDirectoryAuditsToB2bRows(auditPages || []) };
                    } catch (e2) {
                        return { ok: false, error: e2 };
                    }
                })();

                const accumulated = [];
                setProgress(true, 'Gast-Mitglieder pro Team prüfen … 0 / ' + tasks.length);
                let done = 0;
                await runPool(
                    tasks.map((fn) => async () => {
                        const row = await fn();
                        done++;
                        if (row) {
                            accumulated.push(row);
                            lastTeams = accumulated.slice().sort((a, b) => compareDe(a.displayName, b.displayName));
                            renderTeams();
                        }
                        if (done % 4 === 0 || done === tasks.length) {
                            setProgress(true, 'Gast-Mitglieder pro Team prüfen … ' + done + ' / ' + tasks.length);
                        }
                        return row;
                    }),
                    4
                );

                lastTeams = accumulated.slice().sort((a, b) => compareDe(a.displayName, b.displayName));
                if (!lastTeams.length) {
                    renderTeamsLoadingRow(
                        'Kein Team mit Gast-Benutzerkonten in der Mitgliedschaft gefunden.'
                    );
                } else {
                    renderTeams();
                }
                btnCsvTeams.disabled = !lastTeams.length;

                setProgress(true, 'B2B-Einladungen (Verzeichnis-Audit) werden geladen …');
                const b2bRes = await b2bAuditPromise;
                if (b2bRes.ok) {
                    lastInv = b2bRes.data;
                    lastInv.sort((a, b) =>
                        compareDe(a.invitedUserEmailAddress || '', b.invitedUserEmailAddress || '')
                    );
                    renderInv();
                    btnCsvInv.disabled = !lastInv.length;
                    if (invNote) {
                        invNote.textContent =
                            lastInv.length === 0
                                ? 'Keine passenden Audit-Einträge zu Gast-Einladungen (letzte Vorgänge). Hinweis: Es werden nur protokollierte Aktivitäten geladen, keine offenen Einladungs-„Postfächer“.'
                                : lastInv.length +
                                  ' Audit-Einträge zu B2B-/Gast-Einladungen (Verzeichnisprotokoll, lesend).';
                    }
                } else {
                    const e2 = b2bRes.error;
                    lastInv = [];
                    if (invNote) {
                        invNote.textContent =
                            'B2B-Audit konnte nicht geladen werden: ' +
                            (e2 && e2.message ? e2.message : String(e2)) +
                            ' Teams mit Gästen oben sind trotzdem gültig. Bitte in Entra die delegierte Berechtigung AuditLog.Read.All erteilen; der angemeldete Benutzer benötigt außerdem eine passende Rolle (z. B. Security Reader, Reports Reader).';
                    }
                }

                setProgress(
                    true,
                    'Fertig: ' +
                        lastTeams.length +
                        ' Teams mit Gästen' +
                        (lastInv.length ? ', ' + lastInv.length + ' B2B-Audit-Zeilen.' : '.')
                );
                gzSaveSnapshot(lastTeams, lastInv, invNote ? invNote.textContent : '');
                setTimeout(() => setProgress(false, ''), 3200);
            } catch (e) {
                setProgress(true, 'Fehler: ' + (e && e.message ? e.message : String(e)));
            } finally {
                btn.disabled = false;
            }
        });

        btnCsvTeams?.addEventListener('click', () => {
            if (!lastTeams.length) return;
            const lines = ['Team;E-Mail;Anzahl Gaeste;Gaeste (UPN/E-Mail)'];
            for (const r of lastTeams) {
                const upns = r.guests.map((g) => String(g.userPrincipalName || g.mail || '').trim()).join(', ');
                lines.push(
                    [r.displayName, r.mail, String(r.guests.length), upns].map((c) => csvEscape(c)).join(';')
                );
            }
            const csv = '\uFEFF' + lines.join('\n');
            const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = 'teams-mit-gaesten.csv';
            a.click();
            setTimeout(() => URL.revokeObjectURL(a.href), 4000);
        });

        btnCsvInv?.addEventListener('click', () => {
            if (!lastInv.length) return;
            const lines = ['Eingeladen_UPN;Zeit_Ergebnis;Name_oder_Initiator'];
            for (const inv of lastInv) {
                lines.push(
                    [inv.invitedUserEmailAddress || '', inv.status || '', inv.invitedUserDisplayName || '']
                        .map((c) => csvEscape(c))
                        .join(';')
                );
            }
            const csv = '\uFEFF' + lines.join('\n');
            const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = 'b2b-einladungen.csv';
            a.click();
            setTimeout(() => URL.revokeObjectURL(a.href), 4000);
        });
    }

    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', bind);
    else bind();
})();
