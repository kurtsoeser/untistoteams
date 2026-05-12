import { getEl } from '../../shared/utils/dom.js';
import { compareDe, escapeHtml } from '../../shared/utils/strings.js';

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

/** Scopes für den schreibenden Bereich (Gäste verwalten). */
const MGR_READ_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/User.Read.All',
    'https://graph.microsoft.com/AuditLog.Read.All'
];
const MGR_WRITE_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/User.ReadWrite.All',
    'https://graph.microsoft.com/Group.Read.All'
];

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

/* ============================================================
 *  Verwaltung aller Gäste im Tenant (lesen, bearbeiten, löschen)
 * ============================================================ */

/** ID des aktuell angemeldeten Benutzers (zum Selbst-Lösch-Schutz). */
async function gzGetCurrentUserId(token) {
    try {
        const me = await graphJson('GET', '/me?$select=id,userPrincipalName', token, undefined, undefined);
        return me && me.id ? String(me.id) : '';
    } catch {
        return '';
    }
}

/** Alle Gast-Benutzer des Tenants laden. signInActivity ist optional (braucht AuditLog.Read.All). */
async function gzFetchAllGuestUsers(token, onProgress) {
    const selectFields =
        'id,displayName,userPrincipalName,mail,userType,accountEnabled,companyName,jobTitle,givenName,surname,createdDateTime,externalUserState,externalUserStateChangeDateTime,signInActivity';
    const path =
        '/users?$filter=' +
        encodeURIComponent("userType eq 'Guest'") +
        '&$select=' +
        encodeURIComponent(selectFields) +
        '&$count=true&$top=999';
    try {
        return await fetchAllPages(token, path, onProgress, { ConsistencyLevel: 'eventual' });
    } catch (e) {
        // Wenn signInActivity nicht erlaubt ist (fehlende AuditLog.Read.All), fällt der Endpunkt
        // mit 403/400 zurück. In diesem Fall ohne signInActivity erneut versuchen.
        const status = e && e.status ? e.status : 0;
        if (status === 403 || status === 400) {
            const fallback = selectFields.replace(/,signInActivity/, '');
            const pathNoSignIn =
                '/users?$filter=' +
                encodeURIComponent("userType eq 'Guest'") +
                '&$select=' +
                encodeURIComponent(fallback) +
                '&$count=true&$top=999';
            return await fetchAllPages(token, pathNoSignIn, onProgress, { ConsistencyLevel: 'eventual' });
        }
        throw e;
    }
}

function gzLastSignInDate(g) {
    const a = g && g.signInActivity ? g.signInActivity : null;
    const candidates = [
        a && a.lastSignInDateTime,
        a && a.lastNonInteractiveSignInDateTime,
        a && a.lastSuccessfulSignInDateTime
    ].filter(Boolean);
    if (!candidates.length) return null;
    let max = null;
    for (const s of candidates) {
        const t = Date.parse(s);
        if (!isNaN(t) && (max === null || t > max)) max = t;
    }
    return max ? new Date(max) : null;
}

function gzFormatDate(iso) {
    if (!iso) return '';
    try {
        const d = typeof iso === 'string' ? new Date(iso) : iso;
        if (!d || isNaN(d.getTime())) return '';
        return d.toLocaleDateString('de-DE', { year: 'numeric', month: '2-digit', day: '2-digit' });
    } catch {
        return '';
    }
}

function gzDaysSince(d) {
    if (!d) return Infinity;
    return Math.floor((Date.now() - d.getTime()) / 86400000);
}

function gzStatusBadgeHtml(g) {
    const state = String(g.externalUserState || '').toLowerCase();
    if (g.accountEnabled === false) {
        return '<span class="gz-badge gz-badge-disabled">deaktiviert</span>';
    }
    if (state === 'pendingacceptance') {
        return '<span class="gz-badge gz-badge-pending">Pending</span>';
    }
    if (state === 'accepted') {
        return '<span class="gz-badge gz-badge-accepted">Akzeptiert</span>';
    }
    return '<span class="gz-badge gz-badge-disabled">' + escapeHtml(g.externalUserState || '–') + '</span>';
}

async function gzFetchMemberships(token, userId) {
    const out = [];
    let next =
        '/users/' +
        encodeURIComponent(userId) +
        "/memberOf/microsoft.graph.group?$select=id,displayName,mail,resourceProvisioningOptions&$top=200";
    let pages = 0;
    while (next && pages < 30) {
        pages++;
        try {
            const data = await graphJson('GET', next, token, undefined, undefined);
            const vals = data.value || [];
            for (const g of vals) {
                const opts = Array.isArray(g.resourceProvisioningOptions) ? g.resourceProvisioningOptions : [];
                out.push({
                    id: String(g.id || ''),
                    displayName: String(g.displayName || ''),
                    mail: String(g.mail || ''),
                    isTeam: opts.map((s) => String(s).toLowerCase()).indexOf('team') !== -1
                });
            }
            next = data['@odata.nextLink'] || null;
        } catch {
            next = null;
        }
    }
    out.sort((a, b) => compareDe(a.displayName, b.displayName));
    return out;
}

function bindManager() {
    const els = {
        btnLoad: document.getElementById('gzMgrBtnLoad'),
        btnCsv: document.getElementById('gzMgrBtnCsv'),
        search: document.getElementById('gzMgrSearch'),
        filterStatus: document.getElementById('gzMgrFilterStatus'),
        filterStale: document.getElementById('gzMgrFilterStale'),
        tbody: document.getElementById('gzMgrTbody'),
        count: document.getElementById('gzMgrCount'),
        progress: document.getElementById('gzMgrProgress'),
        // Bulk-Auswahl
        selectAll: document.getElementById('gzMgrSelectAll'),
        bulkBar: document.getElementById('gzMgrBulkBar'),
        bulkCount: document.getElementById('gzMgrBulkCount'),
        bulkHint: document.getElementById('gzMgrBulkHint'),
        bulkClear: document.getElementById('gzMgrBulkClear'),
        bulkDelete: document.getElementById('gzMgrBulkDelete'),
        // Edit-Modal
        editModal: document.getElementById('gzEditModal'),
        editUpn: document.getElementById('gzEditUpn'),
        editDisplay: document.getElementById('gzEditDisplayName'),
        editGiven: document.getElementById('gzEditGivenName'),
        editSurname: document.getElementById('gzEditSurname'),
        editCompany: document.getElementById('gzEditCompany'),
        editJobTitle: document.getElementById('gzEditJobTitle'),
        editSave: document.getElementById('gzEditSave'),
        editError: document.getElementById('gzEditError'),
        // Delete-Modal (einzeln)
        delModal: document.getElementById('gzDeleteModal'),
        delDisplay: document.getElementById('gzDelDisplay'),
        delUpn: document.getElementById('gzDelUpn'),
        delMemberships: document.getElementById('gzDelMembershipsBox'),
        delConfirm: document.getElementById('gzDelConfirm'),
        delExecute: document.getElementById('gzDelExecute'),
        delError: document.getElementById('gzDelError'),
        // Bulk-Delete-Modal
        bulkModal: document.getElementById('gzBulkDeleteModal'),
        bulkTitleText: document.getElementById('gzBulkDelTitleText'),
        bulkCountText: document.getElementById('gzBulkDelCountText'),
        bulkList: document.getElementById('gzBulkDelList'),
        bulkConfirm: document.getElementById('gzBulkDelConfirm'),
        bulkExecute: document.getElementById('gzBulkDelExecute'),
        bulkCancel: document.getElementById('gzBulkDelCancel'),
        bulkError: document.getElementById('gzBulkDelError'),
        bulkProgressWrap: document.getElementById('gzBulkDelProgressWrap'),
        bulkProgressBar: document.getElementById('gzBulkDelProgressBar'),
        bulkProgressText: document.getElementById('gzBulkDelProgressText')
    };
    if (!els.btnLoad || !els.tbody) return;
    if (els.btnLoad.dataset.gzMgrBound === '1') return;
    els.btnLoad.dataset.gzMgrBound = '1';

    let allGuests = [];
    let viewGuests = [];
    const selectedIds = new Set();
    let currentEditId = null;
    let currentDeleteId = null;
    let currentDeleteUpn = '';
    let currentUserOid = '';
    let bulkInProgress = false;

    function setProgress(on, text) {
        if (!els.progress) return;
        els.progress.style.display = on ? '' : 'none';
        if (text) els.progress.textContent = String(text);
    }

    function applyFilters() {
        const q = String(els.search.value || '').trim().toLowerCase();
        const status = String(els.filterStatus.value || 'all');
        const stale = String(els.filterStale.value || 'all');

        viewGuests = allGuests.filter((g) => {
            if (q) {
                const hay =
                    (g.displayName || '') +
                    ' ' +
                    (g.userPrincipalName || '') +
                    ' ' +
                    (g.mail || '') +
                    ' ' +
                    (g.companyName || '') +
                    ' ' +
                    (g.jobTitle || '');
                if (hay.toLowerCase().indexOf(q) === -1) return false;
            }
            if (status === 'accepted' && String(g.externalUserState || '').toLowerCase() !== 'accepted') return false;
            if (status === 'pending' && String(g.externalUserState || '').toLowerCase() !== 'pendingacceptance') return false;
            if (status === 'disabled' && g.accountEnabled !== false) return false;

            if (stale !== 'all') {
                const last = gzLastSignInDate(g);
                if (stale === 'never') {
                    if (last) return false;
                } else {
                    const threshold = stale === 'stale90' ? 90 : stale === 'stale180' ? 180 : 365;
                    if (gzDaysSince(last) <= threshold) return false;
                }
            }
            return true;
        });
        renderRows();
    }

    function renderRows() {
        if (!els.tbody) return;
        els.tbody.replaceChildren();
        els.count.textContent =
            viewGuests.length === allGuests.length
                ? allGuests.length + ' Gäste geladen'
                : viewGuests.length + ' von ' + allGuests.length + ' Gästen';
        if (!viewGuests.length) {
            const tr = document.createElement('tr');
            tr.innerHTML =
                '<td colspan="7" class="muted" style="padding:14px 10px;">Keine Treffer für die aktuelle Filter-/Suchkombination.</td>';
            els.tbody.appendChild(tr);
            updateSelectionUi();
            return;
        }
        const frag = document.createDocumentFragment();
        for (const g of viewGuests) {
            const tr = document.createElement('tr');
            tr.dataset.guestId = g.id;
            const last = gzLastSignInDate(g);
            const lastLabel = last
                ? gzFormatDate(last) + ' <span class="muted">(' + gzDaysSince(last) + ' Tage)</span>'
                : '<span class="muted">noch nie</span>';
            const isSelf = currentUserOid && currentUserOid === g.id;
            if (selectedIds.has(g.id)) tr.classList.add('is-selected');
            const cbHtml = isSelf
                ? '<input type="checkbox" disabled title="Sie können sich nicht selbst löschen.">'
                : '<input type="checkbox" data-gz-sel="row" aria-label="Auswählen"' +
                  (selectedIds.has(g.id) ? ' checked' : '') +
                  '>';
            const delBtn = isSelf
                ? '<button type="button" class="btn btn-danger" disabled title="Sie können sich nicht selbst löschen.">' +
                  '<i class="bi bi-trash"></i></button>'
                : '<button type="button" class="btn btn-danger" data-gz-act="del"><i class="bi bi-trash"></i></button>';
            tr.innerHTML =
                '<td class="gz-col-select">' +
                '<label class="gz-checkbox-label">' +
                cbHtml +
                '</label>' +
                '</td>' +
                '<td><strong>' +
                escapeHtml(g.displayName || '–') +
                '</strong>' +
                (g.companyName ? '<br><span class="muted" style="font-size:0.88em;">' + escapeHtml(g.companyName) + '</span>' : '') +
                '</td>' +
                '<td>' +
                escapeHtml(g.userPrincipalName || g.mail || '') +
                (g.mail && g.mail !== g.userPrincipalName
                    ? '<br><span class="muted" style="font-size:0.88em;">' + escapeHtml(g.mail) + '</span>'
                    : '') +
                '</td>' +
                '<td>' +
                gzStatusBadgeHtml(g) +
                '</td>' +
                '<td>' +
                escapeHtml(gzFormatDate(g.createdDateTime)) +
                '</td>' +
                '<td>' +
                lastLabel +
                '</td>' +
                '<td style="text-align:right;">' +
                '<div class="gz-row-actions">' +
                '<button type="button" class="btn" data-gz-act="edit" title="Bearbeiten"><i class="bi bi-pencil"></i></button>' +
                delBtn +
                '</div>' +
                '</td>';
            frag.appendChild(tr);
        }
        els.tbody.appendChild(frag);
        updateSelectionUi();
    }

    function selectableViewIds() {
        const out = [];
        for (const g of viewGuests) {
            if (currentUserOid && currentUserOid === g.id) continue;
            out.push(g.id);
        }
        return out;
    }

    function updateSelectionUi() {
        const total = selectedIds.size;
        const visibleIds = selectableViewIds();
        const visibleSelected = visibleIds.filter((id) => selectedIds.has(id)).length;

        if (els.selectAll) {
            if (!visibleIds.length) {
                els.selectAll.checked = false;
                els.selectAll.indeterminate = false;
                els.selectAll.disabled = true;
            } else {
                els.selectAll.disabled = false;
                if (visibleSelected === 0) {
                    els.selectAll.checked = false;
                    els.selectAll.indeterminate = false;
                } else if (visibleSelected === visibleIds.length) {
                    els.selectAll.checked = true;
                    els.selectAll.indeterminate = false;
                } else {
                    els.selectAll.checked = false;
                    els.selectAll.indeterminate = true;
                }
            }
        }

        if (els.bulkBar) {
            if (total === 0) {
                els.bulkBar.setAttribute('hidden', '');
            } else {
                els.bulkBar.removeAttribute('hidden');
            }
        }
        if (els.bulkCount) els.bulkCount.textContent = String(total);
        if (els.bulkHint) {
            const hiddenSelected = total - visibleSelected;
            if (total === 0) {
                els.bulkHint.textContent = '';
            } else if (hiddenSelected > 0) {
                els.bulkHint.textContent =
                    '(' +
                    visibleSelected +
                    ' aktuell sichtbar, ' +
                    hiddenSelected +
                    ' nicht im aktuellen Filter angezeigt)';
            } else {
                els.bulkHint.textContent =
                    visibleSelected === 1 ? '(1 sichtbarer Treffer)' : '(' + visibleSelected + ' sichtbare Treffer)';
            }
        }
        if (els.bulkDelete) els.bulkDelete.disabled = total === 0 || bulkInProgress;
        if (els.bulkClear) els.bulkClear.disabled = total === 0 || bulkInProgress;
    }

    function toggleRowSelection(id, on) {
        if (!id) return;
        if (currentUserOid && currentUserOid === id) return;
        const tr = els.tbody.querySelector('tr[data-guest-id="' + (CSS && CSS.escape ? CSS.escape(id) : id) + '"]');
        if (on) selectedIds.add(id);
        else selectedIds.delete(id);
        if (tr) tr.classList.toggle('is-selected', on);
        updateSelectionUi();
    }

    function selectAllVisible(on) {
        const ids = selectableViewIds();
        if (on) ids.forEach((id) => selectedIds.add(id));
        else ids.forEach((id) => selectedIds.delete(id));
        // Reine DOM-Synchronisation, ohne Rerender (Performance bei vielen Zeilen):
        const rows = els.tbody.querySelectorAll('tr[data-guest-id]');
        rows.forEach((tr) => {
            const id = tr.dataset.guestId;
            const cb = tr.querySelector('input[data-gz-sel="row"]');
            const selected = selectedIds.has(id);
            if (cb) cb.checked = selected;
            tr.classList.toggle('is-selected', selected);
        });
        updateSelectionUi();
    }

    function clearSelection() {
        selectedIds.clear();
        const rows = els.tbody.querySelectorAll('tr[data-guest-id]');
        rows.forEach((tr) => {
            const cb = tr.querySelector('input[data-gz-sel="row"]');
            if (cb) cb.checked = false;
            tr.classList.remove('is-selected');
        });
        updateSelectionUi();
    }

    // Tabellen-Aktionen via Event-Delegation
    els.tbody.addEventListener('click', (e) => {
        const btn = e.target.closest('button[data-gz-act]');
        if (!btn) return;
        const tr = btn.closest('tr');
        const id = tr && tr.dataset.guestId;
        if (!id) return;
        const guest = allGuests.find((g) => g.id === id);
        if (!guest) return;
        const act = btn.getAttribute('data-gz-act');
        if (act === 'edit') openEdit(guest);
        else if (act === 'del') openDelete(guest);
    });

    // Zeilen-Checkboxen
    els.tbody.addEventListener('change', (e) => {
        const cb = e.target.closest('input[data-gz-sel="row"]');
        if (!cb) return;
        const tr = cb.closest('tr');
        const id = tr && tr.dataset.guestId;
        if (!id) return;
        toggleRowSelection(id, !!cb.checked);
    });

    // Header-Checkbox: alle sichtbaren auswählen / abwählen
    if (els.selectAll) {
        els.selectAll.addEventListener('change', () => {
            selectAllVisible(!!els.selectAll.checked);
        });
    }
    if (els.bulkClear) els.bulkClear.addEventListener('click', clearSelection);
    if (els.bulkDelete) els.bulkDelete.addEventListener('click', openBulkDelete);

    function openModal(modal) {
        if (!modal) return;
        modal.classList.add('is-open');
    }
    function closeModal(modal) {
        if (!modal) return;
        modal.classList.remove('is-open');
    }

    document.querySelectorAll('[data-gz-close]').forEach((el) => {
        el.addEventListener('click', () => {
            const which = el.getAttribute('data-gz-close');
            if (which === 'edit') closeModal(els.editModal);
            else if (which === 'delete') closeModal(els.delModal);
        });
    });
    // Klick außerhalb des Modal-Containers schließt das Modal
    [els.editModal, els.delModal].forEach((m) => {
        if (!m) return;
        m.addEventListener('click', (e) => {
            if (e.target === m) closeModal(m);
        });
    });

    function openEdit(g) {
        currentEditId = g.id;
        els.editUpn.value = g.userPrincipalName || g.mail || '';
        els.editDisplay.value = g.displayName || '';
        els.editGiven.value = g.givenName || '';
        els.editSurname.value = g.surname || '';
        els.editCompany.value = g.companyName || '';
        els.editJobTitle.value = g.jobTitle || '';
        els.editError.textContent = '';
        openModal(els.editModal);
        setTimeout(() => els.editDisplay.focus(), 30);
    }

    async function saveEdit() {
        if (!currentEditId) return;
        const id = currentEditId;
        const original = allGuests.find((g) => g.id === id);
        if (!original) return;
        const patch = {};
        const newDisplay = String(els.editDisplay.value || '').trim();
        const newGiven = String(els.editGiven.value || '').trim();
        const newSurname = String(els.editSurname.value || '').trim();
        const newCompany = String(els.editCompany.value || '').trim();
        const newJob = String(els.editJobTitle.value || '').trim();

        if (!newDisplay) {
            els.editError.textContent = 'Der Anzeigename darf nicht leer sein.';
            els.editDisplay.focus();
            return;
        }
        if (newDisplay !== (original.displayName || '')) patch.displayName = newDisplay;
        if (newGiven !== (original.givenName || '')) patch.givenName = newGiven || null;
        if (newSurname !== (original.surname || '')) patch.surname = newSurname || null;
        if (newCompany !== (original.companyName || '')) patch.companyName = newCompany || null;
        if (newJob !== (original.jobTitle || '')) patch.jobTitle = newJob || null;

        if (!Object.keys(patch).length) {
            els.editError.textContent = 'Es wurden keine Änderungen vorgenommen.';
            return;
        }

        els.editSave.disabled = true;
        els.editError.textContent = 'Speichere …';
        try {
            const token = await getGraphToken(MGR_WRITE_SCOPES);
            const res = await graphRequest('PATCH', '/users/' + encodeURIComponent(id), token, patch, undefined);
            if (!res.ok && res.status !== 204) {
                const text = await res.text();
                let msg = text;
                try {
                    const j = JSON.parse(text);
                    msg = (j && j.error && (j.error.message || JSON.stringify(j.error))) || text;
                } catch {
                    /* keep msg */
                }
                throw new Error('HTTP ' + res.status + ': ' + msg);
            }
            Object.assign(original, {
                displayName: newDisplay,
                givenName: newGiven,
                surname: newSurname,
                companyName: newCompany,
                jobTitle: newJob
            });
            els.editError.textContent = '';
            closeModal(els.editModal);
            applyFilters();
        } catch (e) {
            els.editError.textContent = 'Fehler beim Speichern: ' + (e && e.message ? e.message : String(e));
        } finally {
            els.editSave.disabled = false;
        }
    }
    els.editSave.addEventListener('click', saveEdit);

    async function openDelete(g) {
        currentDeleteId = g.id;
        currentDeleteUpn = String(g.userPrincipalName || g.mail || '');
        els.delDisplay.value = g.displayName || '';
        els.delUpn.value = currentDeleteUpn;
        els.delConfirm.value = '';
        els.delExecute.disabled = true;
        els.delError.textContent = '';
        els.delMemberships.textContent = 'Mitgliedschaften werden geladen …';
        openModal(els.delModal);
        setTimeout(() => els.delConfirm.focus(), 30);

        try {
            const token = await getGraphToken(MGR_WRITE_SCOPES);
            const memberships = await gzFetchMemberships(token, g.id);
            if (!memberships.length) {
                els.delMemberships.innerHTML =
                    '<em>Keine Gruppen-/Team-Mitgliedschaften gefunden.</em>';
            } else {
                const items = memberships
                    .map(
                        (m) =>
                            '<li>' +
                            (m.isTeam ? '<i class="bi bi-microsoft-teams" style="margin-right:6px;"></i>' : '<i class="bi bi-people" style="margin-right:6px;"></i>') +
                            escapeHtml(m.displayName || '–') +
                            (m.mail ? ' <span class="muted">&lt;' + escapeHtml(m.mail) + '&gt;</span>' : '') +
                            '</li>'
                    )
                    .join('');
                els.delMemberships.innerHTML =
                    '<strong>Wird aus folgenden Gruppen/Teams entfernt (' +
                    memberships.length +
                    '):</strong><ul>' +
                    items +
                    '</ul>';
            }
        } catch (e) {
            els.delMemberships.innerHTML =
                '<em>Mitgliedschaften konnten nicht geladen werden: ' +
                escapeHtml(e && e.message ? e.message : String(e)) +
                '</em>';
        }
    }

    els.delConfirm.addEventListener('input', () => {
        const ok = String(els.delConfirm.value || '').trim().toLowerCase() === currentDeleteUpn.toLowerCase() && !!currentDeleteUpn;
        els.delExecute.disabled = !ok;
    });

    async function executeDelete() {
        if (!currentDeleteId) return;
        const id = currentDeleteId;
        if (currentUserOid && currentUserOid === id) {
            els.delError.textContent = 'Sie können sich nicht selbst löschen.';
            return;
        }
        els.delExecute.disabled = true;
        els.delError.textContent = 'Lösche …';
        try {
            const token = await getGraphToken(MGR_WRITE_SCOPES);
            const res = await graphRequest('DELETE', '/users/' + encodeURIComponent(id), token, undefined, undefined);
            if (!res.ok && res.status !== 204) {
                const text = await res.text();
                let msg = text;
                try {
                    const j = JSON.parse(text);
                    msg = (j && j.error && (j.error.message || JSON.stringify(j.error))) || text;
                } catch {
                    /* keep msg */
                }
                throw new Error('HTTP ' + res.status + ': ' + msg);
            }
            allGuests = allGuests.filter((g) => g.id !== id);
            closeModal(els.delModal);
            applyFilters();
        } catch (e) {
            els.delError.textContent = 'Fehler beim Löschen: ' + (e && e.message ? e.message : String(e));
            els.delExecute.disabled = false;
        }
    }
    els.delExecute.addEventListener('click', executeDelete);

    // === Bulk-Löschen =========================================================
    function getBulkTargets() {
        const ids = Array.from(selectedIds);
        const targets = [];
        for (const id of ids) {
            if (currentUserOid && currentUserOid === id) continue;
            const g = allGuests.find((x) => x.id === id);
            if (g) targets.push(g);
        }
        return targets;
    }

    function setBulkExecuteEnabled() {
        if (!els.bulkExecute || !els.bulkConfirm) return;
        const phrase = String(els.bulkConfirm.value || '').trim().toUpperCase();
        const ok = phrase === 'LÖSCHEN' || phrase === 'LOESCHEN';
        els.bulkExecute.disabled = !ok || bulkInProgress;
    }

    function openBulkDelete() {
        if (!els.bulkModal) return;
        const targets = getBulkTargets();
        if (!targets.length) return;
        els.bulkCountText.textContent = String(targets.length);
        els.bulkTitleText.textContent =
            targets.length === 1 ? '1 Gast löschen' : targets.length + ' Gäste löschen';
        const itemsHtml = targets
            .map(
                (g) =>
                    '<li><strong>' +
                    escapeHtml(g.displayName || '–') +
                    '</strong>' +
                    '<br><span class="muted" style="font-size:0.88em;">' +
                    escapeHtml(g.userPrincipalName || g.mail || '') +
                    '</span></li>'
            )
            .join('');
        els.bulkList.innerHTML =
            '<strong>Folgende ' +
            targets.length +
            ' Konten werden entfernt:</strong>' +
            '<ul style="margin:6px 0 0;padding-left:20px;">' +
            itemsHtml +
            '</ul>';
        els.bulkConfirm.value = '';
        els.bulkExecute.disabled = true;
        els.bulkError.textContent = '';
        els.bulkProgressWrap.style.display = 'none';
        els.bulkProgressBar.style.width = '0%';
        els.bulkProgressText.textContent = '0 / ' + targets.length;
        if (els.bulkCancel) els.bulkCancel.disabled = false;
        bulkInProgress = false;
        openModal(els.bulkModal);
        setTimeout(() => els.bulkConfirm.focus(), 30);
    }

    if (els.bulkConfirm) els.bulkConfirm.addEventListener('input', setBulkExecuteEnabled);

    async function executeBulkDelete() {
        if (bulkInProgress) return;
        const targets = getBulkTargets();
        if (!targets.length) return;
        bulkInProgress = true;
        setBulkExecuteEnabled();
        if (els.bulkCancel) els.bulkCancel.disabled = true;
        if (els.bulkClear) els.bulkClear.disabled = true;
        els.bulkError.textContent = '';
        els.bulkProgressWrap.style.display = '';
        els.bulkProgressBar.style.width = '0%';
        els.bulkProgressText.textContent = '0 / ' + targets.length;

        let token;
        try {
            token = await getGraphToken(MGR_WRITE_SCOPES);
        } catch (e) {
            els.bulkError.textContent = 'Token-Fehler: ' + (e && e.message ? e.message : String(e));
            bulkInProgress = false;
            setBulkExecuteEnabled();
            if (els.bulkCancel) els.bulkCancel.disabled = false;
            updateSelectionUi();
            return;
        }

        const succeeded = [];
        const failed = [];
        let done = 0;
        for (const g of targets) {
            try {
                const res = await graphRequest(
                    'DELETE',
                    '/users/' + encodeURIComponent(g.id),
                    token,
                    undefined,
                    undefined
                );
                if (!res.ok && res.status !== 204) {
                    const text = await res.text().catch(() => '');
                    let msg = text;
                    try {
                        const j = JSON.parse(text);
                        msg = (j && j.error && (j.error.message || JSON.stringify(j.error))) || text;
                    } catch {
                        /* keep msg */
                    }
                    throw new Error('HTTP ' + res.status + ': ' + msg);
                }
                succeeded.push(g.id);
            } catch (e) {
                failed.push({
                    id: g.id,
                    upn: g.userPrincipalName || g.mail || '',
                    displayName: g.displayName || '',
                    message: e && e.message ? e.message : String(e)
                });
            }
            done++;
            const pct = Math.round((done / targets.length) * 100);
            els.bulkProgressBar.style.width = pct + '%';
            els.bulkProgressText.textContent = done + ' / ' + targets.length;
        }

        if (succeeded.length) {
            const set = new Set(succeeded);
            allGuests = allGuests.filter((g) => !set.has(g.id));
            succeeded.forEach((id) => selectedIds.delete(id));
            applyFilters();
        }

        bulkInProgress = false;
        if (els.bulkCancel) els.bulkCancel.disabled = false;
        updateSelectionUi();

        if (!failed.length) {
            els.bulkError.style.color = '#22543d';
            els.bulkError.textContent =
                'Alle ' + succeeded.length + ' Konten wurden erfolgreich entfernt.';
            setTimeout(() => {
                closeModal(els.bulkModal);
                els.bulkError.style.color = '';
            }, 1400);
        } else {
            els.bulkError.style.color = '';
            const list = failed
                .slice(0, 5)
                .map(
                    (f) =>
                        '• ' +
                        (f.displayName || f.upn || f.id) +
                        (f.upn ? ' (' + f.upn + ')' : '') +
                        ': ' +
                        f.message
                )
                .join('\n');
            const more = failed.length > 5 ? '\n… und ' + (failed.length - 5) + ' weitere' : '';
            els.bulkError.style.whiteSpace = 'pre-wrap';
            els.bulkError.textContent =
                succeeded.length +
                ' erfolgreich gelöscht, ' +
                failed.length +
                ' fehlgeschlagen:\n' +
                list +
                more;
            setBulkExecuteEnabled();
        }
    }
    if (els.bulkExecute) els.bulkExecute.addEventListener('click', executeBulkDelete);
    // Schließen-Buttons im Bulk-Modal
    document.querySelectorAll('[data-gz-close="bulk"]').forEach((el) => {
        el.addEventListener('click', () => {
            if (bulkInProgress) return; // während Lauf nicht schließen
            closeModal(els.bulkModal);
        });
    });
    if (els.bulkModal) {
        els.bulkModal.addEventListener('click', (e) => {
            if (bulkInProgress) return;
            if (e.target === els.bulkModal) closeModal(els.bulkModal);
        });
    }

    // Suche/Filter
    let searchTimer = null;
    els.search.addEventListener('input', () => {
        if (searchTimer) clearTimeout(searchTimer);
        searchTimer = setTimeout(applyFilters, 150);
    });
    els.filterStatus.addEventListener('change', applyFilters);
    els.filterStale.addEventListener('change', applyFilters);

    // ESC schließt offene Modale
    document.addEventListener('keydown', (e) => {
        if (e.key !== 'Escape') return;
        if (els.editModal && els.editModal.classList.contains('is-open')) closeModal(els.editModal);
        else if (els.delModal && els.delModal.classList.contains('is-open')) closeModal(els.delModal);
        else if (els.bulkModal && els.bulkModal.classList.contains('is-open') && !bulkInProgress) closeModal(els.bulkModal);
    });

    els.btnLoad.addEventListener('click', async () => {
        els.btnLoad.disabled = true;
        els.btnCsv.disabled = true;
        setProgress(true, 'Gäste werden geladen …');
        try {
            const token = await getGraphToken(MGR_READ_SCOPES);
            if (!currentUserOid) currentUserOid = await gzGetCurrentUserId(token);
            const items = await gzFetchAllGuestUsers(token, (p) => {
                setProgress(true, 'Gäste laden … Seite ' + p.page + ', insgesamt ' + p.loaded);
            });
            allGuests = (items || []).slice().sort((a, b) => compareDe(a.displayName || a.userPrincipalName, b.displayName || b.userPrincipalName));
            selectedIds.clear();
            applyFilters();
            els.btnCsv.disabled = !allGuests.length;
            setProgress(true, 'Fertig: ' + allGuests.length + ' Gast-Benutzer.');
            setTimeout(() => setProgress(false, ''), 2400);
        } catch (e) {
            setProgress(true, 'Fehler: ' + (e && e.message ? e.message : String(e)));
        } finally {
            els.btnLoad.disabled = false;
        }
    });

    els.btnCsv.addEventListener('click', () => {
        if (!viewGuests.length) return;
        const header = [
            'Anzeigename',
            'UPN',
            'Mail',
            'Status',
            'Firma',
            'Funktion',
            'Aktiviert',
            'Erstellt',
            'Letzte Anmeldung',
            'Tage_inaktiv'
        ];
        const lines = [header.map(csvEscape).join(';')];
        for (const g of viewGuests) {
            const last = gzLastSignInDate(g);
            const days = last ? String(gzDaysSince(last)) : '';
            const status = g.externalUserState || '';
            lines.push(
                [
                    g.displayName || '',
                    g.userPrincipalName || '',
                    g.mail || '',
                    status,
                    g.companyName || '',
                    g.jobTitle || '',
                    g.accountEnabled === false ? 'nein' : 'ja',
                    gzFormatDate(g.createdDateTime),
                    last ? gzFormatDate(last) : '',
                    days
                ]
                    .map(csvEscape)
                    .join(';')
            );
        }
        const csv = '\uFEFF' + lines.join('\n');
        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = 'tenant-gaeste.csv';
        a.click();
        setTimeout(() => URL.revokeObjectURL(a.href), 4000);
    });
}

function bindAll() {
    bind();
    bindManager();
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', bindAll);
else bindAll();
