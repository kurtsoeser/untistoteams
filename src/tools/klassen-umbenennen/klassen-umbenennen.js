
const GRAPH_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/User.Read.All',
    'https://graph.microsoft.com/Group.ReadWrite.All'
];

const GROUP_SELECT = 'id,displayName,mail,mailNickname,resourceProvisioningOptions';
const STEP_ORDER = [1, 2, 3, 4];

let msalMod = null;
let pca = null;
/** @type {{ id: string, displayName: string, mail: string, mailNickname: string, resourceProvisioningOptions?: string[] }[]} */
let loadedGroups = [];
/** @type {Record<string, { ownersText: string, loading: boolean, err: string }>} */
let ownersCache = Object.create(null);
let showTenantClassesInStep1 = false;
/** @type {{ find: string, replaceWith: string }[]} */
let replaceRules = [];
/** @type {Array<{ id: string, displayNameOld: string, displayNameNew: string, mailNicknameOld: string, mailNicknameNew: string, mailOld: string, hint: string, ok: boolean }>} */
let previewRows = [];

function toast(msg) {
    const el = document.getElementById('toast');
    if (el) {
        el.textContent = msg;
        el.classList.add('show');
        clearTimeout(toast._t);
        toast._t = setTimeout(() => el.classList.remove('show'), 3800);
    } else if (typeof window.ms365ToastOrAlert === 'function') {
        window.ms365ToastOrAlert(msg);
    } else if (typeof window.ms365ShowToast === 'function') {
        window.ms365ShowToast(msg);
    } else {
        window.alert(msg);
    }
}

async function loadMsal() {
    if (msalMod) return msalMod;
    try {
        msalMod = await import('https://esm.sh/@azure/msal-browser@3.26.1');
    } catch {
        msalMod = await import('https://cdn.jsdelivr.net/npm/@azure/msal-browser@3.26.1/+esm');
    }
    return msalMod;
}

function isInteractionRequired(e) {
    return (
        e &&
        (e.name === 'InteractionRequiredAuthError' ||
            e.errorCode === 'interaction_required' ||
            (typeof e.message === 'string' && e.message.indexOf('interaction_required') !== -1))
    );
}

function resolveMsalConfig() {
    let cfg = window.MS365_MSAL_CONFIG;
    if (!cfg) cfg = {};
    let id = String(cfg.clientId || '').trim();
    if (!id) {
        const meta = document.querySelector('meta[name="ms365-graph-client-id"]');
        const fromMeta = meta && meta.getAttribute('content') ? meta.getAttribute('content').trim() : '';
        if (fromMeta) id = fromMeta;
    }
    if (!id) {
        throw new Error(
            'Keine clientId: ms365-config.js fehlt/leer oder blockiert. Seite mit Strg+F5 neu laden.'
        );
    }
    return {
        clientId: id,
        authority: cfg.authority || 'https://login.microsoftonline.com/organizations',
        redirectUri: (cfg.redirectUri || window.location.href.split('#')[0]).trim()
    };
}

async function getPca() {
    const m = await loadMsal();
    const PublicClientApplication = m.PublicClientApplication || (m.default && m.default.PublicClientApplication);
    if (!PublicClientApplication) {
        throw new Error('MSAL: PublicClientApplication nicht gefunden (Import).');
    }
    const cfg = resolveMsalConfig();
    if (!pca) {
        pca = new PublicClientApplication({
            auth: {
                clientId: cfg.clientId,
                authority: cfg.authority,
                redirectUri: cfg.redirectUri
            },
            cache: {
                cacheLocation: 'sessionStorage',
                storeAuthStateInCookie: true
            }
        });
        await pca.initialize();
        await pca.handleRedirectPromise();
    }
    return pca;
}

async function getGraphToken() {
    const instance = await getPca();
    let accounts = instance.getAllAccounts();
    if (!accounts.length) {
        await instance.loginPopup({ scopes: GRAPH_SCOPES, prompt: 'select_account' });
        accounts = instance.getAllAccounts();
    }
    if (!accounts.length) {
        throw new Error('Anmeldung abgebrochen.');
    }
    const req = { scopes: GRAPH_SCOPES, account: accounts[0] };
    try {
        return (await instance.acquireTokenSilent(req)).accessToken;
    } catch (e) {
        if (isInteractionRequired(e)) {
            return (await instance.acquireTokenPopup(req)).accessToken;
        }
        throw e;
    }
}

function sleep(ms) {
    return new Promise(function (r) {
        setTimeout(r, ms);
    });
}

async function graphRequest(method, pathOrUrl, token, body) {
    const url =
        pathOrUrl.indexOf('http') === 0 ? pathOrUrl : 'https://graph.microsoft.com/v1.0' + pathOrUrl;
    let attempt = 0;
    while (true) {
        const headers = { Authorization: 'Bearer ' + token };
        if (body !== undefined) {
            headers['Content-Type'] = 'application/json';
        }
        const res = await fetch(url, {
            method: method,
            headers: headers,
            body: body !== undefined ? JSON.stringify(body) : undefined
        });
        if (res.status === 429 && attempt < 8) {
            const ra = parseInt(res.headers.get('Retry-After') || '5', 10);
            await sleep((isNaN(ra) ? 5 : ra) * 1000);
            attempt++;
            continue;
        }
        return res;
    }
}

async function graphJson(method, pathOrUrl, token, body) {
    const res = await graphRequest(method, pathOrUrl, token, body);
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
        throw new Error(method + ' ' + pathOrUrl + ': ' + msg);
    }
    return data || {};
}

function appendLog(msg, kind) {
    const el = document.getElementById('kuLog');
    if (!el) return;
    const line = document.createElement('div');
    line.textContent = new Date().toLocaleTimeString() + '  ' + msg;
    if (kind === 'err') line.style.color = '#b00020';
    else if (kind === 'ok') line.style.color = '#0d8050';
    else if (kind === 'warn') line.style.color = '#856404';
    else line.style.color = '#212529';
    el.appendChild(line);
    el.scrollTop = el.scrollHeight;
}

function clearLog() {
    const el = document.getElementById('kuLog');
    if (el) el.replaceChildren();
}

function groupHasTeamProvisioning(g) {
    const opts = g && g.resourceProvisioningOptions;
    return Array.isArray(opts) && opts.indexOf('Team') !== -1;
}

async function fetchAllPages(token, initialPath) {
    const out = [];
    let next = initialPath;
    while (next) {
        const data = await graphJson('GET', next, token, undefined);
        const vals = data.value;
        if (Array.isArray(vals)) {
            for (let i = 0; i < vals.length; i++) out.push(vals[i]);
        }
        next = data['@odata.nextLink'] || null;
    }
    return out;
}

function escapeRe(s) {
    return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function odataEscape(s) {
    return String(s).replace(/'/g, "''");
}

function deriveMailNickname(displayName) {
    let s = displayName.trim().toLowerCase().replace(/\s+/g, '');
    s = s.replace(/[^a-z0-9\-]/g, '');
    if (s.length > 64) s = s.slice(0, 64);
    return s;
}

function applyReplaceRules(input) {
    let out = String(input || '');
    for (let i = 0; i < replaceRules.length; i++) {
        const r = replaceRules[i];
        const f = r && r.find != null ? String(r.find) : '';
        const w = r && r.replaceWith != null ? String(r.replaceWith) : '';
        if (!f) continue;
        out = out.split(f).join(w);
    }
    return out;
}

/**
 * Erwartetes Muster: "<Präfix> <Stufe><Kürzel>" z. B. "Klasse 1A", "Klasse 10HAK"
 */
function computeNewDisplayNamePlusOne(displayName, prefix) {
    const p = String(prefix || '').trim();
    if (!p) return null;
    const re = new RegExp('^' + escapeRe(p) + '\\s+(\\d+)([A-Za-z0-9\\-]*)$', 'i');
    const m = String(displayName || '').trim().match(re);
    if (!m) return null;
    const current = parseInt(m[1], 10);
    if (!isFinite(current)) return null;
    const next = current + 1;
    return p + ' ' + String(next) + (m[2] || '');
}

function getFilteredGroups() {
    const onlyTeams = document.getElementById('kuOnlyTeams');
    const filtInp = document.getElementById('kuFilterContains');
    const sub = filtInp && filtInp.value ? String(filtInp.value).trim().toLowerCase() : '';
    let list = loadedGroups.slice();
    if (onlyTeams && onlyTeams.checked) {
        list = list.filter(groupHasTeamProvisioning);
    }
    if (sub) {
        list = list.filter(function (g) {
            return String(g.displayName || '')
                .toLowerCase()
                .indexOf(sub) !== -1;
        });
    }
    return list;
}

function getTenantClassesMatcher() {
    if (typeof window.ms365TenantSettingsLoad !== 'function') return null;
    let settings = null;
    try {
        settings = window.ms365TenantSettingsLoad();
    } catch {
        return null;
    }
    const classes = settings && Array.isArray(settings.classes) ? settings.classes : [];
    if (!classes.length) return null;

    const names = [];
    const codes = [];
    for (let i = 0; i < classes.length; i++) {
        const c = classes[i] || {};
        const name = c.name ? String(c.name).trim() : '';
        const code = c.code ? String(c.code).trim() : '';
        if (name) names.push(name.toLowerCase());
        if (code) codes.push(code);
    }

    const codeRes = codes
        .map(function (code) {
            const esc = escapeRe(code);
            // typ. Klassenkürzel wie "1A", "2BK" sollen als eigenständiges Token matchen
            return new RegExp('(^|\\s)' + esc + '(\\s|$)', 'i');
        })
        .slice(0, 250); // harte Bremse gegen absurd lange Listen

    return function (displayName) {
        const dn = String(displayName || '').trim();
        if (!dn) return false;
        const dnLow = dn.toLowerCase();
        for (let i = 0; i < names.length; i++) {
            if (dnLow.indexOf(names[i]) !== -1) return true;
        }
        for (let j = 0; j < codeRes.length; j++) {
            if (codeRes[j].test(dn)) return true;
        }
        return false;
    };
}

function getTenantClassesList() {
    if (typeof window.ms365TenantSettingsLoad !== 'function') return [];
    try {
        const s = window.ms365TenantSettingsLoad();
        const classes = s && Array.isArray(s.classes) ? s.classes : [];
        return classes
            .map(function (c) {
                return {
                    code: c && c.code ? String(c.code).trim() : '',
                    name: c && c.name ? String(c.name).trim() : ''
                };
            })
            .filter(function (c) {
                return !!(c.code || c.name);
            });
    } catch {
        return [];
    }
}

function getTenantClassesFull() {
    if (typeof window.ms365TenantSettingsLoad !== 'function') return [];
    try {
        const s = window.ms365TenantSettingsLoad();
        const classes = s && Array.isArray(s.classes) ? s.classes : [];
        return classes
            .map(function (c) {
                return {
                    code: c && c.code ? String(c.code).trim() : '',
                    year: c && c.year ? String(c.year).trim() : '',
                    name: c && c.name ? String(c.name).trim() : '',
                    headName: c && c.headName ? String(c.headName).trim() : '',
                    headEmail: c && c.headEmail ? String(c.headEmail).trim() : ''
                };
            })
            .filter(function (c) {
                return !!(c.code || c.name || c.year || c.headName || c.headEmail);
            });
    } catch {
        return [];
    }
}

function resolveSchoolDomainNoAt() {
    try {
        if (typeof window.ms365GetSchoolDomainNoAt === 'function') {
            const d = String(window.ms365GetSchoolDomainNoAt() || '').trim();
            return d || 'ms365.schule';
        }
    } catch {
        // ignore
    }
    try {
        const s = typeof window.ms365TenantSettingsLoad === 'function' ? window.ms365TenantSettingsLoad() : null;
        const d = s && s.domain ? String(s.domain).trim().replace(/^@+/, '') : '';
        if (d) return d;
    } catch {
        // ignore
    }
    return 'ms365.schule';
}

function findTenantClassForDisplayName(displayName, tenantClassesFull) {
    const dn = String(displayName || '').trim();
    if (!dn) return null;
    const dnLow = dn.toLowerCase();
    const list = Array.isArray(tenantClassesFull) ? tenantClassesFull : [];

    // 1) exakter Name-Contains
    for (let i = 0; i < list.length; i++) {
        const c = list[i];
        const n = c && c.name ? String(c.name).trim().toLowerCase() : '';
        if (n && dnLow.indexOf(n) !== -1) return c;
    }

    // 2) Code als Token
    for (let j = 0; j < list.length; j++) {
        const c = list[j];
        const code = c && c.code ? String(c.code).trim() : '';
        if (!code) continue;
        const re = new RegExp('(^|\\s)' + escapeRe(code) + '(\\s|$)', 'i');
        if (re.test(dn)) return c;
    }
    return null;
}

function renderLoadedGroupsTable() {
    const tbody = document.getElementById('kuLoadedGroupsBody');
    if (!tbody) return;
    tbody.replaceChildren();

    const list = getFilteredGroups();
    const tenantClassesFull = getTenantClassesFull();
    const summary = document.getElementById('kuLoadedGroupsSummary');
    if (summary) {
        if (showTenantClassesInStep1) {
            summary.textContent = tenantClassesFull.length
                ? 'Klassen aus Schul‑Einstellungen: ' +
                  tenantClassesFull.length +
                  '. Für ID/Besitzer bitte zusätzlich „Gruppen laden“ ausführen.'
                : 'Keine Klassen in den Schul‑Einstellungen gefunden.';
        } else {
            summary.textContent =
                'Vorschau (Schritt 1): ' +
                list.length +
                ' Gruppe(n) sichtbar. Besitzer werden ggf. nachgeladen.';
        }
    }

    function td(text) {
        const c = document.createElement('td');
        c.style.fontSize = '0.9em';
        c.style.wordBreak = 'break-word';
        c.textContent = text;
        return c;
    }

    // Wenn der Button "Klassen aus Einstellungen laden" genutzt wurde,
    // zeige die Klassenliste an (IDs/Besitzer kommen erst nach dem Laden).
    if (showTenantClassesInStep1) {
        const domain = resolveSchoolDomainNoAt();
        for (let i = 0; i < tenantClassesFull.length; i++) {
            const c = tenantClassesFull[i];
            const tr = document.createElement('tr');
            const displayName = c.name || (c.code ? 'Klasse ' + c.code : '') || '—';
            const mailNick = displayName && displayName !== '—' ? deriveMailNickname(displayName) : '';
            const mailPreview = mailNick ? mailNick + '@' + domain : '—';

            tr.appendChild(td(c.code || '—'));
            tr.appendChild(td(c.year || '—'));
            tr.appendChild(td(c.name || '—'));
            tr.appendChild(td(displayName || '—'));
            tr.appendChild(td(mailNick || '—'));
            tr.appendChild(td(mailPreview || '—'));
            tr.appendChild(td(c.headName || '—'));
            tr.appendChild(td(c.headEmail || '—'));
            tbody.appendChild(tr);
        }
        if (!tenantClassesFull.length) {
            const tr = document.createElement('tr');
            const c = document.createElement('td');
            c.colSpan = 8;
            c.style.color = '#6c757d';
            c.textContent = 'Keine Klassen – bitte in den Schul‑Einstellungen pflegen.';
            tr.appendChild(c);
            tbody.appendChild(tr);
        }
        return; // bewusst keine Gruppen anzeigen in diesem Modus
    }

    // Normal: zeige gefilterte Gruppen
    const domain = resolveSchoolDomainNoAt();
    for (let i = 0; i < list.length; i++) {
        const g = list[i];
        const tr = document.createElement('tr');
        const tc = findTenantClassForDisplayName(g.displayName || '', tenantClassesFull);
        const cached = ownersCache[g.id];
        const ownersText = cached
            ? cached.loading
                ? 'Lade …'
                : cached.err
                  ? 'Fehler: ' + cached.err
                  : cached.ownersText || '—'
            : '—';
        const displayName = g.displayName || '—';
        const mailNick = g.mailNickname || (displayName && displayName !== '—' ? deriveMailNickname(displayName) : '');
        const mailPreview = g.mail || (mailNick ? mailNick + '@' + domain : '—');

        tr.appendChild(td((tc && tc.code) || '—'));
        tr.appendChild(td((tc && tc.year) || '—'));
        tr.appendChild(td((tc && tc.name) || '—'));
        tr.appendChild(td(displayName));
        tr.appendChild(td(mailNick || '—'));
        tr.appendChild(td(mailPreview || '—'));
        // Besitzer aus den Schul‑Einstellungen (Klassenvorstand); Graph-Owners bleiben in ownersText als Fallback sichtbar
        tr.appendChild(td((tc && tc.headName) || '—'));
        tr.appendChild(td((tc && tc.headEmail) || ownersText || '—'));
        tbody.appendChild(tr);
    }

    if (!list.length) {
        const tr = document.createElement('tr');
        const c = document.createElement('td');
        c.colSpan = 8;
        c.style.color = '#6c757d';
        c.textContent = loadedGroups.length
            ? 'Keine Gruppen – bitte Filter prüfen.'
            : 'Keine Gruppen – bitte zuerst „Gruppen laden“ ausführen.';
        tr.appendChild(c);
        tbody.appendChild(tr);
    }
}

async function withConcurrency(items, limit, fn) {
    const q = items.slice();
    const workers = [];
    const cap = Math.max(1, Math.min(limit || 1, 16));
    for (let i = 0; i < cap; i++) {
        workers.push(
            (async function () {
                while (q.length) {
                    const it = q.shift();
                    try {
                        await fn(it);
                    } catch {
                        // Fehler pro Item werden im fn gehandhabt
                    }
                }
            })()
        );
    }
    await Promise.all(workers);
}

async function fetchOwnersForGroup(token, groupId) {
    if (!groupId) return;
    const existing = ownersCache[groupId];
    if (existing && (existing.loading || existing.ownersText || existing.err)) return;
    ownersCache[groupId] = { ownersText: '', loading: true, err: '' };
    renderLoadedGroupsTable();
    try {
        const path =
            '/groups/' +
            encodeURIComponent(groupId) +
            "/owners?$select=id,displayName,mail,userPrincipalName&$top=999";
        const owners = await fetchAllPages(token, path);
        const names = (owners || [])
            .map(function (o) {
                const dn = o && o.displayName ? String(o.displayName).trim() : '';
                const mail = o && (o.mail || o.userPrincipalName) ? String(o.mail || o.userPrincipalName).trim() : '';
                if (dn && mail) return dn + ' (' + mail + ')';
                return dn || mail || '';
            })
            .filter(Boolean);
        ownersCache[groupId] = { ownersText: names.length ? names.join(', ') : '—', loading: false, err: '' };
    } catch (e) {
        ownersCache[groupId] = {
            ownersText: '',
            loading: false,
            err: String(e && e.message ? e.message : e)
        };
    }
    renderLoadedGroupsTable();
}

async function ensureOwnersForVisibleGroups(token) {
    const vis = getFilteredGroups();
    const ids = vis.map(function (g) {
        return g.id;
    });
    await withConcurrency(ids, 6, async function (id) {
        await fetchOwnersForGroup(token, id);
    });
}

async function checkMailNicknameConflict(token, mailNickname, excludeId) {
    if (!mailNickname) return 'Leerer Mail-Nickname.';
    const filter = "mailNickname eq '" + odataEscape(mailNickname) + "'";
    const path = '/groups?$filter=' + encodeURIComponent(filter) + '&$select=id,displayName';
    const data = await graphJson('GET', path, token, undefined);
    const v = data.value || [];
    if (!v.length) return null;
    if (v.length === 1 && v[0].id === excludeId) return null;
    return (
        'Mail-Nickname bereits vergeben (Gruppe: ' +
        (v[0].displayName || v[0].id) +
        ').'
    );
}

async function buildPreview() {
    const prefixInp = document.getElementById('kuPrefix');
    const prefix = prefixInp && prefixInp.value ? prefixInp.value : 'Klasse';

    // Testmodus: Vorschau aus Klassen aus den Schul‑Einstellungen erlauben (ohne Graph / ohne echte Gruppen-IDs)
    if (!loadedGroups.length && showTenantClassesInStep1) {
        const tenantClasses = getTenantClassesFull();
        previewRows = [];
        for (let i = 0; i < tenantClasses.length; i++) {
            const c = tenantClasses[i] || {};
            const oldDn = c.name || (c.code ? 'Klasse ' + c.code : '') || '';
            const newDnRaw = computeNewDisplayNamePlusOne(oldDn, prefix);
            const newDn = newDnRaw ? applyReplaceRules(newDnRaw) : null;
            if (!newDn) {
                previewRows.push({
                    id: 'TEST-' + (c.code || String(i + 1)),
                    displayNameOld: oldDn || '—',
                    displayNameNew: '—',
                    mailNicknameOld: '—',
                    mailNicknameNew: '—',
                    mailOld: '',
                    hint: 'Regel trifft nicht zu (Format passt nicht).',
                    ok: false
                });
                continue;
            }

            const willChangeDn = newDn !== oldDn;
            const ok = !!willChangeDn;
            previewRows.push({
                id: 'TEST-' + (c.code || String(i + 1)),
                displayNameOld: oldDn || '—',
                displayNameNew: newDn,
                mailNicknameOld: '—',
                mailNicknameNew: '—',
                mailOld: '',
                hint: ok ? '' : 'Keine Änderung.',
                ok: ok
            });
        }

        const okCount = previewRows.filter(function (r) {
            return r.ok;
        }).length;
        const summary = document.getElementById('kuPreviewSummary');
        if (summary) {
            summary.textContent =
                'Testdaten (Schul‑Klassen): ' +
                tenantClasses.length +
                '. Davon würden ' +
                okCount +
                ' umbenannt werden.';
        }

        renderPreviewTable();
        const next3 = document.getElementById('kuNext3');
        if (next3) next3.disabled = okCount === 0;
        return true;
    }

    const token = await getGraphToken();
    const groups = getFilteredGroups();
    previewRows = [];

    for (let i = 0; i < groups.length; i++) {
        const g = groups[i];
        const oldDn = g.displayName || '';
        const newDnRaw = computeNewDisplayNamePlusOne(oldDn, prefix);
        const newDn = newDnRaw ? applyReplaceRules(newDnRaw) : null;
        const mailOld = g.mail || '';
        const nickOld = g.mailNickname || '';

        if (!newDn) {
            previewRows.push({
                id: g.id,
                displayNameOld: oldDn,
                displayNameNew: '—',
                mailNicknameOld: nickOld,
                mailNicknameNew: '—',
                mailOld: mailOld,
                hint: 'Regel trifft nicht zu (Format passt nicht).',
                ok: false
            });
            continue;
        }

        let nickNew = nickOld;
        let hint = '';

        const willChangeDn = newDn !== oldDn;
        const willChangeNick = false;

        if (!hint && !willChangeDn && !willChangeNick) {
            hint = 'Keine Änderung (Name und Mail-Nickname bereits passend).';
        }

        const ok = !hint && (willChangeDn || willChangeNick);

        previewRows.push({
            id: g.id,
            displayNameOld: oldDn,
            displayNameNew: newDn,
            mailNicknameOld: nickOld,
            mailNicknameNew: nickNew,
            mailOld: mailOld,
            hint: hint,
            ok: ok
        });
    }

    const okCount = previewRows.filter(function (r) {
        return r.ok;
    }).length;
    const summary = document.getElementById('kuPreviewSummary');
    if (summary) {
        summary.textContent =
            'Gefilterte Gruppen: ' +
            groups.length +
            '. Davon können ' +
            okCount +
            ' umbenannt werden (ohne Konflikt).';
    }

    renderPreviewTable();
    const next3 = document.getElementById('kuNext3');
    if (next3) next3.disabled = okCount === 0;
    return true;
}

function renderPreviewTable() {
    const tbody = document.getElementById('kuPreviewBody');
    if (!tbody) return;
    tbody.replaceChildren();
    const next3 = document.getElementById('kuNext3');

    function recomputePreviewSummaryAndNext() {
        const groupsCount = previewRows.length;
        const okCount = previewRows.filter(function (r) {
            return r.ok;
        }).length;
        const summary = document.getElementById('kuPreviewSummary');
        if (summary) {
            // Wenn wir im Testmodus sind, kommt die Summary schon aus buildPreview(); trotzdem konsistent halten.
            summary.textContent =
                (loadedGroups.length
                    ? 'Gefilterte Gruppen: ' + groupsCount + '. '
                    : 'Testdaten (Schul‑Klassen): ' + groupsCount + '. ') +
                'Davon können ' +
                okCount +
                ' umbenannt werden.';
        }
        if (next3) next3.disabled = okCount === 0;
    }

    function startCellEdit(tdEl, initialValue, onCommit) {
        const prevText = String(initialValue ?? '');
        const input = document.createElement('input');
        input.type = 'text';
        input.value = prevText;
        input.style.width = '100%';
        input.style.boxSizing = 'border-box';
        input.style.padding = '8px 10px';
        input.style.border = '1px solid rgba(94, 114, 228, 0.6)';
        input.style.borderRadius = '10px';
        input.style.font = 'inherit';
        tdEl.replaceChildren(input);
        input.focus();
        input.select();

        function commit(cancelled) {
            const next = cancelled ? prevText : String(input.value ?? '');
            onCommit(next, { cancelled: !!cancelled });
        }

        input.addEventListener('keydown', function (e) {
            if (e.key === 'Enter') {
                e.preventDefault();
                commit(false);
            } else if (e.key === 'Escape') {
                e.preventDefault();
                commit(true);
            }
        });
        input.addEventListener('blur', function () {
            commit(false);
        });
    }

    function normalizeName(s) {
        return String(s || '').trim();
    }

    for (let i = 0; i < previewRows.length; i++) {
        const r = previewRows[i];
        const tr = document.createElement('tr');
        if (r.hint) tr.style.background = 'rgba(255, 193, 7, 0.12)';
        function td(text) {
            const c = document.createElement('td');
            c.style.fontSize = '0.9em';
            c.style.wordBreak = 'break-word';
            c.textContent = text;
            return c;
        }

        const tdOld = td(r.displayNameOld || '–');

        // Editierbar: Anzeigename (neu) per Doppelklick
        const tdNew = td(r.displayNameNew || '–');
        tdNew.title = 'Doppelklick zum Bearbeiten';
        tdNew.style.cursor = 'text';
        tdNew.addEventListener('dblclick', function () {
            startCellEdit(tdNew, r.displayNameNew || '', function (nextRaw, meta) {
                const next = normalizeName(nextRaw);
                const prev = r.displayNameNew || '';
                r.displayNameNew = meta && meta.cancelled ? prev : next;

                // neu bewerten (MailNickname bleibt immer unverändert)
                const oldDn = normalizeName(r.displayNameOld || '');
                const newDn = normalizeName(r.displayNameNew || '');
                const willChangeDn = !!newDn && newDn !== '—' && newDn !== oldDn;
                if (!newDn || newDn === '—') {
                    r.hint = 'Neuer Anzeigename ist leer.';
                    r.ok = false;
                } else if (!willChangeDn) {
                    r.hint = 'Keine Änderung.';
                    r.ok = false;
                } else {
                    r.hint = '';
                    r.ok = true;
                }

                renderPreviewTable();
            });
        });

        const tdHint = td(r.hint || (r.ok ? 'OK' : ''));

        tr.appendChild(tdOld);
        tr.appendChild(tdNew);
        tr.appendChild(tdHint);
        tbody.appendChild(tr);
    }
    if (!previewRows.length) {
        const tr = document.createElement('tr');
        const td = document.createElement('td');
        td.colSpan = 3;
        td.style.color = '#6c757d';
        td.textContent = 'Keine Gruppen – Schritt 1 prüfen.';
        tr.appendChild(td);
        tbody.appendChild(tr);
    }

    recomputePreviewSummaryAndNext();
}

function renderReplaceRulesTable() {
    const tbody = document.getElementById('kuReplaceRulesBody');
    if (!tbody) return;
    tbody.replaceChildren();

    function td(text) {
        const c = document.createElement('td');
        c.style.fontSize = '0.9em';
        c.style.wordBreak = 'break-word';
        c.textContent = text;
        return c;
    }

    if (!replaceRules.length) {
        const tr = document.createElement('tr');
        const c = document.createElement('td');
        c.colSpan = 3;
        c.style.color = '#6c757d';
        c.textContent = 'Keine Regeln – optional hinzufügen.';
        tr.appendChild(c);
        tbody.appendChild(tr);
        return;
    }

    for (let i = 0; i < replaceRules.length; i++) {
        const r = replaceRules[i];
        const tr = document.createElement('tr');
        tr.appendChild(td(r.find || '—'));
        tr.appendChild(td(r.replaceWith || ''));
        const action = document.createElement('td');
        action.style.whiteSpace = 'nowrap';
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'btn btn-small';
        btn.textContent = '✕';
        btn.addEventListener('click', function () {
            replaceRules.splice(i, 1);
            renderReplaceRulesTable();
            renderRulePreview();
        });
        action.appendChild(btn);
        tr.appendChild(action);
        tbody.appendChild(tr);
    }
}

function renderRulePreview() {
    const tbody = document.getElementById('kuRulePreviewBody');
    if (!tbody) return;
    tbody.replaceChildren();

    const prefixInp = document.getElementById('kuPrefix');
    const prefix = prefixInp && prefixInp.value ? prefixInp.value : 'Klasse';

    /** @type {{ old: string, neu: string, hint: string, ok: boolean }[]} */
    const rows = [];

    if (loadedGroups.length) {
        const groups = getFilteredGroups();
        const sample = groups.slice(0, 25);
        for (let i = 0; i < sample.length; i++) {
            const g = sample[i];
            const oldDn = String(g.displayName || '');
            const newDnRaw = computeNewDisplayNamePlusOne(oldDn, prefix);
            const newDn = newDnRaw ? applyReplaceRules(newDnRaw) : '';
            if (!newDn) {
                rows.push({ old: oldDn || '—', neu: '—', hint: 'Regel trifft nicht zu.', ok: false });
                continue;
            }
            const ok = newDn !== oldDn;
            rows.push({ old: oldDn || '—', neu: newDn, hint: ok ? 'OK' : 'Keine Änderung.', ok: ok });
        }
    } else if (showTenantClassesInStep1) {
        const tenantClasses = getTenantClassesFull().slice(0, 25);
        for (let i = 0; i < tenantClasses.length; i++) {
            const c = tenantClasses[i] || {};
            const oldDn = c.name || (c.code ? 'Klasse ' + c.code : '') || '';
            const newDnRaw = computeNewDisplayNamePlusOne(oldDn, prefix);
            const newDn = newDnRaw ? applyReplaceRules(newDnRaw) : '';
            if (!newDn) {
                rows.push({ old: oldDn || '—', neu: '—', hint: 'Regel trifft nicht zu.', ok: false });
                continue;
            }
            const ok = newDn !== oldDn;
            rows.push({ old: oldDn || '—', neu: newDn, hint: ok ? 'OK' : 'Keine Änderung.', ok: ok });
        }
    }

    function td(text) {
        const c = document.createElement('td');
        c.style.fontSize = '0.9em';
        c.style.wordBreak = 'break-word';
        c.textContent = text;
        return c;
    }

    for (let i = 0; i < rows.length; i++) {
        const r = rows[i];
        const tr = document.createElement('tr');
        tr.appendChild(td(r.old));
        tr.appendChild(td(r.neu));
        tr.appendChild(td(r.hint));
        tbody.appendChild(tr);
    }

    if (!rows.length) {
        const tr = document.createElement('tr');
        const c = document.createElement('td');
        c.colSpan = 3;
        c.style.color = '#6c757d';
        c.textContent = 'Keine Daten – bitte zuerst Schritt 1 ausführen.';
        tr.appendChild(c);
        tbody.appendChild(tr);
    }

    const summary = document.getElementById('kuRulePreviewSummary');
    if (summary) {
        const okCount = rows.filter(function (x) {
            return x.ok;
        }).length;
        summary.textContent = 'Vorschau zeigt bis zu 25 Zeilen. OK: ' + okCount + ' / ' + rows.length + '.';
    }
}

function goToStep(step) {
    const n = Number(step);
    const contents = document.querySelectorAll('.ku-step-content');
    for (let i = 0; i < contents.length; i++) {
        const el = contents[i];
        const s = el.getAttribute('data-ku-step');
        if (String(s) === String(n)) el.classList.add('active');
        else el.classList.remove('active');
    }
    const steps = document.querySelectorAll('.ku-steps .step');
    for (let j = 0; j < steps.length; j++) {
        const st = steps[j];
        const s = st.getAttribute('data-ku-step');
        if (String(s) === String(n)) st.classList.add('active');
        else st.classList.remove('active');
    }
    const bar = document.getElementById('kuStepsBar');
    if (bar && typeof window.ms365ApplyStepProgress === 'function') {
        window.ms365ApplyStepProgress(bar, n, STEP_ORDER);
    }
}

async function loadGroups() {
    const status = document.getElementById('kuLoadStatus');
    const next1 = document.getElementById('kuNext1');
    if (next1) next1.disabled = true;
    if (status) status.textContent = 'Lade …';
    try {
        const token = await getGraphToken();
        showTenantClassesInStep1 = false;
        const filter = encodeURIComponent("groupTypes/any(c:c eq 'Unified')");
        const initial =
            '/groups?$filter=' + filter + '&$select=' + encodeURIComponent(GROUP_SELECT) + '&$top=999';
        loadedGroups = await fetchAllPages(token, initial);
        ownersCache = Object.create(null);
        loadedGroups.sort(function (a, b) {
            const an = a && a.displayName ? String(a.displayName) : '';
            const bn = b && b.displayName ? String(b.displayName) : '';
            return an.localeCompare(bn, 'de');
        });
        const n = getFilteredGroups().length;
        if (status) {
            status.textContent =
                'Gesamt ' +
                loadedGroups.length +
                ' einheitliche Gruppe(n). Nach Filter: ' +
                n +
                ' sichtbar für die nächsten Schritte.';
        }
        if (next1) next1.disabled = loadedGroups.length === 0;
        renderLoadedGroupsTable();
        ensureOwnersForVisibleGroups(token);
        toast('Gruppen geladen.');
    } catch (e) {
        const msg = String(e && e.message ? e.message : e);
        if (status) status.textContent = 'Fehler: ' + msg;
        const summary = document.getElementById('kuLoadedGroupsSummary');
        if (summary) summary.textContent = 'Keine Gruppen geladen (siehe Fehlermeldung oben).';
        renderLoadedGroupsTable();
        toast(msg);
    }
}

function loadTenantClassesIntoStep1Table() {
    showTenantClassesInStep1 = true;
    renderLoadedGroupsTable();
    const next1 = document.getElementById('kuNext1');
    if (next1) next1.disabled = false;
    const status = document.getElementById('kuLoadStatus');
    if (status) status.textContent = 'Testdaten geladen (Klassen aus Schul‑Einstellungen).';
}

async function runRename() {
    clearLog();
    const toPatch = previewRows.filter(function (r) {
        return r.ok;
    });
    if (!toPatch.length) {
        toast('Keine gültigen Zeilen zum Umbenennen.');
        return;
    }
    appendLog('Start: ' + toPatch.length + ' Gruppe(n).');
    let ok = 0;
    let fail = 0;
    try {
        const token = await getGraphToken();
        for (let i = 0; i < toPatch.length; i++) {
            const r = toPatch[i];
            const body = {};
            if (r.displayNameNew && r.displayNameNew !== r.displayNameOld) {
                body.displayName = r.displayNameNew;
            }
            if (!Object.keys(body).length) {
                appendLog('Übersprungen (keine Felder): ' + r.displayNameOld, 'warn');
                continue;
            }
            try {
                await graphJson('PATCH', '/groups/' + encodeURIComponent(r.id), token, body);
                appendLog('OK: ' + r.displayNameOld + ' → ' + r.displayNameNew, 'ok');
                ok++;
            } catch (e) {
                fail++;
                appendLog(
                    'Fehler ' +
                        r.displayNameOld +
                        ': ' +
                        (e && e.message ? e.message : e),
                    'err'
                );
            }
        }
        appendLog('Fertig. Erfolg: ' + ok + ', Fehler: ' + fail + '.', fail ? 'warn' : 'ok');
        toast('Umbenennung abgeschlossen.');
    } catch (e) {
        appendLog(String(e && e.message ? e.message : e), 'err');
    }
}

async function onLogin() {
    try {
        await getGraphToken();
        toast('Angemeldet.');
    } catch (e) {
        toast(String(e && e.message ? e.message : e));
    }
}

function bind() {
    const btnLogin = document.getElementById('kuBtnLogin');
    const btnLoad = document.getElementById('kuBtnLoad');
    const btnLoadTenantClasses = document.getElementById('kuBtnLoadTenantClasses');
    const next1 = document.getElementById('kuNext1');
    const next2 = document.getElementById('kuNext2');
    const next3 = document.getElementById('kuNext3');
    const back2 = document.getElementById('kuBack2');
    const back3 = document.getElementById('kuBack3');
    const back4 = document.getElementById('kuBack4');
    const btnRun = document.getElementById('kuBtnRun');
    const filt = document.getElementById('kuFilterContains');
    const onlyTeams = document.getElementById('kuOnlyTeams');
    const prefixInp = document.getElementById('kuPrefix');
    const replFind = document.getElementById('kuReplaceFind');
    const replWith = document.getElementById('kuReplaceWith');
    const replAdd = document.getElementById('kuReplaceAdd');

    if (btnLogin) btnLogin.addEventListener('click', () => onLogin());
    if (btnLoad) btnLoad.addEventListener('click', () => loadGroups());
    if (btnLoadTenantClasses) btnLoadTenantClasses.addEventListener('click', () => loadTenantClassesIntoStep1Table());
    function refreshFilterStatus() {
        const status = document.getElementById('kuLoadStatus');
        const next1b = document.getElementById('kuNext1');
        const n = getFilteredGroups().length;
        if (status) {
            status.textContent =
                'Gesamt ' +
                loadedGroups.length +
                ' Gruppe(n). Nach Filter: ' +
                n +
                ' sichtbar.';
        }
        if (next1b) next1b.disabled = false;
        renderLoadedGroupsTable();
        (async function () {
            try {
                const token = await getGraphToken();
                ensureOwnersForVisibleGroups(token);
            } catch {
                // Ignorieren: Besitzer-Vorschau ist optional
            }
        })();
    }

    if (filt)
        filt.addEventListener('input', function () {
            if (loadedGroups.length) refreshFilterStatus();
        });
    if (onlyTeams)
        onlyTeams.addEventListener('change', function () {
            if (loadedGroups.length) refreshFilterStatus();
        });

    if (prefixInp) {
        prefixInp.addEventListener('input', function () {
            renderRulePreview();
        });
    }
    if (replAdd) {
        replAdd.addEventListener('click', function () {
            const f = replFind && replFind.value ? String(replFind.value) : '';
            const w = replWith && replWith.value ? String(replWith.value) : '';
            if (!f.trim()) {
                toast('Bitte „Suchen“ ausfüllen.');
                return;
            }
            replaceRules.push({ find: f, replaceWith: w });
            if (replFind) replFind.value = '';
            if (replWith) replWith.value = '';
            renderReplaceRulesTable();
            renderRulePreview();
        });
    }

    if (next1) {
        next1.addEventListener('click', function () {
            if (!loadedGroups.length && !showTenantClassesInStep1) {
                toast('Bitte zuerst „Gruppen laden“ oder „Klassen aus Einstellungen laden“.');
                return;
            }
            goToStep(2);
            renderReplaceRulesTable();
            renderRulePreview();
        });
    }
    if (back2) back2.addEventListener('click', () => goToStep(1));
    if (next2) {
        next2.addEventListener('click', async function () {
            try {
                const ok = await buildPreview();
                if (ok) goToStep(3);
            } catch (e) {
                toast(String(e && e.message ? e.message : e));
            }
        });
    }
    if (back3) back3.addEventListener('click', () => goToStep(2));
    if (next3) {
        next3.addEventListener('click', function () {
            goToStep(4);
        });
    }
    if (back4) back4.addEventListener('click', () => goToStep(3));
    if (btnRun) btnRun.addEventListener('click', () => runRename());

    goToStep(1);
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', bind);
} else {
    bind();
}