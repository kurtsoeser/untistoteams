(function () {
    'use strict';

    const STORAGE_KEY = 'ms365-schulstruktur-sync-v1';
    const STORAGE_TENANT_CACHE_KEY = 'ms365-schulstruktur-tenant-cache-v1';

    const GRAPH_SCOPES_TENANT_READ = [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/Group.Read.All'
    ];

    function safeJsonParse(s) {
        try {
            return JSON.parse(String(s));
        } catch {
            return null;
        }
    }

    function uid() {
        // Kurze, stabile ID für lokale Mock-Daten (nicht kryptografisch)
        return 'id-' + Math.random().toString(16).slice(2) + '-' + Date.now().toString(16);
    }

    function normStr(v) {
        return String(v ?? '').trim();
    }

    function loadState() {
        try {
            const raw = localStorage.getItem(STORAGE_KEY);
            const obj = raw ? safeJsonParse(raw) : null;
            const rows = obj && Array.isArray(obj.rows) ? obj.rows : [];
            return { rows };
        } catch {
            return { rows: [] };
        }
    }

    function saveState(state) {
        const rows = state && Array.isArray(state.rows) ? state.rows : [];
        try {
            localStorage.setItem(STORAGE_KEY, JSON.stringify({ rows }));
        } catch {
            // ignore
        }
        return { rows };
    }

    function loadTenantCache() {
        try {
            const raw = localStorage.getItem(STORAGE_TENANT_CACHE_KEY);
            const obj = raw ? safeJsonParse(raw) : null;
            const rows = obj && Array.isArray(obj.rows) ? obj.rows : [];
            return { rows, loadedAt: obj && obj.loadedAt ? String(obj.loadedAt) : '' };
        } catch {
            return { rows: [], loadedAt: '' };
        }
    }

    function saveTenantCache(rows) {
        const out = Array.isArray(rows) ? rows : [];
        try {
            localStorage.setItem(
                STORAGE_TENANT_CACHE_KEY,
                JSON.stringify({ rows: out, loadedAt: new Date().toISOString() })
            );
        } catch {
            // ignore
        }
    }

    function buildDemoRows() {
        const schuljahr = String(new Date().getFullYear()) + '/' + String(new Date().getFullYear() + 1).slice(2);
        const jg1 = { id: uid(), parentId: '', typ: 'Jahrgang', bezeichnung: 'Jahrgang 1', schuljahr, status: 'Aktiv', syncStatus: 'Ausstehend', letzteFehlermeldung: '' };
        const jg2 = { id: uid(), parentId: '', typ: 'Jahrgang', bezeichnung: 'Jahrgang 2', schuljahr, status: 'Aktiv', syncStatus: 'Ok', letzteFehlermeldung: '' };
        const k1a = { id: uid(), parentId: jg1.id, typ: 'Klasse', bezeichnung: '1A', schuljahr, status: 'Aktiv', syncStatus: 'Abweichung', letzteFehlermeldung: 'Mitgliedschaft weicht ab (Mock).' };
        const k1b = { id: uid(), parentId: jg1.id, typ: 'Klasse', bezeichnung: '1B', schuljahr, status: 'Aktiv', syncStatus: 'Ok', letzteFehlermeldung: '' };
        const k2a = { id: uid(), parentId: jg2.id, typ: 'Klasse', bezeichnung: '2A', schuljahr, status: 'Aktiv', syncStatus: 'Fehler', letzteFehlermeldung: 'Team konnte nicht bereitgestellt werden (Mock).' };
        const ar1 = { id: uid(), parentId: '', typ: 'Arbeitsgemeinschaft', bezeichnung: 'ARGE Robotik', schuljahr, status: 'Aktiv', syncStatus: 'Ok', letzteFehlermeldung: '' };
        const ar2 = { id: uid(), parentId: '', typ: 'Arbeitsgemeinschaft', bezeichnung: 'ARGE Chor', schuljahr, status: 'Inaktiv', syncStatus: 'Ausstehend', letzteFehlermeldung: '' };
        return [jg1, jg2, k1a, k1b, k2a, ar1, ar2];
    }

    function byId(rows) {
        const map = new Map();
        for (let i = 0; i < rows.length; i++) {
            const r = rows[i];
            if (r && r.id) map.set(String(r.id), r);
        }
        return map;
    }

    function compareDe(a, b) {
        return String(a || '').localeCompare(String(b || ''), 'de', { sensitivity: 'base' });
    }

    function getEl(id) {
        return document.getElementById(id);
    }

    function toast(msg) {
        if (typeof window.ms365ShowToast === 'function') {
            window.ms365ShowToast(msg);
        } else {
            window.alert(msg);
        }
    }

    function pillClass(syncStatus) {
        const s = String(syncStatus || '');
        if (s === 'Ok') return 'ok';
        if (s === 'Abweichung') return 'warn';
        if (s === 'Fehler') return 'err';
        return '';
    }

    function computeStats(rows) {
        const total = rows.length;
        const aktiv = rows.filter((r) => r && r.status === 'Aktiv').length;
        const abw = rows.filter((r) => r && r.syncStatus === 'Abweichung').length;
        const err = rows.filter((r) => r && r.syncStatus === 'Fehler').length;
        return { total, aktiv, abw, err };
    }

    function computeTenantStats(rows) {
        const total = rows.length;
        const teams = rows.filter((r) => r && r.typ === 'Team').length;
        const groups = rows.filter((r) => r && r.typ === 'Gruppe').length;
        return { total, teams, groups };
    }

    function setModeHint(mode, tenantLoadedAt) {
        const el = getEl('ssModeHint');
        if (!el) return;
        if (mode === 'tenant') {
            el.textContent =
                'Tenant‑Inventar: M365‑Gruppen/Teams werden live per Graph eingelesen (read‑only).' +
                (tenantLoadedAt ? ' Letztes Einlesen: ' + new Date(tenantLoadedAt).toLocaleString() : '');
        } else {
            el.textContent =
                'Hinweis: Das ist die v1‑Mock‑Ansicht (lokal im Browser). Nächster Schritt: SharePoint‑Listen als Datenquelle.';
        }
    }

    function getFilterState() {
        const schuljahr = normStr(getEl('ssFilterSchuljahr')?.value);
        const typ = normStr(getEl('ssFilterTyp')?.value);
        const text = normStr(getEl('ssFilterText')?.value).toLowerCase();
        return { schuljahr, typ, text };
    }

    function applyFilters(rows, mode) {
        const f = getFilterState();
        return rows.filter((r) => {
            if (!r) return false;
            if (mode !== 'tenant') {
                if (f.schuljahr && String(r.schuljahr || '') !== f.schuljahr) return false;
            }
            if (f.typ && String(r.typ || '') !== f.typ) return false;
            if (f.text) {
                const hay = (String(r.bezeichnung || '') + ' ' + String(r.typ || '') + ' ' + String(r.schuljahr || '')).toLowerCase();
                if (hay.indexOf(f.text) === -1) return false;
            }
            return true;
        });
    }

    function buildTreeOrder(rows) {
        const map = byId(rows);
        const children = new Map();
        for (const r of rows) {
            const pid = String(r.parentId || '');
            if (!children.has(pid)) children.set(pid, []);
            children.get(pid).push(r);
        }
        // sort children lists
        for (const [k, list] of children.entries()) {
            list.sort((a, b) => {
                // Jahrgang vor Klasse vor ARGE vor Gruppe (nur optische Sortierung)
                const rank = (x) => (x.typ === 'Jahrgang' ? 1 : x.typ === 'Klasse' ? 2 : x.typ === 'Arbeitsgemeinschaft' ? 3 : 4);
                const ra = rank(a);
                const rb = rank(b);
                if (ra !== rb) return ra - rb;
                return compareDe(a.bezeichnung, b.bezeichnung);
            });
            children.set(k, list);
        }

        const out = [];
        function walk(parentId, depth) {
            const list = children.get(String(parentId || '')) || [];
            for (const r of list) {
                out.push({ r, depth });
                // Schutz gegen versehentliche Zyklen
                if (depth < 10 && map.has(String(r.id))) walk(r.id, depth + 1);
            }
        }
        walk('', 0);
        return out;
    }

    function renderFilters(rows, mode) {
        const sel = getEl('ssFilterSchuljahr');
        if (sel) {
            const prev = sel.value || '';
            const years = Array.from(new Set(rows.map((r) => String(r.schuljahr || '')).filter(Boolean))).sort(compareDe);
            sel.replaceChildren();
            const optAll = document.createElement('option');
            optAll.value = '';
            optAll.textContent = '(alle)';
            sel.appendChild(optAll);
            years.forEach((y) => {
                const o = document.createElement('option');
                o.value = y;
                o.textContent = y;
                sel.appendChild(o);
            });
            if (prev && years.indexOf(prev) !== -1) sel.value = prev;
            // Im Tenant-Modus ist Schuljahr-Filtern sinnlos -> disabled
            sel.disabled = mode === 'tenant';
        }

        const typeSel = getEl('ssFilterTyp');
        if (typeSel) {
            // Typ-Optionen je Modus
            const prevType = typeSel.value || '';
            const opts =
                mode === 'tenant'
                    ? [
                          { v: '', t: '(alle)' },
                          { v: 'Team', t: 'Team' },
                          { v: 'Gruppe', t: 'Gruppe' }
                      ]
                    : [
                          { v: '', t: '(alle)' },
                          { v: 'Jahrgang', t: 'Jahrgang' },
                          { v: 'Klasse', t: 'Klasse' },
                          { v: 'Arbeitsgemeinschaft', t: 'Arbeitsgemeinschaft' },
                          { v: 'Gruppe', t: 'Gruppe' }
                      ];
            typeSel.replaceChildren();
            for (const o of opts) {
                const elO = document.createElement('option');
                elO.value = o.v;
                elO.textContent = o.t;
                typeSel.appendChild(elO);
            }
            if (prevType && opts.some((o) => o.v === prevType)) typeSel.value = prevType;
        }
    }

    function renderStats(rows, mode) {
        const s = computeStats(rows);
        const elTotal = getEl('ssStatEinheiten');
        const elAktiv = getEl('ssStatAktiv');
        const elAbw = getEl('ssStatAbweichung');
        const elErr = getEl('ssStatFehler');
        if (mode === 'tenant') {
            const ts = computeTenantStats(rows);
            if (elTotal) elTotal.textContent = String(ts.total);
            if (elAktiv) elAktiv.textContent = String(ts.teams);
            if (elAbw) elAbw.textContent = String(ts.groups);
            if (elErr) elErr.textContent = '–';
            const labAktiv = elAktiv && elAktiv.parentElement ? elAktiv.parentElement.querySelector('.l') : null;
            const labAbw = elAbw && elAbw.parentElement ? elAbw.parentElement.querySelector('.l') : null;
            const labErr = elErr && elErr.parentElement ? elErr.parentElement.querySelector('.l') : null;
            if (labAktiv) labAktiv.textContent = 'Teams';
            if (labAbw) labAbw.textContent = 'Gruppen';
            if (labErr) labErr.textContent = '—';
            return;
        }
        if (elTotal) elTotal.textContent = String(s.total);
        if (elAktiv) elAktiv.textContent = String(s.aktiv);
        if (elAbw) elAbw.textContent = String(s.abw);
        if (elErr) elErr.textContent = String(s.err);
        const labAktiv = elAktiv && elAktiv.parentElement ? elAktiv.parentElement.querySelector('.l') : null;
        const labAbw = elAbw && elAbw.parentElement ? elAbw.parentElement.querySelector('.l') : null;
        const labErr = elErr && elErr.parentElement ? elErr.parentElement.querySelector('.l') : null;
        if (labAktiv) labAktiv.textContent = 'aktiv';
        if (labAbw) labAbw.textContent = 'Abweichung';
        if (labErr) labErr.textContent = 'Fehler';
    }

    function renderTree(rows, selectedId, mode) {
        const tree = getEl('ssTree');
        if (!tree) return;
        tree.replaceChildren();

        const visible = applyFilters(rows, mode);
        renderStats(visible, mode);

        if (!visible.length) {
            const li = document.createElement('li');
            li.style.padding = '10px 12px';
            li.style.color = '#6c757d';
            li.textContent = rows.length ? 'Keine Treffer für den Filter.' : 'Noch keine Einheiten. Mit „Demo“ oder „Neu“ starten.';
            tree.appendChild(li);
            return;
        }

        const ordered = mode === 'tenant'
            ? visible
                  .slice()
                  .sort((a, b) => compareDe(a.bezeichnung, b.bezeichnung))
                  .map((r) => ({ r, depth: 0 }))
            : buildTreeOrder(visible);

        for (const { r, depth } of ordered) {
            const li = document.createElement('li');
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.dataset.ssSelect = String(r.id);
            btn.setAttribute('aria-current', selectedId && String(selectedId) === String(r.id) ? 'true' : 'false');
            btn.style.paddingLeft = String(12 + depth * 18) + 'px';

            const name = document.createElement('div');
            name.style.minWidth = '0';
            name.style.flex = '1';
            name.style.fontWeight = '900';
            name.style.color = '#32325d';
            name.style.overflow = 'hidden';
            name.style.textOverflow = 'ellipsis';
            name.style.whiteSpace = 'nowrap';
            name.textContent = r.bezeichnung || '(ohne Bezeichnung)';

            const meta = document.createElement('div');
            meta.className = 'pill ' + pillClass(r.syncStatus);
            meta.title = mode === 'tenant' ? 'Art' : 'Sync-Status';
            meta.textContent = mode === 'tenant' ? String(r.typ || '–') : String(r.syncStatus || 'Ausstehend');

            const meta2 = document.createElement('div');
            meta2.className = 'pill';
            meta2.title = mode === 'tenant' ? 'Alias / E-Mail' : 'Typ / Schuljahr';
            meta2.textContent =
                mode === 'tenant'
                    ? String(r.alias || r.mail || '–')
                    : String(r.typ || '–') + (r.schuljahr ? ' · ' + String(r.schuljahr) : '');

            btn.appendChild(name);
            btn.appendChild(meta2);
            btn.appendChild(meta);
            li.appendChild(btn);
            tree.appendChild(li);
        }
    }

    function fillParentSelect(rows, currentId) {
        const sel = getEl('ssUebergeordnet');
        if (!sel) return;
        const prev = sel.value || '';
        const opts = rows
            .filter((r) => r && r.id && String(r.id) !== String(currentId || ''))
            .slice()
            .sort((a, b) => compareDe(a.bezeichnung, b.bezeichnung));

        sel.replaceChildren();
        const none = document.createElement('option');
        none.value = '';
        none.textContent = '(keine)';
        sel.appendChild(none);

        for (const r of opts) {
            const o = document.createElement('option');
            o.value = String(r.id);
            o.textContent = String(r.bezeichnung || '(ohne Bezeichnung)') + ' – ' + String(r.typ || '');
            sel.appendChild(o);
        }
        if (prev && opts.some((r) => String(r.id) === String(prev))) sel.value = prev;
    }

    function showDetail(isOn) {
        const hint = getEl('ssHint');
        const detail = getEl('ssDetail');
        const tenantDetail = getEl('ssTenantDetail');
        if (hint) hint.style.display = isOn ? 'none' : '';
        if (detail) detail.style.display = isOn ? '' : 'none';
        if (tenantDetail) tenantDetail.style.display = 'none';
    }

    function setDetailFromRow(row, rows) {
        if (!row) {
            showDetail(false);
            return;
        }
        showDetail(true);
        fillParentSelect(rows, row.id);
        getEl('ssBezeichnung').value = String(row.bezeichnung || '');
        getEl('ssTyp').value = String(row.typ || 'Gruppe');
        getEl('ssSchuljahr').value = String(row.schuljahr || '');
        getEl('ssStatus').value = String(row.status || 'Aktiv');
        getEl('ssUebergeordnet').value = String(row.parentId || '');
        getEl('ssSyncStatus').value = String(row.syncStatus || 'Ausstehend');
        getEl('ssLetzteFehlermeldung').value = String(row.letzteFehlermeldung || '');
    }

    function showTenantDetail(group) {
        const hint = getEl('ssHint');
        const detail = getEl('ssDetail');
        const tenantDetail = getEl('ssTenantDetail');
        if (hint) hint.style.display = group ? 'none' : '';
        if (detail) detail.style.display = 'none';
        if (tenantDetail) tenantDetail.style.display = group ? '' : 'none';
        if (!group) return;
        getEl('ssTenantName').value = String(group.bezeichnung || '');
        getEl('ssTenantArt').value = String(group.typ || '');
        getEl('ssTenantMail').value = String(group.mail || '');
        getEl('ssTenantAlias').value = String(group.alias || '');
        getEl('ssTenantId').value = String(group.id || '');
    }

    function readDetailToRow(row) {
        const next = Object.assign({}, row);
        next.bezeichnung = normStr(getEl('ssBezeichnung')?.value);
        next.typ = normStr(getEl('ssTyp')?.value) || 'Gruppe';
        next.schuljahr = normStr(getEl('ssSchuljahr')?.value);
        next.status = normStr(getEl('ssStatus')?.value) || 'Aktiv';
        next.parentId = normStr(getEl('ssUebergeordnet')?.value);
        next.syncStatus = normStr(getEl('ssSyncStatus')?.value) || 'Ausstehend';
        next.letzteFehlermeldung = normStr(getEl('ssLetzteFehlermeldung')?.value);
        return next;
    }

    function downloadJson(filename, obj) {
        const blob = new Blob([JSON.stringify(obj, null, 2)], { type: 'application/json;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        a.remove();
        setTimeout(() => URL.revokeObjectURL(url), 500);
    }

    // --- Graph (Tenant read) ---
    let msalMod = null;
    let pca = null;

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
        if (!id) throw new Error('Keine clientId: ms365-config.js fehlt/leer oder blockiert.');
        return {
            clientId: id,
            authority: cfg.authority || 'https://login.microsoftonline.com/organizations',
            redirectUri: (cfg.redirectUri || window.location.href.split('#')[0]).trim()
        };
    }

    async function getPca() {
        const m = await loadMsal();
        const PublicClientApplication = m.PublicClientApplication || (m.default && m.default.PublicClientApplication);
        if (!PublicClientApplication) throw new Error('MSAL: PublicClientApplication nicht gefunden.');
        const cfg = resolveMsalConfig();
        if (!pca) {
            pca = new PublicClientApplication({
                auth: { clientId: cfg.clientId, authority: cfg.authority, redirectUri: cfg.redirectUri },
                cache: { cacheLocation: 'sessionStorage', storeAuthStateInCookie: true }
            });
            await pca.initialize();
            await pca.handleRedirectPromise();
        }
        return pca;
    }

    async function getGraphToken(scopes) {
        const instance = await getPca();
        let accounts = instance.getAllAccounts();
        if (!accounts.length) {
            await instance.loginPopup({ scopes, prompt: 'select_account' });
            accounts = instance.getAllAccounts();
        }
        if (!accounts.length) throw new Error('Anmeldung abgebrochen.');
        const req = { scopes, account: accounts[0] };
        try {
            return (await instance.acquireTokenSilent(req)).accessToken;
        } catch (e) {
            if (isInteractionRequired(e)) return (await instance.acquireTokenPopup(req)).accessToken;
            throw e;
        }
    }

    function sleep(ms) {
        return new Promise((r) => setTimeout(r, ms));
    }

    async function graphRequest(method, pathOrUrl, token) {
        const url = pathOrUrl.indexOf('http') === 0 ? pathOrUrl : 'https://graph.microsoft.com/v1.0' + pathOrUrl;
        let attempt = 0;
        while (true) {
            const res = await fetch(url, { method, headers: { Authorization: 'Bearer ' + token } });
            if (res.status === 429 && attempt < 8) {
                const ra = parseInt(res.headers.get('Retry-After') || '5', 10);
                await sleep((isNaN(ra) ? 5 : ra) * 1000);
                attempt++;
                continue;
            }
            return res;
        }
    }

    async function graphJson(method, pathOrUrl, token) {
        const res = await graphRequest(method, pathOrUrl, token);
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
            const msg = typeof data === 'object' && data && data.error ? JSON.stringify(data.error) : text || String(res.status);
            throw new Error(method + ' ' + pathOrUrl + ': ' + msg);
        }
        return data || {};
    }

    async function fetchAllPages(token, initialPath) {
        const out = [];
        let next = initialPath;
        while (next) {
            const data = await graphJson('GET', next, token);
            const vals = data.value;
            if (Array.isArray(vals)) for (let i = 0; i < vals.length; i++) out.push(vals[i]);
            next = data['@odata.nextLink'] || null;
        }
        return out;
    }

    function groupIsTeam(g) {
        const opts = g && g.resourceProvisioningOptions;
        return Array.isArray(opts) && opts.indexOf('Team') !== -1;
    }

    async function loadTenantGroupsLive() {
        const token = await getGraphToken(GRAPH_SCOPES_TENANT_READ);
        const select = 'id,displayName,mail,mailNickname,createdDateTime,groupTypes,resourceProvisioningOptions';
        const filter = encodeURIComponent("groupTypes/any(c:c eq 'Unified')");
        const initial = '/groups?$filter=' + filter + '&$select=' + encodeURIComponent(select) + '&$top=999';
        const groups = await fetchAllPages(token, initial);
        const mapped = groups.map((g) => {
            const isTeam = groupIsTeam(g);
            return {
                id: String(g.id || ''),
                bezeichnung: String(g.displayName || ''),
                typ: isTeam ? 'Team' : 'Gruppe',
                mail: String(g.mail || ''),
                alias: String(g.mailNickname || ''),
                createdDateTime: String(g.createdDateTime || '')
            };
        }).filter((x) => x.id);
        mapped.sort((a, b) => compareDe(a.bezeichnung, b.bezeichnung));
        saveTenantCache(mapped);
        return mapped;
    }

    function bind() {
        const state = loadState();
        let rowsStruktur = state.rows.slice();
        const tenantCache = loadTenantCache();
        let rowsTenant = tenantCache.rows.slice();
        let selectedId = '';
        /** @type {'struktur'|'tenant'} */
        let mode = 'struktur';

        function rerender() {
            const rows = mode === 'tenant' ? rowsTenant : rowsStruktur;
            renderFilters(rows, mode);
            renderTree(rows, selectedId, mode);
            setModeHint(mode, tenantCache.loadedAt || '');

            const sel = rows.find((r) => String(r.id) === String(selectedId));
            if (mode === 'tenant') {
                showTenantDetail(sel || null);
            } else {
                setDetailFromRow(sel || null, rowsStruktur);
            }
        }

        function select(id) {
            selectedId = id ? String(id) : '';
            rerender();
        }

        function upsert(row) {
            const id = String(row.id);
            const idx = rowsStruktur.findIndex((r) => String(r.id) === id);
            if (idx === -1) rowsStruktur.push(row);
            else rowsStruktur[idx] = row;
            saveState({ rows: rowsStruktur });
        }

        function remove(id) {
            const sid = String(id);
            // entferne auch direkte Kinder (Mock: 1 Ebene reicht fürs UI, reicht für v1)
            const childIds = rowsStruktur.filter((r) => String(r.parentId || '') === sid).map((r) => String(r.id));
            rowsStruktur = rowsStruktur.filter((r) => String(r.id) !== sid && childIds.indexOf(String(r.id)) === -1);
            saveState({ rows: rowsStruktur });
            if (selectedId === sid) selectedId = '';
        }

        // Tree click
        getEl('ssTree')?.addEventListener('click', (ev) => {
            const t = ev.target;
            const btn = t && t.closest ? t.closest('button[data-ss-select]') : null;
            if (!btn) return;
            const id = btn.getAttribute('data-ss-select');
            if (id) select(id);
        });

        // Filter
        const onFilter = () => rerender();
        getEl('ssFilterSchuljahr')?.addEventListener('change', onFilter);
        getEl('ssFilterTyp')?.addEventListener('change', onFilter);
        getEl('ssFilterText')?.addEventListener('input', onFilter);

        // Mode switch
        const ansicht = getEl('ssAnsicht');
        const btnNeu = getEl('ssBtnNeu');
        const btnDemo = getEl('ssBtnDemo');
        const btnReset = getEl('ssBtnReset');
        const btnLogin = getEl('ssBtnTenantLogin');
        const btnLoad = getEl('ssBtnTenantLoad');
        function updateModeUi() {
            const isTenant = mode === 'tenant';
            if (btnNeu) btnNeu.style.display = isTenant ? 'none' : '';
            if (btnDemo) btnDemo.style.display = isTenant ? 'none' : '';
            if (btnReset) btnReset.style.display = isTenant ? 'none' : '';
            if (btnLogin) btnLogin.style.display = isTenant ? '' : 'none';
            if (btnLoad) btnLoad.style.display = isTenant ? '' : 'none';
        }
        if (ansicht) {
            ansicht.addEventListener('change', () => {
                mode = String(ansicht.value || 'struktur') === 'tenant' ? 'tenant' : 'struktur';
                selectedId = '';
                updateModeUi();
                rerender();
            });
        }

        // Buttons
        getEl('ssBtnDemo')?.addEventListener('click', () => {
            rowsStruktur = buildDemoRows();
            saveState({ rows: rowsStruktur });
            selectedId = '';
            rerender();
        });
        getEl('ssBtnReset')?.addEventListener('click', () => {
            if (!window.confirm('Lokale Mock-Daten wirklich löschen?')) return;
            rowsStruktur = [];
            saveState({ rows: rowsStruktur });
            selectedId = '';
            rerender();
        });
        getEl('ssBtnNeu')?.addEventListener('click', () => {
            const r = { id: uid(), parentId: '', typ: 'Klasse', bezeichnung: '', schuljahr: '', status: 'Aktiv', syncStatus: 'Ausstehend', letzteFehlermeldung: '' };
            rowsStruktur.push(r);
            saveState({ rows: rowsStruktur });
            select(r.id);
        });
        getEl('ssBtnNeuUnter')?.addEventListener('click', () => {
            if (!selectedId) return;
            const parent = rowsStruktur.find((r) => String(r.id) === String(selectedId));
            const r = { id: uid(), parentId: selectedId, typ: parent && parent.typ === 'Jahrgang' ? 'Klasse' : 'Gruppe', bezeichnung: '', schuljahr: parent ? String(parent.schuljahr || '') : '', status: 'Aktiv', syncStatus: 'Ausstehend', letzteFehlermeldung: '' };
            rowsStruktur.push(r);
            saveState({ rows: rowsStruktur });
            select(r.id);
        });
        getEl('ssBtnSpeichern')?.addEventListener('click', () => {
            const cur = rowsStruktur.find((r) => String(r.id) === String(selectedId));
            if (!cur) return;
            const next = readDetailToRow(cur);
            if (!next.bezeichnung) {
                window.alert('Bitte eine Bezeichnung eingeben.');
                return;
            }
            // simple cycle guard
            if (next.parentId && String(next.parentId) === String(next.id)) next.parentId = '';
            upsert(next);
            rerender();
        });
        getEl('ssBtnLoeschen')?.addEventListener('click', () => {
            const cur = rowsStruktur.find((r) => String(r.id) === String(selectedId));
            if (!cur) return;
            const label = cur.bezeichnung ? '"' + cur.bezeichnung + '"' : 'diesen Eintrag';
            if (!window.confirm('Wirklich ' + label + ' löschen? (Unterpunkte werden ebenfalls entfernt.)')) return;
            remove(cur.id);
            selectedId = '';
            rerender();
        });

        getEl('ssBtnExport')?.addEventListener('click', () => {
            downloadJson('schulstruktur-sync.json', { rows: rowsStruktur });
        });
        getEl('ssImportFile')?.addEventListener('change', async (e) => {
            const f = e.target.files && e.target.files[0];
            if (!f) return;
            try {
                const text = await f.text();
                const obj = safeJsonParse(text);
                const nextRows = obj && Array.isArray(obj.rows) ? obj.rows : null;
                if (!nextRows) {
                    window.alert('Import fehlgeschlagen: ungültige Datei (erwartet { "rows": [...] }).');
                    return;
                }
                rowsStruktur = nextRows
                    .filter((r) => r && r.id)
                    .map((r) => ({
                        id: String(r.id),
                        parentId: String(r.parentId || ''),
                        typ: normStr(r.typ) || 'Gruppe',
                        bezeichnung: normStr(r.bezeichnung) || '',
                        schuljahr: normStr(r.schuljahr) || '',
                        status: normStr(r.status) || 'Aktiv',
                        syncStatus: normStr(r.syncStatus) || 'Ausstehend',
                        letzteFehlermeldung: normStr(r.letzteFehlermeldung) || ''
                    }));
                saveState({ rows: rowsStruktur });
                selectedId = '';
                rerender();
            } catch (err) {
                window.alert('Import fehlgeschlagen: ' + (err?.message || String(err)));
            } finally {
                e.target.value = '';
            }
        });

        getEl('ssBtnTenantLogin')?.addEventListener('click', async () => {
            try {
                await getGraphToken(GRAPH_SCOPES_TENANT_READ);
                toast('Microsoft angemeldet – Sie können jetzt „Tenant einlesen“ wählen.');
            } catch (e) {
                toast('Anmeldung: ' + (e?.message || String(e)));
            }
        });
        getEl('ssBtnTenantLoad')?.addEventListener('click', async () => {
            const btn = getEl('ssBtnTenantLoad');
            if (btn) btn.disabled = true;
            try {
                toast('Lese Tenant‑Gruppen/Teams …');
                rowsTenant = await loadTenantGroupsLive();
                const cache = loadTenantCache();
                tenantCache.loadedAt = cache.loadedAt;
                toast('Tenant eingelesen: ' + rowsTenant.length + ' Gruppe(n)/Team(s).');
                selectedId = '';
                rerender();
            } catch (e) {
                toast('Tenant einlesen: ' + (e?.message || String(e)));
            } finally {
                if (btn) btn.disabled = false;
            }
        });

        // initial render
        updateModeUi();
        rerender();
    }

    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', bind);
    else bind();
})();

