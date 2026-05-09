(function () {
    'use strict';

    const GRAPH_SCOPES = [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/User.Read.All',
        'https://graph.microsoft.com/User.ReadWrite.All',
        'https://graph.microsoft.com/Group.Read.All'
    ];

    const USER_LIST_SELECT =
        'id,displayName,givenName,surname,mail,userPrincipalName,jobTitle,department,' +
        'officeLocation,mobilePhone,businessPhones,companyName,preferredLanguage,accountEnabled,' +
        'createdDateTime,userType';

    const USER_REFRESH_SELECT = USER_LIST_SELECT;

    const GROUP_MEMBEROF_SELECT = 'id,displayName,mail,mailNickname,groupTypes,securityEnabled,mailEnabled';

    let msalMod = null;
    let pca = null;
    /** @type {Record<string, any>[]} */
    let loadedUsers = [];
    /** @type {string | null} */
    let selectedUserId = null;
    /** @type {'profil' | 'gruppen'} */
    let activeTab = 'profil';
    /** @type {Record<string, any>[] | null} */
    let cachedGroupsForSelection = null;
    /** @type {boolean} */
    let profileEditMode = false;

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

    async function graphRequest(method, pathOrUrl, token, body, extraHeaders) {
        const url =
            pathOrUrl.indexOf('http') === 0 ? pathOrUrl : 'https://graph.microsoft.com/v1.0' + pathOrUrl;
        let attempt = 0;
        while (true) {
            const headers = { Authorization: 'Bearer ' + token };
            if (extraHeaders && typeof extraHeaders === 'object') {
                Object.assign(headers, extraHeaders);
            }
            if (body !== undefined && method !== 'GET' && method !== 'DELETE') {
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

    async function graphDelete(path, token) {
        const res = await graphRequest('DELETE', path, token, undefined);
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
            throw new Error('DELETE ' + path + ': ' + msg);
        }
    }

    function appendLog(msg, kind) {
        const el = document.getElementById('pvLog');
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
        const el = document.getElementById('pvLog');
        if (el) el.replaceChildren();
    }

    async function fetchAllPages(token, initialPath, onProgress) {
        const out = [];
        let next = initialPath;
        let page = 0;
        while (next) {
            page++;
            const data = await graphJson('GET', next, token, undefined);
            const vals = data.value;
            if (Array.isArray(vals)) {
                for (let i = 0; i < vals.length; i++) out.push(vals[i]);
            }
            next = data['@odata.nextLink'] || null;
            if (onProgress) onProgress(out.length, page, !!next);
        }
        return out;
    }

    function norm(s) {
        return String(s || '').trim().toLowerCase();
    }

    function compareStrings(a, b) {
        return String(a || '').localeCompare(String(b || ''), 'de', { sensitivity: 'base' });
    }

    function readSortFromSelect() {
        const sel = document.getElementById('pvSortKey');
        const raw = sel && sel.value ? String(sel.value) : 'displayName:asc';
        const parts = raw.split(':');
        const key = parts[0] || 'displayName';
        const dir = parts[1] === 'desc' ? 'desc' : 'asc';
        return { key: key, dir: dir };
    }

    function formatPhones(u) {
        const m = u && u.mobilePhone ? String(u.mobilePhone).trim() : '';
        const bp = u && Array.isArray(u.businessPhones) ? u.businessPhones.filter(Boolean).join(', ') : '';
        if (m && bp) return m + ' · ' + bp;
        return m || bp || '';
    }

    function formatDate(iso) {
        if (!iso) return '–';
        try {
            const d = new Date(iso);
            if (isNaN(d.getTime())) return String(iso);
            return d.toLocaleString(undefined, {
                dateStyle: 'medium',
                timeStyle: 'short'
            });
        } catch {
            return String(iso);
        }
    }

    function groupTypeLabel(g) {
        if (!g || typeof g !== 'object') return '–';
        const types = g.groupTypes;
        if (Array.isArray(types) && types.indexOf('Unified') !== -1) return 'Microsoft 365 (Unified)';
        if (g.securityEnabled && !g.mailEnabled) return 'Sicherheitsgruppe';
        if (g.mailEnabled && !g.securityEnabled) return 'Verteilerliste';
        if (g.securityEnabled && g.mailEnabled) return 'Mail-aktivierte Sicherheitsgruppe';
        return 'Gruppe';
    }

    function userTypeLabel(ut) {
        const t = String(ut || '').toLowerCase();
        if (t === 'guest') return 'Gast';
        if (t === 'member') return 'Mitglied';
        return ut ? String(ut) : '–';
    }

    function getSelectedUser() {
        if (!selectedUserId) return null;
        return loadedUsers.find(function (x) {
            return x.id === selectedUserId;
        }) || null;
    }

    function setProfileEditMode(on) {
        profileEditMode = !!on;
        updateDetailActionButtons();
        const u = getSelectedUser();
        if (u) renderProfileTab(u, profileEditMode);
    }

    function updateDetailActionButtons() {
        const edit = document.getElementById('pvBtnEdit');
        const save = document.getElementById('pvBtnSave');
        const cancel = document.getElementById('pvBtnCancelEdit');
        const del = document.getElementById('pvBtnDelete');
        const hasSel = !!selectedUserId;
        if (edit) {
            edit.style.display = hasSel && !profileEditMode ? '' : 'none';
            edit.disabled = !hasSel;
        }
        if (save) save.style.display = hasSel && profileEditMode ? '' : 'none';
        if (cancel) cancel.style.display = hasSel && profileEditMode ? '' : 'none';
        if (del) {
            del.style.display = hasSel && !profileEditMode ? '' : 'none';
            del.disabled = !hasSel;
        }
    }

    function getVisibleRows() {
        const filterInp = document.getElementById('pvFilterText');
        const q = filterInp && filterInp.value ? norm(filterInp.value) : '';

        const typeSel = document.getElementById('pvFilterUserType');
        const typeVal = typeSel && typeSel.value ? String(typeSel.value) : '';

        const accSel = document.getElementById('pvFilterAccount');
        const accVal = accSel && accSel.value !== '' ? String(accSel.value) : '';

        const depSel = document.getElementById('pvFilterDepartment');
        const depVal = depSel && depSel.value ? String(depSel.value) : '';

        let rows = loadedUsers.slice();

        if (typeVal) {
            rows = rows.filter(function (u) {
                return String(u.userType || '') === typeVal;
            });
        }

        if (accVal === '1') {
            rows = rows.filter(function (u) {
                return u.accountEnabled === true;
            });
        } else if (accVal === '0') {
            rows = rows.filter(function (u) {
                return u.accountEnabled === false;
            });
        }

        if (depVal) {
            rows = rows.filter(function (u) {
                return String(u.department || '').trim() === depVal;
            });
        }

        if (q) {
            rows = rows.filter(function (u) {
                const blob = [
                    u.displayName,
                    u.givenName,
                    u.surname,
                    u.mail,
                    u.userPrincipalName,
                    u.department,
                    u.jobTitle,
                    u.id,
                    u.officeLocation,
                    u.companyName
                ]
                    .map(function (x) {
                        return norm(x);
                    })
                    .join(' ');
                return blob.indexOf(q) !== -1;
            });
        }

        const sortState = readSortFromSelect();
        const key = sortState.key || 'displayName';
        const dir = sortState.dir === 'desc' ? -1 : 1;

        rows.sort(function (ua, ub) {
            return compareStrings(ua[key] || '', ub[key] || '') * dir;
        });

        return rows;
    }

    function refreshDepartmentFilter() {
        const sel = document.getElementById('pvFilterDepartment');
        if (!sel) return;
        const current = sel.value;
        const set = new Set();
        for (let i = 0; i < loadedUsers.length; i++) {
            const d = loadedUsers[i].department;
            if (d && String(d).trim()) set.add(String(d).trim());
        }
        const list = Array.from(set).sort(function (a, b) {
            return compareStrings(a, b);
        });
        sel.replaceChildren();
        const o0 = document.createElement('option');
        o0.value = '';
        o0.textContent = '(alle)';
        sel.appendChild(o0);
        for (let j = 0; j < list.length; j++) {
            const o = document.createElement('option');
            o.value = list[j];
            o.textContent = list[j];
            sel.appendChild(o);
        }
        if (current && set.has(current)) sel.value = current;
    }

    function updateStatsPanel() {
        const total = loadedUsers.length;
        let members = 0;
        let guests = 0;
        let active = 0;
        for (let i = 0; i < loadedUsers.length; i++) {
            const u = loadedUsers[i];
            if (String(u.userType || '').toLowerCase() === 'guest') guests++;
            else members++;
            if (u.accountEnabled === true) active++;
        }
        const el = function (id, val) {
            const n = document.getElementById(id);
            if (n) n.textContent = val;
        };
        el('pvStatTotal', total ? String(total) : '–');
        el('pvStatMember', total ? String(members) : '–');
        el('pvStatGuest', total ? String(guests) : '–');
        el('pvStatActive', total ? String(active) : '–');
    }

    function updateProgressLine() {
        const progress = document.getElementById('pvProgress');
        if (!progress) return;
        if (!loadedUsers.length) {
            progress.textContent = '';
            return;
        }
        const visible = getVisibleRows();
        const base = 'Geladen: ' + loadedUsers.length + ' Person(en).';
        if (visible.length !== loadedUsers.length) {
            progress.textContent = base + ' Angezeigt: ' + visible.length + ' Treffer.';
        } else {
            progress.textContent = base;
        }
    }

    function renderUserTree() {
        const tree = document.getElementById('pvTree');
        if (!tree) return;
        tree.replaceChildren();
        const rows = getVisibleRows();

        if (!rows.length) {
            const li = document.createElement('li');
            const p = document.createElement('p');
            p.className = 'muted';
            p.style.margin = '0';
            p.style.padding = '14px 12px';
            p.textContent = loadedUsers.length ? 'Keine Treffer für die Filter.' : 'Noch keine Daten – „Personen einlesen“ wählen.';
            li.appendChild(p);
            tree.appendChild(li);
            updateProgressLine();
            return;
        }

        for (let i = 0; i < rows.length; i++) {
            const u = rows[i];
            const li = document.createElement('li');
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'pv-tree-row';
            btn.dataset.pvSelectUser = u.id || '';
            btn.setAttribute('aria-current', selectedUserId && u.id === selectedUserId ? 'true' : 'false');

            const isGuest = String(u.userType || '').toLowerCase() === 'guest';
            const iconWrap = document.createElement('span');
            iconWrap.className = 'pv-tree-icon';
            const icon = document.createElement('i');
            icon.className = isGuest ? 'bi bi-person-badge' : 'bi bi-person-fill';
            icon.setAttribute('aria-hidden', 'true');
            iconWrap.appendChild(icon);

            const main = document.createElement('div');
            main.className = 'pv-tree-main';
            const title = document.createElement('div');
            title.className = 'pv-tree-title';
            title.textContent = u.displayName || u.userPrincipalName || u.mail || '(ohne Namen)';
            const sub = document.createElement('div');
            sub.className = 'pv-tree-sub';
            sub.textContent = u.userPrincipalName || u.mail || u.id || '';
            main.appendChild(title);
            main.appendChild(sub);

            const meta = document.createElement('div');
            meta.className = 'pv-tree-meta';
            const pillType = document.createElement('span');
            pillType.className = 'pill' + (isGuest ? '' : ' ok');
            pillType.textContent = userTypeLabel(u.userType);
            meta.appendChild(pillType);
            if (u.accountEnabled === false) {
                const pillOff = document.createElement('span');
                pillOff.className = 'pill err';
                pillOff.textContent = 'Inaktiv';
                meta.appendChild(pillOff);
            }
            if (u.department) {
                const pillDep = document.createElement('span');
                pillDep.className = 'pill muted-pill';
                pillDep.textContent = String(u.department).trim();
                meta.appendChild(pillDep);
            }

            btn.appendChild(iconWrap);
            btn.appendChild(main);
            btn.appendChild(meta);
            li.appendChild(btn);
            tree.appendChild(li);
        }
        updateProgressLine();
    }

    function dispVal(v) {
        if (v === undefined || v === null || v === '') return '–';
        return String(v);
    }

    function addProfileTextField(root, label, fieldKey, value, editable, fullWidth) {
        const wrap = document.createElement('div');
        wrap.className = 'field' + (fullWidth ? ' field-full' : '');
        const lab = document.createElement('label');
        lab.setAttribute('for', 'pv_f_' + fieldKey);
        lab.textContent = label;
        const inp = document.createElement('input');
        inp.type = 'text';
        inp.id = 'pv_f_' + fieldKey;
        inp.dataset.pvField = fieldKey;
        inp.readOnly = !editable;
        if (editable && (value === undefined || value === null || value === '')) {
            inp.value = '';
        } else {
            inp.value = dispVal(value);
        }
        if (inp.value === '–') inp.style.color = 'var(--muted)';
        else if (!editable) inp.style.color = '#32325d';
        wrap.appendChild(lab);
        wrap.appendChild(inp);
        root.appendChild(wrap);
    }

    function addProfileAccountEnabled(root, u, editable) {
        const wrap = document.createElement('div');
        wrap.className = 'field';
        const lab = document.createElement('label');
        lab.setAttribute('for', 'pv_f_accountEnabled');
        lab.textContent = 'Konto aktiv';
        wrap.appendChild(lab);
        if (editable) {
            const row = document.createElement('div');
            row.className = 'pv-checkbox-row';
            const cb = document.createElement('input');
            cb.type = 'checkbox';
            cb.id = 'pv_f_accountEnabled';
            cb.dataset.pvField = 'accountEnabled';
            cb.checked = u.accountEnabled !== false;
            const l2 = document.createElement('label');
            l2.htmlFor = 'pv_f_accountEnabled';
            l2.style.margin = '0';
            l2.style.fontWeight = '600';
            l2.textContent = 'Konto ist aktiviert';
            row.appendChild(cb);
            row.appendChild(l2);
            wrap.appendChild(row);
        } else {
            const inp = document.createElement('input');
            inp.type = 'text';
            inp.readOnly = true;
            inp.value = u.accountEnabled === false ? 'Nein' : u.accountEnabled === true ? 'Ja' : '–';
            if (inp.value === '–') inp.style.color = 'var(--muted)';
            wrap.appendChild(inp);
        }
        root.appendChild(wrap);
    }

    function renderProfileTab(u, editable) {
        const root = document.getElementById('pvProfileFields');
        if (!root) return;
        root.replaceChildren();
        if (!u) return;

        const ro = !editable;
        addProfileTextField(root, 'Anzeigename', 'displayName', u.displayName, editable, false);
        addProfileTextField(root, 'Vorname', 'givenName', u.givenName, editable, false);
        addProfileTextField(root, 'Nachname', 'surname', u.surname, editable, false);
        addProfileTextField(
            root,
            'Benutzername (UPN)' + (String(u.userType).toLowerCase() === 'guest' ? ' (Gast)' : ''),
            'userPrincipalName',
            u.userPrincipalName,
            editable,
            false
        );
        addProfileTextField(root, 'E-Mail (SMTP)', 'mail', u.mail, editable, false);
        addProfileTextField(root, 'Objekt-ID', '_id', u.id, false, true);
        addProfileTextField(root, 'Kontotyp', '_userType', userTypeLabel(u.userType), false, false);
        addProfileAccountEnabled(root, u, editable);
        addProfileTextField(root, 'Erstellt', '_created', formatDate(u.createdDateTime), false, false);
        addProfileTextField(root, 'Position', 'jobTitle', u.jobTitle, editable, false);
        addProfileTextField(root, 'Abteilung', 'department', u.department, editable, false);
        addProfileTextField(root, 'Firma', 'companyName', u.companyName, editable, false);
        addProfileTextField(root, 'Bürostandort', 'officeLocation', u.officeLocation, editable, false);
        addProfileTextField(root, 'Mobiltelefon', 'mobilePhone', u.mobilePhone, editable, false);
        const bp0 = u.businessPhones && u.businessPhones[0] ? u.businessPhones[0] : '';
        addProfileTextField(root, 'Geschäftstelefon (1. Zeile)', 'businessPhone0', bp0, editable, false);
        addProfileTextField(root, 'Sprache (z. B. de-AT)', 'preferredLanguage', u.preferredLanguage, editable, false);

        if (ro) {
            const inputs = root.querySelectorAll('input[readonly]');
            for (let i = 0; i < inputs.length; i++) {
                inputs[i].style.background = '#f8f9fa';
            }
        }
    }

    function readInputTrim(el) {
        if (!el) return '';
        return String(el.value || '').trim();
    }

    function buildPatchFromForm(u) {
        const root = document.getElementById('pvProfileFields');
        if (!root) return null;

        function get(field) {
            const el = root.querySelector('[data-pv-field="' + field + '"]');
            if (!el) return undefined;
            if (el.type === 'checkbox') return !!el.checked;
            const t = readInputTrim(el);
            return t === '' || t === '–' ? '' : t;
        }

        const patch = {};
        const strFields = [
            'displayName',
            'givenName',
            'surname',
            'userPrincipalName',
            'mail',
            'jobTitle',
            'department',
            'companyName',
            'officeLocation',
            'mobilePhone',
            'preferredLanguage'
        ];
        for (let i = 0; i < strFields.length; i++) {
            const k = strFields[i];
            const nv = get(k);
            if (nv === undefined) continue;
            const ov = u[k] == null ? '' : String(u[k]);
            if (String(nv) !== ov) {
                patch[k] = nv === '' ? null : nv;
            }
        }

        const accEl = root.querySelector('[data-pv-field="accountEnabled"]');
        if (accEl && accEl.type === 'checkbox') {
            const nv = !!accEl.checked;
            const ov = u.accountEnabled !== false;
            if (nv !== ov) patch.accountEnabled = nv;
        }

        const bpNew = get('businessPhone0');
        if (bpNew !== undefined) {
            const ov = u.businessPhones && u.businessPhones[0] ? String(u.businessPhones[0]) : '';
            if (String(bpNew) !== ov) {
                patch.businessPhones = bpNew === '' ? [] : [bpNew];
            }
        }

        return patch;
    }

    async function refreshUserFromGraph(token, userId) {
        const path =
            '/users/' + encodeURIComponent(userId) + '?$select=' + encodeURIComponent(USER_REFRESH_SELECT);
        return graphJson('GET', path, token, undefined);
    }

    function mergeUserIntoList(updated) {
        const idx = loadedUsers.findIndex(function (x) {
            return x.id === updated.id;
        });
        if (idx === -1) {
            loadedUsers.push(updated);
        } else {
            loadedUsers[idx] = updated;
        }
        refreshDepartmentFilter();
        updateStatsPanel();
    }

    async function saveProfilePatch() {
        const u = getSelectedUser();
        if (!u || !profileEditMode) return;
        const root = document.getElementById('pvProfileFields');
        const dnEl = root && root.querySelector('[data-pv-field="displayName"]');
        if (dnEl && readInputTrim(dnEl) === '') {
            toast('Anzeigename darf nicht leer sein.');
            return;
        }
        const upnEl = root && root.querySelector('[data-pv-field="userPrincipalName"]');
        if (upnEl && readInputTrim(upnEl) === '') {
            toast('UPN darf nicht leer sein.');
            return;
        }

        const patch = buildPatchFromForm(u);
        if (!patch || Object.keys(patch).length === 0) {
            toast('Keine Änderungen.');
            setProfileEditMode(false);
            return;
        }
        const saveBtn = document.getElementById('pvBtnSave');
        if (saveBtn) saveBtn.disabled = true;
        try {
            const token = await getGraphToken();
            await graphJson('PATCH', '/users/' + encodeURIComponent(u.id), token, patch);
            const fresh = await refreshUserFromGraph(token, u.id);
            mergeUserIntoList(fresh);
            appendLog('Profil gespeichert (PATCH).', 'ok');
            toast('Gespeichert.');
            profileEditMode = false;
            updateDetailActionButtons();
            renderProfileTab(fresh, false);
            renderUserTree();
        } catch (e) {
            const msg = e && e.message ? e.message : String(e);
            appendLog('PATCH: ' + msg, 'err');
            toast(msg);
        } finally {
            if (saveBtn) saveBtn.disabled = false;
        }
    }

    function sanitizeMailNickname(raw, fallbackFromUpn) {
        let s = String(raw || '').trim();
        if (!s && fallbackFromUpn) {
            const at = fallbackFromUpn.indexOf('@');
            s = at > 0 ? fallbackFromUpn.slice(0, at) : fallbackFromUpn;
        }
        s = s.split('@')[0].replace(/[^a-zA-Z0-9._-]/g, '');
        return s;
    }

    function openCreateModal() {
        const bd = document.getElementById('pvModalCreateBackdrop');
        if (!bd) return;
        const ids = ['pvCreateUpn', 'pvCreateDisplayName', 'pvCreateMailNick', 'pvCreatePassword', 'pvCreateGiven', 'pvCreateSurname', 'pvCreateMail'];
        for (let i = 0; i < ids.length; i++) {
            const el = document.getElementById(ids[i]);
            if (el) el.value = '';
        }
        const f = document.getElementById('pvCreateForcePw');
        if (f) f.checked = true;
        const e = document.getElementById('pvCreateEnabled');
        if (e) e.checked = true;
        bd.classList.add('active');
        bd.setAttribute('aria-hidden', 'false');
    }

    function closeCreateModal() {
        const bd = document.getElementById('pvModalCreateBackdrop');
        if (!bd) return;
        bd.classList.remove('active');
        bd.setAttribute('aria-hidden', 'true');
    }

    async function submitCreateUser() {
        const upn = readInputTrim(document.getElementById('pvCreateUpn'));
        const displayName = readInputTrim(document.getElementById('pvCreateDisplayName'));
        let mailNick = readInputTrim(document.getElementById('pvCreateMailNick'));
        const password = String(document.getElementById('pvCreatePassword')?.value || '');
        const givenName = readInputTrim(document.getElementById('pvCreateGiven'));
        const surname = readInputTrim(document.getElementById('pvCreateSurname'));
        const mail = readInputTrim(document.getElementById('pvCreateMail'));
        const forcePw = document.getElementById('pvCreateForcePw') ? document.getElementById('pvCreateForcePw').checked : true;
        const enabled = document.getElementById('pvCreateEnabled') ? document.getElementById('pvCreateEnabled').checked : true;

        if (!upn || !displayName || !password) {
            toast('UPN, Anzeigename und Kennwort sind Pflichtfelder.');
            return;
        }
        mailNick = sanitizeMailNickname(mailNick, upn);
        if (!mailNick) {
            toast('Mail-Nickname ungültig oder leer.');
            return;
        }

        const body = {
            accountEnabled: enabled,
            displayName: displayName,
            mailNickname: mailNick,
            userPrincipalName: upn,
            passwordProfile: {
                password: password,
                forceChangePasswordNextSignIn: !!forcePw
            }
        };
        if (givenName) body.givenName = givenName;
        if (surname) body.surname = surname;
        if (mail) body.mail = mail;

        const sub = document.getElementById('pvModalCreateSubmit');
        if (sub) sub.disabled = true;
        try {
            const token = await getGraphToken();
            const created = await graphJson('POST', '/users', token, body);
            const id = created && created.id ? created.id : null;
            appendLog('Benutzer angelegt: ' + (created.userPrincipalName || upn), 'ok');
            toast('Benutzer angelegt.');
            closeCreateModal();
            if (id) {
                try {
                    const fresh = await refreshUserFromGraph(token, id);
                    mergeUserIntoList(fresh);
                    selectUser(id);
                } catch {
                    loadedUsers.push(created);
                    refreshDepartmentFilter();
                    updateStatsPanel();
                    if (id) selectUser(id);
                }
            }
            renderUserTree();
        } catch (e) {
            const msg = e && e.message ? e.message : String(e);
            appendLog('Anlegen: ' + msg, 'err');
            toast(msg);
        } finally {
            if (sub) sub.disabled = false;
        }
    }

    function updateDeleteModalUi() {
        const hard = document.getElementById('pvDeleteHard') && document.getElementById('pvDeleteHard').checked;
        const warn = document.getElementById('pvDeleteHardWarn');
        const intro = document.getElementById('pvDeleteSoftIntro');
        const sub = document.getElementById('pvModalDeleteSubmit');
        if (warn) warn.style.display = hard ? 'block' : 'none';
        if (intro) intro.style.opacity = hard ? '0.55' : '1';
        if (sub) {
            sub.textContent = hard ? 'Endgültig löschen' : 'Konto deaktivieren';
            sub.className = hard ? 'btn btn-danger' : 'btn btn-success';
        }
    }

    function openDeleteModal() {
        const u = getSelectedUser();
        if (!u) return;
        const bd = document.getElementById('pvModalDeleteBackdrop');
        const echo = document.getElementById('pvDeleteUpnEcho');
        const inp = document.getElementById('pvDeleteConfirmInput');
        const sub = document.getElementById('pvModalDeleteSubmit');
        const hardChk = document.getElementById('pvDeleteHard');
        if (hardChk) hardChk.checked = false;
        updateDeleteModalUi();
        if (echo) echo.textContent = u.userPrincipalName || u.mail || u.id;
        if (inp) inp.value = '';
        if (sub) sub.disabled = true;
        if (bd) {
            bd.classList.add('active');
            bd.setAttribute('aria-hidden', 'false');
        }
    }

    function closeDeleteModal() {
        const bd = document.getElementById('pvModalDeleteBackdrop');
        if (!bd) return;
        bd.classList.remove('active');
        bd.setAttribute('aria-hidden', 'true');
    }

    function syncDeleteConfirmButton() {
        const u = getSelectedUser();
        const inp = document.getElementById('pvDeleteConfirmInput');
        const sub = document.getElementById('pvModalDeleteSubmit');
        if (!sub || !inp || !u) return;
        const ok = readInputTrim(inp) === String(u.userPrincipalName || '').trim();
        sub.disabled = !ok;
    }

    async function submitDeleteUser() {
        const u = getSelectedUser();
        if (!u) return;
        const inp = document.getElementById('pvDeleteConfirmInput');
        if (!inp || readInputTrim(inp) !== String(u.userPrincipalName || '').trim()) {
            toast('UPN stimmt nicht überein.');
            return;
        }
        const hard = document.getElementById('pvDeleteHard') && document.getElementById('pvDeleteHard').checked;
        const sub = document.getElementById('pvModalDeleteSubmit');
        if (sub) sub.disabled = true;
        try {
            const token = await getGraphToken();
            if (!hard) {
                if (u.accountEnabled === false) {
                    toast('Konto ist bereits deaktiviert.');
                    closeDeleteModal();
                    return;
                }
                await graphJson('PATCH', '/users/' + encodeURIComponent(u.id), token, { accountEnabled: false });
                const fresh = await refreshUserFromGraph(token, u.id);
                mergeUserIntoList(fresh);
                appendLog('Konto deaktiviert: ' + (fresh.userPrincipalName || u.id), 'ok');
                toast('Konto deaktiviert.');
                closeDeleteModal();
                profileEditMode = false;
                cachedGroupsForSelection = null;
                selectUser(fresh.id);
                updateStatsPanel();
                return;
            }

            await graphDelete('/users/' + encodeURIComponent(u.id), token);
            loadedUsers = loadedUsers.filter(function (x) {
                return x.id !== u.id;
            });
            appendLog('Benutzer gelöscht (DELETE): ' + (u.userPrincipalName || u.id), 'ok');
            toast('Dauerhaft gelöscht.');
            closeDeleteModal();
            selectedUserId = null;
            cachedGroupsForSelection = null;
            profileEditMode = false;
            const hint = document.getElementById('pvHint');
            const detail = document.getElementById('pvDetail');
            if (hint) hint.style.display = '';
            if (detail) detail.style.display = 'none';
            updateDetailActionButtons();
            refreshDepartmentFilter();
            updateStatsPanel();
            renderUserTree();
        } catch (e) {
            const msg = e && e.message ? e.message : String(e);
            appendLog((hard ? 'Löschen: ' : 'Deaktivieren: ') + msg, 'err');
            toast(msg);
        } finally {
            if (sub) sub.disabled = false;
            syncDeleteConfirmButton();
        }
    }

    function renderGroupsTable(groups) {
        const tbody = document.getElementById('pvGroupsTbody');
        if (!tbody) return;
        tbody.replaceChildren();

        if (!groups || !groups.length) {
            const tr = document.createElement('tr');
            const td = document.createElement('td');
            td.colSpan = 4;
            td.style.color = '#6c757d';
            td.textContent = 'Keine Gruppenmitgliedschaften gefunden (oder keine Leserechte).';
            tr.appendChild(td);
            tbody.appendChild(tr);
            return;
        }

        const sorted = groups.slice().sort(function (a, b) {
            return compareStrings(a.displayName, b.displayName);
        });

        for (let i = 0; i < sorted.length; i++) {
            const g = sorted[i];
            const tr = document.createElement('tr');
            const tdN = document.createElement('td');
            tdN.textContent = g.displayName || '–';
            const tdM = document.createElement('td');
            tdM.textContent = g.mail || g.mailNickname || '–';
            tdM.style.wordBreak = 'break-all';
            tdM.style.fontSize = '0.9em';
            const tdT = document.createElement('td');
            tdT.textContent = groupTypeLabel(g);
            tdT.style.fontSize = '0.88em';
            const tdI = document.createElement('td');
            tdI.textContent = g.id || '–';
            tdI.style.fontFamily = 'Consolas, monospace';
            tdI.style.fontSize = '0.82em';
            tdI.style.wordBreak = 'break-all';
            tr.appendChild(tdN);
            tr.appendChild(tdM);
            tr.appendChild(tdT);
            tr.appendChild(tdI);
            tbody.appendChild(tr);
        }
    }

    async function fetchUserGroups(token, userId) {
        const path =
            '/users/' +
            encodeURIComponent(userId) +
            '/memberOf/microsoft.graph.group?$select=' +
            encodeURIComponent(GROUP_MEMBEROF_SELECT) +
            '&$top=999';
        return fetchAllPages(token, path, undefined);
    }

    async function loadGroupsForSelected() {
        const prog = document.getElementById('pvGroupsProgress');
        if (!selectedUserId) return;
        if (prog) prog.textContent = 'Lade Gruppen …';

        const tbody = document.getElementById('pvGroupsTbody');
        if (tbody) {
            tbody.replaceChildren();
            const tr = document.createElement('tr');
            const td = document.createElement('td');
            td.colSpan = 4;
            td.style.color = '#6c757d';
            td.textContent = 'Lade …';
            tr.appendChild(td);
            tbody.appendChild(tr);
        }

        try {
            const token = await getGraphToken();
            const groups = await fetchUserGroups(token, selectedUserId);
            cachedGroupsForSelection = groups;
            renderGroupsTable(groups);
            if (prog) prog.textContent = groups.length ? groups.length + ' Gruppe(n).' : 'Keine Einträge.';
            appendLog('Gruppen für ausgewählte Person: ' + groups.length, 'ok');
        } catch (e) {
            cachedGroupsForSelection = [];
            renderGroupsTable([]);
            const msg = e && e.message ? e.message : String(e);
            if (prog) prog.textContent = 'Fehler: ' + msg;
            appendLog('Gruppen laden: ' + msg, 'err');
            toast('Gruppen: ' + msg);
        }
    }

    function setTab(tab) {
        activeTab = tab === 'gruppen' ? 'gruppen' : 'profil';
        const pProf = document.getElementById('pvPanelProfil');
        const pGrp = document.getElementById('pvPanelGruppen');
        const bProf = document.getElementById('pvTabProfil');
        const bGrp = document.getElementById('pvTabGruppen');

        if (pProf) {
            pProf.classList.toggle('active', activeTab === 'profil');
            pProf.setAttribute('aria-hidden', activeTab === 'profil' ? 'false' : 'true');
        }
        if (pGrp) {
            pGrp.classList.toggle('active', activeTab === 'gruppen');
            pGrp.setAttribute('aria-hidden', activeTab === 'gruppen' ? 'false' : 'true');
        }
        if (bProf) {
            bProf.setAttribute('aria-selected', activeTab === 'profil' ? 'true' : 'false');
        }
        if (bGrp) {
            bGrp.setAttribute('aria-selected', activeTab === 'gruppen' ? 'true' : 'false');
        }

        if (activeTab === 'gruppen' && selectedUserId) {
            if (cachedGroupsForSelection === null) {
                loadGroupsForSelected();
            }
        }
    }

    function selectUser(userId) {
        selectedUserId = userId || null;
        cachedGroupsForSelection = null;
        activeTab = 'profil';
        profileEditMode = false;

        const hint = document.getElementById('pvHint');
        const detail = document.getElementById('pvDetail');
        const title = document.getElementById('pvManageTitle');

        if (!selectedUserId) {
            if (hint) hint.style.display = '';
            if (detail) detail.style.display = 'none';
            updateDetailActionButtons();
            renderUserTree();
            return;
        }

        const u = getSelectedUser();

        if (hint) hint.style.display = 'none';
        if (detail) detail.style.display = '';
        if (title) title.textContent = u && u.displayName ? String(u.displayName) : '(ohne Anzeigename)';

        renderProfileTab(u || null, false);
        updateDetailActionButtons();

        if (document.getElementById('pvTabProfil')) {
            document.getElementById('pvTabProfil').setAttribute('aria-selected', 'true');
        }
        if (document.getElementById('pvTabGruppen')) {
            document.getElementById('pvTabGruppen').setAttribute('aria-selected', 'false');
        }
        const pProf = document.getElementById('pvPanelProfil');
        const pGrp = document.getElementById('pvPanelGruppen');
        if (pProf) {
            pProf.classList.add('active');
            pProf.setAttribute('aria-hidden', 'false');
        }
        if (pGrp) {
            pGrp.classList.remove('active');
            pGrp.setAttribute('aria-hidden', 'true');
        }

        renderUserTree();
    }

    async function onLogin() {
        try {
            await getGraphToken();
            toast('Angemeldet.');
            appendLog('Anmeldung erfolgreich.', 'ok');
        } catch (e) {
            appendLog('Anmeldung: ' + (e && e.message ? e.message : String(e)), 'err');
            toast(String(e && e.message ? e.message : e));
        }
    }

    async function loadUsers() {
        const btn = document.getElementById('pvBtnLoad');
        const btnCsv = document.getElementById('pvBtnCsv');
        const progress = document.getElementById('pvProgress');
        if (btn) btn.disabled = true;
        if (btnCsv) btnCsv.disabled = true;
        clearLog();
        loadedUsers = [];
        selectedUserId = null;
        cachedGroupsForSelection = null;
        profileEditMode = false;
        const hint = document.getElementById('pvHint');
        const detail = document.getElementById('pvDetail');
        if (hint) hint.style.display = '';
        if (detail) detail.style.display = 'none';
        updateDetailActionButtons();

        try {
            const token = await getGraphToken();
            appendLog('Lade Benutzer aus dem Verzeichnis …', '');

            const initial =
                '/users?$select=' +
                encodeURIComponent(USER_LIST_SELECT) +
                '&$top=999&$orderby=displayName';

            const users = await fetchAllPages(token, initial, function (count) {
                if (progress) {
                    progress.textContent = 'Gelesen: ' + count + ' Person(en) …';
                }
            });

            users.sort(function (a, b) {
                return compareStrings(a.displayName, b.displayName);
            });
            loadedUsers = users;
            appendLog('Fertig: ' + users.length + ' Person(en).', 'ok');
            refreshDepartmentFilter();
            updateStatsPanel();
            if (progress) progress.textContent = '';
            updateProgressLine();
        } catch (e) {
            appendLog('Laden: ' + (e && e.message ? e.message : String(e)), 'err');
            toast(String(e && e.message ? e.message : e));
            if (progress) progress.textContent = '';
            updateStatsPanel();
        } finally {
            if (btn) btn.disabled = false;
            if (btnCsv) btnCsv.disabled = !loadedUsers.length;
            renderUserTree();
        }
    }

    function exportCsv() {
        if (!loadedUsers.length) {
            toast('Keine Daten zum Exportieren.');
            return;
        }
        const rows = getVisibleRows();
        const headers = [
            'displayName',
            'userPrincipalName',
            'mail',
            'department',
            'jobTitle',
            'id',
            'accountEnabled',
            'userType'
        ];
        const lines = [headers.join(';')];
        for (let i = 0; i < rows.length; i++) {
            const u = rows[i];
            const cells = [];
            for (let h = 0; h < headers.length; h++) {
                const key = headers[h];
                let v = u[key];
                if (v === undefined || v === null) v = '';
                v = String(v).replace(/"/g, '""');
                cells.push('"' + v + '"');
            }
            lines.push(cells.join(';'));
        }
        const blob = new Blob([lines.join('\r\n')], { type: 'text/csv;charset=utf-8' });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = 'personen-export.csv';
        a.click();
        URL.revokeObjectURL(a.href);
        appendLog('CSV exportiert (' + rows.length + ' Zeilen).', 'ok');
    }

    function bind() {
        const btnL = document.getElementById('pvBtnLogin');
        const btnLoad = document.getElementById('pvBtnLoad');
        const btnCsv = document.getElementById('pvBtnCsv');
        const filt = document.getElementById('pvFilterText');
        const tree = document.getElementById('pvTree');
        const reRender = function () {
            renderUserTree();
        };

        if (btnL) btnL.addEventListener('click', () => onLogin());
        if (btnLoad) btnLoad.addEventListener('click', () => loadUsers());
        if (btnCsv) {
            btnCsv.disabled = true;
            btnCsv.addEventListener('click', () => exportCsv());
        }
        if (filt) filt.addEventListener('input', reRender);

        const ft = document.getElementById('pvFilterUserType');
        const fa = document.getElementById('pvFilterAccount');
        const fd = document.getElementById('pvFilterDepartment');
        const fs = document.getElementById('pvSortKey');
        if (ft) ft.addEventListener('change', reRender);
        if (fa) fa.addEventListener('change', reRender);
        if (fd) fd.addEventListener('change', reRender);
        if (fs) fs.addEventListener('change', reRender);

        if (tree) {
            tree.addEventListener('click', function (ev) {
                const t = ev.target;
                if (!t || !t.closest) return;
                const btn = t.closest('button[data-pv-select-user]');
                if (!btn) return;
                const uid = btn.getAttribute('data-pv-select-user');
                selectUser(uid || null);
            });
        }

        document.querySelectorAll('.detail-tab-btn[data-pv-tab]').forEach(function (b) {
            b.addEventListener('click', function () {
                const tab = b.getAttribute('data-pv-tab');
                setTab(tab === 'gruppen' ? 'gruppen' : 'profil');
            });
        });

        const btnNeu = document.getElementById('pvBtnNeu');
        if (btnNeu) btnNeu.addEventListener('click', () => openCreateModal());

        document.getElementById('pvModalCreateClose')?.addEventListener('click', closeCreateModal);
        document.getElementById('pvModalCreateCancel')?.addEventListener('click', closeCreateModal);
        document.getElementById('pvModalCreateSubmit')?.addEventListener('click', () => submitCreateUser());
        document.getElementById('pvModalCreateBackdrop')?.addEventListener('click', function (ev) {
            if (ev.target === ev.currentTarget) closeCreateModal();
        });

        document.getElementById('pvBtnEdit')?.addEventListener('click', function () {
            if (!getSelectedUser()) return;
            setProfileEditMode(true);
        });
        document.getElementById('pvBtnCancelEdit')?.addEventListener('click', function () {
            profileEditMode = false;
            updateDetailActionButtons();
            const u = getSelectedUser();
            if (u) renderProfileTab(u, false);
        });
        document.getElementById('pvBtnSave')?.addEventListener('click', () => saveProfilePatch());
        document.getElementById('pvBtnDelete')?.addEventListener('click', () => openDeleteModal());

        document.getElementById('pvModalDeleteClose')?.addEventListener('click', closeDeleteModal);
        document.getElementById('pvModalDeleteCancel')?.addEventListener('click', closeDeleteModal);
        document.getElementById('pvModalDeleteSubmit')?.addEventListener('click', () => submitDeleteUser());
        document.getElementById('pvDeleteConfirmInput')?.addEventListener('input', syncDeleteConfirmButton);
        document.getElementById('pvDeleteHard')?.addEventListener('change', function () {
            updateDeleteModalUi();
            syncDeleteConfirmButton();
        });
        document.getElementById('pvModalDeleteBackdrop')?.addEventListener('click', function (ev) {
            if (ev.target === ev.currentTarget) closeDeleteModal();
        });

        updateDetailActionButtons();
        renderUserTree();
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', bind);
    } else {
        bind();
    }
})();
