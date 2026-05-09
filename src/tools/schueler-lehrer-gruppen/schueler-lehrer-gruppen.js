(function () {
    'use strict';

    const STORAGE_KEY = 'ms365-schueler-lehrer-gruppen-v1';

    const GRAPH_SCOPES = [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/User.Read.All',
        'https://graph.microsoft.com/Group.ReadWrite.All',
        /** Optional: Team aus M365-Gruppe erstellen (PUT /groups/{id}/team) */
        'https://graph.microsoft.com/Team.Create'
    ];

    const PERSON_SELECT = 'id,displayName,mail,userPrincipalName';

    let msalMod = null;
    let pca = null;
    let slgCurrentStep = 1;

    /** @type {string | null} */
    let resolvedSchuelerId = null;
    /** @type {string | null} */
    let resolvedLehrerId = null;

    function toast(msg) {
        const el = document.getElementById('toast');
        if (el) {
            el.textContent = msg;
            el.classList.add('show');
            clearTimeout(toast._t);
            toast._t = setTimeout(function () {
                el.classList.remove('show');
            }, 3800);
        } else if (typeof window.ms365ToastOrAlert === 'function') {
            window.ms365ToastOrAlert(msg);
        } else if (typeof window.ms365ShowToast === 'function') {
            window.ms365ShowToast(msg);
        } else {
            window.alert(msg);
        }
    }

    function normStr(v) {
        return String(v ?? '').trim();
    }

    function normEmail(v) {
        return normStr(v).toLowerCase();
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

    async function graphRequest(method, path, token, body, extraHeaders) {
        const url = path.indexOf('http') === 0 ? path : 'https://graph.microsoft.com/v1.0' + path;
        let attempt = 0;
        while (true) {
            const headers = { Authorization: 'Bearer ' + token };
            if (extraHeaders && typeof extraHeaders === 'object') {
                for (const k in extraHeaders) {
                    if (Object.prototype.hasOwnProperty.call(extraHeaders, k)) {
                        headers[k] = extraHeaders[k];
                    }
                }
            }
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

    async function graphJson(method, path, token, body, extraHeaders) {
        const res = await graphRequest(method, path, token, body, extraHeaders);
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
            throw new Error(method + ' ' + path + ': ' + msg);
        }
        return data || {};
    }

    function odataEscape(s) {
        return String(s).replace(/'/g, "''");
    }

    function guidLooksValid(s) {
        return /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(
            String(s || '').trim()
        );
    }

    function sanitizeMailNickname(name) {
        let n = String(name || '')
            .replace(/[^0-9a-zA-Z]/g, '')
            .slice(0, 60);
        if (!n) n = 'group';
        return n.toLowerCase();
    }

    function isUnifiedGroup(g) {
        const gt = g && g.groupTypes;
        return Array.isArray(gt) && gt.indexOf('Unified') !== -1;
    }

    function userRef(userId) {
        return 'https://graph.microsoft.com/v1.0/users/' + userId;
    }

    function isDuplicateMemberError(e) {
        const m = String((e && e.message) || e || '');
        return (
            m.indexOf('added object references already exist') !== -1 ||
            m.indexOf('One or more added object references already exist') !== -1 ||
            m.indexOf('already exist') !== -1
        );
    }

    function loadTenantSettings() {
        if (typeof window.ms365TenantSettingsLoad !== 'function') {
            return null;
        }
        return window.ms365TenantSettingsLoad();
    }

    function getSchoolDomainNoAt() {
        const s = loadTenantSettings();
        const d = s && s.domain ? normStr(s.domain) : '';
        if (d) return d;
        if (typeof window.ms365GetSchoolDomainNoAt === 'function') {
            return normStr(window.ms365GetSchoolDomainNoAt());
        }
        return '';
    }

    function collectStudentEmails(settings) {
        const out = [];
        const seen = new Set();
        const students = settings && Array.isArray(settings.students) ? settings.students : [];
        students.forEach(function (row) {
            const em = normEmail(row && row.email);
            if (!em || em.indexOf('@') === -1) return;
            if (seen.has(em)) return;
            seen.add(em);
            out.push(em);
        });
        return out;
    }

    function collectTeacherEmails(settings) {
        const out = [];
        const seen = new Set();
        const teachers = settings && Array.isArray(settings.teachers) ? settings.teachers : [];
        teachers.forEach(function (row) {
            const em = normEmail(row && row.email);
            if (!em || em.indexOf('@') === -1) return;
            if (seen.has(em)) return;
            seen.add(em);
            out.push(em);
        });
        return out;
    }

    /** Rollen aus Schul‑Einstellungen → Verwaltung, die als M365‑Gruppenbesitzer (Direktion) gelten. */
    function isDirektionRole(roleRaw) {
        const r = normStr(roleRaw).toLowerCase();
        if (!r) return false;
        return r.indexOf('direktion') !== -1 || r.indexOf('direktor') !== -1;
    }

    function collectDirektionOwnerEmails(settings) {
        const out = [];
        const seen = new Set();
        const admin = settings && Array.isArray(settings.admin) ? settings.admin : [];
        admin.forEach(function (row) {
            if (!isDirektionRole(row && row.role)) return;
            const em = normEmail(row && row.email);
            if (!em || em.indexOf('@') === -1) return;
            if (seen.has(em)) return;
            seen.add(em);
            out.push(em);
        });
        return out;
    }

    function countStudentsWithAnyData(settings) {
        const students = settings && Array.isArray(settings.students) ? settings.students : [];
        let n = 0;
        students.forEach(function (row) {
            const klasse = normStr(row && (row.klasse || row.class));
            const name = normStr(row && row.name);
            const email = normEmail(row && row.email);
            if (klasse || name || email) n++;
        });
        return n;
    }

    function countTeachers(settings) {
        const teachers = settings && Array.isArray(settings.teachers) ? settings.teachers : [];
        return teachers.length;
    }

    function refreshStep1Ui() {
        const settings = loadTenantSettings();
        const domain = getSchoolDomainNoAt();
        const elDom = document.getElementById('slgDomainPreview');
        const elMS = document.getElementById('slgMailSchuelerPreview');
        const elML = document.getElementById('slgMailLehrerPreview');
        const st = document.getElementById('slgStatStudents');
        const stM = document.getElementById('slgStatStudentsMail');
        const tt = document.getElementById('slgStatTeachers');
        const ttM = document.getElementById('slgStatTeachersMail');
        const warn = document.getElementById('slgTenantWarn');

        if (elDom) elDom.textContent = domain || '(keine Domain in den Schul‑Einstellungen)';
        if (elMS) elMS.textContent = domain ? 'schueler@' + domain : 'schueler@…';
        if (elML) elML.textContent = domain ? 'lehrer@' + domain : 'lehrer@…';

        const studEmails = collectStudentEmails(settings);
        const teachEmails = collectTeacherEmails(settings);
        if (st) st.textContent = String(countStudentsWithAnyData(settings));
        if (stM) stM.textContent = String(studEmails.length);
        if (tt) tt.textContent = String(countTeachers(settings));
        if (ttM) ttM.textContent = String(teachEmails.length);

        if (warn) {
            const lines = [];
            if (!domain) lines.push('Bitte in den Schul‑Einstellungen eine Schul‑Domain eintragen (für die Adress‑Vorschau).');
            if (!studEmails.length) lines.push('Keine Schüler:innen mit E‑Mail in der Liste – Schritt 3 kann dort nichts übernehmen.');
            if (!teachEmails.length) lines.push('Keine Lehrer:innen mit E‑Mail in der Liste – Schritt 3 kann dort nichts übernehmen.');
            const dirEm = collectDirektionOwnerEmails(settings);
            if (!dirEm.length) {
                lines.push(
                    'In der Verwaltungsliste ist keine Rolle „Direktion“ (o. Ä.) mit E‑Mail hinterlegt – beim Anlegen der Gruppen wird sonst der angemeldete Benutzer als Besitzer gesetzt.'
                );
            }
            warn.style.display = lines.length ? 'block' : 'none';
            warn.innerHTML = lines.length ? '<strong>Hinweis:</strong> ' + lines.join(' ') : '';
        }
    }

    function slgStepNum(el) {
        const raw = el.getAttribute('data-slg-step');
        const n = parseFloat(String(raw || '').trim());
        return Number.isFinite(n) ? n : NaN;
    }

    function goToSlgStep(step) {
        slgCurrentStep = step;
        document.querySelectorAll('.slg-step-content').forEach(function (el) {
            el.classList.toggle('active', slgStepNum(el) === step);
        });
        document.querySelectorAll('.slg-steps .step').forEach(function (el) {
            const s = slgStepNum(el);
            el.classList.toggle('active', s === step);
            el.classList.toggle('completed', s < step);
        });
        if (typeof window.ms365ApplyStepProgress === 'function') {
            window.ms365ApplyStepProgress(document.querySelector('.slg-steps'), step, [1, 2, 3, 4]);
        }
        if (step === 4) {
            updateEntraLinks();
        }
    }

    function toggleModeBlocks() {
        const schuelerNew = document.querySelector('input[name="slgSchuelerMode"][value="new"]');
        const schuelerIsNew = schuelerNew && schuelerNew.checked;
        const b1 = document.getElementById('slgSchuelerNewBlock');
        const b2 = document.getElementById('slgSchuelerExistBlock');
        if (b1) b1.style.display = schuelerIsNew ? 'block' : 'none';
        if (b2) b2.style.display = schuelerIsNew ? 'none' : 'block';

        const lehrerNew = document.querySelector('input[name="slgLehrerMode"][value="new"]');
        const lehrerIsNew = lehrerNew && lehrerNew.checked;
        const l1 = document.getElementById('slgLehrerNewBlock');
        const l2 = document.getElementById('slgLehrerExistBlock');
        if (l1) l1.style.display = lehrerIsNew ? 'block' : 'none';
        if (l2) l2.style.display = lehrerIsNew ? 'none' : 'block';
    }

    function setSummary(kind, html, show) {
        const id = kind === 'schueler' ? 'slgSchuelerSummary' : 'slgLehrerSummary';
        const el = document.getElementById(id);
        if (!el) return;
        el.style.display = show ? 'block' : 'none';
        el.innerHTML = html || '';
    }

    async function fetchGroup(token, id) {
        const path =
            '/groups/' +
            encodeURIComponent(id) +
            '?$select=' +
            encodeURIComponent('id,displayName,mail,mailNickname,groupTypes,mailEnabled,securityEnabled');
        return graphJson('GET', path, token, undefined);
    }

    async function findGroupsByMailNickname(token, nickname) {
        const esc = odataEscape(nickname);
        const filter = "mailNickname eq '" + esc + "'";
        const path =
            '/groups?$filter=' +
            encodeURIComponent(filter) +
            '&$select=' +
            encodeURIComponent('id,displayName,mail,mailNickname,groupTypes') +
            '&$top=15';
        const data = await graphJson('GET', path, token, undefined);
        return data.value || [];
    }

    async function findGroupsByMail(token, mail) {
        const esc = odataEscape(mail);
        const filter = "mail eq '" + esc + "'";
        const path =
            '/groups?$filter=' +
            encodeURIComponent(filter) +
            '&$select=' +
            encodeURIComponent('id,displayName,mail,mailNickname,groupTypes') +
            '&$top=15';
        const data = await graphJson('GET', path, token, undefined);
        return data.value || [];
    }

    function escapeSearchPhrase(raw) {
        return String(raw || '')
            .replace(/"/g, '\\"')
            .replace(/\r?\n/g, ' ')
            .trim();
    }

    async function searchUnifiedGroups(token, queryRaw) {
        const q = normStr(queryRaw);
        if (!q) return [];

        // 1) Volltextsuche (intuitiver): displayName, mail, mailNickname, description.
        // Graph benötigt dafür ConsistencyLevel: eventual.
        try {
            const phrase = escapeSearchPhrase(q);
            const aqs =
                '(displayName:' +
                phrase +
                ' OR mail:' +
                phrase +
                ' OR mailNickname:' +
                phrase +
                ' OR description:' +
                phrase +
                ')';
            const path =
                '/groups?$search=' +
                encodeURIComponent('"' + aqs + '"') +
                '&$select=' +
                encodeURIComponent('id,displayName,mail,mailNickname,groupTypes,description') +
                '&$top=25';
            const data = await graphJson('GET', path, token, undefined, { ConsistencyLevel: 'eventual' });
            const list = (data && data.value) || [];
            return list.filter(isUnifiedGroup);
        } catch {
            // 2) Fallback: StartsWith-Filter (funktioniert auch ohne ConsistencyLevel)
        }

        const esc = odataEscape(q);
        const filter =
            "groupTypes/any(c:c eq 'Unified') and (" +
            "startswith(displayName,'" +
            esc +
            "') or startswith(mailNickname,'" +
            esc +
            "') or startswith(mail,'" +
            esc +
            "') )";
        const path =
            '/groups?$filter=' +
            encodeURIComponent(filter) +
            '&$select=' +
            encodeURIComponent('id,displayName,mail,mailNickname,groupTypes') +
            '&$top=25';
        const data = await graphJson('GET', path, token, undefined);
        return data.value || [];
    }

    function normalizeGroupQueryToNickOrMail(raw) {
        const q = normStr(raw);
        if (!q) return { nick: '', mail: '' };
        if (q.indexOf('@') !== -1) {
            const mail = normEmail(q);
            const local = mail.split('@')[0] || '';
            return { nick: sanitizeMailNickname(local), mail: mail };
        }
        return { nick: sanitizeMailNickname(q), mail: '' };
    }

    async function createUnifiedGroup(token, displayName, mailNickname, description) {
        const nick = sanitizeMailNickname(mailNickname);
        const body = {
            displayName: String(displayName).trim(),
            description: description || 'MS365-Schulverwaltung – Schüler:innen/Lehrer:innen',
            mailNickname: nick,
            mailEnabled: true,
            securityEnabled: false,
            groupTypes: ['Unified'],
            visibility: 'Private'
        };
        const group = await graphJson('POST', '/groups', token, body);
        const gid = group.id;
        await sleep(1500);
        let direktionOwnersAdded = 0;
        try {
            const settings = loadTenantSettings();
            const dirEmails = collectDirektionOwnerEmails(settings);
            for (let i = 0; i < dirEmails.length; i++) {
                const em = dirEmails[i];
                try {
                    const u = await resolveUserByEmail(token, em);
                    if (!u || !u.id) continue;
                    try {
                        await graphJson('POST', '/groups/' + encodeURIComponent(gid) + '/owners/$ref', token, {
                            '@odata.id': userRef(u.id)
                        });
                        direktionOwnersAdded++;
                    } catch (e) {
                        if (!isDuplicateMemberError(e)) {
                            /* einzelne Owner-Fehler nicht fatal */
                        }
                    }
                } catch {
                    /* ignore */
                }
                if ((i + 1) % 6 === 0) await sleep(120);
            }
        } catch {
            /* Direktion optional */
        }
        try {
            const me = await graphJson('GET', '/me', token, undefined);
            const meId = me && me.id;
            if (meId) {
                if (direktionOwnersAdded === 0) {
                    try {
                        await graphJson('POST', '/groups/' + encodeURIComponent(gid) + '/owners/$ref', token, {
                            '@odata.id': userRef(meId)
                        });
                    } catch (e) {
                        if (!isDuplicateMemberError(e)) throw e;
                    }
                    try {
                        await graphJson('POST', '/groups/' + encodeURIComponent(gid) + '/members/$ref', token, {
                            '@odata.id': userRef(meId)
                        });
                    } catch (e) {
                        if (!isDuplicateMemberError(e)) throw e;
                    }
                }
            }
        } catch (e) {
            /* Besitzer optional */
        }
        return group;
    }

    function buildPutTeamBody() {
        return {
            memberSettings: {
                allowCreatePrivateChannels: true,
                allowCreateUpdateChannels: true
            },
            messagingSettings: {
                allowUserEditMessages: true,
                allowUserDeleteMessages: true
            },
            funSettings: {
                allowGiphy: true,
                giphyContentRating: 'moderate'
            },
            guestSettings: {
                allowCreateUpdateChannels: false
            }
        };
    }

    async function provisionTeamForGroup(token, gid) {
        const teamUri = '/groups/' + encodeURIComponent(gid) + '/team';
        for (let i = 0; i < 8; i++) {
            try {
                await graphJson('PUT', teamUri, token, buildPutTeamBody());
                return;
            } catch (e) {
                const msg = String(e && e.message ? e.message : e);
                const looksLikeReplication = msg.indexOf('404') !== -1 || msg.indexOf('Request_ResourceNotFound') !== -1;
                if (i < 7 && looksLikeReplication) {
                    await sleep(10000);
                    token = await getGraphToken();
                    continue;
                }
                throw e;
            }
        }
    }

    async function resolveUserByEmail(token, email) {
        const em = normEmail(email);
        if (!em || em.indexOf('@') === -1) return null;
        const esc = odataEscape(em);
        const filter = "(mail eq '" + esc + "' or userPrincipalName eq '" + esc + "')";
        const path =
            '/users?$filter=' +
            encodeURIComponent(filter) +
            '&$select=' +
            encodeURIComponent(PERSON_SELECT) +
            '&$top=5';
        const data = await graphJson('GET', path, token, undefined);
        const list = data.value || [];
        return list[0] || null;
    }

    async function graphAddMember(token, groupId, userId) {
        await graphJson(
            'POST',
            '/groups/' + encodeURIComponent(groupId) + '/members/$ref',
            token,
            {
                '@odata.id': userRef(userId)
            }
        );
    }

    async function ensureDirektionOwnersOnGroup(token, groupId, logLabel) {
        const label = logLabel || 'Besitzer (Direktion)';
        const emails = collectDirektionOwnerEmails(loadTenantSettings());
        let ok = 0;
        for (let i = 0; i < emails.length; i++) {
            const em = emails[i];
            try {
                const u = await resolveUserByEmail(token, em);
                if (!u || !u.id) {
                    appendSyncLog(label + ': Kein Benutzer für ' + em, 'warn');
                    continue;
                }
                try {
                    await graphJson('POST', '/groups/' + encodeURIComponent(groupId) + '/owners/$ref', token, {
                        '@odata.id': userRef(u.id)
                    });
                    ok++;
                    appendSyncLog(label + ': ' + em + ' → Besitzer', 'ok');
                } catch (e) {
                    if (isDuplicateMemberError(e)) {
                        appendSyncLog(label + ': ' + em + ' (war schon Besitzer)', 'warn');
                    } else {
                        appendSyncLog(label + ': ' + em + ' — ' + (e.message || e), 'err');
                    }
                }
            } catch (e) {
                appendSyncLog(label + ': ' + em + ' — ' + (e.message || e), 'err');
            }
            if ((i + 1) % 6 === 0) await sleep(120);
        }
        return ok;
    }

    function appendSyncLog(msg, kind) {
        const el = document.getElementById('slgSyncLog');
        if (!el) return;
        const line = document.createElement('div');
        line.textContent = new Date().toLocaleTimeString() + '  ' + msg;
        if (kind === 'err') line.style.color = '#b00020';
        else if (kind === 'ok') line.style.color = '#0d8050';
        else if (kind === 'warn') line.style.color = '#856404';
        el.appendChild(line);
        el.scrollTop = el.scrollHeight;
    }

    function clearSyncLog() {
        const el = document.getElementById('slgSyncLog');
        if (el) el.replaceChildren();
    }

    function updateEntraLinks() {
        const a1 = document.getElementById('slgLinkSchuelerEntra');
        const a2 = document.getElementById('slgLinkLehrerEntra');
        const sep = document.getElementById('slgLinkSep');
        const base = 'https://entra.microsoft.com/#view/Microsoft_AAD_IAM/GroupDetailsMenuBlade/~/Members/groupId/';
        if (a1 && resolvedSchuelerId) {
            a1.href = base + encodeURIComponent(resolvedSchuelerId);
            a1.style.display = 'inline';
        } else if (a1) {
            a1.style.display = 'none';
        }
        if (a2 && resolvedLehrerId) {
            a2.href = base + encodeURIComponent(resolvedLehrerId);
            a2.style.display = 'inline';
        } else if (a2) {
            a2.style.display = 'none';
        }
        if (sep) {
            sep.style.display = resolvedSchuelerId && resolvedLehrerId ? 'inline' : 'none';
        }
    }

    function persistResolvedIds(kind, group) {
        if (kind === 'schueler') {
            resolvedSchuelerId = group && group.id ? String(group.id) : null;
        } else {
            resolvedLehrerId = group && group.id ? String(group.id) : null;
        }
    }

    function formatGroupSummary(g) {
        if (!g || !g.id) return '';
        const unified = isUnifiedGroup(g) ? 'Microsoft 365‑Gruppe (Unified)' : 'Keine Unified‑Gruppe';
        const mail = normStr(g.mail) || '–';
        const nick = normStr(g.mailNickname) || '–';
        return (
            '<strong>OK:</strong> ' +
            normStr(g.displayName) +
            '<br>Object‑ID: <code>' +
            g.id +
            '</code><br>Mail‑Nickname: <code>' +
            nick +
            '</code> · SMTP: ' +
            mail +
            '<br><span style="color:#084298;">' +
            unified +
            '</span>'
        );
    }

    function clearSearchResults(kind) {
        const el = document.getElementById(kind === 'schueler' ? 'slgSchuelerSearchResults' : 'slgLehrerSearchResults');
        if (!el) return;
        el.style.display = 'none';
        el.replaceChildren();
    }

    function renderSearchResults(kind, list) {
        const el = document.getElementById(kind === 'schueler' ? 'slgSchuelerSearchResults' : 'slgLehrerSearchResults');
        if (!el) return;
        el.replaceChildren();

        const box = document.createElement('div');
        box.style.border = '1px solid #ced4da';
        box.style.borderRadius = '10px';
        box.style.background = '#fff';
        box.style.overflow = 'hidden';

        const head = document.createElement('div');
        head.style.padding = '10px 12px';
        head.style.background = '#f8f9fa';
        head.style.display = 'flex';
        head.style.alignItems = 'center';
        head.style.justifyContent = 'space-between';
        head.style.gap = '10px';
        head.innerHTML =
            '<strong>Treffer</strong> <span style="color:#6c757d;font-size:0.9em;">' +
            String(list.length) +
            '</span>';

        const btnClose = document.createElement('button');
        btnClose.type = 'button';
        btnClose.className = 'btn';
        btnClose.textContent = 'Schließen';
        btnClose.addEventListener('click', function () {
            clearSearchResults(kind);
        });
        head.appendChild(btnClose);

        box.appendChild(head);

        const body = document.createElement('div');
        body.style.maxHeight = '240px';
        body.style.overflow = 'auto';

        if (!list.length) {
            const empty = document.createElement('div');
            empty.style.padding = '10px 12px';
            empty.style.color = '#6c757d';
            empty.textContent = 'Keine passenden Microsoft 365‑Gruppen (Unified) gefunden.';
            body.appendChild(empty);
        } else {
            list.forEach(function (g, idx) {
                const row = document.createElement('div');
                row.style.display = 'grid';
                row.style.gridTemplateColumns = '1fr auto';
                row.style.gap = '10px';
                row.style.padding = '10px 12px';
                row.style.borderTop = idx === 0 ? '0' : '1px solid #eef1f4';
                row.style.alignItems = 'center';

                const left = document.createElement('div');
                const dn = normStr(g && g.displayName) || '(ohne Namen)';
                const mail = normStr(g && g.mail) || '–';
                const nick = normStr(g && g.mailNickname) || '–';
                left.innerHTML =
                    '<div style="font-weight:600;line-height:1.25;">' +
                    escapeHtml(dn) +
                    '</div>' +
                    '<div style="color:#6c757d;font-size:0.9em;line-height:1.35;">' +
                    'Mail‑Nickname: <code>' +
                    escapeHtml(nick) +
                    '</code> · SMTP: ' +
                    escapeHtml(mail) +
                    '</div>' +
                    '<div style="color:#6c757d;font-size:0.85em;line-height:1.35;">' +
                    'Object‑ID: <code>' +
                    escapeHtml(g && g.id ? g.id : '') +
                    '</code>' +
                    '</div>';

                const btn = document.createElement('button');
                btn.type = 'button';
                btn.className = 'btn btn-success';
                btn.textContent = 'Übernehmen';
                btn.addEventListener('click', function () {
                    const inp = document.getElementById(kind === 'schueler' ? 'slgSchuelerGroupId' : 'slgLehrerGroupId');
                    if (inp && g && g.id) inp.value = String(g.id);
                    persistResolvedIds(kind, g);
                    setSummary(kind, formatGroupSummary(g), true);
                    clearSearchResults(kind);
                    toast('Gruppe übernommen.');
                });

                row.appendChild(left);
                row.appendChild(btn);
                body.appendChild(row);
            });
        }

        box.appendChild(body);
        el.appendChild(box);
        el.style.display = 'block';
    }

    async function handleSearchExistingGroups(kind) {
        try {
            const token = await getGraphToken();
            const qEl = document.getElementById(kind === 'schueler' ? 'slgSchuelerGroupQuery' : 'slgLehrerGroupQuery');
            const q = qEl ? qEl.value : '';
            if (!normStr(q)) {
                toast('Bitte oben eine Suche eingeben (Name, Mail‑Nickname oder Gruppen‑E‑Mail).');
                return;
            }
            const list = await searchUnifiedGroups(token, q);
            renderSearchResults(kind, list);
            if (!list.length) toast('Keine passenden Gruppen gefunden.');
        } catch (e) {
            toast('Fehler: ' + (e.message || e));
        }
    }

    async function handleCreateSchueler() {
        const dn = document.getElementById('slgSchuelerDisplayName');
        const nn = document.getElementById('slgSchuelerMailNick');
        const ct = document.getElementById('slgSchuelerCreateTeam');
        const displayName = dn ? dn.value : 'Schüler:innen';
        const mailNick = nn ? nn.value : 'schueler';
        try {
            let token = await getGraphToken();
            const g = await createUnifiedGroup(
                token,
                displayName,
                mailNick,
                'Alle Schüler:innen (MS365-Schulverwaltung / Schul‑Liste)'
            );
            persistResolvedIds('schueler', g);
            setSummary('schueler', formatGroupSummary(g), true);
            if (ct && ct.checked && g && g.id) {
                toast('Gruppe angelegt – Team wird bereitgestellt …');
                await sleep(1500);
                await provisionTeamForGroup(token, g.id);
                toast('Schüler:innen‑Gruppe + Team angelegt.');
            } else {
                toast('Schüler:innen‑Gruppe angelegt.');
            }
        } catch (e) {
            toast('Fehler: ' + (e.message || e));
        }
    }

    async function handleCreateLehrer() {
        const dn = document.getElementById('slgLehrerDisplayName');
        const nn = document.getElementById('slgLehrerMailNick');
        const ct = document.getElementById('slgLehrerCreateTeam');
        const displayName = dn ? dn.value : 'Lehrer:innen';
        const mailNick = nn ? nn.value : 'lehrer';
        try {
            let token = await getGraphToken();
            const g = await createUnifiedGroup(
                token,
                displayName,
                mailNick,
                'Alle Lehrer:innen (MS365-Schulverwaltung / Schul‑Liste)'
            );
            persistResolvedIds('lehrer', g);
            setSummary('lehrer', formatGroupSummary(g), true);
            if (ct && ct.checked && g && g.id) {
                toast('Gruppe angelegt – Team wird bereitgestellt …');
                await sleep(1500);
                await provisionTeamForGroup(token, g.id);
                toast('Lehrer:innen‑Gruppe + Team angelegt.');
            } else {
                toast('Lehrer:innen‑Gruppe angelegt.');
            }
        } catch (e) {
            toast('Fehler: ' + (e.message || e));
        }
    }

    async function handleResolveSchueler() {
        const inp = document.getElementById('slgSchuelerGroupId');
        const id = inp ? normStr(inp.value) : '';
        if (!guidLooksValid(id)) {
            toast('Bitte eine gültige Object‑ID (GUID) eintragen.');
            return;
        }
        try {
            const token = await getGraphToken();
            const g = await fetchGroup(token, id);
            if (!isUnifiedGroup(g)) {
                setSummary(
                    'schueler',
                    '<strong>Warnung:</strong> Diese Gruppe ist keine Microsoft 365‑Gruppe (Unified). Bitte eine passende Gruppe wählen.',
                    true
                );
                persistResolvedIds('schueler', g);
                return;
            }
            persistResolvedIds('schueler', g);
            setSummary('schueler', formatGroupSummary(g), true);
            toast('Schüler:innen‑Gruppe geladen.');
        } catch (e) {
            toast('Fehler: ' + (e.message || e));
        }
    }

    async function handleResolveLehrer() {
        const inp = document.getElementById('slgLehrerGroupId');
        const id = inp ? normStr(inp.value) : '';
        if (!guidLooksValid(id)) {
            toast('Bitte eine gültige Object‑ID (GUID) eintragen.');
            return;
        }
        try {
            const token = await getGraphToken();
            const g = await fetchGroup(token, id);
            if (!isUnifiedGroup(g)) {
                setSummary(
                    'lehrer',
                    '<strong>Warnung:</strong> Diese Gruppe ist keine Microsoft 365‑Gruppe (Unified). Bitte eine passende Gruppe wählen.',
                    true
                );
                persistResolvedIds('lehrer', g);
                return;
            }
            persistResolvedIds('lehrer', g);
            setSummary('lehrer', formatGroupSummary(g), true);
            toast('Lehrer:innen‑Gruppe geladen.');
        } catch (e) {
            toast('Fehler: ' + (e.message || e));
        }
    }

    async function handleFindSchuelerNick() {
        try {
            const token = await getGraphToken();
            const qEl = document.getElementById('slgSchuelerGroupQuery');
            const fallbackNickEl = document.getElementById('slgSchuelerMailNick');
            const q = normalizeGroupQueryToNickOrMail(qEl ? qEl.value : '');
            const nick = q.nick || sanitizeMailNickname(fallbackNickEl ? fallbackNickEl.value : 'schueler') || 'schueler';

            let list = [];
            if (q.mail) {
                list = await findGroupsByMail(token, q.mail);
            }
            if (!list.length && nick) {
                list = await findGroupsByMailNickname(token, nick);
            }
            if (!list.length) {
                toast('Keine Gruppe gefunden (Suche: ' + (q.mail ? q.mail : nick) + ').');
                return;
            }
            const g = list[0];
            const inp = document.getElementById('slgSchuelerGroupId');
            if (inp) inp.value = g.id;
            persistResolvedIds('schueler', g);
            if (!isUnifiedGroup(g)) {
                setSummary(
                    'schueler',
                    '<strong>Warnung:</strong> Gefundene Gruppe ist keine Unified‑Gruppe. ' + formatGroupSummary(g),
                    true
                );
                return;
            }
            setSummary('schueler', formatGroupSummary(g), true);
            toast(list.length > 1 ? 'Mehrere Treffer – erste Gruppe übernommen.' : 'Gruppe gefunden.');
        } catch (e) {
            toast('Fehler: ' + (e.message || e));
        }
    }

    async function handleFindLehrerNick() {
        try {
            const token = await getGraphToken();
            const qEl = document.getElementById('slgLehrerGroupQuery');
            const fallbackNickEl = document.getElementById('slgLehrerMailNick');
            const q = normalizeGroupQueryToNickOrMail(qEl ? qEl.value : '');
            const nick = q.nick || sanitizeMailNickname(fallbackNickEl ? fallbackNickEl.value : 'lehrer') || 'lehrer';

            let list = [];
            if (q.mail) {
                list = await findGroupsByMail(token, q.mail);
            }
            if (!list.length && nick) {
                list = await findGroupsByMailNickname(token, nick);
            }
            if (!list.length) {
                toast('Keine Gruppe gefunden (Suche: ' + (q.mail ? q.mail : nick) + ').');
                return;
            }
            const g = list[0];
            const inp = document.getElementById('slgLehrerGroupId');
            if (inp) inp.value = g.id;
            persistResolvedIds('lehrer', g);
            if (!isUnifiedGroup(g)) {
                setSummary(
                    'lehrer',
                    '<strong>Warnung:</strong> Gefundene Gruppe ist keine Unified‑Gruppe. ' + formatGroupSummary(g),
                    true
                );
                return;
            }
            setSummary('lehrer', formatGroupSummary(g), true);
            toast(list.length > 1 ? 'Mehrere Treffer – erste Gruppe übernommen.' : 'Gruppe gefunden.');
        } catch (e) {
            toast('Fehler: ' + (e.message || e));
        }
    }

    async function syncEmailsToGroup(token, groupId, emails, label) {
        let ok = 0;
        let skip = 0;
        let fail = 0;
        for (let i = 0; i < emails.length; i++) {
            const em = emails[i];
            try {
                const u = await resolveUserByEmail(token, em);
                if (!u || !u.id) {
                    appendSyncLog(label + ': Kein Benutzer für ' + em, 'warn');
                    fail++;
                    continue;
                }
                try {
                    await graphAddMember(token, groupId, u.id);
                    ok++;
                    appendSyncLog(label + ': ' + em + ' → Mitglied', 'ok');
                } catch (e) {
                    if (isDuplicateMemberError(e)) {
                        skip++;
                        appendSyncLog(label + ': ' + em + ' (war schon Mitglied)', 'warn');
                    } else {
                        fail++;
                        appendSyncLog(label + ': ' + em + ' — ' + (e.message || e), 'err');
                    }
                }
            } catch (e) {
                fail++;
                appendSyncLog(label + ': ' + em + ' — ' + (e.message || e), 'err');
            }
            if ((i + 1) % 8 === 0) {
                await sleep(120);
            }
        }
        return { ok: ok, skip: skip, fail: fail };
    }

    async function handleSyncStudents() {
        const settings = loadTenantSettings();
        const emails = collectStudentEmails(settings);
        if (!emails.length) {
            toast('Keine Schüler:innen‑E‑Mails in den Schul‑Einstellungen.');
            return;
        }
        if (!resolvedSchuelerId) {
            toast('Zuerst in Schritt 2 die Schüler:innen‑Gruppe anlegen oder auswählen.');
            return;
        }
        clearSyncLog();
        appendSyncLog('Start: Schüler:innen (' + emails.length + ' Adressen) …', '');
        try {
            let token = await getGraphToken();
            const r = await syncEmailsToGroup(token, resolvedSchuelerId, emails, 'Schüler');
            appendSyncLog(
                'Fertig Schüler:innen: neu ' + r.ok + ', übersprungen ' + r.skip + ', Fehler ' + r.fail + '.',
                'ok'
            );
            await ensureDirektionOwnersOnGroup(token, resolvedSchuelerId, 'Schüler‑Gruppe Besitzer');
            toast('Synchronisation Schüler:innen abgeschlossen.');
        } catch (e) {
            appendSyncLog('Abbruch: ' + (e.message || e), 'err');
            toast('Fehler: ' + (e.message || e));
        }
    }

    async function handleSyncTeachers() {
        const settings = loadTenantSettings();
        const emails = collectTeacherEmails(settings);
        if (!emails.length) {
            toast('Keine Lehrer:innen‑E‑Mails in den Schul‑Einstellungen.');
            return;
        }
        if (!resolvedLehrerId) {
            toast('Zuerst in Schritt 2 die Lehrer:innen‑Gruppe anlegen oder auswählen.');
            return;
        }
        clearSyncLog();
        appendSyncLog('Start: Lehrer:innen (' + emails.length + ' Adressen) …', '');
        try {
            const token = await getGraphToken();
            const r = await syncEmailsToGroup(token, resolvedLehrerId, emails, 'Lehrer');
            appendSyncLog(
                'Fertig Lehrer:innen: neu ' + r.ok + ', übersprungen ' + r.skip + ', Fehler ' + r.fail + '.',
                'ok'
            );
            await ensureDirektionOwnersOnGroup(token, resolvedLehrerId, 'Lehrer‑Gruppe Besitzer');
            toast('Synchronisation Lehrer:innen abgeschlossen.');
        } catch (e) {
            appendSyncLog('Abbruch: ' + (e.message || e), 'err');
            toast('Fehler: ' + (e.message || e));
        }
    }

    function parseManualLines(text) {
        const raw = String(text || '').split(/\r?\n/);
        const out = [];
        const seen = new Set();
        raw.forEach(function (line) {
            const p = normStr(line);
            if (!p || p.indexOf('@') === -1) return;
            const em = normEmail(p);
            if (seen.has(em)) return;
            seen.add(em);
            out.push(em);
        });
        return out;
    }

    async function handleManualAdd() {
        const sel = document.getElementById('slgManualTarget');
        const ta = document.getElementById('slgManualLines');
        const outEl = document.getElementById('slgManualResult');
        const kind = sel && sel.value === 'lehrer' ? 'lehrer' : 'schueler';
        const gid = kind === 'lehrer' ? resolvedLehrerId : resolvedSchuelerId;
        if (!gid) {
            toast('Keine ' + (kind === 'lehrer' ? 'Lehrer:innen' : 'Schüler:innen') + '‑Gruppe – bitte Schritt 2.');
            return;
        }
        const emails = parseManualLines(ta ? ta.value : '');
        if (!emails.length) {
            toast('Bitte mindestens eine gültige E‑Mail‑Zeile eintragen.');
            return;
        }
        try {
            const token = await getGraphToken();
            let ok = 0;
            let skip = 0;
            let fail = 0;
            const lines = [];
            for (let i = 0; i < emails.length; i++) {
                const em = emails[i];
                try {
                    const u = await resolveUserByEmail(token, em);
                    if (!u || !u.id) {
                        fail++;
                        lines.push(em + ' → nicht gefunden');
                        continue;
                    }
                    try {
                        await graphAddMember(token, gid, u.id);
                        ok++;
                        lines.push(em + ' → hinzugefügt');
                    } catch (e) {
                        if (isDuplicateMemberError(e)) {
                            skip++;
                            lines.push(em + ' → war schon Mitglied');
                        } else {
                            fail++;
                            lines.push(em + ' → Fehler: ' + (e.message || e));
                        }
                    }
                } catch (e) {
                    fail++;
                    lines.push(em + ' → ' + (e.message || e));
                }
            }
            if (outEl) {
                outEl.style.display = 'block';
                outEl.innerHTML =
                    '<strong>Ergebnis:</strong> neu ' +
                    ok +
                    ', übersprungen ' +
                    skip +
                    ', Fehler ' +
                    fail +
                    '.<br>' +
                    lines.map(function (x) {
                        return escapeHtml(x);
                    }).join('<br>');
            }
            toast('Manuelle Zuordnung abgeschlossen.');
        } catch (e) {
            toast('Fehler: ' + (e.message || e));
        }
    }

    function escapeHtml(s) {
        return String(s)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    function buildStateObject() {
        return {
            kind: 'ms365-schueler-lehrer-gruppen-v1',
            savedAt: new Date().toISOString(),
            step: slgCurrentStep,
            resolvedSchuelerId: resolvedSchuelerId,
            resolvedLehrerId: resolvedLehrerId,
            slgSchuelerDisplayName: document.getElementById('slgSchuelerDisplayName')
                ? document.getElementById('slgSchuelerDisplayName').value
                : '',
            slgLehrerDisplayName: document.getElementById('slgLehrerDisplayName')
                ? document.getElementById('slgLehrerDisplayName').value
                : '',
            slgSchuelerMailNick: document.getElementById('slgSchuelerMailNick')
                ? document.getElementById('slgSchuelerMailNick').value
                : 'schueler',
            slgLehrerMailNick: document.getElementById('slgLehrerMailNick')
                ? document.getElementById('slgLehrerMailNick').value
                : 'lehrer',
            slgSchuelerCreateTeam: document.getElementById('slgSchuelerCreateTeam')
                ? !!document.getElementById('slgSchuelerCreateTeam').checked
                : false,
            slgLehrerCreateTeam: document.getElementById('slgLehrerCreateTeam')
                ? !!document.getElementById('slgLehrerCreateTeam').checked
                : false,
            slgSchuelerGroupId: document.getElementById('slgSchuelerGroupId')
                ? document.getElementById('slgSchuelerGroupId').value
                : '',
            slgLehrerGroupId: document.getElementById('slgLehrerGroupId')
                ? document.getElementById('slgLehrerGroupId').value
                : '',
            slgSchuelerGroupQuery: document.getElementById('slgSchuelerGroupQuery')
                ? document.getElementById('slgSchuelerGroupQuery').value
                : '',
            slgLehrerGroupQuery: document.getElementById('slgLehrerGroupQuery')
                ? document.getElementById('slgLehrerGroupQuery').value
                : '',
            slgSchuelerMode: document.querySelector('input[name="slgSchuelerMode"]:checked')
                ? document.querySelector('input[name="slgSchuelerMode"]:checked').value
                : 'new',
            slgLehrerMode: document.querySelector('input[name="slgLehrerMode"]:checked')
                ? document.querySelector('input[name="slgLehrerMode"]:checked').value
                : 'new',
            slgManualLines: document.getElementById('slgManualLines') ? document.getElementById('slgManualLines').value : ''
        };
    }

    function applyStateObject(o) {
        if (!o || typeof o !== 'object') return;
        if (o.step !== undefined) {
            const s = parseInt(String(o.step), 10);
            if (s >= 1 && s <= 4) goToSlgStep(s);
            else goToSlgStep(1);
        } else {
            goToSlgStep(1);
        }
        resolvedSchuelerId = o.resolvedSchuelerId ? String(o.resolvedSchuelerId) : null;
        resolvedLehrerId = o.resolvedLehrerId ? String(o.resolvedLehrerId) : null;

        function setVal(id, v) {
            const el = document.getElementById(id);
            if (el && v !== undefined) el.value = String(v);
        }
        setVal('slgSchuelerDisplayName', o.slgSchuelerDisplayName);
        setVal('slgLehrerDisplayName', o.slgLehrerDisplayName);
        setVal('slgSchuelerMailNick', o.slgSchuelerMailNick);
        setVal('slgLehrerMailNick', o.slgLehrerMailNick);
        const ctS = document.getElementById('slgSchuelerCreateTeam');
        if (ctS && o.slgSchuelerCreateTeam !== undefined) ctS.checked = !!o.slgSchuelerCreateTeam;
        const ctL = document.getElementById('slgLehrerCreateTeam');
        if (ctL && o.slgLehrerCreateTeam !== undefined) ctL.checked = !!o.slgLehrerCreateTeam;
        setVal('slgSchuelerGroupId', o.slgSchuelerGroupId);
        setVal('slgLehrerGroupId', o.slgLehrerGroupId);
        setVal('slgSchuelerGroupQuery', o.slgSchuelerGroupQuery);
        setVal('slgLehrerGroupQuery', o.slgLehrerGroupQuery);
        setVal('slgManualLines', o.slgManualLines);

        if (o.slgSchuelerMode) {
            const r = document.querySelector('input[name="slgSchuelerMode"][value="' + o.slgSchuelerMode + '"]');
            if (r) r.checked = true;
        }
        if (o.slgLehrerMode) {
            const r = document.querySelector('input[name="slgLehrerMode"][value="' + o.slgLehrerMode + '"]');
            if (r) r.checked = true;
        }
        toggleModeBlocks();
        updateEntraLinks();
    }

    function saveState() {
        try {
            localStorage.setItem(STORAGE_KEY, JSON.stringify(buildStateObject()));
            toast('Zwischenstand gespeichert.');
        } catch (e) {
            toast('Speichern fehlgeschlagen: ' + (e.message || e));
        }
    }

    function loadState() {
        try {
            const raw = localStorage.getItem(STORAGE_KEY);
            if (!raw) {
                toast('Kein gespeicherter Stand.');
                return;
            }
            const o = JSON.parse(raw);
            applyStateObject(o);
            toast('Stand geladen.');
        } catch (e) {
            toast('Laden fehlgeschlagen: ' + (e.message || e));
        }
    }

    function clearStorage() {
        try {
            localStorage.removeItem(STORAGE_KEY);
            resolvedSchuelerId = null;
            resolvedLehrerId = null;
            setSummary('schueler', '', false);
            setSummary('lehrer', '', false);
            toast('Lokaler Speicher gelöscht.');
        } catch (e) {
            toast('Löschen fehlgeschlagen: ' + (e.message || e));
        }
    }

    function wire() {
        document.getElementById('slgBtnNext1') &&
            document.getElementById('slgBtnNext1').addEventListener('click', function () {
                goToSlgStep(2);
            });
        document.getElementById('slgBtnBack2') &&
            document.getElementById('slgBtnBack2').addEventListener('click', function () {
                goToSlgStep(1);
            });
        document.getElementById('slgBtnNext2') &&
            document.getElementById('slgBtnNext2').addEventListener('click', function () {
                goToSlgStep(3);
            });
        document.getElementById('slgBtnBack3') &&
            document.getElementById('slgBtnBack3').addEventListener('click', function () {
                goToSlgStep(2);
            });
        document.getElementById('slgBtnNext3') &&
            document.getElementById('slgBtnNext3').addEventListener('click', function () {
                goToSlgStep(4);
            });
        document.getElementById('slgBtnBack4') &&
            document.getElementById('slgBtnBack4').addEventListener('click', function () {
                goToSlgStep(3);
            });

        document.getElementById('slgBtnLogin') &&
            document.getElementById('slgBtnLogin').addEventListener('click', async function () {
                try {
                    await getGraphToken();
                    toast('Angemeldet.');
                } catch (e) {
                    toast('Anmeldung: ' + (e.message || e));
                }
            });

        document.getElementById('slgBtnCreateSchueler') &&
            document.getElementById('slgBtnCreateSchueler').addEventListener('click', function () {
                handleCreateSchueler();
            });
        document.getElementById('slgBtnCreateLehrer') &&
            document.getElementById('slgBtnCreateLehrer').addEventListener('click', function () {
                handleCreateLehrer();
            });
        document.getElementById('slgBtnResolveSchueler') &&
            document.getElementById('slgBtnResolveSchueler').addEventListener('click', function () {
                handleResolveSchueler();
            });
        document.getElementById('slgBtnResolveLehrer') &&
            document.getElementById('slgBtnResolveLehrer').addEventListener('click', function () {
                handleResolveLehrer();
            });
        document.getElementById('slgBtnFindSchuelerNick') &&
            document.getElementById('slgBtnFindSchuelerNick').addEventListener('click', function () {
                handleFindSchuelerNick();
            });
        document.getElementById('slgBtnFindLehrerNick') &&
            document.getElementById('slgBtnFindLehrerNick').addEventListener('click', function () {
                handleFindLehrerNick();
            });
        document.getElementById('slgBtnSearchSchuelerGroups') &&
            document.getElementById('slgBtnSearchSchuelerGroups').addEventListener('click', function () {
                handleSearchExistingGroups('schueler');
            });
        document.getElementById('slgBtnSearchLehrerGroups') &&
            document.getElementById('slgBtnSearchLehrerGroups').addEventListener('click', function () {
                handleSearchExistingGroups('lehrer');
            });

        document.querySelectorAll('input[name="slgSchuelerMode"]').forEach(function (r) {
            r.addEventListener('change', toggleModeBlocks);
        });
        document.querySelectorAll('input[name="slgLehrerMode"]').forEach(function (r) {
            r.addEventListener('change', toggleModeBlocks);
        });

        document.getElementById('slgBtnSyncStudents') &&
            document.getElementById('slgBtnSyncStudents').addEventListener('click', function () {
                handleSyncStudents();
            });
        document.getElementById('slgBtnSyncTeachers') &&
            document.getElementById('slgBtnSyncTeachers').addEventListener('click', function () {
                handleSyncTeachers();
            });

        document.getElementById('slgBtnManualAdd') &&
            document.getElementById('slgBtnManualAdd').addEventListener('click', function () {
                handleManualAdd();
            });

        document.getElementById('slgBtnSaveState') &&
            document.getElementById('slgBtnSaveState').addEventListener('click', saveState);
        document.getElementById('slgBtnLoadState') &&
            document.getElementById('slgBtnLoadState').addEventListener('click', loadState);
        document.getElementById('slgBtnClearStorage') &&
            document.getElementById('slgBtnClearStorage').addEventListener('click', clearStorage);
    }

    function init() {
        wire();
        refreshStep1Ui();
        toggleModeBlocks();
        try {
            const raw = localStorage.getItem(STORAGE_KEY);
            if (raw) {
                const o = JSON.parse(raw);
                applyStateObject(o);
            } else {
                goToSlgStep(1);
            }
        } catch {
            goToSlgStep(1);
        }
        updateEntraLinks();
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
