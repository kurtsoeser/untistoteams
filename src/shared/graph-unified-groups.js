(function () {
    'use strict';

    const MSAL_LOADER_IMPORT = (function () {
        const needle = 'graph-unified-groups.js';
        const rel = './msal-loader.js';
        const scripts = document.getElementsByTagName('script');
        for (let i = scripts.length - 1; i >= 0; i--) {
            const src = scripts[i].src || '';
            if (src.indexOf(needle) !== -1) {
                try {
                    return new URL(rel, src).href;
                } catch (_) {}
            }
        }
        try {
            return new URL('src/shared/msal-loader.js', document.baseURI).href;
        } catch (_) {
            return 'src/shared/msal-loader.js';
        }
    })();

    const GRAPH_SCOPES = [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/User.Read.All',
        'https://graph.microsoft.com/User.ReadWrite.All',
        'https://graph.microsoft.com/Group.ReadWrite.All',
        'https://graph.microsoft.com/Team.Create'
    ];

    const PERSON_SELECT = 'id,displayName,mail,userPrincipalName';

    let msalMod = null;
    let pca = null;

    function normStr(v) {
        return String(v ?? '').trim();
    }

    function normEmail(v) {
        return normStr(v).toLowerCase();
    }

    async function loadMsal() {
        if (msalMod) return msalMod;
        const loader = await import(MSAL_LOADER_IMPORT);
        if (typeof loader.loadMsalBrowser !== 'function') {
            throw new Error('MSAL-Loader: loadMsalBrowser fehlt.');
        }
        msalMod = await loader.loadMsalBrowser();
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
                Object.assign(headers, extraHeaders);
            }
            if (body !== undefined) headers['Content-Type'] = 'application/json';
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

    function sanitizeMailNickname(name) {
        let n = String(name || '')
            .replace(/[^0-9a-zA-Z]/g, '')
            .slice(0, 60);
        if (!n) n = 'group';
        return n.toLowerCase();
    }

    /** group.mailNickname: unzulässig u. a. laut Microsoft Learn (validateProperties): @ ( ) \ [ ] " ; : < > , Leerzeichen */
    const GRAPH_MAILNICKNAME_INVALID = /[@()[\]\\";:<>,\s]/;

    function sanitizeUnifiedGroupMailNickname(raw) {
        const s = String(raw ?? '')
            .trim()
            .toLowerCase();
        let out = '';
        for (let i = 0; i < s.length; i++) {
            const c = s.charCodeAt(i);
            if (c < 32 || c === 127 || c > 127) continue;
            const ch = s.charAt(i);
            if (GRAPH_MAILNICKNAME_INVALID.test(ch)) continue;
            out += ch;
        }
        if (!out) out = 'group';
        return out.slice(0, 60);
    }

    function isUnifiedGroup(g) {
        const gt = g && g.groupTypes;
        return Array.isArray(gt) && gt.indexOf('Unified') !== -1;
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
            // fallback
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
            encodeURIComponent('id,displayName,mail,mailNickname,groupTypes,description') +
            '&$top=25';
        const data = await graphJson('GET', path, token, undefined);
        return data.value || [];
    }

    async function fetchGroup(token, id) {
        const path =
            '/groups/' +
            encodeURIComponent(id) +
            '?$select=' +
            encodeURIComponent('id,displayName,mail,mailNickname,groupTypes,description');
        return graphJson('GET', path, token, undefined);
    }

    async function patchGroupDisplayName(token, groupId, displayName, description) {
        const body = {};
        if (displayName !== undefined && displayName !== null) {
            body.displayName = String(displayName).trim();
        }
        if (description !== undefined && description !== null) {
            body.description = String(description).trim();
        }
        if (!Object.keys(body).length) return {};
        return graphJson('PATCH', '/groups/' + encodeURIComponent(groupId), token, body);
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

    async function createUnifiedGroup(token, displayName, mailNickname, description) {
        const nick = sanitizeUnifiedGroupMailNickname(mailNickname);
        const body = {
            displayName: String(displayName).trim(),
            description: description || 'MS365-Schulverwaltung – Microsoft 365-Gruppe',
            mailNickname: nick,
            mailEnabled: true,
            securityEnabled: false,
            groupTypes: ['Unified'],
            visibility: 'Private'
        };
        const group = await graphJson('POST', '/groups', token, body);
        await sleep(1500);
        return group;
    }

    function buildPutTeamBody() {
        return {
            memberSettings: { allowCreatePrivateChannels: true, allowCreateUpdateChannels: true },
            messagingSettings: { allowUserEditMessages: true, allowUserDeleteMessages: true },
            funSettings: { allowGiphy: true, giphyContentRating: 'moderate' },
            guestSettings: { allowCreateUpdateChannels: false }
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

    async function graphAddMember(token, groupId, userId) {
        await graphJson('POST', '/groups/' + encodeURIComponent(groupId) + '/members/$ref', token, {
            '@odata.id': userRef(userId)
        });
    }

    async function ensureOwners(token, groupId, ownerEmails) {
        const emails = Array.isArray(ownerEmails) ? ownerEmails : [];
        let added = 0;
        for (let i = 0; i < emails.length; i++) {
            const em = emails[i];
            try {
                const u = await resolveUserByEmail(token, em);
                if (!u || !u.id) continue;
                try {
                    await graphJson('POST', '/groups/' + encodeURIComponent(groupId) + '/owners/$ref', token, {
                        '@odata.id': userRef(u.id)
                    });
                    added++;
                } catch (e) {
                    if (!isDuplicateMemberError(e)) {
                        // ignore einzelne Fehler
                    }
                }
            } catch {
                // ignore
            }
            if ((i + 1) % 6 === 0) await sleep(120);
        }
        if (added === 0) {
            try {
                const me = await graphJson('GET', '/me', token, undefined);
                const meId = me && me.id;
                if (meId) {
                    try {
                        await graphJson('POST', '/groups/' + encodeURIComponent(groupId) + '/owners/$ref', token, {
                            '@odata.id': userRef(meId)
                        });
                    } catch (e) {
                        if (!isDuplicateMemberError(e)) throw e;
                    }
                    try {
                        await graphJson('POST', '/groups/' + encodeURIComponent(groupId) + '/members/$ref', token, {
                            '@odata.id': userRef(meId)
                        });
                    } catch (e) {
                        if (!isDuplicateMemberError(e)) throw e;
                    }
                }
            } catch {
                // optional
            }
        }
    }

    async function syncEmailsToGroup(token, groupId, emails, label, onLog) {
        const log =
            onLog ||
            function () {
                /* noop */
            };
        let ok = 0;
        let skip = 0;
        let fail = 0;
        for (let i = 0; i < emails.length; i++) {
            const em = emails[i];
            try {
                const u = await resolveUserByEmail(token, em);
                if (!u || !u.id) {
                    log(label + ': Kein Benutzer für ' + em, 'warn');
                    fail++;
                    continue;
                }
                try {
                    await graphAddMember(token, groupId, u.id);
                    ok++;
                    log(label + ': ' + em + ' → Mitglied', 'ok');
                } catch (e) {
                    if (isDuplicateMemberError(e)) {
                        skip++;
                        log(label + ': ' + em + ' (war schon Mitglied)', 'warn');
                    } else {
                        fail++;
                        log(label + ': ' + em + ' — ' + (e.message || e), 'err');
                    }
                }
            } catch (e) {
                fail++;
                log(label + ': ' + em + ' — ' + (e.message || e), 'err');
            }
            if ((i + 1) % 8 === 0) await sleep(120);
        }
        return { ok: ok, skip: skip, fail: fail };
    }

    window.ms365GraphUnifiedGroups = {
        GRAPH_SCOPES,
        PERSON_SELECT,
        getGraphToken,
        sleep,
        graphRequest,
        graphJson,
        odataEscape,
        sanitizeMailNickname,
        sanitizeUnifiedGroupMailNickname,
        isUnifiedGroup,
        searchUnifiedGroups,
        fetchGroup,
        createUnifiedGroup,
        provisionTeamForGroup,
        graphAddMember,
        userRef,
        isDuplicateMemberError,
        resolveUserByEmail,
        ensureOwners,
        syncEmailsToGroup,
        patchGroupDisplayName
    };
})();
