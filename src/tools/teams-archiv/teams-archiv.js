(function () {
    'use strict';

    const GRAPH_SCOPES = [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/TeamSettings.ReadWrite.All',
        'https://graph.microsoft.com/Group.ReadWrite.All'
    ];

    let msalMod = null;
    let pca = null;

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

    async function graphRequest(method, path, token, body) {
        const url = path.indexOf('http') === 0 ? path : 'https://graph.microsoft.com/v1.0' + path;
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

    async function graphJson(method, path, token, body) {
        const res = await graphRequest(method, path, token, body);
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
            throw new Error(method + ' ' + path + ': ' + msg);
        }
        return data || {};
    }

    function appendLog(msg, kind) {
        const el = document.getElementById('taLog');
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
        const el = document.getElementById('taLog');
        if (el) el.replaceChildren();
    }

    function parseTeamsOperationPathFromLocation(locationHeader) {
        if (!locationHeader) return null;
        let loc = String(locationHeader).trim();
        if (loc.indexOf('http') === 0) {
            try {
                const u = new URL(loc);
                loc = u.pathname.replace(/^\/v1\.0/i, '');
            } catch {
                return null;
            }
        }
        const m = loc.match(/\/teams\/([^/]+)\/operations\/([^/?\s]+)/i);
        if (m) return '/teams/' + m[1] + '/operations/' + m[2];
        const m2 = loc.match(/teams\('([^']+)'\)\/operations\('([^']+)'\)/i);
        if (m2) return '/teams/' + m2[1] + '/operations/' + m2[2];
        return null;
    }

    async function pollTeamsAsyncOperation(token, operationPath) {
        const maxAttempts = 90;
        for (let i = 0; i < maxAttempts; i++) {
            await sleep(2000);
            const data = await graphJson('GET', operationPath, token, undefined);
            const st = String(data.status || data.Status || '').toLowerCase();
            if (st === 'succeeded') {
                appendLog('Asynchrone Teams-Operation abgeschlossen.', 'ok');
                return;
            }
            if (st === 'failed') {
                const errMsg =
                    (data.error && (data.error.message || JSON.stringify(data.error))) || JSON.stringify(data);
                throw new Error('Teams-Operation fehlgeschlagen: ' + errMsg);
            }
            if (i > 0 && i % 10 === 0) {
                appendLog('Warte auf Teams-Operation … (' + i * 2 + ' s)', 'warn');
            }
        }
        throw new Error('Timeout: Teams-Operation nicht abgeschlossen.');
    }

    function normGuid(v) {
        const s = String(v || '').trim();
        if (/^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(s)) {
            return s;
        }
        return '';
    }

    function odataEscape(s) {
        return String(s).replace(/'/g, "''");
    }

    function groupHasTeamProvisioning(g) {
        const opts = g && g.resourceProvisioningOptions;
        return Array.isArray(opts) && opts.indexOf('Team') !== -1;
    }

    /**
     * Prüft Teams-Ressource: zuerst resourceProvisioningOptions, sonst einmal GET …/team.
     */
    async function filterGroupsWithTeam(token, groups) {
        const out = [];
        for (const g of groups) {
            if (!g || !g.id) continue;
            if (groupHasTeamProvisioning(g)) {
                out.push(g);
                continue;
            }
            try {
                const team = await graphJson('GET', '/groups/' + encodeURIComponent(g.id) + '/team', token, undefined);
                if (team && (team.id || team.displayName !== undefined)) {
                    out.push(g);
                }
            } catch {
                // keine Team-Ressource
            }
        }
        return out;
    }

    const GROUP_SELECT =
        'id,displayName,mail,mailNickname,resourceProvisioningOptions';

    async function searchTeam(token, query) {
        const q = String(query || '').trim();
        if (!q) throw new Error('Suchbegriff fehlt.');

        if (normGuid(q)) {
            try {
                const g = await graphJson(
                    'GET',
                    '/groups/' + encodeURIComponent(q) + '?$select=' + GROUP_SELECT,
                    token,
                    undefined
                );
                const withTeam = await filterGroupsWithTeam(token, [g]);
                if (!withTeam.length) {
                    throw new Error('Diese GUID ist keine Microsoft-365-Gruppe mit Team.');
                }
                const x = withTeam[0];
                return { id: x.id, displayName: x.displayName || '', mail: x.mail || '' };
            } catch (e) {
                throw new Error('GUID nicht gefunden oder kein Team: ' + (e.message || e));
            }
        }

        let collection = [];
        if (q.indexOf('@') !== -1) {
            const filter = "mail eq '" + odataEscape(q) + "'";
            const data = await graphJson(
                'GET',
                '/groups?$filter=' + encodeURIComponent(filter) + '&$select=' + GROUP_SELECT,
                token,
                undefined
            );
            collection = data.value || [];
        } else {
            const filter = "startswith(displayName,'" + odataEscape(q) + "')";
            const data = await graphJson(
                'GET',
                '/groups?$filter=' +
                    encodeURIComponent(filter) +
                    '&$select=' +
                    GROUP_SELECT +
                    '&$top=15',
                token,
                undefined
            );
            collection = data.value || [];
            if (!collection.length) {
                const ex = "displayName eq '" + odataEscape(q) + "'";
                const data2 = await graphJson(
                    'GET',
                    '/groups?$filter=' +
                        encodeURIComponent(ex) +
                        '&$select=' +
                        GROUP_SELECT +
                        '&$top=15',
                    token,
                    undefined
                );
                collection = data2.value || [];
            }
        }

        const withTeams = await filterGroupsWithTeam(token, collection);
        if (!withTeams.length) {
            throw new Error('Kein Team zu diesen Suchkriterien gefunden (oder keine Teams-Ressource).');
        }
        if (withTeams.length > 1) {
            appendLog(
                'Hinweis: Mehrere Treffer – es wird das erste Team mit Teams-Ressource verwendet. Bitte ggf. die GUID direkt eintragen.',
                'warn'
            );
        }
        const g = withTeams[0];
        return { id: g.id, displayName: g.displayName || '', mail: g.mail || '' };
    }

    async function runArchiveOrUnarchive(archive) {
        const idInp = document.getElementById('taTeamId');
        const spo = document.getElementById('taSpoReadOnly');
        const teamId = normGuid(idInp && idInp.value);
        if (!teamId) {
            toast('Bitte zuerst eine gültige Team-/Gruppen-ID eintragen oder „Team suchen“ verwenden.');
            return;
        }

        const btnA = document.getElementById('taBtnArchive');
        const btnU = document.getElementById('taBtnUnarchive');
        if (btnA) btnA.disabled = true;
        if (btnU) btnU.disabled = true;

        try {
            const token = await getGraphToken();
            const path = '/teams/' + encodeURIComponent(teamId) + (archive ? '/archive' : '/unarchive');
            let body = undefined;
            if (archive && spo && spo.checked) {
                body = { shouldSetSpoSiteReadOnlyForMembers: true };
            }

            appendLog((archive ? 'Archivierung' : 'Aufheben der Archivierung') + ' starten …');
            const res = await graphRequest('POST', path, token, body);

            if (res.status !== 202 && res.status !== 200) {
                const t = await res.text();
                throw new Error('HTTP ' + res.status + ' ' + t);
            }

            const loc = res.headers.get('Location') || res.headers.get('Content-Location');
            const opPath = parseTeamsOperationPathFromLocation(loc);
            if (opPath) {
                appendLog('Asynchrone Verarbeitung (202) – Status wird abgefragt …', 'warn');
                await pollTeamsAsyncOperation(token, opPath);
            } else {
                appendLog('Keine Operation-URL in der Antwort – bitte Status im Admin Center prüfen.', 'warn');
            }

            appendLog(archive ? 'Team archiviert (Anfrage erfolgreich).' : 'Archivierung aufgehoben (Anfrage erfolgreich).', 'ok');
            toast(archive ? 'Archivierung ausgeführt.' : 'Archivierung aufgehoben.');
        } catch (e) {
            const msg = e && e.message ? e.message : String(e);
            appendLog('Fehler: ' + msg, 'err');
            toast('Fehler: ' + msg);
        } finally {
            if (btnA) btnA.disabled = false;
            if (btnU) btnU.disabled = false;
        }
    }

    async function onSearch() {
        const searchInp = document.getElementById('taSearch');
        const idInp = document.getElementById('taTeamId');
        const summary = document.getElementById('taResolvedSummary');
        const q = searchInp && searchInp.value ? searchInp.value.trim() : '';
        if (!q) {
            toast('Bitte E-Mail oder Anzeigename (oder GUID) im Suchfeld eintragen.');
            return;
        }
        clearLog();
        try {
            const token = await getGraphToken();
            appendLog('Suche …');
            const found = await searchTeam(token, q);
            if (idInp) idInp.value = found.id;
            if (summary) {
                summary.style.display = '';
                summary.textContent =
                    'Gefunden: ' +
                    (found.displayName || '(ohne Name)') +
                    (found.mail ? ' · ' + found.mail : '') +
                    ' · ID: ' +
                    found.id;
            }
            appendLog('Team gefunden – ID wurde übernommen.', 'ok');
            toast('Team gefunden.');
        } catch (e) {
            appendLog('Suche: ' + (e && e.message ? e.message : e), 'err');
            toast(String(e && e.message ? e.message : e));
        }
    }

    async function onLogin() {
        const btn = document.getElementById('taBtnLogin');
        if (btn) btn.disabled = true;
        try {
            await getGraphToken();
            toast('Angemeldet – Sie können archivieren oder die Archivierung aufheben.');
        } catch (e) {
            toast('Anmeldung: ' + (e.message || e));
        } finally {
            if (btn) btn.disabled = false;
        }
    }

    function bind() {
        const btnL = document.getElementById('taBtnLogin');
        const btnS = document.getElementById('taBtnSearch');
        const btnA = document.getElementById('taBtnArchive');
        const btnU = document.getElementById('taBtnUnarchive');
        if (btnL) btnL.addEventListener('click', () => onLogin());
        if (btnS) btnS.addEventListener('click', () => onSearch());
        if (btnA) btnA.addEventListener('click', () => runArchiveOrUnarchive(true));
        if (btnU) btnU.addEventListener('click', () => runArchiveOrUnarchive(false));
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', bind);
    } else {
        bind();
    }
})();
