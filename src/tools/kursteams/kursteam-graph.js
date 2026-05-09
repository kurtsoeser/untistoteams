(function () {
    'use strict';

    const MSAL_LOADER_IMPORT = (function () {
        const needle = 'kursteam-graph.js';
        const rel = '../../shared/msal-loader.js';
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
        'https://graph.microsoft.com/Group.ReadWrite.All',
        'https://graph.microsoft.com/User.Read.All',
        /** POST /teams mit teamsTemplates('educationClass') – wie New-Team -Template EDU_Class */
        'https://graph.microsoft.com/Team.Create',
        /**
         * Wird für POST /education/classes mitgefordert; laut Microsoft Learn ist diese API für delegierte Auth oft
         * nicht verfügbar – dann schlägt die Online-Anlage fehl und „Kursteam-Anlage.cmd“ (PowerShell) ist nötig.
         */
        'https://graph.microsoft.com/EduRoster.ReadWrite'
    ];

    let msalMod = null;
    let pca = null;

    function toast(msg) {
        if (typeof window.ms365ToastOrAlert === 'function') {
            window.ms365ToastOrAlert(msg);
        } else if (typeof window.ms365ShowToast === 'function') {
            window.ms365ShowToast(msg);
        } else {
            window.alert(msg);
        }
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
        if (!cfg) {
            cfg = {};
        }
        let id = String(cfg.clientId || '').trim();
        if (!id) {
            const meta = document.querySelector('meta[name="ms365-graph-client-id"]');
            const fromMeta = meta && meta.getAttribute('content') ? meta.getAttribute('content').trim() : '';
            if (fromMeta) {
                id = fromMeta;
            }
        }
        if (!id) {
            throw new Error(
                'Keine clientId: ms365-config.js fehlt/leer oder blockiert. Seite mit Strg+F5 neu laden; im Netzwerk-Tab prüfen, ob ms365-config.js mit 200 lädt. Alternativ meta ms365-graph-client-id in ms365-schooltool.html setzen (Entra-Anwendungs-ID).'
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

    /**
     * @param {{ forceRefresh?: boolean }} [options] forceRefresh: neues Zugriffstoken vom Server holen (z. B. nach neuen
     *   Entra-Berechtigungen), damit der **scp**-Claim alle angeforderten Scopes enthält (sonst „Required scp claim…“).
     */
    async function getGraphToken(options) {
        options = options || {};
        const forceRefresh = options.forceRefresh === true;
        const instance = await getPca();
        let accounts = instance.getAllAccounts();
        if (!accounts.length) {
            await instance.loginPopup({ scopes: GRAPH_SCOPES, prompt: 'select_account' });
            accounts = instance.getAllAccounts();
        }
        if (!accounts.length) {
            throw new Error('Anmeldung abgebrochen.');
        }
        const account = accounts[0];
        const silentReq = {
            scopes: GRAPH_SCOPES,
            account: account,
            forceRefresh: forceRefresh
        };
        try {
            return (await instance.acquireTokenSilent(silentReq)).accessToken;
        } catch (e) {
            if (isInteractionRequired(e)) {
                return (
                    await instance.acquireTokenPopup({
                        scopes: GRAPH_SCOPES,
                        account: account,
                        prompt: 'consent'
                    })
                ).accessToken;
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

    function isGraphDuplicateRefError(err) {
        const msg = String(err && err.message ? err.message : err);
        return /already exist/i.test(msg) || /already exists/i.test(msg);
    }

    function appendLog(msg, kind) {
        const el = document.getElementById('kursteamOnlineLog');
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
        const el = document.getElementById('kursteamOnlineLog');
        if (el) el.replaceChildren();
    }

    /**
     * Microsoft: PowerShell New-Team -Template "EDU_Class" entspricht POST /teams mit teamsTemplates('educationClass')
     * und group@odata.bind auf die Gruppe einer Education-Klasse (nach POST /education/classes).
     */
    function parseTeamsOperationPath(locationHeader) {
        if (!locationHeader) return null;
        const loc = String(locationHeader).trim();
        const m = loc.match(/teams\('([^']+)'\)\/operations\('([^']+)'\)/i);
        if (m) return '/teams/' + m[1] + '/operations/' + m[2];
        const m2 = loc.match(/\/teams\/([^/]+)\/operations\/([^/?\s]+)/i);
        if (m2) return '/teams/' + m2[1] + '/operations/' + m2[2];
        return null;
    }

    async function pollTeamsAsyncOperation(token, operationPath, appendLog) {
        const maxAttempts = 120;
        for (let i = 0; i < maxAttempts; i++) {
            await sleep(2000);
            const data = await graphJson('GET', operationPath, token, undefined);
            const st = String(data.status || data.Status || '').toLowerCase();
            if (st === 'succeeded') {
                appendLog('  Teams: Bereitstellung abgeschlossen (Template educationClass).', 'ok');
                return;
            }
            if (st === 'failed') {
                const errMsg =
                    (data.error && (data.error.message || JSON.stringify(data.error))) ||
                    JSON.stringify(data);
                throw new Error('Team-Bereitstellung fehlgeschlagen: ' + errMsg);
            }
            if (i > 0 && i % 8 === 0) {
                appendLog('  Teams: Warte auf Bereitstellung … (' + i * 2 + ' s)', 'warn');
            }
        }
        throw new Error('Timeout: Team-Bereitstellung (async) nicht abgeschlossen.');
    }

    function sanitizeEducationClassCode(t) {
        const raw = (t && (t.gruppenmail || t.teamName)) || 'Klasse';
        const s = String(raw).replace(/[^a-zA-Z0-9]/g, '');
        return s.substring(0, 50) || 'Klasse';
    }

    /**
     * Vollständiges Kursteam (Aufgaben, Klassennotizbuch, …): laut Microsoft
     * - zuerst Education-Klasse: POST /education/classes (legt die passende M365-Gruppe mit Education-Metadaten an),
     * - dann Team mit teamsTemplates('educationClass'): POST /teams (entspricht PowerShell New-Team -Template EDU_Class).
     * PUT /groups/{id}/team mit specialization educationClass ist dafür kein Ersatz (kein echtes Kursteam).
     */
    async function createEducationClassGroup(token, t, appendLog) {
        const body = {
            '@odata.type': '#microsoft.graph.educationClass',
            displayName: t.teamName,
            mailNickname: t.gruppenmail,
            description: 'Kursteam (WebUntis / MS365-Schulverwaltung)',
            classCode: sanitizeEducationClassCode(t),
            externalSource: 'manual'
        };
        appendLog('  Education: POST /education/classes (legt Klassen-Gruppe für Kursteam an) …', 'warn');
        const edu = await graphJson('POST', '/education/classes', token, body);
        return edu.id;
    }

    /**
     * @param {Error|string} originalErr
     * @returns {Error}
     */
    function buildPostEducationClassesError(originalErr) {
        const raw = originalErr && originalErr.message ? originalErr.message : String(originalErr);
        const isMissingScp =
            /Required scp claim values are not provided/i.test(raw) ||
            (/AccessDenied/i.test(raw) && /scp/i.test(raw));

        if (isMissingScp) {
            return new Error(
                'POST /education/classes: Das Zugriffstoken enthält die nötigen Graph-Berechtigungen nicht ' +
                    '(Microsoft: „Required scp claim values are not provided“ / AccessDenied). ' +
                    'Prüfen Sie in Entra: **delegierte** Berechtigung **EduRoster.ReadWrite** (nicht nur Anwendungsberechtigung), ' +
                    '„Administratorzustimmung“ erteilen. Anschließend **Kursteams jetzt anlegen** erneut – das Tool fordert ' +
                    'ein frisches Token an. Falls es weiter fehlschlägt: Browserdaten für diese Seite löschen oder Inkognito. ' +
                    'Hinweis: Selbst mit korrektem **scp** kann POST /education/classes laut Microsoft Learn für Browser ' +
                    'weiterhin unzulässig sein – dann **Kursteam-Anlage.cmd** (New-Team -Template EDU_Class). ' +
                    'Technisch: ' +
                    raw
            );
        }

        return new Error(
            'Vollständiges Kursteam (Microsoft-Template EDU_Class / educationClass) per Browser oft nicht möglich: ' +
                'POST /education/classes scheitert. Laut Microsoft Learn ist diese API häufig nur mit ' +
                'Anwendungsberechtigung EduRoster.ReadWrite.All (App-Only) vorgesehen, nicht mit delegierter Anmeldung. ' +
                'Alternative: **Kursteam-Anlage.cmd** (PowerShell: New-Team -Template EDU_Class) oder Backend mit App-Only-Token. ' +
                'Technisch: ' +
                raw
        );
    }

    async function addGroupOwnerAndMember(token, gid, ownerId, appendLog) {
        await sleep(2000);
        try {
            await graphJson('POST', '/groups/' + gid + '/owners/$ref', token, {
                '@odata.id': 'https://graph.microsoft.com/v1.0/directoryObjects/' + ownerId
            });
        } catch (e) {
            if (isGraphDuplicateRefError(e)) {
                appendLog(
                    '  Besitzer: bereits gesetzt (häufig, wenn gleicher Admin wie angemeldeter Benutzer).',
                    'warn'
                );
            } else {
                throw e;
            }
        }
        try {
            await graphJson('POST', '/groups/' + gid + '/members/$ref', token, {
                '@odata.id': 'https://graph.microsoft.com/v1.0/directoryObjects/' + ownerId
            });
        } catch (e) {
            if (isGraphDuplicateRefError(e)) {
                appendLog('  Mitglied: bereits gesetzt.', 'warn');
            } else {
                appendLog('  Hinweis (Besitzer als Mitglied): ' + e.message, 'warn');
            }
        }
    }

    /**
     * POST /teams mit teamsTemplates('educationClass') an eine bestehende Education-Klassen-Gruppe.
     * Kein PUT-Fallback: PUT erzeugt kein vollständiges Kursteam.
     */
    async function provisionKursteamPostTeamsEducationTemplate(token, gid, appendLog) {
        const postBody = {
            'template@odata.bind':
                'https://graph.microsoft.com/v1.0/teamsTemplates(\'educationClass\')',
            'group@odata.bind': 'https://graph.microsoft.com/v1.0/groups(\'' + gid + '\')'
        };

        let lastPostErr = null;
        for (let attempt = 0; attempt < 3; attempt++) {
            try {
                const res = await graphRequest('POST', '/teams', token, postBody);
                const text = await res.text();
                if (res.status === 202 || res.status === 200) {
                    const loc = res.headers.get('Location') || res.headers.get('Content-Location');
                    const opPath = parseTeamsOperationPath(loc);
                    if (opPath) {
                        appendLog('  Teams: POST /teams mit Template educationClass (wie EDU_Class) …', 'warn');
                        await pollTeamsAsyncOperation(token, opPath, appendLog);
                    } else {
                        appendLog(
                            '  Teams: POST /teams angenommen (keine Operation-URL – ggf. im Teams-Admin prüfen).',
                            'warn'
                        );
                    }
                    return await getGraphToken();
                }
                if (res.status === 404 && attempt < 2) {
                    appendLog(
                        '  Teams: 404 nach Klassenanlage – Replikation, Warte 10 s …',
                        'warn'
                    );
                    await sleep(10000);
                    token = await getGraphToken();
                    continue;
                }
                lastPostErr = new Error('POST /teams: ' + res.status + ' ' + (text || ''));
                break;
            } catch (e) {
                lastPostErr = e;
                if (attempt < 2 && /404/.test(String(e.message))) {
                    appendLog('  Teams: Wiederholung nach Wartezeit (404) …', 'warn');
                    await sleep(10000);
                    token = await getGraphToken();
                    continue;
                }
                break;
            }
        }

        const detail = lastPostErr && lastPostErr.message ? lastPostErr.message : String(lastPostErr);
        throw new Error(
            'POST /teams (Template educationClass) ist fehlgeschlagen – kein Kursteam angelegt. ' +
                'PUT /team wird nicht verwendet (liefert kein vollständiges Kursteam). Details: ' +
                detail
        );
    }

    async function runKursteamOnline() {
        if (isKursteamGraphOnlineDisabled()) return;
        const snapshotFn = window.ms365GetKursteamSnapshotForGraph;
        if (typeof snapshotFn !== 'function') {
            appendLog('Interner Fehler: Kursteam-Daten nicht verfügbar.', 'err');
            return;
        }
        const pack = snapshotFn();
        if (!pack || !pack.teams || !pack.teams.length) {
            appendLog('Keine gültigen Teams – bitte in Schritt „Teams konfigurieren“ generieren und prüfen.', 'err');
            return;
        }
        const missing = pack.teams.filter(function (t) {
            return !t.besitzer;
        });
        if (missing.length) {
            appendLog('Bitte für alle Teams einen gültigen Besitzer (E-Mail / UPN) im Mandanten eintragen.', 'err');
            return;
        }

        const btnLogin = document.getElementById('kursteamOnlineLogin');
        const btnRun = document.getElementById('kursteamOnlineRun');
        if (btnRun) btnRun.disabled = true;
        if (btnLogin) btnLogin.disabled = true;

        clearLog();
        appendLog('Start – Microsoft Graph (Browser), Kursteams …');
        appendLog(
            'Ziel: echtes Kursteam wie New-Team -Template EDU_Class → POST /education/classes, dann POST /teams mit ' +
                'teamsTemplates(\'educationClass\') (Microsoft Learn: Create team, Create educationClass).',
            'warn'
        );
        appendLog(
            'Wichtig: POST /education/classes ist für delegierte Browser-Anmeldung oft gesperrt – dann schlägt die Anlage fehl; ' +
                'nutzen Sie Schritt 7 „Kursteam-Anlage.cmd“ (PowerShell).',
            'warn'
        );

        let token;
        try {
            token = await getGraphToken({ forceRefresh: true });
        } catch (e) {
            appendLog('Anmeldung/Token: ' + (e.message || e), 'err');
            if (btnRun) btnRun.disabled = false;
            if (btnLogin) btnLogin.disabled = false;
            return;
        }

        const total = pack.teams.length;
        let i = 0;
        for (const t of pack.teams) {
            i++;
            try {
                appendLog('[' + i + '/' + total + '] ' + t.teamName + ' …');

                const owner = await graphJson(
                    'GET',
                    '/users/' + encodeURIComponent(t.besitzer),
                    token,
                    undefined
                );
                const ownerId = owner.id;

                let gid;
                try {
                    gid = await createEducationClassGroup(token, t, appendLog);
                } catch (e) {
                    throw buildPostEducationClassesError(e);
                }

                await addGroupOwnerAndMember(token, gid, ownerId, appendLog);
                token = await provisionKursteamPostTeamsEducationTemplate(token, gid, appendLog);

                appendLog('OK [' + i + '/' + total + '] ' + t.teamName + ' → ' + t.gruppenmail, 'ok');
            } catch (e) {
                appendLog('Fehler [' + i + '/' + total + '] ' + t.teamName + ': ' + (e.message || e), 'err');
            }

            await sleep(2000);
            try {
                token = await getGraphToken();
            } catch (e) {
                appendLog('Token erneuern: ' + (e.message || e), 'err');
                break;
            }
        }

        appendLog('Fertig.', 'ok');
        if (btnRun) btnRun.disabled = false;
        if (btnLogin) btnLogin.disabled = false;
    }

    function isKursteamGraphOnlineDisabled() {
        return !!document.getElementById('kursteamGraphDisabledOverlay');
    }

    async function loginOnly() {
        if (isKursteamGraphOnlineDisabled()) return;
        const btnLogin = document.getElementById('kursteamOnlineLogin');
        if (btnLogin) btnLogin.disabled = true;
        try {
            await getGraphToken({ forceRefresh: true });
            toast('Microsoft angemeldet – Sie können jetzt Kursteams anlegen.');
        } catch (e) {
            toast('Anmeldung: ' + (e.message || e));
        } finally {
            if (btnLogin) btnLogin.disabled = false;
        }
    }

    window.ms365KursteamGraphLogin = loginOnly;
    window.ms365KursteamGraphRun = runKursteamOnline;
})();

