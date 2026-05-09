(function () {
    'use strict';

    const STORAGE_KEY = 'ms365-gruppenerstellung-policy-v1';

    const GRAPH_SCOPES = [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/Group.ReadWrite.All',
        'https://graph.microsoft.com/Directory.ReadWrite.All'
    ];

    let msalMod = null;
    let pca = null;
    let gpCurrentStep = 1;

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
                'Keine clientId: ms365-config.js fehlt/leer oder blockiert. Seite mit Strg+F5 neu laden; im Netzwerk-Tab prüfen, ob ms365-config.js mit 200 lädt.'
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
                typeof data === 'object' && data && data.error ? JSON.stringify(data.error) : text || String(res.status);
            throw new Error(method + ' ' + path + ': ' + msg);
        }
        return data || {};
    }

    function appendLog(msg, kind) {
        const el = document.getElementById('gpOnlineLog');
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
        const el = document.getElementById('gpOnlineLog');
        if (el) el.replaceChildren();
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
        if (!n) n = 'gruppenersteller';
        return n.toLowerCase();
    }

    async function getUnifiedTemplate(token) {
        const data = await graphJson('GET', '/groupSettingTemplates', token);
        const list = data.value || [];
        const t = list.find(function (x) {
            return x && String(x.displayName) === 'Group.Unified';
        });
        if (!t || !t.id) {
            throw new Error('Vorlage Group.Unified nicht gefunden.');
        }
        return t;
    }

    async function getOrCreateUnifiedSetting(token, template) {
        const all = await graphJson('GET', '/groupSettings', token);
        const rows = all.value || [];
        let row = rows.find(function (r) {
            return r && String(r.templateId) === String(template.id);
        });
        if (!row) {
            row = await graphJson('POST', '/groupSettings', token, { templateId: template.id });
        }
        if (!row || !row.id) {
            throw new Error('Verzeichniseinstellung Group.Unified konnte nicht angelegt werden.');
        }
        return row;
    }

    async function applyGroupCreationRestriction(token, groupObjectId) {
        const template = await getUnifiedTemplate(token);
        const setting = await getOrCreateUnifiedSetting(token, template);
        await graphJson('PATCH', '/groupSettings/' + encodeURIComponent(setting.id), token, {
            values: [
                { name: 'EnableGroupCreation', value: 'false' },
                { name: 'GroupCreationAllowedGroupId', value: String(groupObjectId).trim() }
            ]
        });
        return setting.id;
    }

    async function fetchUnifiedSettingStatus(token) {
        const template = await getUnifiedTemplate(token);
        const all = await graphJson('GET', '/groupSettings', token);
        const rows = all.value || [];
        const row = rows.find(function (r) {
            return r && String(r.templateId) === String(template.id);
        });
        if (!row) {
            return { template: template, setting: null, values: {} };
        }
        const map = {};
        (row.values || []).forEach(function (v) {
            if (v && v.name) map[v.name] = v.value;
        });
        return { template: template, setting: row, values: map };
    }

    async function removeUnifiedSetting(token) {
        const st = await fetchUnifiedSettingStatus(token);
        if (!st.setting || !st.setting.id) {
            throw new Error('Keine Group.Unified-Einstellung vorhanden (nichts zu entfernen).');
        }
        await graphJson('DELETE', '/groupSettings/' + encodeURIComponent(st.setting.id), token, undefined);
    }

    async function getGroupById(token, id) {
        return graphJson('GET', '/groups/' + encodeURIComponent(id), token, undefined);
    }

    async function findSecurityGroupByDisplayName(token, displayName) {
        const esc = String(displayName).replace(/'/g, "''");
        const filter =
            "displayName eq '" + esc + "' and securityEnabled eq true and mailEnabled eq false";
        const path = '/groups?$filter=' + encodeURIComponent(filter) + '&$top=25';
        const data = await graphJson('GET', path, token, undefined);
        const list = data.value || [];
        return list;
    }

    async function createSecurityGroup(token, displayName) {
        let nick = sanitizeMailNickname(displayName);
        const body = {
            displayName: String(displayName).trim(),
            description:
                'Nur Mitglieder dieser Gruppe dürfen im Mandanten einheitliche Microsoft 365-Gruppen und Teams anlegen (Entra-Richtlinie Group.Unified: EnableGroupCreation=false, diese Gruppe als GroupCreationAllowedGroupId). Es muss eine reine Sicherheitsgruppe sein (keine mail-aktivierte Gruppe).',
            mailEnabled: false,
            mailNickname: nick,
            securityEnabled: true,
            groupTypes: []
        };
        try {
            return await graphJson('POST', '/groups', token, body);
        } catch (e) {
            if (String(e.message || e).indexOf('mailNickname') !== -1 || /409|conflict/i.test(String(e.message))) {
                body.mailNickname = nick + '-' + Math.random().toString(36).slice(2, 8);
                return await graphJson('POST', '/groups', token, body);
            }
            throw e;
        }
    }

    function entraGroupBladeUrl(tab, groupId) {
        return (
            'https://entra.microsoft.com/#view/Microsoft_AAD_IAM/GroupDetailsMenuBlade/~/' +
            tab +
            '/groupId/' +
            encodeURIComponent(groupId)
        );
    }

    function setResolvedGroup(id, displayName) {
        const hid = document.getElementById('gpResolvedObjectId');
        const out = document.getElementById('gpResolvedSummary');
        const linkOv = document.getElementById('gpLinkOverview');
        const linkOwn = document.getElementById('gpLinkOwners');
        const linkMem = document.getElementById('gpLinkMembers');
        if (hid) hid.value = id || '';
        if (out) {
            out.style.display = id ? 'block' : 'none';
            out.textContent = id
                ? 'Ausgewählte Gruppe: ' + (displayName ? displayName + ' · ' : '') + 'Object-ID ' + id
                : '';
        }
        [linkOv, linkOwn, linkMem].forEach(function (link) {
            if (!link) return;
            if (id) {
                if (link === linkOv) link.href = entraGroupBladeUrl('Overview', id);
                else if (link === linkOwn) link.href = entraGroupBladeUrl('Owners', id);
                else link.href = entraGroupBladeUrl('Members', id);
                link.style.display = 'inline';
            } else {
                link.style.display = 'none';
            }
        });
    }

    function describeGraphGroupKind(g) {
        if (!g) return '–';
        const types = g.groupTypes || [];
        if (types.indexOf('Unified') !== -1) return 'Microsoft 365-Gruppe';
        if (g.securityEnabled && g.mailEnabled) return 'E-Mail-Sicherheitsgruppe';
        if (g.securityEnabled && !g.mailEnabled) return 'Sicherheitsgruppe';
        if (g.mailEnabled) return 'Mail-aktivierte Gruppe';
        return 'Gruppe';
    }

    async function loadGroupDetailsIntoStep2() {
        const hid = document.getElementById('gpResolvedObjectId');
        const gid = hid && hid.value ? hid.value.trim() : '';
        const ph = document.getElementById('gpDetailsPlaceholder');
        const dl = document.getElementById('gpDetailsDl');
        if (!ph || !dl) return;
        if (!guidLooksValid(gid)) {
            ph.style.display = 'block';
            ph.textContent = 'Keine gültige Gruppen-Object-ID – bitte zuerst Schritt 1 abschließen.';
            dl.style.display = 'none';
            return;
        }
        ph.style.display = 'block';
        ph.textContent = 'Lese Gruppe aus Microsoft Graph …';
        dl.style.display = 'none';
        try {
            const token = await getGraphToken();
            const g = await getGroupById(token, gid);
            if (!g.securityEnabled) {
                ph.textContent =
                    'Diese Gruppe ist keine Sicherheitsgruppe (securityEnabled=false). Bitte in Schritt 1 eine reine Sicherheitsgruppe wählen.';
                dl.style.display = 'none';
                return;
            }
            if (g.mailEnabled) {
                ph.textContent =
                    'Diese Gruppe ist mail-aktiviert (E-Mail-Sicherheitsgruppe). Für Group.Unified wird eine reine Sicherheitsgruppe benötigt (mailEnabled=false). Bitte in Schritt 1 eine passende Gruppe anlegen oder auswählen.';
                dl.style.display = 'none';
                return;
            }
            const setTxt = function (id, text) {
                const n = document.getElementById(id);
                if (n) n.textContent = text != null && String(text) !== '' ? String(text) : '–';
            };
            setTxt('gpDetDisplayName', g.displayName);
            setTxt('gpDetType', describeGraphGroupKind(g));
            setTxt('gpDetDescription', g.description);
            setTxt('gpDetMailNick', g.mailNickname);
            setTxt('gpDetId', g.id);
            ph.style.display = 'none';
            dl.style.display = 'block';
        } catch (e) {
            ph.style.display = 'block';
            ph.textContent = 'Gruppe konnte nicht gelesen werden: ' + (e.message || e);
            dl.style.display = 'none';
        }
    }

    function gpStepNum(el) {
        const raw = el.getAttribute('data-gp-step');
        const n = parseFloat(String(raw || '').trim());
        return Number.isFinite(n) ? n : NaN;
    }

    function goToGpStep(step) {
        gpCurrentStep = step;
        document.querySelectorAll('.gp-step-content').forEach(function (el) {
            el.classList.toggle('active', gpStepNum(el) === step);
        });
        document.querySelectorAll('.gp-steps .step').forEach(function (el) {
            const s = gpStepNum(el);
            el.classList.toggle('active', s === step);
            el.classList.toggle('completed', s < step);
        });
        if (typeof window.ms365ApplyStepProgress === 'function') {
            window.ms365ApplyStepProgress(document.querySelector('.gp-steps'), step, [1, 2]);
        }
        if (step === 2) {
            loadGroupDetailsIntoStep2().catch(function () {});
        }
    }

    async function onCreateGroupClick() {
        const nameInp = document.getElementById('gpInputDisplayName');
        const name = nameInp && nameInp.value ? nameInp.value.trim() : '';
        if (!name) {
            toast('Bitte einen Anzeigenamen für die Sicherheitsgruppe eintragen.');
            return;
        }
        const btn = document.getElementById('gpBtnCreateGroup');
        if (btn) btn.disabled = true;
        clearLog();
        appendLog('Sicherheitsgruppe wird angelegt …', '');
        try {
            const token = await getGraphToken();
            const g = await createSecurityGroup(token, name);
            setResolvedGroup(g.id, g.displayName || name);
            appendLog('Sicherheitsgruppe erstellt: ' + (g.displayName || name) + ' (' + g.id + ').', 'ok');
            toast('Sicherheitsgruppe angelegt.');
        } catch (e) {
            appendLog('Fehler: ' + (e.message || e), 'err');
            toast('Fehler: ' + (e.message || e));
        } finally {
            if (btn) btn.disabled = false;
        }
    }

    async function onFindGroupClick() {
        const oidInp = document.getElementById('gpInputObjectId');
        const nameInp = document.getElementById('gpInputDisplayName');
        const rawId = oidInp && oidInp.value ? oidInp.value.trim() : '';
        const name = nameInp && nameInp.value ? nameInp.value.trim() : '';

        const btn = document.getElementById('gpBtnFindGroup');
        if (btn) btn.disabled = true;
        clearLog();
        try {
            const token = await getGraphToken();
            if (guidLooksValid(rawId)) {
                const g = await getGroupById(token, rawId);
                if (!g.securityEnabled) {
                    throw new Error(
                        'Die Gruppe ist keine Sicherheitsgruppe (securityEnabled=false). Bitte eine reine Sicherheitsgruppe verwenden.'
                    );
                }
                if (g.mailEnabled) {
                    throw new Error(
                        'Die Gruppe ist mail-aktiviert (z. B. E-Mail-Sicherheitsgruppe). Für diese Richtlinie ist eine reine Sicherheitsgruppe erforderlich (mailEnabled=false).'
                    );
                }
                setResolvedGroup(g.id, g.displayName);
                appendLog('Gruppe per Object-ID geladen: ' + g.displayName + ' (' + g.id + ').', 'ok');
                toast('Gruppe gefunden.');
            } else if (name) {
                const list = await findSecurityGroupByDisplayName(token, name);
                if (!list.length) {
                    throw new Error(
                        'Keine passende reine Sicherheitsgruppe mit diesem Anzeigenamen gefunden (nur securityEnabled und ohne Mail-Aktivierung).'
                    );
                }
                if (list.length > 1) {
                    appendLog('Mehrere Treffer – bitte Object-ID aus dem Admin Center kopieren.', 'warn');
                }
                const g = list[0];
                setResolvedGroup(g.id, g.displayName);
                appendLog('Gruppe gefunden: ' + g.displayName + ' (' + g.id + ').', 'ok');
                toast('Gruppe gefunden.');
            } else {
                throw new Error('Bitte gültige Object-ID (GUID) eintragen oder Anzeigenamen zur Suche.');
            }
        } catch (e) {
            appendLog('Fehler: ' + (e.message || e), 'err');
            toast('Fehler: ' + (e.message || e));
        } finally {
            if (btn) btn.disabled = false;
        }
    }

    async function onApplyClick() {
        const hid = document.getElementById('gpResolvedObjectId');
        const gid = hid && hid.value ? hid.value.trim() : '';
        if (!guidLooksValid(gid)) {
            toast('Bitte zuerst in Schritt 1 eine gültige reine Sicherheitsgruppe wählen oder anlegen.');
            return;
        }
        const btn = document.getElementById('gpBtnApply');
        const btnLogin = document.getElementById('gpBtnLogin');
        if (btn) btn.disabled = true;
        if (btnLogin) btnLogin.disabled = true;
        clearLog();
        appendLog('Setze Group.Unified (EnableGroupCreation=false, GroupCreationAllowedGroupId) …', '');
        try {
            let token = await getGraphToken();
            await applyGroupCreationRestriction(token, gid);
            appendLog(
                'Einstellung gespeichert. Nur Mitglieder der Sicherheitsgruppe dürfen nun M365-Gruppen/Teams erstellen.',
                'ok'
            );
            toast('Richtlinie gespeichert.');
            await refreshStatusIntoUi(token);
        } catch (e) {
            appendLog('Fehler: ' + (e.message || e), 'err');
            toast('Fehler: ' + (e.message || e));
        } finally {
            if (btn) btn.disabled = false;
            if (btnLogin) btnLogin.disabled = false;
        }
    }

    async function refreshStatusIntoUi(token) {
        const box = document.getElementById('gpStatusPre');
        if (!token) {
            token = await getGraphToken();
        }
        const st = await fetchUnifiedSettingStatus(token);
        let text = '';
        if (!st.setting) {
            text =
                'Keine benutzerdefinierte Group.Unified-Einstellung (Mandanten-Standard).\n' +
                'Hinweis: Ohne diese Einstellung gelten die Microsoft-Standardregeln für die Gruppenerstellung.';
        } else {
            text =
                'Einstellungs-ID: ' +
                st.setting.id +
                '\n' +
                'EnableGroupCreation: ' +
                (st.values.EnableGroupCreation !== undefined ? st.values.EnableGroupCreation : '(nicht gesetzt)') +
                '\n' +
                'GroupCreationAllowedGroupId: ' +
                (st.values.GroupCreationAllowedGroupId || '(leer)') +
                '\n';
        }
        if (box) box.textContent = text;
    }

    async function onRefreshStatusClick() {
        const btn = document.getElementById('gpBtnRefreshStatus');
        if (btn) btn.disabled = true;
        clearLog();
        try {
            const token = await getGraphToken();
            await refreshStatusIntoUi(token);
            appendLog('Status gelesen.', 'ok');
        } catch (e) {
            appendLog('Fehler: ' + (e.message || e), 'err');
            toast('Fehler: ' + (e.message || e));
        } finally {
            if (btn) btn.disabled = false;
        }
    }

    async function onRemoveClick() {
        const ok =
            typeof window.ms365AppDialogConfirm === 'function'
                ? await window.ms365AppDialogConfirm(
                      'Die Verzeichniseinstellung Group.Unified wirklich entfernen? (Entspricht dem Entfernen per PowerShell; Mandant fällt auf Standard zurück.)',
                      { title: 'Einstellung entfernen', okText: 'Entfernen', danger: true }
                  )
                : window.confirm(
                      'Die Verzeichniseinstellung Group.Unified wirklich entfernen? (Entspricht dem Entfernen per PowerShell; Mandant fällt auf Standard zurück.)'
                  );
        if (!ok) {
            return;
        }
        const btn = document.getElementById('gpBtnRemove');
        if (btn) btn.disabled = true;
        clearLog();
        try {
            const token = await getGraphToken();
            await removeUnifiedSetting(token);
            appendLog('Einstellung entfernt.', 'ok');
            toast('Einstellung entfernt.');
            await refreshStatusIntoUi(token);
        } catch (e) {
            appendLog('Fehler: ' + (e.message || e), 'err');
            toast('Fehler: ' + (e.message || e));
        } finally {
            if (btn) btn.disabled = false;
        }
    }

    async function onLoginClick() {
        const btn = document.getElementById('gpBtnLogin');
        if (btn) btn.disabled = true;
        try {
            await getGraphToken();
            toast('Angemeldet – Sie können den Status lesen oder die Einschränkung setzen.');
            appendLog('Anmeldung OK (benötigt Directory.ReadWrite.All mit Administratorzustimmung).', 'ok');
        } catch (e) {
            toast('Anmeldung: ' + (e.message || e));
        } finally {
            if (btn) btn.disabled = false;
        }
    }

    function saveState() {
        try {
            const hid = document.getElementById('gpResolvedObjectId');
            const dn = document.getElementById('gpInputDisplayName');
            const oid = document.getElementById('gpInputObjectId');
            const state = {
                gpCurrentStep: gpCurrentStep,
                groupDisplayName: dn ? dn.value : '',
                groupObjectIdRaw: oid ? oid.value : '',
                resolvedObjectId: hid ? hid.value : ''
            };
            localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
            toast('Gruppenerstellung: Zwischenstand gespeichert.');
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
            const state = JSON.parse(raw);
            const dn = document.getElementById('gpInputDisplayName');
            const oid = document.getElementById('gpInputObjectId');
            if (dn && state.groupDisplayName !== undefined) dn.value = state.groupDisplayName;
            if (oid && state.groupObjectIdRaw !== undefined) oid.value = state.groupObjectIdRaw;
            if (typeof state.gpCurrentStep === 'number' && state.gpCurrentStep >= 1) {
                let st = state.gpCurrentStep;
                if (st > 2) st = 2;
                goToGpStep(st);
            }
            if (state.resolvedObjectId && guidLooksValid(state.resolvedObjectId)) {
                setResolvedGroup(state.resolvedObjectId, '');
            }
            toast('Gruppenerstellung: Stand geladen.');
        } catch (e) {
            toast('Laden fehlgeschlagen: ' + (e.message || e));
        }
    }

    function clearState() {
        try {
            localStorage.removeItem(STORAGE_KEY);
            toast('Gruppenerstellung: lokaler Speicher geleert.');
        } catch (e) {
            toast('Fehler: ' + (e.message || e));
        }
    }

    function init() {
        goToGpStep(1);
        const dn = document.getElementById('gpInputDisplayName');
        if (dn && !dn.value) dn.value = 'GruppenErsteller';

        document.getElementById('gpBtnCreateGroup') &&
            document.getElementById('gpBtnCreateGroup').addEventListener('click', onCreateGroupClick);
        document.getElementById('gpBtnFindGroup') &&
            document.getElementById('gpBtnFindGroup').addEventListener('click', onFindGroupClick);
        document.getElementById('gpBtnNext1') &&
            document.getElementById('gpBtnNext1').addEventListener('click', function () {
                const hid = document.getElementById('gpResolvedObjectId');
                const oidInp = document.getElementById('gpInputObjectId');
                const raw = oidInp && oidInp.value ? oidInp.value.trim() : '';
                if (guidLooksValid(raw)) {
                    setResolvedGroup(raw, '');
                }
                const gid = hid && hid.value ? hid.value.trim() : '';
                if (!guidLooksValid(gid)) {
                    toast(
                        'Bitte zuerst eine Sicherheitsgruppe anlegen, suchen oder eine gültige Object-ID (GUID) eintragen.'
                    );
                    return;
                }
                goToGpStep(2);
            });
        document.getElementById('gpBtnBack2') &&
            document.getElementById('gpBtnBack2').addEventListener('click', function () {
                goToGpStep(1);
            });

        document.getElementById('gpBtnLogin') && document.getElementById('gpBtnLogin').addEventListener('click', onLoginClick);
        document.getElementById('gpBtnApply') && document.getElementById('gpBtnApply').addEventListener('click', onApplyClick);
        document.getElementById('gpBtnRefreshStatus') &&
            document.getElementById('gpBtnRefreshStatus').addEventListener('click', onRefreshStatusClick);
        document.getElementById('gpBtnRemove') && document.getElementById('gpBtnRemove').addEventListener('click', onRemoveClick);
    }

    window.ms365SaveGruppenerstellung = saveState;
    window.ms365LoadGruppenerstellung = loadState;
    window.ms365ClearGruppenerstellung = clearState;
    window.ms365GpGraphLogin = onLoginClick;
    window.ms365GpGraphApply = onApplyClick;

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();

