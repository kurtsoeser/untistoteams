(function () {
    'use strict';

    const STORAGE_KEY = 'ms365-schueler-lehrer-gruppen-v2';

    function gug() {
        const G = window.ms365GraphUnifiedGroups;
        if (!G) throw new Error('graph-unified-groups.js muss vor diesem Skript geladen werden.');
        return G;
    }

    async function getGraphToken() {
        return gug().getGraphToken();
    }

    async function searchUnifiedGroups(token, queryRaw) {
        return gug().searchUnifiedGroups(token, queryRaw);
    }

    async function fetchGroup(token, id) {
        return gug().fetchGroup(token, id);
    }

    async function createUnifiedGroup(token, displayName, mailNickname, description) {
        return gug().createUnifiedGroup(token, displayName, mailNickname, description);
    }

    async function provisionTeamForGroup(token, gid) {
        return gug().provisionTeamForGroup(token, gid);
    }

    /** @type {'schueler' | 'lehrer'} */
    let activeKind = 'schueler';

    /** @type {{ schuelerGroupId: string|null, lehrerGroupId: string|null }} */
    let matched = { schuelerGroupId: null, lehrerGroupId: null };

    /** @type {{ students: string[], teachers: string[], direktion: string[] }} */
    let listCache = { students: [], teachers: [], direktion: [] };

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

    function escapeHtml(s) {
        return String(s)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    async function ensureOwners(token, groupId) {
        return gug().ensureOwners(token, groupId, listCache.direktion || []);
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

    async function syncEmailsToGroup(token, groupId, emails, label) {
        return gug().syncEmailsToGroup(token, groupId, emails, label, appendSyncLog);
    }

    function loadTenantSettings() {
        if (typeof window.ms365TenantSettingsLoad !== 'function') return null;
        return window.ms365TenantSettingsLoad();
    }

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

    function collectEmails(arr) {
        const out = [];
        const seen = new Set();
        (Array.isArray(arr) ? arr : []).forEach(function (row) {
            const em = normEmail(row && row.email);
            if (!em || em.indexOf('@') === -1) return;
            if (seen.has(em)) return;
            seen.add(em);
            out.push(em);
        });
        return out;
    }

    function readLists() {
        const settings = loadTenantSettings();
        listCache.students = collectEmails(settings && settings.students);
        listCache.teachers = collectEmails(settings && settings.teachers);
        listCache.direktion = collectDirektionOwnerEmails(settings);
    }

    function updateLeftListUi() {
        const sCount = document.getElementById('slgSchuelerCount');
        const tCount = document.getElementById('slgLehrerCount');
        if (sCount) sCount.textContent = String(listCache.students.length);
        if (tCount) tCount.textContent = String(listCache.teachers.length);

        const sLine = document.getElementById('slgSchuelerLine');
        const tLine = document.getElementById('slgLehrerLine');
        const sInfo = matched.schuelerGroupId ? 'Gematcht: ' + matched.schuelerGroupId : 'Noch kein Match';
        const tInfo = matched.lehrerGroupId ? 'Gematcht: ' + matched.lehrerGroupId : 'Noch kein Match';
        if (sLine) sLine.textContent = sInfo;
        if (tLine) tLine.textContent = tInfo;
    }

    function renderOwnerPreview() {
        const el = document.getElementById('slgOwnerPreview');
        if (!el) return;
        el.replaceChildren();
        const list = listCache.direktion || [];
        if (!list.length) {
            const p = document.createElement('p');
            p.style.margin = '0';
            p.style.color = '#6c757d';
            p.textContent = 'Keine Direktion‑Owner in den Schul‑Einstellungen gefunden.';
            el.appendChild(p);
            return;
        }
        list.forEach(function (em) {
            const d = document.createElement('div');
            d.textContent = em;
            d.style.padding = '4px 0';
            d.style.borderBottom = '1px solid #eef1f4';
            el.appendChild(d);
        });
    }

    function renderMemberPreview() {
        const el = document.getElementById('slgMemberPreview');
        if (!el) return;
        el.replaceChildren();
        const list = activeKind === 'schueler' ? listCache.students : listCache.teachers;
        const first = list.slice(0, 30);
        if (!first.length) {
            const p = document.createElement('p');
            p.style.margin = '0';
            p.style.color = '#6c757d';
            p.textContent = 'Keine E‑Mails in dieser Liste.';
            el.appendChild(p);
            return;
        }
        first.forEach(function (em) {
            const d = document.createElement('div');
            d.textContent = em;
            d.style.padding = '4px 0';
            d.style.borderBottom = '1px solid #eef1f4';
            el.appendChild(d);
        });
        if (list.length > first.length) {
            const more = document.createElement('div');
            more.className = 'muted';
            more.style.paddingTop = '8px';
            more.textContent = '… und ' + String(list.length - first.length) + ' weitere.';
            el.appendChild(more);
        }
    }

    function setTab(tab) {
        document.querySelectorAll('#slgDetailTabs .detail-tab-btn[data-slg-tab]').forEach(function (b) {
            const on = b.getAttribute('data-slg-tab') === tab;
            b.setAttribute('aria-selected', on ? 'true' : 'false');
        });
        document.querySelectorAll('[data-slg-tab-content]').forEach(function (p) {
            p.classList.toggle('active', p.getAttribute('data-slg-tab-content') === tab);
        });
        if (tab === 'owners') renderOwnerPreview();
        if (tab === 'members') renderMemberPreview();
    }

    function setActiveKind(kind) {
        activeKind = kind === 'lehrer' ? 'lehrer' : 'schueler';
        const title = document.getElementById('slgDetailTitle');
        const sub = document.getElementById('slgDetailSubtitle');
        if (title) title.textContent = activeKind === 'schueler' ? 'Schüler:innen' : 'Lehrer:innen';
        if (sub) sub.textContent = 'Gruppe matchen oder anlegen';

        document.querySelectorAll('button[data-slg-kind]').forEach(function (btn) {
            const on = btn.getAttribute('data-slg-kind') === activeKind;
            btn.setAttribute('aria-current', on ? 'true' : 'false');
        });

        const dn = document.getElementById('slgNewDisplayName');
        const nn = document.getElementById('slgNewMailNick');
        const desc = document.getElementById('slgNewDescription');
        if (activeKind === 'schueler') {
            if (dn && !dn.value) dn.value = 'Schüler:innen';
            if (nn && !nn.value) nn.value = 'schueler';
            if (desc && !desc.value) desc.value = 'Alle Schüler:innen (MS365-Schulverwaltung / Schul‑Liste)';
        } else {
            if (dn && !dn.value) dn.value = 'Lehrer:innen';
            if (nn && !nn.value) nn.value = 'lehrer';
            if (desc && !desc.value) desc.value = 'Alle Lehrer:innen (MS365-Schulverwaltung / Schul‑Liste)';
        }
        renderMatchSummary();
        renderMemberPreview();
        setTab('general');
    }

    function getActiveMatchedId() {
        return activeKind === 'schueler' ? matched.schuelerGroupId : matched.lehrerGroupId;
    }

    function setActiveMatchedId(id) {
        if (activeKind === 'schueler') matched.schuelerGroupId = id;
        else matched.lehrerGroupId = id;
    }

    function renderMatchSummary(group) {
        const summary = document.getElementById('slgMatchSummary');
        const kv = document.getElementById('slgMatchKv');
        const btnOpen = document.getElementById('slgBtnOpenEntra');
        const btnUn = document.getElementById('slgBtnUnmatch');
        const btnRefresh = document.getElementById('slgBtnRefreshGroup');
        const btnSync = document.getElementById('slgBtnSync');

        const gid = getActiveMatchedId();
        const has = !!gid;
        if (btnOpen) btnOpen.disabled = !has;
        if (btnUn) btnUn.disabled = !has;
        if (btnRefresh) btnRefresh.disabled = !has;
        if (btnSync) btnSync.disabled = !has;

        if (!has) {
            if (summary) summary.textContent = 'Noch keine Gruppe gematcht.';
            if (kv) kv.style.display = 'none';
            updateLeftListUi();
            return;
        }
        if (!group) {
            if (summary) summary.innerHTML = 'Gematchte Gruppen‑ID: <code>' + escapeHtml(gid) + '</code>';
            if (kv) kv.style.display = 'none';
            updateLeftListUi();
            return;
        }
        if (summary) summary.innerHTML = '<strong>OK:</strong> ' + escapeHtml(group.displayName || '(ohne Namen)');
        if (kv) kv.style.display = '';
        const kvName = document.getElementById('slgKvName');
        const kvMail = document.getElementById('slgKvMail');
        const kvNick = document.getElementById('slgKvNick');
        const kvId = document.getElementById('slgKvId');
        if (kvName) kvName.textContent = group.displayName || '–';
        if (kvMail) kvMail.textContent = group.mail || '–';
        if (kvNick) kvNick.textContent = group.mailNickname || '–';
        if (kvId) kvId.textContent = group.id || '–';
        updateLeftListUi();
    }

    function renderGroupSearchResults(list) {
        const host = document.getElementById('slgGroupSearchResults');
        if (!host) return;
        host.replaceChildren();
        if (!list || !list.length) {
            host.style.display = 'block';
            const p = document.createElement('div');
            p.className = 'muted';
            p.textContent = 'Keine passenden Microsoft 365‑Gruppen (Unified) gefunden.';
            host.appendChild(p);
            return;
        }
        host.style.display = 'block';

        const box = document.createElement('div');
        box.style.border = '1px solid #ced4da';
        box.style.borderRadius = '12px';
        box.style.background = '#fff';
        box.style.overflow = 'hidden';

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
                '<div style="font-weight:700;line-height:1.25;">' +
                escapeHtml(dn) +
                '</div>' +
                '<div class="muted" style="margin-top:2px;">Mail‑Nickname: <code>' +
                escapeHtml(nick) +
                '</code> · SMTP: ' +
                escapeHtml(mail) +
                '</div>' +
                '<div class="muted" style="margin-top:2px;">Gruppen‑ID: <code>' +
                escapeHtml(g && g.id ? g.id : '') +
                '</code></div>';

            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'btn btn-success';
            btn.textContent = 'Matchen';
            btn.addEventListener('click', function () {
                if (!g || !g.id) return;
                setActiveMatchedId(String(g.id));
                saveState();
                renderMatchSummary(g);
                toast('Gruppe gematcht.');
            });

            row.appendChild(left);
            row.appendChild(btn);
            box.appendChild(row);
        });
        host.appendChild(box);
    }

    async function runSearchGroups() {
        const inp = document.getElementById('slgGroupSearch');
        const q = inp && inp.value ? inp.value.trim() : '';
        if (!q) {
            toast('Bitte einen Suchbegriff eingeben.');
            return;
        }
        try {
            const token = await getGraphToken();
            const list = await searchUnifiedGroups(token, q);
            renderGroupSearchResults(list);
            if (!list.length) toast('Keine passenden Gruppen gefunden.');
        } catch (e) {
            toast('Fehler: ' + (e.message || e));
        }
    }

    async function runCreateAndMatch() {
        const dn = document.getElementById('slgNewDisplayName');
        const nn = document.getElementById('slgNewMailNick');
        const dd = document.getElementById('slgNewDescription');
        const ct = document.getElementById('slgNewCreateTeam');
        const displayName = dn ? dn.value : '';
        const mailNick = nn ? nn.value : '';
        const desc = dd ? dd.value : '';
        if (!normStr(displayName) || !normStr(mailNick)) {
            toast('Bitte Anzeigename und Alias/Mail‑Nickname ausfüllen.');
            return;
        }
        try {
            let token = await getGraphToken();
            const g = await createUnifiedGroup(token, displayName, mailNick, desc);
            await ensureOwners(token, g.id);
            if (ct && ct.checked) {
                toast('Gruppe angelegt – Team wird bereitgestellt …');
                await provisionTeamForGroup(token, g.id);
            }
            setActiveMatchedId(String(g.id));
            saveState();
            renderMatchSummary(g);
            toast('Gruppe angelegt und gematcht.');
        } catch (e) {
            toast('Fehler: ' + (e.message || e));
        }
    }

    async function runRefreshMatchedGroup() {
        const gid = getActiveMatchedId();
        if (!gid) return;
        try {
            const token = await getGraphToken();
            const g = await fetchGroup(token, gid);
            renderMatchSummary(g);
            toast('Gruppe neu geladen.');
        } catch (e) {
            toast('Fehler: ' + (e.message || e));
        }
    }

    function runUnmatch() {
        if (!getActiveMatchedId()) return;
        setActiveMatchedId(null);
        saveState();
        renderMatchSummary();
        toast('Match gelöst.');
    }

    function openEntraForMatched() {
        const gid = getActiveMatchedId();
        if (!gid) return;
        const url =
            'https://entra.microsoft.com/#view/Microsoft_AAD_IAM/GroupDetailsMenuBlade/~/Members/groupId/' +
            encodeURIComponent(gid);
        window.open(url, '_blank', 'noopener');
    }

    async function runSyncMembers() {
        const gid = getActiveMatchedId();
        if (!gid) {
            toast('Zuerst eine Gruppe matchen oder anlegen.');
            return;
        }
        const emails = activeKind === 'schueler' ? listCache.students : listCache.teachers;
        if (!emails.length) {
            toast('Keine E‑Mails in dieser Liste.');
            return;
        }
        clearSyncLog();
        appendSyncLog('Start: ' + (activeKind === 'schueler' ? 'Schüler:innen' : 'Lehrer:innen') + ' (' + emails.length + ' Adressen) …', '');
        try {
            const token = await getGraphToken();
            const label = activeKind === 'schueler' ? 'Schüler' : 'Lehrer';
            const r = await syncEmailsToGroup(token, gid, emails, label);
            appendSyncLog('Fertig: neu ' + r.ok + ', übersprungen ' + r.skip + ', Fehler ' + r.fail + '.', 'ok');
            await ensureOwners(token, gid);
            toast('Synchronisation abgeschlossen.');
        } catch (e) {
            appendSyncLog('Abbruch: ' + (e.message || e), 'err');
            toast('Fehler: ' + (e.message || e));
        }
    }

    function buildStateObject() {
        return {
            kind: STORAGE_KEY,
            savedAt: new Date().toISOString(),
            activeKind: activeKind,
            matched: {
                schuelerGroupId: matched.schuelerGroupId,
                lehrerGroupId: matched.lehrerGroupId
            },
            slgNewDisplayName: document.getElementById('slgNewDisplayName') ? document.getElementById('slgNewDisplayName').value : '',
            slgNewMailNick: document.getElementById('slgNewMailNick') ? document.getElementById('slgNewMailNick').value : '',
            slgNewDescription: document.getElementById('slgNewDescription') ? document.getElementById('slgNewDescription').value : '',
            slgNewCreateTeam: document.getElementById('slgNewCreateTeam') ? !!document.getElementById('slgNewCreateTeam').checked : false
        };
    }

    function applyStateObject(o) {
        if (!o || typeof o !== 'object') return;
        if (o.matched && typeof o.matched === 'object') {
            matched.schuelerGroupId = o.matched.schuelerGroupId ? String(o.matched.schuelerGroupId) : null;
            matched.lehrerGroupId = o.matched.lehrerGroupId ? String(o.matched.lehrerGroupId) : null;
        }
        const dn = document.getElementById('slgNewDisplayName');
        const nn = document.getElementById('slgNewMailNick');
        const dd = document.getElementById('slgNewDescription');
        const ct = document.getElementById('slgNewCreateTeam');
        if (dn && o.slgNewDisplayName !== undefined) dn.value = String(o.slgNewDisplayName);
        if (nn && o.slgNewMailNick !== undefined) nn.value = String(o.slgNewMailNick);
        if (dd && o.slgNewDescription !== undefined) dd.value = String(o.slgNewDescription);
        if (ct && o.slgNewCreateTeam !== undefined) ct.checked = !!o.slgNewCreateTeam;
        setActiveKind(o.activeKind === 'lehrer' ? 'lehrer' : 'schueler');
    }

    function saveState() {
        try {
            const obj = buildStateObject();
            localStorage.setItem(STORAGE_KEY, JSON.stringify(obj));
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.patchSetup === 'function') {
                window.ms365AppDataV2.patchSetup({
                    matched: obj.matched,
                    slgDraft: {
                        activeKind: obj.activeKind,
                        slgNewDisplayName: obj.slgNewDisplayName,
                        slgNewMailNick: obj.slgNewMailNick,
                        slgNewDescription: obj.slgNewDescription,
                        slgNewCreateTeam: obj.slgNewCreateTeam
                    }
                });
            }
        } catch {
            // ignore
        }
    }

    function loadState() {
        let rawLocal = null;
        try {
            rawLocal = localStorage.getItem(STORAGE_KEY);
        } catch {
            rawLocal = null;
        }
        try {
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getSetup === 'function') {
                const su = window.ms365AppDataV2.getSetup();
                const hasIds =
                    su && su.matched && !!(su.matched.schuelerGroupId || su.matched.lehrerGroupId);
                if (hasIds || !rawLocal) {
                    const d = su.slgDraft || {};
                    applyStateObject({
                        matched: su.matched,
                        activeKind: d.activeKind === 'lehrer' ? 'lehrer' : 'schueler',
                        slgNewDisplayName: d.slgNewDisplayName,
                        slgNewMailNick: d.slgNewMailNick,
                        slgNewDescription: d.slgNewDescription,
                        slgNewCreateTeam: d.slgNewCreateTeam
                    });
                    return;
                }
            }
        } catch {
            // ignore
        }
        try {
            if (!rawLocal) return;
            applyStateObject(JSON.parse(rawLocal));
        } catch {
            // ignore
        }
    }

    function clearStorage() {
        try {
            localStorage.removeItem(STORAGE_KEY);
            matched = { schuelerGroupId: null, lehrerGroupId: null };
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.patchSetup === 'function') {
                window.ms365AppDataV2.patchSetup({
                    matched: { schuelerGroupId: null, lehrerGroupId: null }
                });
            }
            saveState();
            renderMatchSummary();
            updateLeftListUi();
            toast('Zurückgesetzt.');
        } catch (e) {
            toast('Löschen fehlgeschlagen: ' + (e.message || e));
        }
    }

    async function onLogin() {
        const btn = document.getElementById('slgBtnLogin');
        if (btn) btn.disabled = true;
        try {
            await getGraphToken();
            toast('Angemeldet.');
        } catch (e) {
            toast('Anmeldung: ' + (e.message || e));
        } finally {
            if (btn) btn.disabled = false;
        }
    }

    function wire() {
        const listHost = document.getElementById('slgListItems');
        if (listHost) {
            listHost.addEventListener('click', function (ev) {
                const t = ev.target;
                if (!t || !t.closest) return;
                const item = t.closest('button[data-slg-kind]');
                if (!item) return;
                const kind = item.getAttribute('data-slg-kind');
                setActiveKind(kind === 'lehrer' ? 'lehrer' : 'schueler');
                saveState();
            });
        }

        document.querySelectorAll('#slgDetailTabs .detail-tab-btn[data-slg-tab]').forEach(function (b) {
            b.addEventListener('click', function () {
                const tab = b.getAttribute('data-slg-tab') || 'general';
                setTab(tab);
            });
        });

        document.getElementById('slgBtnLogin') &&
            document.getElementById('slgBtnLogin').addEventListener('click', function () {
                onLogin();
            });
        document.getElementById('slgBtnReloadLists') &&
            document.getElementById('slgBtnReloadLists').addEventListener('click', function () {
                readLists();
                updateLeftListUi();
                renderOwnerPreview();
                renderMemberPreview();
                toast('Listen neu eingelesen.');
            });

        document.getElementById('slgBtnSearch') &&
            document.getElementById('slgBtnSearch').addEventListener('click', function () {
                runSearchGroups();
            });
        document.getElementById('slgBtnCreate') &&
            document.getElementById('slgBtnCreate').addEventListener('click', function () {
                runCreateAndMatch();
            });
        document.getElementById('slgBtnRefreshGroup') &&
            document.getElementById('slgBtnRefreshGroup').addEventListener('click', function () {
                runRefreshMatchedGroup();
            });
        document.getElementById('slgBtnUnmatch') &&
            document.getElementById('slgBtnUnmatch').addEventListener('click', function () {
                runUnmatch();
            });
        document.getElementById('slgBtnOpenEntra') &&
            document.getElementById('slgBtnOpenEntra').addEventListener('click', function () {
                openEntraForMatched();
            });
        document.getElementById('slgBtnSync') &&
            document.getElementById('slgBtnSync').addEventListener('click', function () {
                runSyncMembers();
            });

        document.getElementById('slgBtnSaveState') &&
            document.getElementById('slgBtnSaveState').addEventListener('click', function () {
                saveState();
                toast('Gespeichert.');
            });
        document.getElementById('slgBtnLoadState') &&
            document.getElementById('slgBtnLoadState').addEventListener('click', function () {
                loadState();
                toast('Geladen.');
            });
        document.getElementById('slgBtnClearStorage') &&
            document.getElementById('slgBtnClearStorage').addEventListener('click', function () {
                clearStorage();
            });
    }

    function init() {
        readLists();
        loadState();
        updateLeftListUi();
        renderMatchSummary();
        renderOwnerPreview();
        renderMemberPreview();
        wire();
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();

