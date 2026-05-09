import './structure-rules.js';
import { allowedStructureChildTypes, structureTreeRowShowsAddChildControl } from './schulstruktur-sync-helpers.js';

(function () {
    'use strict';

    const structureRules = (typeof window !== 'undefined' && window.ms365StructureRules) ? window.ms365StructureRules : null;
    const canReparentStrict = structureRules && typeof structureRules.canReparent === 'function'
        ? structureRules.canReparent
        : () => false;
    const inferRootForType = structureRules && typeof structureRules.inferRootForType === 'function'
        ? structureRules.inferRootForType
        : () => '';

    const STORAGE_KEY = 'ms365-schulstruktur-sync-v1';
    const STORAGE_TENANT_CACHE_KEY = 'ms365-schulstruktur-tenant-cache-v1';
    const STORAGE_MATCH_KEY = 'ms365-schulstruktur-match-v1';
    const UI_MODE_KEY = 'ms365-schulstruktur-sync-ui-mode-v1';

    const GRAPH_SCOPES_TENANT_READ = [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/Group.Read.All'
    ];
    /** Gruppen + Benutzerliste (Schritt 4 Grundkonfiguration); User.Read.All für GET /users. */
    const GRAPH_SCOPES_TENANT_INVENTORY = [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/User.Read.All',
        'https://graph.microsoft.com/Group.Read.All'
    ];
    const GRAPH_SCOPES_TENANT_WRITE = [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/Group.ReadWrite.All'
    ];
    const GRAPH_SCOPES_TENANT_OWNER_MANAGE = [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/User.Read.All',
        'https://graph.microsoft.com/Group.ReadWrite.All'
    ];
    /** POST …/teams/{id}/archive|unarchive (Tenant-Details „Team-Archiv“) */
    const GRAPH_SCOPES_TENANT_TEAM_ARCHIVE = [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/Group.ReadWrite.All',
        'https://graph.microsoft.com/TeamSettings.ReadWrite.All'
    ];
    /** Gruppe per Graph + Benutzer (Person) per Graph POST /users (Administratorzustimmung User.ReadWrite.All). */
    const GRAPH_SCOPES_GRAPH_OBJECT_CREATE = [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/User.Read.All',
        'https://graph.microsoft.com/Group.ReadWrite.All',
        'https://graph.microsoft.com/User.ReadWrite.All'
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

    function dlgAlert(msg, opts) {
        if (typeof window.ms365AppDialogAlert === 'function') return window.ms365AppDialogAlert(msg, opts);
        window.alert(msg);
        return Promise.resolve();
    }

    function dlgConfirm(msg, opts) {
        if (typeof window.ms365AppDialogConfirm === 'function') return window.ms365AppDialogConfirm(msg, opts);
        return Promise.resolve(window.confirm(msg));
    }

    function dlgPrompt(msg, def, opts) {
        if (typeof window.ms365AppDialogPrompt === 'function') return window.ms365AppDialogPrompt(msg, def, opts);
        return Promise.resolve(window.prompt(msg, def));
    }

    function normKey(s) {
        return String(s || '')
            .trim()
            .toLowerCase()
            .replace(/\s+/g, ' ')
            .replace(/[^\p{L}\p{N}\-_. ]/gu, '');
    }

    /** Klasse aus ktKlasse oder Eltern-Knoten „Klasse“; Fach aus ktFach (Differenz/Anlage konsistent). */
    function resolveKursteamKlasseFachForRow(row, rows) {
        if (window.ms365StructureRules && typeof window.ms365StructureRules.resolveKursteamKlasseFach === 'function') {
            return window.ms365StructureRules.resolveKursteamKlasseFach(row, rows);
        }
        const byId = Object.create(null);
        (rows || []).forEach((r) => {
            if (r && r.id != null) byId[String(r.id)] = r;
        });
        let klasse = normStr(row && row.ktKlasse);
        let fach = normStr(row && row.ktFach);
        const pid = normStr(row && row.parentId);
        if (!klasse && pid && byId[pid] && normStr(byId[pid].typ) === 'Klasse') {
            klasse = normStr(byId[pid].bezeichnung);
        }
        return { klasse, fach, hasBoth: !!(klasse && fach) };
    }

    function defaultTenantTargetForTypeStr(typ) {
        const t = String(typ || '');
        if (t === 'Kursteam' || t === 'Klasse') return 'team';
        return 'group';
    }

    function defaultTenantVisibilityForTypeStr(typ) {
        const t = String(typ || '');
        if (t === 'Kursteam') return 'HiddenMembership';
        return 'Private';
    }

    function computeTenantCreateSuggestionFromRow(row, schemaState) {
        const schema = schemaState && typeof schemaState === 'object' ? schemaState : defaultAnlegenSchemas();
        const typ = String(row?.typ || '');
        const displayName = String(row?.bezeichnung || '').trim();
        if (!displayName) return { displayName: '', mailNick: '' };

        if (typ === 'Jahrgang') {
            const y = String(row.jgYear || '').trim();
            const suf = String(row.jgSuffix || '').trim();
            const mailNick = y && suf ? buildJgMailNick(schema, y, suf) : buildMailNickFromLabel(displayName);
            return { displayName, mailNick };
        }
        if (typ === 'Arbeitsgemeinschaft') {
            const code = String(row.argeCode || '').trim();
            const mailNick = code ? buildArgeMailNick(schema, code) : buildMailNickFromLabel(displayName);
            return { displayName, mailNick };
        }
        if (typ === 'Kursteam') {
            const yearPrefix = String(schema.kursteamYearPrefix || '').trim();
            const st = typeof loadState === 'function' ? loadState() : { rows: [] };
            const kt = resolveKursteamKlasseFachForRow(row, st.rows);
            const klasse = kt.klasse;
            const fach = kt.fach;
            const gruppe = String(row.ktGruppe || '').trim();
            const mailNick =
                klasse && fach
                    ? buildKursteamMailNickFromTemplate(schema.kursteamMailNickPattern, { yearPrefix, klasse, fach, gruppe })
                    : buildMailNickFromLabel(displayName);
            return { displayName, mailNick };
        }
        return { displayName, mailNick: buildMailNickFromLabel(displayName) };
    }

    function suggestTenantGroupForUnitFromList(unit, list) {
        if (!unit) return '';
        const uKey = normKey(unit.bezeichnung || '');
        if (!uKey) return '';
        const rows = Array.isArray(list) ? list : [];
        let best = rows.find((g) => normKey(g.bezeichnung) === uKey);
        if (best) return String(best.id);
        best = rows.find((g) => g.alias && normKey(g.alias) === uKey);
        if (best) return String(best.id);
        best = rows.find(
            (g) => normKey(g.bezeichnung).includes(uKey) || (g.alias && normKey(g.alias).includes(uKey))
        );
        return best ? String(best.id) : '';
    }

    /** Abgleich-Vorschlag: Entra-Benutzer für SOLL-Typ „Person“ (Name/E-Mail/Rolle). */
    function suggestTenantUserForPersonFromList(unit, users) {
        if (!unit || String(unit.typ || '') !== 'Person') return '';
        const arr = Array.isArray(users) ? users : [];
        const keys = [];
        const pushK = (x) => {
            const k = normKey(x);
            if (k) keys.push(k);
        };
        pushK(unit.personName);
        pushK(unit.personEmail);
        pushK(unit.bezeichnung);
        if (!keys.length) return '';

        function scoreUser(u) {
            const dn = normKey(u.displayName);
            const upn = normKey(u.userPrincipalName);
            const mail = normKey(u.mail);
            let best = 0;
            for (let i = 0; i < keys.length; i++) {
                const k = keys[i];
                if (!k) continue;
                if (k === dn || k === upn || k === mail) return 100;
                if (dn && (dn.includes(k) || k.includes(dn))) best = Math.max(best, 50);
                if (upn && (upn.includes(k) || k.includes(upn))) best = Math.max(best, 45);
                if (mail && (mail.includes(k) || k.includes(mail))) best = Math.max(best, 45);
            }
            return best;
        }

        let bestId = '';
        let bestScore = 0;
        for (let j = 0; j < arr.length; j++) {
            const u = arr[j];
            const sc = scoreUser(u);
            if (sc > bestScore) {
                bestScore = sc;
                bestId = String(u.id || '');
            }
        }
        return bestId;
    }

    /** Dropdown-Wert: g:{id} = Gruppe/Team, u:{id} = Entra-Benutzer (Klartext für Speichern). */
    function suggestTenantMatchSelectValue(unit, groups, users) {
        if (!unit) return '';
        if (String(unit.typ || '') === 'Person') {
            const uid = suggestTenantUserForPersonFromList(unit, users);
            if (uid) return 'u:' + uid;
        }
        const gid = suggestTenantGroupForUnitFromList(unit, groups);
        return gid ? 'g:' + gid : '';
    }

    function formatEntraUserPickLabel(u) {
        if (!u || typeof u !== 'object') return '';
        const dn = u.displayName ? String(u.displayName).trim() : '';
        const upn = String(u.userPrincipalName || u.mail || '').trim();
        if (dn && upn && dn.toLowerCase() !== upn.toLowerCase()) return dn + ' · ' + upn + ' · Benutzer';
        return (dn || upn || String(u.id || '')) + ' · Benutzer';
    }

    function matchTenantFilterNeedle(raw) {
        return String(raw || '')
            .trim()
            .toLowerCase()
            .replace(/\s+/g, ' ');
    }

    function matchTenantHaystackForGroup(g) {
        return [
            g && g.bezeichnung,
            g && g.typ,
            g && g.alias,
            g && g.mail,
            g && g.description,
            g && g.id
        ]
            .map((x) => String(x || '').toLowerCase())
            .join(' ');
    }

    function matchTenantHaystackForUser(u) {
        return [
            u && u.displayName,
            u && u.userPrincipalName,
            u && u.mail,
            u && u.id,
            formatEntraUserPickLabel(u)
        ]
            .map((x) => String(x || '').toLowerCase())
            .join(' ');
    }

    /**
     * Baut die Tenant-Auswahlliste neu (optional gefiltert). full lists in window.__ms365MatchTenantPickSource.
     * @param {HTMLSelectElement} selTenant
     * @param {string} filterRaw
     * @param {string} selectedValue Wert nach Auswahl (z. B. g:… / u:…)
     */
    function rebuildMatchTenantSelectOptions(selTenant, filterRaw, selectedValue) {
        const src = window.__ms365MatchTenantPickSource;
        const cntEl = getEl('ssMatchTenantFilterCount');
        if (!selTenant || !src || typeof src !== 'object') {
            if (cntEl) cntEl.textContent = '';
            return;
        }
        const needle = matchTenantFilterNeedle(filterRaw);
        const list = Array.isArray(src.groups) ? src.groups : [];
        const users = Array.isArray(src.users) ? src.users : [];
        const prevSel = String(selectedValue != null ? selectedValue : selTenant.value || '');

        selTenant.replaceChildren();
        const opt0 = document.createElement('option');
        opt0.value = '';
        opt0.textContent = '(keine Zuordnung)';
        selTenant.appendChild(opt0);

        const gFiltered = needle
            ? list.filter((g) => matchTenantHaystackForGroup(g).indexOf(needle) !== -1)
            : list.slice();
        const uFiltered = needle
            ? users.filter((u) => matchTenantHaystackForUser(u).indexOf(needle) !== -1)
            : users.slice();

        function selectionInFiltered() {
            if (!prevSel) return true;
            if (prevSel.startsWith('g:')) {
                const id = prevSel.slice(2);
                return gFiltered.some((g) => String(g.id) === id);
            }
            if (prevSel.startsWith('u:')) {
                const id = prevSel.slice(2);
                return uFiltered.some((u) => String(u.id) === id);
            }
            return false;
        }

        if (prevSel && needle && !selectionInFiltered()) {
            let label = prevSel;
            if (prevSel.startsWith('g:')) {
                const id = prevSel.slice(2);
                const g = list.find((x) => String(x.id) === id);
                if (g) label = (g.bezeichnung || id) + ' · ' + (g.typ || '') + ' · aktuell verknüpft';
            } else if (prevSel.startsWith('u:')) {
                const id = prevSel.slice(2);
                const u = users.find((x) => String(x.id) === id);
                if (u) label = formatEntraUserPickLabel(u) + ' · aktuell verknüpft';
            }
            const ox = document.createElement('option');
            ox.value = prevSel;
            ox.textContent = label;
            selTenant.appendChild(ox);
        }

        if (gFiltered.length) {
            const og = document.createElement('optgroup');
            og.label = 'Gruppen / Teams';
            for (let i = 0; i < gFiltered.length; i++) {
                const g = gFiltered[i];
                const o = document.createElement('option');
                o.value = 'g:' + String(g.id || '');
                o.textContent = (g.bezeichnung || '(ohne Name)') + ' · ' + (g.typ || '') + (g.alias ? ' · ' + g.alias : '');
                og.appendChild(o);
            }
            selTenant.appendChild(og);
        }

        if (uFiltered.length) {
            const ou = document.createElement('optgroup');
            ou.label = 'Benutzer (Entra ID)';
            for (let j = 0; j < uFiltered.length; j++) {
                const u = uFiltered[j];
                const o = document.createElement('option');
                o.value = 'u:' + String(u.id || '');
                o.textContent = formatEntraUserPickLabel(u);
                ou.appendChild(o);
            }
            selTenant.appendChild(ou);
        }

        if (prevSel) {
            const ok = Array.from(selTenant.options).some((o) => String(o.value) === prevSel);
            selTenant.value = ok ? prevSel : '';
        } else {
            selTenant.value = '';
        }

        if (cntEl) {
            const total = list.length + users.length;
            const shown = gFiltered.length + uFiltered.length;
            if (!total) cntEl.textContent = 'Noch keine Tenant-Daten – unter „Verwalten“ Tenant einlesen.';
            else if (!needle) cntEl.textContent = String(total) + ' Einträge – Suchfeld nutzen, um die Liste einzugrenzen.';
            else cntEl.textContent = 'Zeige ' + shown + ' von ' + total + ' Einträgen (Filter aktiv).';
        }
    }

    let __matchTenantSearchWired = false;
    function wireMatchTenantSearchOnce() {
        if (__matchTenantSearchWired) return;
        const inp = getEl('ssMatchTenantSearch');
        if (!inp) return;
        __matchTenantSearchWired = true;
        inp.addEventListener('input', () => {
            const sel = getEl('ssMatchTenantGroup');
            if (!sel) return;
            const cur = String(sel.value || '');
            rebuildMatchTenantSelectOptions(sel, inp.value || '', cur);
        });
    }


    function patchStructureRowById(rowId, patch) {
        const st = loadState();
        const idx = st.rows.findIndex((r) => String(r.id) === String(rowId));
        if (idx === -1) return false;
        const rows = st.rows.slice();
        rows[idx] = Object.assign({}, rows[idx], patch);
        saveState({ rows, memberships: st.memberships, settings: st.settings });
        try {
            window.dispatchEvent(new CustomEvent('ms365-structure-changed', { detail: {} }));
        } catch {
            // ignore
        }
        return true;
    }

    function saveMatchLinkPublic(structureId, tenantGroupId, note, tenantUserId) {
        const cur = loadMatchState().links || {};
        const links = Object.assign({}, cur);
        const id = String(structureId || '');
        if (!id) return links;
        const gid = normStr(tenantGroupId);
        const uid = normStr(tenantUserId);
        if (!gid && !uid) delete links[id];
        else {
            links[id] = {
                tenantGroupId: gid,
                tenantUserId: uid,
                note: String(note || ''),
                updatedAt: new Date().toISOString()
            };
        }
        saveMatchState(links);
        try {
            window.__ms365MatchLinks = links;
        } catch {
            // ignore
        }
        try {
            window.dispatchEvent(new CustomEvent('ms365-match-links-changed', { detail: { links } }));
        } catch {
            // ignore
        }
        return links;
    }

    async function graphProvisionStructureGroupRow(row) {
        const st = loadState();
        const schemaState = Object.assign({}, defaultAnlegenSchemas(), st.settings || {});
        const sug = computeTenantCreateSuggestionFromRow(row, schemaState);
        if (!sug.displayName) throw new Error('Keine Bezeichnung.');
        if (!sug.mailNick) throw new Error('Mail‑Nickname leer (Schema prüfen).');
        const desc = normStr(row.beschreibung || '');
        const typ = String(row.typ || '');
        let created;
        if (typ === 'Jahrgang') {
            // Jahrgänge: mail-enabled Sicherheitsgruppe (verschachtelungsfreundlich, kein Team).
            created = await createMailEnabledSecurityGroup(sug.displayName, desc, sug.mailNick);
        } else {
            const vis = defaultTenantVisibilityForTypeStr(row.typ);
            created = await createUnifiedGroup(sug.displayName, desc, sug.mailNick, vis);
        }
        const gid = String(created.id || '').trim();
        if (!gid) throw new Error('Keine Gruppen-ID von Graph.');
        const target = typ === 'Jahrgang' ? 'security' : defaultTenantTargetForTypeStr(row.typ);
        if (target === 'team') await createTeamForGroup(gid);
        const mem = st.memberships[String(row.id)] || { owners: [], members: [] };
        const ownerIds = (mem.owners || []).map((p) => String(p.id || '')).filter(Boolean);
        const memberIds = (mem.members || []).map((p) => String(p.id || '')).filter(Boolean);
        for (let i = 0; i < ownerIds.length; i++) {
            try {
                await addOwnerWithMemberFallback(gid, ownerIds[i]);
            } catch {
                /* optional */
            }
        }
        for (let j = 0; j < memberIds.length; j++) {
            try {
                await addGroupMember(gid, memberIds[j]);
            } catch (e) {
                if (!isGraphDuplicateRefError(e)) throw e;
            }
        }
        patchStructureRowById(row.id, {
            tenantGroupId: gid,
            tenantMailNickname: sug.mailNick,
            tenantTarget: target,
            tenantVisibility: typ === 'Jahrgang' ? '' : defaultTenantVisibilityForTypeStr(row.typ),
            syncStatus: 'Ok',
            letzteFehlermeldung: ''
        });
        saveMatchLinkPublic(String(row.id), gid, 'Schritt 5');
        return { groupId: gid };
    }

    async function graphProvisionPersonRowPublic(row, opts) {
        const o = opts && typeof opts === 'object' ? opts : {};
        const skipConfirm = !!o.skipConfirm;
        const displayName = normStr(row.personName) || normStr(row.bezeichnung);
        const upn = normStr(row.personEmail).toLowerCase();
        if (!displayName) throw new Error('Name fehlt (Person).');
        if (!upn || upn.indexOf('@') === -1) throw new Error('UPN/E-Mail fehlt.');
        if (normStr(row.tenantUserId)) throw new Error('Benutzer bereits verknüpft.');
        const mailNick = mailNicknameFromUpn(upn);
        const pwd = generateGraphTempPassword();
        if (!skipConfirm) {
            if (!(await dlgConfirm('Benutzer in Entra anlegen?\n\n' + displayName + '\n' + upn, { title: 'Entra', okText: 'Anlegen' }))) {
                throw new Error('Abgebrochen.');
            }
        }
        const token = await getGraphToken(GRAPH_SCOPES_GRAPH_OBJECT_CREATE);
        const body = {
            accountEnabled: true,
            displayName,
            mailNickname: mailNick,
            userPrincipalName: upn,
            passwordProfile: {
                forceChangePasswordNextSignIn: true,
                password: pwd
            }
        };
        const created = await graphJson('POST', '/users', token, body, undefined);
        const uid = String(created.id || '').trim();
        if (!uid) throw new Error('Keine Benutzer-ID von Graph.');
        patchStructureRowById(row.id, {
            personName: displayName,
            personEmail: upn,
            tenantUserId: uid,
            syncStatus: 'Ok',
            letzteFehlermeldung: ''
        });
        return { userId: uid, tempPassword: pwd };
    }

    function loadState() {
        try {
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getContainer === 'function') {
                const c = window.ms365AppDataV2.getContainer();
                if (c && c.structure && typeof c.structure === 'object') {
                    const rows = Array.isArray(c.structure.rows) ? c.structure.rows : [];
                    const memberships =
                        c.structure.memberships && typeof c.structure.memberships === 'object' ? c.structure.memberships : {};
                    const settings =
                        c.structure.settings && typeof c.structure.settings === 'object' ? c.structure.settings : {};
                    return { rows, memberships, settings };
                }
            }
            const raw = localStorage.getItem(STORAGE_KEY);
            const obj = raw ? safeJsonParse(raw) : null;
            const rows = obj && Array.isArray(obj.rows) ? obj.rows : [];
            const memberships =
                obj && obj.memberships && typeof obj.memberships === 'object' ? obj.memberships : {};
            const settings =
                obj && obj.settings && typeof obj.settings === 'object' ? obj.settings : {};
            return { rows, memberships, settings };
        } catch {
            return { rows: [], memberships: {}, settings: {} };
        }
    }

    function saveState(state) {
        const rows = state && Array.isArray(state.rows) ? state.rows : [];
        const memberships = state && state.memberships && typeof state.memberships === 'object' ? state.memberships : {};
        const settings = state && state.settings && typeof state.settings === 'object' ? state.settings : {};
        try {
            localStorage.setItem(STORAGE_KEY, JSON.stringify({ rows, memberships, settings }));
        } catch {
            // ignore
        }
        try {
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getContainer === 'function' && typeof window.ms365AppDataV2.setContainer === 'function') {
                const c = window.ms365AppDataV2.getContainer();
                c.structure = { rows, memberships, settings };
                window.ms365AppDataV2.setContainer(c);
            }
        } catch {
            // ignore
        }
        return { rows, memberships };
    }

    function getTenantSettingsDomainFallback() {
        try {
            if (typeof window.ms365TenantSettingsLoad === 'function') {
                const s = window.ms365TenantSettingsLoad();
                const d = s && s.domain ? String(s.domain).trim() : '';
                if (d) return d;
            }
        } catch {
            // ignore
        }
        return '';
    }

    function defaultAnlegenSchemas() {
        const now = new Date();
        const yy = String(now.getFullYear()).slice(-2);
        return {
            domain: getTenantSettingsDomainFallback() || 'ms365.schule',
            kursteamYearPrefix: 'SJ' + yy,
            kursteamPattern: '{yearPrefix} | {klasse} | {fach}',
            kursteamMailNickPattern: 'kt-{yearPrefix}-{klasse}-{fach}',
            jgPrefix: 'jg',
            jgUpper: true,
            argePrefix: 'arge',
            argeUpper: false,
            /** Schuljahrswechsel: maximale Stufenanzahl (3/4/5/8). */
            maxSchulstufen: 5,
            /** Organigramm: Geschwister nebeneinander (klassisch) oder untereinander (weniger Breite). */
            graphLayoutMode: 'horizontal'
        };
    }

    function parseSchoolYearStartYear(label) {
        const m = String(label || '').trim().match(/^(\d{4})\s*\/\s*(\d{2}|\d{4})/);
        return m ? parseInt(m[1], 10) : NaN;
    }

    function nextSchoolYearLabel(cur) {
        const y = parseSchoolYearStartYear(cur);
        if (!isFinite(y)) return currentSchoolYearLabel();
        return String(y + 1) + '/' + String(y + 2).slice(2);
    }

    function gradeFromGraduationYear(gradYear, schoolYearLabel, maxStufen) {
        const gy = String(gradYear || '').trim();
        const sy = parseSchoolYearStartYear(schoolYearLabel);
        const gyi = /^\d{4}$/.test(gy) ? parseInt(gy, 10) : NaN;
        if (!isFinite(gyi) || !isFinite(sy)) return NaN;
        const ms = isFinite(maxStufen) ? Math.max(1, Math.min(12, Math.round(maxStufen))) : 5;
        // Abschlussjahr = Ende der höchsten Stufe. In Schuljahr sy/sy+1 gilt:
        // grade = (maxStufen+1) - (abschlussjahr - sy)
        return (ms + 1) - (gyi - sy);
    }

    function replaceLeadingNumber(label, nextGrade) {
        const s = String(label || '').trim();
        if (!s) return s;
        const g = String(Math.round(nextGrade));
        if (/^\d{1,2}/.test(s)) return s.replace(/^\d{1,2}/, g);
        return s;
    }

    function normalizeGraphLayoutModeInSettings(settings) {
        if (!settings || typeof settings !== 'object') return;
        settings.graphLayoutMode = settings.graphLayoutMode === 'vertical' ? 'vertical' : 'horizontal';
    }

    function normNickPart(s) {
        return String(s || '').trim().replace(/[^A-Za-z0-9-]/g, '');
    }

    function normNickPrefixLower(s, fallback) {
        const t = String(s || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
        return t || String(fallback || '').trim() || '';
    }

    function maybeUpperByFlag(s, upper) {
        const v = String(s || '');
        return upper ? v.toUpperCase() : v.toLowerCase();
    }

    function buildKursteamNameFromTemplate(tpl, ctx) {
        const template = String(tpl || '').trim() || '{yearPrefix} | {klasse} | {fach}';
        return template
            .replaceAll('{yearPrefix}', String(ctx.yearPrefix || ''))
            .replaceAll('{klasse}', String(ctx.klasse || ''))
            .replaceAll('{fach}', String(ctx.fach || ''))
            .replaceAll('{gruppe}', String(ctx.gruppe || ''));
    }

    function buildKursteamMailNickFromTemplate(tpl, ctx) {
        const template = String(tpl || '').trim() || 'kt-{yearPrefix}-{klasse}-{fach}';
        const raw = template
            .replaceAll('{yearPrefix}', String(ctx.yearPrefix || ''))
            .replaceAll('{klasse}', String(ctx.klasse || ''))
            .replaceAll('{fach}', String(ctx.fach || ''))
            .replaceAll('{gruppe}', String(ctx.gruppe || ''));
        return String(raw || '')
            .trim()
            .toLowerCase()
            .replace(/\s+/g, '-')
            .replace(/[^a-z0-9-]/g, '')
            .replace(/-+/g, '-')
            .replace(/^-|-$/g, '');
    }

    function buildJgMailNick(schema, year, suffix) {
        const prefix = normNickPrefixLower(schema?.jgPrefix, 'jg');
        const suf = maybeUpperByFlag(normNickPart(suffix), !!schema?.jgUpper);
        const y = String(year || '').trim().replace(/[^0-9]/g, '').slice(0, 4);
        const sep = suf ? '-' : '';
        return (prefix + y + sep + suf).replace(/[^A-Za-z0-9-]/g, '');
    }

    function buildArgeMailNick(schema, shortCode) {
        const prefix = normNickPrefixLower(schema?.argePrefix, 'arge');
        const code = maybeUpperByFlag(normNickPart(shortCode), !!schema?.argeUpper);
        const sep = code ? '-' : '';
        return (prefix + sep + code).replace(/[^A-Za-z0-9-]/g, '');
    }

    function buildMailNickFromLabel(label) {
        return String(label || '')
            .trim()
            .toLowerCase()
            .replace(/\s+/g, '-')
            .replace(/[^a-z0-9-]/g, '')
            .replace(/-+/g, '-')
            .replace(/^-|-$/g, '');
    }

    function mailNicknameFromUpn(upn) {
        const u = String(upn || '').trim().toLowerCase();
        const local = (u.split('@')[0] || '').trim();
        let nick = buildMailNickFromLabel(local);
        if (!nick) nick = 'u' + String(Math.random()).toString(16).slice(2, 10);
        if (nick.length > 64) nick = nick.slice(0, 64);
        return nick;
    }

    function generateGraphTempPassword() {
        const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789';
        const sym = '!@#$%';
        let s = '';
        for (let i = 0; i < 14; i++) s += chars.charAt(Math.floor(Math.random() * chars.length));
        return s + sym.charAt(Math.floor(Math.random() * sym.length)) + '1aA';
    }

    function escapeHtml(s) {
        return String(s || '')
            .replaceAll('&', '&amp;')
            .replaceAll('<', '&lt;')
            .replaceAll('>', '&gt;')
            .replaceAll('"', '&quot;')
            .replaceAll("'", '&#039;');
    }

    function loadMatchState() {
        try {
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getContainer === 'function') {
                const c = window.ms365AppDataV2.getContainer();
                if (c && c.match && c.match.links && typeof c.match.links === 'object') {
                    return { links: c.match.links };
                }
            }
            const raw = localStorage.getItem(STORAGE_MATCH_KEY);
            const obj = raw ? safeJsonParse(raw) : null;
            const links = obj && obj.links && typeof obj.links === 'object' ? obj.links : {};
            return { links };
        } catch {
            return { links: {} };
        }
    }

    function saveMatchState(links) {
        const out = links && typeof links === 'object' ? links : {};
        try {
            localStorage.setItem(STORAGE_MATCH_KEY, JSON.stringify({ links: out }));
        } catch {
            // ignore
        }
        try {
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getContainer === 'function' && typeof window.ms365AppDataV2.setContainer === 'function') {
                const c = window.ms365AppDataV2.getContainer();
                c.match = { links: out };
                window.ms365AppDataV2.setContainer(c);
            }
        } catch {
            // ignore
        }
        return out;
    }

    function loadTenantCache() {
        try {
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getContainer === 'function') {
                const c = window.ms365AppDataV2.getContainer();
                const cache = c && c.tenant && c.tenant.cache && typeof c.tenant.cache === 'object' ? c.tenant.cache : null;
                if (cache) {
                    const rows = Array.isArray(cache.rows) ? cache.rows : [];
                    const users = Array.isArray(cache.users) ? cache.users : [];
                    return { rows, users, loadedAt: cache.loadedAt ? String(cache.loadedAt) : '' };
                }
            }
            const raw = localStorage.getItem(STORAGE_TENANT_CACHE_KEY);
            const obj = raw ? safeJsonParse(raw) : null;
            const rows = obj && Array.isArray(obj.rows) ? obj.rows : [];
            const users = obj && Array.isArray(obj.users) ? obj.users : [];
            return { rows, users, loadedAt: obj && obj.loadedAt ? String(obj.loadedAt) : '' };
        } catch {
            return { rows: [], users: [], loadedAt: '' };
        }
    }

    function saveTenantCache(rows, users) {
        const out = Array.isArray(rows) ? rows : [];
        const prev = loadTenantCache();
        const u = users !== undefined ? (Array.isArray(users) ? users : []) : prev.users || [];
        try {
            localStorage.setItem(
                STORAGE_TENANT_CACHE_KEY,
                JSON.stringify({ rows: out, users: u, loadedAt: new Date().toISOString() })
            );
        } catch {
            // ignore
        }
        try {
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getContainer === 'function' && typeof window.ms365AppDataV2.setContainer === 'function') {
                const c = window.ms365AppDataV2.getContainer();
                c.tenant = { cache: { rows: out, users: u, loadedAt: new Date().toISOString() } };
                window.ms365AppDataV2.setContainer(c);
            }
        } catch {
            // ignore
        }
    }


    function currentSchoolYearLabel() {
        const y = new Date().getFullYear();
        return String(y) + '/' + String(y + 1).slice(2);
    }

    function normalizeClassLabel(c) {
        const code = c && c.code ? String(c.code).trim() : '';
        const name = c && c.name ? String(c.name).trim() : '';
        return code || name || '';
    }

    function deriveGradeFromClassLabel(label) {
        const m = String(label || '').trim().match(/^(\d{1,2})/);
        return m ? m[1] : '';
    }

    function buildDemoFromTenantSettings() {
        if (typeof window.ms365TenantSettingsLoad !== 'function') return null;
        const s = window.ms365TenantSettingsLoad();
        const classes = s && Array.isArray(s.classes) ? s.classes : [];
        if (!classes.length) return null;

        const schuljahr = currentSchoolYearLabel();
        const gradeMap = new Map();
        const rows = [];

        const grades = Array.from(
            new Set(classes.map((c) => deriveGradeFromClassLabel(normalizeClassLabel(c))).filter(Boolean))
        ).sort(compareDe);
        grades.forEach((g) => {
            const id = uid();
            gradeMap.set(g, id);
            rows.push({
                id,
                parentId: '',
                typ: 'Jahrgang',
                bezeichnung: 'Jahrgang ' + g,
                schuljahr,
                status: 'Aktiv',
                syncStatus: 'Ausstehend',
                letzteFehlermeldung: ''
            });
        });

        classes.forEach((c) => {
            const label = normalizeClassLabel(c);
            if (!label) return;
            const g = deriveGradeFromClassLabel(label);
            const parentId = g && gradeMap.has(g) ? gradeMap.get(g) : '';
            rows.push({
                id: uid(),
                parentId: parentId || '',
                typ: 'Klasse',
                bezeichnung: label,
                schuljahr,
                status: 'Aktiv',
                syncStatus: 'Ausstehend',
                letzteFehlermeldung: ''
            });
        });

        // sinnvolle Standard-Gruppen
        rows.push({ id: uid(), parentId: '', typ: 'Gruppe', bezeichnung: 'Lehrer:innen', schuljahr, status: 'Aktiv', syncStatus: 'Ausstehend', letzteFehlermeldung: '' });
        rows.push({ id: uid(), parentId: '', typ: 'Gruppe', bezeichnung: 'Schüler:innen', schuljahr, status: 'Aktiv', syncStatus: 'Ausstehend', letzteFehlermeldung: '' });

        return { rows, tenantSettings: s };
    }

    function normRoleKey(s) {
        return String(s || '')
            .trim()
            .toLowerCase()
            .replace(/\s+/g, ' ')
            .replace(/[:;]/g, '');
    }

    function buildDemoRows() {
        const fromTenant = buildDemoFromTenantSettings();
        if (fromTenant && fromTenant.rows && fromTenant.rows.length) return fromTenant;

        const schuljahr = currentSchoolYearLabel();
        const jg1 = { id: uid(), parentId: '', typ: 'Jahrgang', bezeichnung: 'Jahrgang 1', schuljahr, status: 'Aktiv', syncStatus: 'Ausstehend', letzteFehlermeldung: '' };
        const jg2 = { id: uid(), parentId: '', typ: 'Jahrgang', bezeichnung: 'Jahrgang 2', schuljahr, status: 'Aktiv', syncStatus: 'Ok', letzteFehlermeldung: '' };
        const k1a = { id: uid(), parentId: jg1.id, typ: 'Klasse', bezeichnung: '1A', schuljahr, status: 'Aktiv', syncStatus: 'Abweichung', letzteFehlermeldung: 'Mitgliedschaft weicht ab (Mock).' };
        const k1b = { id: uid(), parentId: jg1.id, typ: 'Klasse', bezeichnung: '1B', schuljahr, status: 'Aktiv', syncStatus: 'Ok', letzteFehlermeldung: '' };
        const k2a = { id: uid(), parentId: jg2.id, typ: 'Klasse', bezeichnung: '2A', schuljahr, status: 'Aktiv', syncStatus: 'Fehler', letzteFehlermeldung: 'Team konnte nicht bereitgestellt werden (Mock).' };
        const ar1 = { id: uid(), parentId: '', typ: 'Arbeitsgemeinschaft', bezeichnung: 'ARGE Robotik', schuljahr, status: 'Aktiv', syncStatus: 'Ok', letzteFehlermeldung: '' };
        const ar2 = { id: uid(), parentId: '', typ: 'Arbeitsgemeinschaft', bezeichnung: 'ARGE Chor', schuljahr, status: 'Inaktiv', syncStatus: 'Ausstehend', letzteFehlermeldung: '' };
        return { rows: [jg1, jg2, k1a, k1b, k2a, ar1, ar2], tenantSettings: null };
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
            // fallback: kein blocking alert für längere Vorgänge
            try {
                const el = getEl('ssTenantProgressText');
                if (el) el.textContent = String(msg);
            } catch {
                // ignore
            }
        }
    }

    function setTenantProgress(visible, text, ratio) {
        const wrap = getEl('ssTenantProgressWrap');
        const txt = getEl('ssTenantProgressText');
        const bar = getEl('ssTenantProgressBar');
        const pct = getEl('ssTenantProgressPct');
        if (wrap) wrap.style.display = visible ? '' : 'none';
        if (txt && text) txt.textContent = String(text);
        const r = typeof ratio === 'number' && isFinite(ratio) ? Math.max(0, Math.min(1, ratio)) : null;
        if (bar) bar.style.width = r === null ? '0%' : String(Math.round(r * 100)) + '%';
        if (pct) pct.textContent = r === null ? '–' : String(Math.round(r * 100)) + '%';
    }

    function formatDateTimeAT(iso) {
        const s = String(iso || '').trim();
        if (!s) return '';
        try {
            const d = new Date(s);
            if (isNaN(d.getTime())) return s;
            const out = new Intl.DateTimeFormat('de-AT', {
                day: '2-digit',
                month: '2-digit',
                year: 'numeric',
                hour: '2-digit',
                minute: '2-digit',
                hour12: false
            }).format(d);
            return String(out).replace(',', '').replace(/\s+/g, ' ').trim();
        } catch {
            return s;
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
        const m365 = rows.filter((r) => r && r.typ === 'Gruppe').length;
        const sec = rows.filter((r) => r && (r.typ === 'Sicherheitsgruppe' || r.typ === 'E‑Mail‑Sicherheitsgruppe')).length;
        return { total, teams, m365, sec };
    }

    function setModeHint(mode, tenantLoadedAt) {
        const el = getEl('ssModeHint');
        if (!el) return;
        if (mode === 'tenant') {
            el.style.display = '';
            el.textContent =
                'Tenant‑Inventar: Gruppen/Teams werden live per Graph eingelesen. Updates sind für Anzeigename/Beschreibung möglich.' +
                (tenantLoadedAt ? ' Letztes Einlesen: ' + new Date(tenantLoadedAt).toLocaleString() : '');
        } else if (mode === 'match') {
            el.style.display = '';
            el.textContent =
                'Abgleich: SOLL‑Einheiten werden mit bestehenden Tenant‑Gruppen/Teams verknüpft (Mapping lokal gespeichert). Über die Registerkarte „Organigramm“ siehst du die SOLL‑Struktur vernetzt; im Baum und Organigramm kannst du per Drag&Drop umsortieren.' +
                (tenantLoadedAt ? ' Tenant zuletzt eingelesen: ' + new Date(tenantLoadedAt).toLocaleString() : '');
        } else {
            el.textContent = '';
            el.style.display = 'none';
        }
    }

    function getFilterState() {
        const schuljahr = normStr(getEl('ssFilterSchuljahr')?.value);
        const typ = normStr(getEl('ssFilterTyp')?.value);
        const text = normStr(getEl('ssFilterText')?.value).toLowerCase();
        const visibility = normStr(getEl('ssTenantVisibilityFilter')?.value);
        const roster = normStr(getEl('ssTenantRosterFilter')?.value);
        return { schuljahr, typ, text, visibility, roster };
    }

    function applyFilters(rows, mode) {
        const f = getFilterState();
        return rows.filter((r) => {
            if (!r) return false;
            if (mode !== 'tenant') {
                if (f.schuljahr && String(r.schuljahr || '') !== f.schuljahr) return false;
            }
            if (f.typ) {
                if (mode === 'tenant' && f.typ === 'Kursteam') {
                    // Kursteams = HiddenMembership (praktischer Indikator für Kurs-Teams)
                    if (!r.hiddenMembership) return false;
                } else {
                    if (String(r.typ || '') !== f.typ) return false;
                }
            }
            if (mode === 'tenant') {
                if (f.visibility) {
                    const v = String(r.visibility || '');
                    if (v !== f.visibility) return false;
                }
                if (f.roster) {
                    const oc = typeof r.ownerCount === 'number' ? r.ownerCount : -1;
                    const mc = typeof r.memberCount === 'number' ? r.memberCount : -1;
                    if (f.roster === 'noOwners') {
                        if (oc !== 0) return false;
                    } else if (f.roster === 'noMembers') {
                        if (mc !== 0) return false;
                    } else if (f.roster === 'noOwnersNoMembers') {
                        if (oc !== 0 || mc !== 0) return false;
                    }
                }
            }
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
                const rank = (x) =>
                    x.typ === 'Jahrgang'
                        ? 1
                        : x.typ === 'Klasse'
                          ? 2
                          : x.typ === 'Arbeitsgemeinschaft'
                            ? 3
                            : x.typ === 'Kursteam'
                              ? 4
                              : 5;
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

    const STRUCT_TREE_ROOT_STUDENTS = '__root_students__';
    const STRUCT_TREE_ROOT_TEACHERS = '__root_teachers__';
    const STRUCT_TREE_ROOT_ADMIN = '__root_admin__';
    const STRUCT_FOLDER_ARGES = '__ss_folder_arges__';
    const STRUCT_FOLDER_FACHSCHAFTEN = '__ss_folder_fachschaften__';

    /** Nur Anzeige/DnD im Organigramm: keine persistierten Zeilen */
    function isStructureSyntheticGraphNodeId(id) {
        const s = String(id || '');
        if (!s) return false;
        if (s.startsWith('__root_')) return true;
        if (s === STRUCT_FOLDER_ARGES || s === STRUCT_FOLDER_FACHSCHAFTEN) return true;
        if (s.startsWith('__ss_fach__')) return true;
        if (s.startsWith('__ss_kursteams__')) return true;
        return false;
    }

    /** Die drei einklappbaren Hauptäste im Baum (mit Detail-Panel, gespeichert unter settings.structRootDetails). */
    function isStructureTreeRootId(id) {
        const s = String(id || '');
        return s === STRUCT_TREE_ROOT_STUDENTS || s === STRUCT_TREE_ROOT_TEACHERS || s === STRUCT_TREE_ROOT_ADMIN;
    }

    function structureTreeRootDefaultTitle(rootId) {
        const s = String(rootId || '');
        if (s === STRUCT_TREE_ROOT_STUDENTS) return 'Schüler:innen';
        if (s === STRUCT_TREE_ROOT_TEACHERS) return 'Lehrer:innen';
        if (s === STRUCT_TREE_ROOT_ADMIN) return 'Verwaltung';
        return '';
    }

    function structureTreeRootTitle(rootId, structRootDetails) {
        const o = structRootDetails && typeof structRootDetails === 'object' ? structRootDetails[String(rootId)] : null;
        const custom = o && normStr(o.bezeichnung);
        if (custom) return custom;
        return structureTreeRootDefaultTitle(rootId);
    }

    function defaultStructureTreeRootRow(rootId) {
        if (!isStructureTreeRootId(rootId)) return null;
        const typ =
            String(rootId) === STRUCT_TREE_ROOT_STUDENTS
                ? 'SchuelerInnen'
                : String(rootId) === STRUCT_TREE_ROOT_TEACHERS
                  ? 'LehrerInnen'
                  : 'Verwaltung';
        return {
            id: String(rootId),
            parentId: '',
            typ,
            bezeichnung: structureTreeRootDefaultTitle(rootId),
            beschreibung: '',
            schuljahr: '',
            status: 'Aktiv',
            syncStatus: 'Ausstehend',
            letzteFehlermeldung: '',
            jgYear: '',
            jgSuffix: '',
            argeCode: '',
            argeName: '',
            ktKlasse: '',
            ktFach: '',
            ktGruppe: '',
            tenantGroupId: '',
            tenantMailNickname: '',
            tenantTarget: '',
            tenantVisibility: '',
            isStructureTreeRoot: true
        };
    }

    function mergeStructureTreeRootRow(rootId, structRootDetails) {
        const base = defaultStructureTreeRootRow(rootId);
        if (!base) return null;
        const o = structRootDetails && typeof structRootDetails === 'object' ? structRootDetails[String(rootId)] : null;
        if (!o || typeof o !== 'object') return base;
        return Object.assign({}, base, o, {
            id: String(rootId),
            parentId: '',
            typ: base.typ,
            isStructureTreeRoot: true
        });
    }

    function pickStorableStructureTreeRootFields(row) {
        return {
            bezeichnung: normStr(row.bezeichnung),
            beschreibung: normStr(row.beschreibung),
            schuljahr: normStr(row.schuljahr),
            status: normStr(row.status) || 'Aktiv',
            syncStatus: normStr(row.syncStatus) || 'Ausstehend',
            letzteFehlermeldung: normStr(row.letzteFehlermeldung),
            jgYear: normStr(row.jgYear),
            jgSuffix: normStr(row.jgSuffix),
            argeCode: normStr(row.argeCode),
            argeName: normStr(row.argeName),
            ktKlasse: normStr(row.ktKlasse),
            ktFach: normStr(row.ktFach),
            ktGruppe: normStr(row.ktGruppe),
            tenantGroupId: normStr(row.tenantGroupId),
            tenantMailNickname: normStr(row.tenantMailNickname),
            tenantTarget: normStr(row.tenantTarget),
            tenantVisibility: normStr(row.tenantVisibility)
        };
    }

    function isTopVerwaltungStructureNode(r) {
        if (!r) return false;
        const name = String(r.bezeichnung || '').trim().toLowerCase();
        const pid = String(r.parentId || '').trim();
        return !pid && name === 'verwaltung';
    }

    function teacherFachVirtualId(fachLabel) {
        const key = String(fachLabel || '').trim() || '(ohne Fach)';
        return '__ss_fach__' + encodeURIComponent(key).replace(/%/g, '_');
    }

    /** Entfernt die frühere echte Container-„Gruppe Fachschaften“; Kinder werden Top-Level-Fach-Gruppen. */
    function migrateLegacyFachschaftenContainer(rows) {
        if (!rows || !rows.length) return false;
        let changed = false;
        const removeIds = new Set();
        for (const r of rows) {
            if (!r || r.typ !== 'Gruppe') continue;
            const isMarked = r.fachschaftenRoot === true;
            const isNamedContainer =
                String(r.parentId || '') === '' &&
                String(r.bezeichnung || '').trim().toLowerCase() === 'fachschaften' &&
                !r.fachschaftFach;
            if (!isMarked && !isNamedContainer) continue;
            const cid = String(r.id);
            for (const x of rows) {
                if (x && String(x.parentId || '') === cid) {
                    x.parentId = '';
                    if (x.typ === 'Gruppe') x.fachschaftFach = true;
                    changed = true;
                }
            }
            removeIds.add(cid);
            changed = true;
        }
        if (!removeIds.size) return changed;
        for (let i = rows.length - 1; i >= 0; i--) {
            if (rows[i] && removeIds.has(String(rows[i].id))) rows.splice(i, 1);
        }
        return changed;
    }

    /**
     * Keine persistierte „Fachschaften“-Gruppe: nur je Fach eine Gruppe (Top-Level), Anzeige unter virtuellem Ordner wie bei ARGEs.
     * @returns {boolean} true wenn Daten geändert wurden
     */
    function ensureFachschaftFachGruppen(rows, tenantSettings) {
        if (!rows || !rows.length) return false;
        let changed = migrateLegacyFachschaftenContainer(rows);
        const schuljahr = currentSchoolYearLabel();
        const fachKeys = new Set();
        for (const r of rows) {
            if (r && r.typ === 'Kursteam') {
                fachKeys.add(String(r.ktFach || '').trim() || '(ohne Fach)');
            }
        }
        const subj = tenantSettings && Array.isArray(tenantSettings.subjects) ? tenantSettings.subjects : [];
        for (const s of subj) {
            const key = String((s && (s.code || s.name)) || '').trim();
            if (key) fachKeys.add(key);
        }
        for (const f of Array.from(fachKeys).sort(compareDe)) {
            const has = rows.some(
                (r) =>
                    r &&
                    r.typ === 'Gruppe' &&
                    r.fachschaftFach &&
                    String(r.parentId || '') === '' &&
                    (String(r.ktFach || '').trim() === f ||
                        (!String(r.ktFach || '').trim() && String(r.bezeichnung || '').trim() === f))
            );
            if (!has) {
                rows.push({
                    id: uid(),
                    parentId: '',
                    typ: 'Gruppe',
                    bezeichnung: f,
                    ktFach: f,
                    schuljahr,
                    status: 'Aktiv',
                    syncStatus: 'Ausstehend',
                    letzteFehlermeldung: '',
                    fachschaftFach: true
                });
                changed = true;
            }
        }
        return changed;
    }

    function sortStructureTreeChildren(list) {
        list.sort((a, b) => {
            const rank = (x) =>
                x.typ === 'Jahrgang'
                    ? 1
                    : x.typ === 'Klasse'
                      ? 2
                      : x.typ === 'Kursteam'
                        ? 3
                        : x.typ === 'Arbeitsgemeinschaft'
                          ? 4
                          : x.typ === 'Gruppe'
                            ? 5
                            : x.typ === 'Person'
                              ? 6
                              : 99;
            const ra = rank(a);
            const rb = rank(b);
            if (ra !== rb) return ra - rb;
            return compareDe(a.bezeichnung, b.bezeichnung);
        });
    }

    /**
     * Baum wie im Organigramm: virtuelle Wurzeln Schüler:innen / Lehrer:innen / Verwaltung,
     * unter Lehrer:innen Standard-Ordner „ARGEs“ und „Fachschaften“,
     * plus einklappbare Teilbäume (ids teilen sich mit GRAPH_COLLAPSE_KEY).
     * Kursteams erscheinen als echte Knoten unter „Klasse“ (anlegbar, per Abgleich mit Tenant verknüpfbar).
     */
    function buildStructuredTreeOrder(rows, collapsedSet, structRootDetails) {
        const baseRows = (rows || []).filter((r) => r && r.id);
        const idMap = byId(baseRows);
        const children = new Map();
        for (const r of baseRows) {
            const pid = String(r.parentId || '');
            if (!children.has(pid)) children.set(pid, []);
            children.get(pid).push(r);
        }
        for (const list of children.values()) sortStructureTreeChildren(list);

        function virtualRootForTop(r) {
            if (isTopVerwaltungStructureNode(r)) return STRUCT_TREE_ROOT_ADMIN;
            if (inferRootForType(r.typ) === 'LehrerInnen') return STRUCT_TREE_ROOT_TEACHERS;
            return STRUCT_TREE_ROOT_STUDENTS;
        }

        const buckets = new Map([
            [STRUCT_TREE_ROOT_STUDENTS, []],
            [STRUCT_TREE_ROOT_TEACHERS, []],
            [STRUCT_TREE_ROOT_ADMIN, []],
        ]);
        for (const r of baseRows) {
            const pid = String(r.parentId || '');
            if (pid && idMap.has(pid)) continue;
            const b = buckets.get(virtualRootForTop(r));
            if (b) b.push(r);
        }
        for (const list of buckets.values()) sortStructureTreeChildren(list);

        const argeTop = baseRows
            .filter((r) => {
                if (!r || r.typ !== 'Arbeitsgemeinschaft') return false;
                const pid = String(r.parentId || '');
                return !pid || !idMap.has(pid);
            })
            .slice()
            .sort((a, b) => compareDe(a.bezeichnung, b.bezeichnung));

        const teacherOtherTop = baseRows
            .filter((r) => {
                if (!r) return false;
                const pid = String(r.parentId || '');
                if (pid && idMap.has(pid)) return false;
                if (virtualRootForTop(r) !== STRUCT_TREE_ROOT_TEACHERS) return false;
                if (r.typ === 'Arbeitsgemeinschaft') return false;
                if (r.fachschaftFach) return false;
                return true;
            })
            .slice()
            .sort((a, b) => compareDe(a.bezeichnung, b.bezeichnung));

        const fachGruppenTop = baseRows
            .filter((r) => {
                if (!r || r.typ !== 'Gruppe' || !r.fachschaftFach) return false;
                const pid = String(r.parentId || '');
                return !pid || !idMap.has(pid);
            })
            .slice()
            .sort((a, b) => compareDe(a.bezeichnung, b.bezeichnung));

        const virtualRoots = [
            { id: STRUCT_TREE_ROOT_STUDENTS, label: structureTreeRootTitle(STRUCT_TREE_ROOT_STUDENTS, structRootDetails), typ: 'SchuelerInnen' },
            { id: STRUCT_TREE_ROOT_TEACHERS, label: structureTreeRootTitle(STRUCT_TREE_ROOT_TEACHERS, structRootDetails), typ: 'LehrerInnen' },
            { id: STRUCT_TREE_ROOT_ADMIN, label: structureTreeRootTitle(STRUCT_TREE_ROOT_ADMIN, structRootDetails), typ: 'Verwaltung' },
        ];

        const collapsed = collapsedSet || new Set();
        const out = [];

        function walkReal(r, depth) {
            const list = children.get(String(r.id)) || [];
            const hasKids = list.length > 0;
            out.push({ r, depth, virtual: false, hasKids });
            if (collapsed.has(String(r.id))) return;
            for (const k of list) walkReal(k, depth + 1);
        }

        for (const vr of virtualRoots) {
            const rootRowMerged = mergeStructureTreeRootRow(vr.id, structRootDetails);
            const rootSyncStatus = rootRowMerged && rootRowMerged.syncStatus ? String(rootRowMerged.syncStatus) : 'Ausstehend';
            if (vr.id === STRUCT_TREE_ROOT_TEACHERS) {
                const hasFachBranch = fachGruppenTop.length > 0;
                const hasTeacherKids =
                    argeTop.length > 0 || hasFachBranch || teacherOtherTop.length > 0;
                out.push({
                    virtual: true,
                    kind: 'root',
                    rootId: vr.id,
                    label: vr.label,
                    typ: vr.typ,
                    depth: 0,
                    hasKids: hasTeacherKids,
                    rootSyncStatus,
                });
                if (!hasTeacherKids) continue;
                if (collapsed.has(vr.id)) continue;

                out.push({
                    virtual: true,
                    kind: 'folder',
                    rootId: STRUCT_FOLDER_ARGES,
                    label: 'ARGEs',
                    typ: 'Gruppe',
                    typLabel: 'Ordner',
                    depth: 1,
                    hasKids: argeTop.length > 0,
                });
                if (argeTop.length && !collapsed.has(STRUCT_FOLDER_ARGES)) {
                    for (const r of argeTop) walkReal(r, 2);
                }

                out.push({
                    virtual: true,
                    kind: 'folder',
                    rootId: STRUCT_FOLDER_FACHSCHAFTEN,
                    label: 'Fachschaften',
                    typ: 'Gruppe',
                    typLabel: 'Ordner',
                    depth: 1,
                    hasKids: hasFachBranch,
                });
                if (hasFachBranch && !collapsed.has(STRUCT_FOLDER_FACHSCHAFTEN)) {
                    for (const r of fachGruppenTop) walkReal(r, 2);
                }

                for (const r of teacherOtherTop) walkReal(r, 1);
                continue;
            }

            const kids = buckets.get(vr.id) || [];
            out.push({
                virtual: true,
                kind: 'root',
                rootId: vr.id,
                label: vr.label,
                typ: vr.typ,
                depth: 0,
                hasKids: kids.length > 0,
                rootSyncStatus,
            });
            if (!kids.length) continue;
            if (collapsed.has(vr.id)) continue;
            for (const r of kids) walkReal(r, 1);
        }
        return out;
    }

    function typeIcon(typ) {
        const t = String(typ || '');
        if (t === 'Jahrgang') return 'bi-layers';
        if (t === 'Klasse') return 'bi-collection';
        if (t === 'Kursteam') return 'bi-mortarboard';
        if (t === 'Arbeitsgemeinschaft') return 'bi-people-gear';
        if (t === 'Gruppe') return 'bi-people';
        if (t === 'Person') return 'bi-person';
        if (t === 'SchuelerInnen') return 'bi-people-fill';
        if (t === 'LehrerInnen') return 'bi-person-badge';
        if (t === 'Verwaltung') return 'bi-building';
        // Tenant / Microsoft 365 (Verwalten-Tab)
        if (t === 'Team') return 'bi-camera-video-fill';
        if (t === 'Sicherheitsgruppe') return 'bi-shield-lock';
        if (t === 'E-Mail-Sicherheitsgruppe' || t === 'E‑Mail‑Sicherheitsgruppe') return 'bi-envelope-paper';
        return 'bi-folder2';
    }

    /** Icon für echte Zeilen in der Baumansicht (inkl. Kursteam über hiddenMembership im Tenant-Modus). */
    function treeIconForRow(r, mode) {
        if (!r) return typeIcon('');
        if (mode === 'tenant' && r.hiddenMembership) return 'bi-mortarboard';
        if (r.fachschaftFach) return 'bi-journal-text';
        return typeIcon(r.typ);
    }

    /** Virtuelle Baumzeilen (Wurzeln, ARGE-/Fachschaften-Ordner, Fachschaft). */
    function treeIconForVirtualItem(item) {
        if (!item || !item.virtual) return typeIcon('');
        if (item.kind === 'folder') {
            if (String(item.rootId || '') === STRUCT_FOLDER_ARGES) return 'bi-people-gear';
            if (String(item.rootId || '') === STRUCT_FOLDER_FACHSCHAFTEN) return 'bi-journals';
            return 'bi-folder2';
        }
        if (item.kind === 'fach') return 'bi-journal-text';
        if (item.kind === 'kursteams') return 'bi-mortarboard';
        return typeIcon(item.typ);
    }

    /** Organigramm: gleiche Semantik wie Baum für synthetische Ordner/Fächer. */
    function graphNodeIconClass(node) {
        if (!node) return typeIcon('');
        if (node.fachschaftFach) return 'bi-journal-text';
        if (node.isVirtualFach) return 'bi-journal-text';
        if (node.isStructureFolder) {
            const b = String(node.bezeichnung || '');
            if (b === 'ARGEs') return 'bi-people-gear';
            if (b === 'Fachschaften') return 'bi-journals';
            return 'bi-folder2';
        }
        return typeIcon(node.typ);
    }

    const GRAPH_COLLAPSE_KEY = 'ms365-ss-graph-collapsed-v1';

    function loadGraphCollapsedSet() {
        try {
            const raw = localStorage.getItem(GRAPH_COLLAPSE_KEY);
            if (!raw) return new Set();
            const arr = JSON.parse(raw);
            if (!Array.isArray(arr)) return new Set();
            return new Set(arr.map((x) => String(x)));
        } catch {
            return new Set();
        }
    }

    function saveGraphCollapsedSet(set) {
        try {
            localStorage.setItem(GRAPH_COLLAPSE_KEY, JSON.stringify(Array.from(set || []).map((x) => String(x))));
        } catch {
            // ignore
        }
    }

    function computeGraphModel(rows, collapsedSet, structRootDetails, graphLayoutMode) {
        const nodes = new Map();
        const edges = [];
        const layoutVertical = String(graphLayoutMode || '').toLowerCase() === 'vertical';

        const rootStudentsId = '__root_students__';
        const rootTeachersId = '__root_teachers__';
        const rootAdminId = '__root_admin__';
        nodes.set(rootStudentsId, {
            id: rootStudentsId,
            typ: 'SchuelerInnen',
            bezeichnung: structureTreeRootTitle(rootStudentsId, structRootDetails),
            isRoot: true
        });
        nodes.set(rootTeachersId, {
            id: rootTeachersId,
            typ: 'LehrerInnen',
            bezeichnung: structureTreeRootTitle(rootTeachersId, structRootDetails),
            isRoot: true
        });
        nodes.set(rootAdminId, {
            id: rootAdminId,
            typ: 'Verwaltung',
            bezeichnung: structureTreeRootTitle(rootAdminId, structRootDetails),
            isRoot: true
        });

        const folderArgesId = STRUCT_FOLDER_ARGES;
        const folderFachId = STRUCT_FOLDER_FACHSCHAFTEN;
        nodes.set(folderArgesId, {
            id: folderArgesId,
            typ: 'Gruppe',
            bezeichnung: 'ARGEs',
            isStructureFolder: true,
            isRoot: false
        });
        nodes.set(folderFachId, {
            id: folderFachId,
            typ: 'Gruppe',
            bezeichnung: 'Fachschaften',
            isStructureFolder: true,
            isRoot: false
        });

        function isTopVerwaltungNode(r) {
            if (!r) return false;
            const name = String(r.bezeichnung || '').trim().toLowerCase();
            const pid = String(r.parentId || '').trim();
            return !pid && name === 'verwaltung';
        }

        // nodes
        (rows || []).forEach((r) => {
            if (!r || !r.id) return;
            nodes.set(String(r.id), r);
        });

        // adjacency
        const children = new Map();
        function addChild(pid, cid) {
            const k = String(pid || '');
            if (!children.has(k)) children.set(k, []);
            children.get(k).push(String(cid));
        }

        const teacherRootOtherIds = [];

        (rows || []).forEach((r) => {
            if (!r || !r.id) return;
            const id = String(r.id);
            const pid = String(r.parentId || '');
            if (pid && nodes.has(pid)) {
                addChild(pid, id);
                edges.push({ from: pid, to: id });
                return;
            }
            const root = isTopVerwaltungNode(r)
                ? rootAdminId
                : inferRootForType(r.typ) === 'LehrerInnen'
                  ? rootTeachersId
                  : rootStudentsId;
            if (root === rootTeachersId && String(r.typ || '') === 'Arbeitsgemeinschaft') {
                addChild(folderArgesId, id);
                edges.push({ from: folderArgesId, to: id });
                return;
            }
            if (root === rootTeachersId && r.fachschaftFach) {
                addChild(folderFachId, id);
                edges.push({ from: folderFachId, to: id });
                return;
            }
            if (root === rootTeachersId) {
                teacherRootOtherIds.push(id);
                return;
            }
            addChild(root, id);
            edges.push({ from: root, to: id });
        });

        teacherRootOtherIds.sort((a, b) => {
            const ra = nodes.get(a);
            const rb = nodes.get(b);
            return compareDe(String(ra?.bezeichnung || ''), String(rb?.bezeichnung || ''));
        });

        addChild(rootTeachersId, folderArgesId);
        edges.push({ from: rootTeachersId, to: folderArgesId });
        addChild(rootTeachersId, folderFachId);
        edges.push({ from: rootTeachersId, to: folderFachId });
        for (const cid of teacherRootOtherIds) {
            addChild(rootTeachersId, cid);
            edges.push({ from: rootTeachersId, to: cid });
        }

        // sort children for stable layout (Lehrer-Wurzel behält Ordner-Reihenfolge)
        for (const [k, list] of children.entries()) {
            if (k === rootTeachersId) continue;
            list.sort((a, b) => {
                const ra = nodes.get(a);
                const rb = nodes.get(b);
                return compareDe(String(ra?.bezeichnung || ''), String(rb?.bezeichnung || ''));
            });
            children.set(k, list);
        }
        const tKids = children.get(rootTeachersId) || [];
        const want = [folderArgesId, folderFachId];
        const rest = tKids.filter((id) => want.indexOf(String(id)) === -1).sort((a, b) => {
            const ra = nodes.get(a);
            const rb = nodes.get(b);
            return compareDe(String(ra?.bezeichnung || ''), String(rb?.bezeichnung || ''));
        });
        children.set(rootTeachersId, [...want, ...rest]);

        /** „Breite“ bzw. „Höhe“ des Teilbaums in Raster‑Einheiten (bei eingeklapptem Knoten = 1). */
        function subtreeSpan(id) {
            const sid = String(id);
            if (collapsedSet && collapsedSet.has(sid)) return 1;
            const kids = children.get(String(id)) || [];
            if (!kids.length) return 1;
            return kids.reduce((acc, kid) => acc + subtreeSpan(kid), 0);
        }

        const pos = new Map();
        // Spacing für große Schulen: mehr Luft zwischen Karten/Kanten
        const xUnit = 280;
        const yUnit = 170;
        /** Vertikal: Hierarchie nach rechts, Geschwister untereinander (weniger horizontale Gesamtbreite). */
        const xDepthUnit = 300;
        /** Abstand der Knoten-Mittelpunkte in Y (Karte ~56px ±28); ~72 ≈ eine Zeile Luft zwischen den Karten. */
        const ySiblingUnitV = 72;
        const rootBlockGapV = 48;

        if (!layoutVertical) {
            function layout(id, depth, x0) {
                const w = subtreeSpan(id);
                const xCenter = x0 + (w * xUnit) / 2;
                pos.set(String(id), { x: xCenter, y: 40 + depth * yUnit, w, graphLayout: 'horizontal' });
                let cursor = x0;
                const kids = children.get(String(id)) || [];
                if (collapsedSet && collapsedSet.has(String(id))) return;
                for (const kid of kids) {
                    const kw = subtreeSpan(kid);
                    layout(kid, depth + 1, cursor);
                    cursor += kw * xUnit;
                }
            }

            const wS = subtreeSpan(rootStudentsId);
            layout(rootStudentsId, 0, 40);
            const xT = 60 + wS * xUnit + 120;
            layout(rootTeachersId, 0, xT);
            const wT = subtreeSpan(rootTeachersId);
            layout(rootAdminId, 0, xT + wT * xUnit + 120);
        } else {
            function layoutV(id, depth, y0) {
                const h = subtreeSpan(id);
                const xCenter = 80 + depth * xDepthUnit;
                const yCenter = y0 + (h * ySiblingUnitV) / 2;
                pos.set(String(id), { x: xCenter, y: yCenter, w: h, graphLayout: 'vertical' });
                if (collapsedSet && collapsedSet.has(String(id))) return;
                const kids = children.get(String(id)) || [];
                let cursor = y0;
                for (const kid of kids) {
                    const kh = subtreeSpan(kid);
                    layoutV(kid, depth + 1, cursor);
                    cursor += kh * ySiblingUnitV;
                }
            }

            let yBlock = 24;
            layoutV(rootStudentsId, 0, yBlock);
            yBlock += subtreeSpan(rootStudentsId) * ySiblingUnitV + rootBlockGapV;
            layoutV(rootTeachersId, 0, yBlock);
            yBlock += subtreeSpan(rootTeachersId) * ySiblingUnitV + rootBlockGapV;
            layoutV(rootAdminId, 0, yBlock);
        }

        // canvas size
        let maxX = 0;
        let maxY = 0;
        for (const p of pos.values()) {
            if (p.x > maxX) maxX = p.x;
            if (p.y > maxY) maxY = p.y;
        }
        const padX = layoutVertical ? 360 : 320;
        const padY = layoutVertical ? 120 : 240;
        const canvas = {
            width: Math.max(1200, Math.ceil(maxX + padX)),
            height: Math.max(650, Math.ceil(maxY + padY))
        };

        return {
            nodes,
            edges,
            pos,
            canvas,
            rootStudentsId,
            rootTeachersId,
            rootAdminId,
            children,
            graphLayout: layoutVertical ? 'vertical' : 'horizontal'
        };
    }

    function renderGraphView(
        rowsStruktur,
        selectedId,
        onSelect,
        viewport,
        collapsedSet,
        personInfoByRole,
        structRootDetails,
        graphLayoutMode,
        _onGraphAddChild
    ) {
        const wrap = getEl('ssGraphWrap');
        const nodesHost = getEl('ssGraphNodes');
        const edgesSvg = getEl('ssGraphEdges');
        if (!wrap || !nodesHost || !edgesSvg) return null;

        const model = computeGraphModel(rowsStruktur || [], collapsedSet, structRootDetails, graphLayoutMode);
        edgesSvg.setAttribute('width', String(model.canvas.width));
        edgesSvg.setAttribute('height', String(model.canvas.height));
        nodesHost.style.minHeight = String(model.canvas.height) + 'px';
        nodesHost.style.minWidth = String(model.canvas.width) + 'px';

        // Pan/Zoom transform (applied to both layers)
        const vp = viewport && typeof viewport === 'object' ? viewport : { x: 0, y: 0, scale: 1 };
        const tx = Number.isFinite(vp.x) ? vp.x : 0;
        const ty = Number.isFinite(vp.y) ? vp.y : 0;
        const sc = Number.isFinite(vp.scale) ? vp.scale : 1;
        const tr = `translate(${tx}px, ${ty}px) scale(${sc})`;
        edgesSvg.style.transformOrigin = '0 0';
        edgesSvg.style.transform = tr;
        nodesHost.style.transformOrigin = '0 0';
        nodesHost.style.transform = tr;

        // edges
        edgesSvg.replaceChildren();
        const isVert = model.graphLayout === 'vertical';
        for (const e of model.edges) {
            const p1 = model.pos.get(String(e.from));
            const p2 = model.pos.get(String(e.to));
            if (!p1 || !p2) continue;
            const line = document.createElementNS('http://www.w3.org/2000/svg', 'path');
            let d;
            if (isVert) {
                const x1 = p1.x + 120;
                const y1 = p1.y;
                const x2 = p2.x - 120;
                const y2 = p2.y;
                const midX = (x1 + x2) / 2;
                d = `M ${x1} ${y1} C ${midX} ${y1}, ${midX} ${y2}, ${x2} ${y2}`;
            } else {
                const x1 = p1.x;
                const y1 = p1.y + 46;
                const x2 = p2.x;
                const y2 = p2.y - 6;
                const midY = (y1 + y2) / 2;
                d = `M ${x1} ${y1} C ${x1} ${midY}, ${x2} ${midY}, ${x2} ${y2}`;
            }
            line.setAttribute('d', d);
            line.setAttribute('stroke', 'rgba(94,114,228,0.35)');
            line.setAttribute('stroke-width', '2');
            line.setAttribute('fill', 'none');
            edgesSvg.appendChild(line);
        }

        // nodes
        nodesHost.replaceChildren();
        for (const [id, node] of model.nodes.entries()) {
            const p = model.pos.get(String(id));
            if (!p) continue;
            const div = document.createElement('div');
            div.className = 'ss-graph-node';
            const isSelected = String(selectedId || '') === String(id);
            div.setAttribute('data-ss-node-id', String(id));
            div.setAttribute('data-ss-node-type', String(node.typ || ''));
            if (node.isStructureFolder || node.isVirtualFach) div.setAttribute('data-ss-graph-synthetic', '1');
            div.draggable = !node.isRoot && !node.isStructureFolder && !node.isVirtualFach;
            div.style.position = 'absolute';
            div.style.left = String(Math.round(p.x - 120)) + 'px';
            div.style.top = String(Math.round(p.y - 28)) + 'px';
            div.style.width = '240px';
            div.style.padding = '10px 12px';
            div.style.borderRadius = '14px';
            div.style.border = isSelected ? '2px solid rgba(45,206,137,0.7)' : '1px solid rgba(94,114,228,0.22)';
            div.style.background = '#fff';
            div.style.boxShadow = isSelected ? '0 12px 26px rgba(45, 206, 137, 0.16)' : '0 10px 22px rgba(50, 50, 93, 0.10)';
            div.style.cursor = 'pointer';
            div.style.userSelect = 'none';
            div.style.display = 'flex';
            div.style.alignItems = 'center';
            div.style.gap = '10px';
            div.style.fontWeight = '800';
            const kids = model.children.get(String(id)) || [];
            const hasKids = !!kids.length && !node.isRoot && !node.isVirtualFach;
            const isCollapsed = !!(collapsedSet && collapsedSet.has(String(id)));
            const toggleBtn = hasKids
                ? `<button type="button" class="ss-graph-toggle" data-ss-toggle-for="${escapeHtml(String(id))}" title="${isCollapsed ? 'Aufklappen' : 'Einklappen'}" aria-label="${isCollapsed ? 'Aufklappen' : 'Einklappen'}" draggable="false" style="border:1px solid rgba(94,114,228,0.22); background:#fff; border-radius:12px; padding:6px 10px; font-weight:1000; cursor:pointer; line-height:1; min-width:42px;">${isCollapsed ? '▸' : '▾'}${kids.length ? `<span class="muted" style="margin-left:6px;font-weight:900;">${kids.length}</span>` : ''}</button>`
                : '';
            const canAddChildren =
                !node.isStructureFolder && !node.isVirtualFach && allowedStructureChildTypes(String(node.typ || '')).length > 0;
            const plusBtn = canAddChildren
                ? `<button type="button" class="ss-graph-plus" data-ss-plus-for="${escapeHtml(String(id))}" title="Unterpunkt hinzufügen" aria-label="Unterpunkt hinzufügen" style="margin-left:auto; border:1px solid rgba(94,114,228,0.22); background:#fff; border-radius:12px; padding:6px 10px; font-weight:1000; cursor:pointer;">+</button>`
                : '';
            const right = `<div style="margin-left:auto;display:flex;gap:8px;align-items:center;">${toggleBtn}${plusBtn}</div>`;
            const subLabel = (() => {
                if (node.isVirtualFach) {
                    const n = Number(node.kursteamCount || 0);
                    const s = n === 1 ? '1 Kursteam' : String(n) + ' Kursteams';
                    return escapeHtml(s);
                }
                const t = String(node.typ || '');
                if (t !== 'Person') return escapeHtml(node.typ || '');
                const storedN = normStr(node.personName);
                const storedE = normStr(node.personEmail).toLowerCase();
                const info = personInfoByRole && personInfoByRole.get ? personInfoByRole.get(normRoleKey(node.bezeichnung || '')) : null;
                const n = storedN || (info && info.name ? String(info.name).trim() : '');
                const e = storedE || (info && info.email ? String(info.email).trim().toLowerCase() : '');
                if (n && e) return escapeHtml(n + ' · ' + e);
                if (n) return escapeHtml(n);
                if (e) return escapeHtml(e);
                return escapeHtml('Person');
            })();
            div.innerHTML =
                `<i class="bi ${graphNodeIconClass(node)}" style="font-size:1.05em;opacity:0.92;"></i>` +
                `<div style="min-width:0;flex:1;">` +
                `<div style="white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${escapeHtml(node.bezeichnung || '')}</div>` +
                `<div class="muted" style="font-weight:700;font-size:0.86em;margin-top:2px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${subLabel}</div>` +
                `</div>` +
                right;

            if (node.isRoot && isStructureTreeRootId(id)) {
                div.addEventListener('click', (ev) => {
                    const t = ev && ev.target;
                    if (t && t.closest && (t.closest('.ss-graph-toggle') || t.closest('.ss-graph-plus'))) return;
                    if (typeof onSelect === 'function') onSelect(String(id), { openDetails: false });
                });
                div.addEventListener('dblclick', (ev) => {
                    const t = ev && ev.target;
                    if (t && t.closest && (t.closest('.ss-graph-toggle') || t.closest('.ss-graph-plus'))) return;
                    try {
                        ev.preventDefault();
                        ev.stopPropagation();
                    } catch {
                        /* ignore */
                    }
                    if (typeof onSelect === 'function') onSelect(String(id), { openDetails: true });
                });
            } else if (!node.isRoot && !node.isStructureFolder && !node.isVirtualFach) {
                // Single click: nur selektieren (kein Pop-Up).
                div.addEventListener('click', (ev) => {
                    const t = ev && ev.target;
                    if (t && t.closest && (t.closest('.ss-graph-toggle') || t.closest('.ss-graph-plus'))) return;
                    if (typeof onSelect === 'function') onSelect(String(id), { openDetails: false });
                });
                // Double click: Details öffnen.
                div.addEventListener('dblclick', (ev) => {
                    const t = ev && ev.target;
                    if (t && t.closest && (t.closest('.ss-graph-toggle') || t.closest('.ss-graph-plus'))) return;
                    try {
                        ev.preventDefault();
                        ev.stopPropagation();
                    } catch {
                        // ignore
                    }
                    if (typeof onSelect === 'function') onSelect(String(id), { openDetails: true });
                });
            }
            nodesHost.appendChild(div);
        }

        return model;
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
                          { v: 'Kursteam', t: 'Kursteams (HiddenMembership)' },
                          { v: 'Team', t: 'Team' },
                          { v: 'Gruppe', t: 'M365‑Gruppe' },
                          { v: 'Sicherheitsgruppe', t: 'Sicherheitsgruppe' },
                          { v: 'E‑Mail‑Sicherheitsgruppe', t: 'E‑Mail‑Sicherheitsgruppe' }
                      ]
                    : [
                          { v: '', t: '(alle)' },
                          { v: 'Jahrgang', t: 'Jahrgang' },
                          { v: 'Klasse', t: 'Klasse' },
                          { v: 'Arbeitsgemeinschaft', t: 'Arbeitsgemeinschaft' },
                          { v: 'Kursteam', t: 'Kursteam' },
                          { v: 'Gruppe', t: 'Gruppe' },
                          { v: 'Person', t: 'Person' }
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
            if (elAbw) elAbw.textContent = String(ts.m365);
            if (elErr) elErr.textContent = String(ts.sec);
            const labAktiv = elAktiv && elAktiv.parentElement ? elAktiv.parentElement.querySelector('.l') : null;
            const labAbw = elAbw && elAbw.parentElement ? elAbw.parentElement.querySelector('.l') : null;
            const labErr = elErr && elErr.parentElement ? elErr.parentElement.querySelector('.l') : null;
            if (labAktiv) labAktiv.textContent = 'Teams';
            if (labAbw) labAbw.textContent = 'M365‑Gruppen';
            if (labErr) labErr.textContent = 'Sicherheitsgruppen';
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

    function renderTree(rows, selectedId, mode, collapsedSet, onToggleCollapse, structRootDetails, onStructureTreeAdd) {
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

        const coll = collapsedSet || new Set();

        const ordered =
            mode === 'tenant'
                ? visible
                      .slice()
                      .sort((a, b) => compareDe(a.bezeichnung, b.bezeichnung))
                      .map((r) => ({ r, depth: 0 }))
                : buildStructuredTreeOrder(visible, collapsedSet, structRootDetails);

        for (const item of ordered) {
            if (item.virtual) {
                const li = document.createElement('li');
                const isSelectableRoot = (mode === 'struktur' || mode === 'match') && item.kind === 'root';
                const showBranchToggle = (mode === 'struktur' || mode === 'match') && item.hasKids && onToggleCollapse;

                if (showBranchToggle) {
                    li.style.display = 'flex';
                    li.style.alignItems = 'stretch';
                    const togg = document.createElement('button');
                    togg.type = 'button';
                    togg.setAttribute(
                        'aria-expanded',
                        coll.has(String(item.rootId)) ? 'false' : 'true'
                    );
                    togg.setAttribute('aria-label', coll.has(String(item.rootId)) ? 'Bereich aufklappen' : 'Bereich einklappen');
                    togg.style.flexShrink = '0';
                    togg.style.width = '32px';
                    togg.style.alignSelf = 'center';
                    togg.style.border = 'none';
                    togg.style.background = 'transparent';
                    togg.style.cursor = 'pointer';
                    togg.style.color = '#32325d';
                    togg.textContent = coll.has(String(item.rootId)) ? '▸' : '▾';
                    togg.addEventListener('click', (e) => {
                        e.stopPropagation();
                        onToggleCollapse(String(item.rootId));
                    });
                    li.appendChild(togg);
                }

                const isKursteamsPlaceholder = item.kind === 'kursteams';
                const row = isSelectableRoot || isKursteamsPlaceholder ? document.createElement('button') : document.createElement('div');
                if (isSelectableRoot) {
                    row.type = 'button';
                    row.dataset.ssSelect = String(item.rootId);
                    row.dataset.ssType = String(item.typ || '');
                    row.setAttribute(
                        'aria-current',
                        selectedId && String(selectedId) === String(item.rootId) ? 'true' : 'false'
                    );
                    row.style.border = 'none';
                    row.style.background = 'rgba(50, 50, 93, 0.06)';
                    row.style.font = 'inherit';
                    row.style.textAlign = 'left';
                    row.style.cursor = 'pointer';
                    row.style.width = showBranchToggle ? 'auto' : '100%';
                } else if (isKursteamsPlaceholder) {
                    row.type = 'button';
                    row.dataset.ssOpenKursteams = '1';
                    row.dataset.ssKurClassId = String(item.classId || '');
                    row.setAttribute('aria-label', 'Kursteams öffnen');
                    row.style.border = 'none';
                    row.style.background = 'rgba(50, 50, 93, 0.03)';
                    row.style.font = 'inherit';
                    row.style.textAlign = 'left';
                    row.style.cursor = 'pointer';
                    row.style.width = showBranchToggle ? 'auto' : '100%';
                } else {
                    row.style.background = 'rgba(50, 50, 93, 0.06)';
                }
                // Mit Zweig-Toggle: Zeile muss restliche Breite füllen, sonst kleben „Ordner“/Status-Pills am Namen.
                if (showBranchToggle) {
                    row.style.flex = '1';
                    row.style.minWidth = '0';
                } else if (!isSelectableRoot) {
                    row.style.width = '100%';
                }
                row.style.display = 'flex';
                row.style.alignItems = 'center';
                row.style.gap = '8px';
                row.style.padding = '10px 12px';
                row.style.paddingLeft = String(12 + (item.depth || 0) * 18) + 'px';
                row.style.borderRadius = '6px';

                const vIcon = document.createElement('i');
                vIcon.setAttribute('aria-hidden', 'true');
                vIcon.className = 'bi ' + treeIconForVirtualItem(item);
                vIcon.style.flexShrink = '0';
                vIcon.style.fontSize = '1.1em';
                vIcon.style.lineHeight = '1';
                vIcon.style.opacity = '0.92';
                vIcon.style.color = '#5e72e4';

                const name = document.createElement('div');
                name.style.minWidth = '0';
                name.style.flex = '1';
                name.style.fontWeight = '900';
                name.style.color = '#32325d';
                name.style.overflow = 'hidden';
                name.style.textOverflow = 'ellipsis';
                name.style.whiteSpace = 'nowrap';
                name.textContent = item.label;

                const meta2 = document.createElement('div');
                meta2.className = 'pill';
                meta2.title = 'Bereich';
                meta2.textContent = item.typLabel
                    ? item.typLabel
                    : item.typ === 'SchuelerInnen'
                      ? 'Schüler:innen'
                      : item.typ === 'LehrerInnen'
                        ? 'Lehrer:innen'
                        : 'Verwaltung';

                const meta = document.createElement('div');
                meta.className = 'pill';
                meta.title = 'Sync-Status';
                if (item.kind === 'root' && item.rootSyncStatus) {
                    meta.className = 'pill ' + pillClass(item.rootSyncStatus);
                    meta.textContent = String(item.rootSyncStatus);
                } else {
                    meta.textContent = '–';
                }
                if (mode === 'struktur' || mode === 'match') {
                    // Im Struktur-Baum nur Typ/Labels anzeigen – keine Status-Pills.
                    meta.style.display = 'none';
                }

                row.appendChild(vIcon);
                row.appendChild(name);
                row.appendChild(meta2);
                row.appendChild(meta);
                if (structureTreeRowShowsAddChildControl(mode, item) && typeof onStructureTreeAdd === 'function' && isSelectableRoot) {
                    const addB = document.createElement('button');
                    addB.type = 'button';
                    addB.className = 'ss-tree-add-btn';
                    addB.title = 'Unterpunkt hinzufügen';
                    addB.setAttribute('aria-label', 'Unterpunkt hinzufügen');
                    addB.textContent = '+';
                    addB.style.cssText =
                        'flex-shrink:0;align-self:center;width:38px;min-width:38px;height:38px;border:1px solid rgba(94,114,228,0.28);border-radius:10px;background:#fff;font-weight:1000;cursor:pointer;color:#32325d;line-height:1;padding:0;margin:0 4px 0 0;';
                    addB.addEventListener('click', (e) => {
                        e.preventDefault();
                        e.stopPropagation();
                        onStructureTreeAdd(String(item.rootId || ''), String(item.typ || ''), addB);
                    });
                    const wrap = document.createElement('div');
                    wrap.style.cssText = showBranchToggle
                        ? 'display:flex;flex:1;min-width:0;align-items:stretch;gap:0;'
                        : 'display:flex;width:100%;min-width:0;align-items:stretch;gap:0;';
                    row.style.flex = '1';
                    row.style.minWidth = '0';
                    if (isSelectableRoot) row.style.width = 'auto';
                    wrap.appendChild(row);
                    wrap.appendChild(addB);
                    li.appendChild(wrap);
                } else {
                    li.appendChild(row);
                }
                tree.appendChild(li);
                continue;
            }

            const { r, depth, hasKids } = item;
            const li = document.createElement('li');
            const showToggle = (mode === 'struktur' || mode === 'match') && hasKids && onToggleCollapse;
            if (showToggle) {
                li.style.display = 'flex';
                li.style.alignItems = 'stretch';
                const togg = document.createElement('button');
                togg.type = 'button';
                togg.setAttribute('aria-expanded', coll.has(String(r.id)) ? 'false' : 'true');
                togg.setAttribute('aria-label', coll.has(String(r.id)) ? 'Zweig aufklappen' : 'Zweig einklappen');
                togg.style.flexShrink = '0';
                togg.style.width = '32px';
                togg.style.alignSelf = 'center';
                togg.style.border = 'none';
                togg.style.background = 'transparent';
                togg.style.cursor = 'pointer';
                togg.style.color = '#32325d';
                togg.textContent = coll.has(String(r.id)) ? '▸' : '▾';
                togg.addEventListener('click', (e) => {
                    e.stopPropagation();
                    onToggleCollapse(String(r.id));
                });
                li.appendChild(togg);
            }

            const btn = document.createElement('button');
            btn.type = 'button';
            btn.dataset.ssSelect = String(r.id);
            btn.dataset.ssType = String(r.typ || '');
            btn.setAttribute('aria-current', selectedId && String(selectedId) === String(r.id) ? 'true' : 'false');
            btn.style.paddingLeft = String(12 + depth * 18) + 'px';
            btn.style.display = 'flex';
            btn.style.alignItems = 'center';
            btn.style.gap = '8px';
            if (showToggle) btn.style.flex = '1';
            if (mode === 'struktur' || mode === 'match') {
                btn.draggable = true;
            }

            const rowIcon = document.createElement('i');
            rowIcon.setAttribute('aria-hidden', 'true');
            rowIcon.className = 'bi ' + treeIconForRow(r, mode);
            rowIcon.style.flexShrink = '0';
            rowIcon.style.fontSize = '1.08em';
            rowIcon.style.lineHeight = '1';
            rowIcon.style.opacity = '0.92';
            rowIcon.style.color = '#5e72e4';

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
            meta.textContent =
                mode === 'tenant'
                    ? (r.hiddenMembership ? 'Kursteam' : String(r.typ || '–'))
                    : String(r.syncStatus || 'Ausstehend');
            if (mode === 'struktur' || mode === 'match') {
                // Im Struktur-Baum soll nicht der Status im Fokus stehen.
                meta.style.display = 'none';
            }

            const meta2 = document.createElement('div');
            meta2.className = 'pill';
            meta2.title = mode === 'tenant' ? 'Alias / E-Mail' : 'Typ / Schuljahr';
            meta2.textContent =
                mode === 'tenant'
                    ? String(r.alias || r.mail || '–')
                    : mode === 'struktur'
                      ? String(r.typ || '–')
                      : String(r.typ || '–') + (r.schuljahr ? ' · ' + String(r.schuljahr) : '');

            btn.appendChild(rowIcon);
            btn.appendChild(name);
            btn.appendChild(meta2);
            btn.appendChild(meta);
            if (structureTreeRowShowsAddChildControl(mode, item) && typeof onStructureTreeAdd === 'function') {
                const addB = document.createElement('button');
                addB.type = 'button';
                addB.className = 'ss-tree-add-btn';
                addB.title = 'Unterpunkt hinzufügen';
                addB.setAttribute('aria-label', 'Unterpunkt hinzufügen');
                addB.textContent = '+';
                addB.style.cssText =
                    'flex-shrink:0;align-self:center;width:38px;min-width:38px;height:38px;border:1px solid rgba(94,114,228,0.28);border-radius:10px;background:#fff;font-weight:1000;cursor:pointer;color:#32325d;line-height:1;padding:0;margin:0 4px 0 0;';
                addB.addEventListener('click', (e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    onStructureTreeAdd(String(r.id || ''), String(r.typ || ''), addB);
                });
                const wrap = document.createElement('div');
                wrap.style.cssText = 'display:flex;flex:1;min-width:0;align-items:center;gap:0;';
                btn.style.flex = '1';
                btn.style.minWidth = '0';
                btn.style.width = 'auto';
                wrap.appendChild(btn);
                wrap.appendChild(addB);
                li.appendChild(wrap);
            } else {
                li.appendChild(btn);
            }
            tree.appendChild(li);
        }
    }

    function ensureTreeContextMenu() {
        const existing = document.getElementById('ssCtxMenu');
        if (existing) return existing;
        const menu = document.createElement('div');
        menu.id = 'ssCtxMenu';
        menu.setAttribute('role', 'menu');
        menu.style.cssText =
            'position:fixed;z-index:10000;min-width:220px;max-width:min(340px,92vw);background:#fff;border:1px solid rgba(94,114,228,0.22);border-radius:12px;box-shadow:0 24px 70px rgba(0,0,0,0.18);padding:8px;display:none;';

        const head = document.createElement('div');
        head.id = 'ssCtxMenuTitle';
        head.style.cssText =
            'padding:8px 10px 6px;font-weight:1000;color:#32325d;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;';

        const sep = document.createElement('div');
        sep.style.cssText = 'height:1px;background:#eef0f3;margin:6px 0 8px;';

        const btnDelete = document.createElement('button');
        btnDelete.type = 'button';
        btnDelete.id = 'ssCtxDelete';
        btnDelete.setAttribute('role', 'menuitem');
        btnDelete.style.cssText =
            'width:100%;display:flex;align-items:center;gap:10px;padding:10px 12px;border:1px solid rgba(245,54,92,0.22);border-radius:12px;background:rgba(245,54,92,0.06);color:#b00020;font-weight:900;cursor:pointer;font:inherit;text-align:left;';
        btnDelete.innerHTML = '<i class="bi bi-trash" aria-hidden="true"></i><span>Löschen</span>';

        const hint = document.createElement('div');
        hint.id = 'ssCtxHint';
        hint.style.cssText = 'padding:8px 10px 4px;color:#6c757d;font-size:0.9em;line-height:1.35;display:none;';

        menu.appendChild(head);
        menu.appendChild(sep);
        menu.appendChild(btnDelete);
        menu.appendChild(hint);
        document.body.appendChild(menu);

        function hide() {
            menu.style.display = 'none';
            menu.dataset.targetId = '';
            menu.dataset.targetKind = '';
            head.textContent = '';
            hint.style.display = 'none';
            hint.textContent = '';
        }
        // @ts-ignore - lightweight helper
        menu.hide = hide;

        // close on outside click / escape / scroll
        window.addEventListener('click', () => hide(), true);
        window.addEventListener(
            'keydown',
            (ev) => {
                if (ev.key === 'Escape') hide();
            },
            true
        );
        window.addEventListener('scroll', () => hide(), true);

        return menu;
    }

    function fillParentSelect(rows, currentId, childTyp) {
        const sel = getEl('ssUebergeordnet');
        if (!sel) return;
        const prev = sel.value || '';
        const ct = normStr(childTyp);
        const opts = rows
            .filter((r) => r && r.id && String(r.id) !== String(currentId || ''))
            .filter((r) => (ct ? canReparentStrict(ct, String(r.typ || '')) : true))
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

    function isValidStructureParentChild(childTyp, parentId, rows) {
        const c = normStr(childTyp);
        if (!c) return false;
        const pid = normStr(parentId);
        if (!pid) {
            const root = inferRootForType(c);
            return canReparentStrict(c, root);
        }
        const pRow = rows.find((r) => String(r.id) === String(pid));
        if (!pRow) return false;
        return canReparentStrict(c, String(pRow.typ || ''));
    }

    function fillPersonKontaktFields(cur, typEffective, personInfoByRole) {
        const wrap = getEl('ssPersonKontaktWrap');
        const inpN = getEl('ssPersonName');
        const inpE = getEl('ssPersonEmail');
        if (!wrap || !inpN || !inpE) return;
        const t = String(typEffective || '');
        if (t === 'Person' && cur && !cur.isStructureTreeRoot) {
            wrap.style.display = '';
            const labelForRole = normStr(getEl('ssBezeichnung')?.value) || String(cur.bezeichnung || '');
            const fromRole =
                personInfoByRole && typeof personInfoByRole.get === 'function'
                    ? personInfoByRole.get(normRoleKey(labelForRole)) || {}
                    : {};
            const storedN = normStr(cur.personName);
            const storedE = normStr(cur.personEmail).toLowerCase();
            const active = typeof document !== 'undefined' ? document.activeElement : null;
            if (active !== inpN && active !== inpE) {
                inpN.value = storedN || normStr(fromRole.name) || '';
                inpE.value = storedE || normStr(fromRole.email).toLowerCase() || '';
            }
        } else {
            wrap.style.display = 'none';
            inpN.value = '';
            inpE.value = '';
        }
    }

    function refreshStrukturTypDependentUi(ctx) {
        const hintTyp = getEl('ssTypM365Hint');
        const hintOwn = getEl('ssStructOwnerTabHint');
        const hintMem = getEl('ssStructMemberTabHint');
        const besch = getEl('ssBeschreibung');
        const mode = ctx && ctx.mode;
        const selectedId = ctx && ctx.selectedId;
        const rowsStruktur = (ctx && ctx.rowsStruktur) || [];
        const structRootDetails = ctx && ctx.structRootDetails;
        if (mode !== 'struktur' || !selectedId) {
            if (hintTyp) hintTyp.textContent = '';
            const onTabs = ctx && ctx.onPersonDetailTabs;
            if (typeof onTabs === 'function') onTabs('');
            fillPersonKontaktFields(null, '', null);
            return;
        }
        const typFromDom = normStr(getEl('ssTyp')?.value);
        let cur = rowsStruktur.find((r) => String(r.id) === String(selectedId));
        if (!cur && isStructureTreeRootId(selectedId)) {
            cur = mergeStructureTreeRootRow(selectedId, structRootDetails);
        }
        const rowTyp = cur ? String(cur.typ || '') : '';
        const t = typFromDom || rowTyp;

        const typHints = {
            Jahrgang:
                'Organisatorische Jahrgangs‑Einheit unter „Schüler:innen“. Untergeordnet: Klassen. Kein eigenes M365‑Team/Gruppe vorgesehen.',
            Klasse:
                'Klasse unter einem Jahrgang. Darunter: Kursteams (M365‑Teams, oft HiddenMembership) und ggf. Gruppen.',
            Arbeitsgemeinschaft:
                'Lehrer:innen‑Fachgemeinschaft. In M365 oft eine eigene M365‑Gruppe oder ein Team; Owner/Mitglieder steuern Zugriff.',
            Kursteam:
                'Kurs‑Team: in M365 typischerweise ein Team mit HiddenMembership; Owner/Mitglieder entsprechen Lehrkräften und SuS im Kurs.',
            Gruppe:
                'M365‑Gruppe oder Team (SOLL). Beschreibung wird bei „Im Tenant anlegen“ als Graph‑description genutzt; Owner/Mitglieder für das Gruppenobjekt.',
            Person: ''
        };
        if (hintTyp) hintTyp.textContent = typHints[t] || '';

        const defOwn =
            'Owner für diese Einheit (lokal gespeichert). Wenn du oben rechts angemeldet bist, kannst du User bequem über Entra suchen.';
        const defMem =
            'Mitglieder für diese Einheit (lokal gespeichert). Wenn du oben rechts angemeldet bist, kannst du User bequem über Entra suchen.';
        if (hintOwn) hintOwn.textContent = t === 'Person' ? '' : defOwn;
        if (hintMem) hintMem.textContent = t === 'Person' ? '' : defMem;

        const ph = {
            Jahrgang: 'Optional: interner Hinweis zur Jahrgangs‑Einheit.',
            Klasse: 'Optional: interner Hinweis zur Klasse.',
            Arbeitsgemeinschaft: 'Optional: interner Hinweis zur ARGE.',
            Kursteam: 'Optional: interner Hinweis zum Kursteam.',
            Gruppe: 'Wird bei „Im Tenant anlegen“ als Beschreibung der Microsoft‑365‑Gruppe (Graph-Feld description) gesetzt.',
            Person: 'Optional: interne Notiz zur Rolle (kein Ersatz für Benutzerprofile in Entra ID).'
        };
        if (besch && cur && !cur.isStructureTreeRoot && ph[t]) {
            besch.setAttribute('placeholder', ph[t]);
        }

        const onTabs = ctx && ctx.onPersonDetailTabs;
        if (typeof onTabs === 'function') onTabs(t);
        const pir = ctx && ctx.personInfoByRole;
        fillPersonKontaktFields(cur, t, pir);
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
        fillParentSelect(rows, row.id, String(row.typ || ''));
        getEl('ssBezeichnung').value = String(row.bezeichnung || '');
        const besch = getEl('ssBeschreibung');
        if (besch) besch.value = String(row.beschreibung || '');
        getEl('ssTyp').value = String(row.typ || 'Gruppe');
        getEl('ssSchuljahr').value = String(row.schuljahr || '');
        getEl('ssStatus').value = String(row.status || 'Aktiv');
        getEl('ssUebergeordnet').value = String(row.parentId || '');
        getEl('ssSyncStatus').value = String(row.syncStatus || 'Ausstehend');
        getEl('ssLetzteFehlermeldung').value = String(row.letzteFehlermeldung || '');

        // Schema-spezifische Felder (Anlegen)
        const jgYear = getEl('ssJgYear');
        const jgSuffix = getEl('ssJgSuffix');
        const argeCode = getEl('ssArgeCode');
        const argeName = getEl('ssArgeName');
        const ktKlasse = getEl('ssKtKlasse');
        const ktFach = getEl('ssKtFach');
        const ktGruppe = getEl('ssKtGruppe');
        if (jgYear) jgYear.value = String(row.jgYear || '');
        if (jgSuffix) jgSuffix.value = String(row.jgSuffix || '');
        if (argeCode) argeCode.value = String(row.argeCode || '');
        if (argeName) argeName.value = String(row.argeName || '');
        if (ktKlasse) ktKlasse.value = String(row.ktKlasse || '');
        if (ktFach) ktFach.value = String(row.ktFach || '');
        if (ktGruppe) ktGruppe.value = String(row.ktGruppe || '');

        // Tenant-Create Meta (Anlegen)
        const target = getEl('ssStructTenantTarget');
        const vis = getEl('ssStructTenantVisibility');
        const nick = getEl('ssStructTenantMailNick');
        const created = getEl('ssStructTenantCreatedId');
        if (target) target.value = String(row.tenantTarget || '');
        if (vis) vis.value = String(row.tenantVisibility || '');
        if (nick) nick.value = String(row.tenantMailNickname || '');
        if (created) created.value = String(row.tenantGroupId || '');

        const ue = getEl('ssUebergeordnet');
        const typEl = getEl('ssTyp');
        if (row.isStructureTreeRoot) {
            if (ue) {
                ue.disabled = true;
                ue.value = '';
            }
            if (typEl) typEl.disabled = true;
        } else {
            if (ue) ue.disabled = false;
            if (typEl) typEl.disabled = false;
        }
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
        const expEl = getEl('ssTenantExpires');
        if (expEl) {
            const iso = group.expirationDateTime ? String(group.expirationDateTime).trim() : '';
            if (!iso) {
                expEl.value = '';
            } else {
                expEl.value = formatDateTimeAT(iso);
            }
        }
        getEl('ssTenantMail').value = String(group.mail || '');
        const visEl = getEl('ssTenantVisibility');
        if (visEl) {
            const v = String(group.visibility || '').trim();
            // Sichtbarkeit ist nur für Unified Gruppen relevant (Teams/M365-Gruppen)
            const can = group.typ === 'Team' || group.typ === 'Gruppe';
            visEl.disabled = !can;
            if (!can) {
                visEl.value = '';
            } else {
                visEl.value = v === 'Public' ? 'Public' : 'Private';
            }
        }
        getEl('ssTenantAlias').value = String(group.alias || '');
        const desc = getEl('ssTenantDescription');
        if (desc) desc.value = String(group.description || '');
        getEl('ssTenantId').value = String(group.id || '');

        const archWrap = getEl('ssTenantTeamArchiveWrap');
        const archSel = getEl('ssTenantArchiveState');
        const archHint = getEl('ssTenantArchiveHint');
        const archSpo = getEl('ssTenantArchiveSpoReadonly');
        if (archWrap && archSel && archHint && archSpo) {
            const typ = String(group.typ || '');
            const unifiedLike = typ === 'Team' || typ === 'Gruppe';
            archWrap.style.display = unifiedLike ? '' : 'none';
            if (!unifiedLike) {
                archSpo.checked = false;
            } else {
                const st = group.teamIsArchived;
                let cap = group.hasTeamsForArchive;
                if (cap === undefined && (st === true || st === false)) {
                    cap = true;
                }
                const loadingCap = cap === undefined && st === undefined;
                if (loadingCap) {
                    archSel.disabled = true;
                    archSel.value = 'active';
                    archHint.style.display = '';
                    archHint.textContent = 'Teams-Anbindung und Archiv-Status werden ermittelt …';
                    archSpo.disabled = true;
                    archSpo.checked = false;
                } else if (cap === false) {
                    archSel.disabled = true;
                    archSel.value = 'active';
                    archHint.style.display = '';
                    archHint.textContent =
                        'Kein Microsoft Teams an dieser Microsoft 365-Gruppe (nur Gruppe ohne Team/Kursteam) – Teams-Archivierung ist nicht verfügbar.';
                    archSpo.disabled = true;
                    archSpo.checked = false;
                } else if (cap === true && (st === true || st === false)) {
                    archSel.disabled = false;
                    archSel.value = st ? 'archived' : 'active';
                    archHint.style.display = 'none';
                    archHint.textContent = '';
                    archSpo.disabled = archSel.value !== 'archived';
                } else {
                    archSel.disabled = true;
                    archSel.value = 'active';
                    archHint.style.display = '';
                    archHint.textContent =
                        'Teams ist vorhanden, der Archiv-Status konnte nicht gelesen werden. Bitte „Neu laden“ oder Berechtigungen prüfen.';
                    archSpo.disabled = true;
                    archSpo.checked = false;
                }
            }
        }
    }

    function readDetailToRow(row) {
        const next = Object.assign({}, row);
        next.bezeichnung = normStr(getEl('ssBezeichnung')?.value);
        next.beschreibung = normStr(getEl('ssBeschreibung')?.value);
        next.typ = normStr(getEl('ssTyp')?.value) || 'Gruppe';
        next.schuljahr = normStr(getEl('ssSchuljahr')?.value);
        next.status = normStr(getEl('ssStatus')?.value) || 'Aktiv';
        next.parentId = normStr(getEl('ssUebergeordnet')?.value);
        next.syncStatus = normStr(getEl('ssSyncStatus')?.value) || 'Ausstehend';
        next.letzteFehlermeldung = normStr(getEl('ssLetzteFehlermeldung')?.value);
        if (next.typ === 'Person') {
            next.personName = normStr(getEl('ssPersonName')?.value);
            next.personEmail = normStr(getEl('ssPersonEmail')?.value).toLowerCase();
        } else {
            next.personName = '';
            next.personEmail = '';
        }

        // Schema-spezifische Meta (nur für Anlegen; optional)
        next.jgYear = normStr(getEl('ssJgYear')?.value);
        next.jgSuffix = normStr(getEl('ssJgSuffix')?.value);
        next.argeCode = normStr(getEl('ssArgeCode')?.value);
        next.argeName = normStr(getEl('ssArgeName')?.value);
        next.ktKlasse = normStr(getEl('ssKtKlasse')?.value);
        next.ktFach = normStr(getEl('ssKtFach')?.value);
        next.ktGruppe = normStr(getEl('ssKtGruppe')?.value);

        // Tenant-Create Meta (Anlegen)
        next.tenantTarget = normStr(getEl('ssStructTenantTarget')?.value);
        next.tenantVisibility = normStr(getEl('ssStructTenantVisibility')?.value);
        if (row.isStructureTreeRoot) {
            const canon = defaultStructureTreeRootRow(row.id);
            if (canon) {
                next.typ = canon.typ;
                next.parentId = '';
                next.id = String(row.id);
                next.isStructureTreeRoot = true;
            }
        }
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

    function downloadText(filename, text) {
        const blob = new Blob([String(text || '')], { type: 'text/plain;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        a.remove();
        setTimeout(() => URL.revokeObjectURL(url), 500);
    }

    function csvEscape(v) {
        const s = String(v ?? '');
        if (/[",\r\n]/.test(s)) return '"' + s.replaceAll('"', '""') + '"';
        return s;
    }

    function buildKursteamCsvRow(row, memberships, schemaState) {
        const sug = (() => {
            const typ = String(row?.typ || '');
            const displayName = String(row?.bezeichnung || '').trim();
            if (!displayName) return { displayName: '', mailNick: '' };
            if (typ === 'Kursteam') {
                const yearPrefix = String(schemaState.kursteamYearPrefix || '').trim();
                const st = typeof loadState === 'function' ? loadState() : { rows: [] };
                const kt = resolveKursteamKlasseFachForRow(row, st.rows);
                const klasse = kt.klasse;
                const fach = kt.fach;
                const gruppe = String(row.ktGruppe || '').trim();
                const mailNick =
                    klasse && fach
                        ? buildKursteamMailNickFromTemplate(schemaState.kursteamMailNickPattern, { yearPrefix, klasse, fach, gruppe })
                        : buildMailNickFromLabel(displayName);
                return { displayName, mailNick };
            }
            return { displayName, mailNick: buildMailNickFromLabel(displayName) };
        })();

        const mem = memberships[String(row.id)] || { owners: [], members: [] };
        const owners = (mem.owners || [])
            .map((p) => String(p.userPrincipalName || p.mail || '').trim())
            .filter(Boolean)
            .join(';');
        const members = (mem.members || [])
            .map((p) => String(p.userPrincipalName || p.mail || '').trim())
            .filter(Boolean)
            .join(';');

        const visibility = String(row.tenantVisibility || '').trim() || 'HiddenMembership';
        const target = String(row.tenantTarget || '').trim() || 'team';
        return {
            DisplayName: sug.displayName,
            MailNickname: sug.mailNick,
            Visibility: visibility,
            Target: target,
            Owners: owners,
            Members: members
        };
    }

    function buildKursteamCsv(rows, memberships, schemaState) {
        const header = ['DisplayName', 'MailNickname', 'Visibility', 'Target', 'Owners', 'Members'];
        const lines = [header.join(',')];
        for (const r of rows) {
            const o = buildKursteamCsvRow(r, memberships, schemaState);
            lines.push(
                [
                    csvEscape(o.DisplayName),
                    csvEscape(o.MailNickname),
                    csvEscape(o.Visibility),
                    csvEscape(o.Target),
                    csvEscape(o.Owners),
                    csvEscape(o.Members)
                ].join(',')
            );
        }
        return lines.join('\r\n');
    }

    function buildKursteamProvisionScript(csvFileName) {
        const file = String(csvFileName || 'kursteams.csv');
        const lines = [];
        lines.push('# Kursteams anlegen (CSV) – Microsoft Teams PowerShell');
        lines.push('# Voraussetzungen: Install-Module MicrosoftTeams -Scope CurrentUser');
        lines.push('$ErrorActionPreference = "Stop"');
        lines.push('');
        lines.push('Import-Module MicrosoftTeams -ErrorAction SilentlyContinue');
        lines.push('try { Connect-MicrosoftTeams | Out-Null } catch { throw }');
        lines.push('');
        lines.push(`$csv = Import-Csv -Path "${file}"`);
        lines.push('foreach ($r in $csv) {');
        lines.push('  if (-not $r.DisplayName -or -not $r.MailNickname) {');
        lines.push('    Write-Warning "Überspringe: DisplayName/MailNickname fehlt"');
        lines.push('    continue');
        lines.push('  }');
        lines.push('  $vis = if ($r.Visibility -eq "Public") { "Public" } else { "Private" }');
        lines.push('  Write-Host ("Lege an: " + $r.DisplayName + " (" + $r.MailNickname + ")") -ForegroundColor Cyan');
        lines.push('  $team = $null');
        lines.push('  if ($r.Target -eq "group") {');
        lines.push('    # Gruppen-Only ist in Teams PS nicht 1:1 abbildbar; wir legen ein Team an (Unified).');
        lines.push('  }');
        lines.push('  $team = New-Team -DisplayName $r.DisplayName -MailNickname $r.MailNickname -Visibility $vis');
        lines.push('  Start-Sleep -Seconds 2');
        lines.push('  $gid = $team.GroupId');
        lines.push('  if (-not $gid) { Write-Warning "Keine GroupId erhalten – weiter."; continue }');
        lines.push('');
        lines.push('  if ($r.Owners) {');
        lines.push('    $r.Owners.Split(";") | ForEach-Object {');
        lines.push('      $u = $_.Trim(); if ($u) {');
        lines.push('        try { Add-TeamUser -GroupId $gid -User $u -Role Owner | Out-Null } catch { Write-Warning ("Owner fehlgeschlagen: " + $u) }');
        lines.push('      }');
        lines.push('    }');
        lines.push('  }');
        lines.push('  if ($r.Members) {');
        lines.push('    $r.Members.Split(";") | ForEach-Object {');
        lines.push('      $u = $_.Trim(); if ($u) {');
        lines.push('        try { Add-TeamUser -GroupId $gid -User $u -Role Member | Out-Null } catch { Write-Warning ("Member fehlgeschlagen: " + $u) }');
        lines.push('      }');
        lines.push('    }');
        lines.push('  }');
        lines.push('}');
        lines.push('');
        lines.push('Disconnect-MicrosoftTeams | Out-Null');
        lines.push('Write-Host "Fertig." -ForegroundColor Green');
        return lines.join('\n');
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
        // Globaler Login (Header-Widget) – wenn vorhanden, nutzen wir ihn.
        if (typeof window.ms365AuthAcquireToken === 'function') {
            return await window.ms365AuthAcquireToken(scopes);
        }
        const instance = await getPca();
        let accounts = instance.getAllAccounts();
        if (!accounts.length) {
            // In eingebetteten Browsern (z.B. Cursor) bleibt ein Popup gelegentlich schwarz.
            // Redirect-Login ist deutlich robuster.
            try {
                // Nach der Anmeldung wieder zur aktuellen Tool-Seite zurückspringen.
                sessionStorage.setItem('ms365-post-login-url', window.location.href);
            } catch {
                // ignore
            }
            await instance.loginRedirect({ scopes, prompt: 'select_account', redirectStartPage: window.location.href });
            // loginRedirect navigiert weg; Code hier wird normalerweise nicht weiterlaufen.
            throw new Error('Weiterleitung zur Anmeldung …');
        }
        if (!accounts.length) throw new Error('Anmeldung abgebrochen.');
        const req = { scopes, account: accounts[0] };
        try {
            return (await instance.acquireTokenSilent(req)).accessToken;
        } catch (e) {
            if (isInteractionRequired(e)) {
                try {
                    sessionStorage.setItem('ms365-post-login-url', window.location.href);
                } catch {
                    // ignore
                }
                await instance.acquireTokenRedirect({ ...req, redirectStartPage: window.location.href });
                throw new Error('Weiterleitung zur Anmeldung …');
            }
            throw e;
        }
    }

    function sleep(ms) {
        return new Promise((r) => setTimeout(r, ms));
    }

    async function graphRequest(method, pathOrUrl, token, body, extraHeaders) {
        const url = pathOrUrl.indexOf('http') === 0 ? pathOrUrl : 'https://graph.microsoft.com/v1.0' + pathOrUrl;
        let attempt = 0;
        while (true) {
            const headers = { Authorization: 'Bearer ' + token };
            if (extraHeaders && typeof extraHeaders === 'object') {
                Object.assign(headers, extraHeaders);
            }
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
            const msg = typeof data === 'object' && data && data.error ? JSON.stringify(data.error) : text || String(res.status);
            throw new Error(method + ' ' + pathOrUrl + ': ' + msg);
        }
        return data || {};
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

    async function pollTeamsAsyncOperationForTenant(token, operationPath) {
        const maxAttempts = 90;
        for (let i = 0; i < maxAttempts; i++) {
            await sleep(2000);
            const data = await graphJson('GET', operationPath, token, undefined, undefined);
            const st = String(data.status || data.Status || '').toLowerCase();
            if (st === 'succeeded') return;
            if (st === 'failed') {
                const errMsg =
                    (data.error && (data.error.message || JSON.stringify(data.error))) || JSON.stringify(data);
                throw new Error('Teams-Operation fehlgeschlagen: ' + errMsg);
            }
        }
        throw new Error('Timeout: Teams-Archivierung nicht abgeschlossen.');
    }

    async function setTenantTeamArchiveState(teamId, archive, spoReadonlyForMembers) {
        const token = await getGraphToken(GRAPH_SCOPES_TENANT_TEAM_ARCHIVE);
        const path = '/teams/' + encodeURIComponent(teamId) + (archive ? '/archive' : '/unarchive');
        let body = undefined;
        if (archive && spoReadonlyForMembers) {
            body = { shouldSetSpoSiteReadOnlyForMembers: true };
        }
        const res = await graphRequest('POST', path, token, body, undefined);
        if (res.status !== 202 && res.status !== 200) {
            const t = await res.text();
            throw new Error('HTTP ' + res.status + ' ' + t);
        }
        const loc = res.headers.get('Location') || res.headers.get('Content-Location');
        const opPath = parseTeamsOperationPathFromLocation(loc);
        if (opPath) await pollTeamsAsyncOperationForTenant(token, opPath);
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
            if (typeof onProgress === 'function') {
                onProgress({ page, loaded: out.length, hasMore: !!next });
            }
        }
        return out;
    }

    function groupIsTeam(g) {
        const opts = g && g.resourceProvisioningOptions;
        return Array.isArray(opts) && opts.indexOf('Team') !== -1;
    }

    function graphErrorLooksLikeNotFound(err) {
        const s = String((err && err.message) || err || '');
        return /\b404\b/i.test(s) || /ResourceNotFound|ItemNotFound|not found|Request_ResourceNotFound/i.test(s);
    }

    /**
     * Teams-Archiv (Graph) gilt nur für Objekte mit Teams-Ressource (klassisches Team, Kursteam, …).
     * GET /teams/{id} liefert 404, wenn die Unified-Gruppe kein Team hat.
     * @returns {{ hasTeamsForArchive: boolean, teamIsArchived: boolean|null }}
     */
    async function resolveTeamsArchiveStateForUnifiedGroupId(groupId, token) {
        const gid = String(groupId || '').trim();
        if (!gid) return { hasTeamsForArchive: false, teamIsArchived: null };
        try {
            const team = await graphJson(
                'GET',
                '/teams/' + encodeURIComponent(gid) + '?$select=' + encodeURIComponent('id,isArchived'),
                token,
                undefined
            );
            return {
                hasTeamsForArchive: true,
                teamIsArchived: team.isArchived === true
            };
        } catch (e) {
            if (graphErrorLooksLikeNotFound(e)) {
                return { hasTeamsForArchive: false, teamIsArchived: null };
            }
            return { hasTeamsForArchive: true, teamIsArchived: null };
        }
    }

    async function loadTenantGroupsLive(kind, onProgress) {
        const token = await getGraphToken(GRAPH_SCOPES_TENANT_READ);
        const selectBase =
            'id,displayName,description,expirationDateTime,mail,mailNickname,createdDateTime,groupTypes,resourceProvisioningOptions,securityEnabled,mailEnabled,visibility';

        /** @type {any[]} */
        const all = [];
        const k = String(kind || 'm365');

        // 1) M365 (Unified) Gruppen/Teams
        if (k === 'm365' || k === 'both') {
            const filter = encodeURIComponent("groupTypes/any(c:c eq 'Unified')");
            const initial = '/groups?$filter=' + filter + '&$select=' + encodeURIComponent(selectBase) + '&$top=999';
            const groups = await fetchAllPages(token, initial, onProgress, undefined);
            for (const g of groups) all.push(g);
        }

        // 2) Sicherheitsgruppen (ohne Unified)
        if (k === 'security' || k === 'both') {
            // Advanced query: requires ConsistencyLevel:eventual when using "not(...)"
            const filter = encodeURIComponent("securityEnabled eq true and not(groupTypes/any(c:c eq 'Unified'))");
            const initial =
                '/groups?$count=true&$filter=' + filter + '&$select=' + encodeURIComponent(selectBase) + '&$top=999';
            const groups = await fetchAllPages(token, initial, onProgress, { ConsistencyLevel: 'eventual' });
            for (const g of groups) all.push(g);
        }

        const mapped = all
            .map((g) => {
                const isUnified = Array.isArray(g.groupTypes) && g.groupTypes.indexOf('Unified') !== -1;
                // HiddenMembership ist i.d.R. eine Visibility-Variante (visibility === 'HiddenMembership').
                // Manche Tenants liefern zusätzlich groupTypes-Einträge; wir unterstützen beides.
                const vis = String(g.visibility || '').trim();
                const hiddenMembership =
                    vis === 'HiddenMembership' ||
                    (Array.isArray(g.groupTypes) && g.groupTypes.indexOf('HiddenMembership') !== -1);
                const isTeam = isUnified && groupIsTeam(g);
                const isSecurity = !!g.securityEnabled && !isUnified;
                const typeLabel = isTeam
                    ? 'Team'
                    : isUnified
                      ? 'Gruppe'
                      : isSecurity
                        ? g.mailEnabled
                            ? 'E‑Mail‑Sicherheitsgruppe'
                            : 'Sicherheitsgruppe'
                        : 'Gruppe';
                return {
                    id: String(g.id || ''),
                    bezeichnung: String(g.displayName || ''),
                    typ: typeLabel,
                    mail: String(g.mail || ''),
                    alias: String(g.mailNickname || ''),
                    description: String(g.description || ''),
                    expirationDateTime: String(g.expirationDateTime || ''),
                    visibility: vis,
                    hiddenMembership: !!hiddenMembership,
                    createdDateTime: String(g.createdDateTime || '')
                };
            })
            .filter((x) => x.id);

        const seen = new Set();
        const unique = [];
        for (const r of mapped) {
            const id = String(r.id);
            if (seen.has(id)) continue;
            seen.add(id);
            unique.push(r);
        }

        if (typeof onProgress === 'function') {
            onProgress({ phase: 'counts', page: 1, loaded: 0, hasMore: true, total: unique.length });
        }
        const counted = await enrichTenantRowsOwnerMemberCounts(unique, token);
        if (typeof onProgress === 'function') {
            onProgress({ phase: 'counts', page: 1, loaded: unique.length, hasMore: false, total: unique.length });
        }

        counted.sort((a, b) => compareDe(a.bezeichnung, b.bezeichnung));
        const prevUsers = loadTenantCache().users || [];
        saveTenantCache(counted, prevUsers);
        return counted;
    }

    async function loadTenantUsersLive(onProgress) {
        const token = await getGraphToken(GRAPH_SCOPES_TENANT_INVENTORY);
        const select = encodeURIComponent('id,displayName,userPrincipalName,mail,accountEnabled');
        const initial = '/users?$select=' + select + '&$top=999';
        const raw = [];
        let next = initial;
        let page = 0;
        while (next && page < 50 && raw.length < 8000) {
            page++;
            const data = await graphJson('GET', next, token, undefined, undefined);
            const vals = data.value;
            if (Array.isArray(vals)) for (let i = 0; i < vals.length; i++) raw.push(vals[i]);
            next = data['@odata.nextLink'] || null;
            if (typeof onProgress === 'function') {
                onProgress({ phase: 'users', page, loaded: raw.length, hasMore: !!next });
            }
        }
        const mapped = raw
            .map((u) => ({
                id: String(u.id || ''),
                displayName: String(u.displayName || '').trim(),
                userPrincipalName: String(u.userPrincipalName || '').trim().toLowerCase(),
                mail: String(u.mail || '').trim().toLowerCase(),
                accountEnabled: u.accountEnabled !== false
            }))
            .filter((x) => x.id);
        mapped.sort((a, b) => compareDe(a.displayName || a.userPrincipalName, b.displayName || b.userPrincipalName));
        return mapped;
    }

    async function loadTenantInventoryFull(onProgress) {
        const kind = 'both';
        await loadTenantGroupsLive(kind, (p) => {
            if (typeof onProgress !== 'function') return;
            if (p && p.phase === 'counts') onProgress(p);
            else onProgress(Object.assign({ phase: 'groups' }, p));
        });
        const afterG = loadTenantCache();
        const users = await loadTenantUsersLive((p) => {
            if (typeof onProgress === 'function') onProgress(p);
        });
        saveTenantCache(afterG.rows, users);
        return {
            groups: afterG.rows,
            users,
            loadedAt: new Date().toISOString()
        };
    }

    window.ms365TenantInventory = {
        refresh: (onProgress) => loadTenantInventoryFull(onProgress || function () {}),
        readCache: () => loadTenantCache(),
        loadStructureState: () => {
            const st = loadState();
            // Für Schritt 4 (Matching) sollen auch die drei virtuellen Hauptbereiche verknüpfbar sein.
            const settings = st.settings || {};
            const structRootDetails = settings && settings.structRootDetails ? settings.structRootDetails : {};
            const roots = [
                mergeStructureTreeRootRow(STRUCT_TREE_ROOT_STUDENTS, structRootDetails),
                mergeStructureTreeRootRow(STRUCT_TREE_ROOT_TEACHERS, structRootDetails),
                mergeStructureTreeRootRow(STRUCT_TREE_ROOT_ADMIN, structRootDetails)
            ].filter(Boolean);
            return { rows: (st.rows || []).concat(roots), memberships: st.memberships, settings: st.settings };
        },
        loadMatchLinks: () => loadMatchState().links || {},
        suggestGroupForUnit: (unit) => suggestTenantGroupForUnitFromList(unit, loadTenantCache().rows || []),
        saveMatchLink: (structureId, tenantGroupId, note, tenantUserId) =>
            saveMatchLinkPublic(structureId, tenantGroupId, note, tenantUserId),
        computeCreateSuggestion: (row) => {
            const st = loadState();
            const schemaState = Object.assign({}, defaultAnlegenSchemas(), st.settings || {});
            return computeTenantCreateSuggestionFromRow(row, schemaState);
        },
        patchStructureRow: (rowId, patch) => {
            const rid = String(rowId || '');
            if (isStructureTreeRootId(rid)) {
                const st = loadState();
                const settings = Object.assign({}, st.settings || {});
                if (!settings.structRootDetails || typeof settings.structRootDetails !== 'object') settings.structRootDetails = {};
                const cur = mergeStructureTreeRootRow(rid, settings.structRootDetails) || defaultStructureTreeRootRow(rid);
                if (!cur) return false;
                const next = Object.assign({}, cur, patch || {});
                settings.structRootDetails[rid] = pickStorableStructureTreeRootFields(next);
                saveState({ rows: st.rows, memberships: st.memberships, settings });
                try {
                    window.dispatchEvent(new CustomEvent('ms365-structure-changed', { detail: {} }));
                } catch {
                    // ignore
                }
                return true;
            }
            return patchStructureRowById(rid, patch);
        },
        provisionGroupRow: (row) => graphProvisionStructureGroupRow(row),
        provisionPersonRow: (row, opts) => graphProvisionPersonRowPublic(row, opts)
    };


    async function fetchTenantGroupDetail(groupId) {
        const token = await getGraphToken(GRAPH_SCOPES_TENANT_READ);
        const sel =
            'id,displayName,description,expirationDateTime,mail,mailNickname,groupTypes,resourceProvisioningOptions,securityEnabled,mailEnabled,visibility';
        const g = await graphJson('GET', '/groups/' + encodeURIComponent(groupId) + '?$select=' + encodeURIComponent(sel), token, undefined);
        const isUnified = Array.isArray(g.groupTypes) && g.groupTypes.indexOf('Unified') !== -1;
        const vis = String(g.visibility || '').trim();
        const hiddenMembership =
            vis === 'HiddenMembership' ||
            (Array.isArray(g.groupTypes) && g.groupTypes.indexOf('HiddenMembership') !== -1);
        const isTeam = isUnified && groupIsTeam(g);
        const isSecurity = !!g.securityEnabled && !isUnified;
        const typ = isTeam ? 'Team' : isUnified ? 'Gruppe' : isSecurity ? (g.mailEnabled ? 'E‑Mail‑Sicherheitsgruppe' : 'Sicherheitsgruppe') : 'Gruppe';
        /** @type {boolean|null} */
        let teamIsArchived = null;
        /** @type {boolean} */
        let hasTeamsForArchive = false;
        if (isUnified) {
            const ar = await resolveTeamsArchiveStateForUnifiedGroupId(String(g.id || ''), token);
            hasTeamsForArchive = ar.hasTeamsForArchive;
            teamIsArchived = ar.teamIsArchived;
        }
        return {
            id: String(g.id || ''),
            bezeichnung: String(g.displayName || ''),
            typ,
            mail: String(g.mail || ''),
            alias: String(g.mailNickname || ''),
            description: String(g.description || ''),
            expirationDateTime: String(g.expirationDateTime || ''),
            visibility: vis,
            hiddenMembership: !!hiddenMembership,
            teamIsArchived,
            hasTeamsForArchive
        };
    }

    function personLabel(p) {
        if (!p || typeof p !== 'object') return '';
        const dn = p.displayName ? String(p.displayName).trim() : '';
        const upn = p.userPrincipalName || p.mail ? String(p.userPrincipalName || p.mail).trim() : '';
        if (dn && upn && dn !== upn) return dn + ' (' + upn + ')';
        return dn || upn || (p.id ? String(p.id) : '');
    }

    function odataEscape(s) {
        return String(s).replace(/'/g, "''");
    }

    async function graphSearchUsersForOwner(token, query) {
        const q = String(query || '').trim();
        if (!q) return [];
        const esc = odataEscape(q);
        let filter;
        if (q.indexOf('@') !== -1) {
            filter = "(mail eq '" + esc + "' or userPrincipalName eq '" + esc + "')";
        } else {
            filter =
                "(startswith(displayName,'" +
                esc +
                "') or startswith(userPrincipalName,'" +
                esc +
                "') or startswith(mail,'" +
                esc +
                "'))";
        }
        const select = 'id,displayName,mail,userPrincipalName';
        const path =
            '/users?$filter=' +
            encodeURIComponent(filter) +
            '&$select=' +
            encodeURIComponent(select) +
            '&$top=25';
        const data = await graphJson('GET', path, token, undefined);
        return data.value || [];
    }

    async function fetchAllPagesSimple(token, initialPath) {
        const out = [];
        let next = initialPath;
        let pages = 0;
        while (next && pages < 40 && out.length < 4000) {
            pages++;
            const data = await graphJson('GET', next, token, undefined, undefined);
            const vals = data.value;
            if (Array.isArray(vals)) for (let i = 0; i < vals.length; i++) out.push(vals[i]);
            next = data['@odata.nextLink'] || null;
        }
        return out;
    }

    async function fetchGroupOwners(groupId) {
        const token = await getGraphToken(GRAPH_SCOPES_TENANT_OWNER_MANAGE);
        const select = 'id,displayName,mail,userPrincipalName';
        const path =
            '/groups/' +
            encodeURIComponent(groupId) +
            '/owners?$select=' +
            encodeURIComponent(select) +
            '&$top=200';
        const owners = await fetchAllPagesSimple(token, path);
        owners.sort((a, b) => compareDe(personLabel(a), personLabel(b)));
        return owners;
    }

    async function graphGetCollectionCount(token, groupId, segment) {
        const gid = String(groupId || '').trim();
        if (!gid) return -1;
        const seg = segment === 'owners' ? 'owners' : 'members';
        const path = '/groups/' + encodeURIComponent(gid) + '/' + seg + '/$count';
        const res = await graphRequest('GET', path, token, undefined, { ConsistencyLevel: 'eventual' });
        const text = await res.text();
        if (!res.ok) return -1;
        const n = parseInt(String(text).trim(), 10);
        return isNaN(n) ? -1 : n;
    }

    async function fetchGroupMemberCount(groupId) {
        const token = await getGraphToken(GRAPH_SCOPES_TENANT_OWNER_MANAGE);
        return graphGetCollectionCount(token, groupId, 'members');
    }

    async function mapWithConcurrencyLimited(items, limit, fn) {
        const results = new Array(items.length);
        let i = 0;
        async function worker() {
            while (true) {
                const idx = i++;
                if (idx >= items.length) return;
                results[idx] = await fn(items[idx], idx);
            }
        }
        const n = Math.max(1, Math.min(limit, items.length || 1));
        const workers = [];
        for (let w = 0; w < n; w++) workers.push(worker());
        await Promise.all(workers);
        return results;
    }

    /** Nach Gruppenliste: Owner-/Mitglieder-$count für Filter „ohne …“. */
    async function enrichTenantRowsOwnerMemberCounts(rows, token) {
        return await mapWithConcurrencyLimited(rows, 10, async (row) => {
            const id = String(row && row.id ? row.id : '').trim();
            if (!id) return Object.assign({}, row, { ownerCount: -1, memberCount: -1 });
            const oc = await graphGetCollectionCount(token, id, 'owners');
            const mc = await graphGetCollectionCount(token, id, 'members');
            return Object.assign({}, row, { ownerCount: oc, memberCount: mc });
        });
    }

    async function fetchGroupMembers(groupId) {
        const token = await getGraphToken(GRAPH_SCOPES_TENANT_OWNER_MANAGE);
        const select = 'id,displayName,mail,userPrincipalName';
        let next =
            '/groups/' +
            encodeURIComponent(groupId) +
            '/members?$select=' +
            encodeURIComponent(select) +
            '&$top=200';
        const out = [];
        let pages = 0;
        while (next && pages < 40 && out.length < 2000) {
            pages++;
            const data = await graphJson('GET', next, token, undefined, undefined);
            const vals = data.value || [];
            for (let i = 0; i < vals.length; i++) out.push(vals[i]);
            if (out.length >= 2000) break;
            next = data['@odata.nextLink'] || null;
        }
        out.sort((a, b) => compareDe(personLabel(a), personLabel(b)));
        return { items: out, truncated: !!next || out.length >= 2000 };
    }

    function directoryObjectRef(id) {
        return 'https://graph.microsoft.com/v1.0/directoryObjects/' + id;
    }

    async function addGroupOwner(groupId, userId) {
        const token = await getGraphToken(GRAPH_SCOPES_TENANT_OWNER_MANAGE);
        const body = { '@odata.id': directoryObjectRef(userId) };
        await graphJson('POST', '/groups/' + encodeURIComponent(groupId) + '/owners/$ref', token, body);
    }

    async function addGroupMember(groupId, userId) {
        const token = await getGraphToken(GRAPH_SCOPES_TENANT_OWNER_MANAGE);
        const body = { '@odata.id': directoryObjectRef(userId) };
        await graphJson('POST', '/groups/' + encodeURIComponent(groupId) + '/members/$ref', token, body, undefined);
    }

    async function removeGroupOwner(groupId, ownerId) {
        const token = await getGraphToken(GRAPH_SCOPES_TENANT_OWNER_MANAGE);
        await graphJson(
            'DELETE',
            '/groups/' + encodeURIComponent(groupId) + '/owners/' + encodeURIComponent(ownerId) + '/$ref',
            token,
            undefined
        );
    }

    async function removeGroupMember(groupId, memberId) {
        const token = await getGraphToken(GRAPH_SCOPES_TENANT_OWNER_MANAGE);
        await graphJson(
            'DELETE',
            '/groups/' + encodeURIComponent(groupId) + '/members/' + encodeURIComponent(memberId) + '/$ref',
            token,
            undefined,
            undefined
        );
    }

    function isGraphDuplicateRefError(err) {
        const msg = String(err && err.message ? err.message : err);
        return /already exist/i.test(msg) || /already exists/i.test(msg);
    }

    async function addOwnerWithMemberFallback(groupId, userId) {
        try {
            await addGroupOwner(groupId, userId);
        } catch (e1) {
            // Manche Tenants/Policies verlangen, dass Owner auch Member ist.
            try {
                await addGroupMember(groupId, userId);
            } catch (e2) {
                if (!isGraphDuplicateRefError(e2)) throw e2;
            }
            await addGroupOwner(groupId, userId);
        }
    }

    async function updateTenantGroup(groupId, displayName, description, visibility) {
        const token = await getGraphToken(GRAPH_SCOPES_TENANT_WRITE);
        const body = {
            displayName: String(displayName || '').trim(),
            description: String(description || '').trim()
        };
        if (!body.displayName) throw new Error('Bitte einen Anzeigenamen eingeben.');
        const vis = String(visibility || '').trim();
        if (vis === 'Private' || vis === 'Public') {
            body.visibility = vis;
        }
        await graphJson('PATCH', '/groups/' + encodeURIComponent(groupId), token, body, undefined);
    }

    async function renewTenantGroup(groupId) {
        const token = await getGraphToken(GRAPH_SCOPES_TENANT_WRITE);
        // Renew extends expiration based on lifecycle policy (Graph action).
        await graphJson('POST', '/groups/' + encodeURIComponent(groupId) + '/renew', token, undefined, undefined);
    }

    async function deleteTenantGroup(groupId) {
        const token = await getGraphToken(GRAPH_SCOPES_TENANT_WRITE);
        await graphJson('DELETE', '/groups/' + encodeURIComponent(groupId), token, undefined, undefined);
    }

    async function createUnifiedGroup(displayName, description, mailNickname, visibility) {
        const token = await getGraphToken(GRAPH_SCOPES_TENANT_WRITE);
        const body = {
            displayName: String(displayName || '').trim(),
            description: String(description || '').trim(),
            mailEnabled: true,
            mailNickname: String(mailNickname || '').trim(),
            securityEnabled: false,
            groupTypes: ['Unified']
        };
        const vis = String(visibility || '').trim();
        if (vis === 'Private' || vis === 'Public' || vis === 'HiddenMembership') body.visibility = vis;
        if (!body.displayName) throw new Error('Bitte einen Anzeigenamen eingeben.');
        if (!body.mailNickname) throw new Error('Mail‑Nickname fehlt (Vorschlag ist leer).');
        return await graphJson('POST', '/groups', token, body, undefined);
    }

    async function createMailEnabledSecurityGroup(displayName, description, mailNickname) {
        const token = await getGraphToken(GRAPH_SCOPES_TENANT_WRITE);
        const body = {
            displayName: String(displayName || '').trim(),
            description: String(description || '').trim(),
            mailEnabled: true,
            mailNickname: String(mailNickname || '').trim(),
            securityEnabled: true,
            groupTypes: []
        };
        if (!body.displayName) throw new Error('Bitte einen Anzeigenamen eingeben.');
        if (!body.mailNickname) throw new Error('Mail‑Nickname fehlt (Vorschlag ist leer).');
        return await graphJson('POST', '/groups', token, body, undefined);
    }

    async function createTeamForGroup(groupId) {
        const token = await getGraphToken(GRAPH_SCOPES_TENANT_WRITE);
        await graphJson('PUT', '/groups/' + encodeURIComponent(groupId) + '/team', token, {}, undefined);
    }

    function bind() {
        const isEmbedStructure =
            typeof document !== 'undefined' &&
            document.body &&
            String(document.body.getAttribute('data-ss-embed-structure') || '') === 'true';
        const state = loadState();
        let rowsStruktur = state.rows.slice();
        /** @type {Record<string, { owners: any[], members: any[] }>} */
        let memberships = state.memberships || {};
        const schemaState = Object.assign({}, defaultAnlegenSchemas(), state.settings || {});
        normalizeGraphLayoutModeInSettings(schemaState);
        /** @type {{links: Record<string, { tenantGroupId: string, note: string, updatedAt: string }>}} */
        const matchState = loadMatchState();
        /** @type {Record<string, { tenantGroupId: string, note: string, updatedAt: string }>} */
        let links = matchState.links || {};
        const tenantCache = loadTenantCache();
        let rowsTenant = tenantCache.rows.slice();
        let selectedId = '';
        /** @type {Set<string>} */
        let tenantMultiSel = new Set();
        /** @type {'struktur'|'tenant'|'match'} */
        let mode = 'struktur';

        // Default-Verwaltung beim Start sicherstellen (auch bei komplett leeren Daten)
        (function ensureDefaultVerwaltungOnInit() {
            const schuljahr = currentSchoolYearLabel();
            const hasTop = rowsStruktur.some((r) => r && r.typ === 'Gruppe' && String(r.bezeichnung || '').trim().toLowerCase() === 'verwaltung');
            if (hasTop) return;
            const vId = uid();
            rowsStruktur.push({ id: vId, parentId: '', typ: 'Gruppe', bezeichnung: 'Verwaltung', schuljahr, status: 'Aktiv', syncStatus: 'Ausstehend', letzteFehlermeldung: '' });
            // Direktion/Administration sind in der Regel einzelne Personen; Sekretariat oft als Team/Gruppe
            rowsStruktur.push({ id: uid(), parentId: vId, typ: 'Person', bezeichnung: 'Direktion', schuljahr, status: 'Aktiv', syncStatus: 'Ausstehend', letzteFehlermeldung: '' });
            rowsStruktur.push({ id: uid(), parentId: vId, typ: 'Person', bezeichnung: 'Administration', schuljahr, status: 'Aktiv', syncStatus: 'Ausstehend', letzteFehlermeldung: '' });
            rowsStruktur.push({ id: uid(), parentId: vId, typ: 'Gruppe', bezeichnung: 'Sekretariat', schuljahr, status: 'Aktiv', syncStatus: 'Ausstehend', letzteFehlermeldung: '' });
            saveState({ rows: rowsStruktur, memberships, settings: schemaState });
        })();

        /**
         * Verwaltungszweig an Schul‑Einstellungen (admin: Rolle;Name;E‑Mail) angleichen:
         * fehlende Rollen als Person anlegen, Kontaktdaten aus den Einstellungen übernehmen.
         */
        function syncVerwaltungStructureFromTenantSettings(settings, schuljahr) {
            if (!settings || typeof settings !== 'object') return false;
            const admin = Array.isArray(settings.admin) ? settings.admin : [];
            if (!admin.length) return false;
            const sj = String(schuljahr || '').trim() || currentSchoolYearLabel();
            let changed = false;
            let verw = rowsStruktur.find(
                (r) =>
                    r &&
                    r.typ === 'Gruppe' &&
                    String(r.bezeichnung || '').trim().toLowerCase() === 'verwaltung' &&
                    !String(r.parentId || '').trim()
            );
            if (!verw) {
                verw = {
                    id: uid(),
                    parentId: '',
                    typ: 'Gruppe',
                    bezeichnung: 'Verwaltung',
                    schuljahr: sj,
                    status: 'Aktiv',
                    syncStatus: 'Ausstehend',
                    letzteFehlermeldung: ''
                };
                rowsStruktur.push(verw);
                changed = true;
            }
            const verwId = String(verw.id);
            const byRole = new Map();
            for (let i = 0; i < rowsStruktur.length; i++) {
                const r = rowsStruktur[i];
                if (!r || String(r.parentId || '') !== verwId) continue;
                const key = normRoleKey(r.bezeichnung);
                if (key && !byRole.has(key)) byRole.set(key, r);
            }
            for (let j = 0; j < admin.length; j++) {
                const a = admin[j];
                const role = normStr(a && (a.role || a.rolle || a.title));
                if (!role) continue;
                const rk = normRoleKey(role);
                const name = normStr(a && a.name);
                const email = normStr(a && a.email).toLowerCase();
                const existing = byRole.get(rk);
                if (existing) {
                    const nextN = name || '';
                    const nextE = email || '';
                    if (normStr(existing.personName) !== nextN || normStr(existing.personEmail).toLowerCase() !== nextE) {
                        existing.personName = nextN;
                        existing.personEmail = nextE;
                        changed = true;
                    }
                    if (String(existing.typ || '') === 'Person' && normStr(existing.bezeichnung) !== role) {
                        existing.bezeichnung = role;
                        changed = true;
                    }
                    continue;
                }
                rowsStruktur.push({
                    id: uid(),
                    parentId: verwId,
                    typ: 'Person',
                    bezeichnung: role,
                    personName: name,
                    personEmail: email,
                    schuljahr: sj,
                    status: 'Aktiv',
                    syncStatus: 'Ausstehend',
                    letzteFehlermeldung: ''
                });
                byRole.set(rk, rowsStruktur[rowsStruktur.length - 1]);
                changed = true;
            }
            return changed;
        }

        /** @type {{dragId: string, dragType: string}|null} */
        let graphDrag = null;
        let graphWrapDnDBound = false;
        let graphNodesToggleBound = false;
        /** @type {{x:number,y:number,scale:number}} */
        let graphViewport = { x: 0, y: 0, scale: 1 };
        let graphPan = null;
        /** @type {Set<string>} */
        let graphCollapsed = loadGraphCollapsedSet();
        // In der Ersteinrichtung ist die Erwartung: alles sichtbar.
        // Persistierte Collapsed-States aus anderen Seiten sind hier verwirrend.
        try {
            if (isEmbedStructure && /ersteinrichtung\.html$/i.test(String(window.location?.pathname || ''))) {
                graphCollapsed = new Set();
                saveGraphCollapsedSet(graphCollapsed);
            }
        } catch {
            // ignore
        }
        /** @type {Map<string, {name:string,email:string}>} */
        let personInfoByRole = new Map();

        function refreshPersonInfoFromTenantSettings(settings) {
            const next = new Map();
            const s = settings && typeof settings === 'object' ? settings : null;
            const admin = s && Array.isArray(s.admin) ? s.admin : [];
            admin.forEach((a) => {
                const role = normRoleKey(a?.role || a?.rolle || a?.title);
                if (!role) return;
                next.set(role, { name: normStr(a?.name), email: normStr(a?.email).toLowerCase() });
            });
            personInfoByRole = next;
        }

        // Initial: wenn wir in `tenant.html` laufen, Tenant-Settings sind vorhanden
        try {
            if (typeof window.ms365TenantSettingsLoad === 'function') {
                const s = window.ms365TenantSettingsLoad();
                refreshPersonInfoFromTenantSettings(s);
                if (syncVerwaltungStructureFromTenantSettings(s, currentSchoolYearLabel())) {
                    saveState({ rows: rowsStruktur, memberships, settings: schemaState });
                }
            }
        } catch {
            // ignore
        }

        // Graph/Create modal (Grundkonfiguration)
        const modal = getEl('ssStructCreateModal');
        const modalTitle = getEl('ssStructCreateTitle');
        const modalHint = getEl('ssStructCreateHint');
        const modalType = getEl('ssStructCreateType');
        const modalName = getEl('ssStructCreateName');
        const modalYear = getEl('ssStructCreateSchoolYear');
        const modalOk = getEl('ssStructCreateOk');
        const modalCancel = getEl('ssStructCreateCancel');
        const modalClose = getEl('ssStructCreateClose');
        const schemaWrap = getEl('ssStructCreateSchemaWrap');
        const wrapJg = getEl('ssStructCreateSchemaJg');
        const wrapArge = getEl('ssStructCreateSchemaArge');
        const wrapKt = getEl('ssStructCreateSchemaKursteam');
        const inpJgYear = getEl('ssStructCreateJgYear');
        const inpJgSuffix = getEl('ssStructCreateJgSuffix');
        const inpArgeCode = getEl('ssStructCreateArgeCode');
        const inpArgeName = getEl('ssStructCreateArgeName');
        const inpKtKlasse = getEl('ssStructCreateKtKlasse');
        const inpKtFach = getEl('ssStructCreateKtFach');
        const inpKtGruppe = getEl('ssStructCreateKtGruppe');
        let modalParent = { id: '', typ: '' };

        function allowedChildTypes(parentType) {
            return allowedStructureChildTypes(String(parentType || ''));
        }

        function updateCreateSchemaUi() {
            if (!schemaWrap || !wrapJg || !wrapArge || !wrapKt || !modalType) return;
            const t = String(modalType.value || '');
            const show = t === 'Jahrgang' || t === 'Arbeitsgemeinschaft' || t === 'Kursteam';
            schemaWrap.style.display = show ? '' : 'none';
            wrapJg.style.display = t === 'Jahrgang' ? '' : 'none';
            wrapArge.style.display = t === 'Arbeitsgemeinschaft' ? '' : 'none';
            wrapKt.style.display = t === 'Kursteam' ? '' : 'none';
        }

        function openCreateModal(parentId, parentType, preferredChildTyp) {
            if (!modal || !modalType || !modalName || !modalYear) return;
            modalParent = { id: String(parentId || ''), typ: String(parentType || '') };
            const opts = allowedChildTypes(modalParent.typ);
            modalType.replaceChildren();
            for (const t of opts) {
                const o = document.createElement('option');
                o.value = t;
                o.textContent = t === 'Arbeitsgemeinschaft' ? 'ARGE (Arbeitsgemeinschaft)' : t;
                modalType.appendChild(o);
            }
            const pref = String(preferredChildTyp || '').trim();
            modalType.value = opts.indexOf(pref) !== -1 ? pref : opts[0] || '';

            // defaults
            const parentRow = rowsStruktur.find((r) => String(r.id) === String(modalParent.id));
            modalYear.value = parentRow && parentRow.schuljahr ? String(parentRow.schuljahr) : currentSchoolYearLabel();
            modalName.value = '';
            if (inpJgYear) inpJgYear.value = '';
            if (inpJgSuffix) inpJgSuffix.value = '';
            if (inpArgeCode) inpArgeCode.value = '';
            if (inpArgeName) inpArgeName.value = '';
            if (inpKtKlasse) inpKtKlasse.value = parentRow && parentRow.typ === 'Klasse' ? String(parentRow.bezeichnung || '') : '';
            if (inpKtFach) inpKtFach.value = '';
            if (inpKtGruppe) inpKtGruppe.value = '';

            if (modalTitle) modalTitle.textContent = 'Neues Element hinzufügen';
            if (modalHint) {
                const pLab = modalParent.typ === 'SchuelerInnen' ? 'Schüler:innen' : modalParent.typ === 'LehrerInnen' ? 'Lehrer:innen' : modalParent.typ;
                modalHint.textContent =
                    opts.length > 1
                        ? `Unter „${pLab}“ anlegen. Typ: im Feld „Typ“ wählen (oder über das +-Menü im Baum/Organigramm).`
                        : `Unter „${pLab}“ anlegen.`;
            }

            updateCreateSchemaUi();
            closeStructureAddTypePicker();
            modal.classList.add('active');
            setTimeout(() => modalName.focus(), 0);
        }

        let structAddPickerOutsideDown = null;
        let structAddPickerEscDown = null;
        function closeStructureAddTypePicker() {
            const el = getEl('ssStructAddTypePicker');
            if (el) el.style.display = 'none';
            if (structAddPickerOutsideDown) {
                document.removeEventListener('mousedown', structAddPickerOutsideDown, true);
                structAddPickerOutsideDown = null;
            }
            if (structAddPickerEscDown) {
                document.removeEventListener('keydown', structAddPickerEscDown, true);
                structAddPickerEscDown = null;
            }
        }

        function openStructureAddTypePicker(opts, anchorEl, onPick) {
            let el = getEl('ssStructAddTypePicker');
            if (!el) {
                el = document.createElement('div');
                el.id = 'ssStructAddTypePicker';
                el.setAttribute('role', 'dialog');
                el.setAttribute('aria-label', 'Typ wählen');
                el.style.cssText = [
                    'display:none',
                    'position:fixed',
                    'z-index:100000',
                    'min-width:200px',
                    'max-width:min(92vw,340px)',
                    'padding:12px 12px 8px',
                    'border-radius:12px',
                    'border:1px solid rgba(94,114,228,0.35)',
                    'background:#fff',
                    'box-shadow:0 18px 40px rgba(50,50,93,0.22)'
                ].join(';');
                const tit = document.createElement('div');
                tit.textContent = 'Unterpunkt-Typ wählen';
                tit.style.cssText = 'font-weight:1000;font-size:0.92em;color:#32325d;margin:0 0 8px;';
                const row = document.createElement('div');
                row.id = 'ssStructAddTypePickerRow';
                row.style.cssText = 'display:flex;flex-wrap:wrap;gap:6px;align-items:center;';
                el.appendChild(tit);
                el.appendChild(row);
                document.body.appendChild(el);
            }
            const row = getEl('ssStructAddTypePickerRow');
            if (!row) return;
            row.replaceChildren();
            for (const typ of opts) {
                const lab = typ === 'Arbeitsgemeinschaft' ? 'ARGE' : typ;
                const b = document.createElement('button');
                b.type = 'button';
                b.className = 'btn small-btn';
                b.style.margin = '0';
                b.textContent = lab;
                b.addEventListener('click', (e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    closeStructureAddTypePicker();
                    onPick(typ);
                });
                row.appendChild(b);
            }
            const r =
                anchorEl && typeof anchorEl.getBoundingClientRect === 'function'
                    ? anchorEl.getBoundingClientRect()
                    : { left: Math.max(12, window.innerWidth / 2 - 120), top: 96, bottom: 120, right: window.innerWidth / 2 + 120 };
            el.style.display = 'block';
            const ew = el.offsetWidth || 220;
            const eh = el.offsetHeight || 80;
            let left = Math.round(r.left);
            let top = Math.round(r.bottom + 6);
            if (left + ew > window.innerWidth - 8) left = Math.max(8, window.innerWidth - ew - 8);
            if (top + eh > window.innerHeight - 8) top = Math.max(8, Math.round(r.top - eh - 6));
            el.style.left = left + 'px';
            el.style.top = top + 'px';

            structAddPickerOutsideDown = (ev) => {
                if (!el || el.style.display === 'none') return;
                const t = ev && ev.target;
                if (el.contains(t)) return;
                if (anchorEl && anchorEl.contains && anchorEl.contains(t)) return;
                closeStructureAddTypePicker();
            };
            structAddPickerEscDown = (ev) => {
                if (ev && ev.key === 'Escape') closeStructureAddTypePicker();
            };
            setTimeout(() => {
                document.addEventListener('mousedown', structAddPickerOutsideDown, true);
                document.addEventListener('keydown', structAddPickerEscDown, true);
            }, 0);
        }

        function openStructureAddPicker(parentId, parentType, anchorEl) {
            const opts = allowedChildTypes(String(parentType || ''));
            if (!opts.length) return;
            if (opts.length === 1) {
                openCreateModal(parentId, parentType, opts[0]);
                return;
            }
            openStructureAddTypePicker(opts, anchorEl, (typ) => openCreateModal(parentId, parentType, typ));
        }

        function closeCreateModal() {
            if (!modal) return;
            closeStructureAddTypePicker();
            modal.classList.remove('active');
        }

        async function commitCreateModal() {
            if (!modalType || !modalName || !modalYear) return;
            const typ = String(modalType.value || '').trim();
            if (!typ) return;
            const bezeichnung = normStr(modalName.value);
            if (!bezeichnung) {
                await dlgAlert('Bitte eine Bezeichnung eingeben.', { title: 'Eingabe' });
                return;
            }
            const schuljahr = normStr(modalYear.value) || currentSchoolYearLabel();

            // strict rule guard (virtual root handled here)
            if (!canReparentStrict(typ, modalParent.typ)) {
                await dlgAlert('Nicht erlaubt: ' + typ + ' kann nicht unter ' + modalParent.typ + ' angelegt werden.', { title: 'Struktur' });
                return;
            }

            const r = {
                id: uid(),
                parentId: (modalParent.typ === 'SchuelerInnen' || modalParent.typ === 'LehrerInnen') ? '' : String(modalParent.id || ''),
                typ,
                bezeichnung,
                beschreibung: '',
                schuljahr,
                status: 'Aktiv',
                syncStatus: 'Ausstehend',
                letzteFehlermeldung: '',
                jgYear: '',
                jgSuffix: '',
                argeCode: '',
                argeName: '',
                ktKlasse: '',
                ktFach: '',
                ktGruppe: '',
                personName: '',
                personEmail: ''
            };

            if (typ === 'Jahrgang') {
                r.jgYear = normStr(inpJgYear?.value);
                r.jgSuffix = normStr(inpJgSuffix?.value);
            } else if (typ === 'Arbeitsgemeinschaft') {
                r.argeCode = normStr(inpArgeCode?.value);
                r.argeName = normStr(inpArgeName?.value);
            } else if (typ === 'Kursteam') {
                r.ktKlasse = normStr(inpKtKlasse?.value);
                r.ktFach = normStr(inpKtFach?.value);
                r.ktGruppe = normStr(inpKtGruppe?.value);
            }

            rowsStruktur.push(r);
            saveState({ rows: rowsStruktur, memberships, settings: schemaState });
            emitWritebackFromStructure();
            closeCreateModal();
            if (!(await confirmMatchLeaveIfNeeded(String(r.id)))) return;
            select(r.id);
        }

        if (modalType) modalType.addEventListener('change', () => updateCreateSchemaUi());
        if (modalOk) modalOk.addEventListener('click', () => void commitCreateModal());
        if (modalCancel) modalCancel.addEventListener('click', () => closeCreateModal());
        if (modalClose) modalClose.addEventListener('click', () => closeCreateModal());
        if (modal) {
            modal.addEventListener('click', (ev) => {
                if (ev.target === modal) closeCreateModal();
            });
            window.addEventListener('keydown', (ev) => {
                if (ev.key === 'Escape' && modal.classList.contains('active')) closeCreateModal();
            });
        }

        // Graph/Edit modal (Details per Pop-Up)
        const editModal = getEl('ssStructEditModal');
        const editTitle = getEl('ssStructEditTitle');
        const editHint = getEl('ssStructEditHint');
        const editName = getEl('ssStructEditName');
        const editType = getEl('ssStructEditType');
        const editYear = getEl('ssStructEditSchoolYear');
        const editStatus = getEl('ssStructEditStatus');
        const editParent = getEl('ssStructEditParent');
        const editSyncStatus = getEl('ssStructEditSyncStatus');
        const editLastError = getEl('ssStructEditLastError');
        const editBeschreibung = getEl('ssStructEditBeschreibung');
        const editSave = getEl('ssStructEditSave');
        const editDelete = getEl('ssStructEditDelete');
        const editCancel = getEl('ssStructEditCancel');
        const editClose = getEl('ssStructEditClose');
        const editSchemaWrap = getEl('ssStructEditSchemaWrap');
        const editWrapJg = getEl('ssStructEditSchemaJg');
        const editWrapArge = getEl('ssStructEditSchemaArge');
        const editWrapKt = getEl('ssStructEditSchemaKursteam');
        const editJgYear = getEl('ssStructEditJgYear');
        const editJgSuffix = getEl('ssStructEditJgSuffix');
        const editArgeCode = getEl('ssStructEditArgeCode');
        const editArgeName = getEl('ssStructEditArgeName');
        const editKtKlasse = getEl('ssStructEditKtKlasse');
        const editKtFach = getEl('ssStructEditKtFach');
        const editKtGruppe = getEl('ssStructEditKtGruppe');
        let editId = '';
        /** Bearbeiten-Modal für virtuelle Baum-Wurzeln (__root_students__ / …), nicht für echte `rowsStruktur`-Zeilen. */
        let editModalIsStructureTreeRoot = false;

        // Listen <-> Struktur Sync
        let __tenantToStructureGuard = 0;

        function normalizeTenantClassCode(c) {
            const code = c && c.code ? String(c.code).trim().toUpperCase() : '';
            const name = c && c.name ? String(c.name).trim() : '';
            return code || name || '';
        }

        function gradeLabelFromGraduationYear(gradYear, schuljahr) {
            const gy = String(gradYear || '').trim();
            const sy = String(schuljahr || '').trim().slice(0, 4);
            const gyi = /^\d{4}$/.test(gy) ? parseInt(gy, 10) : NaN;
            const syi = /^\d{4}$/.test(sy) ? parseInt(sy, 10) : NaN;
            if (!isFinite(gyi) || !isFinite(syi)) return '';
            const grade = gradeFromGraduationYear(gy, String(schuljahr || ''), schemaState.maxSchulstufen || 5);
            if (!isFinite(grade) || grade < 1 || grade > 5) return '';
            return String(Math.round(grade));
        }

        function ensureJahrgangForClass(classCode, graduationYear, schuljahr) {
            const gy = String(graduationYear || '').trim();
            if (!gy) return '';
            const grade = gradeLabelFromGraduationYear(gy, schuljahr);
            const labelByGrade = ('Jahrgang ' + (grade || '')).trim();
            const existing = rowsStruktur.find((r) => r && r.typ === 'Jahrgang' && String(r.jgYear || '').trim() === gy);
            if (existing) return String(existing.id);
            // Wenn ein Jahrgang bereits manuell existiert (ohne jgYear), nutze ihn statt einen zweiten anzulegen.
            const existingByLabel =
                grade && labelByGrade
                    ? rowsStruktur.find(
                          (r) =>
                              r &&
                              r.typ === 'Jahrgang' &&
                              !String(r.jgYear || '').trim() &&
                              String(r.bezeichnung || '').trim() === labelByGrade
                      )
                    : null;
            if (existingByLabel) {
                existingByLabel.jgYear = gy;
                return String(existingByLabel.id);
            }
            const jg = {
                id: uid(),
                parentId: '',
                typ: 'Jahrgang',
                bezeichnung: labelByGrade,
                schuljahr: schuljahr || currentSchoolYearLabel(),
                status: 'Aktiv',
                syncStatus: 'Ausstehend',
                letzteFehlermeldung: '',
                jgYear: gy,
                jgSuffix: ''
            };
            rowsStruktur.push(jg);
            return String(jg.id);
        }

        function ensureStructureFromTenantSettings(settings, reason) {
            if (!settings || typeof settings !== 'object') return;
            if (__tenantToStructureGuard) return;
            __tenantToStructureGuard++;
            try {
                const schuljahr = currentSchoolYearLabel();
                let structureChanged = syncVerwaltungStructureFromTenantSettings(settings, schuljahr);

                const classes = Array.isArray(settings.classes) ? settings.classes : [];
                if (!classes.length) {
                    if (structureChanged) {
                        saveState({ rows: rowsStruktur, memberships, settings: schemaState });
                        if (reason === 'render' || reason === 'manual-save' || reason === 'autosave') {
                            rerender();
                        }
                    }
                    return;
                }

                // Default-Verwaltung sicherstellen
                (function ensureDefaultVerwaltung() {
                    const hasTop = rowsStruktur.some((r) => r && r.typ === 'Gruppe' && String(r.bezeichnung || '').trim().toLowerCase() === 'verwaltung');
                    if (hasTop) return;
                    const vId = uid();
                    rowsStruktur.push({ id: vId, parentId: '', typ: 'Gruppe', bezeichnung: 'Verwaltung', schuljahr, status: 'Aktiv', syncStatus: 'Ausstehend', letzteFehlermeldung: '' });
                    rowsStruktur.push({ id: uid(), parentId: vId, typ: 'Person', bezeichnung: 'Direktion', schuljahr, status: 'Aktiv', syncStatus: 'Ausstehend', letzteFehlermeldung: '' });
                    rowsStruktur.push({ id: uid(), parentId: vId, typ: 'Person', bezeichnung: 'Administration', schuljahr, status: 'Aktiv', syncStatus: 'Ausstehend', letzteFehlermeldung: '' });
                    rowsStruktur.push({ id: uid(), parentId: vId, typ: 'Gruppe', bezeichnung: 'Sekretariat', schuljahr, status: 'Aktiv', syncStatus: 'Ausstehend', letzteFehlermeldung: '' });
                })();

                // ARGEs/Arbeitsgemeinschaften aus Listen sicherstellen (Top-Level unter Lehrer:innen)
                (function ensureArges() {
                    const arges = Array.isArray(settings.arges) ? settings.arges : [];
                    if (!arges.length) return;
                    const existingByCode = new Map(
                        rowsStruktur
                            .filter((r) => r && r.typ === 'Arbeitsgemeinschaft')
                            .map((r) => [String(r.argeCode || r.bezeichnung || '').trim().toUpperCase(), r])
                            .filter((x) => x[0])
                    );
                    arges.forEach((a) => {
                        const code = String(a.code || '').trim().toUpperCase();
                        if (!code) return;
                        const name = String(a.name || '').trim();
                        const ex = existingByCode.get(code);
                        if (ex) {
                            if (!ex.argeCode) ex.argeCode = code;
                            if (name && !ex.argeName) ex.argeName = name;
                            if (!ex.bezeichnung) ex.bezeichnung = name || code;
                            return;
                        }
                        rowsStruktur.push({
                            id: uid(),
                            parentId: '',
                            typ: 'Arbeitsgemeinschaft',
                            bezeichnung: name || code,
                            schuljahr,
                            status: 'Aktiv',
                            syncStatus: 'Ausstehend',
                            letzteFehlermeldung: '',
                            argeCode: code,
                            argeName: name
                        });
                    });
                })();

                const existingClasses = rowsStruktur.filter((r) => r && r.typ === 'Klasse');
                const byCode = new Map(existingClasses.map((c) => [String(c.bezeichnung || '').trim().toUpperCase(), c]));

                classes.forEach((c) => {
                    const label = normalizeTenantClassCode(c);
                    if (!label) return;
                    const parentId = ensureJahrgangForClass(label, c.year, schuljahr);
                    const ex = byCode.get(label.toUpperCase());
                    if (ex) {
                        if (!ex.parentId && parentId) ex.parentId = parentId;
                        if (!ex.schuljahr) ex.schuljahr = schuljahr;
                        if (c.year && !ex.classGradYear) ex.classGradYear = String(c.year || '').trim();
                        return;
                    }
                    rowsStruktur.push({
                        id: uid(),
                        parentId: parentId || '',
                        typ: 'Klasse',
                        bezeichnung: label,
                        classGradYear: String(c.year || '').trim(),
                        schuljahr,
                        status: 'Aktiv',
                        syncStatus: 'Ausstehend',
                        letzteFehlermeldung: ''
                    });
                });

                // Optional: wenn Domain gesetzt ist, Schema-Domain aktualisieren
                if (settings.domain && typeof settings.domain === 'string' && settings.domain.trim()) {
                    schemaState.domain = String(settings.domain).trim();
                }

                saveState({ rows: rowsStruktur, memberships, settings: schemaState });
                if (reason === 'render' || reason === 'manual-save' || reason === 'autosave') {
                    rerender();
                }
            } finally {
                __tenantToStructureGuard--;
            }
        }

        function emitWritebackFromStructure() {
            try {
                const classCodes = rowsStruktur
                    .filter((r) => r && r.typ === 'Klasse')
                    .map((r) => String(r.bezeichnung || '').trim())
                    .filter(Boolean);
                const subjectCodes = rowsStruktur
                    .filter((r) => r && r.typ === 'Kursteam')
                    .map((r) => String(r.ktFach || '').trim())
                    .filter(Boolean);
                const argeCodes = rowsStruktur
                    .filter((r) => r && r.typ === 'Arbeitsgemeinschaft')
                    .map((r) => String(r.argeCode || r.bezeichnung || '').trim())
                    .filter(Boolean);
                window.dispatchEvent(
                    new CustomEvent('ms365-structure-changed', {
                        detail: {
                            writeback: { classCodes, subjectCodes, argeCodes }
                        }
                    })
                );
            } catch {
                // ignore
            }
        }

        try {
            window.addEventListener('ms365-tenant-settings-changed', (ev) => {
                const d = ev && ev.detail ? ev.detail : null;
                if (!d || !d.settings) return;
                refreshPersonInfoFromTenantSettings(d.settings);
                ensureStructureFromTenantSettings(d.settings, String(d.reason || ''));
                rerender();
            });
        } catch {
            // ignore
        }

        function openEditModalForId(id) {
            if (!editModal || !editName || !editType || !editYear || !editStatus || !editParent) return;
            if (editParent) editParent.disabled = false;
            editModalIsStructureTreeRoot = false;

            if (isStructureTreeRootId(id)) {
                const rootRow = mergeStructureTreeRootRow(id, schemaState.structRootDetails);
                if (!rootRow) return;
                editModalIsStructureTreeRoot = true;
                editId = String(id);
                select(String(id));
                if (editTitle) editTitle.textContent = 'Details: ' + (rootRow.bezeichnung || '(Hauptbereich)');
                if (editHint) {
                    editHint.textContent =
                        'Virtueller Hauptbereich (lokal). Bezeichnung/Beschreibung u. a. gelten u. a. für „Im Tenant anlegen“ und die Baum-Anzeige. Übergeordnet ist nicht wählbar.';
                }

                editName.value = String(rootRow.bezeichnung || '');
                editType.value = String(rootRow.typ || '');
                editYear.value = String(rootRow.schuljahr || currentSchoolYearLabel());
                if (editStatus) editStatus.value = String(rootRow.status || 'Aktiv');
                if (editSyncStatus) editSyncStatus.value = String(rootRow.syncStatus || 'Ausstehend');
                if (editLastError) editLastError.value = String(rootRow.letzteFehlermeldung || '');
                if (editBeschreibung) editBeschreibung.value = String(rootRow.beschreibung || '');

                const editPersonWrap = getEl('ssStructEditPersonKontaktWrap');
                const editPersonName = getEl('ssStructEditPersonName');
                const editPersonEmail = getEl('ssStructEditPersonEmail');
                if (editPersonWrap) editPersonWrap.style.display = 'none';
                if (editPersonName) editPersonName.value = '';
                if (editPersonEmail) editPersonEmail.value = '';

                if (editSchemaWrap && editWrapJg && editWrapArge && editWrapKt) {
                    editSchemaWrap.style.display = 'none';
                    editWrapJg.style.display = 'none';
                    editWrapArge.style.display = 'none';
                    editWrapKt.style.display = 'none';
                }

                editParent.replaceChildren();
                const optTop = document.createElement('option');
                optTop.value = '';
                optTop.textContent = '(oberste Ebene)';
                editParent.appendChild(optTop);
                editParent.value = '';
                editParent.disabled = true;

                editModal.classList.add('active');
                setStructEditTab('allg');
                applyPersonStructEditModalTabs(String(rootRow.typ || ''));
                renderStructEditMembershipUi();
                setTimeout(() => editName.focus(), 0);
                return;
            }

            const row = rowsStruktur.find((r) => String(r.id) === String(id));
            if (!row) return;
            editId = String(row.id);
            if (editTitle) editTitle.textContent = 'Details: ' + (row.bezeichnung || '(ohne Bezeichnung)');
            if (editHint) editHint.textContent = 'Änderungen gelten für die SOLL‑Struktur (lokal). Drag&Drop funktioniert alternativ direkt im Organigramm.';

            editName.value = String(row.bezeichnung || '');
            editType.value = String(row.typ || '');
            editYear.value = String(row.schuljahr || currentSchoolYearLabel());
            if (editStatus) editStatus.value = String(row.status || 'Aktiv');
            if (editSyncStatus) editSyncStatus.value = String(row.syncStatus || 'Ausstehend');
            if (editLastError) editLastError.value = String(row.letzteFehlermeldung || '');
            if (editBeschreibung) editBeschreibung.value = String(row.beschreibung || '');

            // schema fields
            const t = String(row.typ || '');
            const editPersonWrap = getEl('ssStructEditPersonKontaktWrap');
            const editPersonName = getEl('ssStructEditPersonName');
            const editPersonEmail = getEl('ssStructEditPersonEmail');
            const fromRolePerson =
                personInfoByRole && typeof personInfoByRole.get === 'function'
                    ? personInfoByRole.get(normRoleKey(String(row.bezeichnung || ''))) || {}
                    : {};
            if (editPersonWrap && editPersonName && editPersonEmail) {
                if (t === 'Person') {
                    editPersonWrap.style.display = '';
                    editPersonName.value = normStr(row.personName) || normStr(fromRolePerson.name) || '';
                    editPersonEmail.value =
                        normStr(row.personEmail).toLowerCase() || normStr(fromRolePerson.email).toLowerCase() || '';
                } else {
                    editPersonWrap.style.display = 'none';
                    editPersonName.value = '';
                    editPersonEmail.value = '';
                }
            }
            if (editSchemaWrap && editWrapJg && editWrapArge && editWrapKt) {
                const show = t === 'Jahrgang' || t === 'Arbeitsgemeinschaft' || t === 'Kursteam';
                editSchemaWrap.style.display = show ? '' : 'none';
                editWrapJg.style.display = t === 'Jahrgang' ? '' : 'none';
                editWrapArge.style.display = t === 'Arbeitsgemeinschaft' ? '' : 'none';
                editWrapKt.style.display = t === 'Kursteam' ? '' : 'none';
            }
            if (editJgYear) editJgYear.value = String(row.jgYear || '');
            if (editJgSuffix) editJgSuffix.value = String(row.jgSuffix || '');
            if (editArgeCode) editArgeCode.value = String(row.argeCode || '');
            if (editArgeName) editArgeName.value = String(row.argeName || '');
            if (editKtKlasse) editKtKlasse.value = String(row.ktKlasse || '');
            if (editKtFach) editKtFach.value = String(row.ktFach || '');
            if (editKtGruppe) editKtGruppe.value = String(row.ktGruppe || '');

            // parent options (strict)
            editParent.replaceChildren();
            const type = String(row.typ || '');
            const root = inferRootForType(type);
            const optRoot = document.createElement('option');
            optRoot.value = root === 'LehrerInnen' ? '__root_teachers__' : '__root_students__';
            optRoot.textContent = root === 'LehrerInnen' ? '(Lehrer:innen)' : '(Schüler:innen)';

            // Determine allowed parent types
            const candidates = [];
            if (type === 'Klasse') {
                rowsStruktur.filter((x) => x && x.typ === 'Jahrgang').forEach((x) => candidates.push(x));
            } else if (type === 'Kursteam' || type === 'Gruppe') {
                rowsStruktur.filter((x) => x && x.typ === 'Klasse').forEach((x) => candidates.push(x));
                if (type === 'Gruppe') {
                    rowsStruktur.filter((x) => x && x.typ === 'Gruppe' && String(x.id) !== String(row.id)).forEach((x) => candidates.push(x));
                }
            } else if (type === 'Person') {
                rowsStruktur.filter((x) => x && x.typ === 'Gruppe').forEach((x) => candidates.push(x));
            } else if (type === 'Jahrgang' || type === 'Arbeitsgemeinschaft') {
                // root only
            }

            // root allowed?
            if (canReparentStrict(type, root)) editParent.appendChild(optRoot);
            candidates
                .sort((a, b) => compareDe(String(a.bezeichnung || ''), String(b.bezeichnung || '')))
                .forEach((p) => {
                    const ok = canReparentStrict(type, String(p.typ || ''));
                    if (!ok) return;
                    const o = document.createElement('option');
                    o.value = String(p.id);
                    o.textContent = String(p.bezeichnung || '(ohne Bezeichnung)');
                    editParent.appendChild(o);
                });

            // current parent
            const currentPid = String(row.parentId || '');
            if (!currentPid) {
                editParent.value = optRoot.value;
            } else {
                editParent.value = currentPid;
            }

            editModal.classList.add('active');
            setStructEditTab('allg');
            applyPersonStructEditModalTabs(String(row.typ || ''));
            renderStructEditMembershipUi();
            setTimeout(() => editName.focus(), 0);
        }

        function closeEditModal() {
            if (!editModal) return;
            editModal.classList.remove('active');
            editId = '';
            editModalIsStructureTreeRoot = false;
            if (editParent) editParent.disabled = false;
        }

        // Details-Pop-Up Tabs (Allgemein | Owner | Mitglieder)
        const editTabAllgBtn = getEl('ssStructEditTabAllgBtn');
        const editTabOwnerBtn = getEl('ssStructEditTabOwnerBtn');
        const editTabMitglBtn = getEl('ssStructEditTabMitglBtn');
        const editTabAllg = getEl('ssStructEditTabAllg');
        const editTabOwner = getEl('ssStructEditTabOwner');
        const editTabMitgl = getEl('ssStructEditTabMitgl');

        function setStructEditTab(key) {
            const k = String(key || 'allg');
            const isAllg = k === 'allg';
            const isOwner = k === 'own';
            const isMitgl = k === 'mem';
            if (editTabAllgBtn) editTabAllgBtn.classList.toggle('active', isAllg);
            if (editTabOwnerBtn) editTabOwnerBtn.classList.toggle('active', isOwner);
            if (editTabMitglBtn) editTabMitglBtn.classList.toggle('active', isMitgl);
            if (editTabAllgBtn) editTabAllgBtn.setAttribute('aria-selected', isAllg ? 'true' : 'false');
            if (editTabOwnerBtn) editTabOwnerBtn.setAttribute('aria-selected', isOwner ? 'true' : 'false');
            if (editTabMitglBtn) editTabMitglBtn.setAttribute('aria-selected', isMitgl ? 'true' : 'false');
            if (editTabAllg) editTabAllg.classList.toggle('active', isAllg);
            if (editTabOwner) editTabOwner.classList.toggle('active', isOwner);
            if (editTabMitgl) editTabMitgl.classList.toggle('active', isMitgl);
        }

        if (editTabAllgBtn) editTabAllgBtn.addEventListener('click', () => setStructEditTab('allg'));
        if (editTabOwnerBtn) editTabOwnerBtn.addEventListener('click', () => setStructEditTab('own'));
        if (editTabMitglBtn) editTabMitglBtn.addEventListener('click', () => setStructEditTab('mem'));

        function applyPersonStructEditModalTabs(typ) {
            const isPerson = String(typ || '') === 'Person';
            ['ssStructEditTabOwnerBtn', 'ssStructEditTabMitglBtn'].forEach((id) => {
                const el = getEl(id);
                if (!el) return;
                el.style.display = isPerson ? 'none' : '';
                el.disabled = isPerson;
                el.setAttribute('aria-hidden', isPerson ? 'true' : 'false');
                if (isPerson) el.classList.remove('active');
            });
            ['ssStructEditTabOwner', 'ssStructEditTabMitgl'].forEach((id) => {
                const el = getEl(id);
                if (!el) return;
                el.style.display = isPerson ? 'none' : '';
                el.setAttribute('aria-hidden', isPerson ? 'true' : 'false');
                if (isPerson) el.classList.remove('active');
            });
            if (isPerson) setStructEditTab('allg');
        }

        function renderStructEditMembershipUi() {
            const id = String(editId || '');
            if (!id) return;
            const m = getStructMembership(id);
            renderStructPeople(m.owners, 'ssStructEditOwnersList', 'data-ss-struct-edit-remove-owner');
            renderStructPeople(m.members, 'ssStructEditMembersList', 'data-ss-struct-edit-remove-member');
        }

        // Owner/Mitglieder im Details-Pop-Up
        getEl('ssStructEditOwnerSearchBtn')?.addEventListener('click', async () => {
            await runStructUserSearch(getEl('ssStructEditOwnerSearch')?.value || '', 'ssStructEditOwnerSearchResults');
        });
        getEl('ssStructEditMemberSearchBtn')?.addEventListener('click', async () => {
            await runStructUserSearch(getEl('ssStructEditMemberSearch')?.value || '', 'ssStructEditMemberSearchResults');
        });
        getEl('ssStructEditOwnerAddBtn')?.addEventListener('click', () => {
            if (!editId) return;
            const sel = getEl('ssStructEditOwnerSearchResults');
            const raw = sel ? sel.value : '';
            const user = safeJsonParse(raw);
            if (!user) return;
            const m = getStructMembership(editId);
            if (!m.owners.some((x) => String(x.id) === String(user.id))) m.owners.push(user);
            setStructMembership(editId, m);
            renderStructEditMembershipUi();
        });
        getEl('ssStructEditMemberAddBtn')?.addEventListener('click', () => {
            if (!editId) return;
            const sel = getEl('ssStructEditMemberSearchResults');
            const raw = sel ? sel.value : '';
            const user = safeJsonParse(raw);
            if (!user) return;
            const m = getStructMembership(editId);
            if (!m.members.some((x) => String(x.id) === String(user.id))) m.members.push(user);
            setStructMembership(editId, m);
            renderStructEditMembershipUi();
        });
        getEl('ssStructEditOwnersList')?.addEventListener('click', (ev) => {
            if (!editId) return;
            const t = ev && ev.target ? ev.target : null;
            const btn = t && t.closest ? t.closest('[data-ss-struct-edit-remove-owner]') : null;
            if (!btn) return;
            const uid = btn.getAttribute('data-ss-struct-edit-remove-owner') || '';
            const m = getStructMembership(editId);
            m.owners = (m.owners || []).filter((x) => String(x.id) !== String(uid));
            setStructMembership(editId, m);
            renderStructEditMembershipUi();
        });
        getEl('ssStructEditMembersList')?.addEventListener('click', (ev) => {
            if (!editId) return;
            const t = ev && ev.target ? ev.target : null;
            const btn = t && t.closest ? t.closest('[data-ss-struct-edit-remove-member]') : null;
            if (!btn) return;
            const uid = btn.getAttribute('data-ss-struct-edit-remove-member') || '';
            const m = getStructMembership(editId);
            m.members = (m.members || []).filter((x) => String(x.id) !== String(uid));
            setStructMembership(editId, m);
            renderStructEditMembershipUi();
        });

        async function saveEditModal() {
            if (!editId) return;
            if (editModalIsStructureTreeRoot && isStructureTreeRootId(editId)) {
                const nextName = normStr(editName?.value);
                if (!nextName) {
                    await dlgAlert('Bitte eine Bezeichnung eingeben.', { title: 'Eingabe' });
                    return;
                }
                if (!schemaState.structRootDetails || typeof schemaState.structRootDetails !== 'object') {
                    schemaState.structRootDetails = {};
                }
                const cur = mergeStructureTreeRootRow(editId, schemaState.structRootDetails) || defaultStructureTreeRootRow(editId);
                if (!cur) return;
                cur.bezeichnung = nextName;
                if (editBeschreibung) cur.beschreibung = normStr(editBeschreibung.value);
                cur.schuljahr = normStr(editYear?.value) || currentSchoolYearLabel();
                cur.status = normStr(editStatus?.value) || 'Aktiv';
                cur.syncStatus = normStr(editSyncStatus?.value) || 'Ausstehend';
                cur.letzteFehlermeldung = normStr(editLastError?.value);
                schemaState.structRootDetails[String(editId)] = pickStorableStructureTreeRootFields(cur);
                saveState({ rows: rowsStruktur, memberships, settings: schemaState });
                selectedId = String(editId);
                rerender();
                closeEditModal();
                return;
            }

            const row = rowsStruktur.find((r) => String(r.id) === String(editId));
            if (!row) return;
            const nextName = normStr(editName?.value);
            if (!nextName) {
                await dlgAlert('Bitte eine Bezeichnung eingeben.', { title: 'Eingabe' });
                return;
            }
            row.bezeichnung = nextName;
            if (editBeschreibung) row.beschreibung = normStr(editBeschreibung.value);
            row.schuljahr = normStr(editYear?.value) || currentSchoolYearLabel();
            row.status = normStr(editStatus?.value) || 'Aktiv';
            row.syncStatus = normStr(editSyncStatus?.value) || 'Ausstehend';
            row.letzteFehlermeldung = normStr(editLastError?.value);

            const t = String(row.typ || '');
            if (t === 'Person') {
                row.personName = normStr(getEl('ssStructEditPersonName')?.value);
                row.personEmail = normStr(getEl('ssStructEditPersonEmail')?.value).toLowerCase();
            } else {
                row.personName = '';
                row.personEmail = '';
            }
            if (t === 'Jahrgang') {
                row.jgYear = normStr(editJgYear?.value);
                row.jgSuffix = normStr(editJgSuffix?.value);
            } else if (t === 'Arbeitsgemeinschaft') {
                row.argeCode = normStr(editArgeCode?.value);
                row.argeName = normStr(editArgeName?.value);
            } else if (t === 'Kursteam') {
                row.ktKlasse = normStr(editKtKlasse?.value);
                row.ktFach = normStr(editKtFach?.value);
                row.ktGruppe = normStr(editKtGruppe?.value);
            }

            // parent
            const pv = editParent && editParent.value ? String(editParent.value) : '';
            const root = inferRootForType(row.typ);
            const isRootSel = pv === '__root_students__' || pv === '__root_teachers__';
            if (isRootSel) {
                if (!canReparentStrict(row.typ, root)) {
                    await dlgAlert('Nicht erlaubt: ' + row.typ + ' kann nicht direkt unter Root liegen.', { title: 'Struktur' });
                    return;
                }
                row.parentId = '';
            } else {
                const pRow = rowsStruktur.find((r) => String(r.id) === String(pv));
                if (!pRow) {
                    await dlgAlert('Übergeordnetes Element ungültig.', { title: 'Struktur' });
                    return;
                }
                if (!canReparentStrict(row.typ, String(pRow.typ || ''))) {
                    await dlgAlert('Nicht erlaubt: ' + row.typ + ' kann nicht unter ' + pRow.typ + ' liegen.', { title: 'Struktur' });
                    return;
                }
                // cycle guard
                if (String(pRow.id) === String(row.id)) {
                    await dlgAlert('Nicht erlaubt: Zyklus.', { title: 'Struktur' });
                    return;
                }
                row.parentId = String(pRow.id);
            }

            saveState({ rows: rowsStruktur, memberships, settings: schemaState });
            emitWritebackFromStructure();
            selectedId = String(row.id);
            rerender();
            closeEditModal();
        }

        async function deleteEditModal() {
            if (!editId) return;
            if (editModalIsStructureTreeRoot || isStructureTreeRootId(editId)) {
                await dlgAlert('Die Hauptäste Schüler:innen, Lehrer:innen und Verwaltung können nicht gelöscht werden.', { title: 'Löschen' });
                return;
            }
            const row = rowsStruktur.find((r) => String(r.id) === String(editId));
            if (!row) return;
            const label = row.bezeichnung ? `"${row.bezeichnung}"` : 'diesen Eintrag';
            if (!(await dlgConfirm('Wirklich ' + label + ' löschen? (Unterpunkte werden ebenfalls entfernt.)', { title: 'Löschen', okText: 'Löschen', danger: true })))
                return;
            remove(row.id);
            selectedId = '';
            saveState({ rows: rowsStruktur, memberships, settings: schemaState });
            emitWritebackFromStructure();
            rerender();
            closeEditModal();
        }

        if (editSave) editSave.addEventListener('click', () => void saveEditModal());
        if (editDelete) editDelete.addEventListener('click', () => void deleteEditModal());
        if (editCancel) editCancel.addEventListener('click', () => closeEditModal());
        if (editClose) editClose.addEventListener('click', () => closeEditModal());
        if (editModal) {
            editModal.addEventListener('click', (ev) => {
                if (ev.target === editModal) closeEditModal();
            });
            window.addEventListener('keydown', (ev) => {
                if (ev.key === 'Escape' && editModal.classList.contains('active')) closeEditModal();
            });
        }

        function isTenantBulkMode() {
            return mode === 'tenant' && tenantMultiSel && tenantMultiSel.size >= 2;
        }

        function setTenantBulkModeUi(isBulk) {
            // Bulk-Controls (im Owner-Tab)
            const bulkWrap = getEl('ssTenantBulkWrap');
            const bulkCount = getEl('ssTenantBulkCount');
            if (bulkWrap) bulkWrap.style.display = isBulk ? '' : 'none';
            if (bulkCount) bulkCount.textContent = String(tenantMultiSel ? tenantMultiSel.size : 0);

            // Show either Bulk UI or Single UI (avoid duplicated "suchen/treffer")
            const singleWrap = getEl('ssTenantOwnerSingleWrap');
            if (singleWrap) singleWrap.style.display = isBulk ? 'none' : '';

            // Disable other details in bulk mode (Allgemein/Mitglieder + Single-Owner-Controls)
            const allgPanel = getEl('ssTenantTabAllgemein');
            const memPanel = getEl('ssTenantTabMitglieder');
            if (allgPanel) allgPanel.classList.toggle('bulk-disabled', !!isBulk);
            if (memPanel) memPanel.classList.toggle('bulk-disabled', !!isBulk);

            const disableIds = [
                // Allgemein
                'ssTenantName',
                'ssTenantDescription',
                'ssTenantArchiveState',
                'ssTenantArchiveSpoReadonly',
                'ssTenantUpdateBtn',
                'ssTenantReloadBtn',
                'ssTenantRenewBtn',
                'ssTenantDeleteBtn',
                // Owner (Single)
                'ssOwnerSearch',
                'ssOwnerSearchBtn',
                'ssOwnerSearchResults',
                'ssOwnerAddBtn',
                'ssOwnersReloadBtn',
                // Mitglieder
                'ssMemberSearch',
                'ssMemberSearchBtn',
                'ssMemberSearchResults',
                'ssMemberAddBtn',
                'ssMembersReloadBtn'
            ];
            for (const id of disableIds) {
                const el = getEl(id);
                if (!el) continue;
                try {
                    el.disabled = !!isBulk;
                } catch {
                    // ignore
                }
            }

            // Make lists non-interactive in bulk mode (remove buttons etc.)
            const ownersList = getEl('ssOwnersList');
            const membersList = getEl('ssMembersList');
            if (ownersList) ownersList.classList.toggle('bulk-disabled', !!isBulk);
            if (membersList) membersList.classList.toggle('bulk-disabled', !!isBulk);
        }

        function toggleStructureBranch(collapsedId) {
            const sid = String(collapsedId);
            if (graphCollapsed.has(sid)) graphCollapsed.delete(sid);
            else graphCollapsed.add(sid);
            saveGraphCollapsedSet(graphCollapsed);
            rerender();
        }

        function rerender() {
            if (mode === 'struktur' && rowsStruktur && rowsStruktur.length) {
                let st = null;
                try {
                    if (typeof window.ms365TenantSettingsLoad === 'function') {
                        st = window.ms365TenantSettingsLoad();
                    }
                } catch {
                    // ignore
                }
                if (ensureFachschaftFachGruppen(rowsStruktur, st)) {
                    saveState({ rows: rowsStruktur, memberships, settings: schemaState });
                }
            }
            const rows = mode === 'tenant' ? rowsTenant : rowsStruktur;
            renderFilters(rows, mode);
            renderTree(
                rows,
                selectedId,
                mode,
                mode === 'struktur' || mode === 'match' ? graphCollapsed : null,
                mode === 'struktur' || mode === 'match' ? toggleStructureBranch : null,
                mode === 'struktur' || mode === 'match' ? schemaState.structRootDetails : null,
                mode === 'struktur' || mode === 'match' ? openStructureAddPicker : null
            );
            setModeHint(mode, tenantCache.loadedAt || '');

            // Multi-select highlight + bulk UI
            if (mode === 'tenant') {
                try {
                    const tree = getEl('ssTree');
                    const buttons = tree ? tree.querySelectorAll('button[data-ss-select]') : [];
                    buttons.forEach((b) => {
                        const id = b.getAttribute('data-ss-select') || '';
                        b.classList.toggle('is-multi-selected', tenantMultiSel.has(id));
                    });
                } catch {
                    // ignore
                }
                const isBulk = tenantMultiSel.size >= 2;
                setTenantBulkModeUi(isBulk);
                // In Bulk-Modus automatisch Owner-Tab anzeigen (dort ist die Bulk-Maske).
                if (isBulk) setTenantTab('own');
            } else {
                setTenantBulkModeUi(false);
            }

            let sel = rows.find((r) => String(r.id) === String(selectedId));
            if ((mode === 'struktur' || mode === 'match') && !sel && isStructureTreeRootId(selectedId)) {
                sel = mergeStructureTreeRootRow(selectedId, schemaState.structRootDetails);
            }
            if (mode === 'tenant') {
                showTenantDetail(sel || null);
                if (!sel) {
                    ownersCache = [];
                    renderOwnersList([]);
                    fillUserSearchSelect([]);
                }
                const md = getEl('ssMatchDetail');
                if (md) md.style.display = 'none';
            } else {
                setDetailFromRow(sel || null, rowsStruktur);
                const td = getEl('ssTenantDetail');
                if (td) td.style.display = 'none';
            }
            // Schema-Vorschau (Anlegen) aktualisieren
            if (mode === 'struktur') {
                renderAnlegenSchemaUnitUi();
                updateStructTenantCreateUi();
                updateGraphQuickCreateBtn();
                refreshStrukturTypDependentUi({
                    mode,
                    selectedId,
                    rowsStruktur,
                    structRootDetails: schemaState.structRootDetails,
                    onPersonDetailTabs: applyPersonStructTabsForDetailPanel,
                    personInfoByRole
                });
            }

            // Organigramm (SOLL): Anlegen + Abgleich
            if (mode === 'struktur' || mode === 'match') {
                const graphPanel = getEl('ssGraphViewPanel');
                const treePanel = getEl('ssTreeViewPanel');
                const bTree = getEl('ssViewTabTreeBtn');
                const bGraph = getEl('ssViewTabGraphBtn');
                const activeGraph = bGraph && bGraph.getAttribute('aria-selected') === 'true';
                if (graphPanel && treePanel) {
                    graphPanel.classList.toggle('active', !!activeGraph);
                    treePanel.classList.toggle('active', !activeGraph);
                }
                if (activeGraph) {
                    const model = renderGraphView(
                        rowsStruktur,
                        selectedId,
                        (id, meta) => {
                            void (async () => {
                                if (!(await confirmMatchLeaveIfNeeded(id))) return;
                                select(id);
                                if (meta && meta.openDetails) openEditModalForId(id);
                            })();
                        },
                        graphViewport,
                        graphCollapsed,
                        personInfoByRole,
                        schemaState.structRootDetails,
                        schemaState.graphLayoutMode,
                        openStructureAddPicker
                    );
                    updateGraphToolbarState();
                    // Bind DnD after render
                    try {
                        const nodesHost = getEl('ssGraphNodes');
                        const wrap = getEl('ssGraphWrap');
                        if (nodesHost && wrap && model) {
                            if (!graphNodesToggleBound) {
                                graphNodesToggleBound = true;
                                nodesHost.addEventListener('click', (ev) => {
                                    const t = ev && ev.target ? ev.target : null;
                                    const btn = t && typeof t.closest === 'function' ? t.closest('.ss-graph-toggle') : null;
                                    if (!btn) return;
                                    ev.preventDefault();
                                    ev.stopPropagation();
                                    const id = String(btn.getAttribute('data-ss-toggle-for') || '');
                                    if (!id) return;
                                    if (graphCollapsed.has(id)) graphCollapsed.delete(id);
                                    else graphCollapsed.add(id);
                                    saveGraphCollapsedSet(graphCollapsed);
                                    rerender();
                                });
                            }
                            const nodeEls = Array.from(nodesHost.querySelectorAll('.ss-graph-node[data-ss-node-id]'));
                            nodeEls.forEach((el) => {
                                // plus button -> create modal
                                const plus = el.querySelector('button[data-ss-plus-for]');
                                if (plus) {
                                    plus.addEventListener('click', (ev) => {
                                        ev.preventDefault();
                                        ev.stopPropagation();
                                        const pid = el.getAttribute('data-ss-node-id') || '';
                                        const pt = el.getAttribute('data-ss-node-type') || '';
                                        if (typeof openStructureAddPicker === 'function') {
                                            openStructureAddPicker(pid, pt, plus);
                                        }
                                    });
                                }

                                el.addEventListener('dragstart', (ev) => {
                                    const did = el.getAttribute('data-ss-node-id') || '';
                                    if (isStructureSyntheticGraphNodeId(did)) {
                                        try {
                                            ev.preventDefault();
                                        } catch {
                                            // ignore
                                        }
                                        return;
                                    }
                                    const dt = el.getAttribute('data-ss-node-type') || '';
                                    graphDrag = { dragId: did, dragType: dt };
                                    try {
                                        ev.dataTransfer.effectAllowed = 'move';
                                        ev.dataTransfer.setData('text/plain', did);
                                    } catch {
                                        // ignore
                                    }
                                });
                                el.addEventListener('dragend', () => {
                                    graphDrag = null;
                                });
                                el.addEventListener('dragover', (ev) => {
                                    if (!graphDrag || !graphDrag.dragId) return;
                                    const targetId = el.getAttribute('data-ss-node-id') || '';
                                    if (!targetId || targetId === graphDrag.dragId) return;
                                    if (isStructureSyntheticGraphNodeId(targetId)) return;
                                    const targetType = el.getAttribute('data-ss-node-type') || '';
                                    const ok = canReparentStrict(graphDrag.dragType, targetType);
                                    if (!ok) return;
                                    ev.preventDefault();
                                    try {
                                        ev.dataTransfer.dropEffect = 'move';
                                    } catch {
                                        // ignore
                                    }
                                });
                                el.addEventListener('drop', (ev) => {
                                    if (!graphDrag || !graphDrag.dragId) return;
                                    const targetId = el.getAttribute('data-ss-node-id') || '';
                                    const targetType = el.getAttribute('data-ss-node-type') || '';
                                    if (!targetId || targetId === graphDrag.dragId) return;
                                    if (isStructureSyntheticGraphNodeId(targetId)) return;
                                    const ok = canReparentStrict(graphDrag.dragType, targetType);
                                    if (!ok) return;
                                    ev.preventDefault();

                                    const child = rowsStruktur.find((r) => String(r.id) === String(graphDrag.dragId));
                                    if (!child) return;
                                    void (async () => {
                                        if (!(await confirmMatchLeaveIfNeeded(String(child.id)))) return;
                                        // Virtual roots reset to ''
                                        const isRootTarget = targetType === 'SchuelerInnen' || targetType === 'LehrerInnen';
                                        child.parentId = isRootTarget ? '' : String(targetId);
                                        saveState({ rows: rowsStruktur, memberships, settings: schemaState });
                                        emitWritebackFromStructure();
                                        selectedId = String(child.id);
                                        rerender();
                                    })();
                                });
                            });

                            // allow dropping onto background -> root (only if allowed)
                            if (!graphWrapDnDBound) {
                                graphWrapDnDBound = true;
                                wrap.addEventListener('dragover', (ev) => {
                                    if (!graphDrag || !graphDrag.dragId) return;
                                    const rootType = inferRootForType(graphDrag.dragType);
                                    const ok =
                                        rootType === 'SchuelerInnen'
                                            ? canReparentStrict(graphDrag.dragType, 'SchuelerInnen')
                                            : rootType === 'LehrerInnen'
                                              ? canReparentStrict(graphDrag.dragType, 'LehrerInnen')
                                              : false;
                                    if (!ok) return;
                                    ev.preventDefault();
                                });
                                wrap.addEventListener('drop', (ev) => {
                                    if (!graphDrag || !graphDrag.dragId) return;
                                    const rootType = inferRootForType(graphDrag.dragType);
                                    const ok =
                                        rootType === 'SchuelerInnen'
                                            ? canReparentStrict(graphDrag.dragType, 'SchuelerInnen')
                                            : rootType === 'LehrerInnen'
                                              ? canReparentStrict(graphDrag.dragType, 'LehrerInnen')
                                              : false;
                                    if (!ok) return;
                                    ev.preventDefault();
                                    const child = rowsStruktur.find((r) => String(r.id) === String(graphDrag.dragId));
                                    if (!child) return;
                                    void (async () => {
                                        if (!(await confirmMatchLeaveIfNeeded(String(child.id)))) return;
                                        child.parentId = '';
                                        saveState({ rows: rowsStruktur, memberships, settings: schemaState });
                                        emitWritebackFromStructure();
                                        selectedId = String(child.id);
                                        rerender();
                                    })();
                                });

                                // Pan (drag background)
                                wrap.addEventListener('mousedown', (ev) => {
                                    // only left button and only background (not nodes/buttons)
                                    if (ev.button !== 0) return;
                                    const target = ev.target;
                                    const isInteractive = target && (target.closest && target.closest('.ss-graph-node'));
                                    if (isInteractive) return;
                                    graphPan = {
                                        startX: ev.clientX,
                                        startY: ev.clientY,
                                        baseX: graphViewport.x,
                                        baseY: graphViewport.y
                                    };
                                    wrap.style.cursor = 'grabbing';
                                    ev.preventDefault();
                                });
                                window.addEventListener('mousemove', (ev) => {
                                    if (!graphPan) return;
                                    const dx = ev.clientX - graphPan.startX;
                                    const dy = ev.clientY - graphPan.startY;
                                    graphViewport.x = graphPan.baseX + dx;
                                    graphViewport.y = graphPan.baseY + dy;
                                    rerender();
                                });
                                window.addEventListener('mouseup', () => {
                                    if (!graphPan) return;
                                    graphPan = null;
                                    try {
                                        wrap.style.cursor = '';
                                    } catch {
                                        // ignore
                                    }
                                });

                                // Zoom (mouse wheel)
                                wrap.addEventListener('wheel', (ev) => {
                                    // allow page scroll when not over organigramm
                                    ev.preventDefault();
                                    const rect = wrap.getBoundingClientRect();
                                    const cx = ev.clientX - rect.left;
                                    const cy = ev.clientY - rect.top;
                                    const prev = graphViewport.scale;
                                    const delta = ev.deltaY;
                                    const factor = delta > 0 ? 0.9 : 1.1;
                                    let next = prev * factor;
                                    next = Math.max(0.4, Math.min(2.4, next));
                                    const k = next / prev;
                                    // zoom around cursor: adjust translation so the point under cursor stays stable
                                    graphViewport.x = cx - (cx - graphViewport.x) * k;
                                    graphViewport.y = cy - (cy - graphViewport.y) * k;
                                    graphViewport.scale = next;
                                    rerender();
                                }, { passive: false });
                            }
                        }
                    } catch {
                        // ignore
                    }
                }
                if (!activeGraph) {
                    updateGraphToolbarState();
                }
            }


            // Match view uses structure selection but different right panel
            if (mode === 'match') {
                // expose caches for renderMatchDetail helper
                window.__ms365TenantRowsCache = rowsTenant;
                try {
                    window.__ms365TenantUsersCache = loadTenantCache().users || [];
                } catch {
                    window.__ms365TenantUsersCache = [];
                }
                window.__ms365MatchLinks = links;
                const td = getEl('ssTenantDetail');
                if (td) td.style.display = 'none';
                const sd = getEl('ssDetail');
                if (sd) sd.style.display = 'none';
                const md = getEl('ssMatchDetail');
                if (md) md.style.display = sel ? '' : 'none';
                if (sel) {
                    renderMatchDetail(sel);
                }
            }

            // Anlegen: wenn nichts ausgewählt ist, Settings anzeigen
            const settingsPanel = getEl('ssAnlegenSettingsPanel');
            if (settingsPanel) {
                const show = mode === 'struktur';
                settingsPanel.style.display = show ? '' : 'none';
                if (show) {
                    wireSchemaTabsOnce();
                    bindAnlegenSettingsUi();
                }
            }
        }

        async function synchronizeTenantTeamArchiveFlag() {
            const rid = selectedId ? String(selectedId) : '';
            if (mode !== 'tenant' || !rid) return;
            const row = rowsTenant.find((r) => String(r.id) === String(rid));
            const typ = String(row && row.typ ? row.typ : '');
            if (!row || (typ !== 'Team' && typ !== 'Gruppe')) return;
            try {
                const token = await getGraphToken(GRAPH_SCOPES_TENANT_READ);
                const ar = await resolveTeamsArchiveStateForUnifiedGroupId(rid, token);
                if (String(selectedId) !== rid) return;
                const idx = rowsTenant.findIndex((r) => String(r.id) === String(rid));
                if (idx === -1) return;
                rowsTenant[idx] = Object.assign({}, rowsTenant[idx], {
                    hasTeamsForArchive: ar.hasTeamsForArchive,
                    teamIsArchived: ar.teamIsArchived
                });
                saveTenantCache(rowsTenant);
                if (String(selectedId) === rid) {
                    showTenantDetail(rowsTenant[idx]);
                }
            } catch {
                if (String(selectedId) !== rid) return;
                const idx = rowsTenant.findIndex((r) => String(r.id) === String(rid));
                if (idx === -1) return;
                rowsTenant[idx] = Object.assign({}, rowsTenant[idx], {
                    hasTeamsForArchive: true,
                    teamIsArchived: null
                });
                saveTenantCache(rowsTenant);
                if (String(selectedId) === rid) showTenantDetail(rowsTenant[idx]);
            }
        }

        function select(id) {
            selectedId = id ? String(id) : '';
            rerender();
            if (mode === 'tenant' && selectedId) {
                // Owner-Liste automatisch nachladen (für alle Gruppentypen möglich).
                loadOwnersNow(false);
                loadMembersNow(false);
                void synchronizeTenantTeamArchiveFlag();
            }
            if (mode === 'struktur' && selectedId) {
                renderStructMembershipUi();
            }
        }

        /** Gespeicherte Match-Auswahl als g:/u:-Wert (wie im Dropdown), für Abgleich mit UI. */
        function persistedMatchSelectValueForRow(structureId) {
            const users = Array.isArray(window.__ms365TenantUsersCache) ? window.__ms365TenantUsersCache : [];
            const saved = (loadMatchState().links || {})[String(structureId)] || null;
            if (!saved) return '';
            if (normStr(saved.tenantUserId)) return 'u:' + String(saved.tenantUserId);
            if (normStr(saved.tenantGroupId)) {
                const gid = String(saved.tenantGroupId);
                const isUser = users.some((u) => String(u.id) === gid);
                return (isUser ? 'u:' : 'g:') + gid;
            }
            return '';
        }

        function isMatchDraftDirty() {
            if (mode !== 'match' || !selectedId) return false;
            const selTenant = getEl('ssMatchTenantGroup');
            const noteEl = getEl('ssMatchNote');
            if (!selTenant || !noteEl) return false;
            const curVal = String(selTenant.value || '').trim();
            const curNote = String(noteEl.value || '').trim();
            const savedVal = persistedMatchSelectValueForRow(selectedId);
            const savedObj = (loadMatchState().links || {})[String(selectedId)] || null;
            const savedNote = savedObj && savedObj.note ? String(savedObj.note).trim() : '';
            return curVal !== savedVal || curNote !== savedNote;
        }

        /**
         * Vor Wechsel der Strukturauswahl im Abgleich: ungespeicherte Zuordnung speichern, verwerfen oder abbrechen.
         * @param {string} [nextSelectedId] Ziel-ID (leer = Auswahl aufheben / Tab wechseln)
         */
        async function confirmMatchLeaveIfNeeded(nextSelectedId) {
            const next = nextSelectedId != null ? String(nextSelectedId) : '';
            if (mode !== 'match') return true;
            if (!isMatchDraftDirty()) return true;
            if (String(selectedId || '') === next) return true;

            const saveFirst = await dlgConfirm(
                'Die Tenant-Zuordnung wurde geändert, aber noch nicht mit „Verknüpfen“ übernommen. Zuerst speichern und dann wechseln?',
                { title: 'Ungesicherte Zuordnung', okText: 'Speichern und wechseln', cancelText: 'Zurück' }
            );
            if (saveFirst) {
                const selTenant = getEl('ssMatchTenantGroup');
                const note = getEl('ssMatchNote');
                saveLinkForSelected(selTenant ? selTenant.value : '', note ? note.value : '');
                toast('Zuordnung gespeichert.');
                return true;
            }
            const discard = await dlgConfirm(
                'Ohne Speichern wechseln? Die Änderung an der Zuordnung geht verloren.',
                { title: 'Änderung verwerfen', okText: 'Verwerfen und wechseln', cancelText: 'Abbrechen', danger: true }
            );
            return !!discard;
        }

        function upsert(row) {
            const id = String(row.id);
            const idx = rowsStruktur.findIndex((r) => String(r.id) === id);
            if (idx === -1) rowsStruktur.push(row);
            else rowsStruktur[idx] = row;
            saveState({ rows: rowsStruktur, memberships, settings: schemaState });
        }

        function remove(id) {
            const sid = String(id);
            // entferne auch direkte Kinder (Mock: 1 Ebene reicht fürs UI, reicht für v1)
            const childIds = rowsStruktur.filter((r) => String(r.parentId || '') === sid).map((r) => String(r.id));
            rowsStruktur = rowsStruktur.filter((r) => String(r.id) !== sid && childIds.indexOf(String(r.id)) === -1);
            try {
                delete memberships[sid];
                childIds.forEach((cid) => {
                    delete memberships[String(cid)];
                });
            } catch {
                // ignore
            }
            saveState({ rows: rowsStruktur, memberships, settings: schemaState });
            if (selectedId === sid) selectedId = '';
        }

        // Tree click
        getEl('ssTree')?.addEventListener('click', (ev) => {
            void (async () => {
                const t = ev.target;
                const openK = t && t.closest ? t.closest('button[data-ss-open-kursteams]') : null;
                if (openK) {
                    try {
                        ev.preventDefault();
                        ev.stopPropagation();
                    } catch {
                        /* ignore */
                    }
                    const inToolsFolder = String(window.location && window.location.pathname ? window.location.pathname : '').toLowerCase().includes('/tools/');
                    const href = (inToolsFolder ? '' : 'tools/') + 'kursteams.html';
                    window.open(href, '_blank', 'noopener');
                    return;
                }
                const btn = t && t.closest ? t.closest('button[data-ss-select]') : null;
                if (!btn) return;
                const id = btn.getAttribute('data-ss-select');
                if (!id) return;
                if (mode === 'tenant' && (ev.ctrlKey || ev.metaKey)) {
                    // toggle multi selection
                    if (tenantMultiSel.has(id)) tenantMultiSel.delete(id);
                    else tenantMultiSel.add(id);
                    // keep last clicked as details target
                    selectedId = id;
                    rerender();
                    return;
                }
                if (!(await confirmMatchLeaveIfNeeded(id))) {
                    try {
                        ev.preventDefault();
                        ev.stopPropagation();
                    } catch {
                        /* ignore */
                    }
                    return;
                }
                // single select resets multi selection
                if (mode === 'tenant') tenantMultiSel = new Set([id]);
                select(id);
            })();
        });

        // Tree context menu (right click)
        getEl('ssTree')?.addEventListener('contextmenu', (ev) => {
            const t = ev.target;
            const btn = t && t.closest ? t.closest('button[data-ss-select]') : null;
            if (!btn) return;
            const id = String(btn.getAttribute('data-ss-select') || '').trim();
            if (!id) return;

            // Only for structure/match: deleting tenant objects should be explicit in details UI.
            if (mode === 'tenant') return;

            ev.preventDefault();
            ev.stopPropagation();

            const menu = ensureTreeContextMenu();
            const title = document.getElementById('ssCtxMenuTitle');
            const hint = document.getElementById('ssCtxHint');
            const delBtn = document.getElementById('ssCtxDelete');

            // disallow delete for main roots + synthetic nodes
            const isRoot = isStructureTreeRootId(id);
            const isSynthetic = isStructureSyntheticGraphNodeId(id);
            const row = rowsStruktur.find((r) => String(r.id) === String(id));
            const label = row && row.bezeichnung ? String(row.bezeichnung) : isRoot ? structureTreeRootDefaultTitle(id) : '';

            if (title) title.textContent = label || 'Eintrag';
            if (hint) {
                if (isRoot || isSynthetic || !row) {
                    hint.style.display = '';
                    hint.textContent = isRoot
                        ? 'Die Hauptäste können nicht gelöscht werden.'
                        : isSynthetic
                          ? 'Dieser Eintrag ist ein Hinweis/virtueller Knoten und kann nicht gelöscht werden.'
                          : 'Dieser Eintrag kann nicht gelöscht werden.';
                } else {
                    hint.style.display = 'none';
                    hint.textContent = '';
                }
            }

            if (delBtn) {
                delBtn.disabled = !!(isRoot || isSynthetic || !row);
                delBtn.onclick = async () => {
                    // @ts-ignore
                    if (menu && typeof menu.hide === 'function') menu.hide();
                    if (!row) return;
                    const rowLabel = row.bezeichnung ? `"${row.bezeichnung}"` : 'diesen Eintrag';
                    if (!(await dlgConfirm('Wirklich ' + rowLabel + ' löschen? (Unterpunkte werden ebenfalls entfernt.)', { title: 'Löschen', okText: 'Löschen', danger: true })))
                        return;
                    remove(row.id);
                    selectedId = '';
                    saveState({ rows: rowsStruktur, memberships, settings: schemaState });
                    emitWritebackFromStructure();
                    rerender();
                };
            }

            const vw = window.innerWidth || 0;
            const vh = window.innerHeight || 0;
            const rect = menu.getBoundingClientRect();
            let x = ev.clientX;
            let y = ev.clientY;
            // keep inside viewport
            if (x + rect.width + 8 > vw) x = Math.max(8, vw - rect.width - 8);
            if (y + rect.height + 8 > vh) y = Math.max(8, vh - rect.height - 8);
            menu.style.left = x + 'px';
            menu.style.top = y + 'px';
            menu.style.display = 'block';
        });

        getEl('ssTree')?.addEventListener('dblclick', (ev) => {
            void (async () => {
                if (mode !== 'struktur' && mode !== 'match') return;
                const t = ev.target;
                const btn = t && t.closest ? t.closest('button[data-ss-select]') : null;
                if (!btn) return;
                const id = btn.getAttribute('data-ss-select');
                if (!id) return;
                try {
                    ev.preventDefault();
                } catch {
                    /* ignore */
                }
                if (!(await confirmMatchLeaveIfNeeded(id))) return;
                select(id);
                openEditModalForId(id);
            })();
        });

        // Tree drag & drop (SOLL: Anlegen + Abgleich)
        /** @type {{id: string, type: string}|null} */
        let treeDrag = null;

        function isDescendant(childId, maybeAncestorId) {
            const map = new Map(rowsStruktur.map((r) => [String(r.id), r]));
            let cur = map.get(String(maybeAncestorId));
            let guard = 0;
            while (cur && guard++ < 30) {
                const pid = String(cur.parentId || '');
                if (!pid) return false;
                if (pid === String(childId)) return true;
                cur = map.get(pid);
            }
            return false;
        }

        getEl('ssTree')?.addEventListener('dragstart', (ev) => {
            if (mode !== 'struktur' && mode !== 'match') return;
            const t = ev.target;
            const btn = t && t.closest ? t.closest('button[data-ss-select]') : null;
            if (!btn) return;
            const id = btn.getAttribute('data-ss-select') || '';
            const typ = btn.getAttribute('data-ss-type') || '';
            treeDrag = { id, type: typ };
            try {
                ev.dataTransfer.effectAllowed = 'move';
                ev.dataTransfer.setData('text/plain', id);
            } catch {
                // ignore
            }
        });
        getEl('ssTree')?.addEventListener('dragend', () => {
            treeDrag = null;
        });
        getEl('ssTree')?.addEventListener('dragover', (ev) => {
            if ((mode !== 'struktur' && mode !== 'match') || !treeDrag) return;
            const t = ev.target;
            const btn = t && t.closest ? t.closest('button[data-ss-select]') : null;
            if (!btn) {
                // allow dropping to root if rule allows
                const rootType = inferRootForType(treeDrag.type);
                const ok = rootType === 'SchuelerInnen'
                    ? canReparentStrict(treeDrag.type, 'SchuelerInnen')
                    : rootType === 'LehrerInnen'
                      ? canReparentStrict(treeDrag.type, 'LehrerInnen')
                      : false;
                if (ok) ev.preventDefault();
                return;
            }
            const targetId = btn.getAttribute('data-ss-select') || '';
            const targetType = btn.getAttribute('data-ss-type') || '';
            if (!targetId || targetId === treeDrag.id) return;
            if (isDescendant(treeDrag.id, targetId)) return;
            const ok = canReparentStrict(treeDrag.type, targetType);
            if (!ok) return;
            ev.preventDefault();
            try {
                ev.dataTransfer.dropEffect = 'move';
            } catch {
                // ignore
            }
        });
        getEl('ssTree')?.addEventListener('drop', (ev) => {
            if ((mode !== 'struktur' && mode !== 'match') || !treeDrag) return;
            const t = ev.target;
            const btn = t && t.closest ? t.closest('button[data-ss-select]') : null;
            const child = rowsStruktur.find((r) => String(r.id) === String(treeDrag.id));
            if (!child) return;

            if (!btn) {
                const rootType = inferRootForType(treeDrag.type);
                const ok = rootType === 'SchuelerInnen'
                    ? canReparentStrict(treeDrag.type, 'SchuelerInnen')
                    : rootType === 'LehrerInnen'
                      ? canReparentStrict(treeDrag.type, 'LehrerInnen')
                      : false;
                if (!ok) return;
                ev.preventDefault();
                void (async () => {
                    if (!(await confirmMatchLeaveIfNeeded(String(child.id)))) return;
                    child.parentId = '';
                    saveState({ rows: rowsStruktur, memberships, settings: schemaState });
                    emitWritebackFromStructure();
                    selectedId = String(child.id);
                    rerender();
                })();
                return;
            }

            const targetId = btn.getAttribute('data-ss-select') || '';
            const targetType = btn.getAttribute('data-ss-type') || '';
            if (!targetId || targetId === treeDrag.id) return;
            if (isDescendant(treeDrag.id, targetId)) return;
            const ok = canReparentStrict(treeDrag.type, targetType);
            if (!ok) return;
            ev.preventDefault();
            void (async () => {
                if (!(await confirmMatchLeaveIfNeeded(String(child.id)))) return;
                child.parentId = String(targetId);
                saveState({ rows: rowsStruktur, memberships, settings: schemaState });
                emitWritebackFromStructure();
                selectedId = String(child.id);
                rerender();
            })();
        });

        // Filter
        const onFilter = () => rerender();
        getEl('ssFilterSchuljahr')?.addEventListener('change', onFilter);
        getEl('ssFilterTyp')?.addEventListener('change', onFilter);
        getEl('ssTenantVisibilityFilter')?.addEventListener('change', onFilter);
        getEl('ssTenantRosterFilter')?.addEventListener('change', onFilter);
        getEl('ssFilterText')?.addEventListener('input', onFilter);

        // View tabs: Baum ↔ Organigramm (SOLL)
        const viewTreeBtn = getEl('ssViewTabTreeBtn');
        const viewGraphBtn = getEl('ssViewTabGraphBtn');
        const viewTreePanel = getEl('ssTreeViewPanel');
        const viewGraphPanel = getEl('ssGraphViewPanel');
        function setView(v) {
            const isGraph = v === 'graph';
            if (viewTreeBtn) viewTreeBtn.setAttribute('aria-selected', isGraph ? 'false' : 'true');
            if (viewGraphBtn) viewGraphBtn.setAttribute('aria-selected', isGraph ? 'true' : 'false');
            if (viewTreePanel) viewTreePanel.classList.toggle('active', !isGraph);
            if (viewGraphPanel) viewGraphPanel.classList.toggle('active', isGraph);
            rerender();
        }
        if (viewTreeBtn) viewTreeBtn.addEventListener('click', () => setView('tree'));
        if (viewGraphBtn) viewGraphBtn.addEventListener('click', () => setView('graph'));

        // Struktur-Modals: im Browser-Vollbild nur sichtbar, wenn sie Nachkommen des Fullscreen-Knotens sind
        function isGraphFullscreenTarget(el) {
            if (!el) return false;
            const id = el.id || '';
            return id === 'ssGraphViewPanel' || id === 'ssGraphFullscreenContainer' || id === 'ssGraphWrap';
        }
        function syncStructureModalsWithFullscreen() {
            const home = getEl('ssStructModalsHome');
            const createModal = getEl('ssStructCreateModal');
            const editModal = getEl('ssStructEditModal');
            if (!home || !createModal || !editModal) return;
            const fs = document.fullscreenElement;
            if (isGraphFullscreenTarget(fs)) {
                fs.appendChild(createModal);
                fs.appendChild(editModal);
            } else {
                home.appendChild(createModal);
                home.appendChild(editModal);
            }
        }

        // Organigramm Vollbild
        const fsBtn = getEl('ssGraphFullscreenBtn');
        function setFsBtnLabel() {
            if (!fsBtn) return;
            const on = !!document.fullscreenElement;
            fsBtn.innerHTML = on
                ? '<i class="bi bi-fullscreen-exit"></i>Vollbild beenden'
                : '<i class="bi bi-arrows-fullscreen"></i>Vollbild';
        }
        if (fsBtn) {
            fsBtn.addEventListener('click', async () => {
                const container = getEl('ssGraphViewPanel') || getEl('ssGraphFullscreenContainer') || getEl('ssGraphWrap');
                if (!container) return;
                try {
                    if (document.fullscreenElement) {
                        await document.exitFullscreen();
                    } else {
                        // Fullscreen on the whole panel so the header/toolbar stays visible
                        await container.requestFullscreen();
                    }
                } catch (e) {
                    toast('Vollbild nicht möglich: ' + (e?.message || String(e)));
                }
                setFsBtnLabel();
                // Nach Toggle einmal rerendern, damit Canvas-Größen passen
                setTimeout(() => rerender(), 50);
            });
            document.addEventListener('fullscreenchange', () => {
                syncStructureModalsWithFullscreen();
                setFsBtnLabel();
                setTimeout(() => rerender(), 50);
            });
            setFsBtnLabel();
            syncStructureModalsWithFullscreen();
        }

        // Organigramm Toolbar Actions
        const graphEditBtn = getEl('ssGraphEditBtn');
        const graphCopyBtn = getEl('ssGraphCopyBtn');
        const graphDeleteBtn = getEl('ssGraphDeleteBtn');
        const graphCollapseAllBtn = getEl('ssGraphCollapseAllBtn');
        const graphExpandToClassesBtn = getEl('ssGraphExpandToClassesBtn');

        function syncGraphLayoutToggleUi() {
            const hBtn = getEl('ssGraphLayoutHorizontalBtn');
            const vBtn = getEl('ssGraphLayoutVerticalBtn');
            if (!hBtn && !vBtn) return;
            const vert = schemaState.graphLayoutMode === 'vertical';
            if (hBtn) {
                hBtn.classList.toggle('ss-graph-layout-btn--active', !vert);
                hBtn.setAttribute('aria-pressed', vert ? 'false' : 'true');
            }
            if (vBtn) {
                vBtn.classList.toggle('ss-graph-layout-btn--active', !!vert);
                vBtn.setAttribute('aria-pressed', vert ? 'true' : 'false');
            }
        }

        function updateGraphToolbarState() {
            const row = rowsStruktur.find((r) => String(r.id) === String(selectedId));
            const enabled = !!row || isStructureTreeRootId(selectedId);
            if (graphEditBtn) graphEditBtn.disabled = !enabled;
            if (graphCopyBtn) graphCopyBtn.disabled = !row;
            if (graphDeleteBtn) graphDeleteBtn.disabled = !row;
            syncGraphLayoutToggleUi();
        }

        function collapseAllGraph() {
            const next = new Set();
            for (const r of rowsStruktur) {
                if (r && r.id) next.add(String(r.id));
            }
            graphCollapsed = next;
            saveGraphCollapsedSet(graphCollapsed);
            rerender();
        }

        function expandToClassesGraph() {
            const next = new Set();
            for (const r of rowsStruktur) {
                const t = String(r?.typ || '');
                // Unter Schüler:innen: Klassen als "Stop" (Kursteams/Gruppe darunter einklappen)
                // Unter Lehrer:innen: Arbeitsgemeinschaft/Gruppe als "Stop"
                if (t === 'Klasse' || t === 'Arbeitsgemeinschaft' || t === 'Gruppe') {
                    if (r && r.id) next.add(String(r.id));
                }
            }
            graphCollapsed = next;
            saveGraphCollapsedSet(graphCollapsed);
            rerender();
        }

        if (graphCollapseAllBtn) {
            graphCollapseAllBtn.addEventListener('click', () => collapseAllGraph());
        }
        if (graphExpandToClassesBtn) {
            graphExpandToClassesBtn.addEventListener('click', () => expandToClassesGraph());
        }

        function applyGraphLayoutMode(next) {
            const v = next === 'vertical' ? 'vertical' : 'horizontal';
            if (schemaState.graphLayoutMode === v) return;
            schemaState.graphLayoutMode = v;
            normalizeGraphLayoutModeInSettings(schemaState);
            saveState({ rows: rowsStruktur, memberships, settings: schemaState });
            graphViewport = { x: 0, y: 0, scale: 1 };
            rerender();
        }
        getEl('ssGraphLayoutHorizontalBtn')?.addEventListener('click', () => applyGraphLayoutMode('horizontal'));
        getEl('ssGraphLayoutVerticalBtn')?.addEventListener('click', () => applyGraphLayoutMode('vertical'));

        if (graphEditBtn) {
            graphEditBtn.addEventListener('click', () => {
                if (!selectedId) return;
                openEditModalForId(selectedId);
            });
        }
        if (graphDeleteBtn) {
            graphDeleteBtn.addEventListener('click', () => {
                if (!selectedId) return;
                // reuse edit modal delete semantics
                openEditModalForId(selectedId);
                try {
                    const btn = getEl('ssStructEditDelete');
                    if (btn) btn.focus();
                } catch {
                    // ignore
                }
            });
        }
        if (graphCopyBtn) {
            graphCopyBtn.addEventListener('click', async () => {
                const row = rowsStruktur.find((r) => String(r.id) === String(selectedId));
                if (!row) return;
                const payload = JSON.stringify(row, null, 2);
                try {
                    await navigator.clipboard.writeText(payload);
                    toast('Kopiert.');
                } catch (e) {
                    toast('Kopieren fehlgeschlagen: ' + (e?.message || String(e)));
                }
            });
        }

        // Mode switch
        const tabStruktur = getEl('ssTabStrukturTop');
        const tabMatch = getEl('ssTabAbgleichenTop');
        const tabTenant = getEl('ssTabTenantTop');
        const btnNeu = getEl('ssBtnNeu');
        const btnDemo = getEl('ssBtnDemo');
        const btnReset = getEl('ssBtnReset');
        const btnLoad = getEl('ssBtnTenantLoadTop');
        const tenantKindSel = null; // Gruppentyp-Filter entfernt
        const tenantKindWrap = null;
        const liveBanner = getEl('ssLiveBanner');
        const filterTypWrap = getEl('ssFilterTypWrap');
        const filterTypSel = getEl('ssFilterTyp');

        async function setActiveTab(nextMode) {
            // „Anlegen“ wird als Grundkonfiguration verstanden und lebt im Modul „Schul‑Grundeinstellungen“.
            if (nextMode === 'struktur') {
                if (isEmbedStructure) {
                    if (!(await confirmMatchLeaveIfNeeded(''))) return false;
                    mode = 'struktur';
                    selectedId = '';
                    updateModeUi();
                    rerender();
                    return true;
                }
                if (!(await confirmMatchLeaveIfNeeded(''))) return false;
                try {
                    window.location.href = '../tenant.html';
                } catch {
                    // ignore
                }
                return true;
            }
            if (!(await confirmMatchLeaveIfNeeded(''))) return false;
            mode = nextMode === 'tenant' ? 'tenant' : nextMode === 'match' ? 'match' : 'struktur';
            selectedId = '';
            if (tabStruktur) tabStruktur.setAttribute('aria-selected', mode === 'struktur' ? 'true' : 'false');
            if (tabMatch) tabMatch.setAttribute('aria-selected', mode === 'match' ? 'true' : 'false');
            if (tabTenant) tabTenant.setAttribute('aria-selected', mode === 'tenant' ? 'true' : 'false');
            try {
                sessionStorage.setItem(UI_MODE_KEY, mode);
            } catch {
                // ignore
            }
            updateModeUi();
            rerender();
            return true;
        }

        function updateModeUi() {
            if (isEmbedStructure) {
                // Embedded in Schul‑Grundeinstellungen: only structure planning UI is available.
                mode = 'struktur';
            }
            const isTenant = mode === 'tenant';
            const isMatch = mode === 'match';
            const rollWrap = getEl('ssWrapSchoolYearRollBtn');
            const sjWrap = getEl('ssFilterSchuljahrWrap');
            if (rollWrap) rollWrap.style.display = isTenant ? 'none' : '';
            if (sjWrap) sjWrap.style.display = isTenant ? 'none' : '';
            if (btnNeu) btnNeu.style.display = isTenant ? 'none' : '';
            if (btnDemo) btnDemo.style.display = isTenant ? 'none' : '';
            if (btnReset) btnReset.style.display = isTenant ? 'none' : '';
            if (btnLoad) btnLoad.style.display = isTenant ? '' : 'none';
            // Gruppentyp-Filter entfernt -> nichts anzeigen
            const visWrap = getEl('ssTenantVisibilityFilterWrap');
            if (visWrap) visWrap.style.display = isTenant ? '' : 'none';
            const rosterWrap = getEl('ssTenantRosterFilterWrap');
            if (rosterWrap) rosterWrap.style.display = isTenant ? '' : 'none';
            if (liveBanner) liveBanner.style.display = isTenant ? '' : 'none';
            const matchBanner = getEl('ssMatchBanner');
            if (matchBanner) matchBanner.style.display = isMatch ? '' : 'none';
            // Im Tenant-Modus wollen wir "Typ" als Filter (Team/Gruppe/Sicherheitsgruppe …) nutzen.
            if (filterTypWrap) filterTypWrap.style.display = '';
            if (!isTenant) {
                // Tenant-spezifische Filter zurücksetzen, damit sie nicht "unsichtbar" filtern
                const vSel = getEl('ssTenantVisibilityFilter');
                if (vSel) vSel.value = '';
                const rSel = getEl('ssTenantRosterFilter');
                if (rSel) rSel.value = '';
            }

            // Baum/Organigramm: im Tenant-Modus nur Liste (kein Organigramm für Tenant-Inventar)
            const tabBar = getEl('ssStructureViewTabsBar');
            const treePanel = getEl('ssTreeViewPanel');
            const graphPanel = getEl('ssGraphViewPanel');
            const bTree = getEl('ssViewTabTreeBtn');
            const bGraph = getEl('ssViewTabGraphBtn');
            if (isTenant) {
                if (tabBar) tabBar.style.display = 'none';
                if (treePanel) {
                    treePanel.classList.add('active');
                    treePanel.style.display = '';
                }
                if (graphPanel) {
                    graphPanel.classList.remove('active');
                    graphPanel.style.display = 'none';
                }
                if (bTree) bTree.setAttribute('aria-selected', 'true');
                if (bGraph) bGraph.setAttribute('aria-selected', 'false');
            } else {
                if (tabBar) tabBar.style.display = '';
                if (treePanel) treePanel.style.display = '';
                if (graphPanel) graphPanel.style.display = '';
            }

            // Multi selection should not leak across modes
            if (!isTenant) tenantMultiSel = new Set();
        }
        if (tabStruktur) tabStruktur.addEventListener('click', () => void setActiveTab('struktur'));
        if (tabMatch) tabMatch.addEventListener('click', () => void setActiveTab('match'));
        if (tabTenant) tabTenant.addEventListener('click', () => void setActiveTab('tenant'));
        async function reloadTenantNow(reasonText) {
            const btn = getEl('ssBtnTenantLoadTop');
            if (btn) btn.disabled = true;
            try {
                setTenantProgress(
                    true,
                    (reasonText ? reasonText + ' – ' : '') + 'lese M365‑Gruppen/Teams und Benutzer (Entra) …',
                    0.02
                );
                await loadTenantInventoryFull((p) => {
                    if (p && p.phase === 'users') {
                        setTenantProgress(
                            true,
                            `Benutzer: Seite ${p.page} … (${p.loaded})`,
                            Math.min(0.95, 0.5 + (p.page || 0) * 0.04)
                        );
                    } else if (p && p.phase === 'counts') {
                        const tot = Math.max(1, p.total || 1);
                        const ld = Math.min(p.loaded || 0, tot);
                        setTenantProgress(
                            true,
                            `Besitzer/Mitglieder zählen … ${ld} / ${tot}`,
                            Math.min(0.48, 0.12 + (ld / tot) * 0.34)
                        );
                    } else {
                        const ratio = p && p.page ? Math.min(0.45, 0.08 + p.page * 0.08) : 0.12;
                        setTenantProgress(
                            true,
                            p && p.page
                                ? `Gruppen: Seite ${p.page} – ${p.loaded} …` + (p.hasMore ? '' : ' (fertig)')
                                : 'Gruppen …',
                            ratio
                        );
                    }
                });
                const cache = loadTenantCache();
                rowsTenant = cache.rows.slice();
                tenantCache.loadedAt = cache.loadedAt;
                window.__ms365TenantRowsCache = rowsTenant;
                try {
                    window.__ms365TenantUsersCache = Array.isArray(cache.users) ? cache.users : [];
                } catch {
                    // ignore
                }
                const nu = Array.isArray(cache.users) ? cache.users.length : 0;
                setTenantProgress(true, `Fertig: ${rowsTenant.length} Gruppe(n)/Team(s), ${nu} Benutzer.`, 1);
                setTimeout(() => setTenantProgress(false, '', null), 2200);
                selectedId = '';
                rerender();
            } catch (e) {
                setTenantProgress(true, 'Fehler beim Einlesen: ' + (e?.message || String(e)), null);
            } finally {
                if (btn) btn.disabled = false;
            }
        }

        // Gruppentyp-Filter entfernt -> kein handler

        // Buttons
        getEl('ssBtnDemo')?.addEventListener('click', () => {
            const demoPack = buildDemoRows();
            rowsStruktur = demoPack.rows || [];
            memberships = {};

            const ts = demoPack.tenantSettings;
            if (ts && Array.isArray(ts.classes)) {
                const classOwners = new Map();
                ts.classes.forEach((c) => {
                    const label = normalizeClassLabel(c);
                    const mail = c && c.headEmail ? String(c.headEmail).trim().toLowerCase() : '';
                    const name = c && c.headName ? String(c.headName).trim() : '';
                    if (label && (mail || name)) {
                        classOwners.set(label, {
                            id: mail || name,
                            displayName: name || mail,
                            userPrincipalName: mail || ''
                        });
                    }
                });
                rowsStruktur.forEach((u) => {
                    if (!u || u.typ !== 'Klasse') return;
                    const label = String(u.bezeichnung || '').trim();
                    const owner = classOwners.get(label);
                    if (!owner) return;
                    memberships[String(u.id)] = { owners: [owner], members: [] };
                });
            }

            saveState({ rows: rowsStruktur, memberships, settings: schemaState });
            selectedId = '';
            rerender();
        });

        function ensureHeaderTenantLoadButton() {
            if (isEmbedStructure) return;
            try {
                const slot =
                    typeof window.ms365AuthGetActionSlot === 'function' ? window.ms365AuthGetActionSlot() : null;
                if (!slot) return;
                if (document.getElementById('ssHeaderTenantLoadBtn')) return;
                const hb = document.createElement('button');
                hb.type = 'button';
                hb.className = 'btn btn-success small-btn';
                hb.id = 'ssHeaderTenantLoadBtn';
                hb.style.margin = '0';
                hb.innerHTML = '<i class="bi bi-cloud-download"></i>Tenant einlesen';
                hb.addEventListener('click', async () => {
                    try {
                        sessionStorage.setItem(UI_MODE_KEY, 'tenant');
                    } catch {
                        // ignore
                    }
                    if (mode !== 'tenant') {
                        if (!(await setActiveTab('tenant'))) return;
                    }
                    await reloadTenantNow('Starte');
                });
                slot.appendChild(hb);
            } catch {
                // ignore
            }
        }

        // sofort versuchen + nachträglich wenn Widget fertig ist
        ensureHeaderTenantLoadButton();
        try {
            window.addEventListener('ms365-auth-widget-ready', () => ensureHeaderTenantLoadButton(), { once: true });
        } catch {
            // ignore
        }
        getEl('ssBtnReset')?.addEventListener('click', async () => {
            if (!(await dlgConfirm('Lokale Mock-Daten wirklich löschen?', { title: 'Zurücksetzen', okText: 'Löschen', danger: true }))) return;
            rowsStruktur = [];
            memberships = {};
            saveState({ rows: rowsStruktur, memberships, settings: schemaState });
            selectedId = '';
            rerender();
        });

        function applySchoolYearRollforward(targetSchoolYearLabel, maxStufen) {
            const tgt = String(targetSchoolYearLabel || '').trim();
            const ms = isFinite(maxStufen) ? Math.max(1, Math.min(12, Math.round(maxStufen))) : 5;
            schemaState.maxSchulstufen = ms;

            // Jahrgänge: anhand Abschlussjahr (jgYear) neue Stufe berechnen, ggf. archivieren.
            rowsStruktur.forEach((r) => {
                if (!r) return;
                if (String(r.typ || '') === 'Jahrgang') {
                    const g = gradeFromGraduationYear(String(r.jgYear || '').trim(), tgt, ms);
                    if (isFinite(g) && g >= 1 && g <= ms) {
                        r.bezeichnung = 'Jahrgang ' + String(Math.round(g));
                        if (String(r.status || '') === 'Archiviert') r.status = 'Aktiv';
                    } else if (isFinite(g) && g > ms) {
                        r.status = 'Archiviert';
                    }
                    r.schuljahr = tgt;
                }
            });

            // Klassen: neue Stufe aus classGradYear, Bezeichnung vorne anpassen, ggf. archivieren, Schuljahr setzen,
            // parentId auf passenden Jahrgang (jgYear) binden.
            const jahrgangByYear = new Map(
                rowsStruktur
                    .filter((r) => r && String(r.typ || '') === 'Jahrgang' && String(r.jgYear || '').trim())
                    .map((r) => [String(r.jgYear || '').trim(), r])
            );
            rowsStruktur.forEach((r) => {
                if (!r) return;
                if (String(r.typ || '') !== 'Klasse') return;
                const gy = String(r.classGradYear || '').trim();
                if (!gy) {
                    r.schuljahr = tgt;
                    return;
                }
                const g = gradeFromGraduationYear(gy, tgt, ms);
                if (isFinite(g) && g >= 1 && g <= ms) {
                    r.bezeichnung = replaceLeadingNumber(String(r.bezeichnung || ''), g);
                    if (String(r.status || '') === 'Archiviert') r.status = 'Aktiv';
                    const jg = jahrgangByYear.get(gy);
                    if (jg) r.parentId = String(jg.id || '');
                } else if (isFinite(g) && g > ms) {
                    r.status = 'Archiviert';
                }
                r.schuljahr = tgt;
            });

            saveState({ rows: rowsStruktur, memberships, settings: schemaState });
            rerender();
            toast('Schuljahrswechsel angewendet: ' + tgt + ' (Stufen: ' + String(ms) + ').');
        }

        getEl('ssBtnSchoolYearRoll')?.addEventListener('click', async () => {
            if (mode !== 'struktur') return;
            const cur = normStr(getEl('ssFilterSchuljahr')?.value) || currentSchoolYearLabel();
            const suggest = nextSchoolYearLabel(cur);
            const tgt = await dlgPrompt('Neues Schuljahr (z. B. 2027/28)', suggest, { title: 'Schuljahr', inputLabel: 'Bezeichnung' });
            if (tgt == null || !normStr(tgt)) return;
            const msRaw = await dlgPrompt('Max. Schulstufen (3, 4, 5 oder 8)', String(schemaState.maxSchulstufen || 5), {
                title: 'Schulstufen',
                inputLabel: 'Anzahl (3, 4, 5 oder 8)'
            });
            if (msRaw == null) return;
            const ms = parseInt(String(msRaw || '').trim(), 10);
            if (![3, 4, 5, 8].includes(ms)) {
                await dlgAlert('Bitte 3, 4, 5 oder 8 eingeben.', { title: 'Eingabe' });
                return;
            }
            const ok = await dlgConfirm(
                'Schuljahrswechsel auf ' +
                    String(tgt).trim() +
                    ' anwenden?\n\n- Klassen/Jahrgänge werden hochgestuft\n- Ausgelaufene Einheiten werden archiviert\n- Tenant-Verknüpfungen bleiben unverändert',
                { title: 'Schuljahrswechsel', okText: 'Anwenden', danger: true }
            );
            if (!ok) return;
            applySchoolYearRollforward(String(tgt).trim(), ms);
        });
        getEl('ssBtnNeu')?.addEventListener('click', () => {
            void (async () => {
                const r = { id: uid(), parentId: '', typ: 'Klasse', bezeichnung: '', beschreibung: '', schuljahr: '', status: 'Aktiv', syncStatus: 'Ausstehend', letzteFehlermeldung: '' };
                rowsStruktur.push(r);
                saveState({ rows: rowsStruktur, memberships, settings: schemaState });
                if (!(await confirmMatchLeaveIfNeeded(String(r.id)))) return;
                select(r.id);
            })();
        });
        getEl('ssBtnGraphCreate')?.addEventListener('click', async () => {
            if (mode !== 'struktur' || !selectedId) return;
            const cur = rowsStruktur.find((r) => String(r.id) === String(selectedId));
            if (!cur) return;
            const t = String(cur.typ || '');
            if (t === 'Person') await createSelectedPersonUserInTenant();
            else if (t === 'Gruppe') await createSelectedStructInTenant();
        });
        getEl('ssBtnSpeichern')?.addEventListener('click', async () => {
            if (isStructureTreeRootId(selectedId)) {
                const base = mergeStructureTreeRootRow(selectedId, schemaState.structRootDetails);
                if (!base) return;
                const next = readDetailToRow(base);
                if (!next.bezeichnung) {
                    await dlgAlert('Bitte eine Bezeichnung eingeben.', { title: 'Eingabe' });
                    return;
                }
                if (!schemaState.structRootDetails || typeof schemaState.structRootDetails !== 'object') {
                    schemaState.structRootDetails = {};
                }
                schemaState.structRootDetails[String(selectedId)] = pickStorableStructureTreeRootFields(next);
                saveState({ rows: rowsStruktur, memberships, settings: schemaState });
                rerender();
                return;
            }
            const cur = rowsStruktur.find((r) => String(r.id) === String(selectedId));
            if (!cur) return;
            const next = readDetailToRow(cur);
            if (!next.bezeichnung) {
                await dlgAlert('Bitte eine Bezeichnung eingeben.', { title: 'Eingabe' });
                return;
            }
            // simple cycle guard
            if (next.parentId && String(next.parentId) === String(next.id)) next.parentId = '';
            if (!isValidStructureParentChild(next.typ, next.parentId, rowsStruktur)) {
                await dlgAlert(
                    'Die Kombination aus Typ und übergeordneter Einheit ist nicht erlaubt (z. B. „Person“ nur unter „Gruppe“).',
                    { title: 'Struktur' }
                );
                return;
            }
            upsert(next);
            rerender();
        });

        function computeAnlegenSuggestionFromInputs() {
            const typ = normStr(getEl('ssTyp')?.value);
            const domain = String(schemaState.domain || '').trim();
            if (typ === 'Jahrgang') {
                const y = normStr(getEl('ssJgYear')?.value);
                const suf = normStr(getEl('ssJgSuffix')?.value);
                const mailNick = y && suf ? buildJgMailNick(schemaState, y, suf) : '';
                const displayName = normStr(getEl('ssBezeichnung')?.value) || (y && suf ? `Jahrgang ${y} ${suf}` : '');
                return { displayName, mailNick, email: mailNick && domain ? mailNick + '@' + domain : '' };
            }
            if (typ === 'Arbeitsgemeinschaft') {
                const code = normStr(getEl('ssArgeCode')?.value);
                const name = normStr(getEl('ssArgeName')?.value);
                const displayName = name ? `ARGE ${name}` : (code ? `ARGE ${code}` : '');
                const mailNick = code ? buildArgeMailNick(schemaState, code) : '';
                return { displayName, mailNick, email: mailNick && domain ? mailNick + '@' + domain : '' };
            }
            if (typ === 'Kursteam') {
                const yearPrefix = String(schemaState.kursteamYearPrefix || '').trim();
                const klasse = normStr(getEl('ssKtKlasse')?.value);
                const fach = normStr(getEl('ssKtFach')?.value);
                const gruppe = normStr(getEl('ssKtGruppe')?.value);
                const displayName = buildKursteamNameFromTemplate(schemaState.kursteamPattern, { yearPrefix, klasse, fach, gruppe });
                const mailNick = buildKursteamMailNickFromTemplate(schemaState.kursteamMailNickPattern, { yearPrefix, klasse, fach, gruppe });
                return { displayName, mailNick, email: mailNick && domain ? mailNick + '@' + domain : '' };
            }
            return { displayName: '', mailNick: '', email: '' };
        }

        function renderAnlegenSchemaUnitUi() {
            const box = getEl('ssAnlegenSchemaUnitBox');
            const p = getEl('ssAnlegenSchemaUnitPreview');
            const wrapJg = getEl('ssAnlegenSchemaJg');
            const wrapArge = getEl('ssAnlegenSchemaArge');
            const wrapKt = getEl('ssAnlegenSchemaKursteam');
            const btnApply = getEl('ssApplyNameSuggestion');
            if (!box || !p || !wrapJg || !wrapArge || !wrapKt) return;

            if (mode !== 'struktur' || !selectedId) {
                box.style.display = 'none';
                return;
            }
            const cur = rowsStruktur.find((r) => String(r.id) === String(selectedId));
            if (!cur) {
                box.style.display = 'none';
                return;
            }
            const typ = String(cur.typ || '');
            const show = typ === 'Jahrgang' || typ === 'Arbeitsgemeinschaft' || typ === 'Kursteam';
            box.style.display = show ? '' : 'none';
            wrapJg.style.display = typ === 'Jahrgang' ? '' : 'none';
            wrapArge.style.display = typ === 'Arbeitsgemeinschaft' ? '' : 'none';
            wrapKt.style.display = typ === 'Kursteam' ? '' : 'none';

            if (!show) return;
            const s = computeAnlegenSuggestionFromInputs();
            const has = !!(s.displayName || s.mailNick);
            p.innerHTML =
                `<div><strong>Bezeichnung (Vorschlag):</strong> <code>${escapeHtml(s.displayName || '–')}</code></div>` +
                `<div style="margin-top:6px;"><strong>Mail‑Nickname (Vorschau):</strong> <code>${escapeHtml(s.mailNick || '–')}</code>` +
                (s.email ? ` &nbsp;→&nbsp; <code>${escapeHtml(s.email)}</code>` : '') +
                `</div>`;
            if (btnApply) btnApply.disabled = !has;
        }

        // Vorschlag übernehmen
        getEl('ssApplyNameSuggestion')?.addEventListener('click', () => {
            if (mode !== 'struktur' || !selectedId) return;
            const sug = computeAnlegenSuggestionFromInputs();
            if (!sug.displayName) return;
            const inp = getEl('ssBezeichnung');
            if (inp) inp.value = sug.displayName;
        });

        // Live-Vorschau bei Eingaben
        ['ssTyp','ssBezeichnung','ssJgYear','ssJgSuffix','ssArgeCode','ssArgeName','ssKtKlasse','ssKtFach','ssKtGruppe','ssPersonName','ssPersonEmail'].forEach((id) => {
            const el = getEl(id);
            if (!el) return;
            const evt = id === 'ssTyp' ? 'change' : 'input';
            el.addEventListener(evt, () => {
                // Bei Typ-Wechsel: wenn Bezeichnung leer ist, Vorschlag einmalig übernehmen
                if (id === 'ssTyp') {
                    const b = getEl('ssBezeichnung');
                    if (b && !normStr(b.value)) {
                        const sug = computeAnlegenSuggestionFromInputs();
                        if (sug.displayName) b.value = sug.displayName;
                    }
                    if (mode === 'struktur' && selectedId) {
                        fillParentSelect(rowsStruktur, selectedId, normStr(getEl('ssTyp')?.value));
                    }
                }
                if (mode === 'struktur' && selectedId) {
                    refreshStrukturTypDependentUi({
                        mode,
                        selectedId,
                        rowsStruktur,
                        structRootDetails: schemaState.structRootDetails,
                        onPersonDetailTabs: applyPersonStructTabsForDetailPanel,
                        personInfoByRole
                    });
                    updateGraphQuickCreateBtn();
                }
                renderAnlegenSchemaUnitUi();
            });
        });
        getEl('ssBtnLoeschen')?.addEventListener('click', async () => {
            if (isStructureTreeRootId(selectedId)) {
                await dlgAlert('Die Hauptäste Schüler:innen, Lehrer:innen und Verwaltung können nicht gelöscht werden.', { title: 'Löschen' });
                return;
            }
            const cur = rowsStruktur.find((r) => String(r.id) === String(selectedId));
            if (!cur) return;
            const label = cur.bezeichnung ? '"' + cur.bezeichnung + '"' : 'diesen Eintrag';
            if (!(await dlgConfirm('Wirklich ' + label + ' löschen? (Unterpunkte werden ebenfalls entfernt.)', { title: 'Löschen', okText: 'Löschen', danger: true })))
                return;
            remove(cur.id);
            selectedId = '';
            rerender();
        });

        getEl('ssBtnExport')?.addEventListener('click', () => {
            downloadJson('schulstruktur-sync.json', { rows: rowsStruktur, memberships, settings: schemaState });
        });
        getEl('ssImportFile')?.addEventListener('change', async (e) => {
            const f = e.target.files && e.target.files[0];
            if (!f) return;
            try {
                const text = await f.text();
                const obj = safeJsonParse(text);
                const nextRows = obj && Array.isArray(obj.rows) ? obj.rows : null;
                if (!nextRows) {
                    await dlgAlert('Import fehlgeschlagen: ungültige Datei (erwartet { "rows": [...] }).', { title: 'Import' });
                    return;
                }
                rowsStruktur = nextRows
                    .filter((r) => r && r.id)
                    .map((r) => ({
                        id: String(r.id),
                        parentId: String(r.parentId || ''),
                        typ: normStr(r.typ) || 'Gruppe',
                        bezeichnung: normStr(r.bezeichnung) || '',
                        beschreibung: normStr(r.beschreibung) || '',
                        schuljahr: normStr(r.schuljahr) || '',
                        status: normStr(r.status) || 'Aktiv',
                        syncStatus: normStr(r.syncStatus) || 'Ausstehend',
                        letzteFehlermeldung: normStr(r.letzteFehlermeldung) || '',
                        jgYear: normStr(r.jgYear) || '',
                        jgSuffix: normStr(r.jgSuffix) || '',
                        argeCode: normStr(r.argeCode) || '',
                        argeName: normStr(r.argeName) || '',
                        ktKlasse: normStr(r.ktKlasse) || '',
                        ktFach: normStr(r.ktFach) || '',
                        ktGruppe: normStr(r.ktGruppe) || '',
                        tenantGroupId: normStr(r.tenantGroupId) || '',
                        tenantMailNickname: normStr(r.tenantMailNickname) || '',
                        tenantTarget: normStr(r.tenantTarget) || '',
                        tenantVisibility: normStr(r.tenantVisibility) || '',
                        tenantUserId: normStr(r.tenantUserId) || '',
                        personName: normStr(r.personName) || '',
                        personEmail: normStr(r.personEmail).toLowerCase() || ''
                    }));
                memberships = (obj && obj.memberships && typeof obj.memberships === 'object') ? obj.memberships : memberships;
                if (obj && obj.settings && typeof obj.settings === 'object') {
                    Object.assign(schemaState, obj.settings);
                    normalizeGraphLayoutModeInSettings(schemaState);
                }
                saveState({ rows: rowsStruktur, memberships, settings: schemaState });
                selectedId = '';
                rerender();
            } catch (err) {
                await dlgAlert('Import fehlgeschlagen: ' + (err?.message || String(err)), { title: 'Import' });
            } finally {
                e.target.value = '';
            }
        });

        // --- Schulstruktur: Tabs (Allgemein/Owner/Mitglieder) + lokale Mitgliedschaft ---
        const sAllgBtn = getEl('ssStructTabAllgemeinBtn');
        const sOwnBtn = getEl('ssStructTabOwnerBtn');
        const sMemBtn = getEl('ssStructTabMitgliederBtn');
        const spAllg = getEl('ssStructTabAllgemein');
        const spOwn = getEl('ssStructTabOwner');
        const spMem = getEl('ssStructTabMitglieder');
        let structActiveTab = 'allg';

        function setStructTab(next) {
            structActiveTab = next === 'own' ? 'own' : next === 'mem' ? 'mem' : 'allg';
            if (sAllgBtn) sAllgBtn.setAttribute('aria-selected', structActiveTab === 'allg' ? 'true' : 'false');
            if (sOwnBtn) sOwnBtn.setAttribute('aria-selected', structActiveTab === 'own' ? 'true' : 'false');
            if (sMemBtn) sMemBtn.setAttribute('aria-selected', structActiveTab === 'mem' ? 'true' : 'false');
            if (spAllg) spAllg.classList.toggle('active', structActiveTab === 'allg');
            if (spOwn) spOwn.classList.toggle('active', structActiveTab === 'own');
            if (spMem) spMem.classList.toggle('active', structActiveTab === 'mem');
        }

        if (sAllgBtn) sAllgBtn.addEventListener('click', () => setStructTab('allg'));
        if (sOwnBtn) sOwnBtn.addEventListener('click', () => setStructTab('own'));
        if (sMemBtn) sMemBtn.addEventListener('click', () => setStructTab('mem'));
        setStructTab('allg');

        function applyPersonStructTabsForDetailPanel(effectiveTyp) {
            const t = String(effectiveTyp || '');
            const hideOwnerMem = t === 'Person';
            const ownBtn = getEl('ssStructTabOwnerBtn');
            const memBtn = getEl('ssStructTabMitgliederBtn');
            const pOwn = getEl('ssStructTabOwner');
            const pMem = getEl('ssStructTabMitglieder');
            [ownBtn, memBtn].forEach((el) => {
                if (!el) return;
                el.style.display = hideOwnerMem ? 'none' : '';
                el.disabled = hideOwnerMem;
                el.setAttribute('aria-hidden', hideOwnerMem ? 'true' : 'false');
            });
            [pOwn, pMem].forEach((el) => {
                if (!el) return;
                el.style.display = hideOwnerMem ? 'none' : '';
                el.setAttribute('aria-hidden', hideOwnerMem ? 'true' : 'false');
                if (hideOwnerMem) el.classList.remove('active');
            });
            if (hideOwnerMem || !t) setStructTab('allg');
        }

        function defaultTenantTargetForType(typ) {
            return defaultTenantTargetForTypeStr(typ);
        }

        function defaultTenantVisibilityForType(typ) {
            return defaultTenantVisibilityForTypeStr(typ);
        }

        function computeTenantCreateSuggestion(row) {
            return computeTenantCreateSuggestionFromRow(row, schemaState);
        }

        function updateStructTenantCreateUi() {
            const wrap = getEl('ssStructTenantCreateWrap');
            if (!wrap) return;
            if (mode !== 'struktur' || !selectedId) {
                wrap.style.display = 'none';
                return;
            }
            let row = rowsStruktur.find((r) => String(r.id) === String(selectedId));
            if (!row && isStructureTreeRootId(selectedId)) {
                row = mergeStructureTreeRootRow(selectedId, schemaState.structRootDetails);
            }
            if (!row) {
                wrap.style.display = 'none';
                return;
            }
            if (String(row.typ || '') === 'Person') {
                wrap.style.display = 'none';
                return;
            }
            wrap.style.display = '';

            const selTarget = getEl('ssStructTenantTarget');
            const selVis = getEl('ssStructTenantVisibility');
            const inpNick = getEl('ssStructTenantMailNick');
            const inpCreated = getEl('ssStructTenantCreatedId');
            const btn = getEl('ssStructTenantCreateBtn');
            const psWrap = getEl('ssStructKursteamPsWrap');
            const psTa = getEl('ssKursteamPsScript');

            if (selTarget && !selTarget.value) selTarget.value = row.tenantTarget || defaultTenantTargetForType(row.typ);
            if (selVis && !selVis.value) selVis.value = row.tenantVisibility || defaultTenantVisibilityForType(row.typ);

            const sug = computeTenantCreateSuggestion(row);
            const nick = sug.mailNick || '';
            if (inpNick) inpNick.value = nick;
            if (inpCreated) inpCreated.value = String(row.tenantGroupId || '');

            const isKursteam = String(row.typ || '') === 'Kursteam';
            if (psWrap) psWrap.style.display = isKursteam ? '' : 'none';
            if (btn) {
                // Kursteam: prefer PS/CSV instead of online create
                btn.style.display = isKursteam ? 'none' : '';
                btn.disabled = !sug.displayName || !nick || !!row.tenantGroupId;
            }
            if (psTa && isKursteam) {
                psTa.value = buildKursteamProvisionScript('kursteams.csv');
            }
        }

        function updateGraphQuickCreateBtn() {
            const btn = getEl('ssBtnGraphCreate');
            if (!btn) return;
            if (mode !== 'struktur' || !selectedId) {
                btn.style.display = 'none';
                return;
            }
            let row = rowsStruktur.find((r) => String(r.id) === String(selectedId));
            if (!row && isStructureTreeRootId(selectedId)) {
                row = mergeStructureTreeRootRow(selectedId, schemaState.structRootDetails);
            }
            if (!row || row.isStructureTreeRoot) {
                btn.style.display = 'none';
                return;
            }
            const t = String(row.typ || '');
            if (t === 'Person') {
                btn.style.display = '';
                btn.innerHTML = '<i class="bi bi-person-plus"></i>Benutzer anlegen (Graph)';
                const next = readDetailToRow(row);
                const upn = normStr(next.personEmail).toLowerCase();
                const dn = normStr(next.personName) || normStr(next.bezeichnung);
                const hasUpn = upn.includes('@');
                btn.disabled = !hasUpn || !dn || !!normStr(next.tenantUserId);
                return;
            }
            if (t === 'Gruppe') {
                btn.style.display = '';
                btn.innerHTML = '<i class="bi bi-people"></i>M365‑Gruppe anlegen (Graph)';
                const next = readDetailToRow(row);
                const sug = computeTenantCreateSuggestion(next);
                const nick = String(getEl('ssStructTenantMailNick')?.value || sug.mailNick || '').trim();
                btn.disabled = !sug.displayName || !nick || !!normStr(next.tenantGroupId);
                return;
            }
            btn.style.display = 'none';
        }

        async function createSelectedPersonUserInTenant() {
            if (mode !== 'struktur' || !selectedId) return;
            const cur = rowsStruktur.find((r) => String(r.id) === String(selectedId));
            if (!cur || String(cur.typ || '') !== 'Person') return;
            const next = readDetailToRow(cur);
            const displayName = normStr(next.personName) || normStr(next.bezeichnung);
            const upn = normStr(next.personEmail).toLowerCase();
            if (!displayName) {
                await dlgAlert('Bitte einen Anzeigenamen im Feld „Name“ eingeben (und speichern oder hier direkt ausfüllen).', { title: 'Eingabe' });
                return;
            }
            if (!upn || upn.indexOf('@') === -1) {
                await dlgAlert('Bitte eine gültige UPN/E‑Mail im Feld „E‑Mail / UPN“ eingeben.', { title: 'Eingabe' });
                return;
            }
            if (normStr(next.tenantUserId)) {
                await dlgAlert('Für diesen Eintrag ist bereits eine Entra-Benutzer-ID gespeichert.', { title: 'Entra' });
                return;
            }
            const mailNick = mailNicknameFromUpn(upn);
            const pwd = generateGraphTempPassword();
            if (
                !(await dlgConfirm(
                    'Benutzer in Entra ID anlegen (Microsoft Graph)?\n\n' +
                        'Anzeigename: ' +
                        displayName +
                        '\nUPN: ' +
                        upn +
                        '\nMail‑Nickname: ' +
                        mailNick +
                        '\n\nEs wird ein temporäres Kennwort gesetzt (Wechsel beim ersten Anmelden).',
                    { title: 'Entra', okText: 'Anlegen' }
                ))
            ) {
                return;
            }
            const btn = getEl('ssBtnGraphCreate');
            if (btn) btn.disabled = true;
            try {
                setTenantProgress(true, 'Benutzer wird angelegt …', 0.25);
                const token = await getGraphToken(GRAPH_SCOPES_GRAPH_OBJECT_CREATE);
                const body = {
                    accountEnabled: true,
                    displayName,
                    mailNickname: mailNick,
                    userPrincipalName: upn,
                    passwordProfile: {
                        forceChangePasswordNextSignIn: true,
                        password: pwd
                    }
                };
                const created = await graphJson('POST', '/users', token, body, undefined);
                const uid = String(created.id || '').trim();
                if (!uid) throw new Error('Keine Benutzer-ID von Graph erhalten.');
                next.personName = displayName;
                next.personEmail = upn;
                next.tenantUserId = uid;
                next.syncStatus = 'Ok';
                next.letzteFehlermeldung = '';
                upsert(next);
                saveState({ rows: rowsStruktur, memberships, settings: schemaState });
                setTenantProgress(true, 'Benutzer angelegt.', 1);
                setTimeout(() => setTenantProgress(false, '', null), 1600);
                try {
                    await dlgAlert(
                        'Benutzer wurde angelegt.\n\nBitte das temporäre Startkennwort sicher weitergeben (einmalige Anzeige):\n\n' + pwd,
                        { title: 'Kennwort notieren', okText: 'Verstanden' }
                    );
                } catch {
                    // ignore
                }
                rerender();
            } catch (e) {
                const msg = e?.message || String(e);
                next.syncStatus = 'Fehler';
                next.letzteFehlermeldung = msg;
                upsert(next);
                saveState({ rows: rowsStruktur, memberships, settings: schemaState });
                setTenantProgress(true, 'Anlegen fehlgeschlagen: ' + msg, null);
                rerender();
            } finally {
                if (btn) btn.disabled = false;
                updateGraphQuickCreateBtn();
            }
        }

        async function createSelectedStructInTenant() {
            if (mode !== 'struktur' || !selectedId) return;
            let row = rowsStruktur.find((r) => String(r.id) === String(selectedId));
            if (!row && isStructureTreeRootId(selectedId)) {
                row = mergeStructureTreeRootRow(selectedId, schemaState.structRootDetails);
            }
            if (!row) return;
            if (String(row.typ || '') === 'Person') {
                await dlgAlert('Für den Typ „Person“ wird hier kein Microsoft‑365‑Gruppen‑ oder Team‑Objekt angelegt (Person = Benutzerkonto in Entra ID).', {
                    title: 'Hinweis'
                });
                return;
            }
            const btn = getEl('ssStructTenantCreateBtn');
            const graphQuickBtn = getEl('ssBtnGraphCreate');
            if (btn) btn.disabled = true;
            if (graphQuickBtn) graphQuickBtn.disabled = true;
            try {
                const target = normStr(getEl('ssStructTenantTarget')?.value) || row.tenantTarget || defaultTenantTargetForType(row.typ);
                const vis = normStr(getEl('ssStructTenantVisibility')?.value) || row.tenantVisibility || defaultTenantVisibilityForType(row.typ);
                const sug = computeTenantCreateSuggestion(row);
                const nick = String(getEl('ssStructTenantMailNick')?.value || sug.mailNick || '').trim();
                if (!sug.displayName) throw new Error('Bitte zuerst eine Bezeichnung eingeben.');
                if (!nick) throw new Error('Mail‑Nickname ist leer.');
                if (
                    !(await dlgConfirm(
                        `Im Tenant anlegen?\n\nName: ${sug.displayName}\nMailNick: ${nick}\nZiel: ${target === 'team' ? 'Team' : 'Gruppe'}\nSichtbarkeit: ${vis}`,
                        { title: 'Tenant-Anlage', okText: 'Anlegen' }
                    ))
                ) {
                    return;
                }

                setTenantProgress(true, 'Anlegen im Tenant …', 0.15);
                const descForGroup = normStr(getEl('ssBeschreibung')?.value);
                const created = await createUnifiedGroup(sug.displayName, descForGroup, nick, vis);
                const gid = String(created.id || '').trim();
                if (!gid) throw new Error('Anlegen fehlgeschlagen: keine Gruppen-ID erhalten.');

                setTenantProgress(true, 'Angelegt. Übernehme Owner/Mitglieder …', 0.45);
                if (target === 'team') {
                    setTenantProgress(true, 'Team wird erstellt …', 0.55);
                    await createTeamForGroup(gid);
                }

                const mem = memberships[String(row.id)] || { owners: [], members: [] };
                const ownerIds = (mem.owners || []).map((p) => String(p.id || '')).filter(Boolean);
                const memberIds = (mem.members || []).map((p) => String(p.id || '')).filter(Boolean);

                if (ownerIds.length) await mapWithConcurrency(ownerIds, 4, async (uid) => await addOwnerWithMemberFallback(gid, uid));
                if (memberIds.length) await mapWithConcurrency(memberIds, 6, async (uid) => await addGroupMember(gid, uid));

                row.tenantGroupId = gid;
                row.tenantMailNickname = nick;
                row.tenantTarget = target;
                row.tenantVisibility = vis;
                row.beschreibung = descForGroup;
                row.syncStatus = 'Ok';
                row.letzteFehlermeldung = '';
                if (row.isStructureTreeRoot) {
                    if (!schemaState.structRootDetails || typeof schemaState.structRootDetails !== 'object') {
                        schemaState.structRootDetails = {};
                    }
                    schemaState.structRootDetails[String(row.id)] = pickStorableStructureTreeRootFields(row);
                }
                saveState({ rows: rowsStruktur, memberships, settings: schemaState });

                links[String(row.id)] = {
                    tenantGroupId: gid,
                    note: 'Auto: im Tenant angelegt',
                    updatedAt: new Date().toISOString()
                };
                links = saveMatchState(links);

                setTenantProgress(true, 'Fertig: im Tenant angelegt.', 1);
                setTimeout(() => setTenantProgress(false, '', null), 1600);
                rerender();
            } catch (e) {
                const msg = e?.message || String(e);
                row.syncStatus = 'Fehler';
                row.letzteFehlermeldung = msg;
                if (row.isStructureTreeRoot) {
                    if (!schemaState.structRootDetails || typeof schemaState.structRootDetails !== 'object') {
                        schemaState.structRootDetails = {};
                    }
                    schemaState.structRootDetails[String(row.id)] = pickStorableStructureTreeRootFields(row);
                }
                saveState({ rows: rowsStruktur, memberships, settings: schemaState });
                setTenantProgress(true, 'Anlegen fehlgeschlagen: ' + msg, null);
                rerender();
            } finally {
                if (btn) btn.disabled = false;
                if (graphQuickBtn) graphQuickBtn.disabled = false;
                updateGraphQuickCreateBtn();
            }
        }

        getEl('ssStructTenantCreateBtn')?.addEventListener('click', createSelectedStructInTenant);
        ['ssStructTenantTarget', 'ssStructTenantVisibility'].forEach((id) => {
            const el = getEl(id);
            if (!el) return;
            el.addEventListener('change', () => {
                const cur = rowsStruktur.find((r) => String(r.id) === String(selectedId));
                if (!cur) return;
                cur.tenantTarget = normStr(getEl('ssStructTenantTarget')?.value);
                cur.tenantVisibility = normStr(getEl('ssStructTenantVisibility')?.value);
                saveState({ rows: rowsStruktur, memberships, settings: schemaState });
                updateStructTenantCreateUi();
                updateGraphQuickCreateBtn();
            });
        });
        getEl('ssStructTenantMailNick')?.addEventListener('input', () => {
            if (mode === 'struktur' && selectedId) updateGraphQuickCreateBtn();
        });

        function getSelectedKursteamRows() {
            if (!selectedId) return [];
            const row = rowsStruktur.find((r) => String(r.id) === String(selectedId));
            if (!row) return [];
            if (String(row.typ || '') !== 'Kursteam') return [];
            return [row];
        }

        getEl('ssKursteamCsvDownload')?.addEventListener('click', async () => {
            const rows = getSelectedKursteamRows();
            if (!rows.length) {
                await dlgAlert('Bitte zuerst ein Kursteam auswählen.', { title: 'Kursteam' });
                return;
            }
            // ensure latest suggestion/visibility is written to row fields
            const cur = rows[0];
            cur.tenantVisibility = normStr(getEl('ssStructTenantVisibility')?.value) || cur.tenantVisibility || 'HiddenMembership';
            cur.tenantTarget = normStr(getEl('ssStructTenantTarget')?.value) || cur.tenantTarget || 'team';
            saveState({ rows: rowsStruktur, memberships, settings: schemaState });

            const csv = buildKursteamCsv(rows, memberships, schemaState);
            downloadText('kursteams.csv', csv);
        });

        getEl('ssKursteamPsCopy')?.addEventListener('click', async () => {
            const ta = getEl('ssKursteamPsScript');
            const text = ta ? String(ta.value || '') : '';
            if (!text) return;
            try {
                await navigator.clipboard.writeText(text);
            } catch {
                // ignore
            }
        });

        getEl('ssKursteamPsDownload')?.addEventListener('click', () => {
            const ta = getEl('ssKursteamPsScript');
            const text = ta ? String(ta.value || '') : '';
            if (!text) return;
            downloadText('kursteams-provision.ps1', text);
        });

        function renderAnlegenSettingsPreview() {
            const el = getEl('ssSchemaPreview');
            if (!el) return;
            const domain = String(schemaState.domain || '').trim() || '…';
            const yearPrefix = String(schemaState.kursteamYearPrefix || '').trim() || '…';
            const kurTpl = String(schemaState.kursteamPattern || '').trim();
            const kurNickTpl = String(schemaState.kursteamMailNickPattern || '').trim();
            const kurName = buildKursteamNameFromTemplate(kurTpl, { yearPrefix, klasse: '1AK', fach: 'D', gruppe: 'G1' });
            const kurNick = buildKursteamMailNickFromTemplate(kurNickTpl, { yearPrefix, klasse: '1AK', fach: 'D', gruppe: 'G1' });

            const jgNick = buildJgMailNick(schemaState, '2030', 'AK');
            const argeNick = buildArgeMailNick(schemaState, 'M');

            const bK = getEl('ssSchemaTabKursteamBtn');
            const bJ = getEl('ssSchemaTabJahrgangBtn');
            const bA = getEl('ssSchemaTabArgeBtn');
            const active =
                bJ && bJ.getAttribute('aria-selected') === 'true'
                    ? 'jg'
                    : bA && bA.getAttribute('aria-selected') === 'true'
                      ? 'arge'
                      : 'kt';

            if (active === 'jg') {
                el.innerHTML =
                    `<div><strong>Jahrgang (Mail‑Nickname):</strong> <code>${escapeHtml(jgNick)}</code> &nbsp;→&nbsp; <code>${escapeHtml(jgNick)}@${escapeHtml(domain)}</code></div>`;
            } else if (active === 'arge') {
                el.innerHTML =
                    `<div><strong>ARGE (Mail‑Nickname):</strong> <code>${escapeHtml(argeNick)}</code> &nbsp;→&nbsp; <code>${escapeHtml(argeNick)}@${escapeHtml(domain)}</code></div>`;
            } else {
                el.innerHTML =
                    `<div><strong>Kursteam (Anzeige):</strong> <code>${escapeHtml(kurName)}</code></div>` +
                    `<div style="margin-top:6px;"><strong>Kursteam (Mail‑Nickname):</strong> <code>${escapeHtml(kurNick)}</code> &nbsp;→&nbsp; <code>${escapeHtml(kurNick)}@${escapeHtml(domain)}</code></div>`;
            }
        }

        function bindAnlegenSettingsUi() {
            const panel = getEl('ssAnlegenSettingsPanel');
            if (!panel) return;
            const domainEl = getEl('ssSchemaDomain');
            const ypEl = getEl('ssSchemaSchoolYearPrefix');
            const kurTplEl = getEl('ssSchemaKursteamPattern');
            const kurNickTplEl = getEl('ssSchemaKursteamMailNickPattern');
            const jgPrefEl = getEl('ssSchemaJgPrefix');
            const jgUpperEl = getEl('ssSchemaJgUpper');
            const argePrefEl = getEl('ssSchemaArgePrefix');
            const argeUpperEl = getEl('ssSchemaArgeUpper');

            if (domainEl) domainEl.value = String(schemaState.domain || '');
            if (ypEl) ypEl.value = String(schemaState.kursteamYearPrefix || '');
            if (kurTplEl) kurTplEl.value = String(schemaState.kursteamPattern || '');
            if (kurNickTplEl) kurNickTplEl.value = String(schemaState.kursteamMailNickPattern || '');
            if (jgPrefEl) jgPrefEl.value = String(schemaState.jgPrefix || '');
            if (jgUpperEl) jgUpperEl.checked = !!schemaState.jgUpper;
            if (argePrefEl) argePrefEl.value = String(schemaState.argePrefix || '');
            if (argeUpperEl) argeUpperEl.checked = !!schemaState.argeUpper;

            let t;
            const saveNow = () => saveState({ rows: rowsStruktur, memberships, settings: schemaState });
            const onChange = () => {
                clearTimeout(t);
                t = setTimeout(() => {
                    if (domainEl) schemaState.domain = String(domainEl.value || '').trim();
                    if (ypEl) schemaState.kursteamYearPrefix = String(ypEl.value || '').trim();
                    if (kurTplEl) schemaState.kursteamPattern = String(kurTplEl.value || '').trim();
                    if (kurNickTplEl) schemaState.kursteamMailNickPattern = String(kurNickTplEl.value || '').trim();
                    if (jgPrefEl) schemaState.jgPrefix = String(jgPrefEl.value || '').trim();
                    if (jgUpperEl) schemaState.jgUpper = !!jgUpperEl.checked;
                    if (argePrefEl) schemaState.argePrefix = String(argePrefEl.value || '').trim();
                    if (argeUpperEl) schemaState.argeUpper = !!argeUpperEl.checked;
                    saveNow();
                    renderAnlegenSettingsPreview();
                }, 120);
            };

            [domainEl, ypEl, kurTplEl, kurNickTplEl, jgPrefEl, argePrefEl].forEach((x) => x && x.addEventListener('input', onChange));
            [jgUpperEl, argeUpperEl].forEach((x) => x && x.addEventListener('change', onChange));
            renderAnlegenSettingsPreview();
        }

        // Tabs im Schema-Panel (Kursteams/Jahrgang/ARGEs)
        let schemaTabWired = false;
        function wireSchemaTabsOnce() {
            if (schemaTabWired) return;
            schemaTabWired = true;
            const bK = getEl('ssSchemaTabKursteamBtn');
            const bJ = getEl('ssSchemaTabJahrgangBtn');
            const bA = getEl('ssSchemaTabArgeBtn');
            const pK = getEl('ssSchemaTabKursteam');
            const pJ = getEl('ssSchemaTabJahrgang');
            const pA = getEl('ssSchemaTabArge');
            if (!bK || !bJ || !bA || !pK || !pJ || !pA) return;

            function setTab(which) {
                const w = which === 'jg' ? 'jg' : which === 'arge' ? 'arge' : 'kt';
                bK.setAttribute('aria-selected', w === 'kt' ? 'true' : 'false');
                bJ.setAttribute('aria-selected', w === 'jg' ? 'true' : 'false');
                bA.setAttribute('aria-selected', w === 'arge' ? 'true' : 'false');
                pK.classList.toggle('active', w === 'kt');
                pJ.classList.toggle('active', w === 'jg');
                pA.classList.toggle('active', w === 'arge');
                renderAnlegenSettingsPreview();
            }

            bK.addEventListener('click', () => setTab('kt'));
            bJ.addEventListener('click', () => setTab('jg'));
            bA.addEventListener('click', () => setTab('arge'));
            setTab('kt');
        }

        function getStructMembership(unitId) {
            const id = String(unitId || '');
            if (!id) return { owners: [], members: [] };
            const m = memberships[id];
            if (m && typeof m === 'object') {
                return {
                    owners: Array.isArray(m.owners) ? m.owners : [],
                    members: Array.isArray(m.members) ? m.members : []
                };
            }
            return { owners: [], members: [] };
        }

        function setStructMembership(unitId, next) {
            const id = String(unitId || '');
            if (!id) return;
            memberships[id] = {
                owners: Array.isArray(next.owners) ? next.owners : [],
                members: Array.isArray(next.members) ? next.members : []
            };
            saveState({ rows: rowsStruktur, memberships, settings: schemaState });
        }

        function renderStructPeople(list, targetId, removeAttr) {
            const wrap = getEl(targetId);
            if (!wrap) return;
            wrap.replaceChildren();
            const arr = Array.isArray(list) ? list : [];
            if (!arr.length) {
                const p = document.createElement('p');
                p.style.margin = '0';
                p.style.color = '#6c757d';
                p.textContent = 'Keine Einträge.';
                wrap.appendChild(p);
                return;
            }
            arr
                .slice()
                .sort((a, b) => compareDe(personLabel(a), personLabel(b)))
                .forEach((u) => {
                    const row = document.createElement('div');
                    row.style.display = 'flex';
                    row.style.justifyContent = 'space-between';
                    row.style.alignItems = 'flex-start';
                    row.style.gap = '10px';
                    row.style.padding = '8px 0';
                    row.style.borderBottom = '1px solid #e9ecef';
                    const txt = document.createElement('div');
                    txt.style.whiteSpace = 'pre-wrap';
                    txt.style.lineHeight = '1.35';
                    txt.style.fontSize = '0.92em';
                    txt.textContent = personLabel(u) || '–';
                    const btn = document.createElement('button');
                    btn.type = 'button';
                    btn.className = 'btn';
                    btn.style.padding = '6px 10px';
                    btn.style.fontSize = '0.85em';
                    btn.textContent = 'Entfernen';
                    btn.setAttribute(removeAttr, String(u.id || u.userPrincipalName || u.mail || personLabel(u)));
                    row.appendChild(txt);
                    row.appendChild(btn);
                    wrap.appendChild(row);
                });
        }

        function fillStructSearchSelect(users, selectId) {
            const sel = getEl(selectId);
            if (!sel) return;
            sel.replaceChildren();
            if (!users || !users.length) {
                const opt = document.createElement('option');
                opt.value = '';
                opt.textContent = '(keine Treffer)';
                sel.appendChild(opt);
                return;
            }
            for (let i = 0; i < users.length; i++) {
                const u = users[i];
                const opt = document.createElement('option');
                opt.value = u.id || '';
                opt.textContent = personLabel(u) || (u.id ? String(u.id) : '');
                sel.appendChild(opt);
            }
        }

        function renderStructMembershipUi() {
            if (mode !== 'struktur' || !selectedId) return;
            const m = getStructMembership(selectedId);
            renderStructPeople(m.owners, 'ssStructOwnersList', 'data-ss-struct-remove-owner');
            renderStructPeople(m.members, 'ssStructMembersList', 'data-ss-struct-remove-member');
        }

        async function runStructUserSearch(query, selectId) {
            const q = String(query || '').trim();
            if (!q) {
                fillStructSearchSelect([], selectId);
                return;
            }
            try {
                const token = await getGraphToken(GRAPH_SCOPES_TENANT_OWNER_MANAGE);
                const users = await graphSearchUsersForOwner(token, q);
                fillStructSearchSelect(users, selectId);
            } catch (e) {
                toast('Suche: ' + (e?.message || String(e)));
                fillStructSearchSelect([], selectId);
            }
        }

        getEl('ssStructOwnerSearchBtn')?.addEventListener('click', async () => {
            if (mode !== 'struktur') return;
            await runStructUserSearch(getEl('ssStructOwnerSearch')?.value || '', 'ssStructOwnerSearchResults');
        });
        getEl('ssStructMemberSearchBtn')?.addEventListener('click', async () => {
            if (mode !== 'struktur') return;
            await runStructUserSearch(getEl('ssStructMemberSearch')?.value || '', 'ssStructMemberSearchResults');
        });

        function addStructPerson(kind, user) {
            if (!selectedId) return;
            const m = getStructMembership(selectedId);
            const arr = kind === 'owners' ? m.owners : m.members;
            const id = String(user.id || '');
            if (id && arr.some((x) => String(x.id || '') === id)) return;
            arr.push(user);
            if (kind === 'owners') m.owners = arr;
            else m.members = arr;
            setStructMembership(selectedId, m);
            renderStructMembershipUi();
        }

        getEl('ssStructOwnerAddBtn')?.addEventListener('click', () => {
            if (mode !== 'struktur' || !selectedId) return;
            const sel = getEl('ssStructOwnerSearchResults');
            const userId = sel && sel.value ? String(sel.value).trim() : '';
            if (!userId) return toast('Bitte zuerst einen Benutzer aus den Treffern auswählen.');
            // Minimal speichern: id + displayName/upn (falls vorhanden)
            const opt = sel.options[sel.selectedIndex];
            addStructPerson('owners', { id: userId, displayName: opt ? String(opt.textContent || '') : '' });
        });
        getEl('ssStructMemberAddBtn')?.addEventListener('click', () => {
            if (mode !== 'struktur' || !selectedId) return;
            const sel = getEl('ssStructMemberSearchResults');
            const userId = sel && sel.value ? String(sel.value).trim() : '';
            if (!userId) return toast('Bitte zuerst einen Benutzer aus den Treffern auswählen.');
            const opt = sel.options[sel.selectedIndex];
            addStructPerson('members', { id: userId, displayName: opt ? String(opt.textContent || '') : '' });
        });

        getEl('ssStructOwnersList')?.addEventListener('click', (ev) => {
            if (mode !== 'struktur' || !selectedId) return;
            const t = ev.target;
            const btn = t && t.closest ? t.closest('button[data-ss-struct-remove-owner]') : null;
            if (!btn) return;
            const key = btn.getAttribute('data-ss-struct-remove-owner') || '';
            const m = getStructMembership(selectedId);
            m.owners = m.owners.filter((u) => String(u.id || u.displayName || '') !== String(key));
            setStructMembership(selectedId, m);
            renderStructMembershipUi();
        });
        getEl('ssStructMembersList')?.addEventListener('click', (ev) => {
            if (mode !== 'struktur' || !selectedId) return;
            const t = ev.target;
            const btn = t && t.closest ? t.closest('button[data-ss-struct-remove-member]') : null;
            if (!btn) return;
            const key = btn.getAttribute('data-ss-struct-remove-member') || '';
            const m = getStructMembership(selectedId);
            m.members = m.members.filter((u) => String(u.id || u.displayName || '') !== String(key));
            setStructMembership(selectedId, m);
            renderStructMembershipUi();
        });

        getEl('ssBtnTenantLoadTop')?.addEventListener('click', async () => {
            // Wenn ein Login-Redirect passiert, wollen wir nach Rückkehr im "Verwalten"-Tab bleiben.
            try {
                sessionStorage.setItem(UI_MODE_KEY, 'tenant');
            } catch {
                // ignore
            }
            await reloadTenantNow('Starte');
        });

        // --- Bulk Owner add (Tenant) ---
        function fillBulkOwnerSearchSelect(users) {
            const sel = getEl('ssTenantBulkOwnerResults');
            if (!sel) return;
            sel.replaceChildren();
            if (!users || !users.length) {
                const opt = document.createElement('option');
                opt.value = '';
                opt.textContent = '(keine Treffer)';
                sel.appendChild(opt);
                return;
            }
            for (let i = 0; i < users.length; i++) {
                const u = users[i];
                const opt = document.createElement('option');
                opt.value = u.id || '';
                opt.textContent = personLabel(u) || (u.id ? String(u.id) : '');
                sel.appendChild(opt);
            }
        }

        async function mapWithConcurrency(items, limit, fn) {
            const results = new Array(items.length);
            let i = 0;
            async function worker() {
                while (i < items.length) {
                    const idx = i++;
                    results[idx] = await fn(items[idx], idx);
                }
            }
            const n = Math.min(limit, items.length || 1);
            const workers = [];
            for (let w = 0; w < n; w++) workers.push(worker());
            await Promise.all(workers);
            return results;
        }

        getEl('ssTenantBulkClear')?.addEventListener('click', () => {
            if (mode !== 'tenant') return;
            tenantMultiSel = new Set();
            selectedId = '';
            rerender();
        });

        getEl('ssTenantBulkOwnerSearchBtn')?.addEventListener('click', async () => {
            if (mode !== 'tenant') return;
            const q = getEl('ssTenantBulkOwnerSearch')?.value || '';
            const btn = getEl('ssTenantBulkOwnerSearchBtn');
            if (btn) btn.disabled = true;
            try {
                const token = await getGraphToken(GRAPH_SCOPES_TENANT_OWNER_MANAGE);
                const users = await graphSearchUsersForOwner(token, q);
                fillBulkOwnerSearchSelect(users);
                toast('Suche: ' + users.length + ' Treffer.');
            } catch (e) {
                toast('Suche: ' + (e?.message || String(e)));
            } finally {
                if (btn) btn.disabled = false;
            }
        });

        getEl('ssTenantBulkAddOwnerBtn')?.addEventListener('click', async () => {
            if (mode !== 'tenant') return;
            if (tenantMultiSel.size < 2) {
                toast('Bitte mindestens 2 Gruppen/Teams auswählen.');
                return;
            }
            const sel = getEl('ssTenantBulkOwnerResults');
            const userId = sel && sel.value ? String(sel.value).trim() : '';
            if (!userId) {
                toast('Bitte zuerst einen Owner aus den Treffern auswählen.');
                return;
            }
            if (!(await dlgConfirm('Owner wirklich zu ALLEN ausgewählten Gruppen/Teams hinzufügen?', { title: 'Bulk-Owner', okText: 'Hinzufügen' }))) return;

            const ids = Array.from(tenantMultiSel);
            const btn = getEl('ssTenantBulkAddOwnerBtn');
            if (btn) btn.disabled = true;
            try {
                setTenantProgress(true, 'Bulk: Owner wird gesetzt … 0 / ' + ids.length, 0.1);
                let ok = 0;
                let fail = 0;
                await mapWithConcurrency(ids, 4, async (gid, idx) => {
                    try {
                        await addOwnerWithMemberFallback(gid, userId);
                        ok++;
                        const ix = rowsTenant.findIndex((r) => String(r.id) === String(gid));
                        if (ix !== -1) {
                            const prev = rowsTenant[ix].ownerCount;
                            const next =
                                typeof prev === 'number' && prev >= 0 ? Math.max(1, prev + 1) : 1;
                            rowsTenant[ix] = Object.assign({}, rowsTenant[ix], { ownerCount: next });
                        }
                    } catch {
                        fail++;
                    }
                    if (idx % 2 === 0) {
                        const ratio = Math.min(0.95, (idx + 1) / ids.length);
                        setTenantProgress(true, 'Bulk: Owner wird gesetzt … ' + (idx + 1) + ' / ' + ids.length, ratio);
                    }
                });
                try {
                    saveTenantCache(rowsTenant);
                } catch (_) {}
                try {
                    rerender();
                } catch (_) {}
                setTenantProgress(true, 'Bulk fertig. OK: ' + ok + ', Fehler: ' + fail, 1);
                setTimeout(() => setTenantProgress(false, '', null), 1800);
            } catch (e) {
                setTenantProgress(true, 'Bulk fehlgeschlagen: ' + (e?.message || String(e)), null);
            } finally {
                if (btn) btn.disabled = false;
            }
        });

        // --- Abgleichen (Mapping) ---
        function saveLinkForSelected(selectValue, noteText) {
            if (!selectedId) return;
            const raw = String(selectValue || '').trim();
            let gid = '';
            let uid = '';
            if (raw.startsWith('u:')) uid = raw.slice(2).trim();
            else if (raw.startsWith('g:')) gid = raw.slice(2).trim();
            else gid = raw;
            links = saveMatchLinkPublic(String(selectedId), gid, noteText || '', uid);
        }

        function suggestTenantGroupForUnit(unit) {
            return suggestTenantGroupForUnitFromList(unit, rowsTenant || []);
        }

        getEl('ssMatchSaveBtn')?.addEventListener('click', () => {
            if (mode !== 'match' || !selectedId) return;
            const selTenant = getEl('ssMatchTenantGroup');
            const note = getEl('ssMatchNote');
            const val = selTenant && selTenant.value ? String(selTenant.value) : '';
            saveLinkForSelected(val, note ? note.value : '');
            toast(val ? 'Verknüpft.' : 'Verknüpfung gelöst.');
        });
        getEl('ssMatchClearBtn')?.addEventListener('click', () => {
            if (mode !== 'match' || !selectedId) return;
            const selTenant = getEl('ssMatchTenantGroup');
            const note = getEl('ssMatchNote');
            const inpSearch = getEl('ssMatchTenantSearch');
            if (inpSearch) inpSearch.value = '';
            if (selTenant) {
                rebuildMatchTenantSelectOptions(selTenant, '', '');
                selTenant.value = '';
            }
            if (note) note.value = '';
            saveLinkForSelected('', '');
            toast('Verknüpfung gelöst.');
        });
        getEl('ssMatchSuggestBtn')?.addEventListener('click', () => {
            if (mode !== 'match' || !selectedId) return;
            const unit = rowsStruktur.find((r) => String(r.id) === String(selectedId));
            const cache = loadTenantCache();
            const sug = suggestTenantMatchSelectValue(unit, rowsTenant || [], cache.users || []);
            const selTenant = getEl('ssMatchTenantGroup');
            const inpSearch = getEl('ssMatchTenantSearch');
            if (inpSearch) inpSearch.value = '';
            if (selTenant) {
                rebuildMatchTenantSelectOptions(selTenant, '', sug || '');
                selTenant.value = sug || '';
            }
            toast(sug ? 'Vorschlag gesetzt.' : 'Kein passender Vorschlag gefunden.');
        });

        getEl('ssTenantArchiveState')?.addEventListener('change', () => {
            const sel = getEl('ssTenantArchiveState');
            const spo = getEl('ssTenantArchiveSpoReadonly');
            if (!sel || !spo) return;
            spo.disabled = sel.disabled || String(sel.value || '') !== 'archived';
            if (String(sel.value || '') !== 'archived') spo.checked = false;
        });

        getEl('ssTenantReloadBtn')?.addEventListener('click', async () => {
            if (mode !== 'tenant' || !selectedId) return;
            const btn = getEl('ssTenantReloadBtn');
            if (btn) btn.disabled = true;
            try {
                setTenantProgress(true, 'Lade Gruppendetails …', 0.25);
                const fresh = await fetchTenantGroupDetail(selectedId);
                // Update cache row
                const idx = rowsTenant.findIndex((r) => String(r.id) === String(selectedId));
                if (idx !== -1) rowsTenant[idx] = Object.assign({}, rowsTenant[idx], fresh);
                saveTenantCache(rowsTenant);
                setTenantProgress(true, 'Details aktualisiert.', 1);
                setTimeout(() => setTenantProgress(false, '', null), 1200);
                rerender();
            } catch (e) {
                setTenantProgress(true, 'Neu laden: ' + (e?.message || String(e)), null);
            } finally {
                if (btn) btn.disabled = false;
            }
        });

        getEl('ssTenantUpdateBtn')?.addEventListener('click', async () => {
            if (mode !== 'tenant' || !selectedId) return;
            const btn = getEl('ssTenantUpdateBtn');
            const btnReload = getEl('ssTenantReloadBtn');
            const btnRenew = getEl('ssTenantRenewBtn');
            const btnDelete = getEl('ssTenantDeleteBtn');
            const archSel = getEl('ssTenantArchiveState');
            const archSpo = getEl('ssTenantArchiveSpoReadonly');
            const name = getEl('ssTenantName')?.value || '';
            const desc = getEl('ssTenantDescription')?.value || '';
            const vis = getEl('ssTenantVisibility')?.value || '';
            const idxPre = rowsTenant.findIndex((r) => String(r.id) === String(selectedId));
            const rowPre = idxPre === -1 ? null : rowsTenant[idxPre];
            const baselineArchived =
                rowPre && (rowPre.teamIsArchived === true || rowPre.teamIsArchived === false) ? rowPre.teamIsArchived : null;
            const wantArchived =
                archSel && !archSel.disabled && String(archSel.value || '') === 'archived'
                    ? true
                    : archSel && !archSel.disabled
                      ? false
                      : null;
            const typPre = String(rowPre && rowPre.typ ? rowPre.typ : '');
            const unifyArch = typPre === 'Team' || typPre === 'Gruppe';
            const hasTeamForMutation =
                rowPre &&
                (rowPre.hasTeamsForArchive === true ||
                    (rowPre.hasTeamsForArchive === undefined &&
                        (rowPre.teamIsArchived === true || rowPre.teamIsArchived === false)));
            const doArchiveMutation =
                rowPre &&
                unifyArch &&
                !!hasTeamForMutation &&
                wantArchived !== null &&
                baselineArchived !== null &&
                wantArchived !== baselineArchived;
            const spoForArchive = !!(wantArchived && archSpo && archSpo.checked);
            if (btn) btn.disabled = true;
            if (btnReload) btnReload.disabled = true;
            if (btnRenew) btnRenew.disabled = true;
            if (btnDelete) btnDelete.disabled = true;
            if (archSel) archSel.disabled = true;
            if (archSpo) archSpo.disabled = true;
            try {
                if (!(await dlgConfirm('Änderungen im LIVE‑Tenant wirklich speichern?', { title: 'Tenant-Update', okText: 'Speichern' }))) return;
                setTenantProgress(true, 'Update wird durchgeführt …', 0.35);
                await updateTenantGroup(selectedId, name, desc, vis);
                if (doArchiveMutation) {
                    setTenantProgress(
                        true,
                        wantArchived ? 'Team wird archiviert …' : 'Archivierung wird aufgehoben …',
                        0.5
                    );
                    await setTenantTeamArchiveState(selectedId, wantArchived, spoForArchive);
                }
                setTenantProgress(true, 'Update OK – lade Details neu …', 0.75);
                const fresh = await fetchTenantGroupDetail(selectedId);
                const idx = rowsTenant.findIndex((r) => String(r.id) === String(selectedId));
                if (idx !== -1) rowsTenant[idx] = Object.assign({}, rowsTenant[idx], fresh);
                saveTenantCache(rowsTenant);
                setTenantProgress(true, 'Fertig. Änderungen wurden gespeichert.', 1);
                setTimeout(() => setTenantProgress(false, '', null), 1600);
                rerender();
            } catch (e) {
                setTenantProgress(true, 'Update fehlgeschlagen: ' + (e?.message || String(e)), null);
                toast('Update fehlgeschlagen: ' + (e?.message || String(e)));
            } finally {
                if (btn) btn.disabled = false;
                if (btnReload) btnReload.disabled = false;
                if (btnRenew) btnRenew.disabled = false;
                if (btnDelete) btnDelete.disabled = false;
                if (mode === 'tenant' && selectedId) {
                    const r = rowsTenant.find((x) => String(x.id) === String(selectedId));
                    if (r) showTenantDetail(r);
                }
            }
        });

        getEl('ssTenantRenewBtn')?.addEventListener('click', async () => {
            if (mode !== 'tenant' || !selectedId) return;
            if (!(await dlgConfirm('Ablaufdatum dieser Gruppe/dieses Teams verlängern (Renew)?', { title: 'Renew', okText: 'Verlängern' }))) return;
            const btn = getEl('ssTenantRenewBtn');
            const btnReload = getEl('ssTenantReloadBtn');
            const btnUpdate = getEl('ssTenantUpdateBtn');
            const btnDelete = getEl('ssTenantDeleteBtn');
            if (btn) btn.disabled = true;
            if (btnReload) btnReload.disabled = true;
            if (btnUpdate) btnUpdate.disabled = true;
            if (btnDelete) btnDelete.disabled = true;
            try {
                setTenantProgress(true, 'Verlängere Ablaufdatum …', 0.45);
                await renewTenantGroup(selectedId);
                setTenantProgress(true, 'Renew OK – lade Details neu …', 0.75);
                const fresh = await fetchTenantGroupDetail(selectedId);
                const idx = rowsTenant.findIndex((r) => String(r.id) === String(selectedId));
                if (idx !== -1) rowsTenant[idx] = Object.assign({}, rowsTenant[idx], fresh);
                saveTenantCache(rowsTenant);
                setTenantProgress(true, 'Fertig.', 1);
                setTimeout(() => setTenantProgress(false, '', null), 1400);
                rerender();
            } catch (e) {
                setTenantProgress(true, 'Renew fehlgeschlagen: ' + (e?.message || String(e)), null);
            } finally {
                if (btn) btn.disabled = false;
                if (btnReload) btnReload.disabled = false;
                if (btnUpdate) btnUpdate.disabled = false;
                if (btnDelete) btnDelete.disabled = false;
            }
        });

        getEl('ssTenantDeleteBtn')?.addEventListener('click', async () => {
            if (mode !== 'tenant' || !selectedId) return;
            if (
                !(await dlgConfirm(
                    'Diese Gruppe/dieses Team wirklich LÖSCHEN? Dieser Vorgang kann nicht rückgängig gemacht werden.',
                    { title: 'Löschen', okText: 'Endgültig löschen', danger: true }
                ))
            )
                return;
            const btn = getEl('ssTenantDeleteBtn');
            const btnReload = getEl('ssTenantReloadBtn');
            const btnUpdate = getEl('ssTenantUpdateBtn');
            const btnRenew = getEl('ssTenantRenewBtn');
            if (btn) btn.disabled = true;
            if (btnReload) btnReload.disabled = true;
            if (btnUpdate) btnUpdate.disabled = true;
            if (btnRenew) btnRenew.disabled = true;
            try {
                setTenantProgress(true, 'Lösche Gruppe/Team …', 0.45);
                await deleteTenantGroup(selectedId);
                rowsTenant = rowsTenant.filter((r) => String(r.id) !== String(selectedId));
                saveTenantCache(rowsTenant);
                selectedId = '';
                setTenantProgress(true, 'Gelöscht.', 1);
                setTimeout(() => setTenantProgress(false, '', null), 1200);
                rerender();
            } catch (e) {
                setTenantProgress(true, 'Löschen fehlgeschlagen: ' + (e?.message || String(e)), null);
            } finally {
                if (btn) btn.disabled = false;
                if (btnReload) btnReload.disabled = false;
                if (btnUpdate) btnUpdate.disabled = false;
                if (btnRenew) btnRenew.disabled = false;
            }
        });

        function fillUserSearchSelect(users) {
            const sel = getEl('ssOwnerSearchResults');
            if (!sel) return;
            sel.replaceChildren();
            if (!users || !users.length) {
                const opt = document.createElement('option');
                opt.value = '';
                opt.textContent = '(keine Treffer)';
                sel.appendChild(opt);
                return;
            }
            for (let i = 0; i < users.length; i++) {
                const u = users[i];
                const opt = document.createElement('option');
                opt.value = u.id || '';
                opt.textContent = personLabel(u) || (u.id ? String(u.id) : '');
                sel.appendChild(opt);
            }
        }

        function renderOwnersList(owners) {
            const wrap = getEl('ssOwnersList');
            if (!wrap) return;
            wrap.replaceChildren();
            const list = Array.isArray(owners) ? owners : [];
            if (!list.length) {
                const p = document.createElement('p');
                p.style.margin = '0';
                p.style.color = '#6c757d';
                p.textContent = 'Keine Besitzer gefunden (Achtung: das ist meist ein Problem).';
                wrap.appendChild(p);
                return;
            }
            list.forEach((o) => {
                const row = document.createElement('div');
                row.style.display = 'flex';
                row.style.justifyContent = 'space-between';
                row.style.alignItems = 'flex-start';
                row.style.gap = '10px';
                row.style.padding = '8px 0';
                row.style.borderBottom = '1px solid #e9ecef';
                const txt = document.createElement('div');
                txt.style.whiteSpace = 'pre-wrap';
                txt.style.lineHeight = '1.35';
                txt.style.fontSize = '0.92em';
                txt.textContent = personLabel(o) || '–';
                const btn = document.createElement('button');
                btn.type = 'button';
                btn.className = 'btn';
                btn.style.padding = '6px 10px';
                btn.style.fontSize = '0.85em';
                btn.textContent = 'Entfernen';
                btn.dataset.ssRemoveOwner = o.id || '';
                row.appendChild(txt);
                row.appendChild(btn);
                wrap.appendChild(row);
            });
        }

        let ownersCache = [];
        async function loadOwnersNow(showProgress) {
            if (mode !== 'tenant' || !selectedId) return;
            const btnReload = getEl('ssOwnersReloadBtn');
            const btnAdd = getEl('ssOwnerAddBtn');
            const btnSearch = getEl('ssOwnerSearchBtn');
            if (btnReload) btnReload.disabled = true;
            if (btnAdd) btnAdd.disabled = true;
            if (btnSearch) btnSearch.disabled = true;
            try {
                if (showProgress) setTenantProgress(true, 'Lade Owner …', 0.35);
                ownersCache = await fetchGroupOwners(selectedId);
                renderOwnersList(ownersCache);
                try {
                    const ix = rowsTenant.findIndex((r) => String(r.id) === String(selectedId));
                    if (ix !== -1) {
                        rowsTenant[ix] = Object.assign({}, rowsTenant[ix], { ownerCount: ownersCache.length });
                        saveTenantCache(rowsTenant);
                        rerender();
                    }
                } catch (_) {}
                if (showProgress) {
                    setTenantProgress(true, 'Owner geladen: ' + ownersCache.length, 1);
                    setTimeout(() => setTenantProgress(false, '', null), 900);
                }
            } catch (e) {
                setTenantProgress(true, 'Owner laden: ' + (e?.message || String(e)), null);
            } finally {
                if (btnReload) btnReload.disabled = false;
                if (btnAdd) btnAdd.disabled = false;
                if (btnSearch) btnSearch.disabled = false;
            }
        }

        // Owner: Suche
        getEl('ssOwnerSearchBtn')?.addEventListener('click', async () => {
            if (mode !== 'tenant') return;
            const q = getEl('ssOwnerSearch')?.value || '';
            const btn = getEl('ssOwnerSearchBtn');
            if (btn) btn.disabled = true;
            try {
                const token = await getGraphToken(GRAPH_SCOPES_TENANT_OWNER_MANAGE);
                const users = await graphSearchUsersForOwner(token, q);
                fillUserSearchSelect(users);
                toast('Suche: ' + users.length + ' Treffer.');
            } catch (e) {
                toast('Suche: ' + (e?.message || String(e)));
            } finally {
                if (btn) btn.disabled = false;
            }
        });

        // Owner: hinzufügen
        getEl('ssOwnerAddBtn')?.addEventListener('click', async () => {
            if (mode !== 'tenant' || !selectedId) return;
            const sel = getEl('ssOwnerSearchResults');
            const userId = sel && sel.value ? String(sel.value).trim() : '';
            if (!userId) {
                toast('Bitte zuerst einen Benutzer aus den Treffern auswählen.');
                return;
            }
            const btn = getEl('ssOwnerAddBtn');
            if (btn) btn.disabled = true;
            try {
                setTenantProgress(true, 'Owner wird hinzugefügt …', 0.45);
                await addOwnerWithMemberFallback(selectedId, userId);
                setTenantProgress(true, 'Owner hinzugefügt. Lade Liste neu …', 0.8);
                await loadOwnersNow(false);
                setTenantProgress(true, 'Fertig.', 1);
                setTimeout(() => setTenantProgress(false, '', null), 900);
            } catch (e) {
                setTenantProgress(true, 'Owner hinzufügen: ' + (e?.message || String(e)), null);
            } finally {
                if (btn) btn.disabled = false;
            }
        });

        // Owner: neu laden
        getEl('ssOwnersReloadBtn')?.addEventListener('click', async () => {
            await loadOwnersNow(true);
        });

        // Owner: entfernen (delegiert über Container)
        getEl('ssOwnersList')?.addEventListener('click', async (ev) => {
            if (mode !== 'tenant' || !selectedId) return;
            const t = ev.target;
            const btn = t && t.closest ? t.closest('button[data-ss-remove-owner]') : null;
            if (!btn) return;
            const ownerId = btn.getAttribute('data-ss-remove-owner') || '';
            if (!ownerId) return;
            if (ownersCache.length <= 1) {
                toast('Der letzte Besitzer kann nicht entfernt werden.');
                return;
            }
            if (!(await dlgConfirm('Diesen Owner wirklich entfernen?', { title: 'Owner', okText: 'Entfernen', danger: true }))) return;
            try {
                setTenantProgress(true, 'Owner wird entfernt …', 0.45);
                await removeGroupOwner(selectedId, ownerId);
                setTenantProgress(true, 'Owner entfernt. Lade Liste neu …', 0.8);
                await loadOwnersNow(false);
                setTenantProgress(true, 'Fertig.', 1);
                setTimeout(() => setTenantProgress(false, '', null), 900);
            } catch (e) {
                setTenantProgress(true, 'Owner entfernen: ' + (e?.message || String(e)), null);
            }
        });

        function fillMemberSearchSelect(users) {
            const sel = getEl('ssMemberSearchResults');
            if (!sel) return;
            sel.replaceChildren();
            if (!users || !users.length) {
                const opt = document.createElement('option');
                opt.value = '';
                opt.textContent = '(keine Treffer)';
                sel.appendChild(opt);
                return;
            }
            for (let i = 0; i < users.length; i++) {
                const u = users[i];
                const opt = document.createElement('option');
                opt.value = u.id || '';
                opt.textContent = personLabel(u) || (u.id ? String(u.id) : '');
                sel.appendChild(opt);
            }
        }

        function renderMembersList(result, totalCount) {
            const wrap = getEl('ssMembersList');
            if (!wrap) return;
            wrap.replaceChildren();
            const list = result && Array.isArray(result.items) ? result.items : [];
            const truncated = !!(result && result.truncated);

            const head = document.createElement('div');
            head.style.display = 'flex';
            head.style.justifyContent = 'space-between';
            head.style.alignItems = 'baseline';
            head.style.gap = '10px';
            head.style.marginBottom = '8px';
            const left = document.createElement('div');
            left.style.fontWeight = '900';
            left.style.color = '#32325d';
            const totalTxt = totalCount >= 0 ? String(totalCount) : String(list.length);
            left.textContent = 'Mitglieder: ' + totalTxt + (truncated ? ' (Anzeige gekürzt)' : '');
            const right = document.createElement('div');
            right.className = 'pill';
            right.textContent = truncated ? 'gekürzt' : 'vollständig';
            head.appendChild(left);
            head.appendChild(right);
            wrap.appendChild(head);

            if (!list.length) {
                const p = document.createElement('p');
                p.style.margin = '0';
                p.style.color = '#6c757d';
                p.textContent = 'Keine Mitglieder.';
                wrap.appendChild(p);
                return;
            }

            list.forEach((m) => {
                const row = document.createElement('div');
                row.style.display = 'flex';
                row.style.justifyContent = 'space-between';
                row.style.alignItems = 'flex-start';
                row.style.gap = '10px';
                row.style.padding = '8px 0';
                row.style.borderBottom = '1px solid #e9ecef';
                const txt = document.createElement('div');
                txt.style.whiteSpace = 'pre-wrap';
                txt.style.lineHeight = '1.35';
                txt.style.fontSize = '0.92em';
                txt.textContent = personLabel(m) || '–';
                const btn = document.createElement('button');
                btn.type = 'button';
                btn.className = 'btn';
                btn.style.padding = '6px 10px';
                btn.style.fontSize = '0.85em';
                btn.textContent = 'Entfernen';
                btn.dataset.ssRemoveMember = m.id || '';
                row.appendChild(txt);
                row.appendChild(btn);
                wrap.appendChild(row);
            });
        }

        let membersCache = [];
        let membersCountCache = -1;
        async function loadMembersNow(showProgress) {
            if (mode !== 'tenant' || !selectedId) return;
            const btnReload = getEl('ssMembersReloadBtn');
            const btnAdd = getEl('ssMemberAddBtn');
            const btnSearch = getEl('ssMemberSearchBtn');
            if (btnReload) btnReload.disabled = true;
            if (btnAdd) btnAdd.disabled = true;
            if (btnSearch) btnSearch.disabled = true;
            try {
                if (showProgress) setTenantProgress(true, 'Lade Mitglieder …', 0.35);
                // Count (best effort)
                try {
                    membersCountCache = await fetchGroupMemberCount(selectedId);
                } catch {
                    membersCountCache = -1;
                }
                const result = await fetchGroupMembers(selectedId);
                membersCache = result.items || [];
                renderMembersList(result, membersCountCache);
                try {
                    const ix = rowsTenant.findIndex((r) => String(r.id) === String(selectedId));
                    if (ix !== -1 && membersCountCache >= 0) {
                        rowsTenant[ix] = Object.assign({}, rowsTenant[ix], { memberCount: membersCountCache });
                        saveTenantCache(rowsTenant);
                        rerender();
                    }
                } catch (_) {}
                if (showProgress) {
                    setTenantProgress(true, 'Mitglieder geladen.', 1);
                    setTimeout(() => setTenantProgress(false, '', null), 900);
                }
            } catch (e) {
                setTenantProgress(true, 'Mitglieder laden: ' + (e?.message || String(e)), null);
            } finally {
                if (btnReload) btnReload.disabled = false;
                if (btnAdd) btnAdd.disabled = false;
                if (btnSearch) btnSearch.disabled = false;
            }
        }

        // Member: Suche
        getEl('ssMemberSearchBtn')?.addEventListener('click', async () => {
            if (mode !== 'tenant') return;
            const q = getEl('ssMemberSearch')?.value || '';
            const btn = getEl('ssMemberSearchBtn');
            if (btn) btn.disabled = true;
            try {
                const token = await getGraphToken(GRAPH_SCOPES_TENANT_OWNER_MANAGE);
                const users = await graphSearchUsersForOwner(token, q);
                fillMemberSearchSelect(users);
                toast('Suche: ' + users.length + ' Treffer.');
            } catch (e) {
                toast('Suche: ' + (e?.message || String(e)));
            } finally {
                if (btn) btn.disabled = false;
            }
        });

        // Member: hinzufügen
        getEl('ssMemberAddBtn')?.addEventListener('click', async () => {
            if (mode !== 'tenant' || !selectedId) return;
            const sel = getEl('ssMemberSearchResults');
            const userId = sel && sel.value ? String(sel.value).trim() : '';
            if (!userId) {
                toast('Bitte zuerst einen Benutzer aus den Treffern auswählen.');
                return;
            }
            const btn = getEl('ssMemberAddBtn');
            if (btn) btn.disabled = true;
            try {
                setTenantProgress(true, 'Mitglied wird hinzugefügt …', 0.45);
                await addGroupMember(selectedId, userId);
                setTenantProgress(true, 'Mitglied hinzugefügt. Lade Liste neu …', 0.8);
                await loadMembersNow(false);
                setTenantProgress(true, 'Fertig.', 1);
                setTimeout(() => setTenantProgress(false, '', null), 900);
            } catch (e) {
                const msg = e?.message || String(e);
                if (/added object references already exist/i.test(msg) || /already exist/i.test(msg)) {
                    setTenantProgress(true, 'Hinweis: Benutzer war bereits Mitglied.', 1);
                    setTimeout(() => setTenantProgress(false, '', null), 900);
                } else {
                    setTenantProgress(true, 'Mitglied hinzufügen: ' + msg, null);
                }
            } finally {
                if (btn) btn.disabled = false;
            }
        });

        // Member: neu laden
        getEl('ssMembersReloadBtn')?.addEventListener('click', async () => {
            await loadMembersNow(true);
        });

        // Member: entfernen
        getEl('ssMembersList')?.addEventListener('click', async (ev) => {
            if (mode !== 'tenant' || !selectedId) return;
            const t = ev.target;
            const btn = t && t.closest ? t.closest('button[data-ss-remove-member]') : null;
            if (!btn) return;
            const memberId = btn.getAttribute('data-ss-remove-member') || '';
            if (!memberId) return;
            if (!(await dlgConfirm('Dieses Mitglied wirklich entfernen?', { title: 'Mitglied', okText: 'Entfernen', danger: true }))) return;
            try {
                setTenantProgress(true, 'Mitglied wird entfernt …', 0.45);
                await removeGroupMember(selectedId, memberId);
                setTenantProgress(true, 'Mitglied entfernt. Lade Liste neu …', 0.8);
                await loadMembersNow(false);
                setTenantProgress(true, 'Fertig.', 1);
                setTimeout(() => setTenantProgress(false, '', null), 900);
            } catch (e) {
                setTenantProgress(true, 'Mitglied entfernen: ' + (e?.message || String(e)), null);
            }
        });

        // Tenant-Detail Tabs (Allgemein / Owner / Mitglieder)
        const tAllgBtn = getEl('ssTenantTabAllgemeinBtn');
        const tOwnBtn = getEl('ssTenantTabOwnerBtn');
        const tMemBtn = getEl('ssTenantTabMitgliederBtn');
        const pAllg = getEl('ssTenantTabAllgemein');
        const pOwn = getEl('ssTenantTabOwner');
        const pMem = getEl('ssTenantTabMitglieder');
        let tenantActiveTab = 'allg';

        function setTenantTab(next) {
            tenantActiveTab = next === 'own' ? 'own' : next === 'mem' ? 'mem' : 'allg';
            if (tAllgBtn) tAllgBtn.setAttribute('aria-selected', tenantActiveTab === 'allg' ? 'true' : 'false');
            if (tOwnBtn) tOwnBtn.setAttribute('aria-selected', tenantActiveTab === 'own' ? 'true' : 'false');
            if (tMemBtn) tMemBtn.setAttribute('aria-selected', tenantActiveTab === 'mem' ? 'true' : 'false');
            if (pAllg) pAllg.classList.toggle('active', tenantActiveTab === 'allg');
            if (pOwn) pOwn.classList.toggle('active', tenantActiveTab === 'own');
            if (pMem) pMem.classList.toggle('active', tenantActiveTab === 'mem');

            // Lazy-load when switching tabs
            if (mode === 'tenant' && selectedId) {
                if (!isTenantBulkMode()) {
                    if (tenantActiveTab === 'own') loadOwnersNow(false);
                    if (tenantActiveTab === 'mem') loadMembersNow(false);
                }
            }
        }

        if (tAllgBtn) tAllgBtn.addEventListener('click', () => setTenantTab('allg'));
        if (tOwnBtn) tOwnBtn.addEventListener('click', () => setTenantTab('own'));
        if (tMemBtn) tMemBtn.addEventListener('click', () => setTenantTab('mem'));
        setTenantTab('allg');

        // initial render (restore last tab if available)
        function readStartMode() {
            const valid = (m) => m === 'tenant' || m === 'match' || m === 'struktur';
            // 1) URL parameter (?mode=match)
            try {
                const q = new URLSearchParams(String(window.location.search || ''));
                const m = String(q.get('mode') || '').trim();
                if (valid(m)) return m;
            } catch {
                // ignore
            }
            // 2) Explicit force mode on body
            try {
                const m = String(document?.body?.getAttribute('data-ss-force-mode') || '').trim();
                if (valid(m)) return m;
            } catch {
                // ignore
            }
            // 3) Default mode on body (used by pages without tabs)
            try {
                const m = String(document?.body?.getAttribute('data-ss-default-mode') || '').trim();
                if (valid(m)) return m;
            } catch {
                // ignore
            }
            // 4) Restore last
            try {
                const saved = String(sessionStorage.getItem(UI_MODE_KEY) || '').trim();
                if (valid(saved)) return saved;
            } catch {
                // ignore
            }
            // 5) Fallback
            if (!tabStruktur && !tabMatch) return 'tenant';
            return 'struktur';
        }
        mode = readStartMode();
        if (tabStruktur) tabStruktur.setAttribute('aria-selected', mode === 'struktur' ? 'true' : 'false');
        if (tabMatch) tabMatch.setAttribute('aria-selected', mode === 'match' ? 'true' : 'false');
        if (tabTenant) tabTenant.setAttribute('aria-selected', mode === 'tenant' ? 'true' : 'false');
        updateModeUi();
        rerender();
    }

    function renderMatchDetail(unit) {
        const inpUnit = getEl('ssMatchUnit');
        const selTenant = getEl('ssMatchTenantGroup');
        const note = getEl('ssMatchNote');
        if (!inpUnit || !selTenant || !note) return;

        inpUnit.value = (unit.bezeichnung || '(ohne Bezeichnung)') + ' – ' + (unit.typ || '');

        const list = Array.isArray(window.__ms365TenantRowsCache) ? window.__ms365TenantRowsCache : [];
        const users = Array.isArray(window.__ms365TenantUsersCache) ? window.__ms365TenantUsersCache : [];
        window.__ms365MatchTenantPickSource = { groups: list, users };

        const links = window.__ms365MatchLinks || {};
        const cur = links[String(unit.id)] || null;
        let selVal = '';
        if (cur && normStr(cur.tenantUserId)) selVal = 'u:' + String(cur.tenantUserId);
        else if (cur && normStr(cur.tenantGroupId)) {
            const gid = String(cur.tenantGroupId);
            const isUser = users.some((u) => String(u.id) === gid);
            selVal = (isUser ? 'u:' : 'g:') + gid;
        }

        const inpSearch = getEl('ssMatchTenantSearch');
        if (inpSearch) inpSearch.value = '';
        rebuildMatchTenantSelectOptions(selTenant, '', selVal);
        wireMatchTenantSearchOnce();

        note.value = cur && cur.note ? String(cur.note) : '';
    }

    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', bind);
    else bind();
})();

