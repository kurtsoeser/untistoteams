(function () {
    'use strict';

    const STORAGE_KEY_V2 = 'ms365-schooltool-data-v2';
    /** Schema 3: setup (Wizard, SLG-Matches, Fach-/ARGE-Gruppen-Links) */
    const VERSION = 3;
    const SLG_LEGACY_KEY = 'ms365-schueler-lehrer-gruppen-v2';

    function safeJsonParse(s) {
        try {
            return JSON.parse(String(s));
        } catch {
            return null;
        }
    }

    function deepClone(obj) {
        try {
            return JSON.parse(JSON.stringify(obj));
        } catch {
            return obj;
        }
    }

    function currentSchoolYearLabel() {
        const y = new Date().getFullYear();
        return String(y) + '/' + String(y + 1).slice(2);
    }

    function defaultSetup() {
        return {
            wizardStep: 1,
            completedSteps: [],
            finishedAt: null,
            lastVisitedAt: null,
            matched: { schuelerGroupId: null, lehrerGroupId: null, verwaltungGroupId: null },
            slgDraft: {
                activeKind: 'schueler',
                slgNewDisplayName: '',
                slgNewMailNick: '',
                slgNewDescription: '',
                slgNewCreateTeam: false,
                /** Besitzer Lehrer-Sammelgruppe: direktion | teachers | manual */
                slgOwnerSourceLehrer: 'direktion',
                slgOwnerManualEmailsLehrer: '',
                /** Besitzer Schüler-Sammelgruppe: direktion | admin | manual */
                slgOwnerSourceSchueler: 'direktion',
                slgOwnerManualEmailsSchueler: ''
            },
            verwaltungDraft: {
                vwNewDisplayName: 'Schulverwaltung',
                vwNewMailNick: 'verwaltung',
                vwNewDescription: '',
                vwNewCreateTeam: false,
                /** Besitzer Verwaltungs-Sammelgruppe: admin | direktion | manual */
                vwOwnerSource: 'admin',
                vwOwnerManualEmails: ''
            },
            /** Kleinbuchstaben/Ziffern; Vorschau/Anlage Fachgruppen (Einrichtungsassistent) */
            subjectGroupMailPrefix: 'fach',
            /** Kleinbuchstaben/Ziffern; Vorschau/Anlage ARGE-Gruppen */
            argeGroupMailPrefix: 'ag',
            /** E‑Mail (Kleinbuchstaben) → Entra-Benutzer (Einrichtungsassistent, optional) */
            directoryMatchByEmail: {},
            catalogLinks: []
        };
    }

    function normEmailKey(v) {
        return String(v ?? '')
            .trim()
            .toLowerCase();
    }

    /**
     * Präfix für group.mailNickname (Einrichtungsassistent): gleiche Zeichenregel wie
     * sanitizeUnifiedGroupMailNickname in graph-unified-groups.js (Microsoft Learn:
     * directoryObject validateProperties – ungültig u. a. @ ( ) \ [ ] " ; : < > , und Leerzeichen).
     * Zusätzlich nur ASCII, keine Steuerzeichen. Erlaubt u. a. . - _
     */
    function mailNicknamePrefixSanitize(raw, maxLen) {
        const lim = typeof maxLen === 'number' && maxLen > 0 ? maxLen : 24;
        const s = String(raw ?? '')
            .trim()
            .toLowerCase();
        let out = '';
        for (let i = 0; i < s.length; i++) {
            const c = s.charCodeAt(i);
            if (c < 32 || c === 127 || c > 127) continue;
            const ch = s.charAt(i);
            if (/[@()[\]\\";:<>,\s]/.test(ch)) continue;
            out += ch;
        }
        if (out.length > lim) out = out.slice(0, lim);
        return out;
    }

    function normalizeDirectoryMatchByEmail(raw) {
        const out = {};
        const src = raw && typeof raw === 'object' ? raw : {};
        Object.keys(src).forEach(function (k) {
            const em = normEmailKey(k);
            if (!em || em.indexOf('@') === -1) return;
            const v = src[k];
            if (!v || typeof v !== 'object') return;
            if (v.notFound === true) {
                out[em] = {
                    graphUserId: '',
                    displayName: '',
                    userPrincipalName: '',
                    notFound: true,
                    checkedAt: String(v.checkedAt || '')
                };
                return;
            }
            const id = String(v.graphUserId || v.id || '').trim();
            if (!id) return;
            out[em] = {
                graphUserId: id,
                displayName: String(v.displayName || '').trim(),
                userPrincipalName: String(v.userPrincipalName || '').trim(),
                notFound: false,
                checkedAt: String(v.checkedAt || '')
            };
        });
        return out;
    }

    function normCode(v) {
        return String(v ?? '')
            .trim()
            .toUpperCase();
    }

    function normalizeCatalogLink(row) {
        const r = row && typeof row === 'object' ? row : {};
        const kind = r.kind === 'arge' ? 'arge' : 'subject';
        const code = normCode(r.code);
        if (!code) return null;
        const mode = r.mode === 'created' || r.mode === 'matched' ? r.mode : '';
        return {
            kind: kind,
            code: code,
            graphGroupId: r.graphGroupId ? String(r.graphGroupId).trim() : '',
            displayName: String(r.displayName || '').trim(),
            mailNickname: String(r.mailNickname || '').trim(),
            mode: mode,
            syncStatus: String(r.syncStatus || '').trim()
        };
    }

    function normalizeSetup(s) {
        const d = defaultSetup();
        const x = s && typeof s === 'object' ? s : {};
        let ws = parseInt(x.wizardStep, 10);
        const layout9 = x._einrichtungWizardLayout === 9;
        const layout8 = x._einrichtungWizardLayout === 8;
        const layout7 = x._einrichtungWizardLayout === 7;
        const layout6 = x._einrichtungWizardLayout === 6;
        // Nur echte Layout-Upgrades (6/7 → 8); bei bereits 8 oder 9 keine erneute Verschiebung
        if (!layout8 && !layout9 && !isNaN(ws)) {
            // Layout 7 → 8: neuer Schritt „Verwaltung“ vor Lehrkräften (alte 3–7 werden 4–8)
            if (layout7 && ws >= 3 && ws <= 7) ws += 1;
            // Sehr alt: 5 Schritte (4=Katalog, 5=Klassen) → +1 für eingefügte Personen-Schritte
            if (!layout7 && !layout6 && ws >= 4 && ws <= 5) ws += 1;
            // Vorher 6 Schritte: Klassen war Schritt 6 → jetzt Schritt 7
            if (layout6 && ws === 6) ws = 7;
        }
        d.wizardStep = !isNaN(ws) && ws >= 1 && ws <= 9 ? ws : 1;
        d._einrichtungWizardLayout = 9;
        d.completedSteps = Array.isArray(x.completedSteps) ? x.completedSteps.map((t) => String(t)) : [];
        d.finishedAt = x.finishedAt != null && x.finishedAt !== '' ? String(x.finishedAt) : null;
        d.lastVisitedAt = x.lastVisitedAt != null && x.lastVisitedAt !== '' ? String(x.lastVisitedAt) : null;
        const m = x.matched && typeof x.matched === 'object' ? x.matched : {};
        d.matched = {
            schuelerGroupId: m.schuelerGroupId ? String(m.schuelerGroupId).trim() : null,
            lehrerGroupId: m.lehrerGroupId ? String(m.lehrerGroupId).trim() : null,
            verwaltungGroupId: m.verwaltungGroupId ? String(m.verwaltungGroupId).trim() : null
        };
        const dr = x.slgDraft && typeof x.slgDraft === 'object' ? x.slgDraft : {};
        const srcLehrer = String(dr.slgOwnerSourceLehrer || '').trim();
        const srcSchueler = String(dr.slgOwnerSourceSchueler || '').trim();
        d.slgDraft = {
            activeKind: dr.activeKind === 'lehrer' ? 'lehrer' : 'schueler',
            slgNewDisplayName: String(dr.slgNewDisplayName != null ? dr.slgNewDisplayName : ''),
            slgNewMailNick: String(dr.slgNewMailNick != null ? dr.slgNewMailNick : ''),
            slgNewDescription: String(dr.slgNewDescription != null ? dr.slgNewDescription : ''),
            slgNewCreateTeam: !!dr.slgNewCreateTeam,
            slgOwnerSourceLehrer:
                srcLehrer === 'teachers' || srcLehrer === 'manual' ? srcLehrer : 'direktion',
            slgOwnerManualEmailsLehrer: String(dr.slgOwnerManualEmailsLehrer != null ? dr.slgOwnerManualEmailsLehrer : ''),
            slgOwnerSourceSchueler:
                srcSchueler === 'admin' || srcSchueler === 'manual' ? srcSchueler : 'direktion',
            slgOwnerManualEmailsSchueler: String(
                dr.slgOwnerManualEmailsSchueler != null ? dr.slgOwnerManualEmailsSchueler : ''
            )
        };
        const vd = x.verwaltungDraft && typeof x.verwaltungDraft === 'object' ? x.verwaltungDraft : {};
        const vwSrc = String(vd.vwOwnerSource || '').trim();
        d.verwaltungDraft = {
            vwNewDisplayName: String(vd.vwNewDisplayName != null ? vd.vwNewDisplayName : 'Schulverwaltung'),
            vwNewMailNick: mailNicknamePrefixSanitize(vd.vwNewMailNick || 'verwaltung', 60) || 'verwaltung',
            vwNewDescription: String(vd.vwNewDescription != null ? vd.vwNewDescription : ''),
            vwNewCreateTeam: !!vd.vwNewCreateTeam,
            vwOwnerSource: vwSrc === 'direktion' || vwSrc === 'manual' ? vwSrc : 'admin',
            vwOwnerManualEmails: String(vd.vwOwnerManualEmails != null ? vd.vwOwnerManualEmails : '')
        };
        d.subjectGroupMailPrefix = mailNicknamePrefixSanitize(x.subjectGroupMailPrefix, 24) || 'fach';
        d.argeGroupMailPrefix = mailNicknamePrefixSanitize(x.argeGroupMailPrefix, 24) || 'ag';
        const linksIn = Array.isArray(x.catalogLinks) ? x.catalogLinks : [];
        const seen = new Set();
        d.catalogLinks = [];
        linksIn.forEach(function (row) {
            const n = normalizeCatalogLink(row);
            if (!n) return;
            const k = n.kind + ':' + n.code.toLowerCase();
            if (seen.has(k)) return;
            seen.add(k);
            d.catalogLinks.push(n);
        });
        d.directoryMatchByEmail = normalizeDirectoryMatchByEmail(x.directoryMatchByEmail);
        return d;
    }

    function deriveStableNickFromClassRow(cl) {
        if (typeof window.ms365DeriveClassStableMailNickname === 'function') {
            return String(window.ms365DeriveClassStableMailNickname(cl.year || '', cl.code || '') || '')
                .trim()
                .replace(/[^a-zA-Z0-9]/g, '')
                .toLowerCase()
                .slice(0, 60);
        }
        const y = String(cl.year || '').trim();
        const yy = /^\d{4}$/.test(y) ? y : '';
        const code = normCode(cl.code || '');
        const tail = String(code)
            .replace(/[^0-9A-Za-z]/g, '')
            .toLowerCase()
            .slice(0, 24);
        if (!yy || !tail) return '';
        return ('jg' + yy + tail).toLowerCase().slice(0, 60);
    }

    function normalizeClassTeam(row) {
        const r = row && typeof row === 'object' ? row : {};
        let nick = String(r.stableMailNickname || '')
            .trim()
            .replace(/[^a-zA-Z0-9]/g, '')
            .toLowerCase()
            .slice(0, 60);
        if (!nick) return null;
        const mode = r.mode === 'created' || r.mode === 'matched' ? r.mode : '';
        const y = String(r.abschlussJahr || r.year || '').trim();
        const abschlussJahr = /^\d{4}$/.test(y) ? y : '';
        return {
            stableMailNickname: nick,
            graphGroupId: String(r.graphGroupId || '').trim(),
            classCode: normCode(r.classCode || r.code || ''),
            displayName: String(r.displayName || r.name || '').trim(),
            abschlussJahr: abschlussJahr,
            mode: mode,
            educationClassId: String(r.educationClassId || '').trim()
        };
    }

    function normalizeCoreClassTeams(arr) {
        const seen = new Set();
        const out = [];
        (Array.isArray(arr) ? arr : []).forEach(function (row) {
            const n = normalizeClassTeam(row);
            if (!n) return;
            if (seen.has(n.stableMailNickname)) return;
            seen.add(n.stableMailNickname);
            out.push(n);
        });
        return out;
    }

    function classTeamMatchesKlasse(ct, klasseRaw) {
        const k = String(klasseRaw ?? '').trim();
        if (!k) return false;
        const cc = normCode(ct.classCode || '');
        const dn = String(ct.displayName || '').trim();
        if (dn && k === dn) return true;
        const nk = normCode(k);
        if (cc && nk === cc) return true;
        if (cc && cc.length >= 2 && k.toUpperCase().indexOf(cc) !== -1) return true;
        return false;
    }

    function reconcileClassTeamsFromYearClasses(c, _yearKey, classesArr) {
        const classes = Array.isArray(classesArr) ? classesArr : [];
        let teams = normalizeCoreClassTeams(c.core.classTeams || []);
        const byNick = {};
        teams.forEach(function (t) {
            byNick[t.stableMailNickname] = t;
        });
        classes.forEach(function (cl) {
            let nick = String(cl.stableMailNickname || '')
                .trim()
                .replace(/[^a-zA-Z0-9]/g, '')
                .toLowerCase()
                .slice(0, 60);
            if (!nick && cl.year && cl.code) nick = deriveStableNickFromClassRow(cl);
            if (!nick) return;
            let ex = byNick[nick];
            if (!ex) {
                ex = normalizeClassTeam({
                    stableMailNickname: nick,
                    classCode: cl.code,
                    displayName: cl.name,
                    abschlussJahr: cl.year,
                    graphGroupId: '',
                    mode: ''
                });
                if (!ex) return;
                teams.push(ex);
                byNick[nick] = ex;
            } else {
                if (cl.name) ex.displayName = String(cl.name).trim();
                if (cl.code) ex.classCode = normCode(cl.code);
                const yr = String(cl.year || '').trim();
                if (/^\d{4}$/.test(yr)) ex.abschlussJahr = yr;
            }
        });
        c.core.classTeams = normalizeCoreClassTeams(teams);
    }

    function emptyContainer() {
        return {
            version: VERSION,
            core: {
                domain: '',
                subjects: [],
                arges: [],
                teachers: [],
                admin: [],
                classTeams: []
            },
            years: {
                current: currentSchoolYearLabel(),
                byLabel: {}
            },
            structure: {
                rows: [],
                memberships: {},
                settings: {}
            },
            tenant: {
                cache: {
                    rows: [],
                    users: [],
                    loadedAt: ''
                }
            },
            match: {
                links: {}
            },
            setup: defaultSetup()
        };
    }

    function normalizeContainer(obj) {
        const base = emptyContainer();
        const o = obj && typeof obj === 'object' ? obj : {};
        const out = Object.assign({}, base, o);
        out.version = VERSION;

        out.core = Object.assign({}, base.core, (o.core && typeof o.core === 'object' ? o.core : {}));
        out.core.classTeams = normalizeCoreClassTeams(out.core.classTeams);
        out.years = Object.assign({}, base.years, (o.years && typeof o.years === 'object' ? o.years : {}));
        out.years.byLabel = Object.assign({}, base.years.byLabel, (out.years.byLabel && typeof out.years.byLabel === 'object' ? out.years.byLabel : {}));

        out.structure = Object.assign({}, base.structure, (o.structure && typeof o.structure === 'object' ? o.structure : {}));
        out.tenant = Object.assign({}, base.tenant, (o.tenant && typeof o.tenant === 'object' ? o.tenant : {}));
        out.tenant.cache = Object.assign({}, base.tenant.cache, (out.tenant.cache && typeof out.tenant.cache === 'object' ? out.tenant.cache : {}));
        out.match = Object.assign({}, base.match, (o.match && typeof o.match === 'object' ? o.match : {}));
        out.match.links = Object.assign({}, base.match.links, (out.match.links && typeof out.match.links === 'object' ? out.match.links : {}));

        out.setup = normalizeSetup(o.setup);

        if (!out.years.current) out.years.current = currentSchoolYearLabel();
        if (!out.years.byLabel[out.years.current]) out.years.byLabel[out.years.current] = { students: [], classes: [] };

        return out;
    }

    function loadV2Raw() {
        try {
            const raw = localStorage.getItem(STORAGE_KEY_V2);
            if (!raw) return null;
            return safeJsonParse(raw);
        } catch {
            return null;
        }
    }

    function saveV2(container) {
        const normalized = normalizeContainer(container);
        try {
            localStorage.setItem(STORAGE_KEY_V2, JSON.stringify(normalized));
        } catch {
            // ignore
        }
        return normalized;
    }

    function migrateFromV1IfNeeded() {
        const existing = loadV2Raw();
        if (existing && typeof existing === 'object') {
            return saveV2(normalizeContainer(existing));
        }

        // Migrate from legacy keys (best-effort, non-destructive)
        const out = emptyContainer();

        // tenant-settings-core v1
        try {
            const rawCore = localStorage.getItem('ms365-tenant-settings-v1');
            const coreObj = rawCore ? safeJsonParse(rawCore) : null;
            if (coreObj && typeof coreObj === 'object') {
                out.core.domain = String(coreObj.domain || '').trim();
                out.core.subjects = Array.isArray(coreObj.subjects) ? deepClone(coreObj.subjects) : [];
                out.core.arges = Array.isArray(coreObj.arges) ? deepClone(coreObj.arges) : [];
                out.core.teachers = Array.isArray(coreObj.teachers) ? deepClone(coreObj.teachers) : [];
                out.core.admin = Array.isArray(coreObj.admin) ? deepClone(coreObj.admin) : [];

                const cur = out.years.current;
                out.years.byLabel[cur] = {
                    students: Array.isArray(coreObj.students) ? deepClone(coreObj.students) : [],
                    classes: Array.isArray(coreObj.classes) ? deepClone(coreObj.classes) : []
                };
            }
        } catch {
            // ignore
        }

        // schulstruktur-sync v1
        try {
            const rawStruct = localStorage.getItem('ms365-schulstruktur-sync-v1');
            const st = rawStruct ? safeJsonParse(rawStruct) : null;
            if (st && typeof st === 'object') {
                out.structure.rows = Array.isArray(st.rows) ? deepClone(st.rows) : [];
                out.structure.memberships = st.memberships && typeof st.memberships === 'object' ? deepClone(st.memberships) : {};
                out.structure.settings = st.settings && typeof st.settings === 'object' ? deepClone(st.settings) : {};
            }
        } catch {
            // ignore
        }

        try {
            const rawMatch = localStorage.getItem('ms365-schulstruktur-match-v1');
            const m = rawMatch ? safeJsonParse(rawMatch) : null;
            if (m && typeof m === 'object' && m.links && typeof m.links === 'object') {
                out.match.links = deepClone(m.links);
            }
        } catch {
            // ignore
        }

        try {
            const rawCache = localStorage.getItem('ms365-schulstruktur-tenant-cache-v1');
            const c = rawCache ? safeJsonParse(rawCache) : null;
            if (c && typeof c === 'object') {
                out.tenant.cache.rows = Array.isArray(c.rows) ? deepClone(c.rows) : [];
                out.tenant.cache.users = Array.isArray(c.users) ? deepClone(c.users) : [];
                out.tenant.cache.loadedAt = String(c.loadedAt || '');
            }
        } catch {
            // ignore
        }

        return saveV2(out);
    }

    function maybeMergeSlgLocalIntoSetup(c) {
        try {
            const m = c.setup && c.setup.matched;
            if (m && (m.schuelerGroupId || m.lehrerGroupId)) return c;
            const raw = localStorage.getItem(SLG_LEGACY_KEY);
            if (!raw) return c;
            const o = safeJsonParse(raw);
            if (!o || typeof o !== 'object' || !o.matched || typeof o.matched !== 'object') return c;
            const sm = o.matched.schuelerGroupId ? String(o.matched.schuelerGroupId).trim() : '';
            const lm = o.matched.lehrerGroupId ? String(o.matched.lehrerGroupId).trim() : '';
            if (!sm && !lm) return c;
            c.setup = normalizeSetup(c.setup);
            if (sm) c.setup.matched.schuelerGroupId = sm;
            if (lm) c.setup.matched.lehrerGroupId = lm;
            if (o.activeKind === 'lehrer' || o.activeKind === 'schueler') {
                c.setup.slgDraft.activeKind = o.activeKind;
            }
            if (o.slgNewDisplayName !== undefined) c.setup.slgDraft.slgNewDisplayName = String(o.slgNewDisplayName);
            if (o.slgNewMailNick !== undefined) c.setup.slgDraft.slgNewMailNick = String(o.slgNewMailNick);
            if (o.slgNewDescription !== undefined) c.setup.slgDraft.slgNewDescription = String(o.slgNewDescription);
            if (o.slgNewCreateTeam !== undefined) c.setup.slgDraft.slgNewCreateTeam = !!o.slgNewCreateTeam;
            return saveV2(c);
        } catch {
            return c;
        }
    }

    function getContainer() {
        const c = migrateFromV1IfNeeded();
        return maybeMergeSlgLocalIntoSetup(c);
    }

    function setContainer(next) {
        return saveV2(next);
    }

    function setCoreFromTenantSettings(v1Settings) {
        const c = getContainer();
        const keepClassTeams = normalizeCoreClassTeams(c.core.classTeams || []);
        const s = v1Settings && typeof v1Settings === 'object' ? v1Settings : {};
        c.core.domain = String(s.domain || '').trim();
        c.core.subjects = Array.isArray(s.subjects) ? deepClone(s.subjects) : [];
        c.core.arges = Array.isArray(s.arges) ? deepClone(s.arges) : [];
        c.core.teachers = Array.isArray(s.teachers) ? deepClone(s.teachers) : [];
        c.core.admin = Array.isArray(s.admin) ? deepClone(s.admin) : [];
        c.core.classTeams = keepClassTeams;
        const cur = String(c.years.current || currentSchoolYearLabel());
        if (!c.years.byLabel[cur]) c.years.byLabel[cur] = { students: [], classes: [] };
        c.years.byLabel[cur].students = Array.isArray(s.students) ? deepClone(s.students) : [];
        c.years.byLabel[cur].classes = Array.isArray(s.classes) ? deepClone(s.classes) : [];
        reconcileClassTeamsFromYearClasses(c, cur, c.years.byLabel[cur].classes);
        return saveV2(c);
    }

    function listYears() {
        const c = getContainer();
        const by = c && c.years && c.years.byLabel && typeof c.years.byLabel === 'object' ? c.years.byLabel : {};
        return Object.keys(by).map((k) => String(k)).sort();
    }

    function setCurrentYear(label, opts) {
        const y = String(label || '').trim();
        if (!y) throw new Error('Schuljahr fehlt.');
        const o = opts && typeof opts === 'object' ? opts : {};
        const c = getContainer();
        const by = c.years.byLabel || {};
        if (!by[y]) {
            let seed = { students: [], classes: [] };
            const copyFrom = String(o.copyFrom || '').trim();
            if (copyFrom && by[copyFrom]) {
                // Kopie nur von Schüler/Klassen; alles andere ist global.
                seed = {
                    students: deepClone(by[copyFrom].students || []),
                    classes: deepClone(by[copyFrom].classes || [])
                };
            }
            by[y] = seed;
        }
        c.years.byLabel = by;
        c.years.current = y;
        return saveV2(c);
    }

    function getSetup() {
        const c = getContainer();
        return normalizeSetup(c.setup);
    }

    function patchSetup(partial) {
        const c = getContainer();
        const cur = normalizeSetup(c.setup);
        const p = partial && typeof partial === 'object' ? partial : {};
        const pCopy = Object.assign({}, p);
        delete pCopy.directoryMatchByEmail;
        const mergedDir = Object.assign(
            {},
            cur.directoryMatchByEmail || {},
            p.directoryMatchByEmail && typeof p.directoryMatchByEmail === 'object' ? p.directoryMatchByEmail : {}
        );
        if (p.directoryMatchByEmailRemove && typeof p.directoryMatchByEmailRemove === 'object') {
            Object.keys(p.directoryMatchByEmailRemove).forEach(function (k) {
                const em = normEmailKey(k);
                if (em) delete mergedDir[em];
            });
        }
        const next = normalizeSetup(
            Object.assign({}, cur, pCopy, {
                matched: Object.assign({}, cur.matched, p.matched && typeof p.matched === 'object' ? p.matched : {}),
                slgDraft: Object.assign({}, cur.slgDraft, p.slgDraft && typeof p.slgDraft === 'object' ? p.slgDraft : {}),
                verwaltungDraft: Object.assign(
                    {},
                    cur.verwaltungDraft,
                    p.verwaltungDraft && typeof p.verwaltungDraft === 'object' ? p.verwaltungDraft : {}
                ),
                catalogLinks: Array.isArray(p.catalogLinks) ? p.catalogLinks : cur.catalogLinks,
                directoryMatchByEmail: mergedDir
            })
        );
        c.setup = next;
        return saveV2(c);
    }

    function getClassTeamGruppenmailForKlasse(klasseRaw) {
        const c = getContainer();
        const teams = normalizeCoreClassTeams(c.core.classTeams || []);
        for (let i = 0; i < teams.length; i++) {
            if (classTeamMatchesKlasse(teams[i], klasseRaw)) {
                return teams[i].stableMailNickname || '';
            }
        }
        return '';
    }

    function upsertClassTeam(entry) {
        const c = getContainer();
        const n = normalizeClassTeam(entry);
        if (!n) throw new Error('Klassen-Team: stableMailNickname fehlt oder ungültig.');
        let teams = normalizeCoreClassTeams(c.core.classTeams || []);
        const idx = teams.findIndex(function (t) {
            return t.stableMailNickname === n.stableMailNickname;
        });
        if (idx >= 0) teams[idx] = Object.assign({}, teams[idx], n);
        else teams.push(n);
        c.core.classTeams = normalizeCoreClassTeams(teams);
        return saveV2(c);
    }

    function touchWizardVisit(step) {
        const c = getContainer();
        const s = normalizeSetup(c.setup);
        if (typeof step === 'number' && step >= 1 && step <= 9) s.wizardStep = step;
        try {
            s.lastVisitedAt = new Date().toISOString();
        } catch {
            s.lastVisitedAt = '';
        }
        c.setup = s;
        return saveV2(c);
    }

    function exportJson() {
        return getContainer();
    }

    function importJson(obj) {
        // Accept either v2/v3 container or legacy tenant-settings v1 JSON
        const o = obj && typeof obj === 'object' ? obj : null;
        if (!o) throw new Error('Keine gültige JSON.');
        if (o.version >= 2 && o.core && o.structure && o.match) {
            return saveV2(normalizeContainer(o));
        }
        // Legacy: treat as tenant-settings-core v1 payload, update only core+current year
        const cur = getContainer();
        cur.core.domain = String(o.domain || '').trim();
        cur.core.subjects = Array.isArray(o.subjects) ? deepClone(o.subjects) : [];
        cur.core.arges = Array.isArray(o.arges) ? deepClone(o.arges) : [];
        cur.core.teachers = Array.isArray(o.teachers) ? deepClone(o.teachers) : [];
        cur.core.admin = Array.isArray(o.admin) ? deepClone(o.admin) : [];
        const y = String(cur.years.current || currentSchoolYearLabel());
        if (!cur.years.byLabel[y]) cur.years.byLabel[y] = { students: [], classes: [] };
        cur.years.byLabel[y].students = Array.isArray(o.students) ? deepClone(o.students) : [];
        cur.years.byLabel[y].classes = Array.isArray(o.classes) ? deepClone(o.classes) : [];
        if (!Array.isArray(cur.core.classTeams)) cur.core.classTeams = [];
        reconcileClassTeamsFromYearClasses(cur, y, cur.years.byLabel[y].classes);
        return saveV2(cur);
    }

    window.ms365AppDataV2 = {
        STORAGE_KEY_V2,
        VERSION,
        getContainer,
        setContainer,
        exportJson,
        importJson,
        setCoreFromTenantSettings,
        listYears,
        setCurrentYear,
        getSetup,
        patchSetup,
        touchWizardVisit,
        defaultSetup,
        normalizeSetup,
        normalizeClassTeam,
        normalizeCoreClassTeams,
        upsertClassTeam,
        getClassTeamGruppenmailForKlasse,
        reconcileClassTeamsFromYearClasses,
        mailNicknamePrefixSanitize
    };
})();

