(function () {
    'use strict';

    function getEl(id) {
        return document.getElementById(id);
    }

    function normStr(v) {
        return String(v ?? '').trim();
    }

    function normCode(v) {
        return normStr(v).toUpperCase();
    }

    function newId() {
        if (typeof crypto !== 'undefined' && crypto.randomUUID) return crypto.randomUUID();
        return 'id-' + String(Date.now()) + '-' + String(Math.random()).slice(2, 10);
    }

    function appData() {
        return window.ms365AppDataV2;
    }

    function getRows() {
        const api = appData();
        if (!api || typeof api.getContainer !== 'function') return [];
        const c = api.getContainer();
        const rows = c && c.structure && Array.isArray(c.structure.rows) ? c.structure.rows : [];
        return rows;
    }

    function getMemberships() {
        const api = appData();
        if (!api || typeof api.getContainer !== 'function') return {};
        const c = api.getContainer();
        const m = c && c.structure && c.structure.memberships && typeof c.structure.memberships === 'object'
            ? c.structure.memberships
            : {};
        return m;
    }

    function getSettings() {
        const api = appData();
        if (!api || typeof api.getContainer !== 'function') return {};
        const c = api.getContainer();
        const s = c && c.structure && c.structure.settings && typeof c.structure.settings === 'object'
            ? c.structure.settings
            : {};
        return s;
    }

    function saveStructurePatch(patch) {
        const api = appData();
        if (!api || typeof api.getContainer !== 'function' || typeof api.setContainer !== 'function') {
            throw new Error('Lokale Daten (app-data-v2) nicht verfügbar.');
        }
        const c = api.getContainer();
        if (!c.structure) c.structure = { rows: [], memberships: {}, settings: {} };
        if (patch.rows) c.structure.rows = patch.rows;
        if (patch.memberships) c.structure.memberships = patch.memberships;
        if (patch.settings) c.structure.settings = Object.assign({}, c.structure.settings || {}, patch.settings);
        api.setContainer(c);
    }

    function uniqueSchoolYearsFromRows(rows) {
        const ys = new Set();
        rows.forEach(function (r) {
            const y = normStr(r.schuljahr);
            if (y) ys.add(y);
        });
        return Array.from(ys).sort();
    }

    function childrenOf(rows, parentId) {
        const p = String(parentId);
        return rows.filter(function (r) {
            return String(r.parentId || '') === p;
        });
    }

    /** BFS: root zuerst, dann Kinder */
    function orderedSubtreeIds(rows, rootId) {
        const out = [];
        const q = [String(rootId)];
        const seen = new Set();
        while (q.length) {
            const id = q.shift();
            if (seen.has(id)) continue;
            seen.add(id);
            out.push(id);
            childrenOf(rows, id).forEach(function (c) {
                q.push(String(c.id));
            });
        }
        return out;
    }

    function filterTopLevelSelection(rows, selectedIds) {
        const sel = new Set(selectedIds.map(String));
        const byId = {};
        rows.forEach(function (r) {
            byId[String(r.id)] = r;
        });
        function hasSelectedAncestor(id) {
            let cur = String(id);
            for (let guard = 0; guard < 5000; guard++) {
                const r = byId[cur];
                if (!r) break;
                const p = String(r.parentId || '');
                if (!p) break;
                if (sel.has(p)) return true;
                cur = p;
            }
            return false;
        }
        return selectedIds.filter(function (id) {
            return !hasSelectedAncestor(id);
        });
    }

    function duplicateSubtrees(rows, memberships, topRootIds, newSchuljahr) {
        const sj = normStr(newSchuljahr);
        if (!sj) throw new Error('Ziel-Schuljahr fehlt.');
        const idMap = {};
        const newRows = rows.slice();
        const mem = Object.assign({}, memberships);

        topRootIds.forEach(function (rootId) {
            const oldIds = orderedSubtreeIds(rows, rootId);
            oldIds.forEach(function (oid) {
                idMap[oid] = newId();
            });
            oldIds.forEach(function (oid) {
                const src = rows.find(function (x) {
                    return String(x.id) === String(oid);
                });
                if (!src) return;
                const pid = String(src.parentId || '');
                const clone = Object.assign({}, src, {
                    id: idMap[oid],
                    parentId: idMap[pid] ? idMap[pid] : pid,
                    schuljahr: sj,
                    syncStatus: 'Ausstehend',
                    letzteFehlermeldung: '',
                    status: normStr(src.status) || 'Aktiv'
                });
                newRows.push(clone);
                const om = mem[String(oid)];
                if (om && typeof om === 'object') {
                    mem[String(clone.id)] = {
                        owners: Array.isArray(om.owners) ? om.owners.slice() : [],
                        members: Array.isArray(om.members) ? om.members.slice() : []
                    };
                }
            });
        });
        return { rows: newRows, memberships: mem };
    }

    function getOrganisationAssist(settings) {
        const st = settings && settings.organisationAssist && typeof settings.organisationAssist === 'object'
            ? settings.organisationAssist
            : {};
        const cohortPlans = Array.isArray(st.cohortPlans) ? st.cohortPlans : [];
        return { cohortPlans: cohortPlans };
    }

    function setOrganisationAssist(settings, partial) {
        const cur = getOrganisationAssist(settings);
        const next = Object.assign({}, cur, partial || {});
        return Object.assign({}, settings || {}, {
            organisationAssist: next
        });
    }

    function flash(msg, ok) {
        const el = getEl('oaStatus');
        if (!el) return;
        el.textContent = msg;
        el.style.color = ok ? '#146c43' : '#842029';
        el.style.fontWeight = '700';
    }

    function renderStructureTable() {
        const tbody = getEl('oaStructTbody');
        const yearSel = getEl('oaStructYear');
        if (!tbody || !yearSel) return;

        const rows = getRows();
        const years = uniqueSchoolYearsFromRows(rows);
        const prev = normStr(yearSel.value);
        yearSel.innerHTML = '<option value="">Alle Schuljahre</option>';
        years.forEach(function (y) {
            const o = document.createElement('option');
            o.value = y;
            o.textContent = y;
            yearSel.appendChild(o);
        });
        if (prev && years.indexOf(prev) >= 0) yearSel.value = prev;

        const yf = normStr(yearSel.value);
        const filt = rows.filter(function (r) {
            return !yf || normStr(r.schuljahr) === yf;
        });

        tbody.innerHTML = '';
        filt.forEach(function (r) {
            const tr = document.createElement('tr');
            const td0 = document.createElement('td');
            const cb = document.createElement('input');
            cb.type = 'checkbox';
            cb.className = 'oa-row-cb';
            cb.dataset.rowId = String(r.id);
            td0.appendChild(cb);

            function td(text) {
                const t = document.createElement('td');
                t.textContent = text;
                return t;
            }

            tr.appendChild(td0);
            tr.appendChild(td(normStr(r.bezeichnung)));
            tr.appendChild(td(normStr(r.typ)));
            tr.appendChild(td(normStr(r.schuljahr)));
            tr.appendChild(td(normStr(r.status)));
            tr.appendChild(td(normStr(r.syncStatus)));
            tbody.appendChild(tr);
        });

        getEl('oaStructCount').textContent = String(filt.length);
    }

    function selectedRowIds() {
        return Array.from(document.querySelectorAll('.oa-row-cb:checked')).map(function (x) {
            return x.dataset.rowId;
        });
    }

    function renderCohortList() {
        const ul = getEl('oaCohortList');
        if (!ul) return;
        const settings = getSettings();
        const { cohortPlans } = getOrganisationAssist(settings);
        ul.innerHTML = '';
        if (!cohortPlans.length) {
            const li = document.createElement('li');
            li.className = 'muted';
            li.textContent = 'Noch keine Kohorten gespeichert.';
            ul.appendChild(li);
            return;
        }
        cohortPlans.forEach(function (p) {
            const li = document.createElement('li');
            li.textContent =
                (p.kind === 'eltern' ? 'Eltern-Jahrgang' : 'Matura / Abschluss') +
                ' · Abschlussjahr ' +
                String(p.graduationYear || '') +
                ' · ' +
                String(p.displayName || '') +
                ' · mailNickname: ' +
                String(p.mailNickname || '');
            ul.appendChild(li);
        });
    }

    function renderFachTable() {
        const tbody = getEl('oaFachTbody');
        if (!tbody) return;
        const api = appData();
        if (!api || typeof api.getContainer !== 'function' || typeof api.getSetup !== 'function') {
            tbody.innerHTML = '';
            return;
        }
        const c = api.getContainer();
        const subjects = Array.isArray(c.core && c.core.subjects) ? c.core.subjects : [];
        const setup = api.getSetup();
        const prefix = normStr(setup.subjectGroupMailPrefix) || 'fach';
        const links = Array.isArray(setup.catalogLinks) ? setup.catalogLinks : [];
        const byCode = {};
        links.forEach(function (L) {
            if (L && L.kind === 'subject' && L.code) byCode[normCode(L.code)] = L;
        });

        tbody.innerHTML = '';
        subjects.forEach(function (s) {
            const code = normCode(s.code);
            const name = normStr(s.name);
            const tail = api.mailNicknamePrefixSanitize
                ? api.mailNicknamePrefixSanitize(code.replace(/[^0-9A-Za-z._-]/g, '').toLowerCase(), 40)
                : code.toLowerCase();
            const nick = api.mailNicknamePrefixSanitize
                ? api.mailNicknamePrefixSanitize(prefix + '-' + tail, 60)
                : prefix + '-' + tail;
            const ex = byCode[code];
            const tr = document.createElement('tr');
            function td(t) {
                const x = document.createElement('td');
                x.textContent = t;
                return x;
            }
            tr.appendChild(td(code));
            tr.appendChild(td(name));
            tr.appendChild(td(nick));
            tr.appendChild(td(ex && ex.graphGroupId ? 'verknüpft' : ex ? 'Entwurf' : '–'));
            const tda = document.createElement('td');
            const cb = document.createElement('input');
            cb.type = 'checkbox';
            cb.className = 'oa-fach-cb';
            cb.dataset.code = code;
            cb.dataset.nick = nick;
            cb.dataset.dname = 'Fachschaft ' + (name || code);
            if (ex) cb.disabled = true;
            tda.appendChild(cb);
            tr.appendChild(tda);
            tbody.appendChild(tr);
        });
    }

    function renderClassTeamsHint() {
        const el = getEl('oaClassTeamsHint');
        if (!el) return;
        const api = appData();
        if (!api || typeof api.getContainer !== 'function') {
            el.textContent = '';
            return;
        }
        const teams = api.normalizeCoreClassTeams
            ? api.normalizeCoreClassTeams(api.getContainer().core.classTeams || [])
            : [];
        el.textContent = teams.length
            ? 'Lokal gespeicherte Kursteam-Zuordnungen: ' + teams.length + ' Einträge (Details im Werkzeug Kursteams).'
            : 'Noch keine Kursteam-Zuordnungen in den lokalen Daten.';
    }

    function bind() {
        getEl('oaStructYear')?.addEventListener('change', function () {
            renderStructureTable();
        });
        getEl('oaStructRefresh')?.addEventListener('click', function () {
            renderStructureTable();
            renderClassTeamsHint();
            flash('Tabelle aktualisiert.', true);
        });

        getEl('oaSelectAll')?.addEventListener('click', function () {
            document.querySelectorAll('.oa-row-cb').forEach(function (x) {
                x.checked = true;
            });
        });
        getEl('oaSelectNone')?.addEventListener('click', function () {
            document.querySelectorAll('.oa-row-cb').forEach(function (x) {
                x.checked = false;
            });
        });

        getEl('oaBulkYearBtn')?.addEventListener('click', function () {
            try {
                const ids = selectedRowIds();
                const tgt = normStr(getEl('oaBulkYear').value);
                if (!ids.length) throw new Error('Keine Zeilen markiert.');
                if (!tgt) throw new Error('Ziel-Schuljahr fehlt.');
                const rows = getRows().map(function (r) {
                    if (ids.indexOf(String(r.id)) >= 0) return Object.assign({}, r, { schuljahr: tgt });
                    return r;
                });
                saveStructurePatch({ rows: rows });
                renderStructureTable();
                flash('Schuljahr für ' + ids.length + ' Zeile(n) gesetzt.', true);
            } catch (e) {
                flash(e.message || String(e), false);
            }
        });

        getEl('oaArchiveBtn')?.addEventListener('click', function () {
            try {
                const ids = selectedRowIds();
                if (!ids.length) throw new Error('Keine Zeilen markiert.');
                const rows = getRows().map(function (r) {
                    if (ids.indexOf(String(r.id)) >= 0) return Object.assign({}, r, { status: 'Inaktiv' });
                    return r;
                });
                saveStructurePatch({ rows: rows });
                renderStructureTable();
                flash(ids.length + ' Zeile(n) als Inaktiv markiert.', true);
            } catch (e) {
                flash(e.message || String(e), false);
            }
        });

        getEl('oaDupBtn')?.addEventListener('click', function () {
            try {
                const ids = selectedRowIds();
                const tgt = normStr(getEl('oaDupYear').value);
                if (!ids.length) throw new Error('Keine Zeilen markiert.');
                if (!tgt) throw new Error('Ziel-Schuljahr für Kopien fehlt.');
                const rows = getRows();
                const mem = getMemberships();
                const roots = filterTopLevelSelection(rows, ids);
                const { rows: nextRows, memberships: nextMem } = duplicateSubtrees(rows, mem, roots, tgt);
                saveStructurePatch({ rows: nextRows, memberships: nextMem });
                renderStructureTable();
                flash('Unterbäume dupliziert: ' + roots.length + ' Wurzel(n).', true);
            } catch (e) {
                flash(e.message || String(e), false);
            }
        });

        getEl('oaCohortAdd')?.addEventListener('click', function () {
            try {
                const api = appData();
                if (!api || typeof api.getContainer !== 'function' || typeof api.setContainer !== 'function') {
                    throw new Error('app-data-v2 fehlt.');
                }
                const kind = getEl('oaCohortKind').value === 'eltern' ? 'eltern' : 'matura';
                const gy = normStr(getEl('oaCohortYear').value);
                if (!/^\d{4}$/.test(gy)) throw new Error('Abschlussjahr als vierstellige Zahl (z. B. 2026).');
                const prefix = kind === 'eltern' ? 'eltern-jg' : 'matura-jg';
                const nick = api.mailNicknamePrefixSanitize
                    ? api.mailNicknamePrefixSanitize(prefix + gy, 60)
                    : prefix + gy;
                const dname =
                    kind === 'eltern'
                        ? 'Elternschaft Abschlussjahrgang ' + gy
                        : 'Abschlussjahrgang / Matura ' + gy;
                const c = api.getContainer();
                const settings = Object.assign({}, c.structure && c.structure.settings ? c.structure.settings : {});
                const oa = getOrganisationAssist(settings);
                const plans = oa.cohortPlans.slice();
                plans.push({
                    id: newId(),
                    kind: kind,
                    graduationYear: gy,
                    displayName: dname,
                    mailNickname: nick,
                    savedAt: new Date().toISOString()
                });
                const nextSettings = setOrganisationAssist(settings, { cohortPlans: plans });
                saveStructurePatch({ settings: nextSettings });
                renderCohortList();
                flash('Kohorte gespeichert (lokal, Planung).', true);
            } catch (e) {
                flash(e.message || String(e), false);
            }
        });

        getEl('oaCohortClear')?.addEventListener('click', function () {
            try {
                const api = appData();
                if (!api || typeof api.getContainer !== 'function' || typeof api.setContainer !== 'function') {
                    throw new Error('app-data-v2 fehlt.');
                }
                const c = api.getContainer();
                const settings = Object.assign({}, c.structure && c.structure.settings ? c.structure.settings : {});
                const nextSettings = setOrganisationAssist(settings, { cohortPlans: [] });
                saveStructurePatch({ settings: nextSettings });
                renderCohortList();
                flash('Kohortenliste geleert.', true);
            } catch (e) {
                flash(e.message || String(e), false);
            }
        });

        getEl('oaFachRefresh')?.addEventListener('click', function () {
            renderFachTable();
            flash('Fächerliste aktualisiert.', true);
        });

        getEl('oaFachAddLinks')?.addEventListener('click', function () {
            try {
                const api = appData();
                if (!api || typeof api.getSetup !== 'function' || typeof api.patchSetup !== 'function') {
                    throw new Error('Einrichtungsdaten nicht verfügbar.');
                }
                const boxes = Array.from(document.querySelectorAll('.oa-fach-cb:checked'));
                if (!boxes.length) throw new Error('Keine Fächer ausgewählt (bereits verknüpfte sind deaktiviert).');
                const cur = api.getSetup();
                const links = Array.isArray(cur.catalogLinks) ? cur.catalogLinks.slice() : [];
                const have = {};
                links.forEach(function (L) {
                    if (L && L.kind === 'subject' && L.code) have[normCode(L.code)] = true;
                });
                let added = 0;
                boxes.forEach(function (cb) {
                    const code = normCode(cb.dataset.code);
                    if (!code || have[code]) return;
                    have[code] = true;
                    links.push({
                        kind: 'subject',
                        code: code,
                        graphGroupId: '',
                        displayName: normStr(cb.dataset.dname) || 'Fachschaft ' + code,
                        mailNickname: normStr(cb.dataset.nick),
                        mode: '',
                        syncStatus: ''
                    });
                    added++;
                });
                api.patchSetup({ catalogLinks: links });
                renderFachTable();
                flash(added + ' Fachschafts-Link(s) in den Einrichtungsdaten ergänzt (Gruppe in M365 separat anlegen).', true);
            } catch (e) {
                flash(e.message || String(e), false);
            }
        });
    }

    function showTab(which) {
        ['sj', 'coh', 'fach'].forEach(function (w) {
            const tab = getEl('oaTab' + w);
            const panel = getEl('oaPanel' + w);
            if (tab) tab.setAttribute('aria-selected', w === which ? 'true' : 'false');
            if (panel) panel.hidden = w !== which;
        });
    }

    document.querySelectorAll('.oa-tab').forEach(function (btn) {
        btn.addEventListener('click', function () {
            const w = btn.getAttribute('data-oa-tab');
            if (w) showTab(w);
        });
    });

    bind();
    showTab('sj');
    renderStructureTable();
    renderCohortList();
    renderFachTable();
    renderClassTeamsHint();
})();
