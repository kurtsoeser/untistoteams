(function () {
    'use strict';

    function $(sel, root) {
        return (root || document).querySelector(sel);
    }

    function escapeHtml(s) {
        return String(s)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    function normCode(c) {
        return String(c || '')
            .trim()
            .toUpperCase()
            .replace(/\s+/g, '');
    }

    function toast(msg) {
        if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(msg);
        else if (typeof window.ms365ShowToast === 'function') window.ms365ShowToast(msg);
        else window.alert(msg);
    }

    function syncYearClassNames(ct, newDisplay) {
        const adv = window.ms365AppDataV2;
        if (!adv || typeof adv.getContainer !== 'function' || typeof adv.setContainer !== 'function') return;
        const c = adv.getContainer();
        const y = String(c.years && c.years.current ? c.years.current : '');
        const bucket = y && c.years && c.years.byLabel ? c.years.byLabel[y] : null;
        if (!bucket || !Array.isArray(bucket.classes)) return;
        const want = normCode(ct.classCode);
        if (!want) return;
        let changed = false;
        bucket.classes.forEach(function (cl) {
            if (normCode(cl && cl.code) === want) {
                cl.name = newDisplay;
                changed = true;
            }
        });
        if (changed) adv.setContainer(c);
    }

    function render(root, state) {
        const adv = window.ms365AppDataV2;
        const preview = window.ms365GraphRenamePreview;
        if (!adv || !preview) {
            root.innerHTML = '<p style="color:var(--muted)">Datenmodul nicht geladen.</p>';
            return;
        }
        const teams = adv.normalizeCoreClassTeams(adv.getContainer().core.classTeams || []);
        const prefix = state.prefix;

        if (!teams.length) {
            root.innerHTML =
                '<p style="margin:0;line-height:1.5;color:var(--muted)">Noch keine Einträge in <code>classTeams</code>. ' +
                'Anlegen in der <a href="einrichtung.html">Einrichtung</a> oder über die Klassenliste in den <a href="tenant.html">Schul‑Grundeinstellungen</a> (Bereich Schüler, Klassen &amp; Klassen‑Gruppen).</p>';
            return;
        }

        const rows = teams
            .map(function (t) {
                const cur = String(t.displayName || '').trim();
                const sug =
                    preview.computeNewDisplayNamePlusOne(cur, prefix) ||
                    preview.computeNewDisplayNamePlusOne(cur, 'Klasse') ||
                    '';
                const gid = String(t.graphGroupId || '').trim();
                const key = t.stableMailNickname;
                const override = state.overrides[key];
                const inputVal = override !== undefined ? override : sug || cur;
                const chk = state.selected[key] ? ' checked' : '';
                return (
                    '<tr data-nick="' +
                    escapeHtml(key) +
                    '">' +
                    '<td><input type="checkbox" class="ctr-graph-chk" data-nick="' +
                    escapeHtml(key) +
                    '"' +
                    chk +
                    (gid ? '' : ' disabled title="Keine graphGroupId"') +
                    '/></td>' +
                    '<td><code>' +
                    escapeHtml(key) +
                    '</code></td>' +
                    '<td>' +
                    escapeHtml(gid ? gid.slice(0, 8) + '…' : '—') +
                    '</td>' +
                    '<td>' +
                    escapeHtml(cur || '—') +
                    '</td>' +
                    '<td><span style="color:var(--muted)">' +
                    escapeHtml(sug || '—') +
                    '</span></td>' +
                    '<td><input type="text" class="ctr-new-dn" data-nick="' +
                    escapeHtml(key) +
                    '" value="' +
                    escapeHtml(inputVal) +
                    '" style="width:100%;min-width:160px" /></td>' +
                    '</tr>'
                );
            })
            .join('');

        root.innerHTML =
            '<div class="field" style="margin-bottom:12px">' +
            '<label for="ctrPrefix">Präfix für Stufen‑Vorschlag (z. B. <code>Klasse</code>)</label>' +
            '<div style="display:flex;gap:10px;flex-wrap:wrap;align-items:center">' +
            '<input type="text" id="ctrPrefix" value="' +
            escapeHtml(prefix) +
            '" style="max-width:220px" />' +
            '<button type="button" class="btn" id="ctrBtnSuggest">Vorschläge übernehmen</button>' +
            '<button type="button" class="btn" id="ctrBtnLocal">Neue Namen lokal speichern</button>' +
            '<button type="button" class="btn" id="ctrBtnGraph">Ausgewählte in Microsoft 365 schreiben</button>' +
            '</div>' +
            '<small style="display:block;margin-top:8px;line-height:1.45;color:var(--muted)">' +
            'Lokal: Register <code>classTeams</code> und Klassennamen im aktuellen Schuljahr (bei passendem Kürzel). ' +
            'Graph: nur <strong>displayName</strong>, nicht der Mail‑Nickname. Anmeldung oben im Header nutzen.' +
            '</small>' +
            '</div>' +
            '<div style="overflow:auto;border:1px solid var(--border);border-radius:12px">' +
            '<table style="width:100%;margin:0;font-size:0.92em;border-collapse:collapse">' +
            '<thead><tr>' +
            '<th>Graph</th><th>Stabiler Nick</th><th>Gruppe</th><th>Aktuell</th><th>Vorschlag +1</th><th>Neuer Name</th>' +
            '</tr></thead><tbody>' +
            rows +
            '</tbody></table></div>';

        const prefixInp = $('#ctrPrefix', root);
        if (prefixInp) {
            prefixInp.addEventListener('change', function () {
                state.prefix = String(prefixInp.value || '').trim() || 'Klasse';
            });
        }

        root.querySelectorAll('.ctr-new-dn').forEach(function (inp) {
            inp.addEventListener('input', function () {
                const k = inp.getAttribute('data-nick');
                if (k) state.overrides[k] = inp.value;
            });
        });
        root.querySelectorAll('.ctr-graph-chk').forEach(function (cb) {
            cb.addEventListener('change', function () {
                const k = cb.getAttribute('data-nick');
                if (k) state.selected[k] = cb.checked;
            });
        });

        $('#ctrBtnSuggest', root).addEventListener('click', function () {
            const p = prefixInp ? String(prefixInp.value || '').trim() || 'Klasse' : 'Klasse';
            state.prefix = p;
            teams.forEach(function (t) {
                const cur = String(t.displayName || '').trim();
                const sug = preview.computeNewDisplayNamePlusOne(cur, p) || '';
                if (sug) state.overrides[t.stableMailNickname] = sug;
            });
            render(root, state);
        });

        $('#ctrBtnLocal', root).addEventListener('click', function () {
            root.querySelectorAll('.ctr-new-dn').forEach(function (inp) {
                const k = inp.getAttribute('data-nick');
                const v = String(inp.value || '').trim();
                if (!k || !v) return;
                const t = teams.find(function (x) {
                    return x.stableMailNickname === k;
                });
                if (!t) return;
                adv.upsertClassTeam(Object.assign({}, t, { displayName: v }));
                syncYearClassNames(Object.assign({}, t, { displayName: v }), v);
            });
            toast('Klassen‑Anzeigenamen lokal gespeichert.');
            render(root, state);
        });

        $('#ctrBtnGraph', root).addEventListener('click', async function () {
            const G = window.ms365GraphUnifiedGroups;
            const auth = window.ms365AuthAcquireToken;
            if (!G || typeof G.patchGroupDisplayName !== 'function' || typeof auth !== 'function') {
                toast('Graph-Modul oder Anmeldung nicht verfügbar.');
                return;
            }
            const checked = Array.prototype.slice.call(root.querySelectorAll('.ctr-graph-chk:checked'));
            if (!checked.length) {
                toast('Keine Zeilen für Graph ausgewählt.');
                return;
            }
            let token;
            try {
                token = await auth(G.GRAPH_SCOPES || []);
            } catch (e) {
                toast(String((e && e.message) || e || 'Anmeldung fehlgeschlagen'));
                return;
            }
            const teamsNow = adv.normalizeCoreClassTeams(adv.getContainer().core.classTeams || []);
            const byNick = {};
            teamsNow.forEach(function (t) {
                byNick[t.stableMailNickname] = t;
            });
            let ok = 0;
            let fail = 0;
            for (let i = 0; i < checked.length; i++) {
                const cb = checked[i];
                const k = cb.getAttribute('data-nick');
                const inp = k ? root.querySelector('.ctr-new-dn[data-nick="' + k + '"]') : null;
                const v = inp ? String(inp.value || '').trim() : '';
                const t = k ? byNick[k] : null;
                const gid = t && String(t.graphGroupId || '').trim();
                if (!gid || !v) continue;
                try {
                    await G.patchGroupDisplayName(token, gid, v, undefined);
                    adv.upsertClassTeam(Object.assign({}, t, { displayName: v }));
                    syncYearClassNames(Object.assign({}, t, { displayName: v }), v);
                    ok++;
                } catch {
                    fail++;
                }
            }
            toast('Graph: ' + ok + ' OK' + (fail ? ', ' + fail + ' Fehler.' : '.'));
            render(root, state);
        });
    }

    function init() {
        const root = document.getElementById('classTeamsRolloverRoot');
        if (!root) return;

        const state = {
            prefix: 'Klasse',
            overrides: {},
            selected: {}
        };

        function refresh() {
            render(root, state);
        }

        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', refresh);
        } else refresh();

        const sel = document.getElementById('schoolYearSelect');
        if (sel) sel.addEventListener('change', refresh);
    }

    init();
})();
