(function () {
    'use strict';

    const G = window.ms365SpoGraph;
    if (!G) return;

    const SCOPES = [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/Sites.ReadWrite.All'
    ];

    function $(id) {
        return document.getElementById(id);
    }

    function log(msg) {
        const el = $('splLog');
        if (!el) return;
        el.textContent += (el.textContent ? '\n' : '') + msg;
        el.scrollTop = el.scrollHeight;
    }

    function toast(m) {
        if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
        else window.alert(m);
    }

    function loadTeachers() {
        if (typeof window.ms365TenantSettingsLoad !== 'function') {
            throw new Error('Stammdaten nicht geladen (tenant-settings-core.js fehlt?).');
        }
        const s = window.ms365TenantSettingsLoad();
        const teachers = (s && Array.isArray(s.teachers) ? s.teachers : []).filter(function (t) {
            return t && String(t.code || '').trim();
        });
        return teachers;
    }

    async function ensureToken() {
        return await G.getGraphToken(SCOPES);
    }

    async function addColumnsLehrer(siteId, listId, token) {
        const base = G.graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/columns';
        const defs = [
            {
                name: 'LehrerCode',
                displayName: 'Kürzel',
                text: { allowMultipleLines: false, maxLength: 40 }
            },
            {
                name: 'EMail',
                displayName: 'E-Mail',
                text: { allowMultipleLines: false, maxLength: 255 }
            },
            {
                name: 'UPN',
                displayName: 'UPN',
                text: { allowMultipleLines: false, maxLength: 255 }
            }
        ];
        for (let i = 0; i < defs.length; i++) {
            await G.graphJson('POST', base, token, defs[i], 'v1.0');
            await G.sleep(120);
        }
    }

    async function runCreate() {
        $('splLog').textContent = '';
        const webUrl = String($('splSiteUrl').value || '').trim();
        const listTitle = String($('splListName').value || '').trim() || 'Lehrerinnen';
        if (!webUrl) throw new Error('Bitte die Adresse der SharePoint-Website eintragen.');

        const teachers = loadTeachers();
        if (!teachers.length) {
            throw new Error('Keine Lehrkräfte in den Schul-Grundeinstellungen – zuerst unter Stammdaten pflegen.');
        }

        log('Lehrkräfte aus lokalem Speicher: ' + teachers.length);
        const token = await ensureToken();
        log('Löse Website auf …');
        const site = await G.resolveSiteFromWebUrl(token, webUrl);
        const siteId = site && site.id ? String(site.id) : '';
        const siteTitle = site && site.displayName ? String(site.displayName) : '';
        if (!siteId) throw new Error('Site-ID fehlt in der Graph-Antwort.');
        log('Site: ' + (siteTitle || siteId));

        log('Erstelle Liste „' + listTitle + '" …');
        const created = await G.graphJson(
            'POST',
            G.graphPathSite(siteId) + '/lists',
            token,
            {
                displayName: listTitle,
                list: { template: 'genericList' }
            },
            'v1.0'
        );
        const listId = created && created.id ? String(created.id) : '';
        if (!listId) throw new Error('Listen-ID fehlt in der Antwort.');
        log('Liste angelegt, ID: ' + listId);

        log('Füge Spalten hinzu (Kürzel, E-Mail, UPN) …');
        await addColumnsLehrer(siteId, listId, token);
        log('Spalten fertig.');

        const itemsPath = G.graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/items';
        let ok = 0;
        for (let i = 0; i < teachers.length; i++) {
            const t = teachers[i];
            const email = String(t.email || '').trim();
            const name = String(t.name || '').trim() || String(t.code || '').trim();
            const code = String(t.code || '').trim();
            await G.graphJson(
                'POST',
                itemsPath,
                token,
                {
                    fields: {
                        Title: name,
                        LehrerCode: code,
                        EMail: email,
                        UPN: email
                    }
                },
                'v1.0'
            );
            ok++;
            if (ok % 10 === 0) log('… ' + ok + ' Zeilen geschrieben');
            await G.sleep(80);
        }
        log('Fertig: ' + ok + ' Lehrkräfte als Listenelemente.');
        const listWeb = created && created.webUrl ? String(created.webUrl) : '';
        if (listWeb) log('Liste im Browser: ' + listWeb);
        toast('Lehrerliste erstellt und befüllt.');
    }

    $('splBtnRun').addEventListener('click', function () {
        if (!window.confirm('Neue Liste auf der angegebenen Website anlegen und alle Lehrkräfte aus den Grundeinstellungen eintragen?')) return;
        runCreate().catch(function (e) {
            log('FEHLER: ' + (e && e.message ? e.message : String(e)));
            toast('Fehler: ' + (e && e.message ? e.message : e));
        });
    });

    $('splBtnProbe').addEventListener('click', function () {
        $('splLog').textContent = '';
        const webUrl = String($('splSiteUrl').value || '').trim();
        if (!webUrl) {
            toast('Website-URL fehlt.');
            return;
        }
        ensureToken()
            .then(function (token) {
                return G.resolveSiteFromWebUrl(token, webUrl);
            })
            .then(function (site) {
                log('Site gefunden: ' + (site.displayName || '') + '\nid: ' + (site.id || ''));
                if (site.webUrl) log('webUrl: ' + site.webUrl);
                toast('Website erkannt.');
            })
            .catch(function (e) {
                log('FEHLER: ' + (e && e.message ? e.message : String(e)));
                toast('Fehler: ' + (e && e.message ? e.message : e));
            });
    });
})();
