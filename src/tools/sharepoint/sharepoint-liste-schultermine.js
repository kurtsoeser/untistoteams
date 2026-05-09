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
        const el = $('sptLog');
        if (!el) return;
        el.textContent += (el.textContent ? '\n' : '') + msg;
        el.scrollTop = el.scrollHeight;
    }

    function toast(m) {
        if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
        else window.alert(m);
    }

    async function ensureToken() {
        return await G.getGraphToken(SCOPES);
    }

    /** Spalten für Power-Automate / Kalender-Sync (keine Listenelemente). */
    function columnDefsSchultermine() {
        const kategorien = [
            'Schulferien',
            'Feiertag',
            'Unterricht',
            'Prüfung',
            'Veranstaltung',
            'Elternabend',
            'Tag der offenen Tür',
            'sonstiges'
        ];
        return [
            {
                name: 'Beginn',
                displayName: 'Beginn',
                dateTime: { displayAs: 'default', format: 'dateTime' }
            },
            {
                name: 'Ende',
                displayName: 'Ende',
                dateTime: { displayAs: 'default', format: 'dateTime' }
            },
            {
                name: 'Kategorie',
                displayName: 'Kategorie',
                choice: { allowTextEntry: true, choices: kategorien }
            },
            {
                name: 'OutlookEventID',
                displayName: 'OutlookEventID',
                text: { allowMultipleLines: false, maxLength: 512 }
            },
            {
                name: 'Info',
                displayName: 'Info',
                text: { allowMultipleLines: true, maxLength: 8000 }
            },
            {
                name: 'ZeitraumText',
                displayName: 'ZeitraumText',
                text: { allowMultipleLines: false, maxLength: 255 }
            },
            {
                name: 'AllDay',
                displayName: 'AllDay',
                boolean: {}
            }
        ];
    }

    async function addColumns(siteId, listId, token) {
        const base = G.graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/columns';
        const defs = columnDefsSchultermine();
        for (let i = 0; i < defs.length; i++) {
            await G.graphJson('POST', base, token, defs[i], 'v1.0');
            await G.sleep(120);
        }
    }

    async function runCreate() {
        $('sptLog').textContent = '';
        const webUrl = String($('sptSiteUrl').value || '').trim();
        const listTitle = String($('sptListName').value || '').trim() || 'Schultermine';
        if (!webUrl) throw new Error('Bitte die Adresse der SharePoint-Website eintragen.');

        const token = await ensureToken();
        log('Löse Website auf …');
        const site = await G.resolveSiteFromWebUrl(token, webUrl);
        const siteId = site && site.id ? String(site.id) : '';
        const siteTitle = site && site.displayName ? String(site.displayName) : '';
        if (!siteId) throw new Error('Site-ID fehlt in der Graph-Antwort.');
        log('Site: ' + (siteTitle || siteId));

        log('Erstelle leere Liste „' + listTitle + '" …');
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

        log('Füge Spalten hinzu (Beginn, Ende, Kategorie, OutlookEventID, Info, ZeitraumText, AllDay) …');
        await addColumns(siteId, listId, token);
        log('Fertig – keine Zeilen angelegt (Sync z. B. per Power Automate).');
        const listWeb = created && created.webUrl ? String(created.webUrl) : '';
        if (listWeb) log('Liste: ' + listWeb);
        toast('Schultermine-Liste mit Spalten erstellt.');
    }

    $('sptBtnRun').addEventListener('click', function () {
        if (!window.confirm('Neue leere Liste auf der Website anlegen (nur Struktur, keine Termine)?')) return;
        runCreate().catch(function (e) {
            log('FEHLER: ' + (e && e.message ? e.message : String(e)));
            toast('Fehler: ' + (e && e.message ? e.message : e));
        });
    });

    $('sptBtnProbe').addEventListener('click', function () {
        $('sptLog').textContent = '';
        const webUrl = String($('sptSiteUrl').value || '').trim();
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
