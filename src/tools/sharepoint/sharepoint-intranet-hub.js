(function () {
    'use strict';

    const G = window.ms365SpoGraph;
    if (!G) return;

    const SCOPES_GRAPH_SITE = [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/Sites.Read.All',
        'https://graph.microsoft.com/Sites.Create.All'
    ];

    function $(id) {
        return document.getElementById(id);
    }

    function toast(msg) {
        if (typeof window.ms365ToastOrAlert === 'function') {
            window.ms365ToastOrAlert(msg);
        } else {
            window.alert(msg);
        }
    }

    function slugify(s) {
        return String(s || '')
            .trim()
            .toLowerCase()
            .replace(/[^a-z0-9-]/g, '-')
            .replace(/-+/g, '-')
            .replace(/^-|-$/g, '');
    }

    function adminHostFromSpoHost(h) {
        const host = String(h || '').trim().toLowerCase();
        if (!host || host.indexOf('.') === -1) return '';
        const i = host.indexOf('.sharepoint.com');
        if (i === -1) return '';
        const name = host.slice(0, i);
        return name + '-admin.sharepoint.com';
    }

    function buildWebUrl(host, slug) {
        const h = String(host || '').trim().toLowerCase();
        const s = slugify(slug);
        if (!h || !s) return '';
        return 'https://' + h + '/sites/' + s;
    }

    function setPsScript(siteUrl) {
        let host = String($('fHost').value || '').trim();
        if (!host) {
            try {
                host = new URL(siteUrl).hostname;
            } catch {
                host = '';
            }
        }
        const admin = adminHostFromSpoHost(host);
        const ps =
            '# SharePoint Online PowerShell (Microsoft.Online.SharePoint.PowerShell)\n' +
            '# Modul: Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser\n' +
            'Connect-SPOService -Url https://' +
            admin +
            '\n' +
            'Register-SPOHubSite -Site ' +
            JSON.stringify(siteUrl) +
            '\n';
        $('fPsHub').value = ps;
    }

    async function detectHost() {
        $('fLog').textContent = 'Ermittle SharePoint-Host …';
        const token = await G.getGraphToken(SCOPES_GRAPH_SITE);
        const host = await G.getSharePointHostname(token);
        if (!host) throw new Error('Hostname konnte nicht ermittelt werden.');
        $('fHost').value = host;
        $('fLog').textContent = 'SharePoint-Host: ' + host;
        toast('Hostname gesetzt.');
    }

    async function createSite() {
        const host = String($('fHost').value || '').trim();
        if (!host) {
            await detectHost();
        }
        const h2 = String($('fHost').value || '').trim();
        const slug = $('fSlug').value;
        const webUrl = buildWebUrl(h2, slug);
        if (!webUrl) throw new Error('Website-Adresse ungültig (Host + Kurzname prüfen).');

        const title = String($('fTitle').value || '').trim() || 'Intranet';
        const description = String($('fDesc').value || '').trim();
        const locale = String($('fLocale').value || 'de-de').trim() || 'de-de';
        const owner = String($('fOwner').value || '').trim();
        if (!owner) throw new Error('Besitzer (UPN/E-Mail) ist erforderlich.');

        $('fLog').textContent = 'Erstelle Kommunikationswebsite (Graph beta) …';
        const token = await G.getGraphToken(SCOPES_GRAPH_SITE);

        const body = {
            name: title,
            description: description,
            webUrl: webUrl,
            locale: locale,
            shareByEmailEnabled: $('fShareByMail').checked,
            template: 'sitepagepublishing',
            ownerIdentityToResolve: { email: owner }
        };

        const res = await G.graphRequest('POST', '/sites', token, body, 'beta');
        const resText = await res.text();
        let resJson = null;
        if (resText) {
            try {
                resJson = JSON.parse(resText);
            } catch {
                resJson = { raw: resText };
            }
        }

        if (res.status !== 202 && res.status !== 200) {
            throw new Error('Site-Erstellung: HTTP ' + res.status + ' ' + (resText || ''));
        }

        let opUrl = res.headers.get('Location') || res.headers.get('location') || '';
        if (!opUrl && resJson && resJson.location) opUrl = resJson.location;
        if (opUrl && opUrl.indexOf('http') !== 0) {
            opUrl = 'https://graph.microsoft.com' + (opUrl.indexOf('/') === 0 ? '' : '/') + opUrl;
        }
        if (!opUrl) {
            $('fJson').textContent = JSON.stringify(resJson, null, 2);
            throw new Error('Keine Operation-URL (Location-Header). Antwort siehe JSON.');
        }

        $('fLog').textContent = 'Vorgang gestartet, warte auf Abschluss …\n' + opUrl;
        const done = await G.pollRichLongRunningOperation(opUrl, token);
        $('fJson').textContent = JSON.stringify({ createResponse: resJson, operation: done }, null, 2);
        $('fLastSiteUrl').value = webUrl;
        setPsScript(webUrl);
        $('fLog').textContent = 'Site bereit (laut Vorgang): ' + webUrl;
        toast('Kommunikationswebsite erstellt.');
        return webUrl;
    }

    async function registerHub() {
        const siteUrl = String($('fLastSiteUrl').value || $('fManualUrl').value || '').trim();
        if (!siteUrl) throw new Error('Zuerst Site erstellen oder Site-URL eintragen.');

        let host = String($('fHost').value || '').trim();
        if (!host) {
            try {
                host = new URL(siteUrl).hostname;
            } catch {
                host = '';
            }
        }
        if (!host) throw new Error('SharePoint-Host fehlt (Hostname ermitteln oder vollständige Site-URL eintragen).');

        $('fLog').textContent = 'Hole SharePoint-Token und versuche Hub-Registrierung (REST) …';
        const spoScope = 'https://' + host + '/Sites.FullControl.All';
        let spoToken;
        try {
            spoToken = await G.getGraphToken([spoScope]);
        } catch (e) {
            $('fLog').textContent =
                'SharePoint-Token fehlgeschlagen (fehlt API-Zustimmung für Office 365 SharePoint Online / Sites.FullControl.All?).\n' +
                String(e && e.message ? e.message : e);
            setPsScript(siteUrl);
            throw e;
        }

        try {
            const hubJson = await G.registerHubSiteViaSpoRest(siteUrl, spoToken);
            $('fHubJson').textContent = JSON.stringify(hubJson, null, 2);
            $('fLog').textContent = 'Hub-Registrierung über SharePoint REST erfolgreich.\n' + siteUrl;
            toast('Als Hub-Website registriert.');
        } catch (e) {
            $('fHubJson').textContent = String(e && e.message ? e.message : e);
            $('fLog').textContent =
                'REST-Registrierung fehlgeschlagen (häufig: Browser-CORS oder fehlende Rechte). ' +
                'Bitte PowerShell unten ausführen.\n' +
                String(e && e.message ? e.message : e);
            setPsScript(siteUrl);
            toast('Hub: REST fehlgeschlagen – PowerShell verwenden.');
        }
    }

    $('btnHost').addEventListener('click', function () {
        detectHost().catch(function (e) {
            $('fLog').textContent = String(e && e.message ? e.message : e);
            toast('Fehler: ' + (e && e.message ? e.message : e));
        });
    });

    $('btnCreate').addEventListener('click', function () {
        createSite().catch(function (e) {
            $('fLog').textContent = String(e && e.message ? e.message : e);
            toast('Fehler: ' + (e && e.message ? e.message : e));
        });
    });

    $('btnHub').addEventListener('click', function () {
        registerHub().catch(function (e) {
            $('fLog').textContent += '\n' + String(e && e.message ? e.message : e);
        });
    });

    $('btnPsCopy').addEventListener('click', function () {
        const t = $('fPsHub').value;
        if (!t) return;
        navigator.clipboard.writeText(t).then(
            function () {
                toast('PowerShell kopiert.');
            },
            function () {
                toast('Kopieren fehlgeschlagen.');
            }
        );
    });

    function refreshPreview() {
        const u = buildWebUrl($('fHost').value, $('fSlug').value);
        $('fPreviewUrl').textContent = u || '–';
    }
    ['fHost', 'fSlug'].forEach(function (id) {
        const el = $(id);
        if (!el) return;
        el.addEventListener('input', refreshPreview);
    });
    refreshPreview();
})();
