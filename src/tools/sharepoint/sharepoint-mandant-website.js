(function () {
    'use strict';

    const G = window.ms365SpoGraph;
    if (!G) {
        console.error('ms365SpoGraph fehlt (spo-graph-shared.js vor diesem Script laden).');
        return;
    }

    const SCOPES_READ = [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/SharePointTenantSettings.Read.All'
    ];
    const SCOPES_WRITE = [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/SharePointTenantSettings.ReadWrite.All'
    ];

    const SETTINGS_PATH = '/admin/sharepoint/settings';

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

    function mbToGb(mb) {
        const n = Number(mb);
        if (!isFinite(n) || n <= 0) return '';
        return (n / 1024).toFixed(2).replace(/\.?0+$/, '');
    }

    function fillForm(s) {
        $('fSiteCreate').checked = !!s.isSiteCreationEnabled;
        $('fSiteCreateUi').checked = !!s.isSiteCreationUIEnabled;
        $('fPagesCreate').checked = !!s.isSitePagesCreationEnabled;
        $('fStorageAuto').checked = !!s.isSitesStorageLimitAutomatic;

        const paths = Array.isArray(s.availableManagedPathsForSiteCreation) ? s.availableManagedPathsForSiteCreation : [];
        const sel = $('fManagedPath');
        sel.innerHTML = '';
        const current = String(s.siteCreationDefaultManagedPath || '/sites/');
        const setOpts = new Set(paths.length ? paths : ['/sites/', '/teams/']);
        setOpts.add(current);
        setOpts.forEach(function (p) {
            const o = document.createElement('option');
            o.value = p;
            o.textContent = p;
            if (p === current) o.selected = true;
            sel.appendChild(o);
        });

        $('fStorageMb').value =
            s.siteCreationDefaultStorageLimitInMB != null ? String(s.siteCreationDefaultStorageLimitInMB) : '';
        $('fStorageGbHint').textContent =
            s.siteCreationDefaultStorageLimitInMB != null
                ? '≈ ' + mbToGb(s.siteCreationDefaultStorageLimitInMB) + ' GiB (Umrechnung 1024 MiB/GiB)'
                : '–';

        $('fTimezone').value = s.tenantDefaultTimezone != null ? String(s.tenantDefaultTimezone) : '';

        $('fReadonlyPaths').textContent = paths.length ? paths.join(', ') : '–';
    }

    function readFormBody() {
        const storageMb = parseInt($('fStorageMb').value, 10);
        return {
            isSiteCreationEnabled: $('fSiteCreate').checked,
            isSiteCreationUIEnabled: $('fSiteCreateUi').checked,
            isSitePagesCreationEnabled: $('fPagesCreate').checked,
            isSitesStorageLimitAutomatic: $('fStorageAuto').checked,
            siteCreationDefaultManagedPath: String($('fManagedPath').value || '/sites/'),
            siteCreationDefaultStorageLimitInMB: isFinite(storageMb) ? storageMb : 0,
            tenantDefaultTimezone: String($('fTimezone').value || '').trim()
        };
    }

    async function loadSettings() {
        const log = $('spoLog');
        log.textContent = 'Lade Mandanteneinstellungen …';
        const token = await G.getGraphToken(SCOPES_READ);
        const s = await G.graphJson('GET', SETTINGS_PATH, token, undefined, 'v1.0');
        fillForm(s);
        $('spoJson').textContent = JSON.stringify(s, null, 2);
        log.textContent = 'Stand: ' + new Date().toLocaleString('de-AT') + ' (gelesen).';
        toast('Einstellungen geladen.');
    }

    async function saveSettings() {
        const log = $('spoLog');
        log.textContent = 'Speichern …';
        const token = await G.getGraphToken(SCOPES_WRITE);
        const body = readFormBody();
        if (!body.tenantDefaultTimezone) {
            delete body.tenantDefaultTimezone;
        }
        const updated = await G.graphJson('PATCH', SETTINGS_PATH, token, body, 'v1.0');
        fillForm(updated);
        $('spoJson').textContent = JSON.stringify(updated, null, 2);
        log.textContent = 'Gespeichert: ' + new Date().toLocaleString('de-AT');
        toast('Änderungen übernommen.');
    }

    $('btnLoad').addEventListener('click', function () {
        loadSettings().catch(function (e) {
            $('spoLog').textContent = String(e && e.message ? e.message : e);
            toast('Fehler: ' + (e && e.message ? e.message : e));
        });
    });
    $('btnSave').addEventListener('click', function () {
        if (!window.confirm('Geänderte Werte jetzt im Mandanten speichern?')) return;
        saveSettings().catch(function (e) {
            $('spoLog').textContent = String(e && e.message ? e.message : e);
            toast('Fehler: ' + (e && e.message ? e.message : e));
        });
    });

    $('fStorageMb').addEventListener('input', function () {
        const n = parseInt($('fStorageMb').value, 10);
        $('fStorageGbHint').textContent = isFinite(n) && n > 0 ? '≈ ' + mbToGb(n) + ' GiB' : '–';
    });
})();
