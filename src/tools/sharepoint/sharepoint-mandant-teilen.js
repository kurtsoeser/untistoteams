(function () {
    'use strict';

    const G = window.ms365SpoGraph;
    if (!G) return;

    const SCOPES_READ = [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/SharePointTenantSettings.Read.All'
    ];
    const SCOPES_WRITE = [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/SharePointTenantSettings.ReadWrite.All'
    ];

    const SETTINGS_PATH = '/admin/sharepoint/settings';

    const SCOPES_SITES_HOST = [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/Sites.Read.All'
    ];

    /** Felder, die wir im Formular bearbeiten – erscheinen nicht nochmals in „Weitere Mandanten-Werte“. */
    const EDIT_KEYS = {
        sharingCapability: true,
        sharingDomainRestrictionMode: true,
        sharingAllowedDomainList: true,
        sharingBlockedDomainList: true,
        isResharingByExternalUsersEnabled: true,
        isRequireAcceptingUserToMatchInvitedUserEnabled: true
    };

    const FIELD_HELP_DE = {
        allowedDomainGuidsForSyncApp: 'Vertrauenswürdige Domains (GUIDs) für die OneDrive-Synchronisierung.',
        availableManagedPathsForSiteCreation: 'Verwaltete Pfade, unter denen neue Teamwebsites erstellt werden dürfen (nur lesen).',
        deletedUserPersonalSiteRetentionPeriodInDays: 'Aufbewahrung gelöschter OneDrive-Standorte (Tage).',
        excludedFileExtensionsForSyncApp: 'Dateiendungen, die die Sync-App nicht hochlädt.',
        imageTaggingOption: 'Bild-Kennzeichnung / Tags für die Suche (basic/enhanced/disabled).',
        isCommentingOnSitePagesEnabled: 'Kommentare auf modernen SharePoint-Seiten erlaubt.',
        isFileActivityNotificationEnabled: 'Push-Benachrichtigungen für OneDrive-Ereignisse.',
        isLegacyAuthProtocolsEnabled: 'Legacy-Authentifizierung (z. B. ältere Clients) für SharePoint/OneDrive.',
        isLoopEnabled: 'Fluid / Loop-Inhalte in SharePoint.',
        isMacSyncAppEnabled: 'OneDrive-Sync-App für Mac erlaubt.',
        isSharePointMobileNotificationEnabled: 'Mobile Push-Benachrichtigungen für SharePoint.',
        isSharePointNewsfeedEnabled: 'Newsfeed auf modernen Seiten.',
        isSiteCreationEnabled: 'Benutzer dürfen SharePoint-Websites erstellen.',
        isSiteCreationUIEnabled: 'UI zum Erstellen von Websites / freigegebener Bibliothek sichtbar.',
        isSitePagesCreationEnabled: 'Erstellung moderner Seiten auf Sites.',
        isSitesStorageLimitAutomatic: 'Speicherplatz für Sites automatisch verwalten.',
        isSyncButtonHiddenOnPersonalSite: 'Sync-Schaltfläche in OneDrive ausblenden.',
        isUnmanagedSyncAppForTenantRestricted: 'Sync nur auf in Domäne eingebundenen PCs.',
        personalSiteDefaultStorageLimitInMB: 'Standard-Speicherlimit OneDrive pro Benutzer (MB).',
        siteCreationDefaultManagedPath: 'Standard-Pfad für neue Teamwebsites.',
        siteCreationDefaultStorageLimitInMB: 'Standard-Speicherkontingent für neue Sites (MB).',
        tenantDefaultTimezone: 'Standardzeitzone für neu erstellte Sites.',
        idleSessionSignOut: 'Abmeldung bei Inaktivität (Warnung / Timeout).'
    };

    let lastRawSettings = null;

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

    function linesToList(textarea) {
        return String(textarea.value || '')
            .split(/\r?\n/)
            .map(function (s) {
                return s.trim();
            })
            .filter(Boolean);
    }

    function listToText(arr) {
        if (!Array.isArray(arr)) return '';
        return arr.join('\n');
    }

    function capLabelDe(cap) {
        const m = {
            externalUserAndGuestSharing: 'Jeder (auch anonyme Links)',
            externalUserSharingOnly: 'Neue und vorhandene Gäste (mit Anmeldung)',
            existingExternalUserSharingOnly: 'Nur bereits vorhandene Gäste',
            disabled: 'Nur Personen in Ihrer Organisation'
        };
        return m[cap] || cap;
    }

    function domainModeDe(mode) {
        const m = {
            none: 'keine Domain-Listen aktiv',
            allowList: 'nur Domains aus der Erlaubnisliste',
            blockList: 'alle Domains außer denen auf der Sperrliste'
        };
        return m[mode] || mode;
    }

    function syncLadderUi() {
        const cap = String($('fSharingCap').value || '');
        document.querySelectorAll('.spo-ladder-step').forEach(function (btn) {
            const v = btn.getAttribute('data-sharing-cap');
            btn.classList.toggle('is-active', v === cap);
        });
    }

    function renderKlartext() {
        const el = $('spoKlartext');
        const panel = $('panelKlartext');
        if (!el || !panel) return;

        const cap = String($('fSharingCap').value || '');
        const mode = String($('fDomainMode').value || '');
        const allow = linesToList($('fAllowDomains'));
        const block = linesToList($('fBlockDomains'));
        const reshare = $('fReshare').checked;
        const matchInv = $('fMatchInvite').checked;

        const lines = [];
        if (!lastRawSettings) {
            lines.push(
                '<em style="color:#64748b">Noch kein Abruf aus dem Mandanten – untenstehende Zusammenfassung zeigt nur die aktuellen Formularwerte. „Stand laden“ ausführen, um mit der Cloud zu vergleichen.</em>'
            );
        }
        lines.push(
            '<strong>Externes Teilen (Stufe):</strong> ' +
                capLabelDe(cap) +
                ' <span style="color:#64748b">(' +
                cap +
                ')</span>.'
        );
        lines.push('<strong>Domain-Modus:</strong> ' + domainModeDe(mode) + ' <span style="color:#64748b">(' + mode + ')</span>.');
        if (mode === 'allowList') {
            const preview =
                allow.length > 0
                    ? allow
                          .slice(0, 3)
                          .map(function (d) {
                              return escapeHtml(d);
                          })
                          .join('“, „')
                    : '';
            lines.push(
                '<strong>Erlaubte Domains:</strong> ' +
                    (allow.length
                        ? allow.length + ' Einträge – z. B. „' + preview + (allow.length > 3 ? '“ …' : '“')
                        : 'keine Einträge (Liste leer).')
            );
        } else if (mode === 'blockList') {
            lines.push(
                '<strong>Blockierte Domains:</strong> ' +
                    (block.length ? block.length + ' Einträge (Details in den Textfeldern).' : 'keine Einträge (Liste leer).')
            );
        } else {
            lines.push('<strong>Domain-Listen:</strong> werden im Modus „none“ nicht ausgewertet.');
        }
        lines.push(
            '<strong>Gäste erneut teilen:</strong> ' +
                (reshare ? 'erlaubt' : 'nicht erlaubt') +
                ' (<code>isResharingByExternalUsersEnabled</code>).'
        );
        lines.push(
            '<strong>Einladung = Konto:</strong> ' +
                (matchInv ? 'ja, Gast muss mit der eingeladenen Identität anmelden' : 'nein') +
                ' (<code>isRequireAcceptingUserToMatchInvitedUserEnabled</code>).'
        );

        if (lastRawSettings && lastRawSettings.tenantDefaultTimezone) {
            lines.push(
                '<strong>Hinweis aus derselben API-Antwort:</strong> Standardzeitzone für neue Sites ist <code>' +
                    escapeHtml(String(lastRawSettings.tenantDefaultTimezone)) +
                    '</code> (wird im Werkzeug „Websiteerstellung“ bearbeitet).'
            );
        }

        lines.push(
            '<strong>Standard-Freigabelink (Lesen/Bearbeiten &amp; Link-Typ):</strong> steht <em>nicht</em> in <code>sharepointSettings</code> – dafür den Abschnitt <strong>„Datei- und Ordnerlinks“</strong> weiter unten (Icons + PowerShell) oder das SharePoint Admin Center nutzen.'
        );

        el.innerHTML = '<h3>Zusammenfassung</h3><ul><li>' + lines.join('</li><li>') + '</li></ul>';
        panel.style.display = '';
    }

    function escapeHtml(s) {
        return String(s)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    /**
     * Mittlerer Block der Read-only-Kachel: Wert klar vom Beschreibungstext getrennt (kein „JaKommentare…“).
     */
    function formatValueMiddleHtml(val) {
        if (val === null || val === undefined) {
            return '<span class="spo-readonly-value">–</span>';
        }
        if (typeof val === 'boolean') {
            return '<span class="spo-readonly-value">' + (val ? 'Ja' : 'Nein') + '</span>';
        }
        if (Array.isArray(val)) {
            const t = val.length ? val.join(', ') : '(leer)';
            return '<span class="spo-readonly-value">' + escapeHtml(t) + '</span>';
        }
        if (typeof val === 'object') {
            let j;
            try {
                j = JSON.stringify(val, null, 2);
            } catch {
                j = String(val);
            }
            return '<pre class="spo-readonly-json">' + escapeHtml(j) + '</pre>';
        }
        return '<span class="spo-readonly-value">' + escapeHtml(String(val)) + '</span>';
    }

    function renderReadonlyExtras(raw) {
        const grid = $('spoReadonlyGrid');
        const panel = $('panelReadonly');
        if (!grid || !panel || !raw || typeof raw !== 'object') {
            if (panel) panel.style.display = 'none';
            return;
        }
        const keys = Object.keys(raw)
            .filter(function (k) {
                if (k.indexOf('@odata') === 0) return false;
                return !EDIT_KEYS[k];
            })
            .sort(function (a, b) {
                return a.localeCompare(b, 'de');
            });
        if (!keys.length) {
            panel.style.display = 'none';
            return;
        }
        grid.innerHTML = keys
            .map(function (k) {
                const help = FIELD_HELP_DE[k] || 'Wert aus Microsoft Graph /admin/sharepoint/settings.';
                return (
                    '<div class="spo-readonly-pill">' +
                    '<span class="k">' +
                    escapeHtml(k) +
                    '</span>' +
                    formatValueMiddleHtml(raw[k]) +
                    '<span class="d">' +
                    escapeHtml(help) +
                    '</span>' +
                    '</div>'
                );
            })
            .join('');
        panel.style.display = '';
    }

    function fillForm(s) {
        $('fSharingCap').value = String(s.sharingCapability || 'externalUserAndGuestSharing');
        $('fDomainMode').value = String(s.sharingDomainRestrictionMode || 'none');
        $('fAllowDomains').value = listToText(s.sharingAllowedDomainList);
        $('fBlockDomains').value = listToText(s.sharingBlockedDomainList);
        $('fReshare').checked = !!s.isResharingByExternalUsersEnabled;
        $('fMatchInvite').checked = !!s.isRequireAcceptingUserToMatchInvitedUserEnabled;
        syncLadderUi();
    }

    function readPatchBody() {
        return {
            sharingCapability: String($('fSharingCap').value),
            sharingDomainRestrictionMode: String($('fDomainMode').value),
            sharingAllowedDomainList: linesToList($('fAllowDomains')),
            sharingBlockedDomainList: linesToList($('fBlockDomains')),
            isResharingByExternalUsersEnabled: $('fReshare').checked,
            isRequireAcceptingUserToMatchInvitedUserEnabled: $('fMatchInvite').checked
        };
    }

    function wireLadder() {
        document.querySelectorAll('.spo-ladder-step').forEach(function (btn) {
            btn.addEventListener('click', function () {
                const v = btn.getAttribute('data-sharing-cap');
                if (!v) return;
                $('fSharingCap').value = v;
                syncLadderUi();
                renderKlartext();
            });
        });
        $('fSharingCap').addEventListener('change', function () {
            syncLadderUi();
            renderKlartext();
        });
    }

    function wireLiveSummary() {
        ['fAllowDomains', 'fBlockDomains'].forEach(function (id) {
            const n = $(id);
            if (n) n.addEventListener('input', renderKlartext);
        });
        ['fDomainMode', 'fReshare', 'fMatchInvite'].forEach(function (id) {
            const n = $(id);
            if (n) n.addEventListener('change', renderKlartext);
        });
    }

    function normalizeSpHostInput() {
        return String($('spoPsHost').value || '')
            .trim()
            .toLowerCase()
            .replace(/^https?:\/\//, '')
            .split('/')[0];
    }

    function tenantAdminUrlFromHost() {
        const h = normalizeSpHostInput();
        if (!h || h.indexOf('.sharepoint.com') === -1) return '';
        const i = h.indexOf('.sharepoint.com');
        const name = h.slice(0, i);
        if (!name) return '';
        return 'https://' + name + '-admin.sharepoint.com';
    }

    function refreshPsScript() {
        const ta = $('spoPsScript');
        if (!ta) return;
        const host = normalizeSpHostInput();
        const adminUrl = tenantAdminUrlFromHost();
        const perm = String($('spoDefPerm').value || 'View');
        const linkT = String($('spoDefLinkType').value || 'Internal');
        const hostComment = host ? '# SharePoint-Host: ' + host : '# SharePoint-Host: (oben eintragen, z. B. contoso.sharepoint.com)';
        const adminLine = adminUrl
            ? '$AdminUrl = "' + adminUrl + '"'
            : '$AdminUrl = "https://<IHRMANDANT>-admin.sharepoint.com"  # z. B. aus contoso.sharepoint.com → contoso-admin.sharepoint.com';

        ta.value = [
            '# MS365-Schulverwaltung – Standard-Freigabelinks (SharePoint Online)',
            '# Entspricht im Admin Center: Datei- und Ordnerlinks (DefaultLinkPermission / DefaultSharingLinkType)',
            hostComment,
            '',
            adminLine,
            '',
            'Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser -ErrorAction SilentlyContinue',
            'Import-Module Microsoft.Online.SharePoint.PowerShell -ErrorAction Stop',
            'Connect-SPOService -Url $AdminUrl',
            '',
            'Set-SPOTenant -DefaultLinkPermission ' + perm,
            'Set-SPOTenant -DefaultSharingLinkType ' + linkT,
            '',
            'Write-Host "Fertig: DefaultLinkPermission=' + perm + ', DefaultSharingLinkType=' + linkT + '"',
            ''
        ].join('\n');
    }

    function syncPermUi() {
        const v = String($('spoDefPerm').value || 'View');
        document.querySelectorAll('[data-def-perm]').forEach(function (b) {
            b.classList.toggle('is-active', b.getAttribute('data-def-perm') === v);
        });
    }

    function syncLinkTypeUi() {
        const v = String($('spoDefLinkType').value || 'Internal');
        document.querySelectorAll('[data-def-link]').forEach(function (b) {
            b.classList.toggle('is-active', b.getAttribute('data-def-link') === v);
        });
    }

    function wireDefaultLinkPs() {
        document.querySelectorAll('[data-def-perm]').forEach(function (btn) {
            btn.addEventListener('click', function () {
                const v = btn.getAttribute('data-def-perm');
                if (!v) return;
                $('spoDefPerm').value = v;
                syncPermUi();
                refreshPsScript();
            });
        });
        document.querySelectorAll('[data-def-link]').forEach(function (btn) {
            btn.addEventListener('click', function () {
                const v = btn.getAttribute('data-def-link');
                if (v === null || v === undefined || v === '') return;
                $('spoDefLinkType').value = v;
                syncLinkTypeUi();
                refreshPsScript();
            });
        });
        const hostEl = $('spoPsHost');
        if (hostEl) hostEl.addEventListener('input', refreshPsScript);

        const btnHost = $('btnSpoResolveHost');
        if (btnHost) {
            btnHost.addEventListener('click', function () {
                G.getGraphToken(SCOPES_SITES_HOST)
                    .then(function (token) {
                        return G.getSharePointHostname(token);
                    })
                    .then(function (hostname) {
                        if (!hostname) throw new Error('Kein Hostname aus Graph.');
                        $('spoPsHost').value = hostname;
                        refreshPsScript();
                        toast('SharePoint-Host übernommen.');
                    })
                    .catch(function (e) {
                        toast('Hostname per Graph: ' + (e && e.message ? e.message : e));
                    });
            });
        }

        const btnCopy = $('btnSpoPsCopy');
        if (btnCopy) {
            btnCopy.addEventListener('click', function () {
                const t = $('spoPsScript').value;
                if (!t) return;
                navigator.clipboard.writeText(t).then(
                    function () {
                        toast('Skript kopiert.');
                    },
                    function () {
                        toast('Kopieren fehlgeschlagen.');
                    }
                );
            });
        }
        const btnDl = $('btnSpoPsDownload');
        if (btnDl) {
            btnDl.addEventListener('click', function () {
                const t = $('spoPsScript').value;
                if (!t) return;
                const blob = new Blob([t], { type: 'text/plain;charset=utf-8' });
                const a = document.createElement('a');
                a.href = URL.createObjectURL(blob);
                a.download = 'sharepoint-default-freigabe-links.ps1';
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(a.href);
            });
        }

        syncPermUi();
        syncLinkTypeUi();
        refreshPsScript();
    }

    async function loadSettings() {
        $('spoLog').textContent = 'Lade Freigabe-Einstellungen …';
        const token = await G.getGraphToken(SCOPES_READ);
        const s = await G.graphJson('GET', SETTINGS_PATH, token, undefined, 'v1.0');
        lastRawSettings = s;
        fillForm(s);
        $('spoJson').textContent = JSON.stringify(s, null, 2);
        renderKlartext();
        renderReadonlyExtras(s);
        $('spoLog').textContent = 'Stand: ' + new Date().toLocaleString('de-AT') + ' (gelesen).';
        toast('Geladen.');
    }

    async function saveSettings() {
        $('spoLog').textContent = 'Speichern …';
        const token = await G.getGraphToken(SCOPES_WRITE);
        const body = readPatchBody();
        const updated = await G.graphJson('PATCH', SETTINGS_PATH, token, body, 'v1.0');
        lastRawSettings = updated;
        fillForm(updated);
        $('spoJson').textContent = JSON.stringify(updated, null, 2);
        renderKlartext();
        renderReadonlyExtras(updated);
        $('spoLog').textContent = 'Gespeichert: ' + new Date().toLocaleString('de-AT');
        toast('Gespeichert.');
    }

    wireLadder();
    wireLiveSummary();
    wireDefaultLinkPs();

    $('btnLoad').addEventListener('click', function () {
        loadSettings().catch(function (e) {
            $('spoLog').textContent = String(e && e.message ? e.message : e);
            toast('Fehler: ' + (e && e.message ? e.message : e));
        });
    });
    $('btnSave').addEventListener('click', function () {
        if (!window.confirm('Freigabe-Richtlinien jetzt im Mandanten speichern?')) return;
        saveSettings().catch(function (e) {
            $('spoLog').textContent = String(e && e.message ? e.message : e);
            toast('Fehler: ' + (e && e.message ? e.message : e));
        });
    });

    syncLadderUi();
    renderKlartext();
})();
