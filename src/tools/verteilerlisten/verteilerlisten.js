(function () {
    'use strict';

    // Verteilerlisten + mail-aktivierte Sicherheitsgruppen
    // - Lesen: Microsoft Graph (/groups, $filter mailEnabled & securityEnabled)
    // - Schreiben (New/Set/Members/Owners/Delete): nicht über Graph möglich;
    //   wir generieren Exchange-Online-PowerShell-Scripts.

    const CACHE_KEY = 'ms365-verteilerlisten-cache-v1';
    const GRAPH_SCOPES = [
        'https://graph.microsoft.com/Group.Read.All',
        'https://graph.microsoft.com/User.Read'
    ];

    function safeJsonParse(s) {
        try { return JSON.parse(String(s)); } catch { return null; }
    }

    function getEl(id) { return document.getElementById(id); }

    function compareDe(a, b) {
        return String(a || '').localeCompare(String(b || ''), 'de', { sensitivity: 'base' });
    }

    function setProgress(on, text) {
        const el = getEl('vlProgress');
        if (!el) return;
        el.style.display = on ? '' : 'none';
        if (text) el.textContent = String(text);
    }

    async function getGraphToken(scopes) {
        if (typeof window.ms365AuthAcquireToken === 'function') {
            return await window.ms365AuthAcquireToken(scopes);
        }
        throw new Error('Bitte oben rechts anmelden (MSAL-Widget nicht verfügbar).');
    }

    function sleep(ms) { return new Promise((r) => setTimeout(r, ms)); }

    async function graphRequest(method, pathOrUrl, token, body, extraHeaders) {
        const url = pathOrUrl.indexOf('http') === 0 ? pathOrUrl : 'https://graph.microsoft.com/v1.0' + pathOrUrl;
        let attempt = 0;
        while (true) {
            const headers = { Authorization: 'Bearer ' + token };
            if (extraHeaders && typeof extraHeaders === 'object') Object.assign(headers, extraHeaders);
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
            try { data = JSON.parse(text); } catch { data = text; }
        }
        if (!res.ok) {
            const msg = typeof data === 'object' && data && data.error ? JSON.stringify(data.error) : text || String(res.status);
            throw new Error(method + ' ' + pathOrUrl + ': ' + msg);
        }
        return data || {};
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
            if (typeof onProgress === 'function') onProgress({ page, loaded: out.length, hasMore: !!next });
        }
        return out;
    }

    function loadCache() {
        try {
            const raw = localStorage.getItem(CACHE_KEY);
            const obj = raw ? safeJsonParse(raw) : null;
            const rows = obj && Array.isArray(obj.rows) ? obj.rows : [];
            return { rows, loadedAt: obj && obj.loadedAt ? String(obj.loadedAt) : '' };
        } catch {
            return { rows: [], loadedAt: '' };
        }
    }

    function saveCache(rows) {
        const out = Array.isArray(rows) ? rows : [];
        try {
            localStorage.setItem(CACHE_KEY, JSON.stringify({ rows: out, loadedAt: new Date().toISOString() }));
        } catch {
            // ignore
        }
    }

    function classifyType(g) {
        const groupTypes = Array.isArray(g.groupTypes) ? g.groupTypes : [];
        const isUnified = groupTypes.indexOf('Unified') >= 0;
        const mailEnabled = !!g.mailEnabled;
        const securityEnabled = !!g.securityEnabled;
        if (isUnified) return 'unified';
        if (mailEnabled && securityEnabled) return 'sec';
        if (mailEnabled && !securityEnabled) return 'dl';
        return 'other';
    }

    async function loadGroupsLive(onProgress) {
        const token = await getGraphToken(GRAPH_SCOPES);
        // Wichtig:
        // - hideFromAddressLists / proxyAddresses NICHT im $select - das wirft bei
        //   klassischen DLs und mail-akt. Sicherheitsgruppen "ErrorInvalidRequest".
        // - Filter auf einfache "mailEnabled eq true"; Unified-Gruppen werden
        //   clientseitig herausgefiltert (vermeidet Probleme mit not()/advanced query).
        const select = 'id,displayName,mail,mailNickname,description,mailEnabled,securityEnabled,groupTypes,visibility,createdDateTime';
        const filter = encodeURIComponent('mailEnabled eq true');
        const initial =
            '/groups?$filter=' +
            filter +
            '&$select=' +
            encodeURIComponent(select) +
            '&$top=999';
        const groups = await fetchAllPages(token, initial, onProgress);
        const mapped = (groups || [])
            .map((g) => ({
                id: String(g.id || ''),
                name: String(g.displayName || ''),
                mail: String(g.mail || ''),
                alias: String(g.mailNickname || ''),
                description: String(g.description || ''),
                kind: classifyType(g),
                hidden: false,
                createdDateTime: String(g.createdDateTime || ''),
                proxyAddresses: [],
                members: undefined,
                owners: undefined
            }))
            .filter((x) => x.id && (x.kind === 'dl' || x.kind === 'sec'));
        mapped.sort((a, b) => compareDe(a.name, b.name));
        saveCache(mapped);
        return mapped;
    }

    async function loadMembersAndOwners(groupId) {
        const token = await getGraphToken(GRAPH_SCOPES);
        const sel = encodeURIComponent('id,displayName,mail,userPrincipalName');
        const memPath = '/groups/' + encodeURIComponent(groupId) + '/members?$select=' + sel + '&$top=999';
        const ownPath = '/groups/' + encodeURIComponent(groupId) + '/owners?$select=' + sel + '&$top=999';
        const grpPath = '/groups/' + encodeURIComponent(groupId) + '?$select=hideFromAddressLists';
        const [members, owners, group] = await Promise.all([
            fetchAllPages(token, memPath).catch(() => []),
            fetchAllPages(token, ownPath).catch(() => []),
            graphJson('GET', grpPath, token).catch(() => null)
        ]);
        const norm = (x) => ({
            id: String(x.id || ''),
            name: String(x.displayName || ''),
            mail: String(x.mail || ''),
            upn: String(x.userPrincipalName || '')
        });
        return {
            members: (members || []).map(norm),
            owners: (owners || []).map(norm),
            hidden: group && typeof group.hideFromAddressLists === 'boolean' ? !!group.hideFromAddressLists : null
        };
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

    function parseLines(raw) {
        const lines = String(raw || '').split(/\r\n|\n|\r/);
        const seen = new Set();
        const out = [];
        lines.forEach((l) => {
            const t = String(l || '').trim();
            if (!t || t.startsWith('#')) return;
            const key = t.toLowerCase();
            if (seen.has(key)) return;
            seen.add(key);
            out.push(t);
        });
        return out;
    }

    function psQuote(s) {
        return '"' + String(s == null ? '' : s).replace(/"/g, '""') + '"';
    }

    function psHeader(lines) {
        lines.push('# Verteilerlisten / mail-akt. Sicherheitsgruppen (Exchange Online)');
        lines.push('# Voraussetzungen: Install-Module ExchangeOnlineManagement -Scope CurrentUser');
        lines.push('# Anmeldung: Connect-ExchangeOnline (Admin-Konto, z. B. Exchange Administrator)');
        lines.push('');
        lines.push('$ErrorActionPreference = "Stop"');
        lines.push('');
        lines.push('if (-not (Get-Module -ListAvailable ExchangeOnlineManagement)) {');
        lines.push('  Write-Host "ExchangeOnlineManagement fehlt. Installiere mit: Install-Module ExchangeOnlineManagement -Scope CurrentUser" -ForegroundColor Yellow');
        lines.push('}');
        lines.push('Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue');
        lines.push('');
        lines.push('try {');
        lines.push('  Connect-ExchangeOnline -ShowBanner:$false | Out-Null');
        lines.push('} catch {');
        lines.push('  Write-Host ("Connect-ExchangeOnline fehlgeschlagen: " + $_.Exception.Message) -ForegroundColor Red');
        lines.push('  throw');
        lines.push('}');
        lines.push('');
    }

    function psFooter(lines) {
        lines.push('');
        lines.push('Disconnect-ExchangeOnline -Confirm:$false | Out-Null');
        lines.push('Write-Host "Fertig." -ForegroundColor Green');
    }

    function buildReportScript(group) {
        const lines = [];
        psHeader(lines);
        const id = group ? (group.mail || group.alias || group.id || '') : '';
        if (!id) {
            lines.push('Write-Host "Keine Gruppe ausgewählt – nichts zu tun." -ForegroundColor Yellow');
        } else {
            lines.push('$Identity = ' + psQuote(id));
            lines.push('Write-Host ("Gruppe: " + $Identity) -ForegroundColor Cyan');
            lines.push('');
            lines.push('Write-Host "== Eigenschaften ==" -ForegroundColor Cyan');
            lines.push('Get-DistributionGroup -Identity $Identity | Format-List Name,DisplayName,PrimarySmtpAddress,Alias,GroupType,RecipientTypeDetails,ManagedBy,RequireSenderAuthenticationEnabled,ModerationEnabled,ModeratedBy,HiddenFromAddressListsEnabled,MemberJoinRestriction,MemberDepartRestriction,WhenCreatedUTC,EmailAddresses');
            lines.push('');
            lines.push('Write-Host "== Mitglieder ==" -ForegroundColor Cyan');
            lines.push('Get-DistributionGroupMember -Identity $Identity -ResultSize Unlimited | Select-Object Name,PrimarySmtpAddress,RecipientTypeDetails | Sort-Object Name | Format-Table -AutoSize');
        }
        psFooter(lines);
        return lines.join('\n');
    }

    function buildNewScript(opts) {
        const lines = [];
        psHeader(lines);
        const name = String(opts.name || '').trim();
        const alias = String(opts.alias || '').trim();
        const smtp = String(opts.smtp || '').trim();
        const type = opts.type === 'Security' ? 'Security' : 'Distribution';
        const allowExternal = !!opts.allowExternal;
        const initialMembers = Array.isArray(opts.initialMembers) ? opts.initialMembers : [];

        if (!name || !alias) {
            lines.push('Write-Host "Anzeigename und Alias sind Pflicht – Eingaben oben ergänzen." -ForegroundColor Red');
            psFooter(lines);
            return lines.join('\n');
        }

        lines.push('# === Neue ' + (type === 'Security' ? 'mail-akt. Sicherheitsgruppe' : 'Verteilerliste') + ' anlegen ===');
        lines.push('$Name = ' + psQuote(name));
        lines.push('$Alias = ' + psQuote(alias));
        if (smtp) lines.push('$Smtp = ' + psQuote(smtp));
        lines.push('');
        const newArgs = ['-Name $Name', '-DisplayName $Name', '-Alias $Alias'];
        if (smtp) newArgs.push('-PrimarySmtpAddress $Smtp');
        newArgs.push('-Type ' + (type === 'Security' ? 'Security' : 'Distribution'));
        lines.push('New-DistributionGroup ' + newArgs.join(' '));
        lines.push('');

        const setArgs = [];
        if (allowExternal) setArgs.push('-RequireSenderAuthenticationEnabled $false');
        if (setArgs.length) {
            lines.push('Set-DistributionGroup -Identity $Alias ' + setArgs.join(' '));
            lines.push('');
        }

        if (initialMembers.length) {
            lines.push('# === Initiale Mitglieder hinzufügen ===');
            lines.push('$Members = @(' + initialMembers.map(psQuote).join(', ') + ')');
            lines.push('foreach ($m in $Members) {');
            lines.push('  try {');
            lines.push('    Add-DistributionGroupMember -Identity $Alias -Member $m -ErrorAction Stop');
            lines.push('    Write-Host ("Hinzugefügt: " + $m) -ForegroundColor Green');
            lines.push('  } catch {');
            lines.push('    Write-Host ("FEHLER bei " + $m + ": " + $_.Exception.Message) -ForegroundColor Red');
            lines.push('  }');
            lines.push('}');
            lines.push('');
        }
        lines.push('Write-Host "== Ergebnis ==" -ForegroundColor Cyan');
        lines.push('Get-DistributionGroup -Identity $Alias | Format-List Name,DisplayName,PrimarySmtpAddress,Alias,GroupType,RecipientTypeDetails,RequireSenderAuthenticationEnabled');
        psFooter(lines);
        return lines.join('\n');
    }

    function buildSetScript(group, opts) {
        const lines = [];
        psHeader(lines);
        const id = group ? (group.mail || group.alias || group.id || '') : '';
        if (!id) {
            lines.push('Write-Host "Keine Gruppe ausgewählt." -ForegroundColor Red');
            psFooter(lines);
            return lines.join('\n');
        }
        lines.push('$Identity = ' + psQuote(id));
        lines.push('Write-Host ("Bearbeite: " + $Identity) -ForegroundColor Cyan');
        lines.push('');

        const setArgs = [];
        if (opts.newName) {
            lines.push('$NewName = ' + psQuote(opts.newName));
            setArgs.push('-DisplayName $NewName');
            setArgs.push('-Name $NewName');
        }
        if (opts.newSmtp) {
            lines.push('$NewSmtp = ' + psQuote(opts.newSmtp));
            setArgs.push('-PrimarySmtpAddress $NewSmtp');
        }
        if (opts.newAlias) {
            lines.push('$NewAlias = ' + psQuote(opts.newAlias));
            setArgs.push('-Alias $NewAlias');
        }
        if (opts.hidden === 'true') setArgs.push('-HiddenFromAddressListsEnabled $true');
        if (opts.hidden === 'false') setArgs.push('-HiddenFromAddressListsEnabled $false');
        if (opts.allowExternal === 'true') setArgs.push('-RequireSenderAuthenticationEnabled $false');
        if (opts.allowExternal === 'false') setArgs.push('-RequireSenderAuthenticationEnabled $true');
        if (opts.moderated === 'true') setArgs.push('-ModerationEnabled $true');
        if (opts.moderated === 'false') setArgs.push('-ModerationEnabled $false');

        const owners = Array.isArray(opts.owners) ? opts.owners : [];
        if (owners.length) {
            lines.push('$Owners = @(' + owners.map(psQuote).join(', ') + ')');
            setArgs.push('-ManagedBy $Owners');
        }

        if (!setArgs.length) {
            lines.push('Write-Host "Keine Änderungen ausgewählt – nichts zu tun." -ForegroundColor Yellow');
        } else {
            lines.push('Set-DistributionGroup -Identity $Identity ' + setArgs.join(' '));
        }
        lines.push('');
        lines.push('Write-Host "== Stand nach Änderungen ==" -ForegroundColor Cyan');
        lines.push('Get-DistributionGroup -Identity ' + (opts.newAlias ? '$NewAlias' : (opts.newSmtp ? '$NewSmtp' : '$Identity')) +
            ' | Format-List Name,DisplayName,PrimarySmtpAddress,Alias,RequireSenderAuthenticationEnabled,ModerationEnabled,HiddenFromAddressListsEnabled,ManagedBy');
        psFooter(lines);
        return lines.join('\n');
    }

    function buildMembersScript(group, deltas) {
        const lines = [];
        psHeader(lines);
        const id = group ? (group.mail || group.alias || group.id || '') : '';
        if (!id) {
            lines.push('Write-Host "Keine Gruppe ausgewählt." -ForegroundColor Red');
            psFooter(lines);
            return lines.join('\n');
        }
        lines.push('$Identity = ' + psQuote(id));
        lines.push('Write-Host ("Mitglieder pflegen: " + $Identity) -ForegroundColor Cyan');
        lines.push('');

        const adds = [];
        const rems = [];
        (deltas || []).forEach((d) => {
            if (d.startsWith('-')) rems.push(d.substring(1).trim());
            else if (d.startsWith('+')) adds.push(d.substring(1).trim());
            else adds.push(d);
        });

        if (!adds.length && !rems.length) {
            lines.push('Write-Host "Keine Mitglieder angegeben – nichts zu tun." -ForegroundColor Yellow');
        }
        if (adds.length) {
            lines.push('$ToAdd = @(' + adds.map(psQuote).join(', ') + ')');
            lines.push('foreach ($m in $ToAdd) {');
            lines.push('  try {');
            lines.push('    Add-DistributionGroupMember -Identity $Identity -Member $m -ErrorAction Stop');
            lines.push('    Write-Host ("Hinzugefügt: " + $m) -ForegroundColor Green');
            lines.push('  } catch {');
            lines.push('    Write-Host ("FEHLER bei " + $m + ": " + $_.Exception.Message) -ForegroundColor Red');
            lines.push('  }');
            lines.push('}');
            lines.push('');
        }
        if (rems.length) {
            lines.push('$ToRemove = @(' + rems.map(psQuote).join(', ') + ')');
            lines.push('foreach ($m in $ToRemove) {');
            lines.push('  try {');
            lines.push('    Remove-DistributionGroupMember -Identity $Identity -Member $m -Confirm:$false -ErrorAction Stop');
            lines.push('    Write-Host ("Entfernt: " + $m) -ForegroundColor Yellow');
            lines.push('  } catch {');
            lines.push('    Write-Host ("FEHLER bei " + $m + ": " + $_.Exception.Message) -ForegroundColor Red');
            lines.push('  }');
            lines.push('}');
            lines.push('');
        }
        lines.push('Write-Host "== Aktuelle Mitglieder ==" -ForegroundColor Cyan');
        lines.push('Get-DistributionGroupMember -Identity $Identity -ResultSize Unlimited | Select-Object Name,PrimarySmtpAddress | Sort-Object Name | Format-Table -AutoSize');
        psFooter(lines);
        return lines.join('\n');
    }

    function buildOwnersScript(group, owners) {
        const lines = [];
        psHeader(lines);
        const id = group ? (group.mail || group.alias || group.id || '') : '';
        if (!id) {
            lines.push('Write-Host "Keine Gruppe ausgewählt." -ForegroundColor Red');
            psFooter(lines);
            return lines.join('\n');
        }
        lines.push('$Identity = ' + psQuote(id));
        if (!owners.length) {
            lines.push('Write-Host "Keine Besitzer angegeben – nichts zu tun." -ForegroundColor Yellow');
        } else {
            lines.push('$Owners = @(' + owners.map(psQuote).join(', ') + ')');
            lines.push('Set-DistributionGroup -Identity $Identity -ManagedBy $Owners');
            lines.push('Write-Host ("Besitzer gesetzt für: " + $Identity) -ForegroundColor Green');
        }
        lines.push('');
        lines.push('Write-Host "== Aktuelle Besitzer ==" -ForegroundColor Cyan');
        lines.push('(Get-DistributionGroup -Identity $Identity).ManagedBy | ForEach-Object { Write-Host (" - " + $_) }');
        psFooter(lines);
        return lines.join('\n');
    }

    function buildDeleteScript(group) {
        const lines = [];
        psHeader(lines);
        const id = group ? (group.mail || group.alias || group.id || '') : '';
        if (!id) {
            lines.push('Write-Host "Keine Gruppe ausgewählt." -ForegroundColor Red');
            psFooter(lines);
            return lines.join('\n');
        }
        lines.push('$Identity = ' + psQuote(id));
        lines.push('Write-Host ("ACHTUNG: Lösche Gruppe " + $Identity) -ForegroundColor Red');
        lines.push('Write-Host "Bestätigen mit Y, abbrechen mit N." -ForegroundColor Yellow');
        lines.push('$confirm = Read-Host "Wirklich löschen? (Y/N)"');
        lines.push('if ($confirm -ne "Y" -and $confirm -ne "y") {');
        lines.push('  Write-Host "Abgebrochen." -ForegroundColor Yellow');
        lines.push('} else {');
        lines.push('  Remove-DistributionGroup -Identity $Identity -Confirm:$false');
        lines.push('  Write-Host ("Gelöscht: " + $Identity) -ForegroundColor Green');
        lines.push('}');
        psFooter(lines);
        return lines.join('\n');
    }

    function buildExoExportScript() {
        const lines = [];
        lines.push('# Export: Verteilerlisten + mail-akt. Sicherheitsgruppen (Exchange Online) -> JSON');
        lines.push('# Voraussetzungen: Install-Module ExchangeOnlineManagement -Scope CurrentUser');
        lines.push('$ErrorActionPreference = "Stop"');
        lines.push('');
        lines.push('Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue');
        lines.push('try { Connect-ExchangeOnline -ShowBanner:$false | Out-Null } catch { throw }');
        lines.push('');
        lines.push('$outFile = "distribution-groups-export.json"');
        lines.push('$items = New-Object System.Collections.Generic.List[object]');
        lines.push('Get-DistributionGroup -ResultSize Unlimited | ForEach-Object {');
        lines.push('  $g = $_');
        lines.push('  $kind = if ($g.GroupType -match "SecurityEnabled") { "sec" } else { "dl" }');
        lines.push('  $members = @()');
        lines.push('  try {');
        lines.push('    $members = Get-DistributionGroupMember -Identity $g.PrimarySmtpAddress -ResultSize Unlimited |');
        lines.push('      ForEach-Object { [pscustomobject]@{ name = $_.DisplayName; mail = $_.PrimarySmtpAddress; upn = $_.WindowsLiveID; type = $_.RecipientTypeDetails } }');
        lines.push('  } catch {');
        lines.push('    Write-Host ("Mitglieder-Lesen fehlgeschlagen für " + $g.PrimarySmtpAddress) -ForegroundColor Yellow');
        lines.push('  }');
        lines.push('  $owners = @()');
        lines.push('  if ($g.ManagedBy) { $owners = @($g.ManagedBy | ForEach-Object { [pscustomobject]@{ name = $_; mail = ""; upn = "" } }) }');
        lines.push('  $items.Add([pscustomobject]@{');
        lines.push('    id = $g.ExternalDirectoryObjectId');
        lines.push('    name = $g.DisplayName');
        lines.push('    mail = $g.PrimarySmtpAddress');
        lines.push('    alias = $g.Alias');
        lines.push('    description = $g.Description');
        lines.push('    kind = $kind');
        lines.push('    hidden = [bool]$g.HiddenFromAddressListsEnabled');
        lines.push('    requireSenderAuth = [bool]$g.RequireSenderAuthenticationEnabled');
        lines.push('    moderationEnabled = [bool]$g.ModerationEnabled');
        lines.push('    members = $members');
        lines.push('    owners = $owners');
        lines.push('  }) | Out-Null');
        lines.push('}');
        lines.push('$items | ConvertTo-Json -Depth 8 | Out-File -FilePath $outFile -Encoding UTF8');
        lines.push('Disconnect-ExchangeOnline -Confirm:$false | Out-Null');
        lines.push('Write-Host ("OK: " + $items.Count + " Gruppen -> " + $outFile) -ForegroundColor Green');
        return lines.join('\n');
    }

    function bind() {
        const cache = loadCache();
        /** @type {Array<any>} */
        let rows = cache.rows || [];
        let selectedId = '';

        function getSelected() {
            return rows.find((x) => x.id === selectedId) || null;
        }

        function filteredRows() {
            const q = String(getEl('vlSearch')?.value || '').trim().toLowerCase();
            const t = String(getEl('vlTypeFilter')?.value || 'all');
            return rows.filter((r) => {
                if (t === 'dl' && r.kind !== 'dl') return false;
                if (t === 'sec' && r.kind !== 'sec') return false;
                if (!q) return true;
                const hay = (r.name + ' ' + r.mail + ' ' + r.alias + ' ' + (r.description || '')).toLowerCase();
                return hay.includes(q);
            });
        }

        function renderList() {
            const tree = getEl('vlTree');
            if (!tree) return;
            tree.replaceChildren();
            const visible = filteredRows();
            if (!visible.length) {
                const li = document.createElement('li');
                li.style.padding = '10px 12px';
                li.style.color = '#6c757d';
                li.textContent = rows.length ? 'Keine Treffer.' : 'Noch keine Daten – „Graph einlesen“ oder „Import (JSON)“.';
                tree.appendChild(li);
                return;
            }
            visible.forEach((r) => {
                const li = document.createElement('li');
                const btn = document.createElement('button');
                btn.type = 'button';
                btn.dataset.vlSelect = r.id;
                btn.setAttribute('aria-current', selectedId === r.id ? 'true' : 'false');

                const name = document.createElement('div');
                name.style.minWidth = '0';
                name.style.flex = '1';
                name.style.fontWeight = '900';
                name.style.color = '#32325d';
                name.style.overflow = 'hidden';
                name.style.textOverflow = 'ellipsis';
                name.style.whiteSpace = 'nowrap';
                name.textContent = r.name || '(ohne Anzeigename)';

                const typePill = document.createElement('div');
                typePill.className = 'pill ' + (r.kind === 'sec' ? 'sec' : 'dl');
                typePill.textContent = r.kind === 'sec' ? 'Sicherheit (mail‑akt.)' : 'Verteilerliste';

                const mail = document.createElement('div');
                mail.className = 'pill';
                mail.textContent = r.mail || r.alias || '–';

                btn.appendChild(name);
                btn.appendChild(typePill);
                btn.appendChild(mail);

                if (r.hidden) {
                    const h = document.createElement('div');
                    h.className = 'pill hidden';
                    h.textContent = 'GAL versteckt';
                    btn.appendChild(h);
                }
                li.appendChild(btn);
                tree.appendChild(li);
            });
        }

        function renderDetail() {
            const hint = getEl('vlHint');
            const detail = getEl('vlDetail');
            const cur = getSelected();
            if (hint) hint.style.display = cur ? 'none' : '';
            if (detail) detail.style.display = cur ? '' : 'none';
            if (!cur) return;
            getEl('vlName').value = cur.name || '';
            getEl('vlMail').value = cur.mail || '';
            getEl('vlAlias').value = cur.alias || '';
            getEl('vlType').value = cur.kind === 'sec' ? 'Mail-akt. Sicherheitsgruppe' : 'Verteilerliste';
            getEl('vlId').value = cur.id || '';
            renderMembers();
        }

        function renderMembers() {
            const cur = getSelected();
            const out = getEl('vlMembersOut');
            const cnt = getEl('vlMembersCount');
            if (!out || !cur) return;
            const m = Array.isArray(cur.members) ? cur.members : null;
            const o = Array.isArray(cur.owners) ? cur.owners : null;
            if (m === null && o === null) {
                out.innerHTML = '<span class="muted">Noch nicht geladen.</span>';
                if (cnt) cnt.textContent = '';
                return;
            }
            const buildList = (label, arr) => {
                const ul = document.createElement('ul');
                (arr || []).forEach((x) => {
                    const li = document.createElement('li');
                    const id = x.upn || x.mail || x.name || '–';
                    li.textContent = (x.name ? x.name + ' ' : '') + (id ? '<' + id + '>' : '');
                    ul.appendChild(li);
                });
                if (!arr || !arr.length) {
                    const li = document.createElement('li');
                    li.className = 'muted';
                    li.textContent = '(keine)';
                    ul.appendChild(li);
                }
                const wrap = document.createElement('div');
                const head = document.createElement('div');
                head.style.fontWeight = '700';
                head.style.color = '#32325d';
                head.style.marginTop = '6px';
                head.textContent = label;
                wrap.appendChild(head);
                wrap.appendChild(ul);
                return wrap;
            };
            out.replaceChildren();
            out.appendChild(buildList('Besitzer (' + (o ? o.length : 0) + ')', o || []));
            out.appendChild(buildList('Mitglieder (' + (m ? m.length : 0) + ')', m || []));
            if (cnt) cnt.textContent = (m ? m.length : 0) + ' Mitglieder · ' + (o ? o.length : 0) + ' Besitzer';
        }

        function applyActModeUI() {
            const mode = String(getEl('vlActMode')?.value || 'report');
            // Felder ein-/ausblenden
            document.querySelectorAll('[data-act-show]').forEach((el) => {
                const show = (el.getAttribute('data-act-show') || '').split(/\s+/);
                el.style.display = show.indexOf(mode) >= 0 ? '' : 'none';
            });
            // Innerhalb des gemeinsamen Members-Feldes nur das passende Label zeigen
            document.querySelectorAll('[data-only]').forEach((el) => {
                const only = (el.getAttribute('data-only') || '').split(/\s+/);
                el.style.display = only.indexOf(mode) >= 0 ? '' : 'none';
            });
            // Type-Auswahl ist nur bei "new" relevant
            const typeField = getEl('vlActType')?.closest('.field');
            if (typeField) typeField.style.display = (mode === 'new') ? '' : 'none';
            // Members-Textarea bei delete deaktiviert
            const ta = getEl('vlActMembers');
            if (ta) ta.disabled = (mode === 'delete' || mode === 'report');
        }

        function regenerateScript() {
            const out = getEl('vlPsScript');
            if (!out) return;
            const mode = String(getEl('vlActMode')?.value || 'report');
            const cur = getSelected();
            const lines = parseLines(getEl('vlActMembers')?.value || '');
            let script = '';
            if (mode === 'report') {
                script = buildReportScript(cur);
            } else if (mode === 'new') {
                script = buildNewScript({
                    name: getEl('vlNewName')?.value || '',
                    alias: getEl('vlNewAlias')?.value || '',
                    smtp: getEl('vlNewSmtp')?.value || '',
                    type: String(getEl('vlActType')?.value || 'Distribution'),
                    allowExternal: !!getEl('vlNewExternal')?.checked,
                    initialMembers: lines
                });
            } else if (mode === 'set') {
                script = buildSetScript(cur, {
                    newName: getEl('vlSetName')?.value || '',
                    newSmtp: getEl('vlSetSmtp')?.value || '',
                    newAlias: getEl('vlSetAlias')?.value || '',
                    hidden: getEl('vlSetHidden')?.value || '',
                    allowExternal: getEl('vlSetExternal')?.value || '',
                    moderated: getEl('vlSetModerated')?.value || '',
                    owners: lines
                });
            } else if (mode === 'members') {
                script = buildMembersScript(cur, lines);
            } else if (mode === 'owners') {
                script = buildOwnersScript(cur, lines);
            } else if (mode === 'delete') {
                script = buildDeleteScript(cur);
            }
            out.value = script;
        }

        function rerender() {
            renderList();
            renderDetail();
            applyActModeUI();
            regenerateScript();
        }

        // Liste: Auswahl
        getEl('vlTree')?.addEventListener('click', (ev) => {
            const t = ev.target;
            const btn = t && t.closest ? t.closest('button[data-vl-select]') : null;
            if (!btn) return;
            const id = btn.getAttribute('data-vl-select');
            if (!id) return;
            selectedId = String(id);
            rerender();
        });

        getEl('vlSearch')?.addEventListener('input', () => renderList());
        getEl('vlTypeFilter')?.addEventListener('change', () => renderList());

        // Aktions-Eingaben → Skript live aktualisieren
        const actInputs = [
            'vlActMode', 'vlActType',
            'vlNewName', 'vlNewAlias', 'vlNewSmtp', 'vlNewExternal',
            'vlSetName', 'vlSetSmtp', 'vlSetAlias', 'vlSetHidden', 'vlSetExternal', 'vlSetModerated',
            'vlActMembers'
        ];
        actInputs.forEach((id) => {
            const el = getEl(id);
            if (!el) return;
            const evt = (el.tagName === 'SELECT' || el.type === 'checkbox') ? 'change' : 'input';
            el.addEventListener(evt, () => {
                applyActModeUI();
                regenerateScript();
            });
        });

        getEl('vlBtnNew')?.addEventListener('click', () => {
            selectedId = '';
            const sel = getEl('vlActMode');
            if (sel) sel.value = 'new';
            rerender();
            getEl('vlNewName')?.focus();
        });

        getEl('vlCopyMail')?.addEventListener('click', async () => {
            const cur = getSelected();
            const text = cur ? (cur.mail || cur.alias || '') : '';
            if (!text) return;
            try { await navigator.clipboard.writeText(text); } catch { /* ignore */ }
        });

        getEl('vlPsCopy')?.addEventListener('click', async () => {
            const ta = getEl('vlPsScript');
            const text = ta ? String(ta.value || '') : '';
            if (!text) return;
            try { await navigator.clipboard.writeText(text); } catch { /* ignore */ }
        });

        getEl('vlPsDownload')?.addEventListener('click', () => {
            const ta = getEl('vlPsScript');
            const text = ta ? String(ta.value || '') : '';
            if (!text) return;
            const cur = getSelected();
            const mode = String(getEl('vlActMode')?.value || 'report');
            const baseRaw = (cur ? (cur.alias || cur.mail || 'gruppe') : 'gruppe');
            const base = String(baseRaw).replace(/[^a-zA-Z0-9-_]/g, '');
            downloadText(`verteilerliste-${mode}-${base}.ps1`, text);
        });

        // Mitglieder/Besitzer aus Graph nachladen
        getEl('vlBtnLoadMembers')?.addEventListener('click', async () => {
            const cur = getSelected();
            if (!cur) return;
            const btn = getEl('vlBtnLoadMembers');
            if (btn) btn.disabled = true;
            try {
                setProgress(true, 'Lade Mitglieder & Besitzer aus Graph …');
                const { members, owners, hidden } = await loadMembersAndOwners(cur.id);
                cur.members = members;
                cur.owners = owners;
                if (hidden !== null) cur.hidden = hidden;
                saveCache(rows);
                renderList();
                renderMembers();
                setProgress(true, 'Fertig: ' + members.length + ' Mitglieder, ' + owners.length + ' Besitzer.');
                setTimeout(() => setProgress(false, ''), 1400);
            } catch (e) {
                setProgress(true, 'Fehler: ' + (e?.message || String(e)));
            } finally {
                if (btn) btn.disabled = false;
            }
        });

        // Graph einlesen
        async function loadNowGraph() {
            const btn = getEl('vlBtnLoadGraph');
            if (btn) btn.disabled = true;
            try {
                setProgress(true, 'Starte – lese Verteilerlisten und mail‑akt. Sicherheitsgruppen …');
                rows = await loadGroupsLive((p) => {
                    setProgress(true, `Seite ${p.page} – bisher ${p.loaded} …` + (p.hasMore ? '' : ' (fertig)'));
                });
                setProgress(true, `Fertig: ${rows.length} Einträge.`);
                setTimeout(() => setProgress(false, ''), 1600);
                selectedId = '';
                rerender();
            } catch (e) {
                setProgress(true, 'Fehler: ' + (e?.message || String(e)));
            } finally {
                if (btn) btn.disabled = false;
            }
        }
        getEl('vlBtnLoadGraph')?.addEventListener('click', loadNowGraph);

        // EXO-Export
        function showExoExport() {
            const wrap = getEl('vlExoExportWrap');
            const ta = getEl('vlExoExportScript');
            if (wrap) wrap.style.display = '';
            if (ta) ta.value = buildExoExportScript();
        }
        getEl('vlBtnLoadExo')?.addEventListener('click', showExoExport);
        getEl('vlExoHide')?.addEventListener('click', () => {
            const wrap = getEl('vlExoExportWrap');
            if (wrap) wrap.style.display = 'none';
        });
        getEl('vlExoCopy')?.addEventListener('click', async () => {
            const ta = getEl('vlExoExportScript');
            const text = ta ? String(ta.value || '') : '';
            if (!text) return;
            try { await navigator.clipboard.writeText(text); } catch { /* ignore */ }
        });
        getEl('vlExoDownload')?.addEventListener('click', () => {
            const ta = getEl('vlExoExportScript');
            const text = ta ? String(ta.value || '') : '';
            if (!text) return;
            downloadText('distribution-groups-export.ps1', text);
        });

        // Import (JSON)
        getEl('vlImportJson')?.addEventListener('change', async (e) => {
            const f = e.target.files && e.target.files[0];
            if (!f) return;
            try {
                const text = await f.text();
                const obj = safeJsonParse(text);
                const arr = Array.isArray(obj) ? obj : (obj && Array.isArray(obj.rows) ? obj.rows : null);
                if (!arr) throw new Error('Ungültiges JSON. Erwartet: Array oder { rows: [...] }.');
                rows = arr
                    .map((r) => {
                        const kind = String(r.kind || (r.recipientTypeDetails === 'MailUniversalSecurityGroup' ? 'sec' : 'dl')).toLowerCase();
                        return {
                            id: String(r.id || r.ExternalDirectoryObjectId || r.PrimarySmtpAddress || r.alias || r.name || ''),
                            name: String(r.name || r.displayName || r.DisplayName || ''),
                            mail: String(r.mail || r.PrimarySmtpAddress || ''),
                            alias: String(r.alias || r.Alias || r.mailNickname || ''),
                            description: String(r.description || r.Description || ''),
                            kind: kind === 'sec' ? 'sec' : 'dl',
                            hidden: !!(r.hidden || r.HiddenFromAddressListsEnabled),
                            members: Array.isArray(r.members) ? r.members.map((m) => ({
                                id: String(m.id || ''),
                                name: String(m.name || m.DisplayName || ''),
                                mail: String(m.mail || m.PrimarySmtpAddress || ''),
                                upn: String(m.upn || m.WindowsLiveID || m.userPrincipalName || '')
                            })) : undefined,
                            owners: Array.isArray(r.owners) ? r.owners.map((m) => ({
                                id: String(m.id || ''),
                                name: String(m.name || m.DisplayName || ''),
                                mail: String(m.mail || m.PrimarySmtpAddress || ''),
                                upn: String(m.upn || m.WindowsLiveID || m.userPrincipalName || '')
                            })) : undefined
                        };
                    })
                    .filter((x) => x.name || x.mail || x.alias);
                rows.sort((a, b) => compareDe(a.name, b.name));
                saveCache(rows);
                selectedId = '';
                rerender();
                setProgress(true, `Import OK: ${rows.length} Einträge.`);
                setTimeout(() => setProgress(false, ''), 1600);
            } catch (err) {
                setProgress(true, 'Import fehlgeschlagen: ' + (err?.message || String(err)));
            } finally {
                e.target.value = '';
            }
        });

        // Header-Slot-Button (Konsistenz mit Modul „Freigegebene Postfächer“)
        try {
            const slot = typeof window.ms365AuthGetActionSlot === 'function' ? window.ms365AuthGetActionSlot() : null;
            if (slot && !document.getElementById('vlHeaderLoadBtn')) {
                const hb = document.createElement('button');
                hb.type = 'button';
                hb.className = 'btn btn-success';
                hb.id = 'vlHeaderLoadBtn';
                hb.style.margin = '0';
                hb.style.padding = '10px 12px';
                hb.style.borderRadius = '10px';
                hb.innerHTML = '<i class="bi bi-box-seam"></i>Exchange‑Export';
                hb.addEventListener('click', showExoExport);
                slot.appendChild(hb);
            }
        } catch {
            // ignore
        }

        rerender();
    }

    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', bind);
    else bind();
})();
