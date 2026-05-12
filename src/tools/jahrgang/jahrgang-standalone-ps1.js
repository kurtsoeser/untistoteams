function psEscapeSingle(s) {
    return String(s).replace(/'/g, "''");
}

/** Mehrzeiliger Text → eindeutige UPNs (pro Klasse). Gleiche Logik wie in jahrgang.js. */
function parseJgMemberLinesText(raw) {
    const lines = String(raw || '').split(/\r\n|\n|\r/);
    const seen = new Set();
    const out = [];
    lines.forEach((line) => {
        const t = String(line || '').trim();
        if (!t || t.startsWith('#')) return;
        const key = t.toLowerCase();
        if (seen.has(key)) return;
        seen.add(key);
        out.push(t);
    });
    return out;
}

export function buildStandaloneJahrgangPs1(rows, domain, standalone, createTeams, setExchangeSmtp) {
    if (createTeams === undefined) createTeams = true;
    if (setExchangeSmtp === undefined) setExchangeSmtp = true;
    const domainTrim = (domain || '').trim();
    const setExoEffective = setExchangeSmtp && domainTrim.length > 0;
    const stamp = new Date().toISOString();
    const lines = [];
    const scopesLine = '$scopes = @("Group.ReadWrite.All","User.Read.All")';

    if (standalone) {
        lines.push('#Requires -Version 5.1');
        lines.push('# Jahrgangsgruppen (M365 Unified); optional Teams ($Ms365CreateTeams); optional Exchange-SMTP ($Ms365SetExchangeSmtp)');
        lines.push('# Erzeugt in der Browser-App am ' + stamp);
        lines.push('# Daten sind unten eingebettet.');
        lines.push('');
        lines.push('[Console]::OutputEncoding = [System.Text.Encoding]::UTF8');
        lines.push('$ErrorActionPreference = "Continue"');
        lines.push('');
        lines.push('Write-Host ""');
        lines.push('Write-Host "========================================"  -ForegroundColor Cyan');
        lines.push('Write-Host "  Jahrgangsgruppen (Microsoft Graph)"   -ForegroundColor Cyan');
        lines.push('Write-Host "========================================"  -ForegroundColor Cyan');
        lines.push('Write-Host ""');
        lines.push(
            '# Meta-Modul Microsoft.Graph (einheitliche DLL-Versionen; PS-5.1 „4096 Funktionen“ per MaximumFunctionCount)'
        );
        lines.push('$MaximumFunctionCount = 32768');
        lines.push('Write-Host "Lade Microsoft.Graph ..." -ForegroundColor Gray');
        lines.push('try {');
        lines.push('    Import-Module Microsoft.Graph -ErrorAction Stop');
        lines.push('} catch {');
        lines.push('    Write-Host "Microsoft.Graph nicht gefunden – Installation (einmalig, kann einige Minuten dauern) ..." -ForegroundColor Yellow');
        lines.push('    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber');
        lines.push('    Import-Module Microsoft.Graph -ErrorAction Stop');
        lines.push('}');
        lines.push('');
        lines.push(scopesLine);
        lines.push('Write-Host "Starte Microsoft Graph-Anmeldung (Browser/Dialog oder Geraetecode) ..." -ForegroundColor Yellow');
        lines.push('Write-Host "Hinweis: Fenster ggf. im Hintergrund – Taskleiste pruefen." -ForegroundColor Gray');
        lines.push('$script:Ms365OldEap = $ErrorActionPreference');
        lines.push('$ErrorActionPreference = "Stop"');
        lines.push('try {');
        lines.push('    Connect-MgGraph -Scopes $scopes -NoWelcome');
        lines.push('} catch {');
        lines.push('    Write-Host ("Hinweis (interaktive Anmeldung): {0}" -f $_.Exception.Message) -ForegroundColor DarkYellow');
        lines.push('}');
        lines.push('$ErrorActionPreference = $script:Ms365OldEap');
        lines.push('if (-not (Get-MgContext)) {');
        lines.push('    Write-Host ""');
        lines.push('    Write-Host "Kein Graph-Kontext – Geraetecode-Anmeldung (Code erscheint unten, Browser: https://microsoft.com/devicelogin ) ..." -ForegroundColor Yellow');
        lines.push('    $ErrorActionPreference = "Stop"');
        lines.push('    try {');
        lines.push('        Connect-MgGraph -Scopes $scopes -UseDeviceAuthentication -NoWelcome');
        lines.push('    } catch {');
        lines.push('        Write-Error ("Microsoft Graph: Anmeldung fehlgeschlagen: {0}" -f $_.Exception.Message)');
        lines.push('        exit 1');
        lines.push('    }');
        lines.push('    $ErrorActionPreference = $script:Ms365OldEap');
        lines.push('}');
        lines.push('if (-not (Get-MgContext)) {');
        lines.push('    Write-Error "Microsoft Graph: Keine Sitzung – Anmeldung nicht erfolgreich. Skript wird beendet."');
        lines.push('    exit 1');
        lines.push('}');
        lines.push('$mgCtx = Get-MgContext');
        lines.push('Write-Host ("Angemeldet (Mandant: {0})" -f $mgCtx.TenantId) -ForegroundColor Green');
        lines.push('');
    } else {
        lines.push('# Microsoft Graph: Jahrgangsgruppen als Microsoft 365-Gruppen (Unified Group)');
        lines.push('# Voraussetzung: Install-Module Microsoft.Graph');
        lines.push('# https://learn.microsoft.com/powershell/module/microsoft.graph.groups/new-mggroup');
        lines.push('');
        lines.push('Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber -ErrorAction SilentlyContinue');
        lines.push('$MaximumFunctionCount = 32768');
        lines.push('try {');
        lines.push('    Import-Module Microsoft.Graph -ErrorAction Stop');
        lines.push('} catch {');
        lines.push('    Write-Host "Microsoft.Graph nicht gefunden – Installation (einmalig, kann einige Minuten dauern) ..." -ForegroundColor Yellow');
        lines.push('    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber');
        lines.push('    Import-Module Microsoft.Graph -ErrorAction Stop');
        lines.push('}');
        lines.push('');
        lines.push(scopesLine);
        lines.push('Write-Host "Starte Microsoft Graph-Anmeldung (Browser/Dialog oder Geraetecode) ..." -ForegroundColor Yellow');
        lines.push('Write-Host "Hinweis: Fenster ggf. im Hintergrund – Taskleiste pruefen." -ForegroundColor Gray');
        lines.push('$script:Ms365OldEap = $ErrorActionPreference');
        lines.push('$ErrorActionPreference = "Stop"');
        lines.push('try {');
        lines.push('    Connect-MgGraph -Scopes $scopes -NoWelcome');
        lines.push('} catch {');
        lines.push('    Write-Host ("Hinweis (interaktive Anmeldung): {0}" -f $_.Exception.Message) -ForegroundColor DarkYellow');
        lines.push('}');
        lines.push('$ErrorActionPreference = $script:Ms365OldEap');
        lines.push('if (-not (Get-MgContext)) {');
        lines.push('    Write-Host ""');
        lines.push('    Write-Host "Kein Graph-Kontext – Geraetecode-Anmeldung (Code erscheint unten, Browser: https://microsoft.com/devicelogin ) ..." -ForegroundColor Yellow');
        lines.push('    $ErrorActionPreference = "Stop"');
        lines.push('    try {');
        lines.push('        Connect-MgGraph -Scopes $scopes -UseDeviceAuthentication -NoWelcome');
        lines.push('    } catch {');
        lines.push('        throw ("Microsoft Graph: Anmeldung fehlgeschlagen: {0}" -f $_.Exception.Message)');
        lines.push('    }');
        lines.push('    $ErrorActionPreference = $script:Ms365OldEap');
        lines.push('}');
        lines.push('if (-not (Get-MgContext)) {');
        lines.push('    throw "Microsoft Graph: Keine Sitzung – Anmeldung nicht erfolgreich."');
        lines.push('}');
        lines.push('$mgCtx = Get-MgContext');
        lines.push('Write-Host ("Angemeldet (Mandant: {0})" -f $mgCtx.TenantId) -ForegroundColor Green');
        lines.push('');
    }

    lines.push('$Ms365CreateTeams = $' + (createTeams ? 'true' : 'false'));
    lines.push('$Ms365SetExchangeSmtp = $' + (setExoEffective ? 'true' : 'false'));
    lines.push("$Ms365ExchangeDomain = '" + psEscapeSingle(domainTrim) + "'");
    lines.push('');
    if (setExoEffective) {
        lines.push('$script:Ms365ExoConnected = $false');
        lines.push('function Ensure-Ms365ExchangeOnline {');
        lines.push('    if ($script:Ms365ExoConnected) { return }');
        lines.push(
            '    Write-Host "Exchange Online: Modul laden und anmelden (zweiter Dialog) …" -ForegroundColor Yellow'
        );
        lines.push('    try {');
        lines.push('        Import-Module ExchangeOnlineManagement -ErrorAction Stop');
        lines.push('    } catch {');
        lines.push('        Write-Host "Installiere ExchangeOnlineManagement …" -ForegroundColor Yellow');
        lines.push('        Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber');
        lines.push('        Import-Module ExchangeOnlineManagement -ErrorAction Stop');
        lines.push('    }');
        lines.push('    Connect-ExchangeOnline -ShowBanner:$false');
        lines.push('    $script:Ms365ExoConnected = $true');
        lines.push('    Write-Host "Exchange Online: angemeldet." -ForegroundColor Green');
        lines.push('}');
        lines.push('Ensure-Ms365ExchangeOnline');
        lines.push('');
    }
    lines.push('$rows = @(');
    rows.forEach((r, i) => {
        const last = i === rows.length - 1;
        const mems = parseJgMemberLinesText(r.memberLines || '');
        const memPart = mems.map(e => "'" + psEscapeSingle(e) + "'").join(',');
        const dn = String(r.displayName || '').trim();
        lines.push(
            "    [PSCustomObject]@{ Klasse = '" +
                psEscapeSingle(r.klasse) +
                "'; DisplayName = '" +
                psEscapeSingle(dn || r.klasse) +
                "'; MailNickname = '" +
                psEscapeSingle(r.mailNick) +
                "'; OwnerUpn = '" +
                psEscapeSingle(r.owner) +
                "'; Description = 'Jahrgangsgruppe " +
                psEscapeSingle(dn || r.klasse) +
                " (Abschluss " +
                psEscapeSingle(r.jahr) +
                ")'; MemberUpns = @(" +
                memPart +
                ') }' +
                (last ? '' : ',')
        );
    });
    lines.push(')');
    lines.push('');
    lines.push('$i = 0');
    lines.push('foreach ($r in $rows) {');
    lines.push('    $i++');
    lines.push('    try {');
    lines.push("        $owner = Get-MgUser -UserId $r.OwnerUpn -ErrorAction Stop");
    lines.push(
        '        # M365 Unified Group: New-MgGroup -BodyParameter (Bulk-Muster, vgl. https://m365corner.com/m365-powershell/using-new-mggroup-in-graph-powershell.html )'
    );
    lines.push('        $groupBody = @{');
    lines.push('            DisplayName     = $r.DisplayName');
    lines.push('            Description     = $r.Description');
    lines.push('            MailNickname    = $r.MailNickname');
    lines.push('            MailEnabled     = $true');
    lines.push('            SecurityEnabled = $false');
    lines.push('            GroupTypes      = @("Unified")');
    lines.push('            Visibility      = "Private"');
    lines.push('        }');
    lines.push('        $group = New-MgGroup -BodyParameter $groupBody -ErrorAction Stop');
    lines.push('        Start-Sleep -Seconds 2  # Replikation vor Owner-Zuweisung');
    lines.push('        New-MgGroupOwner -GroupId $group.Id -DirectoryObjectId $owner.Id');
    lines.push('        try {');
    lines.push(
        '            $memberRef = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($owner.Id)" }'
    );
    lines.push(
        '            Invoke-MgGraphRequest -Method POST -Uri (' +
            "'" +
            'https://graph.microsoft.com/v1.0/groups/{0}/members/$ref' +
            "'" +
            ' -f $group.Id) -Body ($memberRef | ConvertTo-Json -Compress) -ErrorAction Stop'
    );
    lines.push('        } catch {');
    lines.push(
        '            Write-Host ("Hinweis (Besitzer als Mitglied): {0}" -f $_.Exception.Message) -ForegroundColor DarkGray'
    );
    lines.push('        }');
    lines.push('        if ($r.MemberUpns -and $r.MemberUpns.Count -gt 0) {');
    lines.push('            foreach ($mUpn in $r.MemberUpns) {');
    lines.push('                if ([string]::IsNullOrWhiteSpace($mUpn)) { continue }');
    lines.push('                try {');
    lines.push('                    $trimUpn = $mUpn.Trim()');
    lines.push('                    $memUser = Get-MgUser -UserId $trimUpn -ErrorAction Stop');
    lines.push('                    if ($memUser.Id -eq $owner.Id) { continue }');
    lines.push(
        '                    $memberRefExtra = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($memUser.Id)" }'
    );
    lines.push(
        '                    Invoke-MgGraphRequest -Method POST -Uri ("https://graph.microsoft.com/v1.0/groups/{0}/members/$ref" -f $group.Id) -Body ($memberRefExtra | ConvertTo-Json -Compress) -ErrorAction Stop'
    );
    lines.push('                } catch {');
    lines.push(
        '                    if ($_.Exception.Message -match "already exist") {'
    );
    lines.push(
        '                        Write-Host ("  Hinweis (Mitglied {0}): bereits in der Gruppe." -f $mUpn.Trim()) -ForegroundColor DarkGray'
    );
    lines.push('                    } else {');
    lines.push(
        '                        Write-Host ("  Hinweis (Mitglied {0}): {1}" -f $mUpn.Trim(), $_.Exception.Message) -ForegroundColor DarkGray'
    );
    lines.push('                    }');
    lines.push('                }');
    lines.push('            }');
    lines.push('        }');
    lines.push('        if ($Ms365CreateTeams) {');
    lines.push('            $teamProps = @{');
    lines.push('                memberSettings = @{ allowCreatePrivateChannels = $true; allowCreateUpdateChannels = $true }');
    lines.push(
        '                messagingSettings = @{ allowUserEditMessages = $true; allowUserDeleteMessages = $true }'
    );
    lines.push('                funSettings = @{ allowGiphy = $true; giphyContentRating = "moderate" }');
    lines.push('                guestSettings = @{ allowCreateUpdateChannels = $false }');
    lines.push('            }');
    lines.push('            $teamUri = "https://graph.microsoft.com/v1.0/groups/$($group.Id)/team"');
    lines.push('            for ($ti = 0; $ti -lt 8; $ti++) {');
    lines.push('                try {');
    lines.push(
        '                    Invoke-MgGraphRequest -Method PUT -Uri $teamUri -Body $teamProps -ErrorAction Stop'
    );
    lines.push(
        '                    Write-Host ("Teams: {0} – Team bereitgestellt." -f $r.Klasse) -ForegroundColor Cyan'
    );
    lines.push('                    break');
    lines.push('                } catch {');
    lines.push('                    if ($ti -lt 7) {');
    lines.push(
        '                        Write-Host ("Teams: Warte auf Replikation ({0}/8) …" -f ($ti + 1)) -ForegroundColor DarkYellow'
    );
    lines.push('                        Start-Sleep -Seconds 10');
    lines.push('                    } else {');
    lines.push(
        '                        Write-Warning ("Teams: {0} – Team konnte nicht angelegt werden: {1}" -f $r.Klasse, $_.Exception.Message)'
    );
    lines.push('                    }');
    lines.push('                }');
    lines.push('            }');
    lines.push('        }');
    lines.push('        if ($Ms365SetExchangeSmtp -and $Ms365ExchangeDomain) {');
    lines.push('            $wantedSmtp = "$($r.MailNickname)@$Ms365ExchangeDomain"');
    lines.push('            for ($ei = 0; $ei -lt 6; $ei++) {');
    lines.push('                try {');
    lines.push(
        '                    Set-UnifiedGroup -Identity $group.Id -PrimarySmtpAddress $wantedSmtp -ErrorAction Stop'
    );
    lines.push(
        '                    Write-Host ("Exchange: {0} – PrimarySmtpAddress = {1}" -f $r.Klasse, $wantedSmtp) -ForegroundColor Green'
    );
    lines.push('                    break');
    lines.push('                } catch {');
    lines.push('                    if ($ei -lt 5) {');
    lines.push(
        '                        Write-Host ("Exchange: Warte auf Postfach ({0}/6) …" -f ($ei + 1)) -ForegroundColor DarkYellow'
    );
    lines.push('                        Start-Sleep -Seconds 15');
    lines.push('                    } else {');
    lines.push(
        '                        Write-Warning ("Exchange: {0} – PrimarySmtpAddress nicht gesetzt: {1}" -f $r.Klasse, $_.Exception.Message)'
    );
    lines.push('                    }');
    lines.push('                }');
    lines.push('            }');
    lines.push('        }');
    lines.push(
        '        Write-Host ("OK [{0}/{1}] {2} -> {3}" -f $i, $rows.Count, $r.Klasse, $r.MailNickname) -ForegroundColor Green'
    );
    lines.push('    }');
    lines.push('    catch {');
    lines.push('        $ex = $_.Exception');
    lines.push('        $detail = $ex.Message');
    lines.push('        if ($ex.InnerException) { $detail += " | " + $ex.InnerException.Message }');
    lines.push('        Write-Warning ("Fehler [{0}] {1}: {2}" -f $i, $r.Klasse, $detail)');
    lines.push('    }');
    lines.push('    Start-Sleep -Seconds 2');
    lines.push('}');
    lines.push('');
    lines.push(
        '# SMTP: Graph legt nur mailNickname an. Mit $Ms365SetExchangeSmtp wird die primäre Adresse per Exchange gesetzt.'
    );
    lines.push('# Zieldomain (App): ' + psEscapeSingle(domainTrim || domain));
    lines.push('# Set-UnifiedGroup: https://learn.microsoft.com/powershell/module/exchange/set-unifiedgroup');
    if (setExoEffective) {
        lines.push('if ($script:Ms365ExoConnected) {');
        lines.push(
            '    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}'
        );
        lines.push('}');
        lines.push('');
    }
    if (standalone) {
        lines.push('');
        lines.push('Write-Host ""');
        lines.push('Write-Host "Fertig." -ForegroundColor Cyan');
        lines.push('Read-Host "Enter druecken zum Beenden"');
    }
    return lines.join('\r\n');
}
