/**
 * I/O-Helfer für „Schulstruktur-Sync":
 *
 *  - **CSV-Export** der Kursteams (Browser-Download oder PowerShell-Script-Vorlage).
 *  - **JSON-/Text-Download** über Blob + `<a download>` (DOM).
 *
 * Alle CSV-/Script-Builder sind **pure** Funktionen ohne Seiteneffekte – das
 * macht sie testbar und vom DOM unabhängig. `downloadJson` und `downloadText`
 * benutzen den Browser-DOM und lassen sich daher nur als Smoke-Test in JSDOM
 * sinnvoll abdecken.
 *
 * Die Abhängigkeit auf die Tool-spezifische Funktion `resolveKursteamKlasseFach`
 * (sie kennt die aktuelle Struktur-Tabelle) wird per **Dependency Injection**
 * als Callback hereingereicht – ein Wrapper im Hauptfile übergibt sie aus
 * dem laufenden Zustand.
 */

import {
    buildKursteamMailNickFromTemplate,
    buildMailNickFromLabel
} from './schulstruktur-sync-naming.js';

/**
 * RFC-4180-konformes CSV-Escaping.
 * - Felder mit `"`, `,`, `\r`, `\n` werden in `"…"` gewrappt; eingebettete `"` werden verdoppelt.
 *
 * @param {unknown} v
 * @returns {string}
 */
export function csvEscape(v) {
    const s = String(v ?? '');
    if (/[",\r\n]/.test(s)) return '"' + s.replaceAll('"', '""') + '"';
    return s;
}

/**
 * Spalten-Header für die Kursteam-CSV (Reihenfolge ist relevant für PowerShell-Skript).
 */
export const KURSTEAM_CSV_HEADER = ['DisplayName', 'MailNickname', 'Visibility', 'Target', 'Owners', 'Members'];

/**
 * Baut eine einzelne Datenzeile (Objekt) für die Kursteam-CSV.
 *
 * Erwartet **nur unfältige, deklarative Inputs** und einen Callback
 * `resolveKlasseFach(row)`, der für Kursteam-Zeilen `{ klasse, fach }`
 * aus dem aktuellen Strukturbaum liefert (siehe `structure-rules.js`).
 *
 * @param {{id?: string, typ?: string, bezeichnung?: string, ktGruppe?: string, tenantVisibility?: string, tenantTarget?: string}} row
 * @param {Record<string, { owners?: any[], members?: any[] }>} memberships
 * @param {{ kursteamYearPrefix?: string, kursteamMailNickPattern?: string }} schemaState
 * @param {(row: any) => { klasse?: string, fach?: string, hasBoth?: boolean }} [resolveKlasseFach]
 * @returns {{ DisplayName: string, MailNickname: string, Visibility: string, Target: string, Owners: string, Members: string }}
 */
export function buildKursteamCsvRow(row, memberships, schemaState, resolveKlasseFach) {
    const safeSchema = schemaState || {};
    const sug = (() => {
        const typ = String(row?.typ || '');
        const displayName = String(row?.bezeichnung || '').trim();
        if (!displayName) return { displayName: '', mailNick: '' };
        if (typ === 'Kursteam') {
            const yearPrefix = String(safeSchema.kursteamYearPrefix || '').trim();
            const kt = typeof resolveKlasseFach === 'function' ? resolveKlasseFach(row) : null;
            const klasse = kt && kt.klasse ? String(kt.klasse) : '';
            const fach = kt && kt.fach ? String(kt.fach) : '';
            const gruppe = String(row.ktGruppe || '').trim();
            const mailNick =
                klasse && fach
                    ? buildKursteamMailNickFromTemplate(safeSchema.kursteamMailNickPattern, { yearPrefix, klasse, fach, gruppe })
                    : buildMailNickFromLabel(displayName);
            return { displayName, mailNick };
        }
        return { displayName, mailNick: buildMailNickFromLabel(displayName) };
    })();

    const mem = (memberships && memberships[String(row && row.id)]) || { owners: [], members: [] };
    const owners = (mem.owners || [])
        .map((p) => String(p.userPrincipalName || p.mail || '').trim())
        .filter(Boolean)
        .join(';');
    const members = (mem.members || [])
        .map((p) => String(p.userPrincipalName || p.mail || '').trim())
        .filter(Boolean)
        .join(';');

    const visibility = String((row && row.tenantVisibility) || '').trim() || 'HiddenMembership';
    const target = String((row && row.tenantTarget) || '').trim() || 'team';
    return {
        DisplayName: sug.displayName,
        MailNickname: sug.mailNick,
        Visibility: visibility,
        Target: target,
        Owners: owners,
        Members: members
    };
}

/**
 * Stringifiziert eine Liste Kursteam-Zeilen als CSV (CRLF-Zeilenenden).
 * Header-Reihenfolge: siehe {@link KURSTEAM_CSV_HEADER}.
 *
 * @param {any[]} rows
 * @param {Record<string, { owners?: any[], members?: any[] }>} memberships
 * @param {{ kursteamYearPrefix?: string, kursteamMailNickPattern?: string }} schemaState
 * @param {(row: any) => { klasse?: string, fach?: string }} [resolveKlasseFach]
 * @returns {string} CSV mit `\r\n`-Zeilenenden.
 */
export function buildKursteamCsv(rows, memberships, schemaState, resolveKlasseFach) {
    const lines = [KURSTEAM_CSV_HEADER.join(',')];
    for (const r of rows || []) {
        const o = buildKursteamCsvRow(r, memberships, schemaState, resolveKlasseFach);
        lines.push(
            [
                csvEscape(o.DisplayName),
                csvEscape(o.MailNickname),
                csvEscape(o.Visibility),
                csvEscape(o.Target),
                csvEscape(o.Owners),
                csvEscape(o.Members)
            ].join(',')
        );
    }
    return lines.join('\r\n');
}

/**
 * PowerShell-Vorlage zum Anlegen von Teams aus einer Kursteam-CSV.
 *
 * @param {string} [csvFileName='kursteams.csv']
 * @returns {string} Skript-Text (`\n`-Zeilenenden).
 */
export function buildKursteamProvisionScript(csvFileName) {
    const file = String(csvFileName || 'kursteams.csv');
    const lines = [];
    lines.push('# Kursteams anlegen (CSV) – Microsoft Teams PowerShell');
    lines.push('# Voraussetzungen: Install-Module MicrosoftTeams -Scope CurrentUser');
    lines.push('$ErrorActionPreference = "Stop"');
    lines.push('');
    lines.push('Import-Module MicrosoftTeams -ErrorAction SilentlyContinue');
    lines.push('try { Connect-MicrosoftTeams | Out-Null } catch { throw }');
    lines.push('');
    lines.push(`$csv = Import-Csv -Path "${file}"`);
    lines.push('foreach ($r in $csv) {');
    lines.push('  if (-not $r.DisplayName -or -not $r.MailNickname) {');
    lines.push('    Write-Warning "Überspringe: DisplayName/MailNickname fehlt"');
    lines.push('    continue');
    lines.push('  }');
    lines.push('  $vis = if ($r.Visibility -eq "Public") { "Public" } else { "Private" }');
    lines.push('  Write-Host ("Lege an: " + $r.DisplayName + " (" + $r.MailNickname + ")") -ForegroundColor Cyan');
    lines.push('  $team = $null');
    lines.push('  if ($r.Target -eq "group") {');
    lines.push('    # Gruppen-Only ist in Teams PS nicht 1:1 abbildbar; wir legen ein Team an (Unified).');
    lines.push('  }');
    lines.push('  $team = New-Team -DisplayName $r.DisplayName -MailNickname $r.MailNickname -Visibility $vis');
    lines.push('  Start-Sleep -Seconds 2');
    lines.push('  $gid = $team.GroupId');
    lines.push('  if (-not $gid) { Write-Warning "Keine GroupId erhalten – weiter."; continue }');
    lines.push('');
    lines.push('  if ($r.Owners) {');
    lines.push('    $r.Owners.Split(";") | ForEach-Object {');
    lines.push('      $u = $_.Trim(); if ($u) {');
    lines.push('        try { Add-TeamUser -GroupId $gid -User $u -Role Owner | Out-Null } catch { Write-Warning ("Owner fehlgeschlagen: " + $u) }');
    lines.push('      }');
    lines.push('    }');
    lines.push('  }');
    lines.push('  if ($r.Members) {');
    lines.push('    $r.Members.Split(";") | ForEach-Object {');
    lines.push('      $u = $_.Trim(); if ($u) {');
    lines.push('        try { Add-TeamUser -GroupId $gid -User $u -Role Member | Out-Null } catch { Write-Warning ("Member fehlgeschlagen: " + $u) }');
    lines.push('      }');
    lines.push('    }');
    lines.push('  }');
    lines.push('}');
    lines.push('');
    lines.push('Disconnect-MicrosoftTeams | Out-Null');
    lines.push('Write-Host "Fertig." -ForegroundColor Green');
    return lines.join('\n');
}

/**
 * Generischer Download-Helper für Blob-Inhalte.
 *
 * @param {string} filename
 * @param {Blob} blob
 */
function triggerBlobDownload(filename, blob) {
    if (typeof document === 'undefined' || typeof URL === 'undefined') return;
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 500);
}

/**
 * Bietet ein JSON-Objekt als Download an (UTF-8, 2-Space-Pretty-Print).
 *
 * @param {string} filename
 * @param {unknown} obj
 */
export function downloadJson(filename, obj) {
    const blob = new Blob([JSON.stringify(obj, null, 2)], { type: 'application/json;charset=utf-8' });
    triggerBlobDownload(filename, blob);
}

/**
 * Bietet einen Text als Download an (UTF-8 plain).
 *
 * @param {string} filename
 * @param {string} text
 */
export function downloadText(filename, text) {
    const blob = new Blob([String(text || '')], { type: 'text/plain;charset=utf-8' });
    triggerBlobDownload(filename, blob);
}
