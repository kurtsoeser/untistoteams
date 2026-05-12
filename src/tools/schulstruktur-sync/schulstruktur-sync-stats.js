/**
 * Statistik-/Filter-/Format-Helfer für „Schulstruktur-Sync".
 *
 * Aus `schulstruktur-sync.js` 1:1 ausgelagert (Phase 2 Schnitt 8).
 *
 * Pure Funktionen (gut testbar):
 *  - {@link formatDateTimeAT}: ISO-Datum → de-AT-Formatierung.
 *  - {@link pillClass}: SyncStatus → CSS-Pill-Klasse (`ok`/`warn`/`err`).
 *  - {@link computeStats}: SOLL-Aggregat (`total`, `aktiv`, `abw`, `err`).
 *  - {@link computeTenantStats}: Tenant-Aggregat (`total`, `teams`, `m365`, `sec`).
 *  - {@link applyFiltersPure}: filtert eine Reihen-Liste anhand expliziter
 *    Filter-State (rein, ohne DOM-Zugriff). Wird von {@link applyFilters}
 *    aufgerufen, die zusätzlich den UI-Filter-State aus dem DOM liest.
 *
 * DOM-Lese-Helfer:
 *  - {@link getFilterState}: liest die aktuellen Filter-Werte aus den
 *    Form-Feldern (`#ssFilterSchuljahr` …).
 *  - {@link applyFilters}: Convenience-Wrapper – liest Filter-State aus dem
 *    DOM und wendet ihn auf die Zeilen an.
 */

import { getEl } from '../../shared/utils/dom.js';
import { normStr } from '../../shared/utils/strings.js';

/**
 * Formatiert einen ISO-Zeitstempel in deutschsprachiges AT-Format
 * (`tt.mm.jjjj hh:mm`). Falls der Eingabewert kein gültiges Datum ist, wird
 * der getrimmte Original-String zurückgegeben.
 *
 * @param {string | null | undefined} iso
 * @returns {string}
 */
export function formatDateTimeAT(iso) {
    const s = String(iso || '').trim();
    if (!s) return '';
    try {
        const d = new Date(s);
        if (isNaN(d.getTime())) return s;
        const out = new Intl.DateTimeFormat('de-AT', {
            day: '2-digit',
            month: '2-digit',
            year: 'numeric',
            hour: '2-digit',
            minute: '2-digit',
            hour12: false
        }).format(d);
        return String(out).replace(',', '').replace(/\s+/g, ' ').trim();
    } catch {
        return s;
    }
}

/**
 * Mappt einen Sync-Status auf die zugehörige CSS-„Pill"-Klasse.
 *
 * @param {string | null | undefined} syncStatus
 * @returns {'ok' | 'warn' | 'err' | ''}
 */
export function pillClass(syncStatus) {
    const s = String(syncStatus || '');
    if (s === 'Ok') return 'ok';
    if (s === 'Abweichung') return 'warn';
    if (s === 'Fehler') return 'err';
    return '';
}

/**
 * Aggregiert SOLL-Struktur-Zeilen für die Statistik-Kacheln.
 *
 * @param {any[]} rows
 * @returns {{ total: number, aktiv: number, abw: number, err: number }}
 */
export function computeStats(rows) {
    const list = Array.isArray(rows) ? rows : [];
    const total = list.length;
    const aktiv = list.filter((r) => r && r.status === 'Aktiv').length;
    const abw = list.filter((r) => r && r.syncStatus === 'Abweichung').length;
    const err = list.filter((r) => r && r.syncStatus === 'Fehler').length;
    return { total, aktiv, abw, err };
}

/**
 * Aggregiert Tenant-Inventar-Zeilen für die Statistik-Kacheln.
 *
 * @param {any[]} rows
 * @returns {{ total: number, teams: number, m365: number, sec: number }}
 */
export function computeTenantStats(rows) {
    const list = Array.isArray(rows) ? rows : [];
    const total = list.length;
    const teams = list.filter((r) => r && r.typ === 'Team').length;
    const m365 = list.filter((r) => r && r.typ === 'Gruppe').length;
    const sec = list.filter(
        (r) => r && (r.typ === 'Sicherheitsgruppe' || r.typ === 'E‑Mail‑Sicherheitsgruppe')
    ).length;
    return { total, teams, m365, sec };
}

/**
 * @typedef {object} FilterState
 * @property {string} schuljahr      `"2025/26"` oder Leerstring für „alle".
 * @property {string} typ            Gefilterter Typ (z. B. `"Klasse"`).
 * @property {string} text           Volltext (bereits getrimmt + lowercase).
 * @property {string} visibility     Tenant-Filter: Sichtbarkeit (`Public`/…).
 * @property {string} roster         Tenant-Filter: `noOwners` / `noMembers` / `noOwnersNoMembers`.
 */

/**
 * Liest den Filter-State aus den DOM-Form-Feldern.
 * Achtung: liest synchron – nur im Browser-Kontext aufrufen.
 *
 * @returns {FilterState}
 */
export function getFilterState() {
    const schuljahr = normStr(getEl('ssFilterSchuljahr')?.value);
    const typ = normStr(getEl('ssFilterTyp')?.value);
    const text = normStr(getEl('ssFilterText')?.value).toLowerCase();
    const visibility = normStr(getEl('ssTenantVisibilityFilter')?.value);
    const roster = normStr(getEl('ssTenantRosterFilter')?.value);
    return { schuljahr, typ, text, visibility, roster };
}

/**
 * Wendet einen expliziten Filter-State auf eine Zeilenliste an – **rein**, ohne
 * DOM-Zugriff. Wird vor allem für Tests verwendet.
 *
 * @param {any[]} rows
 * @param {'soll' | 'tenant' | 'match'} mode
 * @param {FilterState} filterState
 * @returns {any[]} Gefilterte Zeilen.
 */
export function applyFiltersPure(rows, mode, filterState) {
    const list = Array.isArray(rows) ? rows : [];
    const f = filterState || { schuljahr: '', typ: '', text: '', visibility: '', roster: '' };
    return list.filter((r) => {
        if (!r) return false;
        if (mode !== 'tenant') {
            if (f.schuljahr && String(r.schuljahr || '') !== f.schuljahr) return false;
        }
        if (f.typ) {
            if (mode === 'tenant' && f.typ === 'Kursteam') {
                // Kursteams = HiddenMembership (praktischer Indikator für Kurs-Teams)
                if (!r.hiddenMembership) return false;
            } else if (String(r.typ || '') !== f.typ) {
                return false;
            }
        }
        if (mode === 'tenant') {
            if (f.visibility) {
                const v = String(r.visibility || '');
                if (v !== f.visibility) return false;
            }
            if (f.roster) {
                const oc = typeof r.ownerCount === 'number' ? r.ownerCount : -1;
                const mc = typeof r.memberCount === 'number' ? r.memberCount : -1;
                if (f.roster === 'noOwners') {
                    if (oc !== 0) return false;
                } else if (f.roster === 'noMembers') {
                    if (mc !== 0) return false;
                } else if (f.roster === 'noOwnersNoMembers') {
                    if (oc !== 0 || mc !== 0) return false;
                }
            }
        }
        if (f.text) {
            const hay = (
                String(r.bezeichnung || '') +
                ' ' +
                String(r.typ || '') +
                ' ' +
                String(r.schuljahr || '')
            ).toLowerCase();
            if (hay.indexOf(f.text) === -1) return false;
        }
        return true;
    });
}

/**
 * Convenience-Wrapper: liest den Filter-State aus dem DOM und wendet ihn an.
 *
 * @param {any[]} rows
 * @param {'soll' | 'tenant' | 'match'} mode
 * @returns {any[]}
 */
export function applyFilters(rows, mode) {
    return applyFiltersPure(rows, mode, getFilterState());
}
