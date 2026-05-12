/**
 * Render-/UI-Helfer (Filter-Selects, Stats-Kacheln, Mode-Hint) für
 * „Schulstruktur-Sync".
 *
 * Aus `schulstruktur-sync.js` 1:1 ausgelagert (Phase 2 Schnitt 10).
 *
 * Reines DOM-Rendering – greift auf bekannte Form-/Stats-IDs zu
 * (`ssFilterSchuljahr`, `ssFilterTyp`, `ssStat…`, `ssModeHint`).
 *
 *  - {@link setModeHint}: Header-Hinweis je Modus (SOLL/Tenant/Match) inkl.
 *    optionalem „letztes Tenant-Einlesen"-Datum.
 *  - {@link renderFilters}: aktualisiert die Schuljahr-/Typ-Dropdowns je nach
 *    Modus und behält den Auswahlwert nach Möglichkeit.
 *  - {@link renderStats}: zeigt Aggregat-Zahlen + Spalten-Labels in den
 *    Kacheln an, abhängig vom Modus.
 *
 * Die zugrundeliegende Aggregations-Logik (`computeStats`,
 * `computeTenantStats`) lebt pure in `schulstruktur-sync-stats.js`.
 */

import { getEl } from '../../shared/utils/dom.js';
import { compareDe } from '../../shared/utils/strings.js';
import {
    computeStats,
    computeTenantStats
} from './schulstruktur-sync-stats.js';

/**
 * Setzt den Header-Hinweis abhängig vom Modus.
 * Im `'soll'`-Modus wird der Hinweis ausgeblendet.
 *
 * @param {'soll' | 'tenant' | 'match'} mode
 * @param {string | number | Date | null | undefined} [tenantLoadedAt]
 *        Optionaler Zeitstempel des letzten Tenant-Einlesens – wird formatiert
 *        angehängt.
 */
export function setModeHint(mode, tenantLoadedAt) {
    const el = getEl('ssModeHint');
    if (!el) return;
    if (mode === 'tenant') {
        el.style.display = '';
        el.textContent =
            'Tenant‑Inventar: Gruppen/Teams werden live per Graph eingelesen. Updates sind für Anzeigename/Beschreibung möglich.' +
            (tenantLoadedAt ? ' Letztes Einlesen: ' + new Date(tenantLoadedAt).toLocaleString() : '');
    } else if (mode === 'match') {
        el.style.display = '';
        el.textContent =
            'Abgleich: SOLL‑Einheiten werden mit bestehenden Tenant‑Gruppen/Teams verknüpft (Mapping lokal gespeichert). Über die Registerkarte „Organigramm" siehst du die SOLL‑Struktur vernetzt; im Baum und Organigramm kannst du per Drag&Drop umsortieren.' +
            (tenantLoadedAt ? ' Tenant zuletzt eingelesen: ' + new Date(tenantLoadedAt).toLocaleString() : '');
    } else {
        el.textContent = '';
        el.style.display = 'none';
    }
}

/**
 * Befüllt die Filter-Dropdowns (Schuljahr + Typ) abhängig vom Modus.
 *  - Schuljahr-Optionen werden aus den Zeilen abgeleitet und alphabetisch
 *    sortiert; im Tenant-Modus ist das Dropdown disabled.
 *  - Typ-Optionen unterscheiden sich SOLL vs. Tenant.
 * Beibehält den vorherigen Auswahlwert, sofern noch zulässig.
 *
 * @param {any[]} rows
 * @param {'soll' | 'tenant' | 'match'} mode
 */
export function renderFilters(rows, mode) {
    const sel = getEl('ssFilterSchuljahr');
    if (sel) {
        const prev = sel.value || '';
        const years = Array.from(
            new Set((rows || []).map((r) => String((r && r.schuljahr) || '')).filter(Boolean))
        ).sort(compareDe);
        sel.replaceChildren();
        const optAll = document.createElement('option');
        optAll.value = '';
        optAll.textContent = '(alle)';
        sel.appendChild(optAll);
        years.forEach((y) => {
            const o = document.createElement('option');
            o.value = y;
            o.textContent = y;
            sel.appendChild(o);
        });
        if (prev && years.indexOf(prev) !== -1) sel.value = prev;
        // Im Tenant-Modus ist Schuljahr-Filtern sinnlos -> disabled
        sel.disabled = mode === 'tenant';
    }

    const typeSel = getEl('ssFilterTyp');
    if (typeSel) {
        const prevType = typeSel.value || '';
        const opts =
            mode === 'tenant'
                ? [
                      { v: '', t: '(alle)' },
                      { v: 'Kursteam', t: 'Kursteams (HiddenMembership)' },
                      { v: 'Team', t: 'Team' },
                      { v: 'Gruppe', t: 'M365‑Gruppe' },
                      { v: 'Sicherheitsgruppe', t: 'Sicherheitsgruppe' },
                      { v: 'E‑Mail‑Sicherheitsgruppe', t: 'E‑Mail‑Sicherheitsgruppe' }
                  ]
                : [
                      { v: '', t: '(alle)' },
                      { v: 'Jahrgang', t: 'Jahrgang' },
                      { v: 'Klasse', t: 'Klasse' },
                      { v: 'Arbeitsgemeinschaft', t: 'Arbeitsgemeinschaft' },
                      { v: 'Kursteam', t: 'Kursteam' },
                      { v: 'Gruppe', t: 'Gruppe' },
                      { v: 'Person', t: 'Person' }
                  ];
        typeSel.replaceChildren();
        for (const o of opts) {
            const elO = document.createElement('option');
            elO.value = o.v;
            elO.textContent = o.t;
            typeSel.appendChild(elO);
        }
        if (prevType && opts.some((o) => o.v === prevType)) typeSel.value = prevType;
    }
}

/**
 * Zeigt die Stats-Kacheln (Einheiten/Aktiv/Abweichung/Fehler bzw. Teams/M365/Sec)
 * an, je nach Modus. Aktualisiert auch die Spalten-Labels (`.l`).
 *
 * @param {any[]} rows
 * @param {'soll' | 'tenant' | 'match'} mode
 */
export function renderStats(rows, mode) {
    const s = computeStats(rows);
    const elTotal = getEl('ssStatEinheiten');
    const elAktiv = getEl('ssStatAktiv');
    const elAbw = getEl('ssStatAbweichung');
    const elErr = getEl('ssStatFehler');
    if (mode === 'tenant') {
        const ts = computeTenantStats(rows);
        if (elTotal) elTotal.textContent = String(ts.total);
        if (elAktiv) elAktiv.textContent = String(ts.teams);
        if (elAbw) elAbw.textContent = String(ts.m365);
        if (elErr) elErr.textContent = String(ts.sec);
        const labAktiv = elAktiv && elAktiv.parentElement ? elAktiv.parentElement.querySelector('.l') : null;
        const labAbw = elAbw && elAbw.parentElement ? elAbw.parentElement.querySelector('.l') : null;
        const labErr = elErr && elErr.parentElement ? elErr.parentElement.querySelector('.l') : null;
        if (labAktiv) labAktiv.textContent = 'Teams';
        if (labAbw) labAbw.textContent = 'M365‑Gruppen';
        if (labErr) labErr.textContent = 'Sicherheitsgruppen';
        return;
    }
    if (elTotal) elTotal.textContent = String(s.total);
    if (elAktiv) elAktiv.textContent = String(s.aktiv);
    if (elAbw) elAbw.textContent = String(s.abw);
    if (elErr) elErr.textContent = String(s.err);
    const labAktiv = elAktiv && elAktiv.parentElement ? elAktiv.parentElement.querySelector('.l') : null;
    const labAbw = elAbw && elAbw.parentElement ? elAbw.parentElement.querySelector('.l') : null;
    const labErr = elErr && elErr.parentElement ? elErr.parentElement.querySelector('.l') : null;
    if (labAktiv) labAktiv.textContent = 'aktiv';
    if (labAbw) labAbw.textContent = 'Abweichung';
    if (labErr) labErr.textContent = 'Fehler';
}
