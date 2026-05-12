/**
 * Notifikations- und Progress-Anzeigen für „Schulstruktur-Sync".
 *
 * Aus `schulstruktur-sync.js` 1:1 ausgelagert (Phase 2 Schnitt 12).
 *
 *  - {@link toast}: zentraler Snackbar-Aufruf; nutzt
 *    `window.ms365ShowToast` falls vorhanden, sonst Fallback in den
 *    Progress-Text der Tenant-Sync-Anzeige.
 *  - {@link setTenantProgress}: aktualisiert die Progress-Bar
 *    (`ssTenantProgressWrap`, `…Text`, `…Bar`, `…Pct`) anhand eines
 *    optionalen Ratio-Werts (0–1).
 *
 * Reines DOM-Modul – greift nur auf bekannte IDs zu, hat keine eigenen
 * Tests (geringer Wert ggü. Mock-Aufwand).
 */

import { getEl } from '../../shared/utils/dom.js';

/**
 * Zeigt eine kurze Nachricht an.
 *
 * Bevorzugt den globalen Snackbar (`window.ms365ShowToast`), fällt
 * andernfalls auf den Progress-Text in der Tenant-Anzeige zurück.
 *
 * @param {unknown} msg Beliebiger Inhalt; wird über `String()` formatiert.
 */
export function toast(msg) {
    if (typeof window !== 'undefined' && typeof window.ms365ShowToast === 'function') {
        window.ms365ShowToast(msg);
    } else {
        // fallback: kein blocking alert für längere Vorgänge
        try {
            const el = getEl('ssTenantProgressText');
            if (el) el.textContent = String(msg);
        } catch {
            // ignore
        }
    }
}

/**
 * Steuert die Tenant-Sync-Progress-Bar im UI.
 *
 * @param {boolean} visible Wenn `true`, wird die Anzeige eingeblendet.
 * @param {string} [text] Optionaler Beschreibungstext.
 * @param {number | null} [ratio] Fortschritt 0–1; bei `null`/ungültig wird
 *        ein „–" angezeigt und der Balken auf 0 % gesetzt.
 */
export function setTenantProgress(visible, text, ratio) {
    const wrap = getEl('ssTenantProgressWrap');
    const txt = getEl('ssTenantProgressText');
    const bar = getEl('ssTenantProgressBar');
    const pct = getEl('ssTenantProgressPct');
    if (wrap) wrap.style.display = visible ? '' : 'none';
    if (txt && text) txt.textContent = String(text);
    const r = typeof ratio === 'number' && isFinite(ratio) ? Math.max(0, Math.min(1, ratio)) : null;
    if (bar) bar.style.width = r === null ? '0%' : String(Math.round(r * 100)) + '%';
    if (pct) pct.textContent = r === null ? '–' : String(Math.round(r * 100)) + '%';
}
