/**
 * Zentrale String-Helfer. Konsolidiert ehemals 40+ lokale Kopien (`normStr`,
 * `normCode`, `escapeHtml`, …) aus den Tool-Files. Siehe ARCHITECTURE.md.
 *
 * Alle Funktionen sind **rein** (keine Seiteneffekte, keine DOM-/Window-Zugriffe).
 */

/** `String(v ?? '').trim()` – tolerant gegen `null` / `undefined`. */
export function normStr(v) {
    return String(v ?? '').trim();
}

/** Trim + UPPERCASE (z. B. Klassencode). */
export function normCode(v) {
    return normStr(v).toUpperCase();
}

/** Trim + lowercase (E-Mail-Vergleich). */
export function normEmail(v) {
    return normStr(v).toLowerCase();
}

/**
 * Header-Key-Normalisierung für CSV/Excel-Spalten:
 * trim, lower, Whitespace raus, Umlaute auflösen, Sonderzeichen entfernen.
 */
export function normHeaderKey(k) {
    return String(k ?? '')
        .trim()
        .toLowerCase()
        .replace(/\s+/g, '')
        .replace(/ä/g, 'ae')
        .replace(/ö/g, 'oe')
        .replace(/ü/g, 'ue')
        .replace(/ß/g, 'ss')
        .replace(/[^a-z0-9]/g, '');
}

/**
 * HTML-Escape für Text-Kontexte. Schützt zusätzlich `'` (strengste der
 * historischen Varianten – abwärtskompatibel zu allen bisherigen
 * lokalen `escapeHtml`-Implementierungen).
 */
export function escapeHtml(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');
}

/** HTML-Escape für Attribut-Kontexte (nur `&` und `"`). */
export function attrEscape(s) {
    return String(s ?? '').replace(/&/g, '&amp;').replace(/"/g, '&quot;');
}

/** Locale-aware deutscher Vergleich (für `Array.prototype.sort`). */
export function compareDe(a, b) {
    return String(a ?? '').localeCompare(String(b ?? ''), 'de', { sensitivity: 'base' });
}
