/**
 * Toleranter `JSON.parse`-Wrapper. Konsolidiert die `safeJsonParse`-Kopien
 * aus diversen Tool-Files (siehe ARCHITECTURE.md).
 *
 * @template T
 * @param {unknown} s     Eingabe (String oder beliebig).
 * @param {T} [fallback]  Rückgabe bei Parse-Fehler. Default: `null`.
 * @returns {any|T}
 */
export function safeJsonParse(s, fallback = null) {
    try {
        return JSON.parse(String(s));
    } catch {
        return fallback;
    }
}
