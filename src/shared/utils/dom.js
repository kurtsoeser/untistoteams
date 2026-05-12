/**
 * Minimal-DOM-Helpers für Tool-UIs.
 */

/** Kurzform für `document.getElementById`. */
export function getEl(id) {
    if (typeof document === 'undefined') return null;
    return document.getElementById(id);
}

/**
 * Toast-Helper. Erwartet ein Element mit `id="toast"`; fügt für die
 * konfigurierte Dauer die CSS-Klasse `show` hinzu.
 *
 * @param {string} msg
 * @param {object} [opts]
 * @param {number} [opts.durationMs=3500]
 * @param {string} [opts.elementId='toast']
 */
export function showToast(msg, opts = {}) {
    if (typeof document === 'undefined') return;
    const { durationMs = 3500, elementId = 'toast' } = opts;
    const el = document.getElementById(elementId);
    if (!el) {
        if (typeof window !== 'undefined' && typeof window.ms365ShowToast === 'function') {
            window.ms365ShowToast(msg);
        }
        return;
    }
    el.textContent = String(msg ?? '');
    el.classList.add('show');
    clearTimeout(showToast._t);
    showToast._t = setTimeout(() => el.classList.remove('show'), durationMs);
}
