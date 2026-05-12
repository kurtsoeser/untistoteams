/**
 * Wrapper auf die zentrale App-Dialog-Implementierung in `app-dialog.js`.
 * Fallback auf die nativen `window.alert/confirm/prompt`, falls
 * `app-dialog.js` (noch) nicht geladen ist – z. B. in sehr frühen
 * Boot-Phasen oder in Tests.
 *
 * Konsolidiert ehemals dutzende lokale `dlgAlert/Confirm/Prompt`-Kopien.
 */

/**
 * @param {string} msg
 * @param {object} [opts] - z. B. `{ title, okText }`
 * @returns {Promise<void>}
 */
export function dlgAlert(msg, opts) {
    if (typeof window !== 'undefined' && typeof window.ms365AppDialogAlert === 'function') {
        return window.ms365AppDialogAlert(msg, opts);
    }
    if (typeof window !== 'undefined') window.alert(msg);
    return Promise.resolve();
}

/**
 * @param {string} msg
 * @param {object} [opts] - z. B. `{ title, okText, danger }`
 * @returns {Promise<boolean>}
 */
export function dlgConfirm(msg, opts) {
    if (typeof window !== 'undefined' && typeof window.ms365AppDialogConfirm === 'function') {
        return window.ms365AppDialogConfirm(msg, opts);
    }
    if (typeof window !== 'undefined') return Promise.resolve(window.confirm(msg));
    return Promise.resolve(false);
}

/**
 * @param {string} msg
 * @param {string} [def]
 * @param {object} [opts] - z. B. `{ title, okText, inputLabel }`
 * @returns {Promise<string|null>}
 */
export function dlgPrompt(msg, def, opts) {
    if (typeof window !== 'undefined' && typeof window.ms365AppDialogPrompt === 'function') {
        return window.ms365AppDialogPrompt(msg, def, opts);
    }
    if (typeof window !== 'undefined') return Promise.resolve(window.prompt(msg, def));
    return Promise.resolve(null);
}
