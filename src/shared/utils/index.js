/**
 * Barrel-Re-Export der gemeinsam genutzten Helper-Module.
 * Erlaubt Aufrufern bei Bedarf einen Sammel-Import:
 *
 *     import { normStr, escapeHtml, safeJsonParse } from '../../shared/utils/index.js';
 *
 * Bevorzugt aber: direkt aus dem konkreten Submodul importieren – das hält
 * den Import-Graph klein.
 */

export { normStr, normCode, normEmail, normHeaderKey, escapeHtml, attrEscape, compareDe } from './strings.js';
export { safeJsonParse } from './json.js';
export { dlgAlert, dlgConfirm, dlgPrompt } from './dialog.js';
export { loadJson, saveJson, removeKey } from './storage.js';
export { getEl, showToast } from './dom.js';
