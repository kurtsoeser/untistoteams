/**
 * Typsichere Hülle um `window.localStorage`. Fängt Quota-Fehler und
 * private-Browsing-Restriktionen ab, ohne den Aufrufer-Code zu verkomplizieren.
 *
 * Konsolidiert wiederkehrende `try { JSON.parse(localStorage.getItem(K)) }`
 * -Muster aus den Tool-Files.
 */

import { safeJsonParse } from './json.js';

function ls() {
    if (typeof window === 'undefined') return null;
    try {
        return window.localStorage || null;
    } catch {
        return null;
    }
}

/**
 * Liest und parst einen JSON-Wert. Bei fehlendem Eintrag oder Parse-Fehler
 * wird `fallback` zurückgegeben.
 *
 * @template T
 * @param {string} key
 * @param {T} [fallback]
 * @returns {any|T}
 */
export function loadJson(key, fallback = null) {
    const store = ls();
    if (!store) return fallback;
    try {
        const raw = store.getItem(key);
        if (raw == null) return fallback;
        const parsed = safeJsonParse(raw, fallback);
        return parsed === null && fallback !== null ? fallback : parsed;
    } catch {
        return fallback;
    }
}

/**
 * Serialisiert `value` als JSON und speichert. Gibt `true` bei Erfolg, sonst `false`.
 *
 * @param {string} key
 * @param {unknown} value
 * @returns {boolean}
 */
export function saveJson(key, value) {
    const store = ls();
    if (!store) return false;
    try {
        store.setItem(key, JSON.stringify(value));
        return true;
    } catch {
        return false;
    }
}

/**
 * Entfernt einen Eintrag. Idempotent.
 * @param {string} key
 * @returns {boolean}
 */
export function removeKey(key) {
    const store = ls();
    if (!store) return false;
    try {
        store.removeItem(key);
        return true;
    } catch {
        return false;
    }
}
