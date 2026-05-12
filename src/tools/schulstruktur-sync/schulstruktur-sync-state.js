/**
 * Persistenz-Layer für „Schulstruktur-Sync".
 *
 * Aus `schulstruktur-sync.js` 1:1 ausgelagert (Phase 2-Pilot). Verhalten
 * identisch: localStorage als Fallback, `window.ms365AppDataV2` als
 * bevorzugter Container, falls verfügbar.
 *
 * Enthält ausschließlich Storage-Helfer – keine UI, keine Graph-API.
 * Damit ohne DOM testbar.
 */

import { safeJsonParse } from '../../shared/utils/json.js';

/** @internal Storage-Keys (intern, nicht exportiert – Zugriff nur über die Helfer). */
const STORAGE_KEY = 'ms365-schulstruktur-sync-v1';
const STORAGE_TENANT_CACHE_KEY = 'ms365-schulstruktur-tenant-cache-v1';
const STORAGE_MATCH_KEY = 'ms365-schulstruktur-match-v1';
/** Geteilte ID-Menge mit der Graph-Ansicht (Tree-Collapse). */
const GRAPH_COLLAPSE_KEY = 'ms365-ss-graph-collapsed-v1';

/**
 * Liest den Strukturbaum, Mitgliedschaften und Settings. Bevorzugt
 * `ms365AppDataV2`, fällt sonst auf `localStorage` zurück.
 *
 * @returns {{ rows: any[], memberships: object, settings: object }}
 */
export function loadState() {
    try {
        if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getContainer === 'function') {
            const c = window.ms365AppDataV2.getContainer();
            if (c && c.structure && typeof c.structure === 'object') {
                const rows = Array.isArray(c.structure.rows) ? c.structure.rows : [];
                const memberships =
                    c.structure.memberships && typeof c.structure.memberships === 'object' ? c.structure.memberships : {};
                const settings =
                    c.structure.settings && typeof c.structure.settings === 'object' ? c.structure.settings : {};
                return { rows, memberships, settings };
            }
        }
        const raw = localStorage.getItem(STORAGE_KEY);
        const obj = raw ? safeJsonParse(raw) : null;
        const rows = obj && Array.isArray(obj.rows) ? obj.rows : [];
        const memberships =
            obj && obj.memberships && typeof obj.memberships === 'object' ? obj.memberships : {};
        const settings =
            obj && obj.settings && typeof obj.settings === 'object' ? obj.settings : {};
        return { rows, memberships, settings };
    } catch {
        return { rows: [], memberships: {}, settings: {} };
    }
}

/**
 * Persistiert den Strukturbaum. Schreibt parallel in `localStorage` UND
 * `ms365AppDataV2`, damit beide Quellen synchron bleiben.
 *
 * @param {{ rows?: any[], memberships?: object, settings?: object }} state
 * @returns {{ rows: any[], memberships: object }}
 */
export function saveState(state) {
    const rows = state && Array.isArray(state.rows) ? state.rows : [];
    const memberships = state && state.memberships && typeof state.memberships === 'object' ? state.memberships : {};
    const settings = state && state.settings && typeof state.settings === 'object' ? state.settings : {};
    try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify({ rows, memberships, settings }));
    } catch {
        // ignore
    }
    try {
        if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getContainer === 'function' && typeof window.ms365AppDataV2.setContainer === 'function') {
            const c = window.ms365AppDataV2.getContainer();
            c.structure = { rows, memberships, settings };
            window.ms365AppDataV2.setContainer(c);
        }
    } catch {
        // ignore
    }
    return { rows, memberships };
}

/**
 * Liest die Match-Links (Struktur-ID → Tenant-Group-ID / -User-ID + Notiz).
 *
 * @returns {{ links: object }}
 */
export function loadMatchState() {
    try {
        if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getContainer === 'function') {
            const c = window.ms365AppDataV2.getContainer();
            if (c && c.match && c.match.links && typeof c.match.links === 'object') {
                return { links: c.match.links };
            }
        }
        const raw = localStorage.getItem(STORAGE_MATCH_KEY);
        const obj = raw ? safeJsonParse(raw) : null;
        const links = obj && obj.links && typeof obj.links === 'object' ? obj.links : {};
        return { links };
    } catch {
        return { links: {} };
    }
}

/**
 * Persistiert die Match-Links.
 * @param {object} links
 * @returns {object} die persistierten Links (immer ein Objekt)
 */
export function saveMatchState(links) {
    const out = links && typeof links === 'object' ? links : {};
    try {
        localStorage.setItem(STORAGE_MATCH_KEY, JSON.stringify({ links: out }));
    } catch {
        // ignore
    }
    try {
        if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getContainer === 'function' && typeof window.ms365AppDataV2.setContainer === 'function') {
            const c = window.ms365AppDataV2.getContainer();
            c.match = { links: out };
            window.ms365AppDataV2.setContainer(c);
        }
    } catch {
        // ignore
    }
    return out;
}

/**
 * Liest den zwischengespeicherten Tenant-Inventar-Snapshot (Gruppen + User
 * + Ladestempel).
 *
 * @returns {{ rows: any[], users: any[], loadedAt: string }}
 */
export function loadTenantCache() {
    try {
        if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getContainer === 'function') {
            const c = window.ms365AppDataV2.getContainer();
            const cache = c && c.tenant && c.tenant.cache && typeof c.tenant.cache === 'object' ? c.tenant.cache : null;
            if (cache) {
                const rows = Array.isArray(cache.rows) ? cache.rows : [];
                const users = Array.isArray(cache.users) ? cache.users : [];
                return { rows, users, loadedAt: cache.loadedAt ? String(cache.loadedAt) : '' };
            }
        }
        const raw = localStorage.getItem(STORAGE_TENANT_CACHE_KEY);
        const obj = raw ? safeJsonParse(raw) : null;
        const rows = obj && Array.isArray(obj.rows) ? obj.rows : [];
        const users = obj && Array.isArray(obj.users) ? obj.users : [];
        return { rows, users, loadedAt: obj && obj.loadedAt ? String(obj.loadedAt) : '' };
    } catch {
        return { rows: [], users: [], loadedAt: '' };
    }
}

/**
 * Schreibt den Tenant-Inventar-Snapshot. Wenn `users` nicht angegeben ist,
 * wird der bestehende User-Cache beibehalten.
 *
 * @param {any[]} rows
 * @param {any[]} [users]
 */
export function saveTenantCache(rows, users) {
    const out = Array.isArray(rows) ? rows : [];
    const prev = loadTenantCache();
    const u = users !== undefined ? (Array.isArray(users) ? users : []) : prev.users || [];
    try {
        localStorage.setItem(
            STORAGE_TENANT_CACHE_KEY,
            JSON.stringify({ rows: out, users: u, loadedAt: new Date().toISOString() })
        );
    } catch {
        // ignore
    }
    try {
        if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getContainer === 'function' && typeof window.ms365AppDataV2.setContainer === 'function') {
            const c = window.ms365AppDataV2.getContainer();
            c.tenant = { cache: { rows: out, users: u, loadedAt: new Date().toISOString() } };
            window.ms365AppDataV2.setContainer(c);
        }
    } catch {
        // ignore
    }
}

/**
 * Liest die Liste der eingeklappten Knoten-IDs aus der Graph-Ansicht
 * als `Set<string>`. Schluckt JSON-Fehler und liefert dann ein leeres Set.
 *
 * @returns {Set<string>}
 */
export function loadGraphCollapsedSet() {
    try {
        const raw = localStorage.getItem(GRAPH_COLLAPSE_KEY);
        if (!raw) return new Set();
        const arr = JSON.parse(raw);
        if (!Array.isArray(arr)) return new Set();
        return new Set(arr.map((x) => String(x)));
    } catch {
        return new Set();
    }
}

/**
 * Persistiert die Collapsed-IDs.
 * @param {Set<string>} set
 */
export function saveGraphCollapsedSet(set) {
    try {
        localStorage.setItem(GRAPH_COLLAPSE_KEY, JSON.stringify(Array.from(set || []).map((x) => String(x))));
    } catch {
        // ignore
    }
}
