/**
 * Verwaltungs-/Admin-Datenmodell für den Einrichtungs-Assistenten.
 * Aus setup-wizard.js ausgelagert (gleiches Verhalten).
 *
 * Die String-Helper (`normStr`, `normEmail`, `normCode`, `escapeHtml`) werden
 * aus `./utils/strings.js` re-exportiert, damit bestehende Imports
 * `import { normStr } from '.../setup-wizard-admin-model.js'` weiterhin
 * funktionieren. Neue Aufrufer sollten direkt aus `utils/strings.js` importieren.
 */

export { normStr, normEmail, normCode, escapeHtml } from './utils/strings.js';

import { normStr, normEmail } from './utils/strings.js';

export const SW_ADMIN_DEFAULT_ROLES = [
    'Direktion',
    'Sekretariat',
    'Administration',
    'Schularzt',
    'Schulwart',
    'IT-Support',
    'Bibliothek'
];

/** Kanonischer Standard-Slot nur bei gesetztem defaultKey (vermeidet Kollision mit freier Rolle „Direktion“). */
export function resolveAdminSlotFromRow(row) {
    if (!row) return null;
    const dk = normStr(row.defaultKey);
    if (!dk) return null;
    const m = SW_ADMIN_DEFAULT_ROLES.find(function (x) {
        return x.toLowerCase() === dk.toLowerCase();
    });
    return m || null;
}

export function migrateAdminRowDefaultKey(row) {
    const o = Object.assign({}, row);
    if (normStr(o.defaultKey)) return o;
    const rl = normStr(o.role);
    const m = SW_ADMIN_DEFAULT_ROLES.find(function (x) {
        return x.toLowerCase() === rl.toLowerCase();
    });
    if (m) o.defaultKey = m;
    return o;
}

export function isDirektionRole(roleRaw) {
    const r = normStr(roleRaw).toLowerCase();
    if (!r) return false;
    return r.indexOf('direktion') !== -1 || r.indexOf('direktor') !== -1;
}

export function isDirektionAdminRow(row) {
    const slot = resolveAdminSlotFromRow(row);
    if (slot && slot.toLowerCase() === 'direktion') return true;
    return isDirektionRole(row && row.role);
}

export function getAdminDisplayRowsFromSettings(settings) {
    const adminArr = settings && Array.isArray(settings.admin) ? settings.admin : [];
    if (adminArr.length) {
        return adminArr.map(function (r) {
            return migrateAdminRowDefaultKey(r);
        });
    }
    return SW_ADMIN_DEFAULT_ROLES.map(function (slot) {
        return { defaultKey: slot, role: slot, name: '', email: '' };
    });
}

export function collectDirektionOwnerEmails(settings) {
    const out = [];
    const seen = new Set();
    const admin = settings && Array.isArray(settings.admin) ? settings.admin : [];
    admin.forEach(function (row) {
        if (!isDirektionAdminRow(row)) return;
        const em = normEmail(row && row.email);
        if (!em || em.indexOf('@') === -1) return;
        if (seen.has(em)) return;
        seen.add(em);
        out.push(em);
    });
    return out;
}

export function collectAdminOwnerEmails(settings) {
    const out = [];
    const seen = new Set();
    const admin = settings && Array.isArray(settings.admin) ? settings.admin : [];
    admin.forEach(function (row) {
        const em = normEmail(row && row.email);
        if (!em || em.indexOf('@') === -1) return;
        if (seen.has(em)) return;
        seen.add(em);
        out.push(em);
    });
    return out;
}

export function collectEmails(arr) {
    const out = [];
    const seen = new Set();
    (Array.isArray(arr) ? arr : []).forEach(function (row) {
        const em = normEmail(row && row.email);
        if (!em || em.indexOf('@') === -1) return;
        if (seen.has(em)) return;
        seen.add(em);
        out.push(em);
    });
    return out;
}

export function randomTempPassword() {
    const u = 'ABCDEFGHJKLMNPQRSTUVWXYZ';
    const l = 'abcdefghijkmnopqrstuvwxyz';
    const d = '23456789';
    const s = '!@#$%&*';
    function pick(set) {
        return set.charAt(Math.floor(Math.random() * set.length));
    }
    let pwd = pick(u) + pick(l) + pick(d) + pick(s);
    for (let i = 0; i < 12; i++) {
        pwd += pick(u + l + d);
    }
    return pwd;
}
