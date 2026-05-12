/**
 * Tenant-Match-Vorschläge und das zugehörige Suchfeld-Wiring für
 * „Schulstruktur-Sync".
 *
 * Aus `schulstruktur-sync.js` 1:1 ausgelagert (Phase 2 Schnitt 7).
 *
 * Pure Funktionen (gut testbar):
 *  - {@link normKey}: aggressive, Unicode-bewusste Normalisierung für Match-Keys.
 *  - {@link suggestTenantGroupForUnitFromList}: schlägt eine Gruppe für eine
 *    Strukturzeile vor (exakt → Alias → enthält).
 *  - {@link suggestTenantUserForPersonFromList}: scoring-basierter
 *    Vorschlag für Person-Zeilen anhand `displayName`/`UPN`/`mail`.
 *  - {@link suggestTenantMatchSelectValue}: kombiniert beide Vorschläge zum
 *    `g:<id>` / `u:<id>` Dropdown-Wert.
 *  - {@link formatEntraUserPickLabel}: einheitliches Anzeige-Label.
 *  - {@link matchTenantFilterNeedle}: Filtertext-Normalisierung.
 *  - {@link matchTenantHaystackForGroup} / {@link matchTenantHaystackForUser}:
 *    erzeugen den Vergleichs-Heuhaufen.
 *
 * DOM-Wiring (nicht testbar):
 *  - {@link rebuildMatchTenantSelectOptions}: baut die Dropdown-Liste neu auf.
 *  - {@link wireMatchTenantSearchOnce}: Event-Wiring fürs Suchfeld (idempotent).
 */

import { getEl } from '../../shared/utils/dom.js';

/**
 * Aggressive Normalisierung für Match-Keys: trim, lower, Whitespace zusammenfassen,
 * alle Zeichen außer Buchstaben/Zahlen und `- _ .` entfernen.
 *
 * @param {unknown} s
 * @returns {string}
 */
export function normKey(s) {
    return String(s || '')
        .trim()
        .toLowerCase()
        .replace(/\s+/g, ' ')
        .replace(/[^\p{L}\p{N}\-_. ]/gu, '');
}

/**
 * Abgleich-Vorschlag: passende Tenant-Gruppe für eine Strukturzeile.
 * Reihenfolge:
 *  1. exakter Match auf `bezeichnung`
 *  2. exakter Match auf `alias`
 *  3. „enthält" auf `bezeichnung` oder `alias`
 *
 * @param {{ bezeichnung?: string } | null} unit
 * @param {Array<{ id: string|number, bezeichnung?: string, alias?: string }>} list
 * @returns {string} Group-ID oder Leerstring.
 */
export function suggestTenantGroupForUnitFromList(unit, list) {
    if (!unit) return '';
    const uKey = normKey(unit.bezeichnung || '');
    if (!uKey) return '';
    const rows = Array.isArray(list) ? list : [];
    let best = rows.find((g) => normKey(g.bezeichnung) === uKey);
    if (best) return String(best.id);
    best = rows.find((g) => g.alias && normKey(g.alias) === uKey);
    if (best) return String(best.id);
    best = rows.find(
        (g) => normKey(g.bezeichnung).includes(uKey) || (g.alias && normKey(g.alias).includes(uKey))
    );
    return best ? String(best.id) : '';
}

/**
 * Abgleich-Vorschlag: Entra-Benutzer für SOLL-Typ „Person" (Name/E-Mail/Rolle).
 * Punktet:
 *  - 100 bei exaktem Match auf `displayName`/`UPN`/`mail`
 *  - 50 bei „enthält"-Match `displayName`
 *  - 45 bei „enthält"-Match `UPN` oder `mail`
 *
 * @param {{ typ?: string, personName?: string, personEmail?: string, bezeichnung?: string } | null} unit
 * @param {Array<{ id?: string|number, displayName?: string, userPrincipalName?: string, mail?: string }>} users
 * @returns {string} User-ID oder Leerstring.
 */
export function suggestTenantUserForPersonFromList(unit, users) {
    if (!unit || String(unit.typ || '') !== 'Person') return '';
    const arr = Array.isArray(users) ? users : [];
    const keys = [];
    const pushK = (x) => {
        const k = normKey(x);
        if (k) keys.push(k);
    };
    pushK(unit.personName);
    pushK(unit.personEmail);
    pushK(unit.bezeichnung);
    if (!keys.length) return '';

    function scoreUser(u) {
        const dn = normKey(u.displayName);
        const upn = normKey(u.userPrincipalName);
        const mail = normKey(u.mail);
        let best = 0;
        for (let i = 0; i < keys.length; i++) {
            const k = keys[i];
            if (!k) continue;
            if (k === dn || k === upn || k === mail) return 100;
            if (dn && (dn.includes(k) || k.includes(dn))) best = Math.max(best, 50);
            if (upn && (upn.includes(k) || k.includes(upn))) best = Math.max(best, 45);
            if (mail && (mail.includes(k) || k.includes(mail))) best = Math.max(best, 45);
        }
        return best;
    }

    let bestId = '';
    let bestScore = 0;
    for (let j = 0; j < arr.length; j++) {
        const u = arr[j];
        const sc = scoreUser(u);
        if (sc > bestScore) {
            bestScore = sc;
            bestId = String(u.id || '');
        }
    }
    return bestId;
}

/**
 * Dropdown-Wert: `g:<id>` für Gruppe/Team, `u:<id>` für Entra-Benutzer.
 * Person-Zeilen werden zuerst gegen User gematcht, sonst Gruppen.
 *
 * @param {object | null} unit
 * @param {any[]} groups
 * @param {any[]} users
 * @returns {string}
 */
export function suggestTenantMatchSelectValue(unit, groups, users) {
    if (!unit) return '';
    if (String(unit.typ || '') === 'Person') {
        const uid = suggestTenantUserForPersonFromList(unit, users);
        if (uid) return 'u:' + uid;
    }
    const gid = suggestTenantGroupForUnitFromList(unit, groups);
    return gid ? 'g:' + gid : '';
}

/**
 * Einheitliches Anzeige-Label für Entra-User im Auswahl-Dropdown.
 *
 * @param {{ id?: string|number, displayName?: string, userPrincipalName?: string, mail?: string } | null} u
 * @returns {string}
 */
export function formatEntraUserPickLabel(u) {
    if (!u || typeof u !== 'object') return '';
    const dn = u.displayName ? String(u.displayName).trim() : '';
    const upn = String(u.userPrincipalName || u.mail || '').trim();
    if (dn && upn && dn.toLowerCase() !== upn.toLowerCase()) return dn + ' · ' + upn + ' · Benutzer';
    return (dn || upn || String(u.id || '')) + ' · Benutzer';
}

/**
 * Normalisiert einen Filtertext für „contains"-Match (lower, getrimmt,
 * Whitespace zusammengefasst).
 *
 * @param {unknown} raw
 * @returns {string}
 */
export function matchTenantFilterNeedle(raw) {
    return String(raw || '')
        .trim()
        .toLowerCase()
        .replace(/\s+/g, ' ');
}

/**
 * Erzeugt den Such-„Heuhaufen" für eine Gruppe (lowercase, mit Trennzeichen).
 *
 * @param {{ bezeichnung?: string, typ?: string, alias?: string, mail?: string, description?: string, id?: string|number } | null} g
 * @returns {string}
 */
export function matchTenantHaystackForGroup(g) {
    return [
        g && g.bezeichnung,
        g && g.typ,
        g && g.alias,
        g && g.mail,
        g && g.description,
        g && g.id
    ]
        .map((x) => String(x || '').toLowerCase())
        .join(' ');
}

/**
 * Erzeugt den Such-„Heuhaufen" für einen Entra-User – inkl. des
 * Anzeige-Labels, damit auch der Label-Text getroffen wird.
 *
 * @param {object | null} u
 * @returns {string}
 */
export function matchTenantHaystackForUser(u) {
    return [
        u && u.displayName,
        u && u.userPrincipalName,
        u && u.mail,
        u && u.id,
        formatEntraUserPickLabel(u)
    ]
        .map((x) => String(x || '').toLowerCase())
        .join(' ');
}

/**
 * Baut die Tenant-Auswahlliste neu auf (optional gefiltert).
 * Quelldaten unter `window.__ms365MatchTenantPickSource = { groups, users }`.
 *
 * @param {HTMLSelectElement | null} selTenant
 * @param {string} filterRaw
 * @param {string} selectedValue Wert nach Auswahl (z. B. `g:…` / `u:…`).
 */
export function rebuildMatchTenantSelectOptions(selTenant, filterRaw, selectedValue) {
    const src = (typeof window !== 'undefined' && window.__ms365MatchTenantPickSource) || null;
    const cntEl = getEl('ssMatchTenantFilterCount');
    if (!selTenant || !src || typeof src !== 'object') {
        if (cntEl) cntEl.textContent = '';
        return;
    }
    const needle = matchTenantFilterNeedle(filterRaw);
    const list = Array.isArray(src.groups) ? src.groups : [];
    const users = Array.isArray(src.users) ? src.users : [];
    const prevSel = String(selectedValue != null ? selectedValue : selTenant.value || '');

    selTenant.replaceChildren();
    const opt0 = document.createElement('option');
    opt0.value = '';
    opt0.textContent = '(keine Zuordnung)';
    selTenant.appendChild(opt0);

    const gFiltered = needle
        ? list.filter((g) => matchTenantHaystackForGroup(g).indexOf(needle) !== -1)
        : list.slice();
    const uFiltered = needle
        ? users.filter((u) => matchTenantHaystackForUser(u).indexOf(needle) !== -1)
        : users.slice();

    function selectionInFiltered() {
        if (!prevSel) return true;
        if (prevSel.startsWith('g:')) {
            const id = prevSel.slice(2);
            return gFiltered.some((g) => String(g.id) === id);
        }
        if (prevSel.startsWith('u:')) {
            const id = prevSel.slice(2);
            return uFiltered.some((u) => String(u.id) === id);
        }
        return false;
    }

    if (prevSel && needle && !selectionInFiltered()) {
        let label = prevSel;
        if (prevSel.startsWith('g:')) {
            const id = prevSel.slice(2);
            const g = list.find((x) => String(x.id) === id);
            if (g) label = (g.bezeichnung || id) + ' · ' + (g.typ || '') + ' · aktuell verknüpft';
        } else if (prevSel.startsWith('u:')) {
            const id = prevSel.slice(2);
            const u = users.find((x) => String(x.id) === id);
            if (u) label = formatEntraUserPickLabel(u) + ' · aktuell verknüpft';
        }
        const ox = document.createElement('option');
        ox.value = prevSel;
        ox.textContent = label;
        selTenant.appendChild(ox);
    }

    if (gFiltered.length) {
        const og = document.createElement('optgroup');
        og.label = 'Gruppen / Teams';
        for (let i = 0; i < gFiltered.length; i++) {
            const g = gFiltered[i];
            const o = document.createElement('option');
            o.value = 'g:' + String(g.id || '');
            o.textContent = (g.bezeichnung || '(ohne Name)') + ' · ' + (g.typ || '') + (g.alias ? ' · ' + g.alias : '');
            og.appendChild(o);
        }
        selTenant.appendChild(og);
    }

    if (uFiltered.length) {
        const ou = document.createElement('optgroup');
        ou.label = 'Benutzer (Entra ID)';
        for (let j = 0; j < uFiltered.length; j++) {
            const u = uFiltered[j];
            const o = document.createElement('option');
            o.value = 'u:' + String(u.id || '');
            o.textContent = formatEntraUserPickLabel(u);
            ou.appendChild(o);
        }
        selTenant.appendChild(ou);
    }

    if (prevSel) {
        const ok = Array.from(selTenant.options).some((o) => String(o.value) === prevSel);
        selTenant.value = ok ? prevSel : '';
    } else {
        selTenant.value = '';
    }

    if (cntEl) {
        const total = list.length + users.length;
        const shown = gFiltered.length + uFiltered.length;
        if (!total) cntEl.textContent = 'Noch keine Tenant-Daten – unter „Verwalten" Tenant einlesen.';
        else if (!needle) cntEl.textContent = String(total) + ' Einträge – Suchfeld nutzen, um die Liste einzugrenzen.';
        else cntEl.textContent = 'Zeige ' + shown + ' von ' + total + ' Einträgen (Filter aktiv).';
    }
}

/** Modul-privater Singleton-Guard für `wireMatchTenantSearchOnce`. */
let __matchTenantSearchWired = false;

/**
 * Verkabelt das Tenant-Suchfeld genau einmal mit `rebuildMatchTenantSelectOptions`.
 * Idempotent – wiederholte Aufrufe sind ein No-Op.
 */
export function wireMatchTenantSearchOnce() {
    if (__matchTenantSearchWired) return;
    const inp = getEl('ssMatchTenantSearch');
    if (!inp) return;
    __matchTenantSearchWired = true;
    inp.addEventListener('input', () => {
        const sel = getEl('ssMatchTenantGroup');
        if (!sel) return;
        const cur = String(sel.value || '');
        rebuildMatchTenantSelectOptions(sel, inp.value || '', cur);
    });
}
