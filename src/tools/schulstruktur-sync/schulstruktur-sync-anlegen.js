/**
 * Tenant-Anlegen-Helfer: Default-Werte für Target/Visibility und der zentrale
 * Anzeige-Namen-Vorschlag.
 *
 * Aus `schulstruktur-sync.js` 1:1 ausgelagert (Phase 2 Schnitt 11).
 *
 *  - {@link defaultTenantTargetForTypeStr}: Default `team`/`group` je Typ.
 *  - {@link defaultTenantVisibilityForTypeStr}: Default
 *    `HiddenMembership`/`Private` je Typ.
 *  - {@link resolveKursteamKlasseFachForRow}: Fallback-Auflösung
 *    Klasse/Fach für ein Kursteam, wenn `window.ms365StructureRules` keine
 *    eigene Implementierung liefert.
 *  - {@link computeTenantCreateSuggestionPure}: pure Variante des
 *    Anzeige-Namen-Vorschlags, mit explizitem `resolveKlasseFach`-Callback.
 *  - {@link normRoleKey}: Normalisierung für Rollen-Labels (für
 *    `personInfoByRole`-Lookups).
 */

import { normStr } from '../../shared/utils/strings.js';
import {
    buildJgMailNick,
    buildArgeMailNick,
    buildMailNickFromLabel,
    buildKursteamMailNickFromTemplate
} from './schulstruktur-sync-naming.js';
import { defaultAnlegenSchemas } from './schulstruktur-sync-demo.js';

/**
 * Default `Target` je Typ.
 *  - `Kursteam`, `Klasse` → `team` (klassisches Microsoft Team)
 *  - sonst → `group` (M365-Gruppe)
 *
 * @param {string} typ
 * @returns {'team' | 'group'}
 */
export function defaultTenantTargetForTypeStr(typ) {
    const t = String(typ || '');
    if (t === 'Kursteam' || t === 'Klasse') return 'team';
    return 'group';
}

/**
 * Default `Visibility` je Typ.
 *  - `Kursteam` → `HiddenMembership` (Schüler:innen sehen Mitglieder nicht)
 *  - sonst → `Private`
 *
 * @param {string} typ
 * @returns {'HiddenMembership' | 'Private'}
 */
export function defaultTenantVisibilityForTypeStr(typ) {
    const t = String(typ || '');
    if (t === 'Kursteam') return 'HiddenMembership';
    return 'Private';
}

/**
 * Lazy-Lookup auf `window.ms365StructureRules.resolveKursteamKlasseFach`.
 * @returns {((row: any, rows: any[]) => any) | null}
 */
function getRulesResolver() {
    const sr = (typeof window !== 'undefined' && window.ms365StructureRules) || null;
    return sr && typeof sr.resolveKursteamKlasseFach === 'function' ? sr.resolveKursteamKlasseFach : null;
}

/**
 * Löst Klasse + Fach für ein Kursteam auf.
 *  1. Bevorzugt `window.ms365StructureRules.resolveKursteamKlasseFach`
 *  2. Fallback: `row.ktKlasse` / `row.ktFach`; wenn `ktKlasse` leer ist und der
 *     Parent eine `Klasse` ist, wird dessen Bezeichnung verwendet.
 *
 * @param {any} row Strukturzeile (typ. `typ: 'Kursteam'`).
 * @param {any[]} rows Alle Strukturzeilen (für Parent-Auflösung).
 * @returns {{ klasse: string, fach: string, hasBoth: boolean }}
 */
export function resolveKursteamKlasseFachForRow(row, rows) {
    const resolver = getRulesResolver();
    if (resolver) return resolver(row, rows);
    const byId = Object.create(null);
    (rows || []).forEach((r) => {
        if (r && r.id != null) byId[String(r.id)] = r;
    });
    let klasse = normStr(row && row.ktKlasse);
    const fach = normStr(row && row.ktFach);
    const pid = normStr(row && row.parentId);
    if (!klasse && pid && byId[pid] && normStr(byId[pid].typ) === 'Klasse') {
        klasse = normStr(byId[pid].bezeichnung);
    }
    return { klasse, fach, hasBoth: !!(klasse && fach) };
}

/**
 * Liefert einen Vorschlag für `{ displayName, mailNick }` zur Anlage einer
 * Tenant-Gruppe/eines Teams.
 *
 * **Pure**: nimmt einen Callback `resolveKlasseFach(row)` und keine globalen
 * Zustände. Für die im Tool übliche Variante siehe
 * {@link computeTenantCreateSuggestionFromRow} im Hauptfile.
 *
 * @param {{ typ?: string, bezeichnung?: string, jgYear?: string, jgSuffix?: string, argeCode?: string, ktGruppe?: string }} row
 * @param {{ kursteamYearPrefix?: string, kursteamMailNickPattern?: string }} [schemaState]
 * @param {(row: any) => { klasse?: string, fach?: string }} [resolveKlasseFach]
 *        Nur für `typ === 'Kursteam'` relevant.
 * @returns {{ displayName: string, mailNick: string }}
 */
export function computeTenantCreateSuggestionPure(row, schemaState, resolveKlasseFach) {
    const schema =
        schemaState && typeof schemaState === 'object' ? schemaState : defaultAnlegenSchemas();
    const typ = String(row?.typ || '');
    const displayName = String(row?.bezeichnung || '').trim();
    if (!displayName) return { displayName: '', mailNick: '' };

    if (typ === 'Jahrgang') {
        const y = String(row.jgYear || '').trim();
        const suf = String(row.jgSuffix || '').trim();
        const mailNick = y && suf ? buildJgMailNick(schema, y, suf) : buildMailNickFromLabel(displayName);
        return { displayName, mailNick };
    }
    if (typ === 'Arbeitsgemeinschaft') {
        const code = String(row.argeCode || '').trim();
        const mailNick = code ? buildArgeMailNick(schema, code) : buildMailNickFromLabel(displayName);
        return { displayName, mailNick };
    }
    if (typ === 'Kursteam') {
        const yearPrefix = String(schema.kursteamYearPrefix || '').trim();
        const kt = typeof resolveKlasseFach === 'function' ? resolveKlasseFach(row) : null;
        const klasse = kt && kt.klasse ? String(kt.klasse) : '';
        const fach = kt && kt.fach ? String(kt.fach) : '';
        const gruppe = String(row.ktGruppe || '').trim();
        const mailNick =
            klasse && fach
                ? buildKursteamMailNickFromTemplate(schema.kursteamMailNickPattern, { yearPrefix, klasse, fach, gruppe })
                : buildMailNickFromLabel(displayName);
        return { displayName, mailNick };
    }
    return { displayName, mailNick: buildMailNickFromLabel(displayName) };
}

/**
 * Normalisiert ein Rollen-Label für einen Map-Lookup
 * (Trim, lower, Whitespace zusammenfassen, `:` und `;` entfernen).
 *
 * @param {unknown} s
 * @returns {string}
 */
export function normRoleKey(s) {
    return String(s || '')
        .trim()
        .toLowerCase()
        .replace(/\s+/g, ' ')
        .replace(/[:;]/g, '');
}
