/**
 * Demo-Daten und „Anlegen"-Schema-Defaults für „Schulstruktur-Sync".
 *
 * Aus `schulstruktur-sync.js` 1:1 ausgelagert (Phase 2 Schnitt 6).
 *
 * Verantwortlichkeiten:
 *  - {@link defaultAnlegenSchemas}: Liefert die Default-Settings für das
 *    Anlegen-Schema (Domain, Year-Prefix, Templates, max. Schulstufen,
 *    Graph-Layout-Mode).
 *  - {@link normalizeGraphLayoutModeInSettings}: kleine Settings-Migration,
 *    damit `graphLayoutMode` immer einer der erlaubten Werte ist.
 *  - {@link normalizeClassLabel} / {@link deriveGradeFromClassLabel}: kleine
 *    String-Helfer rund um Klassen-Labels (Tool-spezifisch).
 *  - {@link buildDemoFromTenantSettings}: Versucht Demo-Zeilen aus den
 *    geladenen Tenant-Settings zu generieren (Klassen → Jahrgang/Klasse-Rows).
 *  - {@link buildDemoRows}: Liefert garantiert eine Demo-Struktur – primär
 *    aus Tenant-Settings, mit Fallback auf einen statischen Mock.
 *
 * Alle Funktionen sind **pure** (mit lazy Lookup auf `window.ms365TenantSettingsLoad`).
 */

import { compareDe } from '../../shared/utils/strings.js';
import { currentSchoolYearLabel } from './schulstruktur-sync-naming.js';
import { uid } from './schulstruktur-sync-tree.js';

/**
 * Lazy-Lookup auf `window.ms365TenantSettingsLoad`.
 * @returns {(() => any) | null}
 */
function getTenantSettingsLoader() {
    return typeof window !== 'undefined' && typeof window.ms365TenantSettingsLoad === 'function'
        ? window.ms365TenantSettingsLoad
        : null;
}

/**
 * Liest die Domain aus den Tenant-Settings, sofern verfügbar.
 * @returns {string} Domain oder Leerstring.
 */
export function getTenantSettingsDomainFallback() {
    try {
        const load = getTenantSettingsLoader();
        if (load) {
            const s = load();
            const d = s && s.domain ? String(s.domain).trim() : '';
            if (d) return d;
        }
    } catch {
        // ignore
    }
    return '';
}

/**
 * Default-Werte für das Anlegen-Schema (Templates, Year-Prefix, max. Schulstufen,
 * Graph-Layout). Werden später mit `state.settings` gemerged.
 *
 * @returns {{
 *   domain: string,
 *   kursteamYearPrefix: string,
 *   kursteamPattern: string,
 *   kursteamMailNickPattern: string,
 *   jgPrefix: string,
 *   jgUpper: boolean,
 *   argePrefix: string,
 *   argeUpper: boolean,
 *   maxSchulstufen: number,
 *   graphLayoutMode: 'horizontal' | 'vertical'
 * }}
 */
export function defaultAnlegenSchemas() {
    const now = new Date();
    const yy = String(now.getFullYear()).slice(-2);
    return {
        domain: getTenantSettingsDomainFallback() || 'ms365.schule',
        kursteamYearPrefix: 'SJ' + yy,
        kursteamPattern: '{yearPrefix} | {klasse} | {fach}',
        kursteamMailNickPattern: 'kt-{yearPrefix}-{klasse}-{fach}',
        jgPrefix: 'jg',
        jgUpper: true,
        argePrefix: 'arge',
        argeUpper: false,
        /** Schuljahrswechsel: maximale Stufenanzahl (3/4/5/8). */
        maxSchulstufen: 5,
        /** Organigramm: Geschwister nebeneinander (klassisch) oder untereinander (weniger Breite). */
        graphLayoutMode: 'horizontal'
    };
}

/**
 * Settings-Migration: stellt sicher, dass `graphLayoutMode` immer einer
 * der erlaubten Werte ist. Mutiert das Argument.
 *
 * @param {object | null | undefined} settings
 */
export function normalizeGraphLayoutModeInSettings(settings) {
    if (!settings || typeof settings !== 'object') return;
    settings.graphLayoutMode = settings.graphLayoutMode === 'vertical' ? 'vertical' : 'horizontal';
}

/**
 * Normalisiert das Klassen-Label (`code` bevorzugt vor `name`).
 *
 * @param {{ code?: string, name?: string } | null | undefined} c
 * @returns {string}
 */
export function normalizeClassLabel(c) {
    const code = c && c.code ? String(c.code).trim() : '';
    const name = c && c.name ? String(c.name).trim() : '';
    return code || name || '';
}

/**
 * Extrahiert die führende Schulstufe aus einem Klassen-Label (z. B. `1A` → `1`).
 *
 * @param {string} label
 * @returns {string} Stufe als String (1–2 Ziffern) oder Leerstring.
 */
export function deriveGradeFromClassLabel(label) {
    const m = String(label || '').trim().match(/^(\d{1,2})/);
    return m ? m[1] : '';
}

/**
 * Baut eine Demo-Strukturzeilen-Liste aus den geladenen Tenant-Settings
 * (Klassen → Jahrgang + Klasse + ein paar Standard-Gruppen).
 *
 * @returns {{ rows: any[], tenantSettings: any } | null} `null`, wenn keine
 *   Klassen in den Settings vorhanden sind (oder kein Loader registriert ist).
 */
export function buildDemoFromTenantSettings() {
    const load = getTenantSettingsLoader();
    if (!load) return null;
    const s = load();
    const classes = s && Array.isArray(s.classes) ? s.classes : [];
    if (!classes.length) return null;

    const schuljahr = currentSchoolYearLabel();
    const gradeMap = new Map();
    const rows = [];

    const grades = Array.from(
        new Set(classes.map((c) => deriveGradeFromClassLabel(normalizeClassLabel(c))).filter(Boolean))
    ).sort(compareDe);
    grades.forEach((g) => {
        const id = uid();
        gradeMap.set(g, id);
        rows.push({
            id,
            parentId: '',
            typ: 'Jahrgang',
            bezeichnung: 'Jahrgang ' + g,
            schuljahr,
            status: 'Aktiv',
            syncStatus: 'Ausstehend',
            letzteFehlermeldung: ''
        });
    });

    classes.forEach((c) => {
        const label = normalizeClassLabel(c);
        if (!label) return;
        const g = deriveGradeFromClassLabel(label);
        const parentId = g && gradeMap.has(g) ? gradeMap.get(g) : '';
        rows.push({
            id: uid(),
            parentId: parentId || '',
            typ: 'Klasse',
            bezeichnung: label,
            schuljahr,
            status: 'Aktiv',
            syncStatus: 'Ausstehend',
            letzteFehlermeldung: ''
        });
    });

    rows.push({
        id: uid(),
        parentId: '',
        typ: 'Gruppe',
        bezeichnung: 'Lehrer:innen',
        schuljahr,
        status: 'Aktiv',
        syncStatus: 'Ausstehend',
        letzteFehlermeldung: ''
    });
    rows.push({
        id: uid(),
        parentId: '',
        typ: 'Gruppe',
        bezeichnung: 'Schüler:innen',
        schuljahr,
        status: 'Aktiv',
        syncStatus: 'Ausstehend',
        letzteFehlermeldung: ''
    });

    return { rows, tenantSettings: s };
}

/**
 * Garantiert eine nicht-leere Demo-Struktur: erst Versuch aus Tenant-Settings,
 * sonst statischer Mock (Jahrgang 1+2, Klassen 1A/1B/2A, zwei ARGEn).
 *
 * @returns {{ rows: any[], tenantSettings: any | null }}
 */
export function buildDemoRows() {
    const fromTenant = buildDemoFromTenantSettings();
    if (fromTenant && fromTenant.rows && fromTenant.rows.length) return fromTenant;

    const schuljahr = currentSchoolYearLabel();
    const jg1 = {
        id: uid(),
        parentId: '',
        typ: 'Jahrgang',
        bezeichnung: 'Jahrgang 1',
        schuljahr,
        status: 'Aktiv',
        syncStatus: 'Ausstehend',
        letzteFehlermeldung: ''
    };
    const jg2 = {
        id: uid(),
        parentId: '',
        typ: 'Jahrgang',
        bezeichnung: 'Jahrgang 2',
        schuljahr,
        status: 'Aktiv',
        syncStatus: 'Ok',
        letzteFehlermeldung: ''
    };
    const k1a = {
        id: uid(),
        parentId: jg1.id,
        typ: 'Klasse',
        bezeichnung: '1A',
        schuljahr,
        status: 'Aktiv',
        syncStatus: 'Abweichung',
        letzteFehlermeldung: 'Mitgliedschaft weicht ab (Mock).'
    };
    const k1b = {
        id: uid(),
        parentId: jg1.id,
        typ: 'Klasse',
        bezeichnung: '1B',
        schuljahr,
        status: 'Aktiv',
        syncStatus: 'Ok',
        letzteFehlermeldung: ''
    };
    const k2a = {
        id: uid(),
        parentId: jg2.id,
        typ: 'Klasse',
        bezeichnung: '2A',
        schuljahr,
        status: 'Aktiv',
        syncStatus: 'Fehler',
        letzteFehlermeldung: 'Team konnte nicht bereitgestellt werden (Mock).'
    };
    const ar1 = {
        id: uid(),
        parentId: '',
        typ: 'Arbeitsgemeinschaft',
        bezeichnung: 'ARGE Robotik',
        schuljahr,
        status: 'Aktiv',
        syncStatus: 'Ok',
        letzteFehlermeldung: ''
    };
    const ar2 = {
        id: uid(),
        parentId: '',
        typ: 'Arbeitsgemeinschaft',
        bezeichnung: 'ARGE Chor',
        schuljahr,
        status: 'Inaktiv',
        syncStatus: 'Ausstehend',
        letzteFehlermeldung: ''
    };
    return { rows: [jg1, jg2, k1a, k1b, k2a, ar1, ar2], tenantSettings: null };
}
