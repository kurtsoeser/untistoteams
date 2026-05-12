/**
 * Tests für `src/tools/schulstruktur-sync/schulstruktur-sync-anlegen.js`.
 *
 *  - `defaultTenantTargetForTypeStr` / `defaultTenantVisibilityForTypeStr`
 *  - `resolveKursteamKlasseFachForRow` (Fallback + `window.ms365StructureRules`)
 *  - `computeTenantCreateSuggestionPure` (Jahrgang/Arge/Kursteam/Default)
 *  - `normRoleKey`
 */

import { describe, it, expect, beforeEach, afterEach } from 'vitest';

import {
    defaultTenantTargetForTypeStr,
    defaultTenantVisibilityForTypeStr,
    resolveKursteamKlasseFachForRow,
    computeTenantCreateSuggestionPure,
    normRoleKey
} from '../src/tools/schulstruktur-sync/schulstruktur-sync-anlegen.js';

describe('defaultTenantTargetForTypeStr', () => {
    it('liefert team für Kursteam und Klasse', () => {
        expect(defaultTenantTargetForTypeStr('Kursteam')).toBe('team');
        expect(defaultTenantTargetForTypeStr('Klasse')).toBe('team');
    });

    it('liefert group für andere Typen', () => {
        expect(defaultTenantTargetForTypeStr('Jahrgang')).toBe('group');
        expect(defaultTenantTargetForTypeStr('Arbeitsgemeinschaft')).toBe('group');
        expect(defaultTenantTargetForTypeStr('Gruppe')).toBe('group');
        expect(defaultTenantTargetForTypeStr('')).toBe('group');
        expect(defaultTenantTargetForTypeStr(undefined)).toBe('group');
        expect(defaultTenantTargetForTypeStr(null)).toBe('group');
    });
});

describe('defaultTenantVisibilityForTypeStr', () => {
    it('liefert HiddenMembership für Kursteam', () => {
        expect(defaultTenantVisibilityForTypeStr('Kursteam')).toBe('HiddenMembership');
    });

    it('liefert Private für andere Typen', () => {
        expect(defaultTenantVisibilityForTypeStr('Klasse')).toBe('Private');
        expect(defaultTenantVisibilityForTypeStr('Jahrgang')).toBe('Private');
        expect(defaultTenantVisibilityForTypeStr('Arbeitsgemeinschaft')).toBe('Private');
        expect(defaultTenantVisibilityForTypeStr('Gruppe')).toBe('Private');
        expect(defaultTenantVisibilityForTypeStr('')).toBe('Private');
        expect(defaultTenantVisibilityForTypeStr(undefined)).toBe('Private');
    });
});

describe('resolveKursteamKlasseFachForRow', () => {
    beforeEach(() => {
        globalThis.window = {};
    });
    afterEach(() => {
        delete globalThis.window;
    });

    it('nutzt window.ms365StructureRules.resolveKursteamKlasseFach falls vorhanden', () => {
        let called = false;
        globalThis.window.ms365StructureRules = {
            resolveKursteamKlasseFach: (row, rows) => {
                called = true;
                expect(row).toEqual({ id: 'kt1' });
                expect(rows).toEqual([{ id: 'kt1' }]);
                return { klasse: '5A', fach: 'M', hasBoth: true, source: 'rules' };
            }
        };
        const r = resolveKursteamKlasseFachForRow({ id: 'kt1' }, [{ id: 'kt1' }]);
        expect(called).toBe(true);
        expect(r).toEqual({ klasse: '5A', fach: 'M', hasBoth: true, source: 'rules' });
    });

    it('Fallback: nimmt ktKlasse + ktFach direkt vom row', () => {
        const r = resolveKursteamKlasseFachForRow(
            { id: 'kt1', ktKlasse: '5A', ktFach: 'M' },
            []
        );
        expect(r).toEqual({ klasse: '5A', fach: 'M', hasBoth: true });
    });

    it('Fallback: leitet klasse vom Parent-Typ Klasse ab, wenn ktKlasse leer', () => {
        const rows = [{ id: 'parent1', typ: 'Klasse', bezeichnung: '5A' }];
        const r = resolveKursteamKlasseFachForRow(
            { id: 'kt1', parentId: 'parent1', ktFach: 'Mathematik' },
            rows
        );
        expect(r).toEqual({ klasse: '5A', fach: 'Mathematik', hasBoth: true });
    });

    it('Fallback: ignoriert Parent wenn nicht Typ Klasse', () => {
        const rows = [{ id: 'parent1', typ: 'Jahrgang', bezeichnung: '5' }];
        const r = resolveKursteamKlasseFachForRow(
            { id: 'kt1', parentId: 'parent1', ktFach: 'Mathematik' },
            rows
        );
        expect(r).toEqual({ klasse: '', fach: 'Mathematik', hasBoth: false });
    });

    it('hasBoth ist false, wenn entweder klasse oder fach leer ist', () => {
        expect(resolveKursteamKlasseFachForRow({ ktKlasse: '5A' }, []).hasBoth).toBe(false);
        expect(resolveKursteamKlasseFachForRow({ ktFach: 'M' }, []).hasBoth).toBe(false);
        expect(resolveKursteamKlasseFachForRow({}, []).hasBoth).toBe(false);
    });

    it('verarbeitet null/undefined ohne crash', () => {
        const r = resolveKursteamKlasseFachForRow(null, null);
        expect(r).toEqual({ klasse: '', fach: '', hasBoth: false });
    });
});

describe('computeTenantCreateSuggestionPure', () => {
    it('liefert leeres Ergebnis bei leerer Bezeichnung', () => {
        expect(computeTenantCreateSuggestionPure({ typ: 'Jahrgang', bezeichnung: '' })).toEqual({
            displayName: '',
            mailNick: ''
        });
        expect(computeTenantCreateSuggestionPure({ typ: 'Klasse' })).toEqual({
            displayName: '',
            mailNick: ''
        });
    });

    it('Jahrgang: nutzt jgYear+jgSuffix für mailNick', () => {
        const r = computeTenantCreateSuggestionPure({
            typ: 'Jahrgang',
            bezeichnung: 'Jahrgang 5',
            jgYear: '2026',
            jgSuffix: 'A'
        });
        expect(r.displayName).toBe('Jahrgang 5');
        expect(r.mailNick).toBeTruthy();
        expect(typeof r.mailNick).toBe('string');
    });

    it('Jahrgang: fällt auf Label-mailNick zurück, wenn jgYear/jgSuffix fehlen', () => {
        const r = computeTenantCreateSuggestionPure({
            typ: 'Jahrgang',
            bezeichnung: 'Jahrgang 5'
        });
        expect(r.displayName).toBe('Jahrgang 5');
        expect(r.mailNick).toBeTruthy();
    });

    it('Arbeitsgemeinschaft: nutzt argeCode wenn vorhanden', () => {
        const r = computeTenantCreateSuggestionPure({
            typ: 'Arbeitsgemeinschaft',
            bezeichnung: 'AG Schach',
            argeCode: 'SCH'
        });
        expect(r.displayName).toBe('AG Schach');
        expect(r.mailNick).toBeTruthy();
    });

    it('Arbeitsgemeinschaft: Label-Fallback wenn argeCode fehlt', () => {
        const r = computeTenantCreateSuggestionPure({
            typ: 'Arbeitsgemeinschaft',
            bezeichnung: 'AG Schach'
        });
        expect(r.displayName).toBe('AG Schach');
        expect(r.mailNick).toBeTruthy();
    });

    it('Kursteam: nutzt klasse+fach via resolveKlasseFach-Callback', () => {
        const r = computeTenantCreateSuggestionPure(
            { typ: 'Kursteam', bezeichnung: 'M 5A' },
            { kursteamYearPrefix: '26', kursteamMailNickPattern: '{yearPrefix}-{klasse}-{fach}' },
            () => ({ klasse: '5A', fach: 'M' })
        );
        expect(r.displayName).toBe('M 5A');
        expect(r.mailNick).toContain('26');
        expect(r.mailNick).toContain('5a');
        expect(r.mailNick).toContain('m');
    });

    it('Kursteam: Label-Fallback wenn resolveKlasseFach unvollständig', () => {
        const r = computeTenantCreateSuggestionPure(
            { typ: 'Kursteam', bezeichnung: 'M 5A' },
            {},
            () => ({ klasse: '', fach: 'M' })
        );
        expect(r.displayName).toBe('M 5A');
        expect(r.mailNick).toBeTruthy();
    });

    it('Kursteam: Label-Fallback wenn kein Callback übergeben wird', () => {
        const r = computeTenantCreateSuggestionPure({
            typ: 'Kursteam',
            bezeichnung: 'M 5A'
        });
        expect(r.displayName).toBe('M 5A');
        expect(r.mailNick).toBeTruthy();
    });

    it('andere Typen: einfach Label-mailNick', () => {
        const r = computeTenantCreateSuggestionPure({
            typ: 'Klasse',
            bezeichnung: 'Klasse 5A'
        });
        expect(r.displayName).toBe('Klasse 5A');
        expect(r.mailNick).toBeTruthy();
    });

    it('trimmt displayName-Whitespace', () => {
        const r = computeTenantCreateSuggestionPure({
            typ: 'Klasse',
            bezeichnung: '  Klasse 5A  '
        });
        expect(r.displayName).toBe('Klasse 5A');
    });

    it('nimmt defaultAnlegenSchemas wenn kein schemaState übergeben wird', () => {
        const r = computeTenantCreateSuggestionPure(
            { typ: 'Jahrgang', bezeichnung: 'Jahrgang 5', jgYear: '2026', jgSuffix: 'A' },
            null
        );
        expect(r.mailNick).toBeTruthy();
    });
});

describe('normRoleKey', () => {
    it('trimmt, lowercased und entfernt :;', () => {
        expect(normRoleKey('  Klassenvorstand:  ')).toBe('klassenvorstand');
        expect(normRoleKey('AG-Leitung;')).toBe('ag-leitung');
        expect(normRoleKey('AG: Leitung')).toBe('ag leitung');
    });

    it('komprimiert mehrfache Whitespaces zu einem Space', () => {
        expect(normRoleKey('  AG    Leitung  ')).toBe('ag leitung');
        expect(normRoleKey('A\t\nB\n\rC')).toBe('a b c');
    });

    it('liefert leeren String für null/undefined/leer', () => {
        expect(normRoleKey(null)).toBe('');
        expect(normRoleKey(undefined)).toBe('');
        expect(normRoleKey('')).toBe('');
        expect(normRoleKey('   ')).toBe('');
    });

    it('lässt Bindestriche und andere Sonderzeichen unverändert', () => {
        expect(normRoleKey('AG-Schach')).toBe('ag-schach');
        expect(normRoleKey('Klassenvorstand/Stv.')).toBe('klassenvorstand/stv.');
    });

    it('zwei Eingaben, die nur in Trim/Case/:; abweichen, sind gleich', () => {
        expect(normRoleKey('Klassenvorstand:')).toBe(normRoleKey('klassenvorstand '));
        expect(normRoleKey('AG; Leitung')).toBe(normRoleKey('ag leitung'));
    });
});
