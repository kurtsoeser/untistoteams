import { describe, it, expect } from 'vitest';
import { loadScript } from './kursteams-vm.mjs';

describe('structure-rules', () => {
    it('canReparent: strict matrix', () => {
        const ctx = loadScript('src/tools/schulstruktur-sync/structure-rules.js');
        const { canReparent } = ctx.ms365StructureRules;

        expect(canReparent('Jahrgang', 'SchuelerInnen')).toBe(true);
        expect(canReparent('Klasse', 'Jahrgang')).toBe(true);
        expect(canReparent('Kursteam', 'Klasse')).toBe(true);
        expect(canReparent('Gruppe', 'Klasse')).toBe(true);
        expect(canReparent('Arbeitsgemeinschaft', 'LehrerInnen')).toBe(true);
        expect(canReparent('Gruppe', 'LehrerInnen')).toBe(true);
        expect(canReparent('Gruppe', 'Gruppe')).toBe(true);
        expect(canReparent('Person', 'Gruppe')).toBe(true);

        // invalid
        expect(canReparent('Klasse', 'SchuelerInnen')).toBe(false);
        expect(canReparent('Kursteam', 'Jahrgang')).toBe(false);
        expect(canReparent('Arbeitsgemeinschaft', 'SchuelerInnen')).toBe(false);
        expect(canReparent('Jahrgang', 'LehrerInnen')).toBe(false);
        expect(canReparent('Gruppe', 'SchuelerInnen')).toBe(false);
        expect(canReparent('Person', 'Klasse')).toBe(false);
        expect(canReparent('Person', 'LehrerInnen')).toBe(false);
    });

    it('inferRootForType', () => {
        const ctx = loadScript('src/tools/schulstruktur-sync/structure-rules.js');
        const { inferRootForType } = ctx.ms365StructureRules;
        expect(inferRootForType('Jahrgang')).toBe('SchuelerInnen');
        expect(inferRootForType('Klasse')).toBe('SchuelerInnen');
        expect(inferRootForType('Kursteam')).toBe('SchuelerInnen');
        expect(inferRootForType('Arbeitsgemeinschaft')).toBe('LehrerInnen');
        expect(inferRootForType('Gruppe')).toBe('LehrerInnen');
        expect(inferRootForType('Person')).toBe('LehrerInnen');
    });

    it('resolveKursteamKlasseFach: Feld oder Eltern-Klasse', () => {
        const ctx = loadScript('src/tools/schulstruktur-sync/structure-rules.js');
        const { resolveKursteamKlasseFach } = ctx.ms365StructureRules;
        const klasseId = 'k-1';
        const rows = [
            { id: klasseId, typ: 'Klasse', bezeichnung: '3B' },
            { id: 'kt-1', typ: 'Kursteam', parentId: klasseId, ktKlasse: '', ktFach: 'M' }
        ];
        expect(resolveKursteamKlasseFach(rows[1], rows)).toEqual({
            klasse: '3B',
            fach: 'M',
            hasBoth: true
        });
        expect(
            resolveKursteamKlasseFach({ id: 'x', typ: 'Kursteam', parentId: '', ktKlasse: '1A', ktFach: 'D' }, rows)
        ).toEqual({ klasse: '1A', fach: 'D', hasBoth: true });
        expect(
            resolveKursteamKlasseFach({ id: 'y', typ: 'Kursteam', parentId: klasseId, ktKlasse: '', ktFach: '' }, rows)
        ).toEqual({ klasse: '3B', fach: '', hasBoth: false });
    });
});

