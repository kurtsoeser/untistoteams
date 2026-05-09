import { describe, it, expect } from 'vitest';
import {
    allowedStructureChildTypes,
    structureTreeRowShowsAddChildControl
} from '../src/tools/schulstruktur-sync/schulstruktur-sync-helpers.js';

describe('schulstruktur-sync-helpers', () => {
    it('allowedStructureChildTypes: bekannte Eltern-Typen', () => {
        expect(allowedStructureChildTypes('SchuelerInnen')).toEqual(['Jahrgang']);
        expect(allowedStructureChildTypes('Klasse')).toEqual(['Kursteam', 'Gruppe']);
        expect(allowedStructureChildTypes('Unbekannt')).toEqual([]);
    });

    it('structureTreeRowShowsAddChildControl: Modus und virtuelle Wurzel', () => {
        expect(structureTreeRowShowsAddChildControl('struktur', null)).toBe(false);
        expect(structureTreeRowShowsAddChildControl('liste', { virtual: true, kind: 'root', typ: 'Klasse' })).toBe(false);
        expect(
            structureTreeRowShowsAddChildControl('struktur', {
                virtual: true,
                kind: 'root',
                typ: 'SchuelerInnen'
            })
        ).toBe(true);
        expect(
            structureTreeRowShowsAddChildControl('match', {
                virtual: true,
                kind: 'folder',
                typ: 'Klasse'
            })
        ).toBe(false);
    });

    it('structureTreeRowShowsAddChildControl: Zeile mit r.typ', () => {
        expect(
            structureTreeRowShowsAddChildControl('match', {
                virtual: false,
                r: { typ: 'Jahrgang' }
            })
        ).toBe(true);
        expect(structureTreeRowShowsAddChildControl('struktur', { virtual: false })).toBe(false);
    });
});
