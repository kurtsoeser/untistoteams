/**
 * Tests für `src/tools/schulstruktur-sync/schulstruktur-sync-mapping.js`.
 *
 *  - `parseMatchSelectValue`
 *  - `applyMatchLinkUpdate`
 *  - `persistedMatchSelectValuePure`
 *  - `computeMatchDraftDirty`
 */

import { describe, it, expect } from 'vitest';

import {
    parseMatchSelectValue,
    applyMatchLinkUpdate,
    persistedMatchSelectValuePure,
    computeMatchDraftDirty
} from '../src/tools/schulstruktur-sync/schulstruktur-sync-mapping.js';

describe('parseMatchSelectValue', () => {
    it('liefert leere IDs für leere/whitespace Eingaben', () => {
        expect(parseMatchSelectValue('')).toEqual({ tenantGroupId: '', tenantUserId: '' });
        expect(parseMatchSelectValue('   ')).toEqual({ tenantGroupId: '', tenantUserId: '' });
        expect(parseMatchSelectValue(null)).toEqual({ tenantGroupId: '', tenantUserId: '' });
        expect(parseMatchSelectValue(undefined)).toEqual({ tenantGroupId: '', tenantUserId: '' });
    });

    it('erkennt u:-Präfix als User-ID', () => {
        expect(parseMatchSelectValue('u:abc-123')).toEqual({
            tenantGroupId: '',
            tenantUserId: 'abc-123'
        });
    });

    it('erkennt g:-Präfix als Gruppen-ID', () => {
        expect(parseMatchSelectValue('g:grp-1')).toEqual({
            tenantGroupId: 'grp-1',
            tenantUserId: ''
        });
    });

    it('behandelt einen unpräfigierten Wert als Gruppen-ID', () => {
        expect(parseMatchSelectValue('legacy-id')).toEqual({
            tenantGroupId: 'legacy-id',
            tenantUserId: ''
        });
    });

    it('trimmt Whitespace nach dem Präfix', () => {
        expect(parseMatchSelectValue('  u:  user-1  ')).toEqual({
            tenantGroupId: '',
            tenantUserId: 'user-1'
        });
    });
});

describe('applyMatchLinkUpdate', () => {
    it('fügt einen neuen Eintrag mit Group-ID hinzu', () => {
        const before = {};
        const after = applyMatchLinkUpdate(before, 'row1', {
            tenantGroupId: 'g1',
            note: 'Notiz'
        });
        expect(after.row1.tenantGroupId).toBe('g1');
        expect(after.row1.tenantUserId).toBe('');
        expect(after.row1.note).toBe('Notiz');
        expect(typeof after.row1.updatedAt).toBe('string');
    });

    it('fügt einen neuen Eintrag mit User-ID hinzu', () => {
        const after = applyMatchLinkUpdate({}, 'row1', { tenantUserId: 'u1' });
        expect(after.row1).toEqual({
            tenantGroupId: '',
            tenantUserId: 'u1',
            note: '',
            updatedAt: expect.any(String)
        });
    });

    it('mutiert die Eingabe nicht (immutable)', () => {
        const before = { row1: { tenantGroupId: 'old', tenantUserId: '', note: '', updatedAt: 'x' } };
        const after = applyMatchLinkUpdate(before, 'row1', { tenantGroupId: 'new' });
        expect(after).not.toBe(before);
        expect(before.row1.tenantGroupId).toBe('old');
        expect(after.row1.tenantGroupId).toBe('new');
    });

    it('entfernt einen Eintrag, wenn beide IDs leer sind', () => {
        const before = { row1: { tenantGroupId: 'g1' }, row2: { tenantGroupId: 'g2' } };
        const after = applyMatchLinkUpdate(before, 'row1', {});
        expect(after.row1).toBeUndefined();
        expect(after.row2).toEqual({ tenantGroupId: 'g2' });
    });

    it('lässt die Map unverändert, wenn structureId leer ist', () => {
        const before = { row1: { tenantGroupId: 'g1' } };
        const after = applyMatchLinkUpdate(before, '', { tenantGroupId: 'x' });
        expect(after).toBe(before);
    });

    it('toleriert null/undefined als currentLinks', () => {
        const after = applyMatchLinkUpdate(null, 'row1', { tenantGroupId: 'g1' });
        expect(after.row1.tenantGroupId).toBe('g1');
    });

    it('konvertiert numerische structureId zu String-Key', () => {
        const after = applyMatchLinkUpdate({}, 42, { tenantGroupId: 'g1' });
        expect(after['42'].tenantGroupId).toBe('g1');
    });

    it('verwendet den expliziten updatedAt, falls übergeben', () => {
        const after = applyMatchLinkUpdate({}, 'row1', {
            tenantGroupId: 'g1',
            updatedAt: '2024-01-01T00:00:00Z'
        });
        expect(after.row1.updatedAt).toBe('2024-01-01T00:00:00Z');
    });

    it('überschreibt einen bestehenden Eintrag', () => {
        const before = { row1: { tenantGroupId: 'old', note: 'alt' } };
        const after = applyMatchLinkUpdate(before, 'row1', {
            tenantGroupId: 'neu',
            note: 'neu'
        });
        expect(after.row1.tenantGroupId).toBe('neu');
        expect(after.row1.note).toBe('neu');
    });
});

describe('persistedMatchSelectValuePure', () => {
    it('liefert leeren String für leere/fehlende Daten', () => {
        expect(persistedMatchSelectValuePure('row1', null)).toBe('');
        expect(persistedMatchSelectValuePure('row1', {})).toBe('');
        expect(persistedMatchSelectValuePure('', { row1: { tenantGroupId: 'g1' } })).toBe('');
    });

    it('liefert u:<id> bei vorhandener tenantUserId', () => {
        const links = { row1: { tenantUserId: 'u1', tenantGroupId: 'g1' } };
        expect(persistedMatchSelectValuePure('row1', links)).toBe('u:u1');
    });

    it('liefert g:<id> bei vorhandener tenantGroupId', () => {
        const links = { row1: { tenantGroupId: 'g1' } };
        expect(persistedMatchSelectValuePure('row1', links)).toBe('g:g1');
    });

    it('liefert u:<id> wenn isUserId-Callback die Group-ID als User markiert', () => {
        const links = { row1: { tenantGroupId: 'g1' } };
        const isUserId = (id) => id === 'g1';
        expect(persistedMatchSelectValuePure('row1', links, isUserId)).toBe('u:g1');
    });

    it('isUserId-Callback wird nur befragt, wenn tenantUserId fehlt', () => {
        const links = { row1: { tenantUserId: 'u1', tenantGroupId: 'g1' } };
        const isUserId = () => {
            throw new Error('should not be called');
        };
        expect(persistedMatchSelectValuePure('row1', links, isUserId)).toBe('u:u1');
    });

    it('konvertiert numerische structureId', () => {
        const links = { 42: { tenantGroupId: 'g1' } };
        expect(persistedMatchSelectValuePure(42, links)).toBe('g:g1');
    });

    it('liefert leeren String, wenn beide IDs leer sind', () => {
        const links = { row1: { tenantGroupId: '', tenantUserId: '' } };
        expect(persistedMatchSelectValuePure('row1', links)).toBe('');
    });
});

describe('computeMatchDraftDirty', () => {
    it('false: nichts gespeichert, leere Eingaben', () => {
        expect(computeMatchDraftDirty(null, '', '')).toBe(false);
        expect(computeMatchDraftDirty(undefined, '', '')).toBe(false);
    });

    it('true: nichts gespeichert, aber Eingabe gesetzt', () => {
        expect(computeMatchDraftDirty(null, 'g:g1', '')).toBe(true);
        expect(computeMatchDraftDirty(null, '', 'Notiz')).toBe(true);
    });

    it('false: Eingabe entspricht dem gespeicherten Stand (group)', () => {
        const saved = { tenantGroupId: 'g1', note: 'Notiz' };
        expect(computeMatchDraftDirty(saved, 'g:g1', 'Notiz')).toBe(false);
    });

    it('false: Eingabe entspricht dem gespeicherten Stand (user)', () => {
        const saved = { tenantUserId: 'u1', note: '' };
        expect(computeMatchDraftDirty(saved, 'u:u1', '')).toBe(false);
    });

    it('true: anderer Wert', () => {
        const saved = { tenantGroupId: 'g1', note: '' };
        expect(computeMatchDraftDirty(saved, 'g:g2', '')).toBe(true);
    });

    it('true: andere Notiz', () => {
        const saved = { tenantGroupId: 'g1', note: 'alt' };
        expect(computeMatchDraftDirty(saved, 'g:g1', 'neu')).toBe(true);
    });

    it('Whitespace im UI wird beim Vergleich ignoriert', () => {
        const saved = { tenantGroupId: 'g1', note: 'Notiz' };
        expect(computeMatchDraftDirty(saved, '  g:g1  ', '  Notiz  ')).toBe(false);
    });

    it('isUserId-Callback macht den g:-Eintrag treffsicher', () => {
        const saved = { tenantGroupId: 'g1', note: '' };
        // gleichgesetzt: g1 ist eigentlich ein User → savedVal = 'u:g1'
        expect(computeMatchDraftDirty(saved, 'u:g1', '', (id) => id === 'g1')).toBe(false);
        expect(computeMatchDraftDirty(saved, 'g:g1', '', (id) => id === 'g1')).toBe(true);
    });

    it('Cleared-Saved + Eingabe leer = false', () => {
        const saved = { tenantGroupId: '', tenantUserId: '', note: '' };
        expect(computeMatchDraftDirty(saved, '', '')).toBe(false);
    });
});
