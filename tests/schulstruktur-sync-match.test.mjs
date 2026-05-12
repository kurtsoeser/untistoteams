import { describe, it, expect } from 'vitest';
import {
    normKey,
    suggestTenantGroupForUnitFromList,
    suggestTenantUserForPersonFromList,
    suggestTenantMatchSelectValue,
    formatEntraUserPickLabel,
    matchTenantFilterNeedle,
    matchTenantHaystackForGroup,
    matchTenantHaystackForUser
} from '../src/tools/schulstruktur-sync/schulstruktur-sync-match.js';

describe('normKey', () => {
    it('trimmt und lowert', () => {
        expect(normKey('  Hallo  ')).toBe('hallo');
    });

    it('fasst Whitespace zusammen', () => {
        expect(normKey('a   b\tc')).toBe('a b c');
    });

    it('entfernt Sonderzeichen, behält Punkt/Dash/Underscore', () => {
        expect(normKey('1A: Klasse!')).toBe('1a klasse');
        expect(normKey('foo.bar-baz_qux')).toBe('foo.bar-baz_qux');
    });

    it('behält Umlaute', () => {
        expect(normKey('Schüler:innen')).toBe('schülerinnen');
    });

    it('liefert Leerstring für null/undefined/leer', () => {
        expect(normKey(null)).toBe('');
        expect(normKey(undefined)).toBe('');
        expect(normKey('')).toBe('');
        expect(normKey('   ')).toBe('');
    });
});

describe('suggestTenantGroupForUnitFromList', () => {
    const groups = [
        { id: 'g1', bezeichnung: '1A', alias: '' },
        { id: 'g2', bezeichnung: 'Klasse 1B', alias: '1b' },
        { id: 'g3', bezeichnung: 'Lehrer:innen', alias: 'lehrer' },
        { id: 'g4', bezeichnung: 'ARGE-Robotik', alias: '' }
    ];

    it('exakter Bezeichnung-Match', () => {
        expect(suggestTenantGroupForUnitFromList({ bezeichnung: '1A' }, groups)).toBe('g1');
    });

    it('exakter Alias-Match (geht vor "enthält")', () => {
        expect(suggestTenantGroupForUnitFromList({ bezeichnung: '1B' }, groups)).toBe('g2');
    });

    it('Substring-Match auf Bezeichnung', () => {
        expect(suggestTenantGroupForUnitFromList({ bezeichnung: 'Robotik' }, groups)).toBe('g4');
    });

    it('liefert Leerstring ohne Match', () => {
        expect(suggestTenantGroupForUnitFromList({ bezeichnung: 'XYZ' }, groups)).toBe('');
    });

    it('liefert Leerstring bei leerer Liste/unit', () => {
        expect(suggestTenantGroupForUnitFromList(null, groups)).toBe('');
        expect(suggestTenantGroupForUnitFromList({ bezeichnung: '' }, groups)).toBe('');
        expect(suggestTenantGroupForUnitFromList({ bezeichnung: '1A' }, [])).toBe('');
    });

    it('robust gegen list nicht-Array', () => {
        expect(suggestTenantGroupForUnitFromList({ bezeichnung: '1A' }, null)).toBe('');
    });
});

describe('suggestTenantUserForPersonFromList', () => {
    const users = [
        { id: 'u1', displayName: 'Max Mustermann', userPrincipalName: 'max@x.at', mail: 'max@x.at' },
        { id: 'u2', displayName: 'Anna Beispiel', userPrincipalName: 'anna@x.at', mail: 'anna@x.at' },
        { id: 'u3', displayName: 'Hans Schmid', userPrincipalName: 'hans@x.at', mail: 'hans@x.at' }
    ];

    it('liefert Leerstring für non-Person-Unit', () => {
        expect(suggestTenantUserForPersonFromList({ typ: 'Gruppe' }, users)).toBe('');
    });

    it('exakter displayName-Match', () => {
        expect(suggestTenantUserForPersonFromList({ typ: 'Person', personName: 'Max Mustermann' }, users)).toBe('u1');
    });

    it('exakter Email-Match', () => {
        expect(suggestTenantUserForPersonFromList({ typ: 'Person', personEmail: 'anna@x.at' }, users)).toBe('u2');
    });

    it('Bezeichnung wird als Fallback genutzt', () => {
        expect(suggestTenantUserForPersonFromList({ typ: 'Person', bezeichnung: 'Hans Schmid' }, users)).toBe('u3');
    });

    it('Substring-Match scored niedriger als exakt', () => {
        const users2 = [
            { id: 'u1', displayName: 'Max Mustermann', userPrincipalName: 'mm@x.at' },
            { id: 'u2', displayName: 'Maximilian Müller', userPrincipalName: 'mm2@x.at' }
        ];
        // "Max" → Substring in beiden → der erste gewinnt (gleicher Score, erstes Auftreten)
        const id = suggestTenantUserForPersonFromList({ typ: 'Person', personName: 'Max' }, users2);
        expect(['u1', 'u2']).toContain(id);
    });

    it('liefert Leerstring ohne Match', () => {
        expect(suggestTenantUserForPersonFromList({ typ: 'Person', personName: 'XYZ' }, users)).toBe('');
    });

    it('liefert Leerstring bei leeren keys', () => {
        expect(suggestTenantUserForPersonFromList({ typ: 'Person' }, users)).toBe('');
    });
});

describe('suggestTenantMatchSelectValue', () => {
    const groups = [{ id: 'g1', bezeichnung: '1A' }];
    const users = [{ id: 'u1', displayName: 'Max Mustermann' }];

    it('Person → "u:<id>"', () => {
        expect(suggestTenantMatchSelectValue({ typ: 'Person', personName: 'Max Mustermann' }, groups, users)).toBe('u:u1');
    });

    it('Gruppe → "g:<id>"', () => {
        expect(suggestTenantMatchSelectValue({ typ: 'Klasse', bezeichnung: '1A' }, groups, users)).toBe('g:g1');
    });

    it('Person ohne User-Match → fällt auf Gruppen-Match zurück', () => {
        const u = { typ: 'Person', bezeichnung: '1A' };
        expect(suggestTenantMatchSelectValue(u, groups, users)).toBe('g:g1');
    });

    it('liefert Leerstring ohne Match', () => {
        expect(suggestTenantMatchSelectValue({ bezeichnung: 'XYZ' }, groups, users)).toBe('');
        expect(suggestTenantMatchSelectValue(null, groups, users)).toBe('');
    });
});

describe('formatEntraUserPickLabel', () => {
    it('DisplayName + UPN + Benutzer-Suffix', () => {
        expect(formatEntraUserPickLabel({ displayName: 'Max Mustermann', userPrincipalName: 'max@x.at' }))
            .toBe('Max Mustermann · max@x.at · Benutzer');
    });

    it('nur UPN, wenn DisplayName fehlt', () => {
        expect(formatEntraUserPickLabel({ userPrincipalName: 'a@b.c' })).toBe('a@b.c · Benutzer');
    });

    it('fällt auf mail zurück, wenn UPN fehlt', () => {
        expect(formatEntraUserPickLabel({ displayName: 'Max', mail: 'm@x.at' }))
            .toBe('Max · m@x.at · Benutzer');
    });

    it('lässt UPN weg, wenn er == DisplayName (case-insensitive)', () => {
        expect(formatEntraUserPickLabel({ displayName: 'max@x.at', userPrincipalName: 'MAX@x.at' }))
            .toBe('max@x.at · Benutzer');
    });

    it('liefert Leerstring für null/non-object', () => {
        expect(formatEntraUserPickLabel(null)).toBe('');
        expect(formatEntraUserPickLabel('string')).toBe('');
    });

    it('fällt auf id zurück, wenn nichts anderes vorhanden', () => {
        expect(formatEntraUserPickLabel({ id: 'u1' })).toBe('u1 · Benutzer');
    });
});

describe('matchTenantFilterNeedle', () => {
    it('trimmt, lowert, fasst Whitespace zusammen', () => {
        expect(matchTenantFilterNeedle('  Foo   BAR  ')).toBe('foo bar');
    });

    it('liefert Leerstring für null/undefined', () => {
        expect(matchTenantFilterNeedle(null)).toBe('');
        expect(matchTenantFilterNeedle(undefined)).toBe('');
    });
});

describe('matchTenantHaystackForGroup', () => {
    it('vereint alle relevanten Felder lowercase', () => {
        const hay = matchTenantHaystackForGroup({
            id: 'g1',
            bezeichnung: 'Klasse 1A',
            typ: 'Klasse',
            alias: '1A',
            mail: 'k1a@x.at',
            description: 'Erste A'
        });
        expect(hay).toContain('klasse 1a');
        expect(hay).toContain('1a');
        expect(hay).toContain('k1a@x.at');
        expect(hay).toContain('erste a');
        expect(hay).toContain('g1');
    });

    it('robust gegen null', () => {
        expect(matchTenantHaystackForGroup(null)).toBe('     ');
    });
});

describe('matchTenantHaystackForUser', () => {
    it('enthält DisplayName, UPN, Mail, Id und Pick-Label', () => {
        const hay = matchTenantHaystackForUser({
            id: 'u1',
            displayName: 'Max Mustermann',
            userPrincipalName: 'max@x.at',
            mail: 'max@x.at'
        });
        expect(hay).toContain('max mustermann');
        expect(hay).toContain('max@x.at');
        expect(hay).toContain('u1');
        expect(hay).toContain('benutzer');
    });

    it('robust gegen null', () => {
        expect(matchTenantHaystackForUser(null)).toMatch(/^\s+$/);
    });
});
