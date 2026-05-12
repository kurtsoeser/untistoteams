import { describe, it, expect } from 'vitest';
import {
    formatDateTimeAT,
    pillClass,
    computeStats,
    computeTenantStats,
    applyFiltersPure
} from '../src/tools/schulstruktur-sync/schulstruktur-sync-stats.js';

describe('formatDateTimeAT', () => {
    it('formatiert ein gültiges ISO-Datum (de-AT)', () => {
        const out = formatDateTimeAT('2026-05-12T14:30:00Z');
        // Format: "tt.mm.jjjj hh:mm" – exakte Zeit hängt von TZ ab.
        expect(out).toMatch(/^\d{2}\.\d{2}\.\d{4}\s+\d{2}:\d{2}$/);
    });

    it('liefert Leerstring für null/undefined/leer', () => {
        expect(formatDateTimeAT(null)).toBe('');
        expect(formatDateTimeAT(undefined)).toBe('');
        expect(formatDateTimeAT('')).toBe('');
        expect(formatDateTimeAT('   ')).toBe('');
    });

    it('gibt den Original-String bei ungültigem Datum zurück', () => {
        expect(formatDateTimeAT('kein-datum')).toBe('kein-datum');
    });

    it('trimmt den Eingabewert vor dem Parsen', () => {
        const out = formatDateTimeAT('  2026-05-12T08:00:00Z  ');
        expect(out).toMatch(/^\d{2}\.\d{2}\.\d{4}/);
    });
});

describe('pillClass', () => {
    it('Ok → ok', () => {
        expect(pillClass('Ok')).toBe('ok');
    });

    it('Abweichung → warn', () => {
        expect(pillClass('Abweichung')).toBe('warn');
    });

    it('Fehler → err', () => {
        expect(pillClass('Fehler')).toBe('err');
    });

    it('alles andere → Leerstring', () => {
        expect(pillClass('Ausstehend')).toBe('');
        expect(pillClass('')).toBe('');
        expect(pillClass(null)).toBe('');
        expect(pillClass(undefined)).toBe('');
    });
});

describe('computeStats', () => {
    it('zählt total/aktiv/abw/err', () => {
        const rows = [
            { status: 'Aktiv', syncStatus: 'Ok' },
            { status: 'Aktiv', syncStatus: 'Abweichung' },
            { status: 'Aktiv', syncStatus: 'Fehler' },
            { status: 'Inaktiv', syncStatus: 'Ok' }
        ];
        expect(computeStats(rows)).toEqual({ total: 4, aktiv: 3, abw: 1, err: 1 });
    });

    it('robust gegen null/undefined-Einträge', () => {
        const rows = [null, undefined, { status: 'Aktiv', syncStatus: 'Ok' }];
        expect(computeStats(rows)).toEqual({ total: 3, aktiv: 1, abw: 0, err: 0 });
    });

    it('leeres Array → Nullen', () => {
        expect(computeStats([])).toEqual({ total: 0, aktiv: 0, abw: 0, err: 0 });
    });

    it('robust gegen nicht-Array Inputs', () => {
        expect(computeStats(null)).toEqual({ total: 0, aktiv: 0, abw: 0, err: 0 });
        expect(computeStats(undefined)).toEqual({ total: 0, aktiv: 0, abw: 0, err: 0 });
    });
});

describe('computeTenantStats', () => {
    it('zählt total/teams/m365/sec', () => {
        const rows = [
            { typ: 'Team' },
            { typ: 'Team' },
            { typ: 'Gruppe' },
            { typ: 'Sicherheitsgruppe' },
            { typ: 'E‑Mail‑Sicherheitsgruppe' }, // Unicode-Bindestrich!
            { typ: 'Sonstiges' }
        ];
        expect(computeTenantStats(rows)).toEqual({ total: 6, teams: 2, m365: 1, sec: 2 });
    });

    it('leeres Array → Nullen', () => {
        expect(computeTenantStats([])).toEqual({ total: 0, teams: 0, m365: 0, sec: 0 });
    });

    it('robust gegen null-Einträge', () => {
        expect(computeTenantStats([null, { typ: 'Team' }])).toEqual({ total: 2, teams: 1, m365: 0, sec: 0 });
    });
});

describe('applyFiltersPure', () => {
    const rowsSoll = [
        { typ: 'Jahrgang', bezeichnung: '5', schuljahr: '2025/26' },
        { typ: 'Klasse', bezeichnung: '5A', schuljahr: '2025/26' },
        { typ: 'Klasse', bezeichnung: '6A', schuljahr: '2026/27' },
        { typ: 'Kursteam', bezeichnung: 'M-5A', schuljahr: '2025/26' }
    ];

    const noFilter = { schuljahr: '', typ: '', text: '', visibility: '', roster: '' };

    it('ohne Filter: alle Zeilen', () => {
        expect(applyFiltersPure(rowsSoll, 'soll', noFilter).length).toBe(4);
    });

    it('Schuljahr-Filter (nur soll/match)', () => {
        const out = applyFiltersPure(rowsSoll, 'soll', { ...noFilter, schuljahr: '2025/26' });
        expect(out.length).toBe(3);
    });

    it('Schuljahr-Filter wird im tenant-Modus ignoriert', () => {
        const out = applyFiltersPure(rowsSoll, 'tenant', { ...noFilter, schuljahr: '2025/26' });
        expect(out.length).toBe(4);
    });

    it('Typ-Filter (exakt)', () => {
        const out = applyFiltersPure(rowsSoll, 'soll', { ...noFilter, typ: 'Klasse' });
        expect(out.map((r) => r.bezeichnung)).toEqual(['5A', '6A']);
    });

    it('Volltext-Filter (case-insensitiv, bereits lowercased) – Substring-Match', () => {
        // '5a' kommt in '5A' und 'M-5A' vor – beide werden gematcht.
        const out = applyFiltersPure(rowsSoll, 'soll', { ...noFilter, text: '5a' });
        expect(out.map((r) => r.bezeichnung).sort()).toEqual(['5A', 'M-5A']);
    });

    it('Volltext sucht auch in Typ und Schuljahr', () => {
        const out = applyFiltersPure(rowsSoll, 'soll', { ...noFilter, text: 'kursteam' });
        expect(out.length).toBe(1);
        const out2 = applyFiltersPure(rowsSoll, 'soll', { ...noFilter, text: '2026/27' });
        expect(out2.length).toBe(1);
    });

    it('tenant + typ=Kursteam → HiddenMembership-Heuristik', () => {
        const rows = [
            { typ: 'Team', bezeichnung: 'KT1', hiddenMembership: true },
            { typ: 'Team', bezeichnung: 'AT', hiddenMembership: false }
        ];
        const out = applyFiltersPure(rows, 'tenant', { ...noFilter, typ: 'Kursteam' });
        expect(out.map((r) => r.bezeichnung)).toEqual(['KT1']);
    });

    it('tenant: visibility-Filter', () => {
        const rows = [
            { typ: 'Team', visibility: 'Public' },
            { typ: 'Team', visibility: 'Private' }
        ];
        const out = applyFiltersPure(rows, 'tenant', { ...noFilter, visibility: 'Private' });
        expect(out.length).toBe(1);
    });

    it('tenant: roster=noOwners → nur ownerCount===0', () => {
        const rows = [
            { typ: 'Team', ownerCount: 0, memberCount: 5 },
            { typ: 'Team', ownerCount: 2, memberCount: 5 },
            { typ: 'Team', memberCount: 5 } // ownerCount undefined → -1, NICHT 0
        ];
        const out = applyFiltersPure(rows, 'tenant', { ...noFilter, roster: 'noOwners' });
        expect(out.length).toBe(1);
    });

    it('tenant: roster=noMembers → nur memberCount===0', () => {
        const rows = [
            { typ: 'Team', ownerCount: 1, memberCount: 0 },
            { typ: 'Team', ownerCount: 1, memberCount: 3 }
        ];
        const out = applyFiltersPure(rows, 'tenant', { ...noFilter, roster: 'noMembers' });
        expect(out.length).toBe(1);
    });

    it('tenant: roster=noOwnersNoMembers → beide 0', () => {
        const rows = [
            { typ: 'Team', ownerCount: 0, memberCount: 0 },
            { typ: 'Team', ownerCount: 0, memberCount: 1 },
            { typ: 'Team', ownerCount: 1, memberCount: 0 }
        ];
        const out = applyFiltersPure(rows, 'tenant', { ...noFilter, roster: 'noOwnersNoMembers' });
        expect(out.length).toBe(1);
    });

    it('soll: visibility/roster werden ignoriert', () => {
        const out = applyFiltersPure(rowsSoll, 'soll', { ...noFilter, visibility: 'Public', roster: 'noOwners' });
        expect(out.length).toBe(4);
    });

    it('Kombi-Filter: Schuljahr + Typ + Text', () => {
        // Typ=Klasse schließt das Kursteam aus → nur '5A' bleibt.
        const out = applyFiltersPure(rowsSoll, 'soll', {
            ...noFilter,
            schuljahr: '2025/26',
            typ: 'Klasse',
            text: '5a'
        });
        expect(out.map((r) => r.bezeichnung)).toEqual(['5A']);
    });

    it('null/undefined-Zeilen werden übersprungen', () => {
        const out = applyFiltersPure([null, undefined, ...rowsSoll], 'soll', noFilter);
        expect(out.length).toBe(4);
    });

    it('robust gegen nicht-Array rows', () => {
        expect(applyFiltersPure(null, 'soll', noFilter)).toEqual([]);
    });
});
