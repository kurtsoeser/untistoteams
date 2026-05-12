import { describe, it, expect } from 'vitest';
import {
    csvEscape,
    buildKursteamCsvRow,
    buildKursteamCsv,
    buildKursteamProvisionScript,
    KURSTEAM_CSV_HEADER
} from '../src/tools/schulstruktur-sync/schulstruktur-sync-io.js';

describe('csvEscape', () => {
    it('lässt einfache Strings unverändert', () => {
        expect(csvEscape('hallo')).toBe('hallo');
        expect(csvEscape('a-b_c')).toBe('a-b_c');
    });

    it('quotet Felder mit Komma', () => {
        expect(csvEscape('a,b')).toBe('"a,b"');
    });

    it('quotet Felder mit Anführungszeichen und verdoppelt sie (RFC 4180)', () => {
        expect(csvEscape('a"b')).toBe('"a""b"');
        expect(csvEscape('"x"')).toBe('"""x"""');
    });

    it('quotet Felder mit Zeilenumbrüchen', () => {
        expect(csvEscape('a\nb')).toBe('"a\nb"');
        expect(csvEscape('a\r\nb')).toBe('"a\r\nb"');
    });

    it('konvertiert null/undefined zu Leerstring', () => {
        expect(csvEscape(null)).toBe('');
        expect(csvEscape(undefined)).toBe('');
    });

    it('konvertiert Zahlen zu Strings', () => {
        expect(csvEscape(42)).toBe('42');
        expect(csvEscape(0)).toBe('0');
    });
});

describe('buildKursteamCsvRow', () => {
    it('non-Kursteam: MailNickname via Label-Slug', () => {
        const row = { id: 'g1', typ: 'Gruppe', bezeichnung: 'Mathe Kollegium' };
        const o = buildKursteamCsvRow(row, {}, {});
        expect(o.DisplayName).toBe('Mathe Kollegium');
        expect(o.MailNickname).toMatch(/^[a-z0-9]/);
        expect(o.MailNickname.length).toBeGreaterThan(0);
    });

    it('Kursteam ohne Klasse/Fach: fällt auf Label-Slug zurück', () => {
        const row = { id: 'kt1', typ: 'Kursteam', bezeichnung: 'Fallback' };
        const o = buildKursteamCsvRow(row, {}, {}, () => ({ klasse: '', fach: '' }));
        expect(o.DisplayName).toBe('Fallback');
        expect(o.MailNickname).toMatch(/fallback/);
    });

    it('Kursteam mit Klasse+Fach: nutzt Template', () => {
        const row = { id: 'kt1', typ: 'Kursteam', bezeichnung: 'M-1A' };
        const schema = {
            kursteamYearPrefix: '2526',
            kursteamMailNickPattern: '{yearPrefix}-{klasse}-{fach}'
        };
        const o = buildKursteamCsvRow(row, {}, schema, () => ({ klasse: '1A', fach: 'M' }));
        expect(o.MailNickname).toBe('2526-1a-m');
    });

    it('Owners/Members werden über userPrincipalName/mail mit Semikolon gejoint', () => {
        const row = { id: 'kt1', typ: 'Kursteam', bezeichnung: 'X' };
        const memberships = {
            kt1: {
                owners: [{ userPrincipalName: 'a@x.at' }, { mail: 'b@x.at' }],
                members: [{ userPrincipalName: 'c@x.at' }]
            }
        };
        const o = buildKursteamCsvRow(row, memberships, {}, () => ({}));
        expect(o.Owners).toBe('a@x.at;b@x.at');
        expect(o.Members).toBe('c@x.at');
    });

    it('Visibility/Target haben Defaults', () => {
        const row = { id: 'kt1', typ: 'Kursteam', bezeichnung: 'X' };
        const o = buildKursteamCsvRow(row, {}, {}, () => ({}));
        expect(o.Visibility).toBe('HiddenMembership');
        expect(o.Target).toBe('team');
    });

    it('Visibility/Target werden aus row übernommen', () => {
        const row = {
            id: 'kt1',
            typ: 'Kursteam',
            bezeichnung: 'X',
            tenantVisibility: 'Public',
            tenantTarget: 'group'
        };
        const o = buildKursteamCsvRow(row, {}, {}, () => ({}));
        expect(o.Visibility).toBe('Public');
        expect(o.Target).toBe('group');
    });

    it('leere Bezeichnung → leerer DisplayName und MailNickname', () => {
        const row = { id: 'x', typ: 'Gruppe', bezeichnung: '' };
        const o = buildKursteamCsvRow(row, {}, {});
        expect(o.DisplayName).toBe('');
        expect(o.MailNickname).toBe('');
    });

    it('robust gegen fehlende memberships', () => {
        const row = { id: 'x', typ: 'Gruppe', bezeichnung: 'Y' };
        const o = buildKursteamCsvRow(row, null, {});
        expect(o.Owners).toBe('');
        expect(o.Members).toBe('');
    });
});

describe('buildKursteamCsv', () => {
    it('liefert immer mind. den Header (auch bei leerem Input)', () => {
        const csv = buildKursteamCsv([], {}, {});
        expect(csv).toBe(KURSTEAM_CSV_HEADER.join(','));
    });

    it('nutzt CRLF als Zeilentrenner', () => {
        const rows = [{ id: 'a', typ: 'Gruppe', bezeichnung: 'X' }];
        const csv = buildKursteamCsv(rows, {}, {});
        expect(csv.split('\r\n').length).toBe(2);
    });

    it('header ist erste Zeile in genau definierter Reihenfolge', () => {
        const csv = buildKursteamCsv([{ id: 'a', typ: 'Gruppe', bezeichnung: 'X' }], {}, {});
        const firstLine = csv.split('\r\n')[0];
        expect(firstLine).toBe('DisplayName,MailNickname,Visibility,Target,Owners,Members');
    });

    it('escapiert problematische Felder', () => {
        const rows = [{ id: 'a', typ: 'Gruppe', bezeichnung: 'X,Y' }];
        const csv = buildKursteamCsv(rows, {}, {});
        expect(csv).toContain('"X,Y"');
    });

    it('mehrere Zeilen werden korrekt zusammengesetzt', () => {
        const rows = [
            { id: 'a', typ: 'Gruppe', bezeichnung: 'A' },
            { id: 'b', typ: 'Gruppe', bezeichnung: 'B' }
        ];
        const csv = buildKursteamCsv(rows, {}, {});
        const lines = csv.split('\r\n');
        expect(lines.length).toBe(3); // header + 2 rows
    });
});

describe('buildKursteamProvisionScript', () => {
    it('verwendet den übergebenen Dateinamen', () => {
        const ps = buildKursteamProvisionScript('mycsv.csv');
        expect(ps).toContain('Import-Csv -Path "mycsv.csv"');
    });

    it('fällt auf "kursteams.csv" zurück', () => {
        const ps = buildKursteamProvisionScript();
        expect(ps).toContain('"kursteams.csv"');
    });

    it('enthält die wichtigen Teams-PS-Befehle', () => {
        const ps = buildKursteamProvisionScript('x.csv');
        expect(ps).toContain('Connect-MicrosoftTeams');
        expect(ps).toContain('New-Team');
        expect(ps).toContain('Add-TeamUser');
        expect(ps).toContain('Disconnect-MicrosoftTeams');
    });

    it('endet mit Fertigmeldung', () => {
        const ps = buildKursteamProvisionScript('x.csv');
        expect(ps).toContain('Write-Host "Fertig."');
    });

    it('iteriert über CSV mit foreach', () => {
        const ps = buildKursteamProvisionScript('x.csv');
        expect(ps).toContain('foreach ($r in $csv)');
    });
});

describe('KURSTEAM_CSV_HEADER', () => {
    it('hat genau die 6 erwarteten Spalten in definierter Reihenfolge', () => {
        expect(KURSTEAM_CSV_HEADER).toEqual([
            'DisplayName',
            'MailNickname',
            'Visibility',
            'Target',
            'Owners',
            'Members'
        ]);
    });
});
