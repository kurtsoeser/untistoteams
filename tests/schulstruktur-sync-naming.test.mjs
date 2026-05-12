import { describe, it, expect } from 'vitest';
import {
    parseSchoolYearStartYear,
    currentSchoolYearLabel,
    nextSchoolYearLabel,
    gradeFromGraduationYear,
    replaceLeadingNumber,
    normNickPart,
    normNickPrefixLower,
    maybeUpperByFlag,
    buildKursteamNameFromTemplate,
    buildKursteamMailNickFromTemplate,
    buildJgMailNick,
    buildArgeMailNick,
    buildMailNickFromLabel,
    mailNicknameFromUpn,
    generateGraphTempPassword
} from '../src/tools/schulstruktur-sync/schulstruktur-sync-naming.js';

describe('Schuljahr-Helfer', () => {
    it('parseSchoolYearStartYear: gültige Formate', () => {
        expect(parseSchoolYearStartYear('2025/26')).toBe(2025);
        expect(parseSchoolYearStartYear('2025/2026')).toBe(2025);
        expect(parseSchoolYearStartYear('  2030 / 31  ')).toBe(2030);
    });

    it('parseSchoolYearStartYear: ungültige Eingabe → NaN', () => {
        expect(parseSchoolYearStartYear('foo')).toBeNaN();
        expect(parseSchoolYearStartYear('')).toBeNaN();
        expect(parseSchoolYearStartYear(null)).toBeNaN();
    });

    it('currentSchoolYearLabel: Format YYYY/YY', () => {
        const lbl = currentSchoolYearLabel();
        expect(lbl).toMatch(/^\d{4}\/\d{2}$/);
        const cy = new Date().getFullYear();
        expect(lbl.startsWith(String(cy))).toBe(true);
    });

    it('nextSchoolYearLabel: aus gültigem Vorgänger', () => {
        expect(nextSchoolYearLabel('2025/26')).toBe('2026/27');
        expect(nextSchoolYearLabel('2099/2100')).toBe('2100/01');
    });

    it('nextSchoolYearLabel: ungültige Eingabe → currentSchoolYearLabel', () => {
        expect(nextSchoolYearLabel('foo')).toBe(currentSchoolYearLabel());
    });

    it('gradeFromGraduationYear: Standard 5 Stufen', () => {
        // Abschlussjahr 2028, Schuljahr 2025/26 → Diff 3, Stufe = 6 - 3 = 3
        expect(gradeFromGraduationYear(2028, '2025/26')).toBe(3);
        // Abschluss == Schuljahranfang → letztes Jahr (Stufe = 6 - 0 = 6, geclampt nicht nötig hier)
        expect(gradeFromGraduationYear(2025, '2025/26')).toBe(6);
    });

    it('gradeFromGraduationYear: maxStufen-Override', () => {
        // 8 Stufen: 2028 - 2025 = 3 → (8+1)-3 = 6
        expect(gradeFromGraduationYear(2028, '2025/26', 8)).toBe(6);
    });

    it('gradeFromGraduationYear: ungültige Eingabe → NaN', () => {
        expect(gradeFromGraduationYear('abcd', '2025/26')).toBeNaN();
        expect(gradeFromGraduationYear(2028, 'foo')).toBeNaN();
    });

    it('replaceLeadingNumber: ersetzt führende Zahl', () => {
        expect(replaceLeadingNumber('1A', 2)).toBe('2A');
        expect(replaceLeadingNumber('12B', 13)).toBe('13B');
        expect(replaceLeadingNumber('A1', 2)).toBe('A1'); // keine Zahl vorne
        expect(replaceLeadingNumber('', 5)).toBe('');
    });
});

describe('Nick-Normalisierung', () => {
    it('normNickPart: nur [A-Za-z0-9-]', () => {
        expect(normNickPart(' Hallo Welt 1A! ')).toBe('HalloWelt1A');
        expect(normNickPart('a-b_c')).toBe('a-bc');
        expect(normNickPart(null)).toBe('');
    });

    it('normNickPrefixLower: lower + nur [a-z0-9]', () => {
        expect(normNickPrefixLower('JG-Gruppe', 'jg')).toBe('jggruppe');
        expect(normNickPrefixLower('', 'jg')).toBe('jg');
        expect(normNickPrefixLower(null, 'fallback')).toBe('fallback');
        expect(normNickPrefixLower('', '')).toBe('');
    });

    it('maybeUpperByFlag', () => {
        expect(maybeUpperByFlag('aBcD', true)).toBe('ABCD');
        expect(maybeUpperByFlag('aBcD', false)).toBe('abcd');
    });
});

describe('Builder', () => {
    it('buildKursteamNameFromTemplate: Default-Template', () => {
        const r = buildKursteamNameFromTemplate('', { yearPrefix: '25', klasse: '3A', fach: 'M' });
        expect(r).toBe('25 | 3A | M');
    });

    it('buildKursteamNameFromTemplate: custom + Gruppe', () => {
        const r = buildKursteamNameFromTemplate('{klasse}-{fach}-{gruppe}', {
            klasse: '3A', fach: 'M', gruppe: 'L'
        });
        expect(r).toBe('3A-M-L');
    });

    it('buildKursteamMailNickFromTemplate: slugifiziert Ergebnis', () => {
        const r = buildKursteamMailNickFromTemplate('', { yearPrefix: '25', klasse: '3A', fach: 'M' });
        expect(r).toBe('kt-25-3a-m');
    });

    it('buildKursteamMailNickFromTemplate: kollabiert Bindestriche', () => {
        const r = buildKursteamMailNickFromTemplate('{klasse}--{fach}', { klasse: '3A', fach: '' });
        expect(r).toBe('3a');
    });

    it('buildJgMailNick: Standard-Pattern', () => {
        const r = buildJgMailNick({ jgPrefix: 'jg', jgUpper: false }, '2025', 'A');
        expect(r).toBe('jg2025-a');
    });

    it('buildJgMailNick: jgUpper schaltet Suffix auf UPPER', () => {
        const r = buildJgMailNick({ jgPrefix: 'jg', jgUpper: true }, '2025', 'a');
        expect(r).toBe('jg2025-A');
    });

    it('buildJgMailNick: leerer Suffix → kein Separator', () => {
        const r = buildJgMailNick({ jgPrefix: 'jg' }, '2025', '');
        expect(r).toBe('jg2025');
    });

    it('buildArgeMailNick: Standard', () => {
        const r = buildArgeMailNick({ argePrefix: 'arge', argeUpper: true }, 'fuss');
        expect(r).toBe('arge-FUSS');
    });

    it('buildArgeMailNick: ohne Code', () => {
        expect(buildArgeMailNick({ argePrefix: 'ag' }, '')).toBe('ag');
    });

    it('buildMailNickFromLabel: Slug aus Bezeichner', () => {
        expect(buildMailNickFromLabel('  Hallo Welt!  ')).toBe('hallo-welt');
        expect(buildMailNickFromLabel('a---b')).toBe('a-b');
        expect(buildMailNickFromLabel('---a---')).toBe('a');
        expect(buildMailNickFromLabel('')).toBe('');
    });

    it('mailNicknameFromUpn: nutzt local-part', () => {
        expect(mailNicknameFromUpn('max.muster@schule.at')).toBe('maxmuster');
        expect(mailNicknameFromUpn('Anna_BAUER@schule.at')).toBe('annabauer');
    });

    it('mailNicknameFromUpn: leere Eingabe → Fallback u<hex>', () => {
        const r = mailNicknameFromUpn('');
        expect(r).toMatch(/^u[0-9a-f]{8}$/);
    });

    it('mailNicknameFromUpn: cap bei 64 Zeichen', () => {
        const long = 'a'.repeat(100) + '@x.at';
        expect(mailNicknameFromUpn(long).length).toBeLessThanOrEqual(64);
    });
});

describe('generateGraphTempPassword', () => {
    it('hat erwartete Länge (14 + 1 + 3 = 18)', () => {
        expect(generateGraphTempPassword().length).toBe(18);
    });

    it('endet immer auf "1aA" (Komplexitätsgarantie)', () => {
        for (let i = 0; i < 10; i++) {
            expect(generateGraphTempPassword().endsWith('1aA')).toBe(true);
        }
    });

    it('verwendet ausschließlich erlaubte Zeichen', () => {
        // Erlaubt: A-Z ohne I/O, a-z ohne l, Ziffern 2-9.
        const allowed = /^[A-HJ-NP-Za-km-z2-9]{14}[!@#$%]1aA$/;
        for (let i = 0; i < 20; i++) {
            expect(generateGraphTempPassword()).toMatch(allowed);
        }
    });
});
