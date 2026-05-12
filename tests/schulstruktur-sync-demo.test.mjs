import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import {
    getTenantSettingsDomainFallback,
    defaultAnlegenSchemas,
    normalizeGraphLayoutModeInSettings,
    normalizeClassLabel,
    deriveGradeFromClassLabel,
    buildDemoFromTenantSettings,
    buildDemoRows
} from '../src/tools/schulstruktur-sync/schulstruktur-sync-demo.js';

/**
 * Setzt `window.ms365TenantSettingsLoad` deterministisch auf `value`.
 * Wird per Test wieder entfernt.
 */
function withTenantSettings(value) {
    globalThis.window = {
        ms365TenantSettingsLoad: typeof value === 'function' ? value : () => value
    };
}

afterEach(() => {
    delete globalThis.window;
});

describe('normalizeClassLabel', () => {
    it('bevorzugt code vor name', () => {
        expect(normalizeClassLabel({ code: '1A', name: 'Klasse 1A' })).toBe('1A');
    });

    it('nutzt name als fallback', () => {
        expect(normalizeClassLabel({ name: 'Klasse 1A' })).toBe('Klasse 1A');
    });

    it('trimmt Whitespace', () => {
        expect(normalizeClassLabel({ code: '  1B  ' })).toBe('1B');
    });

    it('liefert Leerstring bei null/undefined/leer', () => {
        expect(normalizeClassLabel(null)).toBe('');
        expect(normalizeClassLabel(undefined)).toBe('');
        expect(normalizeClassLabel({})).toBe('');
        expect(normalizeClassLabel({ code: '', name: '' })).toBe('');
    });
});

describe('deriveGradeFromClassLabel', () => {
    it('extrahiert führende 1-stellige Ziffer', () => {
        expect(deriveGradeFromClassLabel('1A')).toBe('1');
        expect(deriveGradeFromClassLabel('5b')).toBe('5');
    });

    it('extrahiert führende 2-stellige Ziffer', () => {
        expect(deriveGradeFromClassLabel('10A')).toBe('10');
        expect(deriveGradeFromClassLabel('12-Chemie')).toBe('12');
    });

    it('liefert Leerstring ohne führende Ziffer', () => {
        expect(deriveGradeFromClassLabel('AbendA')).toBe('');
        expect(deriveGradeFromClassLabel('')).toBe('');
        expect(deriveGradeFromClassLabel(null)).toBe('');
    });

    it('trimmt Whitespace vor dem Match', () => {
        expect(deriveGradeFromClassLabel('  3A')).toBe('3');
    });
});

describe('normalizeGraphLayoutModeInSettings', () => {
    it('macht null/undefined zu no-op', () => {
        expect(() => normalizeGraphLayoutModeInSettings(null)).not.toThrow();
        expect(() => normalizeGraphLayoutModeInSettings(undefined)).not.toThrow();
    });

    it('lässt "vertical" stehen', () => {
        const s = { graphLayoutMode: 'vertical' };
        normalizeGraphLayoutModeInSettings(s);
        expect(s.graphLayoutMode).toBe('vertical');
    });

    it('mappt unbekannte Werte auf "horizontal"', () => {
        const s = { graphLayoutMode: 'diagonal' };
        normalizeGraphLayoutModeInSettings(s);
        expect(s.graphLayoutMode).toBe('horizontal');
    });

    it('setzt missing → "horizontal"', () => {
        const s = {};
        normalizeGraphLayoutModeInSettings(s);
        expect(s.graphLayoutMode).toBe('horizontal');
    });

    it('mutiert das Argument in-place', () => {
        const s = { graphLayoutMode: 'foo' };
        const ret = normalizeGraphLayoutModeInSettings(s);
        expect(ret).toBeUndefined();
        expect(s.graphLayoutMode).toBe('horizontal');
    });
});

describe('getTenantSettingsDomainFallback', () => {
    it('liefert Leerstring ohne Loader', () => {
        expect(getTenantSettingsDomainFallback()).toBe('');
    });

    it('liefert die Domain aus den Settings', () => {
        withTenantSettings({ domain: 'hs1.schule.at' });
        expect(getTenantSettingsDomainFallback()).toBe('hs1.schule.at');
    });

    it('trimmt die Domain', () => {
        withTenantSettings({ domain: '  hs2.schule.at  ' });
        expect(getTenantSettingsDomainFallback()).toBe('hs2.schule.at');
    });

    it('schluckt Loader-Fehler robust', () => {
        withTenantSettings(() => {
            throw new Error('boom');
        });
        expect(getTenantSettingsDomainFallback()).toBe('');
    });
});

describe('defaultAnlegenSchemas', () => {
    it('liefert alle erwarteten Felder mit sinnvollen Defaults', () => {
        const s = defaultAnlegenSchemas();
        expect(s.domain).toBe('ms365.schule'); // Fallback ohne Tenant
        expect(s.kursteamYearPrefix).toMatch(/^SJ\d{2}$/);
        expect(s.kursteamPattern).toBe('{yearPrefix} | {klasse} | {fach}');
        expect(s.kursteamMailNickPattern).toBe('kt-{yearPrefix}-{klasse}-{fach}');
        expect(s.jgPrefix).toBe('jg');
        expect(s.jgUpper).toBe(true);
        expect(s.argePrefix).toBe('arge');
        expect(s.argeUpper).toBe(false);
        expect(s.maxSchulstufen).toBe(5);
        expect(s.graphLayoutMode).toBe('horizontal');
    });

    it('übernimmt die Domain aus Tenant-Settings, wenn vorhanden', () => {
        withTenantSettings({ domain: 'meineschule.at' });
        const s = defaultAnlegenSchemas();
        expect(s.domain).toBe('meineschule.at');
    });

    it('Year-Prefix folgt dem aktuellen Jahr (letzte 2 Stellen)', () => {
        const s = defaultAnlegenSchemas();
        const yy = String(new Date().getFullYear()).slice(-2);
        expect(s.kursteamYearPrefix).toBe('SJ' + yy);
    });
});

describe('buildDemoFromTenantSettings', () => {
    it('liefert null ohne Loader', () => {
        expect(buildDemoFromTenantSettings()).toBeNull();
    });

    it('liefert null ohne Klassen', () => {
        withTenantSettings({});
        expect(buildDemoFromTenantSettings()).toBeNull();
    });

    it('liefert null bei classes=[]', () => {
        withTenantSettings({ classes: [] });
        expect(buildDemoFromTenantSettings()).toBeNull();
    });

    it('baut Jahrgang+Klasse+Standard-Gruppen aus Klassen-Liste', () => {
        withTenantSettings({
            classes: [{ code: '1A' }, { code: '1B' }, { code: '2A' }]
        });
        const out = buildDemoFromTenantSettings();
        expect(out).not.toBeNull();
        expect(out.tenantSettings).toEqual({ classes: [{ code: '1A' }, { code: '1B' }, { code: '2A' }] });
        const types = out.rows.map((r) => r.typ);
        expect(types.filter((t) => t === 'Jahrgang').length).toBe(2); // Jg 1 und 2
        expect(types.filter((t) => t === 'Klasse').length).toBe(3);
        expect(types.filter((t) => t === 'Gruppe').length).toBe(2);
    });

    it('Klassen werden ihrem Jahrgang als parent zugeordnet', () => {
        withTenantSettings({ classes: [{ code: '1A' }, { code: '1B' }] });
        const out = buildDemoFromTenantSettings();
        const jg = out.rows.find((r) => r.typ === 'Jahrgang' && r.bezeichnung === 'Jahrgang 1');
        const k1a = out.rows.find((r) => r.bezeichnung === '1A');
        expect(k1a.parentId).toBe(jg.id);
    });

    it('Standard-Gruppen enthalten "Lehrer:innen" und "Schüler:innen"', () => {
        withTenantSettings({ classes: [{ code: '1A' }] });
        const out = buildDemoFromTenantSettings();
        const gruppen = out.rows.filter((r) => r.typ === 'Gruppe').map((r) => r.bezeichnung);
        expect(gruppen).toContain('Lehrer:innen');
        expect(gruppen).toContain('Schüler:innen');
    });

    it('alle Zeilen haben schuljahr=currentSchoolYearLabel und Status-Defaults', () => {
        withTenantSettings({ classes: [{ code: '1A' }] });
        const out = buildDemoFromTenantSettings();
        out.rows.forEach((r) => {
            expect(r.schuljahr).toMatch(/^\d{4}\/\d{2}$/);
            expect(r.status).toBe('Aktiv');
            expect(['Ausstehend']).toContain(r.syncStatus);
            expect(r.letzteFehlermeldung).toBe('');
        });
    });

    it('Klassen ohne Label werden übersprungen', () => {
        withTenantSettings({ classes: [{ code: '1A' }, { code: '' }, {}, { code: '2B' }] });
        const out = buildDemoFromTenantSettings();
        const klassen = out.rows.filter((r) => r.typ === 'Klasse').map((r) => r.bezeichnung);
        expect(klassen).toEqual(['1A', '2B']);
    });
});

describe('buildDemoRows', () => {
    it('fällt auf statischen Mock zurück (ohne Tenant)', () => {
        const out = buildDemoRows();
        expect(out.tenantSettings).toBeNull();
        expect(out.rows.length).toBeGreaterThan(0);

        const bezeichnungen = out.rows.map((r) => r.bezeichnung);
        expect(bezeichnungen).toContain('Jahrgang 1');
        expect(bezeichnungen).toContain('Jahrgang 2');
        expect(bezeichnungen).toContain('1A');
        expect(bezeichnungen).toContain('ARGE Robotik');
    });

    it('bevorzugt Tenant-Settings wenn nicht leer', () => {
        withTenantSettings({ classes: [{ code: '3A' }] });
        const out = buildDemoRows();
        expect(out.tenantSettings).not.toBeNull();
        const bezeichnungen = out.rows.map((r) => r.bezeichnung);
        expect(bezeichnungen).toContain('3A');
        expect(bezeichnungen).toContain('Jahrgang 3');
    });

    it('Mock-Daten haben unterschiedliche Sync-Status (Ok / Abweichung / Fehler)', () => {
        const out = buildDemoRows();
        const stati = new Set(out.rows.map((r) => r.syncStatus));
        expect(stati.size).toBeGreaterThan(1);
    });
});
