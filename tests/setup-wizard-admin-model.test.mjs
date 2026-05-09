import { describe, it, expect } from 'vitest';
import {
    SW_ADMIN_DEFAULT_ROLES,
    getAdminDisplayRowsFromSettings,
    migrateAdminRowDefaultKey,
    normCode,
    normStr,
    randomTempPassword,
    resolveAdminSlotFromRow
} from '../src/shared/setup-wizard-admin-model.js';

describe('setup-wizard-admin-model', () => {
    it('getAdminDisplayRowsFromSettings: leere Admin-Liste → Standardrollen', () => {
        const rows = getAdminDisplayRowsFromSettings({});
        expect(rows.length).toBe(SW_ADMIN_DEFAULT_ROLES.length);
        expect(rows[0]).toEqual({ defaultKey: 'Direktion', role: 'Direktion', name: '', email: '' });
    });

    it('migrateAdminRowDefaultKey setzt defaultKey aus Rolle', () => {
        const m = migrateAdminRowDefaultKey({ role: 'bibliothek', name: 'x', email: '' });
        expect(m.defaultKey).toBe('Bibliothek');
    });

    it('resolveAdminSlotFromRow und normCode', () => {
        expect(resolveAdminSlotFromRow({ defaultKey: 'sekretariat' })).toBe('Sekretariat');
        expect(normCode('  ab12  ')).toBe('AB12');
        expect(normStr('  x  ')).toBe('x');
    });

    it('randomTempPassword: Länge und Zeichenklassen', () => {
        const pwd = randomTempPassword();
        expect(pwd.length).toBeGreaterThanOrEqual(16);
        expect(/[A-Z]/.test(pwd)).toBe(true);
        expect(/[a-z]/.test(pwd)).toBe(true);
        expect(/[0-9]/.test(pwd)).toBe(true);
    });
});
