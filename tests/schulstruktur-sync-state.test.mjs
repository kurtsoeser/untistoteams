import { describe, it, expect, beforeEach } from 'vitest';
import {
    loadState,
    saveState,
    loadMatchState,
    saveMatchState,
    loadTenantCache,
    saveTenantCache,
    loadGraphCollapsedSet,
    saveGraphCollapsedSet
} from '../src/tools/schulstruktur-sync/schulstruktur-sync-state.js';

/**
 * Minimaler `localStorage`-Mock auf `globalThis.window`. Das Modul greift
 * direkt auf `localStorage` als Browser-Global zu, daher exponieren wir
 * die Mock-Implementation unter beiden Namen.
 */
function setupMockEnv() {
    const mem = new Map();
    const store = {
        getItem: (k) => (mem.has(k) ? mem.get(k) : null),
        setItem: (k, v) => mem.set(k, String(v)),
        removeItem: (k) => mem.delete(k),
        clear: () => mem.clear()
    };
    globalThis.localStorage = store;
    globalThis.window = { localStorage: store };
    return { mem, store };
}

describe('schulstruktur-sync-state.js', () => {
    beforeEach(() => {
        setupMockEnv();
    });

    describe('loadState / saveState', () => {
        it('liefert leeren State, wenn nichts gespeichert ist', () => {
            const s = loadState();
            expect(s.rows).toEqual([]);
            expect(s.memberships).toEqual({});
            expect(s.settings).toEqual({});
        });

        it('Roundtrip: gespeicherter State kommt unverändert zurück', () => {
            saveState({
                rows: [{ id: '1', typ: 'Klasse', bezeichnung: '1A' }],
                memberships: { '1': ['user@x.at'] },
                settings: { foo: 'bar' }
            });
            const s = loadState();
            expect(s.rows).toHaveLength(1);
            expect(s.rows[0].bezeichnung).toBe('1A');
            expect(s.memberships['1']).toEqual(['user@x.at']);
            expect(s.settings.foo).toBe('bar');
        });

        it('saveState toleriert ungültige Eingaben', () => {
            const r = saveState(null);
            expect(r.rows).toEqual([]);
            expect(r.memberships).toEqual({});
        });

        it('loadState bevorzugt ms365AppDataV2-Container vor localStorage', () => {
            globalThis.window.ms365AppDataV2 = {
                getContainer: () => ({
                    structure: {
                        rows: [{ id: 'v2' }],
                        memberships: { v2: [] },
                        settings: { src: 'v2' }
                    }
                })
            };
            saveState({ rows: [{ id: 'ls' }] });
            const s = loadState();
            expect(s.rows[0].id).toBe('v2');
            expect(s.settings.src).toBe('v2');
        });
    });

    describe('loadMatchState / saveMatchState', () => {
        it('liefert leere Links als Default', () => {
            expect(loadMatchState()).toEqual({ links: {} });
        });

        it('Roundtrip Match-Links', () => {
            saveMatchState({ 'struct-1': { tenantGroupId: 'g1', note: 'n' } });
            const m = loadMatchState();
            expect(m.links['struct-1'].tenantGroupId).toBe('g1');
        });
    });

    describe('loadTenantCache / saveTenantCache', () => {
        it('liefert leeren Cache als Default', () => {
            const c = loadTenantCache();
            expect(c.rows).toEqual([]);
            expect(c.users).toEqual([]);
            expect(c.loadedAt).toBe('');
        });

        it('Roundtrip Tenant-Cache (rows + users)', () => {
            saveTenantCache(
                [{ id: 'g1', displayName: 'Lehrer:innen' }],
                [{ id: 'u1', mail: 'a@b.at' }]
            );
            const c = loadTenantCache();
            expect(c.rows[0].displayName).toBe('Lehrer:innen');
            expect(c.users[0].mail).toBe('a@b.at');
            expect(c.loadedAt).toMatch(/^\d{4}-\d{2}-\d{2}T/);
        });

        it('saveTenantCache ohne users-Argument behält bestehende User', () => {
            saveTenantCache([{ id: 'g1' }], [{ id: 'u1' }]);
            saveTenantCache([{ id: 'g2' }]); // users weggelassen
            const c = loadTenantCache();
            expect(c.rows[0].id).toBe('g2');
            expect(c.users[0].id).toBe('u1');
        });
    });

    describe('loadGraphCollapsedSet / saveGraphCollapsedSet', () => {
        it('liefert leeres Set als Default', () => {
            const s = loadGraphCollapsedSet();
            expect(s).toBeInstanceOf(Set);
            expect(s.size).toBe(0);
        });

        it('Roundtrip Collapsed-IDs', () => {
            const inp = new Set(['a', 'b', 'c']);
            saveGraphCollapsedSet(inp);
            const out = loadGraphCollapsedSet();
            expect(out.size).toBe(3);
            expect(out.has('a')).toBe(true);
            expect(out.has('b')).toBe(true);
            expect(out.has('c')).toBe(true);
        });

        it('saveGraphCollapsedSet toleriert null/undefined', () => {
            expect(() => saveGraphCollapsedSet(null)).not.toThrow();
            expect(() => saveGraphCollapsedSet(undefined)).not.toThrow();
        });
    });
});
