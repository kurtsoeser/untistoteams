import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { computeGraphModel } from '../src/tools/schulstruktur-sync/schulstruktur-sync-graph-layout.js';
import {
    STRUCT_FOLDER_ARGES,
    STRUCT_FOLDER_FACHSCHAFTEN
} from '../src/tools/schulstruktur-sync/schulstruktur-sync-tree.js';

describe('computeGraphModel', () => {
    /**
     * Mockt `window.ms365StructureRules.inferRootForType` deterministisch
     * (identisch zur Logik in `structure-rules.js`).
     */
    beforeEach(() => {
        globalThis.window = {
            ms365StructureRules: {
                inferRootForType: (typ) => {
                    if (typ === 'Jahrgang' || typ === 'Klasse' || typ === 'Kursteam') return 'SchuelerInnen';
                    if (typ === 'Arbeitsgemeinschaft' || typ === 'Gruppe' || typ === 'Person') return 'LehrerInnen';
                    return '';
                }
            }
        };
    });

    afterEach(() => {
        delete globalThis.window;
    });

    describe('Grundgerüst', () => {
        it('liefert immer die 3 Root-Nodes + 2 synthetische Ordner', () => {
            const m = computeGraphModel([]);
            expect(m.nodes.has('__root_students__')).toBe(true);
            expect(m.nodes.has('__root_teachers__')).toBe(true);
            expect(m.nodes.has('__root_admin__')).toBe(true);
            expect(m.nodes.has(STRUCT_FOLDER_ARGES)).toBe(true);
            expect(m.nodes.has(STRUCT_FOLDER_FACHSCHAFTEN)).toBe(true);
        });

        it('Root-Nodes haben isRoot=true', () => {
            const m = computeGraphModel([]);
            expect(m.nodes.get('__root_students__').isRoot).toBe(true);
            expect(m.nodes.get('__root_teachers__').isRoot).toBe(true);
            expect(m.nodes.get('__root_admin__').isRoot).toBe(true);
        });

        it('synthetische Ordner haben isStructureFolder=true', () => {
            const m = computeGraphModel([]);
            expect(m.nodes.get(STRUCT_FOLDER_ARGES).isStructureFolder).toBe(true);
            expect(m.nodes.get(STRUCT_FOLDER_FACHSCHAFTEN).isStructureFolder).toBe(true);
        });

        it('Custom-Titel aus structRootDetails wird übernommen', () => {
            const m = computeGraphModel([], null, {
                __root_students__: { bezeichnung: 'Schülis' }
            });
            expect(m.nodes.get('__root_students__').bezeichnung).toBe('Schülis');
        });
    });

    describe('Edges & Routing', () => {
        it('Jahrgang ohne Parent → unter Schüler:innen-Root', () => {
            const m = computeGraphModel([{ id: 'j1', typ: 'Jahrgang', bezeichnung: '5' }]);
            const e = m.edges.find((e) => e.to === 'j1');
            expect(e.from).toBe('__root_students__');
        });

        it('ARGE → unter ARGE-Ordner', () => {
            const m = computeGraphModel([
                { id: 'a1', typ: 'Arbeitsgemeinschaft', bezeichnung: 'Fußball' }
            ]);
            const e = m.edges.find((e) => e.to === 'a1');
            expect(e.from).toBe(STRUCT_FOLDER_ARGES);
        });

        it('Fachschaft-Gruppe (fachschaftFach=true) → unter Fachschaften-Ordner', () => {
            const m = computeGraphModel([
                { id: 'f1', typ: 'Gruppe', bezeichnung: 'Mathematik', fachschaftFach: true }
            ]);
            const e = m.edges.find((e) => e.to === 'f1');
            expect(e.from).toBe(STRUCT_FOLDER_FACHSCHAFTEN);
        });

        it('Top-„Verwaltung" → unter Admin-Root', () => {
            const m = computeGraphModel([
                { id: 'v1', typ: 'Gruppe', bezeichnung: 'Verwaltung' }
            ]);
            const e = m.edges.find((e) => e.to === 'v1');
            expect(e.from).toBe('__root_admin__');
        });

        it('Kind mit gültiger parentId behält den Parent', () => {
            const m = computeGraphModel([
                { id: 'k1', typ: 'Klasse', bezeichnung: '1A', parentId: 'j1' },
                { id: 'j1', typ: 'Jahrgang', bezeichnung: '5' }
            ]);
            const e = m.edges.find((e) => e.to === 'k1');
            expect(e.from).toBe('j1');
        });

        it('Lehrer-Wurzel: ARGE-Ordner immer als erstes Kind, dann Fachschaften, dann Rest', () => {
            const m = computeGraphModel([
                { id: 'g1', typ: 'Gruppe', bezeichnung: 'Z-Gruppe' },
                { id: 'a1', typ: 'Arbeitsgemeinschaft', bezeichnung: 'AG' }
            ]);
            const teacherKids = m.children.get('__root_teachers__');
            expect(teacherKids[0]).toBe(STRUCT_FOLDER_ARGES);
            expect(teacherKids[1]).toBe(STRUCT_FOLDER_FACHSCHAFTEN);
            expect(teacherKids).toContain('g1');
        });
    });

    describe('Positionen / Canvas', () => {
        it('horizontal: y nimmt mit Tiefe zu, x ist verschieden für Geschwister', () => {
            const m = computeGraphModel([
                { id: 'j1', typ: 'Jahrgang', bezeichnung: '5' },
                { id: 'k1', typ: 'Klasse', bezeichnung: '1A', parentId: 'j1' },
                { id: 'k2', typ: 'Klasse', bezeichnung: '1B', parentId: 'j1' }
            ]);
            const pj = m.pos.get('j1');
            const pk1 = m.pos.get('k1');
            const pk2 = m.pos.get('k2');
            expect(pk1.y).toBeGreaterThan(pj.y);
            expect(pk2.y).toBeGreaterThan(pj.y);
            expect(pk1.x).not.toBe(pk2.x);
            expect(m.graphLayout).toBe('horizontal');
        });

        it('vertical: x nimmt mit Tiefe zu, y ist verschieden für Geschwister', () => {
            const m = computeGraphModel(
                [
                    { id: 'j1', typ: 'Jahrgang', bezeichnung: '5' },
                    { id: 'k1', typ: 'Klasse', bezeichnung: '1A', parentId: 'j1' },
                    { id: 'k2', typ: 'Klasse', bezeichnung: '1B', parentId: 'j1' }
                ],
                undefined,
                undefined,
                'vertical'
            );
            const pj = m.pos.get('j1');
            const pk1 = m.pos.get('k1');
            const pk2 = m.pos.get('k2');
            expect(pk1.x).toBeGreaterThan(pj.x);
            expect(pk2.x).toBeGreaterThan(pj.x);
            expect(pk1.y).not.toBe(pk2.y);
            expect(m.graphLayout).toBe('vertical');
        });

        it('canvas erfüllt Mindestmaße 1200×650', () => {
            const m = computeGraphModel([]);
            expect(m.canvas.width).toBeGreaterThanOrEqual(1200);
            expect(m.canvas.height).toBeGreaterThanOrEqual(650);
        });

        it('canvas wächst mit größerem Baum', () => {
            const small = computeGraphModel([]);
            const wide = computeGraphModel(
                Array.from({ length: 30 }, (_, i) => ({
                    id: 'j' + i,
                    typ: 'Jahrgang',
                    bezeichnung: String(i)
                }))
            );
            expect(wide.canvas.width).toBeGreaterThanOrEqual(small.canvas.width);
        });

        it('collapsedSet auf Root: Kinder bekommen KEINE pos-Einträge', () => {
            const m = computeGraphModel(
                [
                    { id: 'j1', typ: 'Jahrgang', bezeichnung: '5' },
                    { id: 'k1', typ: 'Klasse', bezeichnung: '1A', parentId: 'j1' }
                ],
                new Set(['__root_students__'])
            );
            expect(m.pos.has('__root_students__')).toBe(true);
            // j1 selbst sollte nicht positioniert sein, weil sein Parent kollabiert ist
            expect(m.pos.has('j1')).toBe(false);
        });
    });

    describe('Determinismus', () => {
        it('zwei Aufrufe mit gleichen Inputs liefern identische Strukturen', () => {
            const rows = [
                { id: 'j1', typ: 'Jahrgang', bezeichnung: '5' },
                { id: 'k1', typ: 'Klasse', bezeichnung: '1A', parentId: 'j1' }
            ];
            const a = computeGraphModel(rows);
            const b = computeGraphModel(rows);
            expect(a.nodes.size).toBe(b.nodes.size);
            expect(a.edges.length).toBe(b.edges.length);
            expect(a.pos.get('j1')).toEqual(b.pos.get('j1'));
            expect(a.canvas).toEqual(b.canvas);
        });

        it('Geschwister sind nach Bezeichnung sortiert (case-insensitive)', () => {
            const m = computeGraphModel([
                { id: 'a', typ: 'Jahrgang', bezeichnung: 'Beta' },
                { id: 'b', typ: 'Jahrgang', bezeichnung: 'alpha' },
                { id: 'c', typ: 'Jahrgang', bezeichnung: 'Gamma' }
            ]);
            const kids = m.children.get('__root_students__');
            // Erwartete Reihenfolge (deutsche Sortierung, case-insensitive)
            expect(kids).toEqual(['b', 'a', 'c']);
        });
    });
});
