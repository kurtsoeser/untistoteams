import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import {
    STRUCT_TREE_ROOT_STUDENTS,
    STRUCT_TREE_ROOT_TEACHERS,
    STRUCT_TREE_ROOT_ADMIN,
    STRUCT_FOLDER_ARGES,
    STRUCT_FOLDER_FACHSCHAFTEN,
    byId,
    uid,
    isStructureSyntheticGraphNodeId,
    isStructureTreeRootId,
    isTopVerwaltungStructureNode,
    teacherFachVirtualId,
    structureTreeRootDefaultTitle,
    structureTreeRootTitle,
    defaultStructureTreeRootRow,
    mergeStructureTreeRootRow,
    pickStorableStructureTreeRootFields,
    migrateLegacyFachschaftenContainer,
    ensureFachschaftFachGruppen,
    sortStructureTreeChildren,
    buildTreeOrder,
    buildStructuredTreeOrder,
    typeIcon,
    treeIconForRow,
    treeIconForVirtualItem,
    graphNodeIconClass
} from '../src/tools/schulstruktur-sync/schulstruktur-sync-tree.js';

describe('Konstanten + kleine Helfer', () => {
    it('Konstanten haben erwartete Werte', () => {
        expect(STRUCT_TREE_ROOT_STUDENTS).toBe('__root_students__');
        expect(STRUCT_TREE_ROOT_TEACHERS).toBe('__root_teachers__');
        expect(STRUCT_TREE_ROOT_ADMIN).toBe('__root_admin__');
        expect(STRUCT_FOLDER_ARGES).toBe('__ss_folder_arges__');
        expect(STRUCT_FOLDER_FACHSCHAFTEN).toBe('__ss_folder_fachschaften__');
    });

    it('byId indexiert Rows', () => {
        const m = byId([{ id: 'a', x: 1 }, { id: 'b', x: 2 }, null, { id: '', x: 3 }]);
        expect(m.size).toBe(2);
        expect(m.get('a').x).toBe(1);
        expect(m.get('b').x).toBe(2);
    });

    it('uid: eindeutig und prefix', () => {
        const a = uid();
        const b = uid();
        expect(a).toMatch(/^id-/);
        expect(a).not.toBe(b);
    });

    it('teacherFachVirtualId: deterministisch', () => {
        expect(teacherFachVirtualId('Mathematik')).toBe('__ss_fach__Mathematik');
        expect(teacherFachVirtualId('')).toBe('__ss_fach__(ohne%20Fach)'.replace(/%/g, '_'));
        // URL-encoded „&" wird zu „_26"
        expect(teacherFachVirtualId('A&B')).toBe('__ss_fach__A_26B');
    });
});

describe('Klassifikationen', () => {
    it('isStructureSyntheticGraphNodeId', () => {
        expect(isStructureSyntheticGraphNodeId(STRUCT_TREE_ROOT_STUDENTS)).toBe(true);
        expect(isStructureSyntheticGraphNodeId(STRUCT_FOLDER_ARGES)).toBe(true);
        expect(isStructureSyntheticGraphNodeId('__ss_fach__M')).toBe(true);
        expect(isStructureSyntheticGraphNodeId('__ss_kursteams__M')).toBe(true);
        expect(isStructureSyntheticGraphNodeId('echte-id')).toBe(false);
        expect(isStructureSyntheticGraphNodeId('')).toBe(false);
    });

    it('isStructureTreeRootId', () => {
        expect(isStructureTreeRootId(STRUCT_TREE_ROOT_STUDENTS)).toBe(true);
        expect(isStructureTreeRootId(STRUCT_TREE_ROOT_TEACHERS)).toBe(true);
        expect(isStructureTreeRootId(STRUCT_TREE_ROOT_ADMIN)).toBe(true);
        expect(isStructureTreeRootId(STRUCT_FOLDER_ARGES)).toBe(false);
        expect(isStructureTreeRootId('x')).toBe(false);
    });

    it('isTopVerwaltungStructureNode', () => {
        expect(isTopVerwaltungStructureNode({ bezeichnung: 'Verwaltung', parentId: '' })).toBe(true);
        expect(isTopVerwaltungStructureNode({ bezeichnung: 'verwaltung' })).toBe(true);
        expect(isTopVerwaltungStructureNode({ bezeichnung: 'Verwaltung', parentId: 'p' })).toBe(false);
        expect(isTopVerwaltungStructureNode({ bezeichnung: 'X' })).toBe(false);
        expect(isTopVerwaltungStructureNode(null)).toBe(false);
    });
});

describe('Title-/Row-Helfer', () => {
    it('structureTreeRootDefaultTitle', () => {
        expect(structureTreeRootDefaultTitle(STRUCT_TREE_ROOT_STUDENTS)).toBe('Schüler:innen');
        expect(structureTreeRootDefaultTitle(STRUCT_TREE_ROOT_TEACHERS)).toBe('Lehrer:innen');
        expect(structureTreeRootDefaultTitle(STRUCT_TREE_ROOT_ADMIN)).toBe('Verwaltung');
        expect(structureTreeRootDefaultTitle('x')).toBe('');
    });

    it('structureTreeRootTitle: Custom überschreibt Default', () => {
        const details = { [STRUCT_TREE_ROOT_STUDENTS]: { bezeichnung: 'Schülis' } };
        expect(structureTreeRootTitle(STRUCT_TREE_ROOT_STUDENTS, details)).toBe('Schülis');
        expect(structureTreeRootTitle(STRUCT_TREE_ROOT_TEACHERS, details)).toBe('Lehrer:innen');
    });

    it('defaultStructureTreeRootRow', () => {
        const r = defaultStructureTreeRootRow(STRUCT_TREE_ROOT_TEACHERS);
        expect(r.id).toBe(STRUCT_TREE_ROOT_TEACHERS);
        expect(r.typ).toBe('LehrerInnen');
        expect(r.bezeichnung).toBe('Lehrer:innen');
        expect(r.isStructureTreeRoot).toBe(true);
        expect(defaultStructureTreeRootRow('x')).toBeNull();
    });

    it('mergeStructureTreeRootRow: Custom übernimmt Felder', () => {
        const merged = mergeStructureTreeRootRow(STRUCT_TREE_ROOT_STUDENTS, {
            [STRUCT_TREE_ROOT_STUDENTS]: { bezeichnung: 'Schülis', beschreibung: 'Alle' }
        });
        expect(merged.bezeichnung).toBe('Schülis');
        expect(merged.beschreibung).toBe('Alle');
        expect(merged.typ).toBe('SchuelerInnen'); // typ ist fix, kann nicht überschrieben werden
    });

    it('pickStorableStructureTreeRootFields: Default-Werte für leere Felder', () => {
        const stored = pickStorableStructureTreeRootFields({ bezeichnung: '  X  ' });
        expect(stored.bezeichnung).toBe('X');
        expect(stored.status).toBe('Aktiv');
        expect(stored.syncStatus).toBe('Ausstehend');
    });
});

describe('Migration / Fachgruppen', () => {
    it('migrateLegacyFachschaftenContainer: explizit markiert', () => {
        const rows = [
            { id: 'c', typ: 'Gruppe', fachschaftenRoot: true, bezeichnung: 'x', parentId: '' },
            { id: 'k1', typ: 'Gruppe', parentId: 'c', bezeichnung: 'Math' }
        ];
        const changed = migrateLegacyFachschaftenContainer(rows);
        expect(changed).toBe(true);
        expect(rows.find((r) => r.id === 'c')).toBeUndefined();
        expect(rows.find((r) => r.id === 'k1').parentId).toBe('');
        expect(rows.find((r) => r.id === 'k1').fachschaftFach).toBe(true);
    });

    it('migrateLegacyFachschaftenContainer: namens-basiert', () => {
        const rows = [
            { id: 'c', typ: 'Gruppe', bezeichnung: 'Fachschaften', parentId: '' },
            { id: 'k1', typ: 'Gruppe', parentId: 'c', bezeichnung: 'Math' }
        ];
        migrateLegacyFachschaftenContainer(rows);
        expect(rows.find((r) => r.id === 'c')).toBeUndefined();
    });

    it('migrateLegacyFachschaftenContainer: keine Änderung wenn nichts passt', () => {
        const rows = [{ id: 'x', typ: 'Klasse' }];
        expect(migrateLegacyFachschaftenContainer(rows)).toBe(false);
    });

    it('ensureFachschaftFachGruppen: legt fehlende Fach-Gruppen aus Kursteams an', () => {
        const rows = [
            { id: 'k1', typ: 'Kursteam', ktFach: 'Mathematik' },
            { id: 'k2', typ: 'Kursteam', ktFach: 'Mathematik' }
        ];
        const changed = ensureFachschaftFachGruppen(rows, { subjects: [] });
        expect(changed).toBe(true);
        const fg = rows.filter((r) => r.fachschaftFach);
        expect(fg).toHaveLength(1);
        expect(fg[0].ktFach).toBe('Mathematik');
    });

    it('ensureFachschaftFachGruppen: auch aus tenantSettings.subjects', () => {
        // Funktion macht Early-Return bei leerem `rows` – daher mit Dummy-Anchor.
        const rows = [{ id: 'anchor', typ: 'Klasse', bezeichnung: '1A' }];
        ensureFachschaftFachGruppen(rows, { subjects: [{ code: 'GS' }, { name: 'Englisch' }] });
        const fg = rows.filter((r) => r.fachschaftFach);
        expect(fg.map((r) => r.ktFach).sort()).toEqual(['Englisch', 'GS']);
    });

    it('ensureFachschaftFachGruppen: leeres rows → false und keine Änderung', () => {
        const rows = [];
        expect(ensureFachschaftFachGruppen(rows, { subjects: [{ code: 'GS' }] })).toBe(false);
        expect(rows).toEqual([]);
    });
});

describe('Sortierung', () => {
    it('sortStructureTreeChildren: Jahrgang vor Klasse vor Kursteam …', () => {
        const list = [
            { typ: 'Gruppe', bezeichnung: 'G' },
            { typ: 'Klasse', bezeichnung: '1A' },
            { typ: 'Jahrgang', bezeichnung: '1' },
            { typ: 'Person', bezeichnung: 'p' },
            { typ: 'Kursteam', bezeichnung: 'M' }
        ];
        sortStructureTreeChildren(list);
        expect(list.map((x) => x.typ)).toEqual(['Jahrgang', 'Klasse', 'Kursteam', 'Gruppe', 'Person']);
    });
});

describe('buildTreeOrder', () => {
    it('flacht 2-Ebenen-Baum korrekt ab', () => {
        const rows = [
            { id: 'A', parentId: '', typ: 'Jahrgang', bezeichnung: '1' },
            { id: 'B', parentId: 'A', typ: 'Klasse', bezeichnung: '1A' },
            { id: 'C', parentId: 'A', typ: 'Klasse', bezeichnung: '1B' }
        ];
        const out = buildTreeOrder(rows);
        expect(out.map((x) => x.r.id)).toEqual(['A', 'B', 'C']);
        expect(out.map((x) => x.depth)).toEqual([0, 1, 1]);
    });

    it('schützt gegen Zyklen (depth-Limit)', () => {
        const rows = [
            { id: 'X', parentId: 'X', typ: 'Klasse', bezeichnung: 'cyc' }
        ];
        const out = buildTreeOrder(rows);
        // Sollte nicht endlos rekursieren – Limit greift
        expect(out.length).toBeLessThan(50);
    });
});

describe('buildStructuredTreeOrder', () => {
    /**
     * Mockt `window.ms365StructureRules.inferRootForType` für deterministische
     * Tests.
     */
    beforeEach(() => {
        globalThis.window = {
            ms365StructureRules: {
                inferRootForType: (typ) => {
                    if (typ === 'Jahrgang' || typ === 'Klasse' || typ === 'Kursteam') return 'SchuelerInnen';
                    if (typ === 'Arbeitsgemeinschaft' || typ === 'Gruppe') return 'LehrerInnen';
                    return '';
                }
            }
        };
    });

    afterEach(() => {
        delete globalThis.window;
    });

    it('virtuelle Wurzeln immer vorhanden, auch bei leerem Input', () => {
        const out = buildStructuredTreeOrder([]);
        const roots = out.filter((x) => x.virtual && x.kind === 'root');
        expect(roots.map((r) => r.rootId)).toEqual([
            STRUCT_TREE_ROOT_STUDENTS,
            STRUCT_TREE_ROOT_TEACHERS,
            STRUCT_TREE_ROOT_ADMIN
        ]);
    });

    it('ARGE landet unter Lehrer:innen / ARGEs-Ordner', () => {
        const rows = [
            { id: 'a1', parentId: '', typ: 'Arbeitsgemeinschaft', bezeichnung: 'Fußball' }
        ];
        const out = buildStructuredTreeOrder(rows);
        // ARGEs-Ordner muss als virtual folder erscheinen, gefolgt von a1 (Tiefe 2)
        const argeFolder = out.find((x) => x.virtual && x.rootId === STRUCT_FOLDER_ARGES);
        expect(argeFolder).toBeDefined();
        const argeRow = out.find((x) => !x.virtual && x.r && x.r.id === 'a1');
        expect(argeRow.depth).toBe(2);
    });

    it('Jahrgang landet unter Schüler:innen', () => {
        const rows = [{ id: 'j1', parentId: '', typ: 'Jahrgang', bezeichnung: '5' }];
        const out = buildStructuredTreeOrder(rows);
        const stu = out.find((x) => x.virtual && x.rootId === STRUCT_TREE_ROOT_STUDENTS);
        const idxStu = out.indexOf(stu);
        const j1 = out.find((x) => !x.virtual && x.r && x.r.id === 'j1');
        const idxJ1 = out.indexOf(j1);
        expect(idxJ1).toBe(idxStu + 1);
        expect(j1.depth).toBe(1);
    });

    it('collapsed root: keine Kinder werden ausgegeben', () => {
        const rows = [{ id: 'j1', parentId: '', typ: 'Jahrgang', bezeichnung: '5' }];
        const collapsed = new Set([STRUCT_TREE_ROOT_STUDENTS]);
        const out = buildStructuredTreeOrder(rows, collapsed);
        const j1 = out.find((x) => !x.virtual && x.r && x.r.id === 'j1');
        expect(j1).toBeUndefined();
    });

    it('Top-„Verwaltung" landet unter Verwaltung-Root', () => {
        const rows = [{ id: 'v1', parentId: '', typ: 'Gruppe', bezeichnung: 'Verwaltung' }];
        const out = buildStructuredTreeOrder(rows);
        const v1 = out.find((x) => !x.virtual && x.r && x.r.id === 'v1');
        expect(v1).toBeDefined();
        // sollte direkt nach STRUCT_TREE_ROOT_ADMIN kommen
        const adminIdx = out.findIndex((x) => x.virtual && x.rootId === STRUCT_TREE_ROOT_ADMIN);
        const v1Idx = out.indexOf(v1);
        expect(v1Idx).toBe(adminIdx + 1);
    });
});

describe('Icons', () => {
    it('typeIcon: bekannte Typen', () => {
        expect(typeIcon('Jahrgang')).toBe('bi-layers');
        expect(typeIcon('Klasse')).toBe('bi-collection');
        expect(typeIcon('Kursteam')).toBe('bi-mortarboard');
        expect(typeIcon('Person')).toBe('bi-person');
        expect(typeIcon('UnknownTyp')).toBe('bi-folder2');
    });

    it('treeIconForRow: Tenant + hiddenMembership → bi-mortarboard', () => {
        expect(treeIconForRow({ typ: 'Gruppe', hiddenMembership: true }, 'tenant')).toBe('bi-mortarboard');
        expect(treeIconForRow({ typ: 'Gruppe', hiddenMembership: true }, 'struct')).toBe('bi-people');
        expect(treeIconForRow({ typ: 'Klasse', fachschaftFach: true }, 'struct')).toBe('bi-journal-text');
    });

    it('treeIconForVirtualItem: Ordner-Icons', () => {
        expect(treeIconForVirtualItem({ virtual: true, kind: 'folder', rootId: STRUCT_FOLDER_ARGES })).toBe('bi-people-gear');
        expect(treeIconForVirtualItem({ virtual: true, kind: 'folder', rootId: STRUCT_FOLDER_FACHSCHAFTEN })).toBe('bi-journals');
        expect(treeIconForVirtualItem(null)).toBe('bi-folder2');
    });

    it('graphNodeIconClass: synthetische Ordner', () => {
        expect(graphNodeIconClass({ isStructureFolder: true, bezeichnung: 'ARGEs' })).toBe('bi-people-gear');
        expect(graphNodeIconClass({ isStructureFolder: true, bezeichnung: 'Fachschaften' })).toBe('bi-journals');
        expect(graphNodeIconClass({ fachschaftFach: true })).toBe('bi-journal-text');
        expect(graphNodeIconClass({ typ: 'Klasse' })).toBe('bi-collection');
    });
});
