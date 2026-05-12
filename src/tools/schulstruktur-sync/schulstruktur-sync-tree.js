/**
 * Tree-/Strukturbaum-Helfer für „Schulstruktur-Sync".
 *
 * Aus `schulstruktur-sync.js` 1:1 ausgelagert (Phase 2 Schnitt 3). Enthält:
 *  - Konstanten der virtuellen Wurzeln und Ordner
 *  - Klassifikationen (Synthetisch, Root, Top-Verwaltung, …)
 *  - Title-/Row-Helfer für die drei virtuellen Wurzeln
 *  - Migrationen (legacy „Fachschaften"-Container) und Fach-Gruppen-Sicherung
 *  - Sortierung + Tree-Builders (`buildTreeOrder`, `buildStructuredTreeOrder`)
 *  - Icon-Bestimmung (Bootstrap-Icons-Klassen)
 *
 * Reine Funktionen — keine DOM-/Storage-Zugriffe. Lediglich
 * `getInferRootForType` greift lazy auf `window.ms365StructureRules`,
 * mit Fallback `() => ''` (identisch zum ursprünglichen Verhalten).
 */

import { normStr, compareDe } from '../../shared/utils/strings.js';
import { currentSchoolYearLabel } from './schulstruktur-sync-naming.js';

// ────────── Konstanten ──────────

/** Virtuelle Wurzel „Schüler:innen". */
export const STRUCT_TREE_ROOT_STUDENTS = '__root_students__';
/** Virtuelle Wurzel „Lehrer:innen". */
export const STRUCT_TREE_ROOT_TEACHERS = '__root_teachers__';
/** Virtuelle Wurzel „Verwaltung". */
export const STRUCT_TREE_ROOT_ADMIN = '__root_admin__';
/** Virtueller Ordner unter „Lehrer:innen": ARGEs. */
export const STRUCT_FOLDER_ARGES = '__ss_folder_arges__';
/** Virtueller Ordner unter „Lehrer:innen": Fachschaften. */
export const STRUCT_FOLDER_FACHSCHAFTEN = '__ss_folder_fachschaften__';

// ────────── kleine Helfer ──────────

/** Indexiert `rows` nach `id`. */
export function byId(rows) {
    const map = new Map();
    for (let i = 0; i < rows.length; i++) {
        const r = rows[i];
        if (r && r.id) map.set(String(r.id), r);
    }
    return map;
}

/** Kurze, stabile ID für lokale Mock-Daten (nicht kryptografisch). */
export function uid() {
    return 'id-' + Math.random().toString(16).slice(2) + '-' + Date.now().toString(16);
}

/**
 * Liefert `inferRootForType` aus `window.ms365StructureRules`, mit Fallback
 * `() => ''` falls die Rules-Datei (noch) nicht geladen ist. So bleibt das
 * Tree-Modul ohne harte Importabhängigkeit zu `structure-rules.js`
 * (die noch ein IIFE-Modul ist).
 */
function getInferRootForType() {
    const sr = (typeof window !== 'undefined' && window.ms365StructureRules) || null;
    return sr && typeof sr.inferRootForType === 'function'
        ? sr.inferRootForType
        : () => '';
}

// ────────── Klassifikationen ──────────

/** Nur Anzeige/DnD im Organigramm: keine persistierten Zeilen. */
export function isStructureSyntheticGraphNodeId(id) {
    const s = String(id || '');
    if (!s) return false;
    if (s.startsWith('__root_')) return true;
    if (s === STRUCT_FOLDER_ARGES || s === STRUCT_FOLDER_FACHSCHAFTEN) return true;
    if (s.startsWith('__ss_fach__')) return true;
    if (s.startsWith('__ss_kursteams__')) return true;
    return false;
}

/** Die drei einklappbaren Hauptäste im Baum (mit Detail-Panel). */
export function isStructureTreeRootId(id) {
    const s = String(id || '');
    return s === STRUCT_TREE_ROOT_STUDENTS || s === STRUCT_TREE_ROOT_TEACHERS || s === STRUCT_TREE_ROOT_ADMIN;
}

/** Top-Level-Eintrag „Verwaltung" (oberste Hierarchie, ohne Parent). */
export function isTopVerwaltungStructureNode(r) {
    if (!r) return false;
    const name = String(r.bezeichnung || '').trim().toLowerCase();
    const pid = String(r.parentId || '').trim();
    return !pid && name === 'verwaltung';
}

/** Synthetische Knoten-ID für eine Lehrer-Fach-Gruppe. */
export function teacherFachVirtualId(fachLabel) {
    const key = String(fachLabel || '').trim() || '(ohne Fach)';
    return '__ss_fach__' + encodeURIComponent(key).replace(/%/g, '_');
}

// ────────── Title-Helfer für virtuelle Wurzeln ──────────

export function structureTreeRootDefaultTitle(rootId) {
    const s = String(rootId || '');
    if (s === STRUCT_TREE_ROOT_STUDENTS) return 'Schüler:innen';
    if (s === STRUCT_TREE_ROOT_TEACHERS) return 'Lehrer:innen';
    if (s === STRUCT_TREE_ROOT_ADMIN) return 'Verwaltung';
    return '';
}

export function structureTreeRootTitle(rootId, structRootDetails) {
    const o = structRootDetails && typeof structRootDetails === 'object' ? structRootDetails[String(rootId)] : null;
    const custom = o && normStr(o.bezeichnung);
    if (custom) return custom;
    return structureTreeRootDefaultTitle(rootId);
}

// ────────── Row-Helfer für virtuelle Wurzeln ──────────

export function defaultStructureTreeRootRow(rootId) {
    if (!isStructureTreeRootId(rootId)) return null;
    const typ =
        String(rootId) === STRUCT_TREE_ROOT_STUDENTS
            ? 'SchuelerInnen'
            : String(rootId) === STRUCT_TREE_ROOT_TEACHERS
              ? 'LehrerInnen'
              : 'Verwaltung';
    return {
        id: String(rootId),
        parentId: '',
        typ,
        bezeichnung: structureTreeRootDefaultTitle(rootId),
        beschreibung: '',
        schuljahr: '',
        status: 'Aktiv',
        syncStatus: 'Ausstehend',
        letzteFehlermeldung: '',
        jgYear: '',
        jgSuffix: '',
        argeCode: '',
        argeName: '',
        ktKlasse: '',
        ktFach: '',
        ktGruppe: '',
        tenantGroupId: '',
        tenantMailNickname: '',
        tenantTarget: '',
        tenantVisibility: '',
        isStructureTreeRoot: true
    };
}

export function mergeStructureTreeRootRow(rootId, structRootDetails) {
    const base = defaultStructureTreeRootRow(rootId);
    if (!base) return null;
    const o = structRootDetails && typeof structRootDetails === 'object' ? structRootDetails[String(rootId)] : null;
    if (!o || typeof o !== 'object') return base;
    return Object.assign({}, base, o, {
        id: String(rootId),
        parentId: '',
        typ: base.typ,
        isStructureTreeRoot: true
    });
}

export function pickStorableStructureTreeRootFields(row) {
    return {
        bezeichnung: normStr(row.bezeichnung),
        beschreibung: normStr(row.beschreibung),
        schuljahr: normStr(row.schuljahr),
        status: normStr(row.status) || 'Aktiv',
        syncStatus: normStr(row.syncStatus) || 'Ausstehend',
        letzteFehlermeldung: normStr(row.letzteFehlermeldung),
        jgYear: normStr(row.jgYear),
        jgSuffix: normStr(row.jgSuffix),
        argeCode: normStr(row.argeCode),
        argeName: normStr(row.argeName),
        ktKlasse: normStr(row.ktKlasse),
        ktFach: normStr(row.ktFach),
        ktGruppe: normStr(row.ktGruppe),
        tenantGroupId: normStr(row.tenantGroupId),
        tenantMailNickname: normStr(row.tenantMailNickname),
        tenantTarget: normStr(row.tenantTarget),
        tenantVisibility: normStr(row.tenantVisibility)
    };
}

// ────────── Migration / Fachgruppen-Sicherung ──────────

/** Entfernt die frühere echte Container-„Gruppe Fachschaften"; Kinder werden Top-Level-Fach-Gruppen. */
export function migrateLegacyFachschaftenContainer(rows) {
    if (!rows || !rows.length) return false;
    let changed = false;
    const removeIds = new Set();
    for (const r of rows) {
        if (!r || r.typ !== 'Gruppe') continue;
        const isMarked = r.fachschaftenRoot === true;
        const isNamedContainer =
            String(r.parentId || '') === '' &&
            String(r.bezeichnung || '').trim().toLowerCase() === 'fachschaften' &&
            !r.fachschaftFach;
        if (!isMarked && !isNamedContainer) continue;
        const cid = String(r.id);
        for (const x of rows) {
            if (x && String(x.parentId || '') === cid) {
                x.parentId = '';
                if (x.typ === 'Gruppe') x.fachschaftFach = true;
                changed = true;
            }
        }
        removeIds.add(cid);
        changed = true;
    }
    if (!removeIds.size) return changed;
    for (let i = rows.length - 1; i >= 0; i--) {
        if (rows[i] && removeIds.has(String(rows[i].id))) rows.splice(i, 1);
    }
    return changed;
}

/**
 * Keine persistierte „Fachschaften"-Gruppe: nur je Fach eine Gruppe
 * (Top-Level), Anzeige unter virtuellem Ordner wie bei ARGEs.
 * @returns {boolean} `true` wenn Daten geändert wurden.
 */
export function ensureFachschaftFachGruppen(rows, tenantSettings) {
    if (!rows || !rows.length) return false;
    let changed = migrateLegacyFachschaftenContainer(rows);
    const schuljahr = currentSchoolYearLabel();
    const fachKeys = new Set();
    for (const r of rows) {
        if (r && r.typ === 'Kursteam') {
            fachKeys.add(String(r.ktFach || '').trim() || '(ohne Fach)');
        }
    }
    const subj = tenantSettings && Array.isArray(tenantSettings.subjects) ? tenantSettings.subjects : [];
    for (const s of subj) {
        const key = String((s && (s.code || s.name)) || '').trim();
        if (key) fachKeys.add(key);
    }
    for (const f of Array.from(fachKeys).sort(compareDe)) {
        const has = rows.some(
            (r) =>
                r &&
                r.typ === 'Gruppe' &&
                r.fachschaftFach &&
                String(r.parentId || '') === '' &&
                (String(r.ktFach || '').trim() === f ||
                    (!String(r.ktFach || '').trim() && String(r.bezeichnung || '').trim() === f))
        );
        if (!has) {
            rows.push({
                id: uid(),
                parentId: '',
                typ: 'Gruppe',
                bezeichnung: f,
                ktFach: f,
                schuljahr,
                status: 'Aktiv',
                syncStatus: 'Ausstehend',
                letzteFehlermeldung: '',
                fachschaftFach: true
            });
            changed = true;
        }
    }
    return changed;
}

// ────────── Sortierung ──────────

export function sortStructureTreeChildren(list) {
    list.sort((a, b) => {
        const rank = (x) =>
            x.typ === 'Jahrgang'
                ? 1
                : x.typ === 'Klasse'
                  ? 2
                  : x.typ === 'Kursteam'
                    ? 3
                    : x.typ === 'Arbeitsgemeinschaft'
                      ? 4
                      : x.typ === 'Gruppe'
                        ? 5
                        : x.typ === 'Person'
                          ? 6
                          : 99;
        const ra = rank(a);
        const rb = rank(b);
        if (ra !== rb) return ra - rb;
        return compareDe(a.bezeichnung, b.bezeichnung);
    });
}

// ────────── Tree-Builder ──────────

/** Einfacher Tree-Walk: Liste flach, je Eintrag `{ r, depth }`. */
export function buildTreeOrder(rows) {
    const map = byId(rows);
    const children = new Map();
    for (const r of rows) {
        const pid = String(r.parentId || '');
        if (!children.has(pid)) children.set(pid, []);
        children.get(pid).push(r);
    }
    for (const [k, list] of children.entries()) {
        list.sort((a, b) => {
            // Jahrgang vor Klasse vor ARGE vor Gruppe (nur optische Sortierung)
            const rank = (x) =>
                x.typ === 'Jahrgang'
                    ? 1
                    : x.typ === 'Klasse'
                      ? 2
                      : x.typ === 'Arbeitsgemeinschaft'
                        ? 3
                        : x.typ === 'Kursteam'
                          ? 4
                          : 5;
            const ra = rank(a);
            const rb = rank(b);
            if (ra !== rb) return ra - rb;
            return compareDe(a.bezeichnung, b.bezeichnung);
        });
        children.set(k, list);
    }

    const out = [];
    function walk(parentId, depth) {
        const list = children.get(String(parentId || '')) || [];
        for (const r of list) {
            out.push({ r, depth });
            // Schutz gegen versehentliche Zyklen
            if (depth < 10 && map.has(String(r.id))) walk(r.id, depth + 1);
        }
    }
    walk('', 0);
    return out;
}

/**
 * Baum wie im Organigramm: virtuelle Wurzeln Schüler:innen / Lehrer:innen /
 * Verwaltung, unter Lehrer:innen Standard-Ordner „ARGEs" und „Fachschaften",
 * plus einklappbare Teilbäume (ids teilen sich mit GRAPH_COLLAPSE_KEY).
 * Kursteams erscheinen als echte Knoten unter „Klasse".
 */
export function buildStructuredTreeOrder(rows, collapsedSet, structRootDetails) {
    const inferRootForType = getInferRootForType();
    const baseRows = (rows || []).filter((r) => r && r.id);
    const idMap = byId(baseRows);
    const children = new Map();
    for (const r of baseRows) {
        const pid = String(r.parentId || '');
        if (!children.has(pid)) children.set(pid, []);
        children.get(pid).push(r);
    }
    for (const list of children.values()) sortStructureTreeChildren(list);

    function virtualRootForTop(r) {
        if (isTopVerwaltungStructureNode(r)) return STRUCT_TREE_ROOT_ADMIN;
        if (inferRootForType(r.typ) === 'LehrerInnen') return STRUCT_TREE_ROOT_TEACHERS;
        return STRUCT_TREE_ROOT_STUDENTS;
    }

    const buckets = new Map([
        [STRUCT_TREE_ROOT_STUDENTS, []],
        [STRUCT_TREE_ROOT_TEACHERS, []],
        [STRUCT_TREE_ROOT_ADMIN, []],
    ]);
    for (const r of baseRows) {
        const pid = String(r.parentId || '');
        if (pid && idMap.has(pid)) continue;
        const b = buckets.get(virtualRootForTop(r));
        if (b) b.push(r);
    }
    for (const list of buckets.values()) sortStructureTreeChildren(list);

    const argeTop = baseRows
        .filter((r) => {
            if (!r || r.typ !== 'Arbeitsgemeinschaft') return false;
            const pid = String(r.parentId || '');
            return !pid || !idMap.has(pid);
        })
        .slice()
        .sort((a, b) => compareDe(a.bezeichnung, b.bezeichnung));

    const teacherOtherTop = baseRows
        .filter((r) => {
            if (!r) return false;
            const pid = String(r.parentId || '');
            if (pid && idMap.has(pid)) return false;
            if (virtualRootForTop(r) !== STRUCT_TREE_ROOT_TEACHERS) return false;
            if (r.typ === 'Arbeitsgemeinschaft') return false;
            if (r.fachschaftFach) return false;
            return true;
        })
        .slice()
        .sort((a, b) => compareDe(a.bezeichnung, b.bezeichnung));

    const fachGruppenTop = baseRows
        .filter((r) => {
            if (!r || r.typ !== 'Gruppe' || !r.fachschaftFach) return false;
            const pid = String(r.parentId || '');
            return !pid || !idMap.has(pid);
        })
        .slice()
        .sort((a, b) => compareDe(a.bezeichnung, b.bezeichnung));

    const virtualRoots = [
        { id: STRUCT_TREE_ROOT_STUDENTS, label: structureTreeRootTitle(STRUCT_TREE_ROOT_STUDENTS, structRootDetails), typ: 'SchuelerInnen' },
        { id: STRUCT_TREE_ROOT_TEACHERS, label: structureTreeRootTitle(STRUCT_TREE_ROOT_TEACHERS, structRootDetails), typ: 'LehrerInnen' },
        { id: STRUCT_TREE_ROOT_ADMIN, label: structureTreeRootTitle(STRUCT_TREE_ROOT_ADMIN, structRootDetails), typ: 'Verwaltung' },
    ];

    const collapsed = collapsedSet || new Set();
    const out = [];

    function walkReal(r, depth) {
        const list = children.get(String(r.id)) || [];
        const hasKids = list.length > 0;
        out.push({ r, depth, virtual: false, hasKids });
        if (collapsed.has(String(r.id))) return;
        for (const k of list) walkReal(k, depth + 1);
    }

    for (const vr of virtualRoots) {
        const rootRowMerged = mergeStructureTreeRootRow(vr.id, structRootDetails);
        const rootSyncStatus = rootRowMerged && rootRowMerged.syncStatus ? String(rootRowMerged.syncStatus) : 'Ausstehend';
        if (vr.id === STRUCT_TREE_ROOT_TEACHERS) {
            const hasFachBranch = fachGruppenTop.length > 0;
            const hasTeacherKids =
                argeTop.length > 0 || hasFachBranch || teacherOtherTop.length > 0;
            out.push({
                virtual: true,
                kind: 'root',
                rootId: vr.id,
                label: vr.label,
                typ: vr.typ,
                depth: 0,
                hasKids: hasTeacherKids,
                rootSyncStatus,
            });
            if (!hasTeacherKids) continue;
            if (collapsed.has(vr.id)) continue;

            out.push({
                virtual: true,
                kind: 'folder',
                rootId: STRUCT_FOLDER_ARGES,
                label: 'ARGEs',
                typ: 'Gruppe',
                typLabel: 'Ordner',
                depth: 1,
                hasKids: argeTop.length > 0,
            });
            if (argeTop.length && !collapsed.has(STRUCT_FOLDER_ARGES)) {
                for (const r of argeTop) walkReal(r, 2);
            }

            out.push({
                virtual: true,
                kind: 'folder',
                rootId: STRUCT_FOLDER_FACHSCHAFTEN,
                label: 'Fachschaften',
                typ: 'Gruppe',
                typLabel: 'Ordner',
                depth: 1,
                hasKids: hasFachBranch,
            });
            if (hasFachBranch && !collapsed.has(STRUCT_FOLDER_FACHSCHAFTEN)) {
                for (const r of fachGruppenTop) walkReal(r, 2);
            }

            for (const r of teacherOtherTop) walkReal(r, 1);
            continue;
        }

        const kids = buckets.get(vr.id) || [];
        out.push({
            virtual: true,
            kind: 'root',
            rootId: vr.id,
            label: vr.label,
            typ: vr.typ,
            depth: 0,
            hasKids: kids.length > 0,
            rootSyncStatus,
        });
        if (!kids.length) continue;
        if (collapsed.has(vr.id)) continue;
        for (const r of kids) walkReal(r, 1);
    }
    return out;
}

// ────────── Icons ──────────

export function typeIcon(typ) {
    const t = String(typ || '');
    if (t === 'Jahrgang') return 'bi-layers';
    if (t === 'Klasse') return 'bi-collection';
    if (t === 'Kursteam') return 'bi-mortarboard';
    if (t === 'Arbeitsgemeinschaft') return 'bi-people-gear';
    if (t === 'Gruppe') return 'bi-people';
    if (t === 'Person') return 'bi-person';
    if (t === 'SchuelerInnen') return 'bi-people-fill';
    if (t === 'LehrerInnen') return 'bi-person-badge';
    if (t === 'Verwaltung') return 'bi-building';
    // Tenant / Microsoft 365 (Verwalten-Tab)
    if (t === 'Team') return 'bi-camera-video-fill';
    if (t === 'Sicherheitsgruppe') return 'bi-shield-lock';
    if (t === 'E-Mail-Sicherheitsgruppe' || t === 'E‑Mail‑Sicherheitsgruppe') return 'bi-envelope-paper';
    return 'bi-folder2';
}

/** Icon für echte Zeilen in der Baumansicht (inkl. Kursteam über hiddenMembership im Tenant-Modus). */
export function treeIconForRow(r, mode) {
    if (!r) return typeIcon('');
    if (mode === 'tenant' && r.hiddenMembership) return 'bi-mortarboard';
    if (r.fachschaftFach) return 'bi-journal-text';
    return typeIcon(r.typ);
}

/** Virtuelle Baumzeilen (Wurzeln, ARGE-/Fachschaften-Ordner, Fachschaft). */
export function treeIconForVirtualItem(item) {
    if (!item || !item.virtual) return typeIcon('');
    if (item.kind === 'folder') {
        if (String(item.rootId || '') === STRUCT_FOLDER_ARGES) return 'bi-people-gear';
        if (String(item.rootId || '') === STRUCT_FOLDER_FACHSCHAFTEN) return 'bi-journals';
        return 'bi-folder2';
    }
    if (item.kind === 'fach') return 'bi-journal-text';
    if (item.kind === 'kursteams') return 'bi-mortarboard';
    return typeIcon(item.typ);
}

/** Organigramm: gleiche Semantik wie Baum für synthetische Ordner/Fächer. */
export function graphNodeIconClass(node) {
    if (!node) return typeIcon('');
    if (node.fachschaftFach) return 'bi-journal-text';
    if (node.isVirtualFach) return 'bi-journal-text';
    if (node.isStructureFolder) {
        const b = String(node.bezeichnung || '');
        if (b === 'ARGEs') return 'bi-people-gear';
        if (b === 'Fachschaften') return 'bi-journals';
        return 'bi-folder2';
    }
    return typeIcon(node.typ);
}
