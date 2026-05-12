/**
 * Organigramm-Layout und DOM-Rendering für „Schulstruktur-Sync".
 *
 * Aus `schulstruktur-sync.js` 1:1 ausgelagert (Phase 2 Schnitt 4):
 *
 *  - {@link computeGraphModel}: **pure** Layout-Berechnung – baut aus den
 *    Struktur-Zeilen ein Knoten-/Kanten-Modell mit Positionen und
 *    Canvas-Größe. Hat keine DOM-Abhängigkeit und ist gut testbar.
 *
 *  - {@link renderGraphView}: zeichnet das Modell als HTML/SVG ins DOM
 *    (`#ssGraphNodes`, `#ssGraphEdges`). Nimmt `normRoleKey` als
 *    optionalen Parameter, damit das Modul keine Abhängigkeit auf den
 *    Tool-spezifischen Rollen-Normalizer braucht.
 *
 * Layout-Strategien:
 *  - **horizontal** (Default): Hierarchie nach unten, Geschwister nebeneinander.
 *  - **vertical**: Hierarchie nach rechts, Geschwister untereinander.
 */

import { normStr, escapeHtml, compareDe } from '../../shared/utils/strings.js';
import { getEl } from '../../shared/utils/dom.js';
import {
    STRUCT_FOLDER_ARGES,
    STRUCT_FOLDER_FACHSCHAFTEN,
    structureTreeRootTitle,
    graphNodeIconClass,
    isStructureTreeRootId
} from './schulstruktur-sync-tree.js';
import { allowedStructureChildTypes } from './schulstruktur-sync-helpers.js';

/**
 * Lazy-Lookup auf `window.ms365StructureRules.inferRootForType`.
 * Fallback `() => ''` falls die Rules-Datei (IIFE) noch nicht geladen ist.
 */
function getInferRootForType() {
    const sr = (typeof window !== 'undefined' && window.ms365StructureRules) || null;
    return sr && typeof sr.inferRootForType === 'function'
        ? sr.inferRootForType
        : () => '';
}

/**
 * Berechnet das Layout-Modell des Organigramms.
 *
 * @param {any[]} rows               Struktur-Zeilen.
 * @param {Set<string>} [collapsedSet] IDs von eingeklappten Knoten.
 * @param {object} [structRootDetails] Detail-Daten zu den virtuellen Wurzeln.
 * @param {'horizontal'|'vertical'} [graphLayoutMode]
 * @returns {{
 *   nodes: Map<string, any>,
 *   edges: { from: string, to: string }[],
 *   pos: Map<string, { x: number, y: number, w: number, graphLayout: 'horizontal'|'vertical' }>,
 *   canvas: { width: number, height: number },
 *   rootStudentsId: string,
 *   rootTeachersId: string,
 *   rootAdminId: string,
 *   children: Map<string, string[]>,
 *   graphLayout: 'horizontal'|'vertical'
 * }}
 */
export function computeGraphModel(rows, collapsedSet, structRootDetails, graphLayoutMode) {
    const inferRootForType = getInferRootForType();
    const nodes = new Map();
    const edges = [];
    const layoutVertical = String(graphLayoutMode || '').toLowerCase() === 'vertical';

    const rootStudentsId = '__root_students__';
    const rootTeachersId = '__root_teachers__';
    const rootAdminId = '__root_admin__';
    nodes.set(rootStudentsId, {
        id: rootStudentsId,
        typ: 'SchuelerInnen',
        bezeichnung: structureTreeRootTitle(rootStudentsId, structRootDetails),
        isRoot: true
    });
    nodes.set(rootTeachersId, {
        id: rootTeachersId,
        typ: 'LehrerInnen',
        bezeichnung: structureTreeRootTitle(rootTeachersId, structRootDetails),
        isRoot: true
    });
    nodes.set(rootAdminId, {
        id: rootAdminId,
        typ: 'Verwaltung',
        bezeichnung: structureTreeRootTitle(rootAdminId, structRootDetails),
        isRoot: true
    });

    const folderArgesId = STRUCT_FOLDER_ARGES;
    const folderFachId = STRUCT_FOLDER_FACHSCHAFTEN;
    nodes.set(folderArgesId, {
        id: folderArgesId,
        typ: 'Gruppe',
        bezeichnung: 'ARGEs',
        isStructureFolder: true,
        isRoot: false
    });
    nodes.set(folderFachId, {
        id: folderFachId,
        typ: 'Gruppe',
        bezeichnung: 'Fachschaften',
        isStructureFolder: true,
        isRoot: false
    });

    function isTopVerwaltungNode(r) {
        if (!r) return false;
        const name = String(r.bezeichnung || '').trim().toLowerCase();
        const pid = String(r.parentId || '').trim();
        return !pid && name === 'verwaltung';
    }

    (rows || []).forEach((r) => {
        if (!r || !r.id) return;
        nodes.set(String(r.id), r);
    });

    const children = new Map();
    function addChild(pid, cid) {
        const k = String(pid || '');
        if (!children.has(k)) children.set(k, []);
        children.get(k).push(String(cid));
    }

    const teacherRootOtherIds = [];

    (rows || []).forEach((r) => {
        if (!r || !r.id) return;
        const id = String(r.id);
        const pid = String(r.parentId || '');
        if (pid && nodes.has(pid)) {
            addChild(pid, id);
            edges.push({ from: pid, to: id });
            return;
        }
        const root = isTopVerwaltungNode(r)
            ? rootAdminId
            : inferRootForType(r.typ) === 'LehrerInnen'
              ? rootTeachersId
              : rootStudentsId;
        if (root === rootTeachersId && String(r.typ || '') === 'Arbeitsgemeinschaft') {
            addChild(folderArgesId, id);
            edges.push({ from: folderArgesId, to: id });
            return;
        }
        if (root === rootTeachersId && r.fachschaftFach) {
            addChild(folderFachId, id);
            edges.push({ from: folderFachId, to: id });
            return;
        }
        if (root === rootTeachersId) {
            teacherRootOtherIds.push(id);
            return;
        }
        addChild(root, id);
        edges.push({ from: root, to: id });
    });

    teacherRootOtherIds.sort((a, b) => {
        const ra = nodes.get(a);
        const rb = nodes.get(b);
        return compareDe(String(ra?.bezeichnung || ''), String(rb?.bezeichnung || ''));
    });

    addChild(rootTeachersId, folderArgesId);
    edges.push({ from: rootTeachersId, to: folderArgesId });
    addChild(rootTeachersId, folderFachId);
    edges.push({ from: rootTeachersId, to: folderFachId });
    for (const cid of teacherRootOtherIds) {
        addChild(rootTeachersId, cid);
        edges.push({ from: rootTeachersId, to: cid });
    }

    // Sortiere Geschwister stabil (Lehrer-Wurzel behält Ordner-Reihenfolge)
    for (const [k, list] of children.entries()) {
        if (k === rootTeachersId) continue;
        list.sort((a, b) => {
            const ra = nodes.get(a);
            const rb = nodes.get(b);
            return compareDe(String(ra?.bezeichnung || ''), String(rb?.bezeichnung || ''));
        });
        children.set(k, list);
    }
    const tKids = children.get(rootTeachersId) || [];
    const want = [folderArgesId, folderFachId];
    const rest = tKids.filter((id) => want.indexOf(String(id)) === -1).sort((a, b) => {
        const ra = nodes.get(a);
        const rb = nodes.get(b);
        return compareDe(String(ra?.bezeichnung || ''), String(rb?.bezeichnung || ''));
    });
    children.set(rootTeachersId, [...want, ...rest]);

    /** „Breite" bzw. „Höhe" des Teilbaums in Raster-Einheiten (bei eingeklapptem Knoten = 1). */
    function subtreeSpan(id) {
        const sid = String(id);
        if (collapsedSet && collapsedSet.has(sid)) return 1;
        const kids = children.get(String(id)) || [];
        if (!kids.length) return 1;
        return kids.reduce((acc, kid) => acc + subtreeSpan(kid), 0);
    }

    const pos = new Map();
    // Spacing für große Schulen: mehr Luft zwischen Karten/Kanten
    const xUnit = 280;
    const yUnit = 170;
    /** Vertikal: Hierarchie nach rechts, Geschwister untereinander (weniger horizontale Gesamtbreite). */
    const xDepthUnit = 300;
    /** Abstand der Knoten-Mittelpunkte in Y (Karte ~56px ±28); ~72 ≈ eine Zeile Luft zwischen den Karten. */
    const ySiblingUnitV = 72;
    const rootBlockGapV = 48;

    if (!layoutVertical) {
        function layout(id, depth, x0) {
            const w = subtreeSpan(id);
            const xCenter = x0 + (w * xUnit) / 2;
            pos.set(String(id), { x: xCenter, y: 40 + depth * yUnit, w, graphLayout: 'horizontal' });
            const kids = children.get(String(id)) || [];
            if (collapsedSet && collapsedSet.has(String(id))) return;
            let cursor = x0;
            for (const kid of kids) {
                const kw = subtreeSpan(kid);
                layout(kid, depth + 1, cursor);
                cursor += kw * xUnit;
            }
        }

        const wS = subtreeSpan(rootStudentsId);
        layout(rootStudentsId, 0, 40);
        const xT = 60 + wS * xUnit + 120;
        layout(rootTeachersId, 0, xT);
        const wT = subtreeSpan(rootTeachersId);
        layout(rootAdminId, 0, xT + wT * xUnit + 120);
    } else {
        function layoutV(id, depth, y0) {
            const h = subtreeSpan(id);
            const xCenter = 80 + depth * xDepthUnit;
            const yCenter = y0 + (h * ySiblingUnitV) / 2;
            pos.set(String(id), { x: xCenter, y: yCenter, w: h, graphLayout: 'vertical' });
            if (collapsedSet && collapsedSet.has(String(id))) return;
            const kids = children.get(String(id)) || [];
            let cursor = y0;
            for (const kid of kids) {
                const kh = subtreeSpan(kid);
                layoutV(kid, depth + 1, cursor);
                cursor += kh * ySiblingUnitV;
            }
        }

        let yBlock = 24;
        layoutV(rootStudentsId, 0, yBlock);
        yBlock += subtreeSpan(rootStudentsId) * ySiblingUnitV + rootBlockGapV;
        layoutV(rootTeachersId, 0, yBlock);
        yBlock += subtreeSpan(rootTeachersId) * ySiblingUnitV + rootBlockGapV;
        layoutV(rootAdminId, 0, yBlock);
    }

    let maxX = 0;
    let maxY = 0;
    for (const p of pos.values()) {
        if (p.x > maxX) maxX = p.x;
        if (p.y > maxY) maxY = p.y;
    }
    const padX = layoutVertical ? 360 : 320;
    const padY = layoutVertical ? 120 : 240;
    const canvas = {
        width: Math.max(1200, Math.ceil(maxX + padX)),
        height: Math.max(650, Math.ceil(maxY + padY))
    };

    return {
        nodes,
        edges,
        pos,
        canvas,
        rootStudentsId,
        rootTeachersId,
        rootAdminId,
        children,
        graphLayout: layoutVertical ? 'vertical' : 'horizontal'
    };
}

/**
 * Rendert das Organigramm in `#ssGraphNodes` und `#ssGraphEdges`.
 *
 * @param {any[]} rowsStruktur
 * @param {string|null} selectedId
 * @param {(id: string, opts: { openDetails: boolean }) => void} onSelect
 * @param {{ x: number, y: number, scale: number }} [viewport]
 * @param {Set<string>} [collapsedSet]
 * @param {Map<string, { name?: string, email?: string }>} [personInfoByRole]
 * @param {object} [structRootDetails]
 * @param {'horizontal'|'vertical'} [graphLayoutMode]
 * @param {Function} [_onGraphAddChild] (reserviert)
 * @param {(s: string) => string} [normRoleKey] Optionaler Rollen-Normalizer (Tool-spezifisch).
 * @returns {ReturnType<typeof computeGraphModel> | null} Das berechnete Modell – oder `null`, falls die DOM-Hosts fehlen.
 */
export function renderGraphView(
    rowsStruktur,
    selectedId,
    onSelect,
    viewport,
    collapsedSet,
    personInfoByRole,
    structRootDetails,
    graphLayoutMode,
    _onGraphAddChild,
    normRoleKey
) {
    const roleKey = typeof normRoleKey === 'function' ? normRoleKey : (s) => String(s || '').toLowerCase();
    const wrap = getEl('ssGraphWrap');
    const nodesHost = getEl('ssGraphNodes');
    const edgesSvg = getEl('ssGraphEdges');
    if (!wrap || !nodesHost || !edgesSvg) return null;

    const model = computeGraphModel(rowsStruktur || [], collapsedSet, structRootDetails, graphLayoutMode);
    edgesSvg.setAttribute('width', String(model.canvas.width));
    edgesSvg.setAttribute('height', String(model.canvas.height));
    nodesHost.style.minHeight = String(model.canvas.height) + 'px';
    nodesHost.style.minWidth = String(model.canvas.width) + 'px';

    // Pan/Zoom-Transform (auf beide Layer angewendet)
    const vp = viewport && typeof viewport === 'object' ? viewport : { x: 0, y: 0, scale: 1 };
    const tx = Number.isFinite(vp.x) ? vp.x : 0;
    const ty = Number.isFinite(vp.y) ? vp.y : 0;
    const sc = Number.isFinite(vp.scale) ? vp.scale : 1;
    const tr = `translate(${tx}px, ${ty}px) scale(${sc})`;
    edgesSvg.style.transformOrigin = '0 0';
    edgesSvg.style.transform = tr;
    nodesHost.style.transformOrigin = '0 0';
    nodesHost.style.transform = tr;

    edgesSvg.replaceChildren();
    const isVert = model.graphLayout === 'vertical';
    for (const e of model.edges) {
        const p1 = model.pos.get(String(e.from));
        const p2 = model.pos.get(String(e.to));
        if (!p1 || !p2) continue;
        const line = document.createElementNS('http://www.w3.org/2000/svg', 'path');
        let d;
        if (isVert) {
            const x1 = p1.x + 120;
            const y1 = p1.y;
            const x2 = p2.x - 120;
            const y2 = p2.y;
            const midX = (x1 + x2) / 2;
            d = `M ${x1} ${y1} C ${midX} ${y1}, ${midX} ${y2}, ${x2} ${y2}`;
        } else {
            const x1 = p1.x;
            const y1 = p1.y + 46;
            const x2 = p2.x;
            const y2 = p2.y - 6;
            const midY = (y1 + y2) / 2;
            d = `M ${x1} ${y1} C ${x1} ${midY}, ${x2} ${midY}, ${x2} ${y2}`;
        }
        line.setAttribute('d', d);
        line.setAttribute('stroke', 'rgba(94,114,228,0.35)');
        line.setAttribute('stroke-width', '2');
        line.setAttribute('fill', 'none');
        edgesSvg.appendChild(line);
    }

    nodesHost.replaceChildren();
    for (const [id, node] of model.nodes.entries()) {
        const p = model.pos.get(String(id));
        if (!p) continue;
        const div = document.createElement('div');
        div.className = 'ss-graph-node';
        const isSelected = String(selectedId || '') === String(id);
        div.setAttribute('data-ss-node-id', String(id));
        div.setAttribute('data-ss-node-type', String(node.typ || ''));
        if (node.isStructureFolder || node.isVirtualFach) div.setAttribute('data-ss-graph-synthetic', '1');
        div.draggable = !node.isRoot && !node.isStructureFolder && !node.isVirtualFach;
        div.style.position = 'absolute';
        div.style.left = String(Math.round(p.x - 120)) + 'px';
        div.style.top = String(Math.round(p.y - 28)) + 'px';
        div.style.width = '240px';
        div.style.padding = '10px 12px';
        div.style.borderRadius = '14px';
        div.style.border = isSelected ? '2px solid rgba(45,206,137,0.7)' : '1px solid rgba(94,114,228,0.22)';
        div.style.background = '#fff';
        div.style.boxShadow = isSelected ? '0 12px 26px rgba(45, 206, 137, 0.16)' : '0 10px 22px rgba(50, 50, 93, 0.10)';
        div.style.cursor = 'pointer';
        div.style.userSelect = 'none';
        div.style.display = 'flex';
        div.style.alignItems = 'center';
        div.style.gap = '10px';
        div.style.fontWeight = '800';
        const kids = model.children.get(String(id)) || [];
        const hasKids = !!kids.length && !node.isRoot && !node.isVirtualFach;
        const isCollapsed = !!(collapsedSet && collapsedSet.has(String(id)));
        const toggleBtn = hasKids
            ? `<button type="button" class="ss-graph-toggle" data-ss-toggle-for="${escapeHtml(String(id))}" title="${isCollapsed ? 'Aufklappen' : 'Einklappen'}" aria-label="${isCollapsed ? 'Aufklappen' : 'Einklappen'}" draggable="false" style="border:1px solid rgba(94,114,228,0.22); background:#fff; border-radius:12px; padding:6px 10px; font-weight:1000; cursor:pointer; line-height:1; min-width:42px;">${isCollapsed ? '▸' : '▾'}${kids.length ? `<span class="muted" style="margin-left:6px;font-weight:900;">${kids.length}</span>` : ''}</button>`
            : '';
        const canAddChildren =
            !node.isStructureFolder && !node.isVirtualFach && allowedStructureChildTypes(String(node.typ || '')).length > 0;
        const plusBtn = canAddChildren
            ? `<button type="button" class="ss-graph-plus" data-ss-plus-for="${escapeHtml(String(id))}" title="Unterpunkt hinzufügen" aria-label="Unterpunkt hinzufügen" style="margin-left:auto; border:1px solid rgba(94,114,228,0.22); background:#fff; border-radius:12px; padding:6px 10px; font-weight:1000; cursor:pointer;">+</button>`
            : '';
        const right = `<div style="margin-left:auto;display:flex;gap:8px;align-items:center;">${toggleBtn}${plusBtn}</div>`;
        const subLabel = (() => {
            if (node.isVirtualFach) {
                const n = Number(node.kursteamCount || 0);
                const s = n === 1 ? '1 Kursteam' : String(n) + ' Kursteams';
                return escapeHtml(s);
            }
            const t = String(node.typ || '');
            if (t !== 'Person') return escapeHtml(node.typ || '');
            const storedN = normStr(node.personName);
            const storedE = normStr(node.personEmail).toLowerCase();
            const info = personInfoByRole && personInfoByRole.get ? personInfoByRole.get(roleKey(node.bezeichnung || '')) : null;
            const n = storedN || (info && info.name ? String(info.name).trim() : '');
            const e = storedE || (info && info.email ? String(info.email).trim().toLowerCase() : '');
            if (n && e) return escapeHtml(n + ' · ' + e);
            if (n) return escapeHtml(n);
            if (e) return escapeHtml(e);
            return escapeHtml('Person');
        })();
        div.innerHTML =
            `<i class="bi ${graphNodeIconClass(node)}" style="font-size:1.05em;opacity:0.92;"></i>` +
            `<div style="min-width:0;flex:1;">` +
            `<div style="white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${escapeHtml(node.bezeichnung || '')}</div>` +
            `<div class="muted" style="font-weight:700;font-size:0.86em;margin-top:2px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${subLabel}</div>` +
            `</div>` +
            right;

        if (node.isRoot && isStructureTreeRootId(id)) {
            div.addEventListener('click', (ev) => {
                const t = ev && ev.target;
                if (t && t.closest && (t.closest('.ss-graph-toggle') || t.closest('.ss-graph-plus'))) return;
                if (typeof onSelect === 'function') onSelect(String(id), { openDetails: false });
            });
            div.addEventListener('dblclick', (ev) => {
                const t = ev && ev.target;
                if (t && t.closest && (t.closest('.ss-graph-toggle') || t.closest('.ss-graph-plus'))) return;
                try {
                    ev.preventDefault();
                    ev.stopPropagation();
                } catch {
                    /* ignore */
                }
                if (typeof onSelect === 'function') onSelect(String(id), { openDetails: true });
            });
        } else if (!node.isRoot && !node.isStructureFolder && !node.isVirtualFach) {
            // Single click: nur selektieren (kein Pop-Up).
            div.addEventListener('click', (ev) => {
                const t = ev && ev.target;
                if (t && t.closest && (t.closest('.ss-graph-toggle') || t.closest('.ss-graph-plus'))) return;
                if (typeof onSelect === 'function') onSelect(String(id), { openDetails: false });
            });
            // Double click: Details öffnen.
            div.addEventListener('dblclick', (ev) => {
                const t = ev && ev.target;
                if (t && t.closest && (t.closest('.ss-graph-toggle') || t.closest('.ss-graph-plus'))) return;
                try {
                    ev.preventDefault();
                    ev.stopPropagation();
                } catch {
                    // ignore
                }
                if (typeof onSelect === 'function') onSelect(String(id), { openDetails: true });
            });
        }
        nodesHost.appendChild(div);
    }

    return model;
}
