/**
 * Reine Hilfslogik für Strukturbaum / Organigramm (keine DOM-, kein Graph).
 * Aus schulstruktur-sync.js ausgelagert, Verhalten unverändert.
 */

/** Erlaubte Kind-Typen (gleiche Logik wie Anlegen-Modal / Organigramm „+“). */
export function allowedStructureChildTypes(parentType) {
    const p = String(parentType || '');
    if (p === 'SchuelerInnen') return ['Jahrgang'];
    if (p === 'LehrerInnen') return ['Arbeitsgemeinschaft', 'Gruppe'];
    if (p === 'Jahrgang') return ['Klasse'];
    if (p === 'Klasse') return ['Kursteam', 'Gruppe'];
    if (p === 'Gruppe') return ['Gruppe', 'Person'];
    return [];
}

export function structureTreeRowShowsAddChildControl(mode, item) {
    if ((mode !== 'struktur' && mode !== 'match') || !item) return false;
    if (item.virtual) {
        if (item.kind !== 'root') return false;
        return allowedStructureChildTypes(String(item.typ || '')).length > 0;
    }
    if (!item.r) return false;
    return allowedStructureChildTypes(String(item.r.typ || '')).length > 0;
}
