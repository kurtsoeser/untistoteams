(function () {
    'use strict';

    function norm(v) {
        return String(v || '').trim();
    }

    /**
     * Strikte Drop-Regeln für Schul-Organisationsstruktur.
     * parentType kann auch ein virtueller Root sein: 'SchuelerInnen' | 'LehrerInnen'
     */
    function canReparent(childType, parentType) {
        const c = norm(childType);
        const p = norm(parentType);

        // Virtuelle Root-Container
        if (p === 'SchuelerInnen') return c === 'Jahrgang';
        if (p === 'LehrerInnen') return c === 'Arbeitsgemeinschaft' || c === 'Gruppe';

        // Echte Einheiten
        if (p === 'Jahrgang') return c === 'Klasse';
        if (p === 'Klasse') return c === 'Kursteam' || c === 'Gruppe';
        if (p === 'Gruppe') return c === 'Gruppe' || c === 'Person';

        // Sonst nicht erlaubt
        return false;
    }

    function inferRootForType(type) {
        const t = norm(type);
        if (t === 'Jahrgang' || t === 'Klasse' || t === 'Kursteam') return 'SchuelerInnen';
        if (t === 'Arbeitsgemeinschaft') return 'LehrerInnen';
        if (t === 'Gruppe') return 'LehrerInnen';
        if (t === 'Person') return 'LehrerInnen';
        return '';
    }

    /**
     * Index id -> Zeile (Schritt-3-Struktur).
     */
    function buildStructureRowIndex(rows) {
        const m = Object.create(null);
        (rows || []).forEach((r) => {
            if (r && r.id != null) m[String(r.id)] = r;
        });
        return m;
    }

    /**
     * Kursteam braucht Klasse + Fach für Mail-Nickname/Graph.
     * Klasse darf aus dem Feld ktKlasse oder — falls leer — aus dem direkten Eltern-Knoten vom Typ „Klasse“ stammen.
     */
    function resolveKursteamKlasseFach(row, rows) {
        const byId = buildStructureRowIndex(rows);
        let klasse = norm(row && row.ktKlasse);
        let fach = norm(row && row.ktFach);
        const pid = norm(row && row.parentId);
        if (!klasse && pid && byId[pid]) {
            const p = byId[pid];
            if (norm(p.typ) === 'Klasse') klasse = norm(p.bezeichnung);
        }
        return {
            klasse,
            fach,
            hasBoth: !!(klasse && fach)
        };
    }

    window.ms365StructureRules = {
        canReparent,
        inferRootForType,
        buildStructureRowIndex,
        resolveKursteamKlasseFach
    };
})();

