
/**
 * Rohdaten → gefilterte Zeilen (Fach-Ausschluss, Klassenpflicht, optional Duplikate).
 * @param {Array} rawData
 * @param {string[]} excludeSubjects bereits normalisierte Fach-Kürzel (z. B. upper case)
 * @param {boolean} removeDuplicates
 * @returns {{ filtered: Array, removedByFilter: number, removedByDuplicate: number }}
 */
function applyRowFilters(rawData, excludeSubjects, removeDuplicates) {
    const ex = Array.isArray(excludeSubjects) ? excludeSubjects : [];

    let filtered = (rawData || []).filter((row) => {
        if (!row.fach || !row.lehrer) return false;
        const fach = row.fach.toUpperCase().trim();
        if (ex.includes(fach)) return false;
        if (!row.klasse || row.klasse.trim() === '') return false;
        return true;
    });

    const countAfterPass1 = filtered.length;
    const removedByFilter = (rawData || []).length - countAfterPass1;

    if (removeDuplicates) {
        const seen = new Set();
        filtered = filtered.filter((row) => {
            const key = `${row.klasse}-${row.fach}-${row.lehrer}-${row.gruppe}`;
            if (seen.has(key)) return false;
            seen.add(key);
            return true;
        });
    }

    const removedByDuplicate = countAfterPass1 - filtered.length;

    return { filtered, removedByFilter, removedByDuplicate };
}

window.ms365KursteamFilterLogic = {
    applyRowFilters
};
