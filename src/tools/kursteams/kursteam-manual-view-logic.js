
function normFilterToken(s) {
    return String(s || '').trim().toUpperCase();
}

/**
 * @param {Array} filteredData
 * @param {{ key?: string, dir?: number }} manualSort
 * @param {{ klasse?: string, fach?: string, lehrer?: string }} filters bereits normalisierte Such-Tokens (upper)
 * @returns {Array<{ row: object, index: number }>}
 */
function applyManualFiltersAndSort(filteredData, manualSort, filters) {
    const f = filters || {};
    const out = [];
    (filteredData || []).forEach((row, index) => {
        const klasse = normFilterToken(row.klasse);
        const fach = normFilterToken(row.fach);
        const lehrer = normFilterToken(row.lehrer);
        if (f.klasse && !klasse.includes(f.klasse)) return;
        if (f.fach && !fach.includes(f.fach)) return;
        if (f.lehrer && !lehrer.includes(f.lehrer)) return;
        out.push({ row, index });
    });

    const key = manualSort && manualSort.key;
    const dir = (manualSort && manualSort.dir) || 1;
    if (key) {
        out.sort((a, b) => {
            const av = normFilterToken(a.row[key] || '');
            const bv = normFilterToken(b.row[key] || '');
            const cmp = av.localeCompare(bv, 'de');
            if (cmp !== 0) return cmp * dir;
            return (a.row.id > b.row.id ? 1 : a.row.id < b.row.id ? -1 : 0) * dir;
        });
    }
    return out;
}

window.ms365KursteamManualViewLogic = {
    normFilterToken,
    applyManualFiltersAndSort
};
