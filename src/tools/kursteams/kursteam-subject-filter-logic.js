
function normalizeSubjectToken(s) {
    return String(s || '').trim().toUpperCase();
}

function parseExcludeSubjectsFromString(value) {
    return String(value || '')
        .split(',')
        .map(normalizeSubjectToken)
        .filter((x) => x.length > 0);
}

function uniqSortedSubjectTokens(tokens) {
    const uniq = Array.from(new Set((tokens || []).map(normalizeSubjectToken).filter(Boolean)));
    uniq.sort((a, b) => a.localeCompare(b, 'de'));
    return uniq;
}

function collectSubjectsFromRows(rows) {
    const set = new Set();
    (rows || []).forEach((r) => {
        const t = normalizeSubjectToken(r && r.fach);
        if (t) set.add(t);
    });
    return Array.from(set).sort((a, b) => a.localeCompare(b, 'de'));
}

function subjectFilterSummaryText(availableCount, excludedCount) {
    const a = Number(availableCount) || 0;
    const e = Number(excludedCount) || 0;
    if (!a) {
        return 'Noch keine Daten: Importieren Sie zuerst Zeilen in Schritt 1 oder fügen Sie manuell Unterrichtszeilen hinzu.';
    }
    return `${a} Fach/Fächer erkannt. ${e} ausgeschlossen.`;
}

window.ms365KursteamSubjectFilterLogic = {
    normalizeSubjectToken,
    parseExcludeSubjectsFromString,
    uniqSortedSubjectTokens,
    collectSubjectsFromRows,
    subjectFilterSummaryText
};
