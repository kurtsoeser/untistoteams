
const KT = window.ms365KursteamTeamNames;
window.ms365AssertModules({ KT }, 'kursteam-team-build.js');

/**
 * Erzeugt die Team-Liste aus gefilterten Unterrichtszeilen (ohne Duplikat-Auflösung).
 * @param {Array} rows ns.filteredData
 * @param {object} options
 */
function buildTeamEntriesFromRows(rows, options) {
    const yearPrefix = options.yearPrefix;
    const emailDomain = options.emailDomain;
    const separator = options.separator != null ? options.separator : ' | ';
    const pattern = options.pattern;
    const combineClassNames = options.combineClassNames;
    const buildGruppenmailBase = options.buildGruppenmailBase;
    const INVALID_CHARS_REPLACE = options.INVALID_CHARS_REPLACE;
    const INVALID_CHARS_TEST = options.INVALID_CHARS_TEST;
    const teacherEmailMapping = options.teacherEmailMapping || {};

    return (rows || []).map((row) => {
        let klasseForName = row.klasse;
        if (row.klasse && row.klasse.includes(',')) klasseForName = combineClassNames(row.klasse);

        const teamName = pattern
            ? KT.buildTeamNameFromPattern(pattern, {
                  yearPrefix,
                  klasse: klasseForName,
                  fach: row.fach,
                  gruppe: row.gruppe
              })
            : `${yearPrefix}${separator}${klasseForName}${separator}${row.fach}`;
        let klasseForGruppenmail = klasseForName;
        try {
            const adv = window.ms365AppDataV2;
            if (adv && typeof adv.getClassTeamGruppenmailForKlasse === 'function') {
                const stable = adv.getClassTeamGruppenmailForKlasse(row.klasse);
                if (stable) klasseForGruppenmail = stable;
            }
        } catch {
            // ignore
        }
        const gruppenmailRaw = buildGruppenmailBase(yearPrefix, klasseForGruppenmail, row.fach, row.gruppe).replace(/\s+/g, '-');

        const originalGruppenmail = gruppenmailRaw;
        let gruppenmail = gruppenmailRaw.replace(INVALID_CHARS_REPLACE, '');

        let besitzer = '';
        const lehrerCode = row.lehrer.toUpperCase().trim();
        if (teacherEmailMapping[lehrerCode]) {
            besitzer = teacherEmailMapping[lehrerCode];
        } else {
            besitzer = row.lehrer.toLowerCase().trim().replace(/\s+/g, '.');
            besitzer = besitzer.replace(INVALID_CHARS_REPLACE, '');
            if (!besitzer.includes('@')) besitzer += emailDomain;
        }

        const hasInvalidChars = INVALID_CHARS_TEST.test(originalGruppenmail);
        const isValid = !hasInvalidChars && teamName && gruppenmail && besitzer && gruppenmail.length > 0;
        const mappingUsed = !!teacherEmailMapping[lehrerCode];

        return {
            teamName,
            gruppenmail,
            besitzer,
            isValid,
            error: hasInvalidChars ? 'Ungültige Zeichen in Gruppenmail' : !isValid ? 'Unvollständige Daten' : null,
            originalClass: row.klasse,
            gruppe: row.gruppe,
            mappingUsed,
            lehrerCode,
            mailNicknameAdjusted: false
        };
    });
}

window.ms365KursteamTeamBuild = {
    buildTeamEntriesFromRows
};
