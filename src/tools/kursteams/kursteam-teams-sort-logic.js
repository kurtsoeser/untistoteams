
function getTeamsSortValue(team, key) {
    if (key === 'status') return team.isValid ? '1' : '0';
    return String(team[key] ?? '').toUpperCase();
}

/**
 * @param {Array} teamsData
 * @param {{ key?: string, dir?: number }} teamsSort
 * @returns {Array<{ team: object, index: number }>}
 */
function sortTeamsWithIndices(teamsData, teamsSort) {
    const sort = teamsSort || { key: 'teamName', dir: 1 };
    const view = (teamsData || []).map((team, index) => ({ team, index }));
    view.sort((a, b) => {
        const ak = getTeamsSortValue(a.team, sort.key);
        const bk = getTeamsSortValue(b.team, sort.key);
        const cmp = ak.localeCompare(bk, 'de');
        if (cmp !== 0) return cmp * sort.dir;
        return (a.index - b.index) * sort.dir;
    });
    return view;
}

window.ms365KursteamTeamsSortLogic = {
    getTeamsSortValue,
    sortTeamsWithIndices
};
