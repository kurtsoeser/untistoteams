import './kursteam-utils.js';
import './kursteam-team-names.js';
import './kursteam-filter-logic.js';
import './kursteam-subject-filter-logic.js';
import './kursteam-subject-filter-ui.js';
import './kursteam-manual-view-logic.js';
import './kursteam-data-tables-ui.js';
import './kursteam-teams-sort-logic.js';
import './kursteam-team-build.js';
import './kursteam-team-name-builder-ui.js';
import './kursteam-teams-table-ui.js';
import './kursteam-teams-actions.js';
import './kursteam-ui.js';
import './kursteam-storage.js';
import './kursteam-import.js';
import './kursteam-teacher-mapping.js';
import './kursteam-members.js';
import './kursteam-steps-export.js';
import './kursteam-graph.js';

const ns = (window.ms365Kursteam = window.ms365Kursteam || {});

if (typeof window.ms365AssertModules !== 'function') {
    throw new Error(
        'ms365-module-guard.js muss vor kursteam-teams.js geladen werden (tools/kursteams.html).'
    );
}

const KUI = window.ms365KursteamTeamNameBuilderUI;
const KSub = window.ms365KursteamSubjectFilterUI;
const KActions = window.ms365KursteamTeamsActions;
window.ms365AssertModules(
    {
        teamNames: window.ms365KursteamTeamNames,
        KUI,
        KSub,
        KActions
    },
    'kursteam-teams.js'
);

KUI.mount(ns);
KSub.mount(ns);
KActions.mount(ns);

window.startKursteamFromWebuntis = ns.startKursteamFromWebuntis;
window.startKursteamManual = ns.startKursteamManual;
window.addManualDataRow = ns.addManualDataRow;
window.addManualDataRowInline = ns.addManualDataRowInline;
window.applyFilters = ns.applyFilters;
window.resetFilters = ns.resetFilters;
window.generateTeamNames = ns.generateTeamNames;
window.addManualKursteamTeam = ns.addManualKursteamTeam;

if (typeof ns.refreshSubjectFilterUI === 'function') ns.refreshSubjectFilterUI();
if (typeof ns.renderTeamNameBuilder === 'function') ns.renderTeamNameBuilder();
