
const KF = window.ms365KursteamFilterLogic;
const KTB = window.ms365KursteamTeamBuild;
const KTU = window.ms365KursteamTeamsTableUI;
const KData = window.ms365KursteamDataTablesUI;
window.ms365AssertModules(
    { KF, KTB, KTU, KData },
    'kursteam-teams-actions.js'
);

/**
 * Datenzeilen, Filter, Teamgenerierung, Team-Modals – nach KUI/KSub, inkl. KData.mount.
 * @param {object} ns window.ms365Kursteam
 */
function mount(ns) {
    ns.updateDataRowField = function updateDataRowField(rowId, field, value) {
        const idx = ns.filteredData.findIndex((r) => r && r.id === rowId);
        if (idx >= 0) {
            ns.filteredData[idx][field] = value;
        }
        const ridx = ns.rawData.findIndex((r) => r && r.id === rowId);
        if (ridx >= 0) {
            ns.rawData[ridx][field] = value;
        }
        ns.invalidateTeams();
    };

    KData.mount(ns);

    ns.applyFilters = function applyFilters() {
        const excludeSubjects = ns.parseExcludeSubjectsFromInput();
        const removeDuplicates = document.getElementById('removeDuplicates').checked;

        try {
            const btn = document.getElementById('btnApplyFilters');
            if (btn) btn.className = 'btn';
        } catch (e) {
            /* ignore */
        }

        const r = KF.applyRowFilters(ns.rawData, excludeSubjects, removeDuplicates);
        ns.filteredData = r.filtered;
        ns.invalidateTeams();
        document.getElementById('filteredRecords').textContent = r.filtered.length;
        document.getElementById('removedDuplicates').textContent = r.removedByFilter + r.removedByDuplicate;
        document.getElementById('filterStats').style.display = 'block';
        ns.displayFilteredData();
    };

    ns.addManualDataRowInline = function addManualDataRowInline() {
        const id = Date.now() + Math.random();
        const row = {
            id,
            klasse: '',
            fach: '',
            lehrer: '',
            gruppe: '',
            original: { manualInline: true }
        };
        ns.rawData.push(row);
        ns.filteredData.push(row);
        ns.kursteamEntryMode = ns.kursteamEntryMode === 'unset' ? 'manual' : ns.kursteamEntryMode;
        ns.invalidateTeams();
        if (typeof ns.refreshSubjectFilterUI === 'function') ns.refreshSubjectFilterUI();
        ns.clearManualFilterInputs();
        if (typeof ns.displayEditableData === 'function') ns.displayEditableData();

        try {
            const tbody = document.getElementById('editableDataTableBody');
            const lastRow = tbody ? tbody.lastElementChild : null;
            if (lastRow && lastRow.children && lastRow.children.length >= 1) {
                const tdKlasse = lastRow.children[0];
                ns.setCellEditMode(tdKlasse, id, 'klasse');
            }
        } catch (e) {
            /* ignore */
        }
    };

    ns.removeRow = function removeRow(index) {
        const row = ns.filteredData[index];
        ns.filteredData.splice(index, 1);
        if (row && row.id !== undefined && row.id !== null) {
            const ri = ns.rawData.findIndex((r) => r.id === row.id);
            if (ri >= 0) ns.rawData.splice(ri, 1);
        }
        if (typeof ns.refreshSubjectFilterUI === 'function') ns.refreshSubjectFilterUI();
        ns.invalidateTeams();
        ns.displayFilteredData();
        if (ns.currentStep === 3 && typeof ns.displayEditableData === 'function') ns.displayEditableData();
        document.getElementById('filteredRecords').textContent = ns.filteredData.length;
    };

    ns.startKursteamFromWebuntis = function startKursteamFromWebuntis() {
        ns.kursteamEntryMode = 'webuntis';
        ns.goToStep(1);
    };

    ns.startKursteamManual = function startKursteamManual() {
        ns.kursteamEntryMode = 'manual';
        ns.rawData = [];
        ns.filteredData = [];
        document.getElementById('totalRecords').textContent = '0';
        document.getElementById('uniqueSubjects').textContent = '0';
        document.getElementById('uniqueTeachers').textContent = '0';
        document.getElementById('importStats').style.display = 'none';
        const fi = document.getElementById('fileInput');
        if (fi) fi.value = '';
        ns.invalidateTeams();
        ns.goToStep(3);
        if (typeof ns.refreshSubjectFilterUI === 'function') ns.refreshSubjectFilterUI();
        document.getElementById('filterStats').style.display = 'none';
        document.getElementById('dataTableContainer').style.display = 'none';
        document.getElementById('continueBtn2').style.display = 'none';
        const tbody = document.getElementById('dataTableBody');
        if (tbody) tbody.replaceChildren();
    };

    ns.addManualDataRow = function addManualDataRow() {
        ns.openModal(
            'Unterrichtszeile hinzufügen',
            '<label for="manualKlasse">Klasse</label><input type="text" id="manualKlasse" autocomplete="off" placeholder="z. B. 5A">' +
                '<label for="manualFach">Fach</label><input type="text" id="manualFach" autocomplete="off" placeholder="z. B. D">' +
                '<label for="manualLehrer">Lehrkraft (Kürzel)</label><input type="text" id="manualLehrer" autocomplete="off" placeholder="z. B. MEI">' +
                '<label for="manualGruppe">Schülergruppe (optional)</label><input type="text" id="manualGruppe" autocomplete="off" placeholder="leer oder z. B. G1">',
            () => {
                const klasse = document.getElementById('manualKlasse').value.trim();
                const fach = document.getElementById('manualFach').value.trim();
                const lehrer = document.getElementById('manualLehrer').value.trim();
                const gruppe = document.getElementById('manualGruppe').value.trim();
                if (!klasse || !fach || !lehrer) {
                    ns.showToast('Bitte Klasse, Fach und Lehrkraft ausfüllen.');
                    return;
                }
                const id = Date.now() + Math.random();
                const row = {
                    id,
                    klasse,
                    fach,
                    lehrer,
                    gruppe: gruppe || '',
                    original: {}
                };
                ns.rawData.push(row);
                ns.filteredData.push(row);
                ns.kursteamEntryMode = ns.kursteamEntryMode === 'unset' ? 'manual' : ns.kursteamEntryMode;
                if (typeof ns.refreshSubjectFilterUI === 'function') ns.refreshSubjectFilterUI();
                ns.invalidateTeams();
                ns.closeModal();
                document.getElementById('filteredRecords').textContent = ns.filteredData.length;
                document.getElementById('filterStats').style.display = 'block';
                ns.displayFilteredData();
            }
        );
    };

    ns.resetFilters = function resetFilters() {
        ns.filteredData = [...ns.rawData];
        ns.setExcludeSubjectsInput(['ORD', 'DIR', 'KV']);
        document.getElementById('removeDuplicates').checked = true;
        if (typeof ns.refreshSubjectFilterUI === 'function') ns.refreshSubjectFilterUI();
        try {
            const btn = document.getElementById('btnApplyFilters');
            if (btn) btn.className = 'btn btn-success';
        } catch (e) {
            /* ignore */
        }
        ns.applyFilters();
    };

    ns.generateTeamNames = function generateTeamNames() {
        try {
            const btn = document.getElementById('btnGenerateTeamNames');
            if (btn) btn.className = 'btn kursteam-generate-teams-btn';
        } catch (e) {
            /* ignore */
        }

        const yearPrefix = document.getElementById('yearPrefix').value;
        const emailDomain =
            typeof window.ms365GetTeacherEmailDomainSuffix === 'function'
                ? window.ms365GetTeacherEmailDomainSuffix()
                : '@';
        const separator = document.getElementById('teamSeparator') ? document.getElementById('teamSeparator').value : ' | ';
        const pattern = document.getElementById('teamNameBuilder') ? ns.getPatternFromBuilder() : null;

        ns.teamsData = KTB.buildTeamEntriesFromRows(ns.filteredData, {
            yearPrefix,
            emailDomain,
            separator,
            pattern,
            combineClassNames: ns.combineClassNames,
            buildGruppenmailBase: ns.buildGruppenmailBase,
            INVALID_CHARS_REPLACE: ns.INVALID_CHARS_REPLACE,
            INVALID_CHARS_TEST: ns.INVALID_CHARS_TEST,
            teacherEmailMapping: ns.teacherEmailMapping
        });

        const dupCount = ns.resolveDuplicateGruppenmails(ns.teamsData);
        document.getElementById('duplicateMailAdjustments').textContent = dupCount;
        ns.teamsGenerated = true;
        ns.displayTeamsData();
    };

    ns.displayTeamsData = function displayTeamsData() {
        KTU.render(ns);
    };

    ns.editTeam = function editTeam(index) {
        const team = ns.teamsData[index];
        if (team && team.ktManualDraft) return;
        ns.openModal(
            'Team bearbeiten',
            '<label for="editName">Team-Name</label><input type="text" id="editName" value="' +
                ns.attrEscape(team.teamName) +
                '">' +
                '<label for="editMail">Gruppenmail</label><input type="text" id="editMail" value="' +
                ns.attrEscape(team.gruppenmail) +
                '">' +
                '<label for="editOwner">Besitzer</label><input type="email" id="editOwner" value="' +
                ns.attrEscape(team.besitzer) +
                '">',
            () => {
                const newName = document.getElementById('editName').value.trim();
                const newMail = document.getElementById('editMail').value.trim();
                const newOwner = document.getElementById('editOwner').value.trim();
                if (!newName || !newMail || !newOwner) {
                    ns.showToast('Bitte alle Felder ausfüllen.');
                    return;
                }
                ns.teamsData[index] = {
                    ...team,
                    teamName: newName,
                    gruppenmail: newMail,
                    besitzer: newOwner,
                    isValid: true,
                    error: null
                };
                ns.closeModal();
                ns.displayTeamsData();
            }
        );
    };

    ns.deleteTeam = function deleteTeam(index) {
        ns.confirmModal('Team löschen', 'Dieses Team wirklich aus der Liste entfernen?', () => {
            ns.teamsData.splice(index, 1);
            if (ns.teamsData.length === 0) ns.teamsGenerated = false;
            ns.displayTeamsData();
        });
    };

    ns.commitManualTeamDraftRow = function commitManualTeamDraftRow(index) {
        const team = ns.teamsData[index];
        if (!team || !team.ktManualDraft) return;
        const tbody = document.getElementById('teamsTableBody');
        const tr = tbody && tbody.querySelector('tr[data-team-draft-index="' + index + '"]');
        if (!tr) return;
        const inpName = tr.querySelector('input[data-team-draft-field="teamName"]');
        const inpMail = tr.querySelector('input[data-team-draft-field="gruppenmail"]');
        const inpOwn = tr.querySelector('input[data-team-draft-field="besitzer"]');
        const v1 = (inpName && inpName.value.trim()) || '';
        const v2 = (inpMail && inpMail.value.trim()) || '';
        const v3 = (inpOwn && inpOwn.value.trim().toLowerCase()) || '';

        if (!v1 && !v2 && !v3) {
            ns.teamsData.splice(index, 1);
            if (ns.teamsData.length === 0) ns.teamsGenerated = false;
            ns.displayTeamsData();
            return;
        }

        team.teamName = v1;
        const originalGruppenmail = v2;
        team.gruppenmail = v2.replace(ns.INVALID_CHARS_REPLACE, '');
        team.besitzer = v3;
        const filled = !!(v1 && v2 && v3);

        if (!filled) {
            team.isValid = false;
            team.error = 'Unvollständige Daten';
            team.mappingUsed = false;
            ns.displayTeamsData();
            setTimeout(() => {
                const tb2 = document.getElementById('teamsTableBody');
                const tr2 = tb2 && tb2.querySelector('tr[data-team-draft-index="' + index + '"]');
                if (!tr2) return;
                const inputs = tr2.querySelectorAll('input.kt-team-draft-input');
                const next = Array.from(inputs).find((i) => !String(i.value || '').trim()) || inputs[0];
                if (next) next.focus();
            }, 0);
            return;
        }

        const hasInvalidChars = ns.INVALID_CHARS_TEST.test(originalGruppenmail);
        team.mappingUsed = true;
        team.lehrerCode = '';
        const isValid = !hasInvalidChars && team.gruppenmail.length > 0;
        team.isValid = isValid;
        team.error = hasInvalidChars ? 'Ungültige Zeichen in Gruppenmail' : !isValid ? 'Unvollständige Daten' : null;
        delete team.ktManualDraft;
        ns.resolveDuplicateGruppenmails(ns.teamsData);
        ns.displayTeamsData();
        ns.showToast('Team übernommen.');
    };

    ns.addManualKursteamTeam = function addManualKursteamTeam() {
        ns.teamsData.push({
            teamName: '',
            gruppenmail: '',
            besitzer: '',
            isValid: false,
            error: 'Unvollständige Daten',
            originalClass: '',
            gruppe: '',
            mappingUsed: false,
            lehrerCode: '',
            mailNicknameAdjusted: false,
            ktManualDraft: true
        });
        ns.teamsGenerated = true;
        ns.displayTeamsData();
        const idx = ns.teamsData.length - 1;
        setTimeout(() => {
            const tb = document.getElementById('teamsTableBody');
            const tr = tb && tb.querySelector('tr[data-team-draft-index="' + idx + '"]');
            const inp = tr && tr.querySelector('input[data-team-draft-field="teamName"]');
            if (inp) inp.focus();
        }, 0);
    };
}

window.ms365KursteamTeamsActions = {
    mount
};
