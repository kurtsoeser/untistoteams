
const ns = (window.ms365Kursteam = window.ms365Kursteam || {});

/** Ab Schema 2: data-step entspricht der angezeigten Schrittnummer (0–8). */
const KURSTEAM_STEP_SCHEMA = 2;

function migrateKursteamStepFromStorage(step, storedSchema) {
    if (storedSchema >= KURSTEAM_STEP_SCHEMA) return step;
    const legacy = {
        2.5: 3,
        3: 4,
        4: 5,
        5: 6,
        6: 7,
        5.5: 8
    };
    return Object.prototype.hasOwnProperty.call(legacy, step) ? legacy[step] : step;
}

ns.saveStateToStorage = function saveStateToStorage() {
    try {
        const state = {
            stepSchema: KURSTEAM_STEP_SCHEMA,
            rawData: ns.rawData,
            filteredData: ns.filteredData,
            teamsData: ns.teamsData,
            teacherEmailMapping: ns.teacherEmailMapping,
            teamsGenerated: ns.teamsGenerated,
            currentStep: ns.currentStep,
            yearPrefix: document.getElementById('yearPrefix').value,
            schoolDomain:
                typeof window.ms365GetSchoolDomainNoAt === 'function'
                    ? window.ms365GetSchoolDomainNoAt()
                    : '',
            teamSeparator: document.getElementById('teamSeparator').value,
            teamNamePattern: ns.teamNamePattern || null,
            excludeSubjects: document.getElementById('excludeSubjects').value,
            removeDuplicates: document.getElementById('removeDuplicates').checked,
            kursteamEntryMode: ns.kursteamEntryMode,
            studentRosterRaw: ns.studentRosterRaw || '',
            studentRosterPreferGroup: document.getElementById('studentRosterPreferGroup')?.checked ?? true,
            studentRosterSkipCombinedClasses: document.getElementById('studentRosterSkipCombinedClasses')?.checked ?? true,
            studentRosterHideNoMatch: document.getElementById('studentRosterHideNoMatch')?.checked ?? true,
            studentRosterTeamSelection: ns.studentRosterTeamSelection || {}
        };
        localStorage.setItem(ns.STORAGE_KEY, JSON.stringify(state));
        ns.showToast('Kursteams: Zwischenstand gespeichert.');
    } catch (e) {
        ns.showToast('Speichern fehlgeschlagen: ' + e.message);
    }
};

ns.loadStateFromStorage = function loadStateFromStorage() {
    try {
        const raw = localStorage.getItem(ns.STORAGE_KEY);
        if (!raw) {
            ns.showToast('Kein gespeicherter Stand gefunden.');
            return;
        }
        const state = JSON.parse(raw);
        ns.rawData = state.rawData || [];
        ns.filteredData = state.filteredData || [];
        ns.teamsData = state.teamsData || [];
        ns.teacherEmailMapping = state.teacherEmailMapping || {};
        ns.teamsGenerated = !!state.teamsGenerated;
        ns.kursteamEntryMode =
            state.kursteamEntryMode === 'manual' || state.kursteamEntryMode === 'webuntis'
                ? state.kursteamEntryMode
                : 'unset';

        document.getElementById('yearPrefix').value = state.yearPrefix || 'SJ26';
        if (typeof window.ms365SetSchoolDomainNoAt === 'function') {
            const sd = state.schoolDomain;
            const legacy = state.emailDomain;
            if (sd !== undefined && sd !== null && String(sd).trim() !== '') {
                window.ms365SetSchoolDomainNoAt(sd);
            } else if (legacy !== undefined && legacy !== null && String(legacy).trim() !== '') {
                window.ms365SetSchoolDomainNoAt(
                    String(legacy)
                        .trim()
                        .replace(/^@+/, '')
                );
            }
        }
        document.getElementById('teamSeparator').value = state.teamSeparator !== undefined ? state.teamSeparator : ' | ';
        ns.teamNamePattern = state.teamNamePattern || null;
        if (typeof ns.renderTeamNameBuilder === 'function') ns.renderTeamNameBuilder();
        document.getElementById('excludeSubjects').value = state.excludeSubjects !== undefined ? state.excludeSubjects : 'ORD,DIR,KV';
        document.getElementById('removeDuplicates').checked = state.removeDuplicates !== false;
        if (typeof ns.refreshSubjectFilterUI === 'function') ns.refreshSubjectFilterUI();

        ns.studentRosterRaw = state.studentRosterRaw || '';
        const pref = document.getElementById('studentRosterPreferGroup');
        const skip = document.getElementById('studentRosterSkipCombinedClasses');
        const hide = document.getElementById('studentRosterHideNoMatch');
        if (pref) pref.checked = state.studentRosterPreferGroup !== false;
        if (skip) skip.checked = state.studentRosterSkipCombinedClasses !== false;
        if (hide) hide.checked = state.studentRosterHideNoMatch !== false;
        if (ns.studentRosterRaw && typeof ns.parseStudentRosterFromText === 'function') {
            ns.parseStudentRosterFromText(ns.studentRosterRaw);
        }
        ns.studentRosterTeamSelection = state.studentRosterTeamSelection || {};
        if (typeof ns.refreshStudentRosterUI === 'function') ns.refreshStudentRosterUI();

        if (ns.rawData.length) {
            document.getElementById('totalRecords').textContent = ns.rawData.length;
            document.getElementById('uniqueSubjects').textContent = new Set(ns.rawData.map(r => r.fach).filter(f => f)).size;
            document.getElementById('uniqueTeachers').textContent = new Set(ns.rawData.map(r => r.lehrer).filter(l => l)).size;
            document.getElementById('importStats').style.display = 'block';
        }
        if (ns.filteredData.length) {
            document.getElementById('filteredRecords').textContent = ns.filteredData.length;
            document.getElementById('filterStats').style.display = 'block';
            if (typeof ns.displayFilteredData === 'function') ns.displayFilteredData();
        }
        if (Object.keys(ns.teacherEmailMapping).length) {
            document.getElementById('teacherCount').textContent = Object.keys(ns.teacherEmailMapping).length;
            document.getElementById('teacherMappingInfo').style.display = 'block';
        }
        if (ns.teamsData.length && ns.teamsGenerated) {
            if (typeof ns.displayTeamsData === 'function') ns.displayTeamsData();
        }

        const hasRows = state.rawData && state.rawData.length > 0;
        const stepRaw = state.currentStep !== undefined ? state.currentStep : (hasRows ? 1 : 0);
        const step = migrateKursteamStepFromStorage(stepRaw, state.stepSchema);
        if (typeof ns.goToStep === 'function') ns.goToStep(step);

        ns.showToast('Kursteams: Stand geladen.');
    } catch (e) {
        ns.showToast('Laden fehlgeschlagen: ' + e.message);
    }
};

ns.clearStorage = function clearStorage() {
    ns.confirmModal('Lokalen Speicher löschen', 'Den gespeicherten Zwischenstand für Kursteams in diesem Browser wirklich löschen?', () => {
        try {
            localStorage.removeItem(ns.STORAGE_KEY);
            ns.showToast('Kursteams: Lokaler Speicher wurde geleert.');
        } catch (e) {
            ns.showToast('Fehler: ' + e.message);
        }
    });
};

