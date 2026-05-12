
const KM = window.ms365KursteamManualViewLogic;
window.ms365AssertModules({ KM }, 'kursteam-data-tables-ui.js');

/**
 * Import-Tabelle (Schritt 2), bearbeitbare Tabelle (Schritt 3), kleine Team-Vorschau.
 * @param {object} ns window.ms365Kursteam
 */
function mount(ns) {
    function setCellEditMode(td, rowId, field) {
        if (!td || td.dataset.editing === '1') return;
        td.dataset.editing = '1';

        const originalText = td.textContent;
        const input = document.createElement('input');
        input.type = 'text';
        input.value = originalText === '-' ? '' : originalText;
        input.style.width = '100%';
        input.style.padding = '6px 8px';
        input.style.border = '1px solid #ced4da';
        input.style.borderRadius = '6px';
        input.style.fontSize = '0.95em';
        input.style.boxSizing = 'border-box';

        td.replaceChildren(input);
        input.focus();
        input.select();

        const commit = () => {
            const val = input.value.trim();
            ns.updateDataRowField(rowId, field, val);
            td.dataset.editing = '0';
            td.textContent = val || (field === 'gruppe' ? '-' : '');
            if (typeof ns.refreshSubjectFilterUI === 'function') ns.refreshSubjectFilterUI();
        };
        const cancel = () => {
            td.dataset.editing = '0';
            td.textContent = originalText;
        };

        input.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                commit();
            } else if (e.key === 'Escape') {
                e.preventDefault();
                cancel();
            }
        });
        input.addEventListener('blur', commit);
    }

    ns.setCellEditMode = setCellEditMode;

    function getManualFilterState() {
        return {
            klasse: KM.normFilterToken(document.getElementById('manualFilterKlasse')?.value),
            fach: KM.normFilterToken(document.getElementById('manualFilterFach')?.value),
            lehrer: KM.normFilterToken(document.getElementById('manualFilterLehrer')?.value)
        };
    }

    ns.clearManualFilterInputs = function clearManualFilterInputs() {
        const k = document.getElementById('manualFilterKlasse');
        const f = document.getElementById('manualFilterFach');
        const l = document.getElementById('manualFilterLehrer');
        if (k) k.value = '';
        if (f) f.value = '';
        if (l) l.value = '';
    };

    function ensureManualSortState() {
        if (!ns.manualSort) ns.manualSort = { key: 'klasse', dir: 1 };
    }

    function updateManualSortIndicators() {
        const table = document.getElementById('editableDataTable');
        if (!table) return;
        const ths = table.querySelectorAll('th[data-sort-key]');
        ths.forEach((th) => {
            const label = th.dataset.label || th.textContent.replace(/[▲▼]\s*$/, '').trim();
            th.dataset.label = label;
            th.textContent = label;
            if (ns.manualSort && th.dataset.sortKey === ns.manualSort.key) {
                const ind = document.createElement('span');
                ind.className = 'kt-sort-indicator';
                ind.textContent = ns.manualSort.dir === -1 ? '▼' : '▲';
                th.appendChild(ind);
            }
        });
    }

    function wireManualSortAndFilterOnce() {
        if (wireManualSortAndFilterOnce._wired) return;
        wireManualSortAndFilterOnce._wired = true;

        const table = document.getElementById('editableDataTable');
        if (table) {
            table.querySelectorAll('th[data-sort-key]').forEach((th) => {
                th.addEventListener('click', () => {
                    ensureManualSortState();
                    const k = th.dataset.sortKey;
                    if (ns.manualSort.key === k) ns.manualSort.dir = ns.manualSort.dir === 1 ? -1 : 1;
                    else ns.manualSort = { key: k, dir: 1 };
                    ns.displayEditableData();
                });
            });
        }

        const onInput = () => ns.displayEditableData();
        const k = document.getElementById('manualFilterKlasse');
        const f = document.getElementById('manualFilterFach');
        const l = document.getElementById('manualFilterLehrer');
        if (k) k.addEventListener('input', onInput);
        if (f) f.addEventListener('input', onInput);
        if (l) l.addEventListener('input', onInput);

        const reset = document.getElementById('manualFilterReset');
        if (reset) {
            reset.addEventListener('click', () => {
                ns.clearManualFilterInputs();
                ns.displayEditableData();
            });
        }
    }

    ns.displayFilteredData = function displayFilteredData() {
        const tbody = document.getElementById('dataTableBody');
        tbody.replaceChildren();
        ns.filteredData.forEach((row, index) => {
            const tr = document.createElement('tr');
            const td1 = document.createElement('td');
            td1.textContent = row.klasse;
            const td2 = document.createElement('td');
            td2.textContent = row.fach;
            const td3 = document.createElement('td');
            td3.textContent = row.lehrer;
            const td4 = document.createElement('td');
            td4.textContent = row.gruppe || '-';
            const tdAction = document.createElement('td');
            tdAction.className = 'kt-action-col';
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'btn btn-small btn-danger kt-delete-btn';
            btn.textContent = 'X';
            btn.title = 'Zeile löschen';
            btn.setAttribute('aria-label', 'Zeile löschen');
            btn.addEventListener('click', () => ns.removeRow(index));
            tdAction.appendChild(btn);

            tr.append(td1, td2, td3, td4, tdAction);
            tbody.appendChild(tr);
        });
        const hasRows = ns.filteredData.length > 0;
        document.getElementById('dataTableContainer').style.display = hasRows ? 'block' : 'none';
        document.getElementById('continueBtn2').style.display = hasRows ? 'inline-block' : 'none';
    };

    ns.displayEditableData = function displayEditableData() {
        const container = document.getElementById('editableDataTableContainer');
        const tbody = document.getElementById('editableDataTableBody');
        if (!container || !tbody) return;

        wireManualSortAndFilterOnce();

        tbody.replaceChildren();
        ensureManualSortState();
        const view = KM.applyManualFiltersAndSort(ns.filteredData, ns.manualSort, getManualFilterState());
        view.forEach(({ row, index }) => {
            const tr = document.createElement('tr');

            const tdKlasse = document.createElement('td');
            tdKlasse.textContent = row.klasse || '';
            tdKlasse.addEventListener('dblclick', () => setCellEditMode(tdKlasse, row.id, 'klasse'));

            const tdFach = document.createElement('td');
            tdFach.textContent = row.fach || '';
            tdFach.addEventListener('dblclick', () => setCellEditMode(tdFach, row.id, 'fach'));

            const tdLehrer = document.createElement('td');
            tdLehrer.textContent = row.lehrer || '';
            tdLehrer.addEventListener('dblclick', () => setCellEditMode(tdLehrer, row.id, 'lehrer'));

            const tdGruppe = document.createElement('td');
            tdGruppe.textContent = row.gruppe ? row.gruppe : '-';
            tdGruppe.addEventListener('dblclick', () => setCellEditMode(tdGruppe, row.id, 'gruppe'));

            const tdAction = document.createElement('td');
            tdAction.className = 'kt-action-col';
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'btn btn-small btn-danger kt-delete-btn';
            btn.textContent = 'X';
            btn.title = 'Zeile löschen';
            btn.setAttribute('aria-label', 'Zeile löschen');
            btn.addEventListener('click', () => {
                ns.removeRow(index);
                ns.displayEditableData();
            });
            tdAction.appendChild(btn);

            tr.append(tdKlasse, tdFach, tdLehrer, tdGruppe, tdAction);
            tbody.appendChild(tr);
        });

        container.style.display = view.length ? 'block' : 'none';
        updateManualSortIndicators();
    };

    ns.displayManualTeamsPreview = function displayManualTeamsPreview() {
        const wrap = document.getElementById('manualTeamsPreviewContainer');
        const body = document.getElementById('manualTeamsPreviewBody');
        if (!wrap || !body) return;

        if (!ns.teamsGenerated || !Array.isArray(ns.teamsData) || ns.teamsData.length === 0) {
            wrap.style.display = 'none';
            body.replaceChildren();
            return;
        }

        body.replaceChildren();
        ns.teamsData.forEach((team) => {
            const tr = document.createElement('tr');
            const td1 = document.createElement('td');
            td1.textContent = team.teamName;
            const td2 = document.createElement('td');
            td2.textContent = team.gruppenmail;
            const td3 = document.createElement('td');
            td3.textContent = team.besitzer;
            const td4 = document.createElement('td');
            td4.textContent = team.isValid ? '✅' : '❌';
            tr.append(td1, td2, td3, td4);
            body.appendChild(tr);
        });
        wrap.style.display = 'block';
    };
}

window.ms365KursteamDataTablesUI = {
    mount
};
