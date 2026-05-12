import { getEl } from '../../shared/utils/dom.js';

const GRAPH_READ_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Group.Read.All'
];
const GRAPH_WRITE_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Group.ReadWrite.All'
];

function getCore() {
    const c = window.ms365LeereGruppenCore;
    if (!c) {
        throw new Error('Skript "leere-gruppen-core.js" wurde nicht (vor) geladen.');
    }
    return c;
}

function bind() {
    const core = getCore();
    const {
        compareDe,
        escapeHtml,
        rowsToCsv,
        getGraphToken,
        graphRequest,
        fetchAllPages,
        fetchCount,
        runPool,
        kindBadgesHtml,
        buildRow,
        buildGroupsListInitialPath
    } = core;

    const progressEl = getEl('lgrProgress');
    const tbody = getEl('lgrTbody');
    const btn = getEl('lgrBtnRun');
    const btnCsv = getEl('lgrBtnCsv');
    const searchInput = getEl('lgrSearch');
    const filterKind = getEl('lgrFilterKind');
    const filterProblem = getEl('lgrFilterProblem');
    const resultCount = getEl('lgrResultCount');

    const selectAllCb = getEl('lgrSelectAll');
    const bulkBar = getEl('lgrBulkBar');
    const bulkCount = getEl('lgrBulkCount');
    const bulkHint = getEl('lgrBulkHint');
    const bulkClear = getEl('lgrBulkClear');
    const bulkDelete = getEl('lgrBulkDelete');

    const delModal = getEl('lgrDeleteModal');
    const delDisplay = getEl('lgrDelDisplay');
    const delMail = getEl('lgrDelMail');
    const delConfirm = getEl('lgrDelConfirm');
    const delExecute = getEl('lgrDelExecute');
    const delError = getEl('lgrDelError');

    const bulkModal = getEl('lgrBulkDeleteModal');
    const bulkDelTitleText = getEl('lgrBulkDelTitleText');
    const bulkDelCountText = getEl('lgrBulkDelCountText');
    const bulkDelList = getEl('lgrBulkDelList');
    const bulkDelConfirm = getEl('lgrBulkDelConfirm');
    const bulkDelExecute = getEl('lgrBulkDelExecute');
    const bulkDelCancel = getEl('lgrBulkDelCancel');
    const bulkDelError = getEl('lgrBulkDelError');
    const bulkDelProgressWrap = getEl('lgrBulkDelProgressWrap');
    const bulkDelProgressBar = getEl('lgrBulkDelProgressBar');
    const bulkDelProgressText = getEl('lgrBulkDelProgressText');

    /** @type {Array<object>} */
    let lastRows = [];
    let viewRows = [];
    const selectedIds = new Set();
    const sortState = { key: 'displayName', dir: 1 };
    let currentDelete = null;
    let bulkInProgress = false;

    function setProgress(on, text) {
        if (!progressEl) return;
        progressEl.style.display = on ? '' : 'none';
        if (text) progressEl.textContent = String(text);
    }

    function setFilterControlsEnabled(on) {
        if (searchInput) searchInput.disabled = !on;
        if (filterKind) filterKind.disabled = !on;
        if (filterProblem) filterProblem.disabled = !on;
        if (selectAllCb) selectAllCb.disabled = !on;
    }

    function matchesKindFilter(row, kind) {
        switch (kind) {
            case 'm365':
                return !!row.isUnified;
            case 'team':
                return !!row.isTeam;
            case 'security':
                return !!row.isSecurity;
            case 'mail':
                return !!row.isMail;
            case 'other':
                return !row.isUnified && !row.isTeam && !row.isSecurity && !row.isMail;
            default:
                return true;
        }
    }

    function matchesProblemFilter(row, mode) {
        const noOwners = row.owners === 0;
        const noMembers = row.members === 0;
        switch (mode) {
            case 'problem':
                return noOwners || noMembers;
            case 'no-owners':
                return noOwners;
            case 'no-members':
                return noMembers;
            case 'empty':
                return noOwners && noMembers;
            default:
                return true;
        }
    }

    function applyFiltersAndSort() {
        const q = String(searchInput?.value || '').trim().toLowerCase();
        const kind = String(filterKind?.value || 'all');
        const problem = String(filterProblem?.value || 'all');

        viewRows = lastRows.filter((r) => {
            if (q) {
                const hay =
                    (r.displayName || '') +
                    ' ' +
                    (r.mail || '') +
                    ' ' +
                    (r.kind || '') +
                    ' ' +
                    (r.flags || '');
                if (hay.toLowerCase().indexOf(q) === -1) return false;
            }
            if (!matchesKindFilter(r, kind)) return false;
            if (!matchesProblemFilter(r, problem)) return false;
            return true;
        });

        const key = sortState.key;
        const dir = sortState.dir;
        viewRows.sort((a, b) => {
            if (key === 'owners' || key === 'members') {
                return ((a[key] ?? 0) - (b[key] ?? 0)) * dir;
            }
            return compareDe(a[key], b[key]) * dir;
        });

        renderRows();
        updateSortIndicators();
    }

    function updateSortIndicators() {
        const ths = document.querySelectorAll('table.lgr-table th.is-sortable');
        ths.forEach((th) => {
            const label = th.dataset.label || th.textContent.replace(/[▲▼]\s*$/, '').trim();
            th.dataset.label = label;
            th.textContent = label;
            if (th.dataset.sortKey === sortState.key) {
                const ind = document.createElement('span');
                ind.className = 'lgr-sort-ind';
                ind.textContent = sortState.dir === -1 ? '▼' : '▲';
                th.appendChild(ind);
            }
        });
    }

    function renderRows() {
        if (!tbody) return;
        tbody.replaceChildren();

        if (resultCount) {
            if (!lastRows.length) {
                resultCount.textContent = '';
            } else if (viewRows.length === lastRows.length) {
                resultCount.textContent = lastRows.length + ' Gruppen';
            } else {
                resultCount.textContent = viewRows.length + ' von ' + lastRows.length + ' Gruppen';
            }
        }

        if (!viewRows.length) {
            if (lastRows.length) {
                const tr = document.createElement('tr');
                tr.innerHTML =
                    '<td colspan="8" class="muted" style="padding:14px 10px;">Keine Treffer für die aktuellen Filter/Suche.</td>';
                tbody.appendChild(tr);
            }
            updateSelectionUi();
            return;
        }

        const frag = document.createDocumentFragment();
        for (const r of viewRows) {
            const tr = document.createElement('tr');
            tr.dataset.groupId = r.id;
            if (selectedIds.has(r.id)) tr.classList.add('is-selected');

            const cbHtml =
                '<input type="checkbox" data-lgr-sel="row" aria-label="Auswählen"' +
                (selectedIds.has(r.id) ? ' checked' : '') +
                '>';

            const ownersTxt = r.owners < 0 ? '–' : String(r.owners);
            const membersTxt = r.members < 0 ? '–' : String(r.members);

            tr.innerHTML =
                '<td class="lgr-col-select">' +
                '<label class="lgr-checkbox-label">' +
                cbHtml +
                '</label>' +
                '</td>' +
                '<td>' +
                escapeHtml(r.displayName) +
                '</td>' +
                '<td>' +
                escapeHtml(r.mail) +
                '</td>' +
                '<td>' +
                kindBadgesHtml(r) +
                '</td>' +
                '<td class="num">' +
                ownersTxt +
                '</td>' +
                '<td class="num">' +
                membersTxt +
                '</td>' +
                '<td>' +
                escapeHtml(r.flags) +
                '</td>' +
                '<td class="lgr-col-actions">' +
                '<button type="button" class="btn btn-danger lgr-row-delete" data-lgr-act="del" title="Gruppe löschen" aria-label="Gruppe löschen">' +
                '<i class="bi bi-trash"></i>' +
                '</button>' +
                '</td>';
            frag.appendChild(tr);
        }
        tbody.appendChild(frag);
        updateSelectionUi();
    }

    function visibleSelectableIds() {
        return viewRows.map((r) => r.id);
    }

    function updateSelectionUi() {
        const total = selectedIds.size;
        const visIds = visibleSelectableIds();
        const visSelected = visIds.filter((id) => selectedIds.has(id)).length;

        if (selectAllCb) {
            if (!visIds.length) {
                selectAllCb.checked = false;
                selectAllCb.indeterminate = false;
                selectAllCb.disabled = !lastRows.length;
            } else {
                selectAllCb.disabled = false;
                if (visSelected === 0) {
                    selectAllCb.checked = false;
                    selectAllCb.indeterminate = false;
                } else if (visSelected === visIds.length) {
                    selectAllCb.checked = true;
                    selectAllCb.indeterminate = false;
                } else {
                    selectAllCb.checked = false;
                    selectAllCb.indeterminate = true;
                }
            }
        }

        if (bulkBar) {
            if (total === 0) bulkBar.setAttribute('hidden', '');
            else bulkBar.removeAttribute('hidden');
        }
        if (bulkCount) bulkCount.textContent = String(total);
        if (bulkHint) {
            const hiddenSelected = total - visSelected;
            if (total === 0) bulkHint.textContent = '';
            else if (hiddenSelected > 0) {
                bulkHint.textContent =
                    '(' +
                    visSelected +
                    ' aktuell sichtbar, ' +
                    hiddenSelected +
                    ' nicht im aktuellen Filter angezeigt)';
            } else {
                bulkHint.textContent =
                    visSelected === 1 ? '(1 sichtbarer Treffer)' : '(' + visSelected + ' sichtbare Treffer)';
            }
        }
        if (bulkDelete) bulkDelete.disabled = total === 0 || bulkInProgress;
        if (bulkClear) bulkClear.disabled = total === 0 || bulkInProgress;
    }

    function toggleRowSelection(id, on) {
        if (!id) return;
        const tr = tbody?.querySelector(
            'tr[data-group-id="' + (window.CSS && CSS.escape ? CSS.escape(id) : id) + '"]'
        );
        if (on) selectedIds.add(id);
        else selectedIds.delete(id);
        if (tr) tr.classList.toggle('is-selected', on);
        updateSelectionUi();
    }

    function selectAllVisible(on) {
        const ids = visibleSelectableIds();
        if (on) ids.forEach((id) => selectedIds.add(id));
        else ids.forEach((id) => selectedIds.delete(id));
        const rows = tbody?.querySelectorAll('tr[data-group-id]') || [];
        rows.forEach((tr) => {
            const id = tr.dataset.groupId;
            const cb = tr.querySelector('input[data-lgr-sel="row"]');
            const sel = selectedIds.has(id);
            if (cb) cb.checked = sel;
            tr.classList.toggle('is-selected', sel);
        });
        updateSelectionUi();
    }

    function clearSelection() {
        selectedIds.clear();
        const rows = tbody?.querySelectorAll('tr[data-group-id]') || [];
        rows.forEach((tr) => {
            const cb = tr.querySelector('input[data-lgr-sel="row"]');
            if (cb) cb.checked = false;
            tr.classList.remove('is-selected');
        });
        updateSelectionUi();
    }

    document.querySelectorAll('table.lgr-table th.is-sortable').forEach((th) => {
        th.addEventListener('click', () => {
            if (!lastRows.length) return;
            const k = th.dataset.sortKey;
            if (!k) return;
            if (sortState.key === k) sortState.dir = sortState.dir === 1 ? -1 : 1;
            else {
                sortState.key = k;
                sortState.dir = 1;
            }
            applyFiltersAndSort();
        });
    });

    let searchTimer = null;
    searchInput?.addEventListener('input', () => {
        if (searchTimer) clearTimeout(searchTimer);
        searchTimer = setTimeout(applyFiltersAndSort, 120);
    });
    filterKind?.addEventListener('change', applyFiltersAndSort);
    filterProblem?.addEventListener('change', applyFiltersAndSort);

    tbody?.addEventListener('click', (e) => {
        const button = e.target.closest('button[data-lgr-act]');
        if (!button) return;
        const tr = button.closest('tr');
        const id = tr && tr.dataset.groupId;
        if (!id) return;
        const row = lastRows.find((x) => x.id === id);
        if (!row) return;
        if (button.getAttribute('data-lgr-act') === 'del') openSingleDelete(row);
    });
    tbody?.addEventListener('change', (e) => {
        const cb = e.target.closest('input[data-lgr-sel="row"]');
        if (!cb) return;
        const tr = cb.closest('tr');
        const id = tr && tr.dataset.groupId;
        if (!id) return;
        toggleRowSelection(id, !!cb.checked);
    });

    selectAllCb?.addEventListener('change', () => selectAllVisible(!!selectAllCb.checked));
    bulkClear?.addEventListener('click', clearSelection);
    bulkDelete?.addEventListener('click', openBulkDelete);

    function openModal(modal) {
        if (modal) modal.classList.add('is-open');
    }
    function closeModal(modal) {
        if (modal) modal.classList.remove('is-open');
    }

    document.querySelectorAll('[data-lgr-close]').forEach((el) => {
        el.addEventListener('click', () => {
            const which = el.getAttribute('data-lgr-close');
            if (which === 'delete') closeModal(delModal);
            else if (which === 'bulk') {
                if (!bulkInProgress) closeModal(bulkModal);
            }
        });
    });
    [delModal, bulkModal].forEach((m) => {
        if (!m) return;
        m.addEventListener('click', (e) => {
            if (e.target !== m) return;
            if (m === bulkModal && bulkInProgress) return;
            closeModal(m);
        });
    });
    document.addEventListener('keydown', (e) => {
        if (e.key !== 'Escape') return;
        if (delModal && delModal.classList.contains('is-open')) closeModal(delModal);
        else if (bulkModal && bulkModal.classList.contains('is-open') && !bulkInProgress) closeModal(bulkModal);
    });

    function openSingleDelete(row) {
        currentDelete = row;
        if (delDisplay) delDisplay.value = row.displayName || '';
        if (delMail) {
            const detail = [row.mail, row.kind].filter(Boolean).join(' · ');
            delMail.value = detail || '–';
        }
        if (delConfirm) delConfirm.value = '';
        if (delError) delError.textContent = '';
        if (delExecute) delExecute.disabled = true;
        openModal(delModal);
        setTimeout(() => delConfirm?.focus(), 30);
    }

    delConfirm?.addEventListener('input', () => {
        if (!currentDelete) return;
        const target = String(currentDelete.displayName || '').trim();
        const typed = String(delConfirm.value || '').trim();
        if (delExecute) {
            delExecute.disabled = !target || typed.toLowerCase() !== target.toLowerCase();
        }
    });

    async function executeSingleDelete() {
        if (!currentDelete) return;
        const row = currentDelete;
        if (delExecute) delExecute.disabled = true;
        if (delError) delError.textContent = 'Lösche …';
        try {
            const token = await getGraphToken(GRAPH_WRITE_SCOPES);
            const res = await graphRequest(
                'DELETE',
                '/groups/' + encodeURIComponent(row.id),
                token,
                undefined,
                undefined
            );
            if (!res.ok && res.status !== 204) {
                const text = await res.text();
                let msg = text;
                try {
                    const j = JSON.parse(text);
                    msg = (j && j.error && (j.error.message || JSON.stringify(j.error))) || text;
                } catch {
                    /* keep msg */
                }
                throw new Error('HTTP ' + res.status + ': ' + msg);
            }
            removeRowsLocally([row.id]);
            if (delError) delError.textContent = '';
            closeModal(delModal);
        } catch (e) {
            if (delError) delError.textContent = 'Fehler beim Löschen: ' + (e && e.message ? e.message : String(e));
            if (delExecute) delExecute.disabled = false;
        }
    }
    delExecute?.addEventListener('click', executeSingleDelete);

    function openBulkDelete() {
        if (selectedIds.size === 0) return;
        const rows = lastRows.filter((r) => selectedIds.has(r.id));
        if (bulkDelTitleText)
            bulkDelTitleText.textContent =
                rows.length === 1 ? '1 Gruppe löschen' : rows.length + ' Gruppen löschen';
        if (bulkDelCountText) bulkDelCountText.textContent = String(rows.length);
        if (bulkDelList) {
            const items = rows
                .map(
                    (r) =>
                        '<li>' +
                        kindBadgesHtml(r) +
                        ' <strong>' +
                        escapeHtml(r.displayName) +
                        '</strong>' +
                        (r.mail ? ' <span class="muted">&lt;' + escapeHtml(r.mail) + '&gt;</span>' : '') +
                        '</li>'
                )
                .join('');
            bulkDelList.innerHTML = '<ul>' + items + '</ul>';
        }
        if (bulkDelConfirm) bulkDelConfirm.value = '';
        if (bulkDelExecute) bulkDelExecute.disabled = true;
        if (bulkDelError) bulkDelError.textContent = '';
        if (bulkDelProgressWrap) bulkDelProgressWrap.style.display = 'none';
        if (bulkDelProgressBar) bulkDelProgressBar.style.width = '0%';
        if (bulkDelProgressText) bulkDelProgressText.textContent = '0 / ' + rows.length;
        openModal(bulkModal);
        setTimeout(() => bulkDelConfirm?.focus(), 30);
    }

    bulkDelConfirm?.addEventListener('input', () => {
        const ok = String(bulkDelConfirm.value || '').trim() === 'LÖSCHEN';
        if (bulkDelExecute) bulkDelExecute.disabled = !ok || bulkInProgress;
    });

    async function executeBulkDelete() {
        const ids = Array.from(selectedIds);
        const rows = lastRows.filter((r) => ids.indexOf(r.id) !== -1);
        if (!rows.length) return;

        bulkInProgress = true;
        if (bulkDelExecute) bulkDelExecute.disabled = true;
        if (bulkDelCancel) bulkDelCancel.disabled = true;
        if (bulkDelError) bulkDelError.textContent = '';
        if (bulkDelProgressWrap) bulkDelProgressWrap.style.display = '';
        if (bulkDelProgressText) bulkDelProgressText.textContent = '0 / ' + rows.length;
        if (bulkDelProgressBar) bulkDelProgressBar.style.width = '0%';

        const errors = [];
        const okIds = [];
        let token;
        try {
            token = await getGraphToken(GRAPH_WRITE_SCOPES);
        } catch (e) {
            if (bulkDelError)
                bulkDelError.textContent = 'Anmeldung fehlgeschlagen: ' + (e && e.message ? e.message : String(e));
            bulkInProgress = false;
            if (bulkDelCancel) bulkDelCancel.disabled = false;
            if (bulkDelExecute) bulkDelExecute.disabled = false;
            return;
        }

        let done = 0;
        for (const r of rows) {
            try {
                const res = await graphRequest(
                    'DELETE',
                    '/groups/' + encodeURIComponent(r.id),
                    token,
                    undefined,
                    undefined
                );
                if (res.ok || res.status === 204) {
                    okIds.push(r.id);
                } else {
                    const text = await res.text();
                    let msg = text;
                    try {
                        const j = JSON.parse(text);
                        msg = (j && j.error && (j.error.message || JSON.stringify(j.error))) || text;
                    } catch {
                        /* keep msg */
                    }
                    errors.push({ row: r, message: 'HTTP ' + res.status + ': ' + msg });
                }
            } catch (e) {
                errors.push({ row: r, message: (e && e.message) || String(e) });
            }
            done++;
            if (bulkDelProgressText) bulkDelProgressText.textContent = done + ' / ' + rows.length;
            if (bulkDelProgressBar)
                bulkDelProgressBar.style.width = Math.round((done / rows.length) * 100) + '%';
        }

        removeRowsLocally(okIds);

        bulkInProgress = false;
        if (bulkDelCancel) bulkDelCancel.disabled = false;

        if (!errors.length) {
            closeModal(bulkModal);
            setProgress(true, okIds.length + ' Gruppe(n) gelöscht.');
            setTimeout(() => setProgress(false, ''), 2800);
        } else {
            if (bulkDelExecute) bulkDelExecute.disabled = false;
            const lines = errors
                .slice(0, 6)
                .map((e) => '• ' + (e.row.displayName || e.row.id) + ': ' + e.message)
                .join('\n');
            const more = errors.length > 6 ? '\n… +' + (errors.length - 6) + ' weitere Fehler.' : '';
            if (bulkDelError) {
                bulkDelError.textContent =
                    okIds.length +
                    ' erfolgreich, ' +
                    errors.length +
                    ' fehlgeschlagen:\n' +
                    lines +
                    more;
                bulkDelError.style.whiteSpace = 'pre-wrap';
            }
        }
    }
    bulkDelExecute?.addEventListener('click', executeBulkDelete);

    function removeRowsLocally(ids) {
        if (!ids || !ids.length) return;
        const idSet = new Set(ids);
        lastRows = lastRows.filter((r) => !idSet.has(r.id));
        ids.forEach((id) => selectedIds.delete(id));
        applyFiltersAndSort();
        btnCsv.disabled = !lastRows.length;
    }

    btnCsv?.addEventListener('click', () => {
        const rows = viewRows.length ? viewRows : lastRows;
        if (!rows.length) return;
        const cols = [
            { label: 'Anzeigename', value: (r) => r.displayName },
            { label: 'E-Mail', value: (r) => r.mail },
            { label: 'Typ', value: (r) => r.kind },
            { label: 'Besitzer', value: (r) => (r.owners < 0 ? '' : r.owners) },
            { label: 'Mitglieder', value: (r) => (r.members < 0 ? '' : r.members) },
            { label: 'Hinweise', value: (r) => r.flags }
        ];
        const csv = rowsToCsv(rows, cols);
        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = 'leere-gruppen-report.csv';
        a.click();
        setTimeout(() => URL.revokeObjectURL(a.href), 4000);
    });

    btn?.addEventListener('click', async () => {
        const scopeMode = String(getEl('lgrScope')?.value || 'unified');
        btn.disabled = true;
        btnCsv.disabled = true;
        lastRows = [];
        viewRows = [];
        selectedIds.clear();
        setFilterControlsEnabled(false);
        renderRows();
        updateSelectionUi();
        setProgress(true, 'Gruppen werden geladen …');

        try {
            const token = await getGraphToken(GRAPH_READ_SCOPES);
            const { path, headers } = buildGroupsListInitialPath(scopeMode);
            const groups = await fetchAllPages(
                token,
                path,
                (p) => setProgress(true, 'Gruppen laden … Seite ' + p.page + ', ' + p.loaded + ' Gruppen'),
                headers
            );

            const tasks = groups.map((g) => async () => {
                const id = String(g.id || '');
                if (!id) return null;
                const [owners, members] = await Promise.all([
                    fetchCount(token, id, 'owners'),
                    fetchCount(token, id, 'members')
                ]);
                return buildRow(g, owners, members);
            });

            setProgress(true, 'Besitzer/Mitglieder zählen … 0 / ' + tasks.length);
            const concurrency = 4;
            let done = 0;
            const enriched = await runPool(
                tasks.map((fn) => async () => {
                    const row = await fn();
                    done++;
                    if (done % 5 === 0 || done === tasks.length) {
                        setProgress(true, 'Besitzer/Mitglieder zählen … ' + done + ' / ' + tasks.length);
                    }
                    return row;
                }),
                concurrency
            );

            lastRows = enriched.filter(Boolean);
            lastRows.sort((a, b) => compareDe(a.displayName, b.displayName));
            setFilterControlsEnabled(true);
            applyFiltersAndSort();
            btnCsv.disabled = !lastRows.length;
            setProgress(true, 'Fertig: ' + lastRows.length + ' Gruppen ausgewertet.');
            setTimeout(() => setProgress(false, ''), 2400);
        } catch (e) {
            setProgress(true, 'Fehler: ' + (e && e.message ? e.message : String(e)));
        } finally {
            btn.disabled = false;
        }
    });

    updateSortIndicators();
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', bind);
else bind();
