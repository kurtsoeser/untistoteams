import * as P from './wtg-parse.js';
import { buildWtgStateSnapshot as buildWtgStateSnapshotImpl, applyWtgImportedState as applyWtgImportedStateImpl } from './wtg-state.js';
import './wtg-graph.js';
import { normStr } from '../../shared/utils/strings.js';

let wtgCurrentStep = 1;
/** @type {{ displayName: string, mailNick: string, mailNickExplicit: boolean, kind: 'team'|'group', owner: string, memberLines: string }[]} */
let wtgRows = [];
/** @type {{ displayName: string, mailNick: string, mailNickExplicit: boolean, kind: 'team'|'group' }[]} */
let wtgPreviewRows = [];
let wtgSuppressTextareaSync = false;

function showToast(msg) {
    const el = document.getElementById('toast');
    if (!el) return;
    el.textContent = msg;
    el.classList.add('show');
    clearTimeout(showToast._t);
    showToast._t = setTimeout(() => el.classList.remove('show'), 3500);
}
window.ms365ShowToast = window.ms365ShowToast || showToast;

function getPageStep() {
    const raw = document.body ? document.body.getAttribute('data-wtg-page-step') : '';
    const n = parseFloat(String(raw || '').trim());
    return Number.isFinite(n) ? n : 1;
}

function stepUrl(step) {
    const m = {
        1: 'weitere-teams-gruppen.html',
        2: 'weitere-teams-gruppen-2-besitzer.html',
        3: 'weitere-teams-gruppen-3-mitglieder.html',
        4: 'weitere-teams-gruppen-4-einstellungen.html',
        5: 'weitere-teams-gruppen-5-ausfuehren.html'
    };
    return m[step] || m[1];
}

function navigateToStep(step) {
    wtgCurrentStep = step;
    // auf Multi-Page: lieber speichern und dann navigieren
    try {
        saveWtgState();
    } catch {
        /* ignore */
    }
    if (window.location && typeof window.location.href === 'string') {
        const base = window.location.href.split('?')[0].split('#')[0];
        const next = base.replace(/[^/]*$/, stepUrl(step));
        window.location.href = next;
    }
}

function loadStoredStateSafe() {
    try {
        const raw = localStorage.getItem(WTG_STORAGE_KEY);
        if (!raw) return null;
        const obj = JSON.parse(raw);
        return obj && typeof obj === 'object' ? obj : null;
    } catch {
        return null;
    }
}

function hasRowsInState(state) {
    const rows = state && Array.isArray(state.wtgRows) ? state.wtgRows : [];
    return rows.length > 0;
}

function allOwnersSetInState(state) {
    const rows = state && Array.isArray(state.wtgRows) ? state.wtgRows : [];
    if (!rows.length) return false;
    return rows.every((r) => !!String(r && r.owner ? r.owner : '').trim());
}

function canJumpToStep(targetStep) {
    if (targetStep <= 1) return { ok: true };
    const st = loadStoredStateSafe();
    if (!st || !hasRowsInState(st)) {
        return { ok: false, reason: 'Bitte zuerst in Schritt 1 eine Liste erfassen.' };
    }
    if (targetStep >= 5 && !allOwnersSetInState(st)) {
        return { ok: false, reason: 'Bitte zuerst in Schritt 2 für alle Einträge einen Besitzer (UPN) eintragen.' };
    }
    return { ok: true };
}

function wtgStepNum(el) {
    const raw = el.getAttribute('data-wtg-step');
    const n = parseFloat(String(raw || '').trim());
    return Number.isFinite(n) ? n : NaN;
}

function goToWtgStep(step) {
    wtgCurrentStep = step;
    // Legacy Single-Page Verhalten: falls noch irgendwo alle Steps in einer Datei sind.
    const stepContents = document.querySelectorAll('.wtg-step-content');
    if (stepContents && stepContents.length > 1) {
        stepContents.forEach((el) => {
            el.classList.toggle('active', wtgStepNum(el) === step);
        });
        document.querySelectorAll('.wtg-steps .step').forEach((el) => {
            const s = wtgStepNum(el);
            el.classList.toggle('active', s === step);
            el.classList.toggle('completed', s < step);
        });
        if (step === 3) rebuildWtgMembersTableFromRows();
        if (typeof window.ms365ApplyStepProgress === 'function') {
            window.ms365ApplyStepProgress(document.querySelector('.wtg-steps'), step, [1, 2, 3, 4, 5]);
        }
        return;
    }
    // Multi-Page: navigate
    navigateToStep(step);
}

function getDomain() {
    if (typeof window.ms365GetSchoolDomainNoAt === 'function') return window.ms365GetSchoolDomainNoAt();
    return '';
}

function getPrefix() {
    const el = document.getElementById('wtgDefaultPrefix');
    const raw = (el && el.value ? el.value : '').trim();
    return raw.toLowerCase().replace(/[^a-z0-9]/g, '').slice(0, 20);
}

function maybeUpper(s) {
    const el = document.getElementById('wtgUpperNick');
    const upper = el ? !!el.checked : false;
    return upper ? s.toUpperCase() : s.toLowerCase();
}

function wtgNickDeps() {
    return { getPrefix, maybeUpper };
}

let wtgPreviewDebounce;
function scheduleWtgPreviewFromTextarea() {
    clearTimeout(wtgPreviewDebounce);
    wtgPreviewDebounce = setTimeout(() => {
        syncWtgPreviewFromTextarea();
        renderWtgPreviewTableBody();
    }, 120);
}

function scheduleWtgPreviewRowsOnly() {
    clearTimeout(wtgPreviewDebounce);
    wtgPreviewDebounce = setTimeout(() => {
        if (wtgPreviewRows.length) {
            P.recomputeWtgPreviewMailNicks(wtgPreviewRows, wtgNickDeps());
            renderWtgPreviewTableBody();
        } else {
            syncWtgPreviewFromTextarea();
            renderWtgPreviewTableBody();
        }
    }, 120);
}

function syncWtgPreviewFromTextarea() {
    const ta = document.getElementById('wtgLines');
    const text = ta ? ta.value : '';
    const { parsed } = P.parseWtgInputLines(text, wtgNickDeps());
    wtgPreviewRows = parsed.map((r) => ({
        displayName: r.displayName,
        mailNick: r.mailNick,
        mailNickExplicit: !!r.mailNickExplicit,
        kind: r.kind
    }));
}

function syncTextareaFromWtgPreviewRows() {
    if (wtgSuppressTextareaSync || !wtgPreviewRows.length) return;
    const ta = document.getElementById('wtgLines');
    if (!ta) return;
    const lines = P.serializePreviewRowsToLines(wtgPreviewRows);
    wtgSuppressTextareaSync = true;
    ta.value = lines.join('\n');
    wtgSuppressTextareaSync = false;
}

function startCellEdit(td, initialValue, onCommit) {
    const prevText = String(initialValue ?? '');
    const input = document.createElement('input');
    input.type = 'text';
    input.value = prevText;
    input.className = 'cell-editor';
    input.style.width = '100%';
    input.style.boxSizing = 'border-box';
    input.style.padding = '8px 10px';
    input.style.border = '1px solid #5e72e4';
    input.style.borderRadius = '10px';
    input.style.font = 'inherit';
    td.replaceChildren(input);
    input.focus();
    input.select();

    const commit = () => onCommit(normStr(input.value));
    const cancel = () => onCommit(prevText, { cancelled: true });
    input.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
            e.preventDefault();
            commit();
        } else if (e.key === 'Escape') {
            e.preventDefault();
            cancel();
        }
    });
    input.addEventListener('blur', () => commit());
}

function renderWtgPreviewTableBody() {
    const tbody = document.getElementById('wtgPreviewBody');
    if (!tbody) return;
    try {
        if (!wtgPreviewRows.length) {
            const ta = document.getElementById('wtgLines');
            const raw = ta ? ta.value : '';
            const nonEmpty = raw
                .split(/\r\n|\n|\r/)
                .filter((l) => l.trim() && !l.trim().startsWith('#')).length;
            tbody.innerHTML = nonEmpty
                ? '<tr><td colspan="4" style="color:#6c757d;">Keine gültigen Zeilen – Format prüfen.</td></tr>'
                : '<tr><td colspan="4" style="color:#6c757d;">Noch keine Zeilen – oben einfügen oder „+ Zeile hinzufügen“.</td></tr>';
            return;
        }
        const domain = getDomain() || '…';
        tbody.replaceChildren();
        wtgPreviewRows.forEach((r, i) => {
            const tr = document.createElement('tr');
            tr.dataset.wtgIndex = String(i);

            const td1 = document.createElement('td');
            td1.textContent = r.displayName || '';
            td1.title = 'Doppelklick zum Bearbeiten';
            td1.addEventListener('dblclick', () => {
                startCellEdit(td1, r.displayName, (next, meta) => {
                    const prev = wtgPreviewRows[i]?.displayName || '';
                    wtgPreviewRows[i].displayName = meta && meta.cancelled ? prev : next;
                    P.recomputeWtgPreviewMailNicks(wtgPreviewRows, wtgNickDeps());
                    syncTextareaFromWtgPreviewRows();
                    renderWtgPreviewTableBody();
                });
            });

            const td2 = document.createElement('td');
            const sel = document.createElement('select');
            sel.style.width = '100%';
            sel.innerHTML = '<option value="team">Team</option><option value="group">Gruppe</option>';
            sel.value = r.kind === 'team' ? 'team' : 'group';
            sel.addEventListener('change', () => {
                wtgPreviewRows[i].kind = sel.value === 'team' ? 'team' : 'group';
                syncTextareaFromWtgPreviewRows();
            });
            td2.appendChild(sel);

            const td3 = document.createElement('td');
            td3.textContent = r.mailNick || '';
            td3.style.fontFamily = 'Consolas,monospace';
            td3.style.fontSize = '0.9em';
            td3.title = 'Doppelklick zum Bearbeiten. Leer lassen = automatisch.';
            td3.addEventListener('dblclick', () => {
                startCellEdit(td3, r.mailNickExplicit ? r.mailNick : '', (next, meta) => {
                    const prevNick = wtgPreviewRows[i]?.mailNick || '';
                    const prevExplicit = !!wtgPreviewRows[i]?.mailNickExplicit;
                    const raw = meta && meta.cancelled ? (prevExplicit ? prevNick : '') : next;
                    const cleaned = maybeUpper(String(raw || '').replace(/[^A-Za-z0-9-]/g, '')).slice(0, 64);
                    wtgPreviewRows[i].mailNickExplicit = normStr(raw) !== '';
                    if (wtgPreviewRows[i].mailNickExplicit) {
                        wtgPreviewRows[i].mailNick = cleaned;
                    }
                    P.recomputeWtgPreviewMailNicks(wtgPreviewRows, wtgNickDeps());
                    syncTextareaFromWtgPreviewRows();
                    renderWtgPreviewTableBody();
                });
            });

            const td4 = document.createElement('td');
            td4.textContent = (r.mailNick || '') + '@' + domain;

            tr.append(td1, td2, td3, td4);
            tbody.appendChild(tr);
        });
    } catch (e) {
        console.error('WTG-Vorschau:', e);
        tbody.innerHTML =
            '<tr><td colspan="4" style="color:#dc3545;">Vorschau konnte nicht berechnet werden. Konsole prüfen.</td></tr>';
    }
}

function rebuildWtgOwnerTableFromRows() {
    const domain = getDomain();
    const tbody = document.getElementById('wtgOwnerBody');
    if (!tbody) return;
    tbody.replaceChildren();
    wtgRows.forEach((row, index) => {
        const tr = document.createElement('tr');
        const td1 = document.createElement('td');
        td1.textContent = row.displayName;
        const tdKind = document.createElement('td');
        tdKind.textContent = row.kind === 'team' ? 'Team' : 'Gruppe';
        const td2 = document.createElement('td');
        td2.textContent = row.mailNick + '@' + domain;
        td2.style.fontFamily = 'Consolas,monospace';
        td2.style.fontSize = '0.9em';
        const td3 = document.createElement('td');
        const inp = document.createElement('input');
        inp.type = 'email';
        inp.placeholder = 'besitzer@' + domain;
        inp.style.width = '100%';
        inp.style.padding = '8px';
        inp.value = row.owner || '';
        inp.addEventListener('input', () => {
            wtgRows[index].owner = inp.value.trim();
        });
        td3.appendChild(inp);
        tr.append(td1, tdKind, td2, td3);
        tbody.appendChild(tr);
    });
}

function parseMemberLinesText(raw) {
    const lines = String(raw || '').split(/\r\n|\n|\r/);
    const seen = new Set();
    const out = [];
    lines.forEach((line) => {
        const t = String(line || '').trim();
        if (!t || t.startsWith('#')) return;
        const key = t.toLowerCase();
        if (seen.has(key)) return;
        seen.add(key);
        out.push(t);
    });
    return out;
}

function rebuildWtgMembersTableFromRows() {
    const domain = getDomain();
    const tbody = document.getElementById('wtgMembersBody');
    if (!tbody) return;
    tbody.replaceChildren();
    wtgRows.forEach((row, index) => {
        const tr = document.createElement('tr');
        const td1 = document.createElement('td');
        td1.textContent = row.displayName;
        const tdKind = document.createElement('td');
        tdKind.textContent = row.kind === 'team' ? 'Team' : 'Gruppe';
        const td2 = document.createElement('td');
        td2.textContent = row.mailNick + '@' + domain;
        td2.style.fontFamily = 'Consolas,monospace';
        td2.style.fontSize = '0.9em';
        const td3 = document.createElement('td');
        const ta = document.createElement('textarea');
        ta.className = 'wtg-member-lines';
        ta.rows = 4;
        ta.style.width = '100%';
        ta.style.minWidth = '220px';
        ta.style.padding = '8px';
        ta.style.fontFamily = 'Consolas,monospace';
        ta.style.fontSize = '0.9em';
        ta.style.boxSizing = 'border-box';
        ta.setAttribute('autocomplete', 'off');
        ta.placeholder = 'person@' + domain;
        ta.value = row.memberLines != null ? row.memberLines : '';
        ta.addEventListener('input', () => {
            wtgRows[index].memberLines = ta.value;
        });
        td3.appendChild(ta);
        tr.append(td1, tdKind, td2, td3);
        tbody.appendChild(tr);
    });
}

function getAdminAsOwner() {
    const el = document.getElementById('wtgAdminAsOwner');
    return el ? !!el.checked : true;
}
function getDefaultVisibilityPrivate() {
    const el = document.getElementById('wtgDefaultVisibilityPrivate');
    return el ? !!el.checked : true;
}

const WTG_STORAGE_KEY = 'ms365-wtg-state-v1';

function saveWtgState() {
    try {
        const prefixEl = document.getElementById('wtgDefaultPrefix');
        const upperEl = document.getElementById('wtgUpperNick');
        const linesEl = document.getElementById('wtgLines');
        const adminEl = document.getElementById('wtgAdminAsOwner');
        const visEl = document.getElementById('wtgDefaultVisibilityPrivate');
        const state = {
            wtgCurrentStep,
            wtgRows,
            wtgDefaultPrefix: prefixEl ? prefixEl.value : '',
            wtgUpperNick: upperEl ? !!upperEl.checked : false,
            wtgAdminAsOwner: adminEl ? !!adminEl.checked : getAdminAsOwner(),
            wtgDefaultVisibilityPrivate: visEl ? !!visEl.checked : getDefaultVisibilityPrivate(),
            wtgLines: linesEl ? linesEl.value : ''
        };
        localStorage.setItem(WTG_STORAGE_KEY, JSON.stringify(state));
        showToast('Weitere Teams & Gruppen: Zwischenstand gespeichert.');
    } catch (e) {
        showToast('Speichern fehlgeschlagen: ' + e.message);
    }
}

function loadWtgState() {
    try {
        const raw = localStorage.getItem(WTG_STORAGE_KEY);
        if (!raw) {
            showToast('Kein gespeicherter Stand.');
            return;
        }
        const state = JSON.parse(raw);
        wtgCurrentStep = typeof state.wtgCurrentStep === 'number' ? state.wtgCurrentStep : 1;
        wtgRows = Array.isArray(state.wtgRows) ? state.wtgRows : [];
        const prefixEl = document.getElementById('wtgDefaultPrefix');
        if (prefixEl) prefixEl.value = state.wtgDefaultPrefix || '';
        const upperEl = document.getElementById('wtgUpperNick');
        if (upperEl) upperEl.checked = !!state.wtgUpperNick;
        const adminEl = document.getElementById('wtgAdminAsOwner');
        if (adminEl) adminEl.checked = state.wtgAdminAsOwner !== undefined ? !!state.wtgAdminAsOwner : true;
        const visEl = document.getElementById('wtgDefaultVisibilityPrivate');
        if (visEl)
            visEl.checked =
                state.wtgDefaultVisibilityPrivate !== undefined ? !!state.wtgDefaultVisibilityPrivate : true;
        const linesEl = document.getElementById('wtgLines');
        if (linesEl) linesEl.value = state.wtgLines || '';

        if (wtgRows.length) {
            rebuildWtgOwnerTableFromRows();
            rebuildWtgMembersTableFromRows();
        } else {
            const ob = document.getElementById('wtgOwnerBody');
            if (ob) ob.replaceChildren();
            const mb = document.getElementById('wtgMembersBody');
            if (mb) mb.replaceChildren();
        }
        const pageStep = getPageStep();
        // Multi-Page: nicht automatisch umleiten (User ist ja schon auf einer Seite)
        wtgCurrentStep = Math.min(Math.max(1, pageStep), 5);
        if (pageStep === 1) scheduleWtgPreviewFromTextarea();
        showToast('Weitere Teams & Gruppen: Stand geladen.');
    } catch (e) {
        showToast('Laden fehlgeschlagen: ' + e.message);
    }
}

function clearWtgState() {
    if (!confirm('Gespeicherten Zwischenstand wirklich löschen?')) return;
    try {
        localStorage.removeItem(WTG_STORAGE_KEY);
        wtgCurrentStep = 1;
        wtgRows = [];
        wtgPreviewRows = [];
        document.getElementById('wtgDefaultPrefix').value = '';
        document.getElementById('wtgUpperNick').checked = false;
        const adminEl = document.getElementById('wtgAdminAsOwner');
        if (adminEl) adminEl.checked = true;
        const visEl = document.getElementById('wtgDefaultVisibilityPrivate');
        if (visEl) visEl.checked = true;
        document.getElementById('wtgLines').value = '';
        document.getElementById('wtgParseError').style.display = 'none';
        const ob = document.getElementById('wtgOwnerBody');
        if (ob) ob.replaceChildren();
        const mb = document.getElementById('wtgMembersBody');
        if (mb) mb.replaceChildren();
        goToWtgStep(1);
        renderWtgPreviewTableBody();
        showToast('Speicher geleert.');
    } catch (e) {
        showToast('Fehler: ' + e.message);
    }
}

window.ms365SaveWtg = saveWtgState;
window.ms365LoadWtg = loadWtgState;
window.ms365ClearWtg = clearWtgState;

window.ms365GetWtgSnapshotForGraph = function () {
    return {
        rows: wtgRows.map((r) => ({
            displayName: r.displayName,
            mailNick: r.mailNick,
            kind: r.kind === 'team' ? 'team' : 'group',
            owner: r.owner,
            memberEmails: parseMemberLinesText(r.memberLines || '')
        })),
        adminAsOwner: getAdminAsOwner(),
        defaultVisibilityPrivate: getDefaultVisibilityPrivate()
    };
};

function wtgBuildStateSnapshot() {
    return buildWtgStateSnapshotImpl({
        wtgCurrentStep,
        wtgDefaultPrefix: document.getElementById('wtgDefaultPrefix')?.value ?? '',
        wtgUpperNick: !!document.getElementById('wtgUpperNick')?.checked,
        wtgAdminAsOwner: getAdminAsOwner(),
        wtgDefaultVisibilityPrivate: getDefaultVisibilityPrivate(),
        wtgLines: document.getElementById('wtgLines')?.value ?? '',
        rows: wtgRows || [],
        normStr: P.normStr
    });
}

function applyImportedWtgState(obj) {
    applyWtgImportedStateImpl(obj, {
        normStr: P.normStr,
        syncWtgPreviewFromTextarea,
        setWtgRows: (rows) => {
            wtgRows = rows;
        },
        rebuildWtgOwnerTableFromRows,
        rebuildWtgMembersTableFromRows,
        scheduleWtgPreviewRowsOnly,
        getWtgPreviewRows: () => wtgPreviewRows,
        showToast
    });
}

function downloadBlob(filename, text, mime) {
    const blob = new Blob([text], { type: mime || 'text/plain;charset=utf-8' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
    URL.revokeObjectURL(a.href);
}

function downloadJson(filename, obj) {
    downloadBlob(filename, JSON.stringify(obj, null, 2), 'application/json;charset=utf-8');
}

// UI wiring
const linesEl = document.getElementById('wtgLines');
if (linesEl) {
    linesEl.addEventListener('input', () => {
        if (wtgSuppressTextareaSync) return;
        scheduleWtgPreviewFromTextarea();
    });
    linesEl.addEventListener('paste', () => setTimeout(() => !wtgSuppressTextareaSync && scheduleWtgPreviewFromTextarea(), 0));
}
['schoolEmailDomain', 'wtgDefaultPrefix'].forEach((id) => {
    const el = document.getElementById(id);
    if (el) el.addEventListener('input', scheduleWtgPreviewRowsOnly);
});
const upperEl = document.getElementById('wtgUpperNick');
if (upperEl) upperEl.addEventListener('change', scheduleWtgPreviewRowsOnly);

const addRowBtn = document.getElementById('wtgPreviewAddRow');
if (addRowBtn) {
    addRowBtn.addEventListener('click', () => {
        wtgPreviewRows.push({ displayName: '', mailNick: '', mailNickExplicit: false, kind: 'group' });
        P.recomputeWtgPreviewMailNicks(wtgPreviewRows, wtgNickDeps());
        renderWtgPreviewTableBody();
    });
}

const btnParseAndGo2 = document.getElementById('wtgParseAndGo2');
if (btnParseAndGo2) btnParseAndGo2.addEventListener('click', () => {
    const errEl = document.getElementById('wtgParseError');
    if (errEl) errEl.style.display = 'none';

    if (!wtgPreviewRows.length) syncWtgPreviewFromTextarea();
    if (!wtgPreviewRows.length) {
        if (errEl) {
            errEl.textContent = 'Bitte mindestens eine Zeile eintragen oder in der Vorschau eine Zeile hinzufügen.';
            errEl.style.display = 'block';
        }
        return;
    }

    P.recomputeWtgPreviewMailNicks(wtgPreviewRows, wtgNickDeps());
    const rowErrors = [];
    wtgPreviewRows.forEach((r, idx) => {
        if (!(r.displayName || '').trim()) rowErrors.push('Vorschau Zeile ' + (idx + 1) + ': Anzeigename fehlt.');
        if (!(r.mailNick || '').trim()) rowErrors.push('Vorschau Zeile ' + (idx + 1) + ': Mail-Nickname fehlt.');
    });
    if (rowErrors.length) {
        if (errEl) {
            errEl.textContent = rowErrors.join('\n');
            errEl.style.display = 'block';
        }
        return;
    }

    const ownerByKey = new Map(wtgRows.map((r) => [r.displayName.toLowerCase(), r.owner]));
    const memberByKey = new Map(wtgRows.map((r) => [r.displayName.toLowerCase(), r.memberLines || '']));

    wtgRows = wtgPreviewRows.map((r) => ({
        displayName: r.displayName.trim(),
        mailNick: r.mailNick,
        mailNickExplicit: !!r.mailNickExplicit,
        kind: r.kind === 'team' ? 'team' : 'group',
        owner: ownerByKey.get(r.displayName.trim().toLowerCase()) || '',
        memberLines: memberByKey.get(r.displayName.trim().toLowerCase()) || ''
    }));

    // Seite 1 → Seite 2
    navigateToStep(2);
});

const btnGoTo3 = document.getElementById('wtgGoTo3');
if (btnGoTo3) btnGoTo3.addEventListener('click', () => navigateToStep(3));
const backMem = document.getElementById('wtgMemberBack');
if (backMem) backMem.addEventListener('click', () => navigateToStep(2));
const nextMem = document.getElementById('wtgMemberNext');
if (nextMem) nextMem.addEventListener('click', () => navigateToStep(4));
const btnGoTo5 = document.getElementById('wtgGoTo5');
if (btnGoTo5) btnGoTo5.addEventListener('click', () => {
    const missing = wtgRows.filter((r) => !r.owner);
    if (missing.length) {
        showToast('Bitte für alle Einträge einen Besitzer (UPN) eintragen.');
        navigateToStep(2);
        return;
    }
    navigateToStep(5);
});

// step header keyboard support
document.querySelectorAll('.wtg-steps .step').forEach((el) => {
    el.setAttribute('tabindex', '0');
    el.addEventListener('click', () => {
        const s = wtgStepNum(el);
        if (!Number.isFinite(s)) return;
        // Rückwärts immer erlauben; vorwärts nur, wenn Voraussetzungen erfüllt sind.
        const current = getPageStep();
        if (s <= current) {
            goToWtgStep(s);
            return;
        }
        const check = canJumpToStep(s);
        if (!check.ok) {
            showToast(check.reason || 'Bitte zuerst die vorherigen Schritte abschließen.');
            return;
        }
        goToWtgStep(s);
    });
    el.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' || e.key === ' ') {
            e.preventDefault();
            el.click();
        }
    });
});

// bottom toolbar wiring
const btnSave = document.getElementById('btnSaveState');
if (btnSave) btnSave.addEventListener('click', saveWtgState);
const btnLoad = document.getElementById('btnLoadState');
if (btnLoad) btnLoad.addEventListener('click', loadWtgState);
const btnClear = document.getElementById('btnClearStorage');
if (btnClear) btnClear.addEventListener('click', clearWtgState);

const btnExport = document.getElementById('btnExportWtgJson');
if (btnExport) {
    btnExport.addEventListener('click', () => {
        const snap = wtgBuildStateSnapshot();
        downloadJson(`weitere-teams-gruppen-export-${new Date().toISOString().slice(0, 10)}.json`, snap);
        showToast('Export: JSON erstellt.');
    });
}
const fileJson = document.getElementById('wtgImportJsonFile');
const btnImport = document.getElementById('btnImportWtgJson');
if (btnImport && fileJson) {
    btnImport.addEventListener('click', () => fileJson.click());
    fileJson.addEventListener('change', async (e) => {
        const f = e.target.files && e.target.files[0];
        if (!f) return;
        try {
            const text = await f.text();
            const obj = JSON.parse(text);
            applyImportedWtgState(obj);
        } catch (err) {
            showToast('Import: JSON fehlgeschlagen: ' + (err?.message || String(err)));
        } finally {
            fileJson.value = '';
        }
    });
}

// init
const stepsInit = document.querySelector('.wtg-steps');
if (stepsInit && typeof window.ms365ApplyStepProgress === 'function') {
    const s = getPageStep();
    wtgCurrentStep = s;
    window.ms365ApplyStepProgress(stepsInit, s, [1, 2, 3, 4, 5]);
}
// Page init
const pageStep = getPageStep();
if (pageStep === 1) {
    scheduleWtgPreviewFromTextarea();
} else {
    loadWtgState();
}

