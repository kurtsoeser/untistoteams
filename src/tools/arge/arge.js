import * as P from './arge-parse.js';
import { buildStateSnapshot as argeBuildStateSnapshotImpl, applyImportedState as argeApplyImportedStateImpl } from './arge-state.js';
import { exportArgeCsv as argeExportCsvImpl, parseArgeCsvToRows } from './arge-csv.js';
import { buildStandaloneArgePs1 as buildStandaloneArgePs1Impl } from './arge-standalone-ps1.js';
import './arge-graph.js';

let argeCurrentStep = 1;
/** @type {{ displayName: string, mailNick: string, owner: string, description: string, memberLines: string }[]} */
let argeRows = [];
/** Schritt 1: bearbeitbare Vorschau (wie Jahrgangsgruppen) */
/** @type {{ displayName: string, mailNick: string, owner: string, description: string, technicalSlug: string, mailNickExplicit: boolean }[]} */
let argePreviewRows = [];
let argeSuppressTextareaSync = false;

const panelW = document.getElementById('panelWebuntis');
const panelJ = document.getElementById('panelJahrgang');
const panelA = document.getElementById('panelArge');
const panelG = document.getElementById('panelGruppenPolicy');

const btnModeW = document.getElementById('modeWebuntis');
const btnModeJ = document.getElementById('modeJahrgang');
const btnModeA = document.getElementById('modeArge');
const btnModeG = document.getElementById('modeGruppenPolicy');

const normStr = P.normStr;
const normSubjectKey = P.normSubjectKey;
const buildSubjectCodeFromName = P.buildSubjectCodeFromName;
const getSubjectCodeForDisplayName = P.getSubjectCodeForDisplayName;
const toNickBaseFromName = P.toNickBaseFromName;
const subjectForSlug = P.subjectForSlug;
const looksLikeSubjectCode = P.looksLikeSubjectCode;
const normSubjectCode = P.normSubjectCode;
const parseArgeLine = P.parseArgeLine;
const displayNameFromSubjectLine = P.displayNameFromSubjectLine;

function showToast(msg) {
    const el = document.getElementById('toast');
    if (!el) return;
    el.textContent = msg;
    el.classList.add('show');
    clearTimeout(showToast._t);
    showToast._t = setTimeout(() => el.classList.remove('show'), 3500);
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

    const commit = () => {
        const next = normStr(input.value);
        onCommit(next);
    };
    const cancel = () => {
        onCommit(prevText, { cancelled: true });
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
    input.addEventListener('blur', () => commit());
}

function setMode(which) {
    const w = which === 'webuntis';
    const j = which === 'jahrgang';
    const a = which === 'arge';
    const g = which === 'gruppenerstellung';
    if (panelW) panelW.style.display = w ? '' : 'none';
    if (panelJ) panelJ.style.display = j ? '' : 'none';
    if (panelA) panelA.style.display = a ? '' : 'none';
    if (panelG) panelG.style.display = g ? '' : 'none';
    if (btnModeW) btnModeW.classList.toggle('btn-success', w);
    if (btnModeJ) btnModeJ.classList.toggle('btn-success', j);
    if (btnModeA) btnModeA.classList.toggle('btn-success', a);
    if (btnModeG) btnModeG.classList.toggle('btn-success', g);
    const sdb = document.getElementById('schoolDomainBar');
    if (sdb) sdb.style.display = g ? 'none' : '';
    if (a) {
        adoptArgeLinesFromTenantSettingsIfEmpty();
        scheduleArgePreviewRefresh();
    }
}

if (btnModeW) btnModeW.addEventListener('click', () => setMode('webuntis'));
if (btnModeJ) btnModeJ.addEventListener('click', () => setMode('jahrgang'));
if (btnModeA) btnModeA.addEventListener('click', () => setMode('arge'));
if (btnModeG) btnModeG.addEventListener('click', () => setMode('gruppenerstellung'));

function applyInitialModeFromUrl() {
    try {
        const mode = new URLSearchParams(window.location.search).get('mode');
        if (!mode) return;
        if (mode.toLowerCase() === 'arge') setMode('arge');
        else if (mode.toLowerCase() === 'gruppenerstellung' || mode.toLowerCase() === 'grouppolicy')
            setMode('gruppenerstellung');
    } catch {
        // ignore
    }
}

function argeStepNum(el) {
    const raw = el.getAttribute('data-arge-step');
    const n = parseFloat(String(raw || '').trim());
    return Number.isFinite(n) ? n : NaN;
}

function goToArgeStep(step) {
    argeCurrentStep = step;
    document.querySelectorAll('.arge-step-content').forEach(el => {
        el.classList.toggle('active', argeStepNum(el) === step);
    });
    document.querySelectorAll('.arge-steps .step').forEach(el => {
        const s = argeStepNum(el);
        el.classList.toggle('active', s === step);
        el.classList.toggle('completed', s < step);
    });
    if (step === 1) {
        if (argeRows.length) {
            argePreviewRows = argeRows.map(r => ({
                displayName: r.displayName,
                mailNick: r.mailNick,
                owner: '',
                description: r.description,
                technicalSlug: toNickBaseFromName(subjectForSlug(r.displayName)),
                mailNickExplicit: true,
                subjectName: normStr(subjectForSlug(r.displayName)),
                subjectCode: normStr(r.subjectCode || '')
            }));
            renderArgePreviewTableBody();
        } else {
            scheduleArgePreviewFromTextarea();
        }
    }
    if (step === 3) {
        rebuildArgeMembersTableFromRows();
    }
    if (typeof window.ms365ApplyStepProgress === 'function') {
        window.ms365ApplyStepProgress(document.querySelector('.arge-steps'), step, [1, 2, 3, 4, 5]);
    }
}

function getDomain() {
    if (typeof window.ms365GetSchoolDomainNoAt === 'function') {
        return window.ms365GetSchoolDomainNoAt();
    }
    return '';
}

function getPrefix() {
    const el = document.getElementById('argeDefaultPrefix');
    const raw = (el && el.value ? el.value : '').trim();
    // ARGE soll standardmäßig "arge-<kürzel>" ergeben, auch wenn das Feld leer ist
    if (!raw) return 'arge';
    return raw.toLowerCase().replace(/[^a-z0-9]/g, '');
}

/**
 * Leere ARGE-Liste aus den zentralen Schul‑Einstellungen füllen.
 * Bevorzugt den Tab „ARGE“ (Stammdaten `arges`); nur wenn dort nichts liegt, Fallback auf Fächer (`subjects`).
 * Nur `Kürzel;Bezeichnung` — keine dritte Spalte (Fächer-Zuordnung), damit die ARGE‑Zeilenparser nicht mit Mail‑Nicknames kollidieren.
 */
function adoptArgeLinesFromTenantSettingsIfEmpty() {
    const ta = document.getElementById('argeLines');
    if (!ta) return;
    if (normStr(ta.value)) return;
    if (typeof window.ms365TenantSettingsLoad !== 'function') return;
    const s = window.ms365TenantSettingsLoad();
    const arges = Array.isArray(s?.arges) ? s.arges : [];
    if (arges.length) {
        const lines = arges
            .map((a) => {
                const code = normSubjectCode(normStr(a?.code));
                const name = normStr(a?.name);
                if (code && name) return `${code};${name}`;
                if (code) return code;
                if (name) return name;
                return '';
            })
            .filter(Boolean);
        if (!lines.length) return;
        ta.value = lines.join('\n');
        scheduleArgePreviewFromTextarea();
        showToast('ARGE: Liste aus Schul‑Einstellungen (ARGE‑Stammdaten) übernommen.');
        return;
    }
    const subjects = Array.isArray(s?.subjects) ? s.subjects : [];
    if (!subjects.length) return;
    const lines = subjects
        .map((x) => {
            const code = normStr(x?.code);
            const name = normStr(x?.name);
            if (code && name) return `${normSubjectCode(code)};${name}`;
            return name || normSubjectCode(code);
        })
        .filter(Boolean);
    if (!lines.length) return;
    ta.value = lines.join('\n');
    scheduleArgePreviewFromTextarea();
    showToast('ARGE: Keine ARGE‑Stammdaten — Fächer aus Schul‑Einstellungen übernommen (Fallback).');
}

let argeTenantSubjectsDebounce;
function scheduleTenantSubjectsSyncFromArgeTextarea() {
    clearTimeout(argeTenantSubjectsDebounce);
    argeTenantSubjectsDebounce = setTimeout(() => {
        try {
            const ta = document.getElementById('argeLines');
            if (!ta) return;
            if (typeof window.ms365TenantSettingsLoad !== 'function' || typeof window.ms365TenantSettingsSave !== 'function')
                return;
            const current = window.ms365TenantSettingsLoad();
            const existing = Array.isArray(current?.subjects) ? current.subjects : [];
            const codeByNameKey = new Map(
                existing
                    .map((x) => ({ code: normStr(x?.code), name: normStr(x?.name) }))
                    .filter((x) => x.name && x.code)
                    .map((x) => [normSubjectKey(x.name), x.code])
            );

            const rawLines = String(ta.value || '').split(/\r\n|\n|\r/);
            const seen = new Set();
            const subjects = [];
            rawLines.forEach((line) => {
                const t = normStr(line);
                if (!t || t.startsWith('#')) return;
                const parsed = parseArgeLine(t);
                if (!parsed) return;
                const name = normStr(parsed.subjectName);
                const explicitCode = normStr(parsed.subjectCode);
                if (!name) return;
                const key = normSubjectKey(name);
                if (seen.has(key)) return;
                seen.add(key);
                const keepCode = codeByNameKey.get(key);
                const code = explicitCode ? normSubjectCode(explicitCode) : keepCode || buildSubjectCodeFromName(name);
                subjects.push({ code, name });
            });

            const domain =
                typeof window.ms365GetSchoolDomainNoAt === 'function' ? window.ms365GetSchoolDomainNoAt() : normStr(current?.domain);
            window.ms365TenantSettingsSave({
                ...current,
                domain,
                subjects
            });
        } catch {
            // ignore (kein Toast, damit Tippen nicht nervt)
        }
    }, 250);
}

function maybeUpper(s) {
    const el = document.getElementById('argeUpperNick');
    const upper = el ? !!el.checked : false;
    return upper ? s.toUpperCase() : s.toLowerCase();
}

function argeNickDeps() {
    return {
        getPrefix,
        maybeUpper,
        getSubjectCodeForDisplayName: P.getSubjectCodeForDisplayName
    };
}

function buildMailNickname(displayName, subjectCode) {
    return P.buildMailNickname(displayName, subjectCode, argeNickDeps());
}

/**
 * Parst die Textarea: eine Zeile pro Fach oder optional Anzeigename;MailNickname.
 * @returns {{ parsed: { displayName: string, mailNick: string, owner: string, description: string, technicalSlug: string }[], errors: string[] }}
 */
function parseArgeInput() {
    const ta = document.getElementById('argeLines');
    if (!ta) {
        return { parsed: [], errors: [] };
    }
    return P.parseArgeInputLines(ta.value, argeNickDeps());
}

let argePreviewDebounce;
function scheduleArgePreviewFromTextarea() {
    clearTimeout(argePreviewDebounce);
    argePreviewDebounce = setTimeout(() => {
        syncArgePreviewFromTextarea();
        renderArgePreviewTableBody();
    }, 120);
}

function scheduleArgePreviewRowsOnly() {
    clearTimeout(argePreviewDebounce);
    argePreviewDebounce = setTimeout(() => {
        if (argePreviewRows.length) {
            recomputeArgePreviewMailNicks();
            updateArgePreviewMailCellsDom();
        } else {
            syncArgePreviewFromTextarea();
            renderArgePreviewTableBody();
        }
    }, 120);
}

/** @deprecated — durch scheduleArgePreviewFromTextarea / scheduleArgePreviewRowsOnly ersetzt */
function scheduleArgePreviewRefresh() {
    scheduleArgePreviewFromTextarea();
}

function syncArgePreviewFromTextarea() {
    const { parsed } = parseArgeInput();
    argePreviewRows = parsed.map(r => ({ ...r }));
    recomputeArgePreviewMailNicks();
}

function recomputeArgePreviewMailNicks() {
    P.recomputeArgePreviewMailNicks(argePreviewRows, argeNickDeps());
}

function updateArgePreviewMailCellsDom() {
    const tbody = document.getElementById('argePreviewBody');
    if (!tbody) return;
    const domain = getDomain() || '…';
    argePreviewRows.forEach((r, i) => {
        const tr = tbody.querySelector(`tr[data-arge-index="${i}"]`);
        if (!tr) return;
        const tds = tr.querySelectorAll('td');
        if (tds.length < 5) return;
        const tech = r.technicalSlug || toNickBaseFromName(subjectForSlug(r.displayName));
        const code = getSubjectCodeForDisplayName(r.displayName, r.subjectCode);
        tds[1].innerHTML = `<code>${code || ''}</code>`;
        tds[2].textContent = tech;
        tds[2].style.fontFamily = 'Consolas,monospace';
        tds[2].style.fontSize = '0.9em';
        tds[3].textContent = r.mailNick;
        tds[3].style.fontFamily = 'Consolas,monospace';
        tds[3].style.fontSize = '0.9em';
        tds[4].textContent = r.mailNick + '@' + domain;
    });
}

function syncTextareaFromArgePreviewRows() {
    if (argeSuppressTextareaSync || !argePreviewRows.length) return;
    const ta = document.getElementById('argeLines');
    if (!ta) return;
    const lines = P.serializePreviewRowsToLines(argePreviewRows, {
        normStr: P.normStr,
        subjectForSlug: P.subjectForSlug,
        normSubjectCode: P.normSubjectCode,
        looksLikeSubjectCode: P.looksLikeSubjectCode
    });
    argeSuppressTextareaSync = true;
    ta.value = lines.join('\n');
    argeSuppressTextareaSync = false;
}

function renderArgePreviewTableBody() {
    const tbody = document.getElementById('argePreviewBody');
    if (!tbody) return;
    try {
        if (!argePreviewRows.length) {
            const ta = document.getElementById('argeLines');
            const raw = ta ? ta.value : '';
            const nonEmpty = raw
                .split(/\r\n|\n|\r/)
                .filter(l => l.trim() && !l.trim().startsWith('#')).length;
            if (nonEmpty) {
                tbody.innerHTML =
                    '<tr><td colspan="5" style="color:#6c757d;">Keine gültigen Zeilen – Format prüfen (eine Zeile pro Fach oder <code>Anzeigename;MailNickname</code>).</td></tr>';
            } else {
                tbody.innerHTML =
                    '<tr><td colspan="5" style="color:#6c757d;">Noch keine Zeilen – oben Fächer einfügen oder „+ Zeile hinzufügen“.</td></tr>';
            }
            return;
        }

        const domain = getDomain() || '…';
        tbody.replaceChildren();
        argePreviewRows.forEach((r, i) => {
            const tr = document.createElement('tr');
            tr.dataset.argeIndex = String(i);

            const td1 = document.createElement('td');
            td1.textContent = r.displayName || '';
            td1.title = 'Doppelklick zum Bearbeiten';
            td1.addEventListener('dblclick', () => {
                startCellEdit(td1, r.displayName, (next, meta) => {
                    const prev = argePreviewRows[i]?.displayName || '';
                    argePreviewRows[i].displayName = meta && meta.cancelled ? prev : next;
                    recomputeArgePreviewMailNicks();
                    syncTextareaFromArgePreviewRows();
                    scheduleTenantSubjectsSyncFromArgeTextarea();
                    renderArgePreviewTableBody();
                });
            });

            const td2 = document.createElement('td');
            td2.innerHTML = `<code>${getSubjectCodeForDisplayName(r.displayName, r.subjectCode) || ''}</code>`;
            td2.title = 'Fach-Kürzel (aus Schul‑Einstellungen; sonst automatisch erzeugt)';
            td2.addEventListener('dblclick', () => {
                startCellEdit(td2, r.subjectCode, (next, meta) => {
                    const prev = argePreviewRows[i]?.subjectCode || '';
                    const raw = meta && meta.cancelled ? prev : next;
                    const v = looksLikeSubjectCode(raw) ? normSubjectCode(raw) : '';
                    argePreviewRows[i].subjectCode = v;
                    // wenn Nick nicht explizit ist: anhand Kürzel neu berechnen
                    recomputeArgePreviewMailNicks();
                    syncTextareaFromArgePreviewRows();
                    scheduleTenantSubjectsSyncFromArgeTextarea();
                    renderArgePreviewTableBody();
                });
            });

            const td3 = document.createElement('td');
            const tech = r.technicalSlug || toNickBaseFromName(subjectForSlug(r.displayName));
            td3.textContent = tech;
            td3.style.fontFamily = 'Consolas,monospace';
            td3.style.fontSize = '0.9em';

            const td4 = document.createElement('td');
            td4.textContent = r.mailNick || '';
            td4.style.fontFamily = 'Consolas,monospace';
            td4.style.fontSize = '0.9em';
            td4.title =
                'Doppelklick zum Bearbeiten. Leer lassen und Anzeigename/Kürzel ändern = automatische Erzeugung.';
            td4.addEventListener('dblclick', () => {
                startCellEdit(td4, r.mailNickExplicit ? r.mailNick : '', (next, meta) => {
                    const prevNick = argePreviewRows[i]?.mailNick || '';
                    const prevExplicit = !!argePreviewRows[i]?.mailNickExplicit;
                    const raw = meta && meta.cancelled ? (prevExplicit ? prevNick : '') : next;
                    const cleaned = maybeUpper(String(raw || '').replace(/[^A-Za-z0-9-]/g, ''));
                    argePreviewRows[i].mailNickExplicit = normStr(raw) !== '';
                    if (argePreviewRows[i].mailNickExplicit) {
                        argePreviewRows[i].mailNick = cleaned;
                    }
                    recomputeArgePreviewMailNicks();
                    syncTextareaFromArgePreviewRows();
                    renderArgePreviewTableBody();
                });
            });

            const td5 = document.createElement('td');
            td5.textContent = r.mailNick + '@' + domain;

            tr.append(td1, td2, td3, td4, td5);
            tbody.appendChild(tr);
        });
    } catch (e) {
        console.error('ARGE-Vorschau:', e);
        tbody.innerHTML =
            '<tr><td colspan="5" style="color:#dc3545;">Vorschau konnte nicht berechnet werden. Seite neu laden oder Konsole prüfen.</td></tr>';
    }
}

/**
 * Parst die ARGE-Liste neu und übernimmt Besitzer aus dem vorherigen Stand (gleicher Anzeigename).
 * @returns {{ ok: true } | { ok: false, errors: string[] }}
 */
function syncArgeRowsFromInputPreservingOwners() {
    if (argePreviewRows.length) {
        syncTextareaFromArgePreviewRows();
    }
    const ta = document.getElementById('argeLines');
    const text = ta ? ta.value : '';
    const r = P.syncRowsFromInputPreservingOwners({
        text,
        previousArgeRows: argeRows,
        deps: argeNickDeps()
    });
    if (!r.ok) return r;
    argeRows = r.rows;
    rebuildArgeOwnerTableFromRows();
    return { ok: true };
}

const ARGE_STORAGE_KEY = 'ms365-arge-state-v2';

/** Alte Reihenfolge 1=Grundlagen, 2=Liste, 3=Besitzer → neue 1=Liste, 2=Besitzer, 3=Einstellungen */
function migrateArgeStepFromV1(step) {
    const m = { 1: 3, 2: 1, 3: 2, 4: 4 };
    const n = m[step];
    return n !== undefined ? n : step;
}

/** v2: 1–4 (Liste, Besitzer, Einstellungen, Ausführen) → v3: 1–5 mit Mitglieder als Schritt 3 */
function migrateArgeStepFromV2ToV3(step) {
    const m = { 1: 1, 2: 2, 3: 4, 4: 5 };
    const n = m[step];
    return n !== undefined ? n : step;
}

function getArgeCreateTeams() {
    const el = document.getElementById('argeCreateTeams');
    return el ? !!el.checked : true;
}

function getArgeExchangeSmtp() {
    const el = document.getElementById('argeExchangeSmtp');
    return el ? !!el.checked : true;
}

function getArgeAdminAsOwner() {
    const el = document.getElementById('argeAdminAsOwner');
    return el ? !!el.checked : true;
}

/** Mehrzeiliger Text → eindeutige UPNs (pro Gruppe). */
function parseMemberLinesText(raw) {
    const lines = String(raw || '').split(/\r\n|\n|\r/);
    const seen = new Set();
    const out = [];
    lines.forEach(line => {
        const t = String(line || '').trim();
        if (!t || t.startsWith('#')) return;
        const key = t.toLowerCase();
        if (seen.has(key)) return;
        seen.add(key);
        out.push(t);
    });
    return out;
}

function rebuildArgeMembersTableFromRows() {
    const domain = getDomain();
    const tbody = document.getElementById('argeMembersBody');
    if (!tbody) return;
    tbody.replaceChildren();
    argeRows.forEach((row, index) => {
        const tr = document.createElement('tr');
        const td1 = document.createElement('td');
        td1.textContent = row.displayName;
        const tdCode = document.createElement('td');
        tdCode.innerHTML = `<code>${getSubjectCodeForDisplayName(row.displayName, row.subjectCode) || ''}</code>`;
        tdCode.title = 'Fach-Kürzel (aus Schul‑Einstellungen; sonst automatisch erzeugt)';
        const td2 = document.createElement('td');
        td2.textContent = row.mailNick + '@' + domain;
        td2.style.fontFamily = 'Consolas,monospace';
        td2.style.fontSize = '0.9em';
        const td3 = document.createElement('td');
        const ta = document.createElement('textarea');
        ta.className = 'arge-member-lines';
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
            argeRows[index].memberLines = ta.value;
            refreshArgeScriptIfStep5();
        });
        ta.addEventListener('paste', () => setTimeout(refreshArgeScriptIfStep5, 0));
        td3.appendChild(ta);
        tr.append(td1, tdCode, td2, td3);
        tbody.appendChild(tr);
    });
}

function refreshArgeScriptIfStep5() {
    if (argeCurrentStep !== 5 || !argeRows.length) return;
    const missing = argeRows.filter(r => !r.owner);
    if (missing.length) return;
    const pre = document.getElementById('argePowerShellScript');
    if (pre)
        pre.textContent = buildStandaloneArgePs1(
            false,
            getArgeCreateTeams(),
            getArgeExchangeSmtp(),
            getArgeAdminAsOwner()
        );
}

function rebuildArgeOwnerTableFromRows() {
    const domain = getDomain();
    const tbody = document.getElementById('argeOwnerBody');
    if (!tbody) return;
    tbody.replaceChildren();
    argeRows.forEach((row, index) => {
        const tr = document.createElement('tr');
        const td1 = document.createElement('td');
        td1.textContent = row.displayName;
        const tdCode = document.createElement('td');
        tdCode.innerHTML = `<code>${getSubjectCodeForDisplayName(row.displayName, row.subjectCode) || ''}</code>`;
        tdCode.title = 'Fach-Kürzel (aus Schul‑Einstellungen; sonst automatisch erzeugt)';
        const td2 = document.createElement('td');
        td2.textContent = row.mailNick + '@' + domain;
        const td3 = document.createElement('td');
        td3.textContent = row.mailNick;
        const td4 = document.createElement('td');
        const inp = document.createElement('input');
        inp.type = 'email';
        inp.placeholder = 'besitzer@' + domain;
        inp.style.width = '100%';
        inp.style.padding = '8px';
        inp.value = row.owner || '';
        inp.addEventListener('input', () => {
            argeRows[index].owner = inp.value.trim();
        });
        td4.appendChild(inp);
        tr.append(td1, tdCode, td2, td3, td4);
        tbody.appendChild(tr);
    });
}

function saveArgeState() {
    try {
        const state = {
            argeStepOrder: 'v3',
            argeCurrentStep,
            argeRows,
            argeDefaultPrefix: document.getElementById('argeDefaultPrefix').value,
            argeUpperNick: document.getElementById('argeUpperNick').checked,
            argeCreateTeams: getArgeCreateTeams(),
            argeExchangeSmtp: getArgeExchangeSmtp(),
            argeAdminAsOwner: getArgeAdminAsOwner(),
            argeLines: document.getElementById('argeLines').value,
            argePowerShellScript: document.getElementById('argePowerShellScript').textContent
        };
        localStorage.setItem(ARGE_STORAGE_KEY, JSON.stringify(state));
        showToast('ARGEs: Zwischenstand gespeichert.');
    } catch (e) {
        showToast('Speichern fehlgeschlagen: ' + e.message);
    }
}

function loadArgeState() {
    try {
        let raw = localStorage.getItem(ARGE_STORAGE_KEY);
        if (!raw) {
            raw = localStorage.getItem('ms365-arge-state-v1');
        }
        if (!raw) {
            showToast('Kein gespeicherter Stand für ARGEs.');
            return;
        }
        const state = JSON.parse(raw);
        let step = typeof state.argeCurrentStep === 'number' ? state.argeCurrentStep : 1;
        if (state.argeStepOrder === 'v3') {
            /* Schritte 1–5 unverändert */
        } else if (state.argeStepOrder === 'v2') {
            step = migrateArgeStepFromV2ToV3(step);
        } else {
            step = migrateArgeStepFromV1(step);
            step = migrateArgeStepFromV2ToV3(step);
        }
        argeCurrentStep = step;
        argeRows = Array.isArray(state.argeRows) ? state.argeRows : [];
        argeRows.forEach(function (row) {
            if (row.memberLines === undefined) {
                row.memberLines = '';
            }
        });
        if (state.argeMemberEmails !== undefined && String(state.argeMemberEmails || '').trim() !== '') {
            const legacy = String(state.argeMemberEmails);
            argeRows.forEach(function (row) {
                if (!String(row.memberLines || '').trim()) {
                    row.memberLines = legacy;
                }
            });
        }
        if (
            typeof window.ms365SetSchoolDomainNoAt === 'function' &&
            state.argeDomain !== undefined &&
            String(state.argeDomain).trim() !== ''
        ) {
            window.ms365SetSchoolDomainNoAt(state.argeDomain);
        }
        document.getElementById('argeDefaultPrefix').value =
            state.argeDefaultPrefix !== undefined ? state.argeDefaultPrefix : '';
        document.getElementById('argeUpperNick').checked = !!state.argeUpperNick;
        const argeTeamsEl = document.getElementById('argeCreateTeams');
        if (argeTeamsEl) {
            argeTeamsEl.checked = state.argeCreateTeams !== undefined ? !!state.argeCreateTeams : true;
        }
        const argeExoEl = document.getElementById('argeExchangeSmtp');
        if (argeExoEl) {
            argeExoEl.checked = state.argeExchangeSmtp !== undefined ? !!state.argeExchangeSmtp : true;
        }
        const argeAdminEl = document.getElementById('argeAdminAsOwner');
        if (argeAdminEl) {
            argeAdminEl.checked = state.argeAdminAsOwner !== undefined ? !!state.argeAdminAsOwner : true;
        }
        document.getElementById('argeLines').value = state.argeLines || '';
        document.getElementById('argeParseError').style.display = 'none';
        const pre = document.getElementById('argePowerShellScript');
        if (pre && state.argePowerShellScript !== undefined) {
            pre.textContent = state.argePowerShellScript;
        }
        if (argeRows.length) {
            rebuildArgeOwnerTableFromRows();
            rebuildArgeMembersTableFromRows();
        } else {
            document.getElementById('argeOwnerBody').replaceChildren();
            const amb = document.getElementById('argeMembersBody');
            if (amb) amb.replaceChildren();
        }
        goToArgeStep(Math.min(Math.max(1, argeCurrentStep), 5));
        scheduleArgePreviewFromTextarea();
        showToast('ARGEs: Stand geladen.');
    } catch (e) {
        showToast('Laden fehlgeschlagen: ' + e.message);
    }
}

function clearArgeState() {
    if (!confirm('Gespeicherten Zwischenstand für ARGEs wirklich löschen?')) {
        return;
    }
    try {
        localStorage.removeItem(ARGE_STORAGE_KEY);
        localStorage.removeItem('ms365-arge-state-v1');
        argeCurrentStep = 1;
        argeRows = [];
        document.getElementById('argeDefaultPrefix').value = '';
        document.getElementById('argeUpperNick').checked = false;
        const argeTeamsClear = document.getElementById('argeCreateTeams');
        if (argeTeamsClear) argeTeamsClear.checked = true;
        const argeExoClear = document.getElementById('argeExchangeSmtp');
        if (argeExoClear) argeExoClear.checked = true;
        const argeAdminClear = document.getElementById('argeAdminAsOwner');
        if (argeAdminClear) argeAdminClear.checked = true;
        document.getElementById('argeLines').value = '';
        document.getElementById('argeParseError').style.display = 'none';
        document.getElementById('argeOwnerBody').replaceChildren();
        const argeMemBodyClear = document.getElementById('argeMembersBody');
        if (argeMemBodyClear) argeMemBodyClear.replaceChildren();
        document.getElementById('argePowerShellScript').textContent = '';
        argePreviewRows = [];
        goToArgeStep(1);
        scheduleArgePreviewFromTextarea();
        showToast('ARGEs: Speicher geleert.');
    } catch (e) {
        showToast('Fehler: ' + e.message);
    }
}

window.ms365SaveArge = saveArgeState;
window.ms365LoadArge = loadArgeState;
window.ms365ClearArge = clearArgeState;
window.ms365ShowToast = showToast;

/**
 * Snapshot für Online-Ausführung (Microsoft Graph im Browser), siehe arge-graph.js
 * @returns {{ rows: { displayName: string, mailNick: string, owner: string, description: string }[], createTeams: boolean, exchangeSmtp: boolean }}
 */
window.ms365GetArgeSnapshotForGraph = function () {
    return {
        rows: argeRows.map(function (r) {
            return {
                displayName: r.displayName,
                mailNick: r.mailNick,
                owner: r.owner,
                description: r.description,
                memberEmails: parseMemberLinesText(r.memberLines || '')
            };
        }),
        createTeams: getArgeCreateTeams(),
        exchangeSmtp: getArgeExchangeSmtp(),
        adminAsOwner: getArgeAdminAsOwner()
    };
};

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

function argeBuildStateSnapshot() {
    return argeBuildStateSnapshotImpl({
        argeCurrentStep,
        argeDefaultPrefix: document.getElementById('argeDefaultPrefix')?.value ?? '',
        argeUpperNick: !!document.getElementById('argeUpperNick')?.checked,
        argeCreateTeams: getArgeCreateTeams(),
        argeExchangeSmtp: getArgeExchangeSmtp(),
        argeAdminAsOwner: getArgeAdminAsOwner(),
        argeLines: document.getElementById('argeLines')?.value ?? '',
        rows: argeRows || [],
        normStr: P.normStr,
        subjectForSlug: P.subjectForSlug
    });
}

function applyImportedArgeState(obj) {
    argeApplyImportedStateImpl(obj, {
        normStr: P.normStr,
        subjectForSlug: P.subjectForSlug,
        looksLikeSubjectCode: P.looksLikeSubjectCode,
        normSubjectCode: P.normSubjectCode,
        displayNameFromSubjectLine: P.displayNameFromSubjectLine,
        syncArgePreviewFromTextarea,
        setArgeRows: (rows) => {
            argeRows = rows;
        },
        rebuildArgeOwnerTableFromRows,
        rebuildArgeMembersTableFromRows,
        scheduleArgePreviewRowsOnly,
        refreshArgeScriptIfStep5,
        showToast,
        getArgePreviewRows: () => argePreviewRows
    });
}

function exportArgeCsv() {
    const { csv, filename } = argeExportCsvImpl({
        domain: getDomain(),
        argeRows,
        argePreviewRows,
        syncArgePreviewFromTextarea,
        normStr,
        subjectForSlug,
        getSubjectCodeForDisplayName
    });
    downloadBlob(filename, csv, 'text/csv;charset=utf-8');
    showToast('ARGE: CSV exportiert.');
}

function importArgeCsvText(text) {
    const rows = parseArgeCsvToRows(text, {
        normStr,
        displayNameFromSubjectLine,
        subjectForSlug,
        normSubjectCode
    });
    if (rows === null) return;
    applyImportedArgeState({ rows });
}

function buildStandaloneArgePs1(standalone, createTeams, setExchangeSmtp, adminAsOwner) {
    return buildStandaloneArgePs1Impl(argeRows, getDomain(), standalone, createTeams, setExchangeSmtp, adminAsOwner);
}

function downloadArgeStandalonePackage() {
    if (!argeRows.length) {
        showToast('Keine ARGE-Daten – zuerst ARGE-Liste, Besitzer und Einstellungen durchgehen.');
        return;
    }
    const missing = argeRows.filter(r => !r.owner);
    if (missing.length) {
        showToast('Bitte für alle ARGEs einen Besitzer eintragen.');
        return;
    }
    if (typeof window.ms365BuildPolyglotCmd !== 'function') {
        showToast('polyglot-cmd.js fehlt – Seite neu laden.');
        return;
    }
    if (
        getArgeExchangeSmtp() &&
        (typeof window.ms365IsTenantSchoolDomainConfigured !== 'function' ||
            !window.ms365IsTenantSchoolDomainConfigured())
    ) {
        showToast('Für die Exchange-Option legen Sie die E-Mail-Domain der Schule in den Schul‑Einstellungen fest.');
        if (typeof window.ms365ShowTenantDomainRequiredModal === 'function') {
            window.ms365ShowTenantDomainRequiredModal();
        }
        return;
    }
    const ps1 = buildStandaloneArgePs1(true, getArgeCreateTeams(), getArgeExchangeSmtp(), getArgeAdminAsOwner());
    const cmd = window.ms365BuildPolyglotCmd({
        title: 'ARGE-Gruppen-Anlage',
        echoLine: 'Starte ARGE-Gruppen-Anlage Microsoft Graph ...',
        psBody: ps1
    });
    downloadBlob('ARGE-Gruppen-Anlage.cmd', cmd);
    showToast('ARGE-Gruppen-Anlage.cmd heruntergeladen – Doppelklick zum Start.');
}

window.downloadArgeStandalonePackage = downloadArgeStandalonePackage;

// UI Wiring — Vorschau zuerst, damit Eingabe auch bei späteren Fehlern funktioniert
const argeLinesEl = document.getElementById('argeLines');
if (argeLinesEl) {
    argeLinesEl.addEventListener('input', () => {
        if (argeSuppressTextareaSync) return;
        scheduleArgePreviewFromTextarea();
        scheduleTenantSubjectsSyncFromArgeTextarea();
    });
    argeLinesEl.addEventListener('paste', () =>
        setTimeout(() => {
            if (!argeSuppressTextareaSync) scheduleArgePreviewFromTextarea();
            scheduleTenantSubjectsSyncFromArgeTextarea();
        }, 0)
    );
}
['schoolEmailDomain', 'argeDefaultPrefix'].forEach(id => {
    const el = document.getElementById(id);
    if (el) {
        el.addEventListener('input', scheduleArgePreviewRowsOnly);
        el.addEventListener('input', refreshArgeScriptIfStep5);
    }
});
const argeUpperEl = document.getElementById('argeUpperNick');
if (argeUpperEl) argeUpperEl.addEventListener('change', scheduleArgePreviewRowsOnly);
const argeTeamsEl = document.getElementById('argeCreateTeams');
if (argeTeamsEl) argeTeamsEl.addEventListener('change', refreshArgeScriptIfStep5);
const argeExoEl = document.getElementById('argeExchangeSmtp');
if (argeExoEl) argeExoEl.addEventListener('change', refreshArgeScriptIfStep5);
const argeAdminAsOwnerEl = document.getElementById('argeAdminAsOwner');
if (argeAdminAsOwnerEl) argeAdminAsOwnerEl.addEventListener('change', refreshArgeScriptIfStep5);
document.getElementById('argeBack1').addEventListener('click', () => goToArgeStep(1));
document.getElementById('argeGoTo3').addEventListener('click', () => goToArgeStep(3));
const argeMemberBack = document.getElementById('argeMemberBack');
if (argeMemberBack) argeMemberBack.addEventListener('click', () => goToArgeStep(2));
const argeMemberNext = document.getElementById('argeMemberNext');
if (argeMemberNext) argeMemberNext.addEventListener('click', () => goToArgeStep(4));
document.getElementById('argeBack2').addEventListener('click', () => goToArgeStep(3));
document.getElementById('argeBack3').addEventListener('click', () => goToArgeStep(4));

const argePreviewAddRow = document.getElementById('argePreviewAddRow');
if (argePreviewAddRow) {
    argePreviewAddRow.addEventListener('click', () => {
        argePreviewRows.push({
            displayName: '',
            mailNick: '',
            owner: '',
            description: '',
            technicalSlug: '',
            mailNickExplicit: false
        });
        recomputeArgePreviewMailNicks();
        renderArgePreviewTableBody();
    });
}

document.getElementById('argeParseAndGo3').addEventListener('click', () => {
    const errEl = document.getElementById('argeParseError');
    errEl.style.display = 'none';
    if (
        typeof window.ms365IsTenantSchoolDomainConfigured !== 'function' ||
        !window.ms365IsTenantSchoolDomainConfigured()
    ) {
        errEl.textContent =
            'Bitte legen Sie die E-Mail-Domain der Schule in den Schul‑Einstellungen fest (Seite „Schul‑Einstellungen“).';
        errEl.style.display = 'block';
        if (typeof window.ms365ShowTenantDomainRequiredModal === 'function') {
            window.ms365ShowTenantDomainRequiredModal();
        }
        return;
    }

    if (!argePreviewRows.length) {
        syncArgePreviewFromTextarea();
    }
    if (!argePreviewRows.length) {
        errEl.textContent =
            'Bitte mindestens eine ARGE-Zeile eintragen oder in der Vorschau eine Zeile hinzufügen und ausfüllen.';
        errEl.style.display = 'block';
        return;
    }

    recomputeArgePreviewMailNicks();
    const rowErrors = [];
    argePreviewRows.forEach((r, idx) => {
        if (!(r.displayName || '').trim()) {
            rowErrors.push('Vorschau Zeile ' + (idx + 1) + ': Anzeigename fehlt.');
        }
        if (!(r.mailNick || '').trim()) {
            rowErrors.push('Vorschau Zeile ' + (idx + 1) + ': Mail-Nickname fehlt.');
        }
    });
    if (rowErrors.length) {
        errEl.textContent = rowErrors.join('\n');
        errEl.style.display = 'block';
        return;
    }

    const ownerByKey = new Map(argeRows.map(r => [r.displayName.toLowerCase(), r.owner]));
    const memberLinesByKey = new Map(argeRows.map(r => [r.displayName.toLowerCase(), r.memberLines || '']));
    argeRows = argePreviewRows.map(r => ({
        displayName: r.displayName.trim(),
        mailNick: r.mailNick,
        owner: ownerByKey.get(r.displayName.trim().toLowerCase()) || '',
        memberLines: memberLinesByKey.get(r.displayName.trim().toLowerCase()) || '',
        description: 'ARGE-Gruppe: ' + r.displayName.trim(),
        subjectName: normStr(r.subjectName || subjectForSlug(r.displayName.trim())),
        subjectCode: normStr(r.subjectCode || '')
    }));

    rebuildArgeOwnerTableFromRows();

    goToArgeStep(2);
});

document.getElementById('argeGoTo4').addEventListener('click', () => {
    const sync = syncArgeRowsFromInputPreservingOwners();
    if (!sync.ok) {
        const errEl = document.getElementById('argeParseError');
        if (errEl) {
            errEl.textContent = sync.errors.join('\n');
            errEl.style.display = 'block';
        }
        showToast(sync.errors[0] || 'ARGE-Liste konnte nicht verarbeitet werden.');
        goToArgeStep(1);
        scheduleArgePreviewFromTextarea();
        return;
    }
    const missing = argeRows.filter(r => !r.owner);
    if (missing.length) {
        showToast('Bitte für alle ARGEs einen Besitzer (UPN) eintragen (Schritt 2).');
        goToArgeStep(2);
        return;
    }
    if (
        getArgeExchangeSmtp() &&
        (typeof window.ms365IsTenantSchoolDomainConfigured !== 'function' ||
            !window.ms365IsTenantSchoolDomainConfigured())
    ) {
        showToast('Für die Exchange-Option legen Sie die E-Mail-Domain der Schule in den Schul‑Einstellungen fest.');
        if (typeof window.ms365ShowTenantDomainRequiredModal === 'function') {
            window.ms365ShowTenantDomainRequiredModal();
        }
        return;
    }
    document.getElementById('argeParseError').style.display = 'none';
    document.getElementById('argePowerShellScript').textContent = buildStandaloneArgePs1(
        false,
        getArgeCreateTeams(),
        getArgeExchangeSmtp(),
        getArgeAdminAsOwner()
    );
    goToArgeStep(5);
});

document.getElementById('argeCopyScript').addEventListener('click', () => {
    const t = document.getElementById('argePowerShellScript').textContent;
    navigator.clipboard.writeText(t).then(() => showToast('Script kopiert.'));
});

// step header keyboard support
document.querySelectorAll('.arge-steps .step').forEach(el => {
    el.setAttribute('tabindex', '0');
    el.addEventListener('click', () => {
        const s = argeStepNum(el);
        if (s <= argeCurrentStep || el.classList.contains('completed')) {
            goToArgeStep(s);
        }
    });
    el.addEventListener('keydown', e => {
        if (e.key === 'Enter' || e.key === ' ') {
            e.preventDefault();
            el.click();
        }
    });
});

const elArgeStepsInit = document.querySelector('.arge-steps');
if (elArgeStepsInit && typeof window.ms365ApplyStepProgress === 'function') {
    window.ms365ApplyStepProgress(elArgeStepsInit, argeCurrentStep, [1, 2, 3, 4, 5]);
}

applyInitialModeFromUrl();
if (panelA && panelA.style.display !== 'none') {
    adoptArgeLinesFromTenantSettingsIfEmpty();
    scheduleArgePreviewRefresh();
}

window.addEventListener('ms365-tenant-settings-changed', () => {
    adoptArgeLinesFromTenantSettingsIfEmpty();
    scheduleArgePreviewFromTextarea();
});

// bottom toolbar wiring (save/load + import/export)
const btnSaveState = document.getElementById('btnSaveState');
if (btnSaveState) btnSaveState.addEventListener('click', () => saveArgeState());
const btnLoadState = document.getElementById('btnLoadState');
if (btnLoadState) btnLoadState.addEventListener('click', () => loadArgeState());
const btnClearStorage = document.getElementById('btnClearStorage');
if (btnClearStorage) btnClearStorage.addEventListener('click', () => clearArgeState());

const btnExportJson = document.getElementById('btnExportArgeJson');
if (btnExportJson) {
    btnExportJson.addEventListener('click', () => {
        const snap = argeBuildStateSnapshot();
        downloadJson(`arge-export-${new Date().toISOString().slice(0, 10)}.json`, snap);
        showToast('ARGE: JSON exportiert.');
    });
}

const fileJson = document.getElementById('argeImportJsonFile');
const btnImportJson = document.getElementById('btnImportArgeJson');
if (btnImportJson && fileJson) {
    btnImportJson.addEventListener('click', () => fileJson.click());
    fileJson.addEventListener('change', async (e) => {
        const f = e.target.files && e.target.files[0];
        if (!f) return;
        try {
            const text = await f.text();
            const obj = JSON.parse(text);
            applyImportedArgeState(obj);
        } catch (err) {
            showToast('ARGE: JSON Import fehlgeschlagen: ' + (err?.message || String(err)));
        } finally {
            fileJson.value = '';
        }
    });
}

const btnExportCsv = document.getElementById('btnExportArgeCsv');
if (btnExportCsv) btnExportCsv.addEventListener('click', () => exportArgeCsv());

const fileCsv = document.getElementById('argeImportCsvFile');
const btnImportCsv = document.getElementById('btnImportArgeCsv');
if (btnImportCsv && fileCsv) {
    btnImportCsv.addEventListener('click', () => fileCsv.click());
    fileCsv.addEventListener('change', async (e) => {
        const f = e.target.files && e.target.files[0];
        if (!f) return;
        try {
            const text = await f.text();
            importArgeCsvText(text);
        } catch (err) {
            showToast('ARGE: CSV Import fehlgeschlagen: ' + (err?.message || String(err)));
        } finally {
            fileCsv.value = '';
        }
    });
}

