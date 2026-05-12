
const ns = (window.ms365Kursteam = window.ms365Kursteam || {});

// State (wird von mehreren Modulen genutzt)
ns.rawData = [];
ns.filteredData = [];
ns.teamsData = [];
/** @type {'unset'|'webuntis'|'manual'} */
ns.kursteamEntryMode = 'unset';
ns.currentStep = 0;
ns.teacherEmailMapping = {};
ns.teamsGenerated = false;

// DOM Cache
ns.dom = {
    uploadArea: document.getElementById('uploadArea'),
    fileInput: document.getElementById('fileInput'),
    teacherUploadArea: document.getElementById('teacherUploadArea'),
    teacherFileInput: document.getElementById('teacherFileInput'),
    appModal: document.getElementById('appModal'),
    modalTitle: document.getElementById('modalTitle'),
    modalBody: document.getElementById('modalBody'),
    modalCancel: document.getElementById('modalCancel'),
    modalOk: document.getElementById('modalOk')
};

ns.showToast = function showToast(msg) {
    const el = document.getElementById('toast');
    el.textContent = msg;
    el.classList.add('show');
    clearTimeout(ns.showToast._t);
    ns.showToast._t = setTimeout(() => el.classList.remove('show'), 3500);
};

let modalOkHandler = null;

ns.openModal = function openModal(title, bodyHtml, onOk) {
    ns.dom.modalTitle.textContent = title;
    ns.dom.modalBody.innerHTML = bodyHtml;
    modalOkHandler = onOk;
    ns.dom.appModal.classList.add('open');
};

ns.closeModal = function closeModal() {
    ns.dom.appModal.classList.remove('open');
    modalOkHandler = null;
    ns.dom.modalBody.innerHTML = '';
};

ns.confirmModal = function confirmModal(title, message, onConfirm) {
    ns.openModal(title, '<p>' + ns.escapeHtml(message) + '</p>', () => {
        ns.closeModal();
        onConfirm();
    });
};

ns.invalidateTeams = function invalidateTeams() {
    ns.teamsData = [];
    ns.teamsGenerated = false;
    document.getElementById('teamsTableContainer').style.display = 'none';
    document.getElementById('validationResults').style.display = 'none';
    const preview = document.getElementById('manualTeamsPreviewContainer');
    if (preview) preview.style.display = 'none';
    // Nur „Weiter zur Anlage“ (Schritt Teams konfigurieren) ausblenden – nicht continueBtn3
    // (Lehrer zuordnen → Team-Konfiguration), das ist reine Schritt-Navigation.
    const c4 = document.getElementById('continueBtn4');
    if (c4) c4.style.display = 'none';
    // UX: Teams generieren wieder als "primäre Aktion" markieren.
    const gen = document.getElementById('btnGenerateTeamNames');
    if (gen) gen.className = 'btn btn-success kursteam-generate-teams-btn';
    const manRow = document.getElementById('kursteamManualAddRow');
    if (manRow) manRow.style.display = 'none';
};

// Modal wiring
ns.dom.modalCancel.addEventListener('click', ns.closeModal);
ns.dom.modalOk.addEventListener('click', () => {
    if (typeof modalOkHandler === 'function') modalOkHandler();
});
ns.dom.appModal.addEventListener('click', (e) => {
    if (e.target === ns.dom.appModal) ns.closeModal();
});

// Toolbar Save/Load/Clear delegiert je nach aktivem Panel
function getActivePanelMode() {
    const pw = document.getElementById('panelWebuntis');
    const pj = document.getElementById('panelJahrgang');
    const pa = document.getElementById('panelArge');
    const pg = document.getElementById('panelGruppenPolicy');
    const hidden = el => !el || window.getComputedStyle(el).display === 'none';
    if (!hidden(pw)) return 'webuntis';
    if (!hidden(pj)) return 'jahrgang';
    if (!hidden(pa)) return 'arge';
    if (!hidden(pg)) return 'gruppenerstellung';
    return 'webuntis';
}

document.getElementById('btnSaveState').addEventListener('click', () => {
    const mode = getActivePanelMode();
    if (mode === 'jahrgang' && typeof window.ms365SaveJahrgang === 'function') return window.ms365SaveJahrgang();
    if (mode === 'arge' && typeof window.ms365SaveArge === 'function') return window.ms365SaveArge();
    if (mode === 'gruppenerstellung' && typeof window.ms365SaveGruppenerstellung === 'function') return window.ms365SaveGruppenerstellung();
    if (typeof ns.saveStateToStorage === 'function') ns.saveStateToStorage();
});
document.getElementById('btnLoadState').addEventListener('click', () => {
    const mode = getActivePanelMode();
    if (mode === 'jahrgang' && typeof window.ms365LoadJahrgang === 'function') return window.ms365LoadJahrgang();
    if (mode === 'arge' && typeof window.ms365LoadArge === 'function') return window.ms365LoadArge();
    if (mode === 'gruppenerstellung' && typeof window.ms365LoadGruppenerstellung === 'function') return window.ms365LoadGruppenerstellung();
    if (typeof ns.loadStateFromStorage === 'function') ns.loadStateFromStorage();
});
document.getElementById('btnClearStorage').addEventListener('click', () => {
    const mode = getActivePanelMode();
    if (mode === 'jahrgang' && typeof window.ms365ClearJahrgang === 'function') return window.ms365ClearJahrgang();
    if (mode === 'arge' && typeof window.ms365ClearArge === 'function') return window.ms365ClearArge();
    if (mode === 'gruppenerstellung' && typeof window.ms365ClearGruppenerstellung === 'function') return window.ms365ClearGruppenerstellung();
    if (typeof ns.clearStorage === 'function') ns.clearStorage();
});

// Globale Helfer (für andere Module/HTML)
window.ms365ShowToast = ns.showToast;

