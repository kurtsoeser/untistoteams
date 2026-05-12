
const ns = (window.ms365Kursteam = window.ms365Kursteam || {});

function upsertTenantTeachersFromMapping(mapping) {
    if (!mapping || typeof mapping !== 'object') return;
    if (typeof window.ms365TenantSettingsLoad !== 'function' || typeof window.ms365TenantSettingsSave !== 'function') return;

    const current = window.ms365TenantSettingsLoad();
    const teachers = Array.isArray(current && current.teachers) ? [...current.teachers] : [];
    const index = new Map();
    teachers.forEach((t, i) => {
        const code = String(t && t.code ? t.code : '').trim().toUpperCase();
        if (code) index.set(code, i);
    });

    Object.entries(mapping).forEach(([codeRaw, emailRaw]) => {
        const code = String(codeRaw || '').trim().toUpperCase();
        const email = String(emailRaw || '').trim().toLowerCase();
        if (!code || !email || !email.includes('@')) return;
        if (index.has(code)) {
            const i = index.get(code);
            const prev = teachers[i] || {};
            teachers[i] = { ...prev, code, email };
        } else {
            teachers.push({ code, name: '', email });
            index.set(code, teachers.length - 1);
        }
    });

    window.ms365TenantSettingsSave({ ...current, teachers });
}

function loadTenantTeacherEmailsIfEmpty() {
    if (typeof window.ms365TenantSettingsGetTeacherEmailMap !== 'function') return;
    const map = window.ms365TenantSettingsGetTeacherEmailMap();
    if (!map || !Object.keys(map).length) return;
    // Schul‑Einstellungen sind Basis: fehlende Einträge ergänzen, bestehende (manuelle) nicht überschreiben.
    ns.teacherEmailMapping = ns.teacherEmailMapping || {};
    Object.entries(map).forEach(([k, v]) => {
        const kk = String(k || '').trim().toUpperCase();
        if (!kk) return;
        if (!ns.teacherEmailMapping[kk]) ns.teacherEmailMapping[kk] = v;
    });
    const el = document.getElementById('teacherCount');
    if (el) el.textContent = Object.keys(ns.teacherEmailMapping).length;
    const info = document.getElementById('teacherMappingInfo');
    if (info) info.style.display = 'block';
}

/** Gleiche Spalten wie WebUntis-Lehrerliste (Kürzel, E-Mail, …). */
ns.getKuerzelFromLehrerRow = function getKuerzelFromLehrerRow(row) {
    const r = ns.normalizeImportedRowKeys(row);
    let v = (r.Kürzel || r.Kuerzel || r.kuerzel || r.Code || r.code || r.Lehrer || '').toString().trim();
    if (v) return v;
    for (const k of Object.keys(r)) {
        const kl = k.toLowerCase().replace(/\s+/g, '');
        if (/^(kürzel|kuerzel|code|lehrer|abbrev)$/.test(kl) || /kuerzel|kürzel/.test(kl)) {
            const x = r[k];
            if (x != null && String(x).trim()) return String(x).trim();
        }
    }
    const vals = Object.values(r).map(x => (x == null ? '' : String(x).trim())).filter(Boolean);
    if (vals.length >= 2) {
        if (vals[1].includes('@') && !vals[0].includes('@')) return vals[0];
        if (vals[0].includes('@') && !vals[1].includes('@')) return vals[1];
    }
    if (vals.length >= 1 && !vals[0].includes('@')) return vals[0];
    return '';
};

ns.getEmailFromLehrerRow = function getEmailFromLehrerRow(row) {
    const r = ns.normalizeImportedRowKeys(row);
    let v = (r['E-Mail'] || r.Email || r.email || r.Mail || r.mail || '').toString().trim().toLowerCase();
    if (v) return v;
    for (const k of Object.keys(r)) {
        const kl = k.toLowerCase().replace(/\s+/g, '');
        if (/^(e-?mail|email|mail)$/.test(kl) || /^e-?mail$/i.test(k)) {
            const x = r[k];
            if (x != null && String(x).trim()) return String(x).trim().toLowerCase();
        }
    }
    const vals = Object.values(r).map(x => (x == null ? '' : String(x).trim())).filter(Boolean);
    if (vals.length >= 2) {
        if (vals[1].includes('@')) return vals[1].toLowerCase();
        if (vals[0].includes('@')) return vals[0].toLowerCase();
    }
    if (vals.length === 1 && vals[0].includes('@')) return vals[0].toLowerCase();
    return '';
};

/** Wenn SheetJS die Kopfzeile nicht erkennt: Zeilen mit erstem Semikolon/Komma als Trenner. */
function parseTeacherListCsvLineByLine(text) {
    let t = String(text || '');
    if (t.charCodeAt(0) === 0xfeff) t = t.slice(1);
    const lines = t.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
    if (lines.length < 2) return [];
    const out = [];
    for (let i = 1; i < lines.length; i++) {
        const line = lines[i];
        let sep = line.indexOf(';');
        if (sep < 0) sep = line.indexOf(',');
        if (sep < 0) continue;
        const k = line.slice(0, sep).trim();
        const em = line.slice(sep + 1).trim();
        if (!k || !em.includes('@')) continue;
        if (/^k(ue|ü)rzel$/i.test(k.replace(/\s/g, ''))) continue;
        out.push({ Kürzel: k, 'E-Mail': em });
    }
    return out;
}

function parseTeacherCsvToJsonData(text) {
    let s = String(text || '');
    if (s.charCodeAt(0) === 0xfeff) s = s.slice(1);
    let workbook = XLSX.read(s, { type: 'string', FS: ';' });
    let firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    let jsonData = XLSX.utils.sheet_to_json(firstSheet);
    if (!jsonData.length || Object.keys(jsonData[0] || {}).length < 2) {
        const wb2 = XLSX.read(s, { type: 'string', FS: ',' });
        const sh2 = wb2.Sheets[wb2.SheetNames[0]];
        const j2 = XLSX.utils.sheet_to_json(sh2);
        if (j2.length) jsonData = j2;
    }
    return jsonData;
}

/**
 * Eine Zeile aus Copy-Paste: Kürzel und E-Mail (Tab, Semikolon, |, mehrere Leerzeichen oder ein Leerzeichen).
 * @returns {{ kuerzel: string, email: string } | null}
 */
function parseTeacherEmailPasteLine(line) {
    const t = String(line || '').trim();
    if (!t || t.startsWith('#')) return null;
    let kuerzel = '';
    let email = '';
    if (t.includes('\t')) {
        const p = t.split(/\t+/).map(s => s.trim()).filter(Boolean);
        if (p.length >= 2) {
            kuerzel = p[0];
            email = p.slice(1).join(' ').trim();
        }
    }
    if (!kuerzel && t.includes(';')) {
        const semi = t.indexOf(';');
        const left = t.slice(0, semi).trim();
        const right = t.slice(semi + 1).trim();
        if (left && right) {
            kuerzel = left;
            email = right;
        }
    }
    if (!kuerzel && /\|/.test(t)) {
        const p = t.split(/\s*\|\s*/).map(s => s.trim()).filter(Boolean);
        if (p.length >= 2) {
            kuerzel = p[0];
            email = p.slice(1).join(' ').trim();
        }
    }
    if (!kuerzel && /\s{2,}/.test(t)) {
        const p = t.split(/\s{2,}/).map(s => s.trim()).filter(Boolean);
        if (p.length >= 2) {
            kuerzel = p[0];
            email = p.slice(1).join(' ').trim();
        }
    }
    if (!kuerzel) {
        const m = t.match(/^(\S+)\s+(.+)$/);
        if (m) {
            kuerzel = m[1].trim();
            email = m[2].trim();
        }
    }
    if (!kuerzel || !email) return null;
    email = email.toLowerCase();
    if (!email.includes('@')) return null;
    return { kuerzel: kuerzel.toUpperCase(), email };
}

ns.processTeacherMapping = function processTeacherMapping(data, options) {
    const replace = !options || options.replace !== false;
    if (replace) ns.teacherEmailMapping = {};
    data.forEach(row => {
        const r = ns.normalizeImportedRowKeys(row);
        const kuerzel = ns.getKuerzelFromLehrerRow(r).toUpperCase();
        const email = ns.getEmailFromLehrerRow(r);
        if (!kuerzel || !email) return;
        const kNorm = kuerzel.replace(/Ü/g, 'U').replace(/ü/g, 'U');
        if (/^KUERZEL$/i.test(kNorm) && (!email || /^e-?mail$/i.test(email))) return;
        ns.teacherEmailMapping[kuerzel] = email;
    });
    document.getElementById('teacherCount').textContent = Object.keys(ns.teacherEmailMapping).length;
    document.getElementById('teacherMappingInfo').style.display = 'block';
    upsertTenantTeachersFromMapping(ns.teacherEmailMapping);
    if (ns.currentStep === 4) ns.updateTeacherStats();
    else ns.displayTeacherMappingTable();
};

ns.displayTeacherMappingTable = function displayTeacherMappingTable() {
    const tbody = document.getElementById('teacherMappingBody');
    tbody.replaceChildren();
    Object.entries(ns.teacherEmailMapping).forEach(([kuerzel, email]) => {
        const tr = document.createElement('tr');
        const td1 = document.createElement('td');
        const strong = document.createElement('strong');
        strong.textContent = kuerzel;
        td1.appendChild(strong);
        const td2 = document.createElement('td');
        td2.textContent = email;
        const td3 = document.createElement('td');
        td3.textContent = '-';
        const td4 = document.createElement('td');
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'btn btn-small btn-danger kt-delete-btn';
        btn.textContent = 'X';
        btn.title = 'Zuordnung löschen';
        btn.setAttribute('aria-label', 'Zuordnung löschen');
        btn.addEventListener('click', () => ns.removeTeacherMapping(kuerzel));
        td4.appendChild(btn);
        tr.append(td1, td2, td3, td4);
        tbody.appendChild(tr);
    });
};

ns.displayTeacherMappingTableWithUsage = function displayTeacherMappingTableWithUsage(requiredTeachers) {
    const tbody = document.getElementById('teacherMappingBody');
    tbody.replaceChildren();
    Object.entries(ns.teacherEmailMapping).forEach(([kuerzel, email]) => {
        const isUsed = requiredTeachers.includes(kuerzel);
        const tr = document.createElement('tr');
        if (!isUsed) tr.style.opacity = '0.6';
        const td1 = document.createElement('td');
        const strong = document.createElement('strong');
        strong.textContent = kuerzel;
        td1.appendChild(strong);
        const td2 = document.createElement('td');
        td2.textContent = email;
        const td3 = document.createElement('td');
        const span = document.createElement('span');
        span.style.color = isUsed ? '#28a745' : '#6c757d';
        span.textContent = isUsed ? '✓ Wird verwendet' : 'Nicht benötigt';
        td3.appendChild(span);
        const td4 = document.createElement('td');
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'btn btn-small btn-danger kt-delete-btn';
        btn.textContent = 'X';
        btn.title = 'Zuordnung löschen';
        btn.setAttribute('aria-label', 'Zuordnung löschen');
        btn.addEventListener('click', () => ns.removeTeacherMapping(kuerzel));
        td4.appendChild(btn);
        tr.append(td1, td2, td3, td4);
        tbody.appendChild(tr);
    });
};

ns.toggleTeacherMapping = function toggleTeacherMapping() {
    const table = document.getElementById('teacherMappingTable');
    table.style.display = table.style.display === 'none' ? 'block' : 'none';
};

ns.clearTeacherMapping = function clearTeacherMapping() {
    ns.confirmModal('Zuordnungen löschen', 'Alle Lehrer-Zuordnungen wirklich löschen?', () => {
        ns.teacherEmailMapping = {};
        document.getElementById('teacherMappingInfo').style.display = 'none';
        document.getElementById('teacherMappingTable').style.display = 'none';
        if (ns.currentStep === 4) ns.updateTeacherStats();
    });
};

ns.removeTeacherMapping = function removeTeacherMapping(kuerzel) {
    delete ns.teacherEmailMapping[kuerzel];
    document.getElementById('teacherCount').textContent = Object.keys(ns.teacherEmailMapping).length;
    // Löschungen werden bewusst NICHT automatisch in den Schul‑Grundeinstellungen gespiegelt,
    // um dort nicht versehentlich Daten zu verlieren. (Schul‑Einstellungen bleiben die Basisquelle.)
    if (ns.currentStep === 4) ns.updateTeacherStats();
    else ns.displayTeacherMappingTable();
    if (Object.keys(ns.teacherEmailMapping).length === 0) {
        document.getElementById('teacherMappingInfo').style.display = 'none';
        document.getElementById('teacherMappingTable').style.display = 'none';
    }
};

ns.addTeacherMapping = function addTeacherMapping() {
    ns.openModal(
        'Lehrer-Zuordnung hinzufügen',
        '<label for="modalKuerzel">Lehrer-Kürzel</label><input type="text" id="modalKuerzel" autocomplete="off">' +
            '<label for="modalEmail">E-Mail-Adresse</label><input type="email" id="modalEmail" autocomplete="off">',
        () => {
            const k = document.getElementById('modalKuerzel').value.trim();
            const em = document.getElementById('modalEmail').value.trim().toLowerCase();
            if (!k || !em) {
                ns.showToast('Bitte Kürzel und E-Mail ausfüllen.');
                return;
            }
            ns.teacherEmailMapping[k.toUpperCase()] = em;
            document.getElementById('teacherCount').textContent = Object.keys(ns.teacherEmailMapping).length;
            document.getElementById('teacherMappingInfo').style.display = 'block';
            ns.closeModal();
            upsertTenantTeachersFromMapping({ [k.toUpperCase()]: em });
            if (ns.currentStep === 4) ns.updateTeacherStats();
            else ns.displayTeacherMappingTable();
        }
    );
};

ns.importTeacherEmailsFromPaste = function importTeacherEmailsFromPaste() {
    const textarea = document.getElementById('teacherEmailPasteInput');
    if (!textarea) return;
    const lines = textarea.value.split(/\r?\n/);
    let added = 0;
    let invalid = 0;
    const delta = {};
    lines.forEach(line => {
        const parsed = parseTeacherEmailPasteLine(line);
        if (!parsed) {
            if (String(line).trim() && !String(line).trim().startsWith('#')) invalid++;
            return;
        }
        ns.teacherEmailMapping[parsed.kuerzel] = parsed.email;
        delta[parsed.kuerzel] = parsed.email;
        added++;
    });
    document.getElementById('teacherCount').textContent = Object.keys(ns.teacherEmailMapping).length;
    document.getElementById('teacherMappingInfo').style.display = 'block';
    if (added > 0) upsertTenantTeachersFromMapping(delta);
    if (ns.currentStep === 4) ns.updateTeacherStats();
    else ns.displayTeacherMappingTable();
    if (added > 0) {
        ns.showToast(
            invalid > 0
                ? `${added} Zuordnung(en) übernommen (${invalid} Zeile(n) übersprungen).`
                : `${added} Zuordnung(en) übernommen.`
        );
    } else if (invalid > 0) {
        ns.showToast('Keine gültigen Zeilen – pro Zeile Kürzel und E-Mail (Tab, Semikolon oder Leerzeichen).');
    } else {
        ns.showToast('Nichts eingefügt – bitte Zeilen mit Kürzel und E-Mail eintragen.');
    }
};

ns.downloadTeacherLehrerTemplateCsv = function downloadTeacherLehrerTemplateCsv() {
    const csv =
        '\uFEFFKürzel;E-Mail\n' +
        'MU;max.mustermann@schule.de\n' +
        'BME;anna.beispiel@schule.de\n';
    ns.downloadBlob('Lehrerliste-Vorlage.csv', csv, 'text/csv;charset=utf-8');
    ns.showToast('CSV-Vorlage heruntergeladen.');
};

ns.downloadTeacherLehrerTemplateXlsx = function downloadTeacherLehrerTemplateXlsx() {
    if (typeof XLSX === 'undefined' || !XLSX.utils || !XLSX.writeFile) {
        ns.showToast('Excel-Bibliothek nicht geladen – Seite neu laden.');
        return;
    }
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([
        ['Kürzel', 'E-Mail'],
        ['MU', 'max.mustermann@schule.de'],
        ['BME', 'anna.beispiel@schule.de']
    ]);
    XLSX.utils.book_append_sheet(wb, ws, 'Lehrer');
    XLSX.writeFile(wb, 'Lehrerliste-Vorlage.xlsx');
    ns.showToast('Excel-Vorlage heruntergeladen.');
};

function handleTeacherFile(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const name = (file.name || '').toLowerCase();
            let jsonData;
            if (name.endsWith('.csv')) {
                const buf = new Uint8Array(e.target.result);
                const textUtf8 = new TextDecoder('utf-8').decode(buf);
                jsonData = parseTeacherCsvToJsonData(textUtf8);
                ns.processTeacherMapping(jsonData, { replace: true });
                let n = Object.keys(ns.teacherEmailMapping).length;
                if (n === 0) {
                    const manual = parseTeacherListCsvLineByLine(textUtf8);
                    if (manual.length) ns.processTeacherMapping(manual, { replace: true });
                    n = Object.keys(ns.teacherEmailMapping).length;
                }
                if (n === 0) {
                    try {
                        const text1252 = new TextDecoder('windows-1252').decode(buf);
                        jsonData = parseTeacherCsvToJsonData(text1252);
                        ns.processTeacherMapping(jsonData, { replace: true });
                        n = Object.keys(ns.teacherEmailMapping).length;
                        if (n === 0) {
                            const manual2 = parseTeacherListCsvLineByLine(text1252);
                            if (manual2.length) ns.processTeacherMapping(manual2, { replace: true });
                            n = Object.keys(ns.teacherEmailMapping).length;
                        }
                    } catch (encErr) {
                        /* windows-1252 u. a. nicht verfügbar */
                    }
                }
                ns.showToast(
                    n
                        ? `Lehrer-Liste: ${n} Zuordnung(en) aus Datei geladen.`
                        : 'Datei enthielt keine gültigen Zeilen (Spalten Kürzel und E-Mail).'
                );
            } else {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                jsonData = XLSX.utils.sheet_to_json(firstSheet);
                ns.processTeacherMapping(jsonData, { replace: true });
                const n = Object.keys(ns.teacherEmailMapping).length;
                ns.showToast(
                    n
                        ? `Lehrer-Liste: ${n} Zuordnung(en) aus Datei geladen.`
                        : 'Datei enthielt keine gültigen Zeilen (Spalten Kürzel und E-Mail).'
                );
            }
        } catch (error) {
            ns.showToast('Fehler beim Lesen der Lehrer-Datei: ' + error.message);
        }
    };
    reader.readAsArrayBuffer(file);
}

// Upload wiring (robust: auch wenn DOM-Cache leer ist)
{
    const area = (ns.dom && ns.dom.teacherUploadArea) || document.getElementById('teacherUploadArea');
    const input = (ns.dom && ns.dom.teacherFileInput) || document.getElementById('teacherFileInput');
    if (area && input) {
        area.addEventListener('click', () => input.click());

        const onDragOver = (e) => {
            e.preventDefault();
            e.stopPropagation();
            try {
                if (e.dataTransfer) e.dataTransfer.dropEffect = 'copy';
            } catch {
                // ignore
            }
            area.classList.add('dragover');
        };
        const onDragLeave = (e) => {
            e.preventDefault();
            e.stopPropagation();
            area.classList.remove('dragover');
        };
        const onDrop = (e) => {
            e.preventDefault();
            e.stopPropagation();
            area.classList.remove('dragover');
            const files = (e.dataTransfer && e.dataTransfer.files) || null;
            if (files && files.length > 0) handleTeacherFile(files[0]);
        };

        area.addEventListener('dragenter', onDragOver);
        area.addEventListener('dragover', onDragOver);
        area.addEventListener('dragleave', onDragLeave);
        area.addEventListener('drop', onDrop);

        input.addEventListener('change', (e) => {
            if (e.target.files.length > 0) handleTeacherFile(e.target.files[0]);
        });
    }
}

// Global exports für HTML onclick
window.toggleTeacherMapping = ns.toggleTeacherMapping;
window.clearTeacherMapping = ns.clearTeacherMapping;
window.addTeacherMapping = ns.addTeacherMapping;
window.importTeacherEmailsFromPaste = ns.importTeacherEmailsFromPaste;
window.downloadTeacherLehrerTemplateCsv = ns.downloadTeacherLehrerTemplateCsv;
window.downloadTeacherLehrerTemplateXlsx = ns.downloadTeacherLehrerTemplateXlsx;

// Schul‑Grundeinstellungen: Lehrerliste optional vorbefüllen (falls vorhanden)
loadTenantTeacherEmailsIfEmpty();

