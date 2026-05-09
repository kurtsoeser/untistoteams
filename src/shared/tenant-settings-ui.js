(function () {
    'use strict';

    function normStr(v) {
        return String(v ?? '').trim();
    }

    function dlgAlert(msg, opts) {
        if (typeof window.ms365AppDialogAlert === 'function') {
            return window.ms365AppDialogAlert(msg, opts);
        }
        window.alert(msg);
        return Promise.resolve();
    }

    function dlgConfirm(msg, opts) {
        if (typeof window.ms365AppDialogConfirm === 'function') {
            return window.ms365AppDialogConfirm(msg, opts);
        }
        return Promise.resolve(window.confirm(msg));
    }

    function dlgPrompt(msg, def, opts) {
        if (typeof window.ms365AppDialogPrompt === 'function') {
            return window.ms365AppDialogPrompt(msg, def, opts);
        }
        return Promise.resolve(window.prompt(msg, def));
    }

    function normCode(v) {
        return normStr(v).toUpperCase();
    }

    function safeJsonParse(s) {
        try {
            return JSON.parse(String(s));
        } catch {
            return null;
        }
    }

    function normHeaderKey(k) {
        return String(k ?? '')
            .trim()
            .toLowerCase()
            .replace(/\s+/g, '')
            .replace(/ä/g, 'ae')
            .replace(/ö/g, 'oe')
            .replace(/ü/g, 'ue')
            .replace(/ß/g, 'ss')
            .replace(/[^a-z0-9]/g, '');
    }

    function getField(row, candidates) {
        if (!row || typeof row !== 'object') return '';
        const map = new Map();
        Object.keys(row).forEach((k) => map.set(normHeaderKey(k), row[k]));
        for (const c of candidates) {
            const v = map.get(normHeaderKey(c));
            if (v != null && String(v).trim() !== '') return String(v).trim();
        }
        return '';
    }

    function ensureXlsxReady() {
        return typeof XLSX !== 'undefined' && XLSX.utils && typeof XLSX.read === 'function';
    }

    function sheetToJsonRows(workbook) {
        const sheetName = workbook.SheetNames && workbook.SheetNames[0];
        if (!sheetName) return [];
        const sheet = workbook.Sheets[sheetName];
        if (!sheet) return [];
        return XLSX.utils.sheet_to_json(sheet, { defval: '' });
    }

    function parseCsvTextToJsonRows(text) {
        if (!ensureXlsxReady()) return [];
        let s = String(text || '');
        if (s.charCodeAt(0) === 0xfeff) s = s.slice(1);
        let wb = XLSX.read(s, { type: 'string', FS: ';' });
        let rows = sheetToJsonRows(wb);
        if (!rows.length) {
            wb = XLSX.read(s, { type: 'string', FS: ',' });
            rows = sheetToJsonRows(wb);
        }
        return rows;
    }

    function downloadXlsxTemplate(filename, aoa, sheetName) {
        if (!ensureXlsxReady() || typeof XLSX.writeFile !== 'function') return false;
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(aoa);
        XLSX.utils.book_append_sheet(wb, ws, sheetName || 'Daten');
        XLSX.writeFile(wb, filename);
        return true;
    }

    /** @param {{ name: string, aoa: any[][] }[]} sheets */
    function downloadXlsxMultiSheet(filename, sheets) {
        if (!ensureXlsxReady() || typeof XLSX.writeFile !== 'function') return false;
        const wb = XLSX.utils.book_new();
        (sheets || []).forEach((sh) => {
            const rawName = String(sh.name || 'Daten').replace(/[:\\/?*[\]]/g, '-');
            const sn = rawName.slice(0, 31) || 'Daten';
            const ws = XLSX.utils.aoa_to_sheet(sh.aoa || []);
            XLSX.utils.book_append_sheet(wb, ws, sn);
        });
        XLSX.writeFile(wb, filename);
        return true;
    }

    /**
     * Generischer Excel/CSV-Import → JSON-Zeilen (erstes Arbeitsblatt).
     * @param {File} file
     * @param {(rows: object[]) => void} onRows
     * @param {(msg: string) => void} [onError]
     */
    function importSpreadsheetFileToJsonRows(file, onRows, onError) {
        if (!file) return;
        if (!ensureXlsxReady()) {
            if (onError) onError('Import: Excel-Bibliothek nicht geladen – Seite neu laden.');
            return;
        }
        const name = String(file.name || '').toLowerCase();
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                let jsonRows = [];
                if (name.endsWith('.csv')) {
                    const buf = new Uint8Array(e.target.result);
                    const tryDecoders = ['utf-8', 'windows-1252'];
                    for (const enc of tryDecoders) {
                        try {
                            const text = new TextDecoder(enc).decode(buf);
                            jsonRows = parseCsvTextToJsonRows(text);
                            if (jsonRows.length) break;
                        } catch {
                            // ignore
                        }
                    }
                } else {
                    const data = new Uint8Array(e.target.result);
                    const wb = XLSX.read(data, { type: 'array' });
                    jsonRows = sheetToJsonRows(wb);
                }
                onRows(jsonRows || []);
            } catch (err) {
                if (onError) onError('Import fehlgeschlagen: ' + (err?.message || String(err)));
            }
        };
        reader.readAsArrayBuffer(file);
    }

    /** Lehrer-Zeilen wie in der Textarea: Kürzel;Name;E-Mail pro Zeile */
    function teacherJsonRowsToSemicolonLines(jsonRows) {
        const out = [];
        (jsonRows || []).forEach((r) => {
            const code = getField(r, ['kürzel', 'kuerzel', 'code', 'lehrer', 'abbrev', 'abbreviation']);
            let name = getField(r, ['name', 'lehrername', 'anzeigename', 'displayname']);
            let email = getField(r, ['e-mail', 'email', 'mail', 'upn']);
            const c = normCode(code);
            if (!c) return;

            const nameNorm = normStr(name);
            const emailNorm = normStr(email).toLowerCase();
            const nameLooksLikeEmail = nameNorm.includes('@');
            const emailLooksLikeEmail = emailNorm.includes('@');

            if (nameLooksLikeEmail && (!emailNorm || !emailLooksLikeEmail)) {
                email = nameNorm;
                name = '';
            }

            out.push({ code: c, name: normStr(name), email: normStr(email).toLowerCase() });
        });
        return out.map((x) => [x.code, x.name || '', x.email || ''].filter(Boolean).join(';')).join('\n');
    }

    window.ms365TeacherListImport = {
        isXlsxReady: ensureXlsxReady,
        downloadTemplate() {
            return downloadXlsxTemplate(
                'Lehrerliste-Vorlage.xlsx',
                [
                    ['Kürzel', 'Name', 'E-Mail'],
                    ['MU', 'Max Mustermann', 'max.mustermann@schule.de'],
                    ['BME', 'Anna Beispiel', 'anna.beispiel@schule.de']
                ],
                'Lehrer'
            );
        },
        downloadCsvTemplate() {
            try {
                const BOM = '\ufeff';
                const body = ['Kürzel;Name;E-Mail', 'MU;Max Mustermann;max.mustermann@schule.de'].join('\r\n');
                const blob = new Blob([BOM + body], { type: 'text/csv;charset=utf-8' });
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = 'Lehrerliste-Vorlage.csv';
                document.body.appendChild(a);
                a.click();
                a.remove();
                setTimeout(() => URL.revokeObjectURL(url), 250);
                return true;
            } catch {
                return false;
            }
        },
        importFile(file, onLines, onError) {
            importSpreadsheetFileToJsonRows(
                file,
                (rows) => {
                    if (onLines) onLines(teacherJsonRowsToSemicolonLines(rows));
                },
                onError
            );
        }
    };

    /** Schüler-Zeilen: Klasse;Name;E-Mail */
    function studentJsonRowsToSemicolonLines(jsonRows) {
        const out = [];
        (jsonRows || []).forEach((r) => {
            const klasse = getField(r, ['klasse', 'class', 'zug', 'gruppe', 'k']);
            let name = getField(r, ['name', 'schueler', 'schüler', 'vollname', 'displayname']);
            let email = getField(r, ['e-mail', 'email', 'mail', 'upn']);
            const nameNorm = normStr(name);
            const emailNorm = normStr(email).toLowerCase();
            const nameLooksLikeEmail = nameNorm.includes('@');
            const emailLooksLikeEmail = emailNorm.includes('@');
            if (nameLooksLikeEmail && (!emailNorm || !emailLooksLikeEmail)) {
                email = nameNorm;
                name = '';
            }
            const k = normStr(klasse);
            if (!k && !normStr(name) && !normStr(email)) return;
            out.push({ klasse: k, name: normStr(name), email: normStr(email).toLowerCase() });
        });
        return out.map((x) => [x.klasse, x.name || '', x.email || ''].filter(Boolean).join(';')).join('\n');
    }

    function adminJsonRowsToSemicolonLines(jsonRows) {
        const out = [];
        (jsonRows || []).forEach((r) => {
            const role = getField(r, ['rolle', 'role', 'position', 'funktion']);
            const name = getField(r, ['name', 'anzeigename', 'displayname']);
            const email = getField(r, ['e-mail', 'email', 'mail', 'upn']);
            if (!normStr(role) && !normStr(name) && !normStr(email)) return;
            out.push({ role: normStr(role), name: normStr(name), email: normStr(email).toLowerCase() });
        });
        return out.map((x) => [x.role, x.name || '', x.email || ''].filter(Boolean).join(';')).join('\n');
    }

    function subjectsJsonRowsToSemicolonLines(jsonRows) {
        const out = [];
        (jsonRows || []).forEach((r) => {
            const code = getField(r, ['kürzel', 'kuerzel', 'code', 'fach', 'abbrev']);
            const name = getField(r, ['name', 'bezeichnung', 'fachname', 'displayname']);
            const c = normCode(code);
            if (!c) return;
            out.push({ code: c, name: normStr(name) });
        });
        return out.map((x) => [x.code, x.name || ''].join(';')).join('\n');
    }

    function argesJsonRowsToSemicolonLines(jsonRows) {
        const out = [];
        (jsonRows || []).forEach((r) => {
            const code = getField(r, ['kürzel', 'kuerzel', 'code', 'arge', 'kuerzelarge']);
            const name = getField(r, ['name', 'bezeichnung', 'displayname']);
            const subj = getField(r, ['fächer', 'faecher', 'subjects', 'fachzuordnung', 'faecherkuerzel']);
            const c = normCode(code);
            if (!c) return;
            let line = c + ';' + normStr(name);
            if (normStr(subj)) line += ';' + normStr(subj);
            out.push(line);
        });
        return out.join('\n');
    }

    function classesJsonRowsToSemicolonLines(jsonRows) {
        const out = [];
        (jsonRows || []).forEach((r) => {
            const code = getField(r, ['kürzel', 'kuerzel', 'code', 'klassekurz']);
            const year = getField(r, ['abschlussjahr', 'jahr', 'year', 'jahrgang']);
            const name = normStr(getField(r, ['anzeigename', 'klasse', 'name', 'displayname']));
            const headName = normStr(getField(r, ['kvname', 'klassenvorstand', 'vorstand', 'headname']));
            const headEmail = normStr(getField(r, ['kvmail', 'kv-email', 'e-mailkv', 'heademail', 'email'])).toLowerCase();
            const c = normCode(code);
            const y = /^\d{4}$/.test(normStr(year)) ? normStr(year) : '';
            if (!c && !name && !y && !headName && !headEmail) return;
            if (y) {
                out.push([c, y, name, headName, headEmail].join(';'));
            } else {
                out.push([c, name, headName, headEmail].join(';'));
            }
        });
        return out.join('\n');
    }

    function findWorksheetName(wb, aliases) {
        const names = wb.SheetNames || [];
        const want = (aliases || []).map((a) => normHeaderKey(a));
        for (const sn of names) {
            const nk = normHeaderKey(sn);
            if (want.indexOf(nk) >= 0) return sn;
        }
        for (const sn of names) {
            const nk = normHeaderKey(sn);
            for (const w of want) {
                if (!w) continue;
                if (nk === w || nk.indexOf(w) === 0 || w.indexOf(nk) === 0) return sn;
            }
        }
        return null;
    }

    function importFileToWorkbook(file, onWorkbook, onError) {
        if (!file) return;
        if (!ensureXlsxReady()) {
            if (onError) onError('Import: Excel-Bibliothek nicht geladen – Seite neu laden.');
            return;
        }
        const name = String(file.name || '').toLowerCase();
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                let wb;
                if (name.endsWith('.csv')) {
                    let s = String(e.target.result || '');
                    if (s.charCodeAt(0) === 0xfeff) s = s.slice(1);
                    wb = XLSX.read(s, { type: 'string', FS: ';' });
                    if (!sheetToJsonRows(wb).length) wb = XLSX.read(s, { type: 'string', FS: ',' });
                } else {
                    wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
                }
                if (onWorkbook) onWorkbook(wb);
            } catch (err) {
                if (onError) onError('Import fehlgeschlagen: ' + (err?.message || String(err)));
            }
        };
        if (name.endsWith('.csv')) reader.readAsText(file, 'utf-8');
        else reader.readAsArrayBuffer(file);
    }

    window.ms365StudentListImport = {
        isXlsxReady: ensureXlsxReady,
        downloadTemplate() {
            return downloadXlsxTemplate(
                'Schuelerliste-Vorlage.xlsx',
                [
                    ['Klasse', 'Name', 'E-Mail'],
                    ['1AK', 'Lisa Beispiel', 'lisa.beispiel@schule.de'],
                    ['1AK', 'Max Muster', 'max.muster@schule.de']
                ],
                'Schueler'
            );
        },
        downloadCsvTemplate() {
            try {
                const BOM = '\ufeff';
                const body = ['Klasse;Name;E-Mail', '1AK;Lisa Beispiel;lisa.beispiel@schule.de'].join('\r\n');
                const blob = new Blob([BOM + body], { type: 'text/csv;charset=utf-8' });
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = 'Schuelerliste-Vorlage.csv';
                document.body.appendChild(a);
                a.click();
                a.remove();
                setTimeout(() => URL.revokeObjectURL(url), 250);
                return true;
            } catch {
                return false;
            }
        },
        importFile(file, onLines, onError) {
            importSpreadsheetFileToJsonRows(
                file,
                (rows) => {
                    if (onLines) onLines(studentJsonRowsToSemicolonLines(rows));
                },
                onError
            );
        }
    };

    window.ms365SchuldatenMasterImport = {
        isXlsxReady: ensureXlsxReady,
        downloadTemplate() {
            return downloadXlsxMultiSheet('MS365-Schuldaten-Vorlage.xlsx', [
                {
                    name: 'Verwaltung',
                    aoa: [
                        ['Rolle', 'Name', 'E-Mail'],
                        ['Direktion', 'Direktorin Beispiel', 'direktion@schule.de'],
                        ['Sekretariat', 'Sekretariat', 'sekretariat@schule.de']
                    ]
                },
                {
                    name: 'Lehrer',
                    aoa: [
                        ['Kürzel', 'Name', 'E-Mail'],
                        ['MU', 'Max Mustermann', 'max.mustermann@schule.de'],
                        ['BME', 'Anna Beispiel', 'anna.beispiel@schule.de']
                    ]
                },
                {
                    name: 'Schueler',
                    aoa: [
                        ['Klasse', 'Name', 'E-Mail'],
                        ['1AK', 'Lisa Beispiel', 'lisa.beispiel@schule.de'],
                        ['1AK', 'Max Muster', 'max.muster@schule.de']
                    ]
                },
                {
                    name: 'Faecher',
                    aoa: [
                        ['Kürzel', 'Name'],
                        ['D', 'Deutsch'],
                        ['M', 'Mathematik'],
                        ['E', 'Englisch']
                    ]
                },
                {
                    name: 'ARGE',
                    aoa: [
                        ['Kürzel', 'Name', 'Fächer'],
                        ['SPRACHEN', 'Sprachen', 'D,E'],
                        ['NAWI', 'Naturwissenschaften', 'BIO,CH,PH']
                    ]
                },
                {
                    name: 'Klassen',
                    aoa: [
                        ['Kürzel', 'Abschlussjahr', 'Anzeigename', 'KV-Name', 'KV-E-Mail'],
                        ['HMA', '2031', '1HMA', 'Max Mustermann', 'max.mustermann@schule.de'],
                        ['1AK', '2030', '1A-Klasse', 'Anna Beispiel', 'anna.beispiel@schule.de']
                    ]
                }
            ]);
        },
        importFile(file, onPayload, onError) {
            const name = String(file.name || '').toLowerCase();
            if (name.endsWith('.csv')) {
                if (onError) onError('Gesamt-Import: Bitte die XLSX-Vorlage verwenden (mehrere Arbeitsblätter).');
                return;
            }
            importFileToWorkbook(
                file,
                (wb) => {
                    try {
                        const out = {};
                        const snV = findWorksheetName(wb, ['verwaltung', 'administration']);
                        const snL = findWorksheetName(wb, ['lehrer', 'lehrerinnen', 'teachers']);
                        const snS = findWorksheetName(wb, ['schueler', 'schuler', 'schüler', 'students', 'schuelerinnen']);
                        const snF = findWorksheetName(wb, ['faecher', 'fächer', 'subjects']);
                        const snA = findWorksheetName(wb, ['arge', 'arbeitsgruppen', 'arbeitsgemeinschaften']);
                        const snK = findWorksheetName(wb, ['klassen', 'classes']);
                        function sheetLines(sn, conv) {
                            if (!sn || !wb.Sheets[sn]) return null;
                            const rows = XLSX.utils.sheet_to_json(wb.Sheets[sn], { defval: '' });
                            if (!rows || !rows.length) return null;
                            const lines = conv(rows);
                            return normStr(lines).length ? lines : null;
                        }
                        const v = sheetLines(snV, adminJsonRowsToSemicolonLines);
                        if (v != null) out.verwaltungLines = v;
                        const l = sheetLines(snL, teacherJsonRowsToSemicolonLines);
                        if (l != null) out.lehrerLines = l;
                        const s = sheetLines(snS, studentJsonRowsToSemicolonLines);
                        if (s != null) out.schuelerLines = s;
                        const f = sheetLines(snF, subjectsJsonRowsToSemicolonLines);
                        if (f != null) out.faecherLines = f;
                        const a = sheetLines(snA, argesJsonRowsToSemicolonLines);
                        if (a != null) out.argeLines = a;
                        const k = sheetLines(snK, classesJsonRowsToSemicolonLines);
                        if (k != null) out.klassenLines = k;
                        if (onPayload) onPayload(out);
                    } catch (err) {
                        if (onError) onError(String(err?.message || err));
                    }
                },
                onError
            );
        }
    };

    // UI binding (optional; nur wenn Elemente existieren)
    function bindUi() {
        const form = document.getElementById('tenantSettingsForm');
        if (!form) return;

        if (typeof window.ms365TenantSettingsLoad !== 'function' || typeof window.ms365TenantSettingsSave !== 'function') {
            return;
        }

        const parseLinesToSubjects = window.ms365TenantSettingsParseSubjectsLines;
        const parseLinesToArges = window.ms365TenantSettingsParseArgesLines;
        const parseLinesToTeachers = window.ms365TenantSettingsParseTeachersLines;
        const parseLinesToAdmin = window.ms365TenantSettingsParseAdminLines;
        const parseLinesToStudents = window.ms365TenantSettingsParseStudentsLines;
        const parseLinesToClasses = window.ms365TenantSettingsParseClassesLines;
        const load = window.ms365TenantSettingsLoad;
        const save = window.ms365TenantSettingsSave;

        const taSubjects = document.getElementById('tenantSubjectsLines');
        const subjectsTbody = document.getElementById('tenantSubjectsTableBody');
        const btnAddSubjectRow = document.getElementById('tenantSubjectsAddRow');
        const taArges = document.getElementById('tenantArgesLines');
        const argesTbody = document.getElementById('tenantArgesTableBody');
        const btnAddArgeRow = document.getElementById('tenantArgesAddRow');
        const taTeachers = document.getElementById('tenantTeachersLines');
        const teachersTbody = document.getElementById('tenantTeachersTableBody');
        const btnAddTeacherRow = document.getElementById('tenantTeachersAddRow');
        const taAdmin = document.getElementById('tenantAdminLines');
        const adminTbody = document.getElementById('tenantAdminTableBody');
        const btnAddAdminRow = document.getElementById('tenantAdminAddRow');
        const taStudents = document.getElementById('tenantStudentsLines');
        const studentsTbody = document.getElementById('tenantStudentsTableBody');
        const btnAddStudentRow = document.getElementById('tenantStudentsAddRow');
        const taClasses = document.getElementById('tenantClassesLines');
        const classesTbody = document.getElementById('tenantClassesTableBody');
        const btnAddClassRow = document.getElementById('tenantClassesAddRow');
        const fileSubjects = document.getElementById('tenantSubjectsImportFile');
        const fileArges = document.getElementById('tenantArgesImportFile');
        const fileTeachers = document.getElementById('tenantTeachersImportFile');
        const fileStudents = document.getElementById('tenantStudentsImportFile');
        const fileClasses = document.getElementById('tenantClassesImportFile');
        const btnSubjectsTpl = document.getElementById('tenantSubjectsTemplateXlsx');
        const btnArgesTpl = document.getElementById('tenantArgesTemplateXlsx');
        const btnTeachersTpl = document.getElementById('tenantTeachersTemplateXlsx');
        const btnStudentsTpl = document.getElementById('tenantStudentsTemplateXlsx');
        const btnClassesTpl = document.getElementById('tenantClassesTemplateXlsx');
        const btnSave = document.getElementById('tenantSettingsSave');
        const btnReload = document.getElementById('tenantSettingsReload');
        const btnExport = document.getElementById('tenantSettingsExport');
        const btnExportHeader = document.getElementById('tenantSettingsExportHeader');
        const fileImport = document.getElementById('tenantSettingsImportFile');
        const btnClear = document.getElementById('tenantSettingsClear');
        const summary = document.getElementById('tenantSettingsSummary');
        const inpDefaultGradYear = null;
        const domainInput = document.getElementById('schoolEmailDomain');
        const schoolYearSelect = document.getElementById('schoolYearSelect');
        const schoolYearAddBtn = document.getElementById('schoolYearAddBtn');

        function currentSchoolYearLabel() {
            const y = new Date().getFullYear();
            return String(y) + '/' + String(y + 1).slice(2);
        }

        function getDisplayedSchoolYearLabel() {
            try {
                if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getContainer === 'function') {
                    const c = window.ms365AppDataV2.getContainer();
                    const cur = c && c.years ? String(c.years.current || '').trim() : '';
                    if (cur) return cur;
                }
            } catch {
                // ignore
            }
            if (schoolYearSelect) return String(schoolYearSelect.value || '').trim();
            return '';
        }

        function seedDemoDataIfEmptyStorage() {
            // Demo-Daten nur beim allerersten Start (wenn noch nichts gespeichert ist)
            try {
                const raw = localStorage.getItem('ms365-tenant-settings-v1');
                if (raw) return false;
            } catch {
                // wenn localStorage nicht lesbar ist: nicht seeden
                return false;
            }

            const demo = {
                domain: 'ms365.schule',
                subjects: [
                    { code: 'M', name: 'Mathematik' },
                    { code: 'D', name: 'Deutsch' },
                    { code: 'E', name: 'Englisch' }
                ],
                arges: [
                    { code: 'SPRACHEN', name: 'Sprachen', subjects: ['D', 'E'] },
                    { code: 'NAWI', name: 'Naturwissenschaften', subjects: ['BIO', 'CH', 'PH'] }
                ],
                teachers: [
                    { code: 'LEH', name: 'Vorname Lehrer', email: 'vorname.lehrer@ms365.schule' },
                    { code: 'MUS', name: 'Max Muster', email: 'max.muster@ms365.schule' }
                ],
                students: [
                    { klasse: '1A', name: 'Anna Beispiel', email: 'anna.beispiel@ms365.schule' },
                    { klasse: '1A', name: 'Ben Demo', email: 'ben.demo@ms365.schule' },
                    { klasse: '1B', name: 'Carla Test', email: 'carla.test@ms365.schule' },
                    { klasse: '2A', name: 'David Probe', email: 'david.probe@ms365.schule' },
                    { klasse: '2A', name: 'Eva Sample', email: 'eva.sample@ms365.schule' }
                ],
                classes: [
                    { code: '1A', year: '2030', name: 'Klasse 1A', headName: 'Vorname Lehrer', headEmail: 'vorname.lehrer@ms365.schule' },
                    { code: '1B', year: '2030', name: 'Klasse 1B', headName: 'Max Muster', headEmail: 'max.muster@ms365.schule' },
                    { code: '2A', year: '2030', name: 'Klasse 2A', headName: 'Vorname Lehrer', headEmail: 'vorname.lehrer@ms365.schule' }
                ]
            };

            const saved = save(demo);
            // Domain auch in der UI sichtbar machen
            try {
                if (typeof window.ms365SetSchoolDomainNoAt === 'function') {
                    window.ms365SetSchoolDomainNoAt(saved.domain);
                }
            } catch {
                // ignore
            }
            return true;
        }

        let autoSaveTimer = null;
        let __syncGuard = 0;

        function dispatchTenantSettingsChanged(saved, reason) {
            try {
                if (__syncGuard) return;
                window.dispatchEvent(
                    new CustomEvent('ms365-tenant-settings-changed', {
                        detail: { settings: saved, reason: String(reason || '') }
                    })
                );
            } catch {
                // ignore
            }
        }

        function autoSaveNow() {
            const subjects = typeof parseLinesToSubjects === 'function' ? parseLinesToSubjects(taSubjects ? taSubjects.value : '') : [];
            const arges = typeof parseLinesToArges === 'function' ? parseLinesToArges(taArges ? taArges.value : '') : [];
            const teachers = typeof parseLinesToTeachers === 'function' ? parseLinesToTeachers(taTeachers ? taTeachers.value : '') : [];
            const admin = typeof parseLinesToAdmin === 'function' ? parseLinesToAdmin(taAdmin ? taAdmin.value : '') : [];
            const students = typeof parseLinesToStudents === 'function' ? parseLinesToStudents(taStudents ? taStudents.value : '') : [];
            const classes = typeof parseLinesToClasses === 'function' ? parseLinesToClasses(taClasses ? taClasses.value : '') : [];
            const domain =
                typeof window.ms365GetSchoolDomainNoAt === 'function' ? window.ms365GetSchoolDomainNoAt() : '';
            const saved = save({ domain, subjects, arges, teachers, admin, students, classes });
            dispatchTenantSettingsChanged(saved, 'autosave');
        }

        function argesToLines(rows) {
            return (rows || [])
                .map((x) => {
                    const list = Array.isArray(x.subjects) ? x.subjects : [];
                    return `${normCode(x.code)};${normStr(x.name || '')};${list.map((s) => normCode(s)).filter(Boolean).join(',')}`.trim();
                })
                .filter(Boolean)
                .join('\n');
        }

        function getArgesFromTextarea() {
            return typeof parseLinesToArges === 'function' ? parseLinesToArges(taArges ? taArges.value : '') : [];
        }

        function setArgesTextareaFromRows(rows) {
            if (!taArges) return;
            taArges.value = argesToLines(rows);
        }

        function renderArgesTableFromTextarea() {
            if (!argesTbody) return;
            const rows = getArgesFromTextarea();
            argesTbody.replaceChildren();

            if (!rows.length) {
                const tr = document.createElement('tr');
                const td = document.createElement('td');
                td.colSpan = 4;
                td.style.color = '#6c757d';
                td.textContent = 'Noch keine Einträge – oben einfügen oder „+ Zeile“.';
                tr.appendChild(td);
                argesTbody.appendChild(tr);
                return;
            }

            rows.forEach((row, idx) => {
                const tr = document.createElement('tr');

                const tdCode = document.createElement('td');
                tdCode.textContent = row.code || '';
                tdCode.title = 'Doppelklick zum Bearbeiten';
                tdCode.addEventListener('dblclick', () => {
                    startCellEdit(tdCode, row.code, (next, meta) => {
                        const all = getArgesFromTextarea();
                        if (!all[idx]) return renderArgesTableFromTextarea();
                        const prev = all[idx].code;
                        all[idx].code = meta && meta.cancelled ? prev : normCode(next);
                        setArgesTextareaFromRows(all);
                        renderArgesTableFromTextarea();
                        scheduleAutoSave();
                    });
                });

                const tdName = document.createElement('td');
                tdName.textContent = row.name || '';
                tdName.title = 'Doppelklick zum Bearbeiten';
                tdName.addEventListener('dblclick', () => {
                    startCellEdit(tdName, row.name, (next, meta) => {
                        const all = getArgesFromTextarea();
                        if (!all[idx]) return renderArgesTableFromTextarea();
                        const prev = all[idx].name;
                        all[idx].name = meta && meta.cancelled ? prev : normStr(next);
                        setArgesTextareaFromRows(all);
                        renderArgesTableFromTextarea();
                        scheduleAutoSave();
                    });
                });

                const tdSubjects = document.createElement('td');
                tdSubjects.textContent = (Array.isArray(row.subjects) ? row.subjects : []).join(', ');
                tdSubjects.title = 'Doppelklick zum Bearbeiten';
                tdSubjects.addEventListener('dblclick', () => {
                    startCellEdit(tdSubjects, (Array.isArray(row.subjects) ? row.subjects : []).join(','), (next, meta) => {
                        const all = getArgesFromTextarea();
                        if (!all[idx]) return renderArgesTableFromTextarea();
                        const prev = all[idx].subjects;
                        all[idx].subjects =
                            meta && meta.cancelled
                                ? prev
                                : String(next || '')
                                      .split(/[,\s|]+/)
                                      .map((x) => normCode(x))
                                      .filter(Boolean);
                        setArgesTextareaFromRows(all);
                        renderArgesTableFromTextarea();
                        scheduleAutoSave();
                    });
                });

                const tdAction = document.createElement('td');
                tdAction.className = 'action-cell';
                const btnDel = document.createElement('button');
                btnDel.type = 'button';
                btnDel.className = 'mini-btn';
                btnDel.textContent = '✕';
                btnDel.title = 'Zeile löschen';
                btnDel.addEventListener('click', () => {
                    const all = getArgesFromTextarea();
                    all.splice(idx, 1);
                    setArgesTextareaFromRows(all);
                    renderArgesTableFromTextarea();
                    scheduleAutoSave();
                });
                tdAction.appendChild(btnDel);

                tr.append(tdCode, tdName, tdSubjects, tdAction);
                argesTbody.appendChild(tr);
            });
        }

        function scheduleAutoSave() {
            if (autoSaveTimer) clearTimeout(autoSaveTimer);
            autoSaveTimer = setTimeout(() => {
                autoSaveTimer = null;
                try {
                    autoSaveNow();
                } catch {
                    // ignore (z.B. während Import/Reset)
                }
            }, 450);
        }

        function setSummary(text, kind) {
            if (!summary) return;
            summary.style.display = 'block';
            summary.textContent = text;
            summary.dataset.kind = kind || 'info';
        }

        function subjectsToLines(rows) {
            return (rows || [])
                .map((x) => `${normCode(x.code)};${normStr(x.name || '')}`.trim())
                .filter(Boolean)
                .join('\n');
        }

        function getSubjectsFromTextarea() {
            return typeof parseLinesToSubjects === 'function' ? parseLinesToSubjects(taSubjects ? taSubjects.value : '') : [];
        }

        function setSubjectsTextareaFromRows(rows) {
            if (!taSubjects) return;
            taSubjects.value = subjectsToLines(rows);
        }

        function renderSubjectsTableFromTextarea() {
            if (!subjectsTbody) return;
            const rows = getSubjectsFromTextarea();
            subjectsTbody.replaceChildren();

            if (!rows.length) {
                const tr = document.createElement('tr');
                const td = document.createElement('td');
                td.colSpan = 3;
                td.style.color = '#6c757d';
                td.textContent = 'Noch keine Einträge – oben einfügen oder „+ Zeile“.';
                tr.appendChild(td);
                subjectsTbody.appendChild(tr);
                return;
            }

            rows.forEach((row, idx) => {
                const tr = document.createElement('tr');

                const tdCode = document.createElement('td');
                tdCode.innerHTML = `<code>${row.code || ''}</code>`;
                tdCode.title = 'Doppelklick zum Bearbeiten';
                tdCode.addEventListener('dblclick', () => {
                    startCellEdit(tdCode, row.code, (next, meta) => {
                        const all = getSubjectsFromTextarea();
                        if (!all[idx]) return renderSubjectsTableFromTextarea();
                        const prev = all[idx].code;
                        all[idx].code = meta && meta.cancelled ? prev : normCode(next);
                        setSubjectsTextareaFromRows(all);
                        renderSubjectsTableFromTextarea();
                        scheduleAutoSave();
                    });
                });

                const tdName = document.createElement('td');
                tdName.textContent = row.name || '';
                tdName.title = 'Doppelklick zum Bearbeiten';
                tdName.addEventListener('dblclick', () => {
                    startCellEdit(tdName, row.name, (next, meta) => {
                        const all = getSubjectsFromTextarea();
                        if (!all[idx]) return renderSubjectsTableFromTextarea();
                        const prev = all[idx].name;
                        all[idx].name = meta && meta.cancelled ? prev : normStr(next);
                        setSubjectsTextareaFromRows(all);
                        renderSubjectsTableFromTextarea();
                        scheduleAutoSave();
                    });
                });

                const tdAction = document.createElement('td');
                tdAction.className = 'action-cell';
                const btnDel = document.createElement('button');
                btnDel.type = 'button';
                btnDel.className = 'mini-btn';
                btnDel.textContent = '✕';
                btnDel.title = 'Zeile löschen';
                btnDel.addEventListener('click', () => {
                    const all = getSubjectsFromTextarea();
                    all.splice(idx, 1);
                    setSubjectsTextareaFromRows(all);
                    renderSubjectsTableFromTextarea();
                    scheduleAutoSave();
                });
                tdAction.appendChild(btnDel);

                tr.append(tdCode, tdName, tdAction);
                subjectsTbody.appendChild(tr);
            });
        }

        function teachersToLines(rows) {
            return (rows || [])
                .map((x) => `${normCode(x.code)};${normStr(x.name || '')};${normStr(x.email || '').toLowerCase()}`.trim())
                .filter(Boolean)
                .join('\n');
        }

        function adminToLines(rows) {
            return (rows || [])
                .map((x) => `${normStr(x.role || '')};${normStr(x.name || '')};${normStr(x.email || '').toLowerCase()}`.trim())
                .filter(Boolean)
                .join('\n');
        }

        function getAdminFromTextarea() {
            return typeof parseLinesToAdmin === 'function' ? parseLinesToAdmin(taAdmin ? taAdmin.value : '') : [];
        }

        function setAdminTextareaFromRows(rows) {
            if (!taAdmin) return;
            taAdmin.value = adminToLines(rows);
        }

        function getTeachersFromTextarea() {
            return typeof parseLinesToTeachers === 'function' ? parseLinesToTeachers(taTeachers ? taTeachers.value : '') : [];
        }

        function setTeachersTextareaFromRows(rows) {
            if (!taTeachers) return;
            taTeachers.value = teachersToLines(rows);
        }

        function startCellEdit(td, initialValue, onCommit) {
            const prevText = String(initialValue ?? '');
            const input = document.createElement('input');
            input.className = 'cell-editor';
            input.type = 'text';
            input.value = prevText;
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

        function renderTeachersTableFromTextarea() {
            if (!teachersTbody) return;
            const rows = getTeachersFromTextarea();
            teachersTbody.replaceChildren();

            if (!rows.length) {
                const tr = document.createElement('tr');
                const td = document.createElement('td');
                td.colSpan = 4;
                td.style.color = '#6c757d';
                td.textContent = 'Noch keine Einträge – oben einfügen oder „+ Zeile“.';
                tr.appendChild(td);
                teachersTbody.appendChild(tr);
                return;
            }

            rows.forEach((row, idx) => {
                const tr = document.createElement('tr');

                const tdCode = document.createElement('td');
                tdCode.innerHTML = `<code>${row.code || ''}</code>`;
                tdCode.title = 'Doppelklick zum Bearbeiten';
                tdCode.addEventListener('dblclick', () => {
                    startCellEdit(tdCode, row.code, (next, meta) => {
                        const all = getTeachersFromTextarea();
                        if (!all[idx]) return renderTeachersTableFromTextarea();
                        const prev = all[idx].code;
                        all[idx].code = meta && meta.cancelled ? prev : normCode(next);
                        setTeachersTextareaFromRows(all);
                        renderTeachersTableFromTextarea();
                    });
                });

                const tdName = document.createElement('td');
                tdName.textContent = row.name || '';
                tdName.title = 'Doppelklick zum Bearbeiten';
                tdName.addEventListener('dblclick', () => {
                    startCellEdit(tdName, row.name, (next, meta) => {
                        const all = getTeachersFromTextarea();
                        if (!all[idx]) return renderTeachersTableFromTextarea();
                        const prev = all[idx].name;
                        all[idx].name = meta && meta.cancelled ? prev : normStr(next);
                        setTeachersTextareaFromRows(all);
                        renderTeachersTableFromTextarea();
                    });
                });

                const tdEmail = document.createElement('td');
                tdEmail.textContent = row.email || '';
                tdEmail.title = 'Doppelklick zum Bearbeiten';
                tdEmail.addEventListener('dblclick', () => {
                    startCellEdit(tdEmail, row.email, (next, meta) => {
                        const all = getTeachersFromTextarea();
                        if (!all[idx]) return renderTeachersTableFromTextarea();
                        const prev = all[idx].email;
                        all[idx].email = meta && meta.cancelled ? prev : normStr(next).toLowerCase();
                        setTeachersTextareaFromRows(all);
                        renderTeachersTableFromTextarea();
                    });
                });

                const tdAction = document.createElement('td');
                tdAction.className = 'action-cell';
                const btnDel = document.createElement('button');
                btnDel.type = 'button';
                btnDel.className = 'mini-btn';
                btnDel.textContent = '✕';
                btnDel.title = 'Zeile löschen';
                btnDel.addEventListener('click', () => {
                    const all = getTeachersFromTextarea();
                    all.splice(idx, 1);
                    setTeachersTextareaFromRows(all);
                    renderTeachersTableFromTextarea();
                });
                tdAction.appendChild(btnDel);

                tr.append(tdCode, tdName, tdEmail, tdAction);
                teachersTbody.appendChild(tr);
            });
        }

        function renderAdminTableFromTextarea() {
            if (!adminTbody) return;
            const rows = getAdminFromTextarea();
            adminTbody.replaceChildren();

            if (!rows.length) {
                const tr = document.createElement('tr');
                const td = document.createElement('td');
                td.colSpan = 4;
                td.style.color = '#6c757d';
                td.textContent = 'Noch keine Einträge – oben einfügen oder „+ Zeile“.';
                tr.appendChild(td);
                adminTbody.appendChild(tr);
                return;
            }

            rows.forEach((row, idx) => {
                const tr = document.createElement('tr');

                const tdRole = document.createElement('td');
                tdRole.textContent = row.role || '';
                tdRole.title = 'Doppelklick zum Bearbeiten';
                tdRole.addEventListener('dblclick', () => {
                    startCellEdit(tdRole, row.role, (next, meta) => {
                        const all = getAdminFromTextarea();
                        if (!all[idx]) return renderAdminTableFromTextarea();
                        const prev = all[idx].role;
                        all[idx].role = meta && meta.cancelled ? prev : normStr(next);
                        setAdminTextareaFromRows(all);
                        renderAdminTableFromTextarea();
                        scheduleAutoSave();
                    });
                });

                const tdName = document.createElement('td');
                tdName.textContent = row.name || '';
                tdName.title = 'Doppelklick zum Bearbeiten';
                tdName.addEventListener('dblclick', () => {
                    startCellEdit(tdName, row.name, (next, meta) => {
                        const all = getAdminFromTextarea();
                        if (!all[idx]) return renderAdminTableFromTextarea();
                        const prev = all[idx].name;
                        all[idx].name = meta && meta.cancelled ? prev : normStr(next);
                        setAdminTextareaFromRows(all);
                        renderAdminTableFromTextarea();
                        scheduleAutoSave();
                    });
                });

                const tdEmail = document.createElement('td');
                tdEmail.textContent = row.email || '';
                tdEmail.title = 'Doppelklick zum Bearbeiten';
                tdEmail.addEventListener('dblclick', () => {
                    startCellEdit(tdEmail, row.email, (next, meta) => {
                        const all = getAdminFromTextarea();
                        if (!all[idx]) return renderAdminTableFromTextarea();
                        const prev = all[idx].email;
                        all[idx].email = meta && meta.cancelled ? prev : normStr(next).toLowerCase();
                        setAdminTextareaFromRows(all);
                        renderAdminTableFromTextarea();
                        scheduleAutoSave();
                    });
                });

                const tdAction = document.createElement('td');
                tdAction.className = 'action-cell';
                const btnDel = document.createElement('button');
                btnDel.type = 'button';
                btnDel.className = 'mini-btn';
                btnDel.textContent = '✕';
                btnDel.title = 'Zeile löschen';
                btnDel.addEventListener('click', () => {
                    const all = getAdminFromTextarea();
                    all.splice(idx, 1);
                    setAdminTextareaFromRows(all);
                    renderAdminTableFromTextarea();
                    scheduleAutoSave();
                });
                tdAction.appendChild(btnDel);

                tr.append(tdRole, tdName, tdEmail, tdAction);
                adminTbody.appendChild(tr);
            });
        }

        function studentsToLines(rows) {
            return (rows || [])
                .map((x) => `${normStr(x.klasse || '')};${normStr(x.name || '')};${normStr(x.email || '').toLowerCase()}`.trim())
                .filter(Boolean)
                .join('\n');
        }

        function getStudentsFromTextarea() {
            return typeof parseLinesToStudents === 'function' ? parseLinesToStudents(taStudents ? taStudents.value : '') : [];
        }

        function setStudentsTextareaFromRows(rows) {
            if (!taStudents) return;
            taStudents.value = studentsToLines(rows);
        }

        function renderStudentsTableFromTextarea() {
            if (!studentsTbody) return;
            const rows = getStudentsFromTextarea();
            studentsTbody.replaceChildren();

            if (!rows.length) {
                const tr = document.createElement('tr');
                const td = document.createElement('td');
                td.colSpan = 4;
                td.style.color = '#6c757d';
                td.textContent = 'Noch keine Einträge – oben einfügen oder „+ Zeile“.';
                tr.appendChild(td);
                studentsTbody.appendChild(tr);
                return;
            }

            rows.forEach((row, idx) => {
                const tr = document.createElement('tr');

                const tdClass = document.createElement('td');
                tdClass.innerHTML = `<code>${row.klasse || ''}</code>`;
                tdClass.title = 'Doppelklick zum Bearbeiten';
                tdClass.addEventListener('dblclick', () => {
                    startCellEdit(tdClass, row.klasse, (next, meta) => {
                        const all = getStudentsFromTextarea();
                        if (!all[idx]) return renderStudentsTableFromTextarea();
                        const prev = all[idx].klasse;
                        all[idx].klasse = meta && meta.cancelled ? prev : normStr(next);
                        setStudentsTextareaFromRows(all);
                        renderStudentsTableFromTextarea();
                    });
                });

                const tdName = document.createElement('td');
                tdName.textContent = row.name || '';
                tdName.title = 'Doppelklick zum Bearbeiten';
                tdName.addEventListener('dblclick', () => {
                    startCellEdit(tdName, row.name, (next, meta) => {
                        const all = getStudentsFromTextarea();
                        if (!all[idx]) return renderStudentsTableFromTextarea();
                        const prev = all[idx].name;
                        all[idx].name = meta && meta.cancelled ? prev : normStr(next);
                        setStudentsTextareaFromRows(all);
                        renderStudentsTableFromTextarea();
                    });
                });

                const tdEmail = document.createElement('td');
                tdEmail.textContent = row.email || '';
                tdEmail.title = 'Doppelklick zum Bearbeiten';
                tdEmail.addEventListener('dblclick', () => {
                    startCellEdit(tdEmail, row.email, (next, meta) => {
                        const all = getStudentsFromTextarea();
                        if (!all[idx]) return renderStudentsTableFromTextarea();
                        const prev = all[idx].email;
                        all[idx].email = meta && meta.cancelled ? prev : normStr(next).toLowerCase();
                        setStudentsTextareaFromRows(all);
                        renderStudentsTableFromTextarea();
                    });
                });

                const tdAction = document.createElement('td');
                tdAction.className = 'action-cell';
                const btnDel = document.createElement('button');
                btnDel.type = 'button';
                btnDel.className = 'mini-btn';
                btnDel.textContent = '✕';
                btnDel.title = 'Zeile löschen';
                btnDel.addEventListener('click', () => {
                    const all = getStudentsFromTextarea();
                    all.splice(idx, 1);
                    setStudentsTextareaFromRows(all);
                    renderStudentsTableFromTextarea();
                });
                tdAction.appendChild(btnDel);

                tr.append(tdClass, tdName, tdEmail, tdAction);
                studentsTbody.appendChild(tr);
            });
        }

        function classesToLines(rows) {
            return (rows || [])
                .map((x) => {
                    const y = normStr(x.year || '');
                    const year = /^\d{4}$/.test(y) ? y : '';
                    return `${normCode(x.code)};${year};${normStr(x.name || '')};${normStr(x.headName || '')};${normStr(x.headEmail || '').toLowerCase()}`.trim();
                })
                .filter(Boolean)
                .join('\n');
        }

        function getClassesFromTextarea() {
            return typeof parseLinesToClasses === 'function' ? parseLinesToClasses(taClasses ? taClasses.value : '') : [];
        }

        function setClassesTextareaFromRows(rows) {
            if (!taClasses) return;
            taClasses.value = classesToLines(rows);
        }

        function renderClassesTableFromTextarea() {
            if (!classesTbody) return;
            const rows = getClassesFromTextarea();
            classesTbody.replaceChildren();

            if (!rows.length) {
                const tr = document.createElement('tr');
                const td = document.createElement('td');
                td.colSpan = 6;
                td.style.color = '#6c757d';
                td.textContent = 'Noch keine Einträge – oben einfügen oder „+ Zeile“.';
                tr.appendChild(td);
                classesTbody.appendChild(tr);
                return;
            }

            rows.forEach((row, idx) => {
                const tr = document.createElement('tr');

                const tdCode = document.createElement('td');
                tdCode.innerHTML = `<code>${row.code || ''}</code>`;
                tdCode.title = 'Doppelklick zum Bearbeiten';
                tdCode.addEventListener('dblclick', () => {
                    startCellEdit(tdCode, row.code, (next, meta) => {
                        const all = getClassesFromTextarea();
                        if (!all[idx]) return renderClassesTableFromTextarea();
                        const prev = all[idx].code;
                        all[idx].code = meta && meta.cancelled ? prev : normCode(next);
                        setClassesTextareaFromRows(all);
                        renderClassesTableFromTextarea();
                    });
                });

                const tdYear = document.createElement('td');
                tdYear.textContent = row.year || '';
                tdYear.title = 'Doppelklick zum Bearbeiten';
                tdYear.addEventListener('dblclick', () => {
                    startCellEdit(tdYear, row.year, (next, meta) => {
                        const all = getClassesFromTextarea();
                        if (!all[idx]) return renderClassesTableFromTextarea();
                        const prev = all[idx].year || '';
                        const n = normStr(next);
                        all[idx].year = meta && meta.cancelled ? prev : /^\d{4}$/.test(n) ? n : '';
                        setClassesTextareaFromRows(all);
                        renderClassesTableFromTextarea();
                    });
                });

                const tdName = document.createElement('td');
                tdName.textContent = row.name || '';
                tdName.title = 'Doppelklick zum Bearbeiten';
                tdName.addEventListener('dblclick', () => {
                    startCellEdit(tdName, row.name, (next, meta) => {
                        const all = getClassesFromTextarea();
                        if (!all[idx]) return renderClassesTableFromTextarea();
                        const prev = all[idx].name;
                        all[idx].name = meta && meta.cancelled ? prev : normStr(next);
                        setClassesTextareaFromRows(all);
                        renderClassesTableFromTextarea();
                    });
                });

                const tdHead = document.createElement('td');
                tdHead.textContent = row.headName || '';
                tdHead.title = 'Doppelklick zum Bearbeiten';
                tdHead.addEventListener('dblclick', () => {
                    startCellEdit(tdHead, row.headName, (next, meta) => {
                        const all = getClassesFromTextarea();
                        if (!all[idx]) return renderClassesTableFromTextarea();
                        const prev = all[idx].headName;
                        all[idx].headName = meta && meta.cancelled ? prev : normStr(next);
                        setClassesTextareaFromRows(all);
                        renderClassesTableFromTextarea();
                    });
                });

                const tdEmail = document.createElement('td');
                tdEmail.textContent = row.headEmail || '';
                tdEmail.title = 'Doppelklick zum Bearbeiten';
                tdEmail.addEventListener('dblclick', () => {
                    startCellEdit(tdEmail, row.headEmail, (next, meta) => {
                        const all = getClassesFromTextarea();
                        if (!all[idx]) return renderClassesTableFromTextarea();
                        const prev = all[idx].headEmail;
                        all[idx].headEmail = meta && meta.cancelled ? prev : normStr(next).toLowerCase();
                        setClassesTextareaFromRows(all);
                        renderClassesTableFromTextarea();
                    });
                });

                const tdAction = document.createElement('td');
                tdAction.className = 'action-cell';
                const btnDel = document.createElement('button');
                btnDel.type = 'button';
                btnDel.className = 'mini-btn';
                btnDel.textContent = '✕';
                btnDel.title = 'Zeile löschen';
                btnDel.addEventListener('click', () => {
                    const all = getClassesFromTextarea();
                    all.splice(idx, 1);
                    setClassesTextareaFromRows(all);
                    renderClassesTableFromTextarea();
                });
                tdAction.appendChild(btnDel);

                tr.append(tdCode, tdYear, tdName, tdHead, tdEmail, tdAction);
                classesTbody.appendChild(tr);
            });
        }

        function renderFromStorage() {
            const s = load();
            // Domain in UI-Feld zurückschreiben (wird auch von school-domain.js genutzt)
            try {
                if (domainInput) domainInput.value = normStr(s.domain || '');
                if (typeof window.ms365SetSchoolDomainNoAt === 'function') {
                    const d = normStr(s.domain || '').replace(/^@+/, '');
                    if (d) window.ms365SetSchoolDomainNoAt(d);
                }
            } catch {
                // ignore
            }
            if (taSubjects) {
                taSubjects.value = (s.subjects || []).map((x) => `${x.code};${x.name || ''}`.trim()).join('\n');
            }
            if (taArges) {
                taArges.value = (s.arges || [])
                    .map((x) => `${x.code};${x.name || ''};${(x.subjects || []).join(',')}`.trim())
                    .join('\n');
            }
            if (taTeachers) {
                taTeachers.value = (s.teachers || [])
                    .map((x) => `${x.code};${x.name || ''};${x.email || ''}`.trim())
                    .join('\n');
            }
            if (taAdmin) {
                taAdmin.value = (s.admin || [])
                    .map((x) => `${x.role || ''};${x.name || ''};${x.email || ''}`.trim())
                    .join('\n');
            }
            if (taStudents) {
                taStudents.value = (s.students || [])
                    .map((x) => `${x.klasse || ''};${x.name || ''};${x.email || ''}`.trim())
                    .join('\n');
            }
            if (taClasses) {
                taClasses.value = (s.classes || [])
                    .map((x) => `${x.code || ''};${x.year || ''};${x.name || ''};${x.headName || ''};${x.headEmail || ''}`.trim())
                    .join('\n');
            }
            renderSubjectsTableFromTextarea();
            renderArgesTableFromTextarea();
            renderTeachersTableFromTextarea();
            renderAdminTableFromTextarea();
            renderStudentsTableFromTextarea();
            renderClassesTableFromTextarea();
            renderSchoolYearSelectFromV2();
            const yLbl = getDisplayedSchoolYearLabel() || currentSchoolYearLabel();
            setSummary(
                `Aktueller Stand: schulweit ${(s.subjects || []).length} Fächer, ${(s.arges || []).length} ARGEs, ${(s.teachers || []).length} Lehrkräfte, ${(s.admin || []).length} Verwaltung — für Schuljahr ${yLbl}: ${(s.students || []).length} Schüler, ${(s.classes || []).length} Klassen.`,
                'ok'
            );
            dispatchTenantSettingsChanged(s, 'render');
        }

        function renderSchoolYearSelectFromV2() {
            if (!schoolYearSelect) return;
            try {
                if (!window.ms365AppDataV2 || typeof window.ms365AppDataV2.getContainer !== 'function') {
                    schoolYearSelect.replaceChildren();
                    const o = document.createElement('option');
                    o.value = currentSchoolYearLabel();
                    o.textContent = currentSchoolYearLabel();
                    schoolYearSelect.appendChild(o);
                    schoolYearSelect.value = o.value;
                    return;
                }
                const c = window.ms365AppDataV2.getContainer();
                const cur = c && c.years ? String(c.years.current || '') : '';
                const years = typeof window.ms365AppDataV2.listYears === 'function' ? window.ms365AppDataV2.listYears() : [];
                const list = years.length ? years : (cur ? [cur] : [currentSchoolYearLabel()]);
                schoolYearSelect.replaceChildren();
                list.forEach((y) => {
                    const opt = document.createElement('option');
                    opt.value = String(y);
                    opt.textContent = String(y);
                    schoolYearSelect.appendChild(opt);
                });
                schoolYearSelect.value = cur && list.includes(cur) ? cur : list[0];
            } catch {
                // ignore
            }
        }

        function setCurrentSchoolYearInV2(nextLabel, opts) {
            try {
                if (!window.ms365AppDataV2 || typeof window.ms365AppDataV2.setCurrentYear !== 'function') return false;
                window.ms365AppDataV2.setCurrentYear(String(nextLabel || '').trim(), opts || {});
                return true;
            } catch {
                return false;
            }
        }

        function downloadJson(filename, obj) {
            const blob = new Blob([JSON.stringify(obj, null, 2)], { type: 'application/json;charset=utf-8' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            a.remove();
            setTimeout(() => URL.revokeObjectURL(url), 250);
        }

        function importFileToRows(file, onRows) {
            importSpreadsheetFileToJsonRows(file, onRows, (msg) => setSummary(msg, 'warn'));
        }

        function importSubjectsRows(jsonRows) {
            const out = [];
            (jsonRows || []).forEach((r) => {
                const code = getField(r, ['kürzel', 'kuerzel', 'code', 'fach', 'abk', 'abkuerzung', 'abbreviation']);
                const name = getField(r, ['name', 'fachname', 'bezeichnung', 'subject', 'subjectname']);
                const c = normCode(code);
                if (!c) return;
                out.push({ code: c, name: normStr(name) });
            });
            if (taSubjects) taSubjects.value = out.map((x) => `${x.code};${x.name || ''}`.trim()).join('\n');
            setSummary(`Fächer importiert: ${out.length} (schulweit)`, 'ok');
        }

        function importArgesRows(jsonRows) {
            const out = [];
            (jsonRows || []).forEach((r) => {
                const code = getField(r, ['kürzel', 'kuerzel', 'code', 'arge', 'abk', 'abkuerzung']);
                const name = getField(r, ['name', 'bezeichnung', 'titel', 'title']);
                const subs = getField(r, ['faecher', 'fächer', 'subjects', 'fach', 'subjectcodes']);
                const c = normCode(code);
                if (!c) return;
                const subjects = String(subs || '')
                    .split(/[,\s|]+/)
                    .map((x) => normCode(x))
                    .filter(Boolean);
                out.push({ code: c, name: normStr(name), subjects });
            });
            if (taArges) taArges.value = argesToLines(out);
            renderArgesTableFromTextarea();
            setSummary(`ARGEs importiert: ${out.length} (schulweit)`, 'ok');
        }

        function importTeachersRows(jsonRows) {
            const out = [];
            (jsonRows || []).forEach((r) => {
                const code = getField(r, ['kürzel', 'kuerzel', 'code', 'lehrer', 'abbrev', 'abbreviation']);
                let name = getField(r, ['name', 'lehrername', 'anzeigename', 'displayname']);
                let email = getField(r, ['e-mail', 'email', 'mail', 'upn']);
                const c = normCode(code);
                if (!c) return;

                // Heuristik für Teillisten: wenn "Name" eigentlich eine E-Mail ist (enthält @),
                // dann korrekt zuordnen statt E-Mail als Name zu speichern.
                const nameNorm = normStr(name);
                const emailNorm = normStr(email).toLowerCase();
                const nameLooksLikeEmail = nameNorm.includes('@');
                const emailLooksLikeEmail = emailNorm.includes('@');

                if (nameLooksLikeEmail && (!emailNorm || !emailLooksLikeEmail)) {
                    email = nameNorm;
                    name = '';
                }

                out.push({ code: c, name: normStr(name), email: normStr(email).toLowerCase() });
            });
            if (taTeachers) taTeachers.value = teachersToLines(out);
            renderTeachersTableFromTextarea();
            setSummary(`Lehrkräfte importiert: ${out.length} (schulweit)`, 'ok');
        }

        function importStudentsRows(jsonRows) {
            const out = [];
            (jsonRows || []).forEach((r) => {
                let klasse = getField(r, ['klasse', 'class', 'gruppe', 'group']);
                let name = getField(r, ['name', 'schueler', 'schüler', 'anzeigename', 'displayname']);
                let email = getField(r, ['e-mail', 'email', 'mail', 'upn']);
                if (!klasse && !name && !email) return;

                // Heuristik für Teillisten: wenn "Name" eigentlich eine E-Mail ist (enthält @),
                // dann korrekt zuordnen statt E-Mail als Name zu speichern.
                const nameNorm = normStr(name);
                const emailNorm = normStr(email).toLowerCase();
                const nameLooksLikeEmail = nameNorm.includes('@');
                const emailLooksLikeEmail = emailNorm.includes('@');

                if (nameLooksLikeEmail && (!emailNorm || !emailLooksLikeEmail)) {
                    email = nameNorm;
                    name = '';
                }

                out.push({ klasse: normStr(klasse), name: normStr(name), email: normStr(email).toLowerCase() });
            });
            if (taStudents) taStudents.value = studentsToLines(out);
            renderStudentsTableFromTextarea();
            const ySt = getDisplayedSchoolYearLabel() || currentSchoolYearLabel();
            setSummary(`Schüler importiert: ${out.length} (Schuljahr ${ySt})`, 'ok');
        }

        function importClassesRows(jsonRows) {
            const out = [];
            (jsonRows || []).forEach((r) => {
                let code = getField(r, ['abkürzung', 'abkuerzung', 'abk', 'kuerzel', 'kürzel', 'code', 'klasseabk', 'classcode']);
                let year = getField(r, ['abschlussjahr', 'abschluss', 'year', 'graduationyear']);
                let name = getField(r, ['klasse', 'class', 'name', 'bezeichnung', 'classname']);
                let headName = getField(r, ['klassenvorstand', 'klassenvorstandname', 'kv', 'kvname', 'vorstand', 'head', 'headname']);
                let headEmail = getField(r, ['klassenvorstandemail', 'kvemail', 'e-mail', 'email', 'mail', 'upn', 'heademail']);
                if (!code && !year && !name && !headName && !headEmail) return;

                // Heuristik: falls "Klassenvorstand" eigentlich E-Mail ist
                const hn = normStr(headName);
                const he = normStr(headEmail).toLowerCase();
                if (hn.includes('@') && (!he || !he.includes('@'))) {
                    headEmail = hn;
                    headName = '';
                }

                out.push({
                    code: normCode(code),
                    year: /^\d{4}$/.test(normStr(year)) ? normStr(year) : '',
                    name: normStr(name),
                    headName: normStr(headName),
                    headEmail: normStr(headEmail).toLowerCase()
                });
            });
            if (taClasses) taClasses.value = classesToLines(out);
            renderClassesTableFromTextarea();
            const yCl = getDisplayedSchoolYearLabel() || currentSchoolYearLabel();
            setSummary(`Klassen importiert: ${out.length} (Schuljahr ${yCl})`, 'ok');
        }

        if (btnSave) {
            btnSave.addEventListener('click', () => {
                const subjects = typeof parseLinesToSubjects === 'function' ? parseLinesToSubjects(taSubjects ? taSubjects.value : '') : [];
                const teachers = typeof parseLinesToTeachers === 'function' ? parseLinesToTeachers(taTeachers ? taTeachers.value : '') : [];
                const admin = typeof parseLinesToAdmin === 'function' ? parseLinesToAdmin(taAdmin ? taAdmin.value : '') : [];
                const students = typeof parseLinesToStudents === 'function' ? parseLinesToStudents(taStudents ? taStudents.value : '') : [];
                const classes = typeof parseLinesToClasses === 'function' ? parseLinesToClasses(taClasses ? taClasses.value : '') : [];
                const domain =
                    typeof window.ms365GetSchoolDomainNoAt === 'function' ? window.ms365GetSchoolDomainNoAt() : '';
                const arges = typeof parseLinesToArges === 'function' ? parseLinesToArges(taArges ? taArges.value : '') : [];
                const saved = save({ domain, subjects, arges, teachers, admin, students, classes });
                const ySave = getDisplayedSchoolYearLabel() || currentSchoolYearLabel();
                setSummary(
                    `Gespeichert: schulweit ${(saved.subjects || []).length} Fächer, ${(saved.arges || []).length} ARGEs, ${(saved.teachers || []).length} Lehrkräfte, ${(saved.admin || []).length} Verwaltung — für Schuljahr ${ySave}: ${(saved.students || []).length} Schüler, ${(saved.classes || []).length} Klassen.`,
                    'ok'
                );
                renderSubjectsTableFromTextarea();
                renderArgesTableFromTextarea();
                renderTeachersTableFromTextarea();
                renderAdminTableFromTextarea();
                renderStudentsTableFromTextarea();
                renderClassesTableFromTextarea();
                dispatchTenantSettingsChanged(saved, 'manual-save');
            });
        }

        // Struktur -> Listen (Writeback)
        function mergeIntoSubjects(existing, codesToEnsure) {
            const out = (existing || []).slice();
            const seen = new Set(out.map((s) => String(s.code || '').toUpperCase()).filter(Boolean));
            (codesToEnsure || []).forEach((c) => {
                const code = normCode(c);
                if (!code || seen.has(code)) return;
                seen.add(code);
                out.push({ code, name: '' });
            });
            out.sort((a, b) => normCode(a.code).localeCompare(normCode(b.code)));
            return out;
        }

        function mergeIntoClasses(existing, codesToEnsure) {
            const out = (existing || []).slice();
            const seen = new Set(out.map((c) => normCode(c.code || c.name || '')).filter(Boolean));
            (codesToEnsure || []).forEach((c) => {
                const code = normCode(c);
                if (!code || seen.has(code)) return;
                seen.add(code);
                out.push({ code, name: `Klasse ${code}`, year: '', headName: '', headEmail: '' });
            });
            out.sort((a, b) => normCode(a.code).localeCompare(normCode(b.code)));
            return out;
        }

        function mergeIntoArges(existing, codesToEnsure) {
            const out = (existing || []).slice();
            const seen = new Set(out.map((a) => normCode(a.code || '')).filter(Boolean));
            (codesToEnsure || []).forEach((c) => {
                const code = normCode(c);
                if (!code || seen.has(code)) return;
                seen.add(code);
                out.push({ code, name: code, subjects: [] });
            });
            out.sort((a, b) => normCode(a.code).localeCompare(normCode(b.code)));
            return out;
        }

        function applyWriteback(detail) {
            if (!detail || !detail.writeback) return;
            if (typeof load !== 'function' || typeof save !== 'function') return;
            const current = load();
            const next = Object.assign({}, current);
            if (detail.writeback.subjectCodes) {
                next.subjects = mergeIntoSubjects(current.subjects || [], detail.writeback.subjectCodes);
            }
            if (detail.writeback.argeCodes) {
                next.arges = mergeIntoArges(current.arges || [], detail.writeback.argeCodes);
            }
            if (detail.writeback.classCodes) {
                next.classes = mergeIntoClasses(current.classes || [], detail.writeback.classCodes);
            }
            __syncGuard++;
            const saved = save(next);
            // UI aktualisieren
            try {
                if (taSubjects) taSubjects.value = subjectsToLines(saved.subjects || []);
                if (taArges) taArges.value = argesToLines(saved.arges || []);
                if (taClasses) taClasses.value = classesToLines(saved.classes || []);
                renderSubjectsTableFromTextarea();
                renderArgesTableFromTextarea();
                renderClassesTableFromTextarea();
                setSummary('Listen wurden aus der Struktur ergänzt.', 'ok');
            } catch {
                // ignore
            }
            __syncGuard--;
            dispatchTenantSettingsChanged(saved, 'writeback');
        }

        try {
            window.addEventListener('ms365-structure-changed', (ev) => {
                if (__syncGuard) return;
                applyWriteback(ev && ev.detail ? ev.detail : null);
                try {
                    if (typeof window.__ms365TenantStepsRefreshMatch === 'function') window.__ms365TenantStepsRefreshMatch();
                } catch {
                    // ignore
                }
            });
            window.addEventListener('ms365-match-links-changed', () => {
                try {
                    if (typeof window.__ms365TenantStepsRefreshMatch === 'function') window.__ms365TenantStepsRefreshMatch();
                } catch {
                    // ignore
                }
            });
        } catch {
            // ignore
        }

        if (fileSubjects) {
            fileSubjects.addEventListener('change', (e) => {
                const f = e.target.files && e.target.files[0];
                importFileToRows(f, (rows) => importSubjectsRows(rows));
                fileSubjects.value = '';
            });
        }
        if (fileArges) {
            fileArges.addEventListener('change', (e) => {
                const f = e.target.files && e.target.files[0];
                importFileToRows(f, (rows) => importArgesRows(rows));
                fileArges.value = '';
            });
        }
        if (fileTeachers) {
            fileTeachers.addEventListener('change', (e) => {
                const f = e.target.files && e.target.files[0];
                importFileToRows(f, (rows) => importTeachersRows(rows));
                fileTeachers.value = '';
            });
        }
        if (fileStudents) {
            fileStudents.addEventListener('change', (e) => {
                const f = e.target.files && e.target.files[0];
                importFileToRows(f, (rows) => importStudentsRows(rows));
                fileStudents.value = '';
            });
        }
        if (fileClasses) {
            fileClasses.addEventListener('change', (e) => {
                const f = e.target.files && e.target.files[0];
                importFileToRows(f, (rows) => importClassesRows(rows));
                fileClasses.value = '';
            });
        }

        if (btnSubjectsTpl) {
            btnSubjectsTpl.addEventListener('click', () => {
                const ok = downloadXlsxTemplate(
                    'Faecherliste-Vorlage.xlsx',
                    [
                        ['Kürzel', 'Name'],
                        ['D', 'Deutsch'],
                        ['M', 'Mathematik'],
                        ['E', 'Englisch']
                    ],
                    'Faecher'
                );
                if (!ok) setSummary('Vorlage: Excel-Bibliothek nicht geladen – Seite neu laden.', 'warn');
            });
        }
        if (btnArgesTpl) {
            btnArgesTpl.addEventListener('click', () => {
                const ok = downloadXlsxTemplate(
                    'ARGE-Liste-Vorlage.xlsx',
                    [
                        ['Kürzel', 'Name', 'Fächer'],
                        ['SPRACHEN', 'Sprachen', 'D,E,FS2'],
                        ['NAWI', 'Naturwissenschaften', 'BIO,CH,PH']
                    ],
                    'ARGEs'
                );
                if (!ok) setSummary('Vorlage: Excel-Bibliothek nicht geladen – Seite neu laden.', 'warn');
            });
        }
        if (btnTeachersTpl) {
            btnTeachersTpl.addEventListener('click', () => {
                const ok = downloadXlsxTemplate(
                    'Lehrerliste-Vorlage.xlsx',
                    [
                        ['Kürzel', 'Name', 'E-Mail'],
                        ['MU', 'Max Mustermann', 'max.mustermann@schule.de'],
                        ['BME', 'Anna Beispiel', 'anna.beispiel@schule.de']
                    ],
                    'Lehrer'
                );
                if (!ok) setSummary('Vorlage: Excel-Bibliothek nicht geladen – Seite neu laden.', 'warn');
            });
        }
        if (btnStudentsTpl) {
            btnStudentsTpl.addEventListener('click', () => {
                const ok = downloadXlsxTemplate(
                    'Schuelerliste-Vorlage.xlsx',
                    [
                        ['Klasse', 'Name', 'E-Mail'],
                        ['1AK', 'Max Mustermann', 'max.mustermann@schule.de'],
                        ['1AK', 'Anna Beispiel', 'anna.beispiel@schule.de']
                    ],
                    'Schueler'
                );
                if (!ok) setSummary('Vorlage: Excel-Bibliothek nicht geladen – Seite neu laden.', 'warn');
            });
        }
        if (btnClassesTpl) {
            btnClassesTpl.addEventListener('click', () => {
                const ok = downloadXlsxTemplate(
                    'Klassenliste-Vorlage.xlsx',
                    [
                        ['Abkürzung', 'Abschlussjahr', 'Klasse', 'Klassenvorstand', 'E-Mail'],
                        ['1AK', '2030', '1A-Klasse', 'Max Mustermann', 'max.mustermann@schule.de'],
                        ['2BK', '2029', '2B-Klasse', 'Anna Beispiel', 'anna.beispiel@schule.de']
                    ],
                    'Klassen'
                );
                if (!ok) setSummary('Vorlage: Excel-Bibliothek nicht geladen – Seite neu laden.', 'warn');
            });
        }

        if (btnReload) {
            btnReload.addEventListener('click', () => renderFromStorage());
        }

        if (schoolYearSelect && !schoolYearSelect.dataset.bound) {
            schoolYearSelect.dataset.bound = '1';
            schoolYearSelect.addEventListener('change', () => {
                const y = String(schoolYearSelect.value || '').trim();
                if (!y) return;
                setCurrentSchoolYearInV2(y);
                renderFromStorage();
                setSummary('Schuljahr gewechselt: ' + y + ' — Schüler- und Klassenlisten beziehen sich nun auf dieses Jahr.', 'ok');
            });
        }
        if (schoolYearAddBtn && !schoolYearAddBtn.dataset.bound) {
            schoolYearAddBtn.dataset.bound = '1';
            schoolYearAddBtn.addEventListener('click', () => {
                void (async () => {
                    const cur = schoolYearSelect ? String(schoolYearSelect.value || '').trim() : '';
                    const suggest = (function () {
                        const m = cur.match(/^(\d{4})\s*\/\s*(\d{2}|\d{4})/);
                        if (!m) return '';
                        const y = parseInt(m[1], 10);
                        if (!isFinite(y)) return '';
                        return String(y + 1) + '/' + String(y + 2).slice(2);
                    })();
                    const next = await dlgPrompt('Neues Schuljahr (z. B. 2027/28)', suggest || currentSchoolYearLabel(), {
                        title: 'Schuljahr',
                        inputLabel: 'Bezeichnung'
                    });
                    if (next == null || !normStr(next)) return;
                    const copy = await dlgConfirm('Schüler & Klassen aus dem aktuellen Schuljahr übernehmen?', {
                        title: 'Schuljahr',
                        okText: 'Ja, übernehmen',
                        cancelText: 'Nein'
                    });
                    setCurrentSchoolYearInV2(next, copy && cur ? { copyFrom: cur } : {});
                    renderFromStorage();
                    if (schoolYearSelect) schoolYearSelect.value = String(next).trim();
                    setSummary('Neues Schuljahr angelegt: ' + String(next).trim(), 'ok');
                })();
            });
        }

        if (btnExport) {
            btnExport.addEventListener('click', () => {
                try {
                    if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.exportJson === 'function') {
                        downloadJson('ms365-schooltool-data-v2.json', window.ms365AppDataV2.exportJson());
                        return;
                    }
                } catch {
                    // ignore
                }
                downloadJson('schule-einstellungen.json', load());
            });
        }
        if (btnExportHeader) {
            btnExportHeader.addEventListener('click', () => {
                try {
                    if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.exportJson === 'function') {
                        downloadJson('ms365-schooltool-data-v2.json', window.ms365AppDataV2.exportJson());
                        return;
                    }
                } catch {
                    // ignore
                }
                downloadJson('schule-einstellungen.json', load());
            });
        }

        if (btnClear) {
            btnClear.addEventListener('click', () => {
                try {
                    localStorage.removeItem('ms365-tenant-settings-v1');
                } catch {
                    // ignore
                }
                // UI/Domain wieder auf Standard zurücksetzen
                try {
                    const domainInput = document.getElementById('schoolEmailDomain');
                    if (domainInput) domainInput.value = 'ms365.schule';
                    if (typeof window.ms365SetSchoolDomainNoAt === 'function') {
                        window.ms365SetSchoolDomainNoAt('ms365.schule');
                    }
                } catch {
                    // ignore
                }
                if (taSubjects) taSubjects.value = '';
                if (taArges) taArges.value = '';
                if (taTeachers) taTeachers.value = '';
                if (taAdmin) taAdmin.value = '';
                if (taStudents) taStudents.value = '';
                if (taClasses) taClasses.value = '';
                renderSubjectsTableFromTextarea();
                renderArgesTableFromTextarea();
                renderTeachersTableFromTextarea();
                renderAdminTableFromTextarea();
                renderStudentsTableFromTextarea();
                renderClassesTableFromTextarea();
                setSummary('Schul‑Grundeinstellungen gelöscht (nur lokaler Browser-Speicher).', 'warn');
            });
        }

        if (fileImport) {
            fileImport.addEventListener('change', async (e) => {
                const f = e.target.files && e.target.files[0];
                if (!f) return;
                try {
                    const text = await f.text();
                    const obj = safeJsonParse(text);
                    if (!obj) {
                        setSummary('Import fehlgeschlagen: keine gültige JSON-Datei.', 'warn');
                        return;
                    }
                    try {
                        if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.importJson === 'function') {
                            window.ms365AppDataV2.importJson(obj);
                        }
                    } catch (e) {
                        setSummary('Import fehlgeschlagen: ' + (e?.message || String(e)), 'warn');
                        return;
                    }
                    const saved = save(obj);
                    if (taSubjects) taSubjects.value = (saved.subjects || []).map((x) => `${x.code};${x.name || ''}`.trim()).join('\n');
                    if (taArges) taArges.value = (saved.arges || []).map((x) => `${x.code};${x.name || ''};${(x.subjects || []).join(',')}`.trim()).join('\n');
                    if (taTeachers) taTeachers.value = (saved.teachers || []).map((x) => `${x.code};${x.name || ''};${x.email || ''}`.trim()).join('\n');
                    if (taAdmin) taAdmin.value = (saved.admin || []).map((x) => `${x.role || ''};${x.name || ''};${x.email || ''}`.trim()).join('\n');
                    if (taStudents) taStudents.value = (saved.students || []).map((x) => `${x.klasse || ''};${x.name || ''};${x.email || ''}`.trim()).join('\n');
                    if (taClasses) taClasses.value = (saved.classes || []).map((x) => `${x.code || ''};${x.year || ''};${x.name || ''};${x.headName || ''};${x.headEmail || ''}`.trim()).join('\n');
                    renderSubjectsTableFromTextarea();
                    renderArgesTableFromTextarea();
                    renderTeachersTableFromTextarea();
                    renderAdminTableFromTextarea();
                    renderStudentsTableFromTextarea();
                    renderClassesTableFromTextarea();
                    const yImp = getDisplayedSchoolYearLabel() || currentSchoolYearLabel();
                    setSummary(
                        `Import OK: schulweit ${(saved.subjects || []).length} Fächer, ${(saved.arges || []).length} ARGEs, ${(saved.teachers || []).length} Lehrkräfte, ${(saved.admin || []).length} Verwaltung — für Schuljahr ${yImp}: ${(saved.students || []).length} Schüler, ${(saved.classes || []).length} Klassen.`,
                        'ok'
                    );
                } catch (err) {
                    setSummary('Import fehlgeschlagen: ' + (err?.message || String(err)), 'warn');
                } finally {
                    fileImport.value = '';
                }
            });
        }

        if (domainInput) {
            domainInput.addEventListener('input', () => scheduleAutoSave());
            domainInput.addEventListener('change', () => scheduleAutoSave());
        }
        // Kein Standard-Abschlussjahr mehr in den Schul‑Grundeinstellungen
        if (taSubjects) taSubjects.addEventListener('input', () => scheduleAutoSave());
        if (taSubjects) taSubjects.addEventListener('input', () => renderSubjectsTableFromTextarea());

        if (taArges) {
            taArges.addEventListener('input', () => renderArgesTableFromTextarea());
            taArges.addEventListener('input', () => scheduleAutoSave());
        }

        if (btnAddSubjectRow) {
            btnAddSubjectRow.addEventListener('click', () => {
                const all = getSubjectsFromTextarea();
                all.push({ code: '', name: '' });
                setSubjectsTextareaFromRows(all);
                renderSubjectsTableFromTextarea();
                scheduleAutoSave();
            });
        }

        if (btnAddArgeRow) {
            btnAddArgeRow.addEventListener('click', () => {
                const all = getArgesFromTextarea();
                all.push({ code: '', name: '', subjects: [] });
                setArgesTextareaFromRows(all);
                renderArgesTableFromTextarea();
                scheduleAutoSave();
            });
        }

        if (taTeachers) {
            taTeachers.addEventListener('input', () => renderTeachersTableFromTextarea());
            taTeachers.addEventListener('input', () => scheduleAutoSave());
        }
        if (taAdmin) {
            taAdmin.addEventListener('input', () => renderAdminTableFromTextarea());
            taAdmin.addEventListener('input', () => scheduleAutoSave());
        }
        if (btnAddTeacherRow) {
            btnAddTeacherRow.addEventListener('click', () => {
                const all = getTeachersFromTextarea();
                all.push({ code: '', name: '', email: '' });
                setTeachersTextareaFromRows(all);
                renderTeachersTableFromTextarea();
                scheduleAutoSave();
            });
        }
        if (btnAddAdminRow) {
            btnAddAdminRow.addEventListener('click', () => {
                const all = getAdminFromTextarea();
                all.push({ role: '', name: '', email: '' });
                setAdminTextareaFromRows(all);
                renderAdminTableFromTextarea();
                scheduleAutoSave();
            });
        }

        if (taStudents) {
            taStudents.addEventListener('input', () => renderStudentsTableFromTextarea());
            taStudents.addEventListener('input', () => scheduleAutoSave());
        }
        if (btnAddStudentRow) {
            btnAddStudentRow.addEventListener('click', () => {
                const all = getStudentsFromTextarea();
                all.push({ klasse: '', name: '', email: '' });
                setStudentsTextareaFromRows(all);
                renderStudentsTableFromTextarea();
                scheduleAutoSave();
            });
        }

        if (taClasses) {
            taClasses.addEventListener('input', () => renderClassesTableFromTextarea());
            taClasses.addEventListener('input', () => scheduleAutoSave());
        }
        if (btnAddClassRow) {
            btnAddClassRow.addEventListener('click', () => {
                const all = getClassesFromTextarea();
                all.push({ code: '', name: '', headName: '', headEmail: '' });
                setClassesTextareaFromRows(all);
                renderClassesTableFromTextarea();
                scheduleAutoSave();
            });
        }

        const seeded = seedDemoDataIfEmptyStorage();
        renderFromStorage();
        if (seeded) {
            setSummary(
                'Demo-Daten geladen: Domain, Fächer und Lehrkräfte (schulweit); Schüler und Klassen für das aktuelle Schuljahr. Du kannst alles anpassen oder löschen.',
                'ok'
            );
        }

        // Schritt 4 / 5: Tenant-IST, Match (Dropdown+Speichern), Differenz + Graph-Anlage
        function setTenantDeltaProgress(visible, text, ratio) {
            const wrap = document.getElementById('tenantDeltaProgressWrap');
            const txt = document.getElementById('tenantDeltaProgressText');
            const bar = document.getElementById('tenantDeltaProgressBar');
            const pct = document.getElementById('tenantDeltaProgressPct');
            if (wrap) wrap.style.display = visible ? '' : 'none';
            if (txt && text) txt.textContent = String(text);
            const r = typeof ratio === 'number' && isFinite(ratio) ? Math.max(0, Math.min(1, ratio)) : null;
            if (bar) bar.style.width = r === null ? '0%' : String(Math.round(r * 100)) + '%';
            if (pct) pct.textContent = r === null ? '–' : String(Math.round(r * 100)) + ' %';
        }

        function collectTenantDeltaItems(api) {
            const out = [];
            if (!api || typeof api.loadStructureState !== 'function') return out;
            const { rows } = api.loadStructureState();
            const links = api.loadMatchLinks() || {};
            (rows || []).forEach((r) => {
                if (!r) return;
                const t = String(r.typ || '');
                const gid =
                    normStr(r.tenantGroupId) ||
                    (links[String(r.id)] && normStr(links[String(r.id)].tenantGroupId));
                const groupLike =
                    t === 'Gruppe' ||
                    t === 'Arbeitsgemeinschaft' ||
                    t === 'Klasse' ||
                    t === 'Jahrgang' ||
                    t === 'SchuelerInnen' ||
                    t === 'LehrerInnen' ||
                    t === 'Verwaltung';
                if (groupLike) {
                    if (!gid) {
                        const sug = typeof api.computeCreateSuggestion === 'function' ? api.computeCreateSuggestion(r) : null;
                        const can = !!(sug && sug.displayName && sug.mailNick);
                        let action = can
                            ? 'M365‑Gruppe/Team per Graph anlegen (Mail‑Nickname vorhanden).'
                            : 'Schema/Bezeichnung prüfen (Mail‑Nickname leer).';
                        if (t === 'Klasse' && can) {
                            action = 'Klassen‑Team/Gruppe per Graph anlegen (Schema: meist Microsoft‑365‑Gruppe).';
                        } else if (t === 'Jahrgang' && can) {
                            action = 'Jahrgangs‑Gruppe per Graph anlegen (Schema: Jahrgang/Jg‑Suffix oder Anzeigename).';
                        } else if (t === 'Arbeitsgemeinschaft' && can) {
                            action = 'ARGE als Microsoft‑365‑Gruppe per Graph anlegen.';
                        }
                        out.push({
                            kind: 'group',
                            r,
                            action,
                            canProvision: can
                        });
                    }
                    return;
                }
                if (t === 'Kursteam') {
                    if (gid) return;
                    const kt =
                        window.ms365StructureRules &&
                        typeof window.ms365StructureRules.resolveKursteamKlasseFach === 'function'
                            ? window.ms365StructureRules.resolveKursteamKlasseFach(r, rows)
                            : {
                                  klasse: normStr(r.ktKlasse),
                                  fach: normStr(r.ktFach),
                                  hasBoth: !!(normStr(r.ktKlasse) && normStr(r.ktFach))
                              };
                    if (!kt.hasBoth) {
                        out.push({
                            kind: 'group',
                            r,
                            action:
                                'Kursteam: Klasse und Fach fehlen (Feld „Klasse“/„Fach“ oder Kursteam unter einer „Klasse“-Zeile mit gesetztem Fach).',
                            canProvision: false
                        });
                        return;
                    }
                    const sug = typeof api.computeCreateSuggestion === 'function' ? api.computeCreateSuggestion(r) : null;
                    const can = !!(sug && sug.displayName && sug.mailNick);
                    out.push({
                        kind: 'group',
                        r,
                        action: can ? 'Kursteam (Unified + Team) per Graph anlegen.' : 'Kursteam: Mail‑Nickname/Schema prüfen.',
                        canProvision: can
                    });
                    return;
                }
                const linkRow = links[String(r.id)] || null;
                const linkedUserForPerson = normStr(r.tenantUserId) || (linkRow && normStr(linkRow.tenantUserId));
                if (t === 'Person' && !linkedUserForPerson) {
                    const dn = normStr(r.personName) || normStr(r.bezeichnung);
                    const em = normStr(r.personEmail).toLowerCase();
                    const can = !!(dn && em && em.indexOf('@') !== -1);
                    out.push({
                        kind: 'person',
                        r,
                        action: can
                            ? 'Entra‑Benutzer per Graph anlegen (Name + E‑Mail vorhanden).'
                            : 'Person: Name und gültige E‑Mail/UPN in Schritt 3 ergänzen.',
                        canProvision: can
                    });
                }
            });
            return out;
        }

        function renderTenantInventoryMatchRows() {
            const tbody = document.getElementById('tenantInvMatchTbody');
            const sum = document.getElementById('tenantInvSummary');
            if (!tbody) return;
            const api = window.ms365TenantInventory;
            if (!api || typeof api.loadStructureState !== 'function') {
                tbody.innerHTML =
                    '<tr><td colspan="3" class="muted">Technik wird geladen … Seite ggf. neu laden, falls diese Zeile bleibt.</td></tr>';
                return;
            }
            const { rows } = api.loadStructureState();
            const links = api.loadMatchLinks() || {};
            const cache = typeof api.readCache === 'function' ? api.readCache() : { rows: [], users: [] };
            const groups = cache.rows || [];
            const users = cache.users || [];
            const orgRows = (rows || []).filter((r) => {
                if (!r) return false;
                const t = String(r.typ || '');
                return (
                    t === 'SchuelerInnen' ||
                    t === 'LehrerInnen' ||
                    t === 'Verwaltung' ||
                    t === 'Jahrgang' ||
                    t === 'Klasse' ||
                    t === 'Gruppe' ||
                    t === 'Kursteam' ||
                    t === 'Person' ||
                    t === 'Arbeitsgemeinschaft'
                );
            });
            if (!orgRows.length) {
                tbody.innerHTML =
                    '<tr><td colspan="3" class="muted">Keine Einträge vom Typ Jahrgang, Klasse, Gruppe, Kursteam, ARGE (Arbeitsgemeinschaft) oder Person in der SOLL‑Struktur (Schritt 3).</td></tr>';
                if (sum) {
                    sum.style.display = '';
                    sum.textContent =
                        'Cache: ' + groups.length + ' Gruppe(n)/Team(s), ' + users.length + ' Benutzerkonto/-konten (nach „Tenant laden“).';
                }
                return;
            }
            tbody.replaceChildren();
            orgRows.forEach((r) => {
                const tr = document.createElement('tr');
                const rid = String(r.id || '');
                const linked =
                    normStr(r.tenantGroupId) ||
                    (links[rid] && normStr(links[rid].tenantGroupId)) ||
                    '';
                const c1 = document.createElement('td');
                c1.textContent = (r.bezeichnung || '–') + ' · ' + (r.typ || '');
                const c2 = document.createElement('td');
                const c3 = document.createElement('td');
                c3.className = 'tenant-inv-actions';
                const t = String(r.typ || '');
                if (t === 'Person') {
                    const pNote = document.createElement('div');
                    pNote.className = 'muted';
                    pNote.style.fontSize = '0.82em';
                    pNote.style.lineHeight = '1.35';
                    const linkP = links[rid] || null;
                    const uidShow = normStr(r.tenantUserId) || (linkP && normStr(linkP.tenantUserId));
                    let label = '';
                    if (uidShow) {
                        const hit = users.find(function (u) {
                            return u && String(u.id) === String(uidShow);
                        });
                        label = hit
                            ? (normStr(hit.displayName) || normStr(hit.userPrincipalName) || uidShow) + ' · ' + uidShow
                            : uidShow;
                    }
                    pNote.textContent = uidShow ? 'Entra: ' + label : '— · Benutzer (Abgleich / Schritt 5)';
                    c2.appendChild(pNote);
                    c3.textContent = '';
                } else {
                    const sel = document.createElement('select');
                    sel.setAttribute('data-tenant-inv-sel', rid);
                    const o0 = document.createElement('option');
                    o0.value = '';
                    o0.textContent = '(keine Entra‑Gruppe)';
                    sel.appendChild(o0);
                    groups.forEach((g) => {
                        const o = document.createElement('option');
                        o.value = String(g.id || '');
                        o.textContent = (g.bezeichnung || '(ohne Name)') + ' · ' + (g.typ || '') + (g.alias ? ' · ' + g.alias : '');
                        sel.appendChild(o);
                    });
                    sel.value = linked || '';
                    c2.appendChild(sel);
                    const bSug = document.createElement('button');
                    bSug.type = 'button';
                    bSug.className = 'btn small-btn tenant-inv-icon-btn';
                    bSug.setAttribute('data-tenant-inv-suggest', rid);
                    bSug.setAttribute('aria-label', 'Vorschlag');
                    bSug.title = 'Vorschlag ins Dropdown (ohne Speichern)';
                    bSug.innerHTML = '<i class="bi bi-magic"></i>';
                    const bSave = document.createElement('button');
                    bSave.type = 'button';
                    bSave.className = 'btn btn-success small-btn tenant-inv-icon-btn';
                    bSave.setAttribute('data-tenant-inv-save', rid);
                    bSave.setAttribute('aria-label', 'Speichern');
                    bSave.title = 'Verknüpfung speichern';
                    bSave.innerHTML = '<i class="bi bi-check2"></i>';
                    const bClr = document.createElement('button');
                    bClr.type = 'button';
                    bClr.className = 'btn small-btn tenant-inv-icon-btn';
                    bClr.setAttribute('data-tenant-inv-clear', rid);
                    bClr.setAttribute('aria-label', 'Verknüpfung löschen');
                    bClr.title = 'Verknüpfung löschen';
                    bClr.innerHTML = '<i class="bi bi-x-lg"></i>';
                    c3.appendChild(bSug);
                    c3.appendChild(bSave);
                    c3.appendChild(bClr);
                }
                tr.appendChild(c1);
                tr.appendChild(c2);
                tr.appendChild(c3);
                tbody.appendChild(tr);
            });
            if (sum) {
                sum.style.display = '';
                sum.textContent =
                    'Cache: ' +
                    groups.length +
                    ' Gruppe(n)/Team(s), ' +
                    users.length +
                    ' Benutzer · ' +
                    orgRows.length +
                    ' Match‑Zeilen.';
            }
        }

        function renderTenantDeltaRows() {
            const tbody = document.getElementById('tenantDeltaTbody');
            const sum = document.getElementById('tenantDeltaSummary');
            if (!tbody) return;
            const api = window.ms365TenantInventory;
            if (!api || typeof api.loadStructureState !== 'function') {
                tbody.innerHTML = '<tr><td colspan="3" class="muted">–</td></tr>';
                return;
            }
            const items = collectTenantDeltaItems(api);
            tbody.replaceChildren();
            if (!items.length) {
                tbody.innerHTML =
                    '<tr><td colspan="3" class="muted">Keine offenen Differenzen (Kursteams nur mit gültigem Klasse‑/Fach‑Kontext; Personen nur mit Name und E‑Mail).</td></tr>';
            } else {
                items.forEach((it) => {
                    const tr = document.createElement('tr');
                    const a = document.createElement('td');
                    a.textContent = it.r.bezeichnung || '–';
                    const b = document.createElement('td');
                    b.textContent = it.r.typ || '';
                    const c = document.createElement('td');
                    c.style.display = 'flex';
                    c.style.flexWrap = 'wrap';
                    c.style.gap = '8px';
                    c.style.alignItems = 'center';
                    const span = document.createElement('span');
                    span.textContent = it.action;
                    span.style.flex = '1';
                    span.style.minWidth = '120px';
                    c.appendChild(span);
                    if (it.canProvision && typeof api.provisionGroupRow === 'function' && it.kind === 'group') {
                        const one = document.createElement('button');
                        one.type = 'button';
                        one.className = 'btn btn-success small-btn';
                        one.setAttribute('data-tenant-delta-one-group', String(it.r.id));
                        one.textContent = 'Jetzt anlegen';
                        c.appendChild(one);
                    }
                    if (it.canProvision && typeof api.provisionPersonRow === 'function' && it.kind === 'person') {
                        const one = document.createElement('button');
                        one.type = 'button';
                        one.className = 'btn btn-success small-btn';
                        one.setAttribute('data-tenant-delta-one-person', String(it.r.id));
                        one.textContent = 'Jetzt anlegen';
                        c.appendChild(one);
                    }
                    tr.appendChild(a);
                    tr.appendChild(b);
                    tr.appendChild(c);
                    tbody.appendChild(tr);
                });
            }
            if (sum) {
                sum.style.display = '';
                const prov = items.filter((x) => x.canProvision).length;
                sum.textContent =
                    items.length +
                    ' offene Position(en), davon ' +
                    prov +
                    ' mit Graph‑Anlage möglich (laut Schema).';
            }
        }

        window.__ms365TenantStepsRefreshMatch = function () {
            renderTenantInventoryMatchRows();
            renderTenantDeltaRows();
        };

        const invTbody = document.getElementById('tenantInvMatchTbody');
        if (invTbody && !invTbody.dataset.tenantInvBound) {
            invTbody.dataset.tenantInvBound = '1';
            invTbody.addEventListener('click', (ev) => {
                const api = window.ms365TenantInventory;
                if (!api || typeof api.saveMatchLink !== 'function') return;
                const t = ev.target;
                const saveB = t.closest && t.closest('[data-tenant-inv-save]');
                const sugB = t.closest && t.closest('[data-tenant-inv-suggest]');
                const clrB = t.closest && t.closest('[data-tenant-inv-clear]');
                if (sugB) {
                    const rid = sugB.getAttribute('data-tenant-inv-suggest');
                    const row = (api.loadStructureState().rows || []).find((x) => String(x.id) === String(rid));
                    if (!row) return;
                    const gid = typeof api.suggestGroupForUnit === 'function' ? api.suggestGroupForUnit(row) : '';
                    const sel = invTbody.querySelector('select[data-tenant-inv-sel="' + rid + '"]');
                    if (sel && gid) sel.value = gid;
                    return;
                }
                if (clrB) {
                    const rid = clrB.getAttribute('data-tenant-inv-clear');
                    api.saveMatchLink(rid, '', '');
                    try {
                        if (typeof api.patchStructureRow === 'function') {
                            api.patchStructureRow(rid, { tenantGroupId: '', tenantMailNickname: '', syncStatus: 'Ausstehend' });
                        }
                    } catch {
                        // ignore
                    }
                    window.__ms365TenantStepsRefreshMatch();
                    return;
                }
                if (saveB) {
                    const rid = saveB.getAttribute('data-tenant-inv-save');
                    const sel = invTbody.querySelector('select[data-tenant-inv-sel="' + rid + '"]');
                    const gid = sel && sel.value ? String(sel.value).trim() : '';
                    api.saveMatchLink(rid, gid, '');
                    if (typeof api.patchStructureRow === 'function') {
                        if (gid) {
                            api.patchStructureRow(rid, {
                                tenantGroupId: gid,
                                syncStatus: 'Ok',
                                letzteFehlermeldung: ''
                            });
                        } else {
                            api.patchStructureRow(rid, {
                                tenantGroupId: '',
                                tenantMailNickname: '',
                                syncStatus: 'Ausstehend',
                                letzteFehlermeldung: ''
                            });
                        }
                    }
                    window.__ms365TenantStepsRefreshMatch();
                }
            });
        }

        const deltaTbody = document.getElementById('tenantDeltaTbody');
        if (deltaTbody && !deltaTbody.dataset.tenantDeltaBound) {
            deltaTbody.dataset.tenantDeltaBound = '1';
            deltaTbody.addEventListener('click', async (ev) => {
                const api = window.ms365TenantInventory;
                if (!api) return;
                const gBtn = ev.target.closest && ev.target.closest('[data-tenant-delta-one-group]');
                const pBtn = ev.target.closest && ev.target.closest('[data-tenant-delta-one-person]');
                const id = (gBtn && gBtn.getAttribute('data-tenant-delta-one-group')) || (pBtn && pBtn.getAttribute('data-tenant-delta-one-person'));
                if (!id) return;
                const row = (api.loadStructureState().rows || []).find((x) => String(x.id) === String(id));
                if (!row) return;
                try {
                    setTenantDeltaProgress(true, 'Anlegen …', 0.2);
                    if (gBtn) await api.provisionGroupRow(row);
                    if (pBtn) {
                        const res = await api.provisionPersonRow(row, {});
                        if (res && res.tempPassword) {
                            await dlgAlert('Benutzer angelegt. Einmaliges Kennwort:\n\n' + res.tempPassword, {
                                title: 'Kennwort notieren',
                                okText: 'Verstanden'
                            });
                        }
                    }
                } catch (e) {
                    await dlgAlert('Fehler: ' + (e && e.message ? e.message : String(e)), { title: 'Fehler' });
                } finally {
                    setTenantDeltaProgress(false, '', null);
                    window.__ms365TenantStepsRefreshMatch();
                }
            });
        }

        async function runTenantInventoryAutomatch() {
            const api = window.ms365TenantInventory;
            const elSt = document.getElementById('tenantInvStatus');
            if (!api || typeof api.loadStructureState !== 'function' || typeof api.suggestGroupForUnit !== 'function') {
                if (elSt) elSt.textContent = 'Automatching: Technik nicht bereit.';
                return;
            }
            const cache = typeof api.readCache === 'function' ? api.readCache() : { rows: [] };
            const gr = cache.rows || [];
            if (!gr.length) {
                await dlgAlert('Zuerst „Tenant laden“ ausführen, damit Entra-Gruppen für das Automatching vorliegen.', {
                    title: 'Automatching'
                });
                return;
            }
            const gidSet = new Set(gr.map((g) => String(g.id || '').trim()).filter(Boolean));
            const st0 = api.loadStructureState();
            const rows = st0.rows || [];
            let links = api.loadMatchLinks() || {};
            const orgRows = rows.filter((r) => {
                if (!r || r.isStructureTreeRoot) return false;
                const ty = String(r.typ || '');
                return (
                    ty === 'Jahrgang' ||
                    ty === 'Klasse' ||
                    ty === 'Gruppe' ||
                    ty === 'Kursteam' ||
                    ty === 'Arbeitsgemeinschaft'
                );
            });
            let saved = 0;
            let skippedHas = 0;
            let skippedNoHit = 0;
            for (let i = 0; i < orgRows.length; i++) {
                const r = orgRows[i];
                const rid = String(r.id || '');
                const linked =
                    normStr(r.tenantGroupId) ||
                    (links[rid] && normStr(links[rid].tenantGroupId));
                if (linked) {
                    skippedHas++;
                    continue;
                }
                const gid = String(api.suggestGroupForUnit(r) || '').trim();
                if (!gid || !gidSet.has(gid)) {
                    skippedNoHit++;
                    continue;
                }
                try {
                    api.saveMatchLink(rid, gid, 'Automatching');
                    if (typeof api.patchStructureRow === 'function') {
                        api.patchStructureRow(rid, {
                            tenantGroupId: gid,
                            syncStatus: 'Ok',
                            letzteFehlermeldung: ''
                        });
                    }
                    saved++;
                    links = api.loadMatchLinks() || links;
                    r.tenantGroupId = gid;
                } catch {
                    skippedNoHit++;
                }
            }
            window.__ms365TenantStepsRefreshMatch();
            const parts = ['Automatching: ' + saved + ' neu verknüpft.'];
            if (skippedHas) parts.push(skippedHas + ' bereits gesetzt.');
            if (skippedNoHit) parts.push(skippedNoHit + ' ohne Treffer.');
            if (elSt) elSt.textContent = parts.join(' ');
        }

        const invAuto = document.getElementById('tenantInvAutoMatchBtn');
        if (invAuto && !invAuto.dataset.tenantInvAutoBound) {
            invAuto.dataset.tenantInvAutoBound = '1';
            invAuto.addEventListener('click', () => void runTenantInventoryAutomatch());
        }

        const invBtn = document.getElementById('tenantInvRefreshBtn');
        if (invBtn) {
            invBtn.addEventListener('click', async () => {
                const st = document.getElementById('tenantInvStatus');
                const api = window.ms365TenantInventory;
                if (!api || typeof api.refresh !== 'function') {
                    if (st) st.textContent = 'Schulstruktur‑Modul nicht geladen. Seite neu laden.';
                    return;
                }
                invBtn.disabled = true;
                if (st) st.textContent = 'Lade Daten über Microsoft Graph …';
                try {
                    await api.refresh((ev) => {
                        if (!st) return;
                        const ph = ev && ev.phase === 'users' ? 'Benutzer' : 'Gruppen/Teams';
                        const pg = ev && ev.page != null ? ' (Seite ' + ev.page + ')' : '';
                        st.textContent = ph + ' werden geladen' + pg + ' …';
                    });
                    const c = api.readCache();
                    if (st) {
                        st.textContent =
                            'Fertig: ' +
                            (c.rows || []).length +
                            ' Gruppe(n)/Team(s), ' +
                            (c.users || []).length +
                            ' Benutzerkonto/-konten.';
                    }
                    renderTenantInventoryMatchRows();
                    renderTenantDeltaRows();
                } catch (e) {
                    if (st) st.textContent = 'Fehler: ' + (e && e.message ? e.message : String(e));
                } finally {
                    invBtn.disabled = false;
                }
            });
        }
        const deltaBtn = document.getElementById('tenantDeltaRefreshBtn');
        if (deltaBtn) {
            deltaBtn.addEventListener('click', () => renderTenantDeltaRows());
        }

        async function runBatchProvision(kind) {
            const api = window.ms365TenantInventory;
            if (!api) return;
            const items = collectTenantDeltaItems(api).filter((x) => x.canProvision && x.kind === kind);
            if (!items.length) {
                await dlgAlert(
                    kind === 'group'
                        ? 'Keine anlegbaren Gruppen (Jahrgang, Klasse, Gruppe, Kursteam, ARGE – Schema prüfen).'
                        : 'Keine anlegbaren Personen (Name + E‑Mail).',
                    { title: 'Delta-Anlage' }
                );
                return;
            }
            const n = items.length;
            const ok = await dlgConfirm(
                kind === 'group'
                    ? n + ' Gruppe(n)/Team(s) (inkl. Klasse, Jahrgang, ARGE, Kursteam) wirklich in Entra anlegen?'
                    : n + ' Benutzerkonto/-konten wirklich in Entra anlegen?',
                { title: 'Entra-Anlage', okText: 'Anlegen', danger: true }
            );
            if (!ok) return;
            const btnG = document.getElementById('tenantDeltaProvisionGroupsBtn');
            const btnP = document.getElementById('tenantDeltaProvisionPersonsBtn');
            if (btnG) btnG.disabled = true;
            if (btnP) btnP.disabled = true;
            let pwdNotes = [];
            for (let i = 0; i < items.length; i++) {
                const it = items[i];
                const ratio = (i + 0.35) / n;
                setTenantDeltaProgress(true, 'Anlegen ' + (i + 1) + ' / ' + n + ' …', ratio);
                try {
                    if (kind === 'group') await api.provisionGroupRow(it.r);
                    else {
                        const res = await api.provisionPersonRow(it.r, { skipConfirm: true });
                        if (res && res.tempPassword) pwdNotes.push((it.r.bezeichnung || it.r.personName || '') + ': ' + res.tempPassword);
                    }
                } catch (e) {
                    await dlgAlert('Abbruch bei Position ' + (i + 1) + ': ' + (e && e.message ? e.message : String(e)), {
                        title: 'Fehler'
                    });
                    break;
                }
            }
            if (btnG) btnG.disabled = false;
            if (btnP) btnP.disabled = false;
            setTenantDeltaProgress(true, 'Fertig.', 1);
            setTimeout(() => setTenantDeltaProgress(false, '', null), 1400);
            if (pwdNotes.length) {
                await dlgAlert('Temporäre Kennwörter (bitte sicher notieren):\n\n' + pwdNotes.join('\n'), {
                    title: 'Kennwörter',
                    okText: 'Verstanden'
                });
            }
            window.__ms365TenantStepsRefreshMatch();
        }

        const btnBatchG = document.getElementById('tenantDeltaProvisionGroupsBtn');
        if (btnBatchG) btnBatchG.addEventListener('click', () => runBatchProvision('group'));
        const btnBatchP = document.getElementById('tenantDeltaProvisionPersonsBtn');
        if (btnBatchP) btnBatchP.addEventListener('click', () => runBatchProvision('person'));

        renderTenantInventoryMatchRows();
        renderTenantDeltaRows();

        // Accordion: immer nur EIN Schritt offen (details.step)
        try {
            const steps = Array.from(document.querySelectorAll('details.step'));
            steps.forEach((d) => {
                d.addEventListener('toggle', () => {
                    if (!d.open) return;
                    steps.forEach((o) => {
                        if (o !== d) o.open = false;
                    });
                });
            });
        } catch {
            // ignore
        }
    }

    bindUi();
})();

