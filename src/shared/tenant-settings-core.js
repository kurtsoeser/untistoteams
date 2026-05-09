(function () {
    'use strict';

    const STORAGE_KEY = 'ms365-tenant-settings-v1';
    const CURRENT_VERSION = 2;

    function normStr(v) {
        return String(v ?? '').trim();
    }

    function normCode(v) {
        return normStr(v).toUpperCase();
    }

    /** Stabiler Mail-Nickname für Klassen-M365-Gruppe: jg{YYYY}{codeAlphaNum} (Kursteam/Umbenennen). */
    function deriveClassStableMailNickname(yearRaw, codeRaw) {
        const y = normStr(yearRaw);
        const yy = /^\d{4}$/.test(y) ? y : '';
        const code = normCode(codeRaw);
        const tail = String(code || '')
            .replace(/[^0-9A-Za-z]/g, '')
            .toLowerCase()
            .slice(0, 24);
        if (!yy || !tail) return '';
        return ('jg' + yy + tail).toLowerCase().slice(0, 60);
    }

    function safeJsonParse(s) {
        try {
            return JSON.parse(String(s));
        } catch {
            return null;
        }
    }

    function loadRaw() {
        try {
            const raw = localStorage.getItem(STORAGE_KEY);
            if (!raw) return null;
            return safeJsonParse(raw);
        } catch {
            return null;
        }
    }

    function normalizeSettings(obj) {
        const o = obj && typeof obj === 'object' ? obj : {};
        const domain =
            typeof window.ms365GetSchoolDomainNoAt === 'function'
                ? window.ms365GetSchoolDomainNoAt()
                : normStr(o.domain);

        const subjectsIn = Array.isArray(o.subjects) ? o.subjects : [];
        const argesIn = Array.isArray(o.arges) ? o.arges : [];
        const teachersIn = Array.isArray(o.teachers) ? o.teachers : [];
        const adminIn = Array.isArray(o.admin) ? o.admin : (Array.isArray(o.administration) ? o.administration : []);
        const studentsIn = Array.isArray(o.students) ? o.students : [];
        const classesIn = Array.isArray(o.classes) ? o.classes : [];

        const subjectsSeen = new Set();
        const subjects = [];
        subjectsIn.forEach((s) => {
            const code = normCode(s?.code);
            const name = normStr(s?.name);
            if (!code) return;
            const key = code.toLowerCase();
            if (subjectsSeen.has(key)) return;
            subjectsSeen.add(key);
            subjects.push({ code, name });
        });

        const argesSeen = new Set();
        const arges = [];
        argesIn.forEach((a) => {
            const code = normCode(a?.code);
            const name = normStr(a?.name);
            const subjectsRaw = Array.isArray(a?.subjects) ? a.subjects : Array.isArray(a?.faecher) ? a.faecher : [];
            const subjects = (subjectsRaw || [])
                .map((x) => normCode(x))
                .filter(Boolean);
            if (!code) return;
            const key = code.toLowerCase();
            if (argesSeen.has(key)) return;
            argesSeen.add(key);
            arges.push({ code, name, subjects });
        });

        const teachersSeen = new Set();
        const teachers = [];
        teachersIn.forEach((t) => {
            const code = normCode(t?.code);
            const name = normStr(t?.name);
            const email = normStr(t?.email).toLowerCase();
            if (!code) return;
            const key = code.toLowerCase();
            if (teachersSeen.has(key)) return;
            teachersSeen.add(key);
            teachers.push({ code, name, email });
        });

        const adminSeen = new Set();
        const admin = [];
        adminIn.forEach((a) => {
            const role = normStr(a?.role || a?.rolle || a?.title);
            const name = normStr(a?.name);
            const email = normStr(a?.email).toLowerCase();
            const defaultKey = normStr(a?.defaultKey);
            if (!role && !name && !email) return;
            const key = (defaultKey || role || name || email).toLowerCase();
            if (adminSeen.has(key)) return;
            adminSeen.add(key);
            const row = { role, name, email };
            if (defaultKey) row.defaultKey = defaultKey;
            admin.push(row);
        });

        const students = [];
        studentsIn.forEach((s) => {
            const klasse = normStr(s?.klasse || s?.class || s?.group || s?.Klassse || s?.Klasse);
            const name = normStr(s?.name);
            const email = normStr(s?.email).toLowerCase();
            if (!klasse && !name && !email) return;
            students.push({ klasse, name, email });
        });

        const classesSeen = new Set();
        const classes = [];
        classesIn.forEach((c) => {
            const code = normCode(c?.code);
            const name = normStr(c?.name || c?.klasse || c?.Klasse);
            const yearRaw = normStr(c?.year || c?.abschlussjahr || c?.Abschlussjahr || c?.graduationYear || '');
            const year = /^\d{4}$/.test(yearRaw) ? yearRaw : '';
            const headName = normStr(c?.headName || c?.klassenvorstandName || c?.kvName);
            const headEmail = normStr(c?.headEmail || c?.klassenvorstandEmail || c?.kvEmail).toLowerCase();
            let stableMailNickname = normStr(c?.stableMailNickname || '')
                .replace(/[^a-zA-Z0-9]/g, '')
                .toLowerCase()
                .slice(0, 60);
            if (!stableMailNickname && year && code) {
                stableMailNickname = deriveClassStableMailNickname(year, code);
            }
            if (!code && !name && !year && !headName && !headEmail) return;
            const key = (code || name).toLowerCase();
            if (classesSeen.has(key)) return;
            classesSeen.add(key);
            classes.push({ code, name, year, headName, headEmail, stableMailNickname });
        });

        return {
            version: CURRENT_VERSION,
            domain: normStr(domain),
            subjects,
            arges,
            teachers,
            admin,
            students,
            classes
        };
    }

    function save(settings) {
        const normalized = normalizeSettings(settings);
        try {
            localStorage.setItem(STORAGE_KEY, JSON.stringify(normalized));
        } catch {
            // ignore
        }
        try {
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.setCoreFromTenantSettings === 'function') {
                window.ms365AppDataV2.setCoreFromTenantSettings(normalized);
            }
        } catch {
            // ignore
        }
        if (typeof window.ms365SetSchoolDomainNoAt === 'function' && normalized.domain) {
            window.ms365SetSchoolDomainNoAt(normalized.domain);
        }
        return normalized;
    }

    function load() {
        try {
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getContainer === 'function') {
                const c = window.ms365AppDataV2.getContainer();
                if (c && c.core && c.years) {
                    const cur = String((c.years && c.years.current) || '');
                    const y = (c.years && c.years.byLabel && cur && c.years.byLabel[cur]) ? c.years.byLabel[cur] : { students: [], classes: [] };
                    return normalizeSettings({
                        domain: c.core.domain,
                        subjects: c.core.subjects,
                        arges: c.core.arges,
                        teachers: c.core.teachers,
                        admin: c.core.admin,
                        students: y.students,
                        classes: y.classes
                    });
                }
            }
        } catch {
            // ignore
        }
        const raw = loadRaw();
        return normalizeSettings(raw || {});
    }

    function getTeacherEmailMap() {
        const s = load();
        const map = {};
        s.teachers.forEach((t) => {
            if (t.code && t.email) map[t.code] = t.email;
        });
        return map;
    }

    function parseDelimitedLines(text) {
        const lines = String(text || '').split(/\r\n|\n|\r/);
        const out = [];
        lines.forEach((line) => {
            const t = normStr(line);
            if (!t || t.startsWith('#')) return;
            const parts = t
                .split(/[;\t,|]/)
                .map((x) => normStr(x))
                .filter(Boolean);
            if (!parts.length) return;
            out.push(parts);
        });
        return out;
    }

    function parseLinesToSubjects(text) {
        const out = [];
        parseDelimitedLines(text).forEach((parts) => {
            const code = normCode(parts[0] || '');
            const name = normStr(parts.slice(1).join(' '));
            if (!code) return;
            out.push({ code, name });
        });
        return out;
    }

    function parseLinesToTeachers(text) {
        const out = [];
        parseDelimitedLines(text).forEach((parts) => {
            const code = normCode(parts[0] || '');
            const name = normStr(parts[1] || '');
            const email = normStr(parts[2] || '').toLowerCase();
            if (!code) return;
            out.push({ code, name, email });
        });
        return out;
    }

    function parseLinesToStudents(text) {
        const out = [];
        parseDelimitedLines(text).forEach((parts) => {
            const klasse = normStr(parts[0] || '');
            const name = normStr(parts[1] || '');
            const email = normStr(parts[2] || '').toLowerCase();
            if (!klasse && !name && !email) return;
            out.push({ klasse, name, email });
        });
        return out;
    }

    function parseLinesToClasses(text) {
        const out = [];
        parseDelimitedLines(text).forEach((parts) => {
            const code = normCode(parts[0] || '');
            // Unterstützte Formate:
            // - code;name;headName;headEmail (alt)
            // - code;year;name;headName;headEmail (neu)
            // - code;name;year;headName;headEmail (tolerant)
            let year = '';
            let name = '';
            let headName = '';
            let headEmail = '';

            if (parts.length >= 2 && /^\d{4}$/.test(normStr(parts[1] || ''))) {
                year = normStr(parts[1] || '');
                name = normStr(parts[2] || '');
                headName = normStr(parts[3] || '');
                headEmail = normStr(parts[4] || '').toLowerCase();
            } else if (parts.length >= 3 && /^\d{4}$/.test(normStr(parts[2] || ''))) {
                name = normStr(parts[1] || '');
                year = normStr(parts[2] || '');
                headName = normStr(parts[3] || '');
                headEmail = normStr(parts[4] || '').toLowerCase();
            } else {
                name = normStr(parts[1] || '');
                headName = normStr(parts[2] || '');
                headEmail = normStr(parts[3] || '').toLowerCase();
            }

            const y = /^\d{4}$/.test(year) ? year : '';
            if (!code && !name && !y && !headName && !headEmail) return;
            const stableMailNickname = deriveClassStableMailNickname(y, code);
            out.push({ code, name, year: y, headName, headEmail, stableMailNickname });
        });
        return out;
    }

    // Public API (kompatibel zu bisher)
    window.ms365TenantSettingsLoad = load;
    window.ms365TenantSettingsSave = save;
    window.ms365TenantSettingsGetTeacherEmailMap = getTeacherEmailMap;
    window.ms365TenantSettingsParseSubjectsLines = parseLinesToSubjects;
    window.ms365TenantSettingsParseArgesLines = function (text) {
        const out = [];
        parseDelimitedLines(text).forEach((parts) => {
            const code = normCode(parts[0] || '');
            const name = normStr(parts[1] || '');
            const subjRaw = normStr(parts.slice(2).join(' '));
            const subjects = subjRaw
                ? subjRaw
                      .split(/[,\s|]+/)
                      .map((x) => normCode(x))
                      .filter(Boolean)
                : [];
            if (!code) return;
            out.push({ code, name, subjects });
        });
        return out;
    };
    window.ms365TenantSettingsParseTeachersLines = parseLinesToTeachers;
    window.ms365TenantSettingsParseAdminLines = function (text) {
        const out = [];
        parseDelimitedLines(text).forEach((parts) => {
            const role = normStr(parts[0] || '');
            const name = normStr(parts[1] || '');
            const email = normStr(parts[2] || '').toLowerCase();
            if (!role && !name && !email) return;
            out.push({ role, name, email });
        });
        return out;
    };
    window.ms365TenantSettingsParseStudentsLines = parseLinesToStudents;
    window.ms365TenantSettingsParseClassesLines = parseLinesToClasses;
    window.ms365DeriveClassStableMailNickname = deriveClassStableMailNickname;
})();

