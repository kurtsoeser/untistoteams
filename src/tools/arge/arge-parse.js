import { normStr } from '../../shared/utils/strings.js';

function normSubjectKey(v) {
    return normStr(v)
        .toLowerCase()
        .replace(/\s+/g, ' ')
        .replace(/[äÄ]/g, 'ae')
        .replace(/[öÖ]/g, 'oe')
        .replace(/[üÜ]/g, 'ue')
        .replace(/ß/g, 'ss');
}

function buildSubjectCodeFromName(name) {
    const raw = normStr(name);
    if (!raw) return '';
    let s = raw
        .replace(/[äÄ]/g, 'AE')
        .replace(/[öÖ]/g, 'OE')
        .replace(/[üÜ]/g, 'UE')
        .replace(/ß/g, 'SS')
        .toUpperCase()
        .replace(/[^A-Z0-9]+/g, '');
    if (!s) s = 'SUBJECT';
    return s.slice(0, 12);
}

function getTenantSubjectsByNameKey() {
    try {
        if (typeof window.ms365TenantSettingsLoad !== 'function') return new Map();
        const s = window.ms365TenantSettingsLoad();
        const subjects = Array.isArray(s?.subjects) ? s.subjects : [];
        const map = new Map();
        subjects.forEach((x) => {
            const name = normStr(x?.name);
            const code = normStr(x?.code);
            if (!name || !code) return;
            map.set(normSubjectKey(name), code);
        });
        return map;
    } catch {
        return new Map();
    }
}

function subjectForSlug(line) {
    let t = String(line || '').trim();
    const stripped = t.replace(/^ARGE\s+/i, '').trim();
    return stripped || t;
}

function looksLikeSubjectCode(s) {
    const t = normStr(s);
    if (!t) return false;
    return /^[A-Za-z0-9]{1,12}$/.test(t);
}

function normSubjectCode(s) {
    return normStr(s).toUpperCase().replace(/[^A-Z0-9]/g, '').slice(0, 12);
}

function getSubjectCodeForDisplayName(displayName, fallbackSubjectCode) {
    const fb = normStr(fallbackSubjectCode);
    if (looksLikeSubjectCode(fb)) return normSubjectCode(fb);
    const name = normStr(subjectForSlug(displayName));
    if (!name) return '';
    const map = getTenantSubjectsByNameKey();
    const existing = map.get(normSubjectKey(name));
    return existing || buildSubjectCodeFromName(name);
}

function toNickBaseFromName(displayName) {
    let s = String(displayName || '').trim();
    s = s
        .replace(/[äÄ]/g, 'ae')
        .replace(/[öÖ]/g, 'oe')
        .replace(/[üÜ]/g, 'ue')
        .replace(/ß/g, 'ss');
    s = s.replace(/[^A-Za-z0-9]+/g, '-').replace(/-+/g, '-');
    s = s.replace(/^-+|-+$/g, '');
    return s;
}

/** Anzeigename der M365-Gruppe aus einer einfachen Fach-Zeile */
function displayNameFromSubjectLine(line) {
    const t = line.trim();
    if (!t) return '';
    if (/^ARGE\s+/i.test(t)) return t;
    return 'ARGE ' + t;
}

/**
 * Eingabeformate (pro Zeile):
 * - "Deutsch"
 * - "D;Deutsch" (Kürzel;Fach)
 * - "D;Deutsch;arge-deutsch" (Kürzel;Fach;MailNickname)
 * - "ARGE Deutsch;arge-deutsch" (Anzeigename;MailNickname)
 */
function parseArgeLine(t) {
    const parts = String(t || '')
        .split(/[;\t]/)
        .map((x) => normStr(x))
        .filter(Boolean);
    if (!parts.length) return null;

    if (parts.length >= 2 && /^ARGE\s+/i.test(parts[0])) {
        const displayName = parts[0];
        const subjectName = normStr(subjectForSlug(displayName));
        const subjectCode = '';
        return { mode: 'display', displayName, subjectName, subjectCode, explicitNick: parts[1] || '' };
    }

    if (parts.length >= 2 && looksLikeSubjectCode(parts[0])) {
        const subjectCode = normSubjectCode(parts[0]);
        const subjectName = normStr(parts[1]);
        if (!subjectName) return null;
        const displayName = displayNameFromSubjectLine(subjectName);
        const explicitNick = parts[2] || '';
        return { mode: 'subject', displayName, subjectName, subjectCode, explicitNick };
    }

    if (parts.length >= 2) {
        const displayName = parts[0];
        const subjectName = normStr(subjectForSlug(displayName));
        const subjectCode = '';
        return { mode: 'display', displayName, subjectName, subjectCode, explicitNick: parts[1] || '' };
    }

    const subjectName = parts[0];
    const displayName = displayNameFromSubjectLine(subjectName);
    return { mode: 'simple', displayName, subjectName, subjectCode: '', explicitNick: '' };
}

function resolveDuplicateNicks(rows) {
    const seen = new Map();
    rows.forEach((r) => {
        const base = r.mailNick;
        let candidate = base;
        let n = 2;
        while (seen.has(candidate)) {
            candidate = base + '-' + n;
            n++;
        }
        r.mailNick = candidate;
        seen.set(candidate, true);
    });
}

/** Mail-Nickname: bevorzugt aus Fach-Kürzel (z.B. arge-enws), sonst aus Fachname */
function buildMailNickname(displayName, subjectCode, deps) {
    const pre = deps.getPrefix();
    const code = normStr(subjectCode);
    if (looksLikeSubjectCode(code)) {
        const combined = pre ? pre + '-' + normSubjectCode(code) : normSubjectCode(code);
        return deps.maybeUpper(combined).replace(/[^A-Za-z0-9-]/g, '');
    }
    const base = toNickBaseFromName(subjectForSlug(displayName));
    if (!base) return '';
    const combined = pre ? pre + '-' + base : base;
    return deps.maybeUpper(combined).replace(/[^A-Za-z0-9-]/g, '');
}

/**
 * @param {string} text Textarea-Inhalt
 * @param {{ getPrefix: () => string, maybeUpper: (s: string) => string }} deps
 */
function parseArgeInputLines(text, deps) {
    const lines = String(text || '').split(/\r\n|\n|\r/);
    const parsed = [];
    const errors = [];
    const seen = new Set();
    lines.forEach((line, idx) => {
        const t = line.trim();
        if (!t || t.startsWith('#')) return;
        const pl = parseArgeLine(t);
        if (!pl) return;

        let displayName = pl.displayName;
        let mailNick;
        let technicalSlug = toNickBaseFromName(subjectForSlug(displayName));
        const subjectName = pl.subjectName || subjectForSlug(displayName);
        const subjectCode = pl.subjectCode || '';

        let mailNickExplicit = false;
        const explicitNick = pl.explicitNick || '';
        if (explicitNick) {
            mailNick = deps.maybeUpper(explicitNick.replace(/[^A-Za-z0-9-]/g, ''));
            mailNickExplicit = true;
        } else {
            mailNick = buildMailNickname(displayName, subjectCode, deps);
        }

        if (!displayName) return;
        if (!mailNick) {
            errors.push('Zeile ' + (idx + 1) + ': Mail-Nickname konnte nicht erzeugt werden.');
            return;
        }
        const key = displayName.toLowerCase();
        if (seen.has(key)) return;
        seen.add(key);
        parsed.push({
            displayName,
            mailNick,
            owner: '',
            memberLines: '',
            description: 'ARGE-Gruppe: ' + displayName,
            technicalSlug,
            mailNickExplicit,
            subjectName,
            subjectCode: subjectCode ? normSubjectCode(subjectCode) : ''
        });
    });
    return { parsed, errors };
}

/**
 * @param {Array} rows
 * @param {{ getPrefix: () => string, maybeUpper: (s: string) => string, getSubjectCodeForDisplayName: (dn: string, fb: string) => string }} deps
 */
function recomputeArgePreviewMailNicks(rows, deps) {
    rows.forEach((r) => {
        r.technicalSlug = toNickBaseFromName(subjectForSlug(r.displayName));
        r.subjectName = normStr(r.subjectName || subjectForSlug(r.displayName));
        r.subjectCode = normStr(r.subjectCode || '');
        if (!looksLikeSubjectCode(r.subjectCode)) {
            r.subjectCode = deps.getSubjectCodeForDisplayName(r.displayName, r.subjectCode);
        }
        if (!r.mailNickExplicit) {
            r.mailNick = buildMailNickname(r.displayName, r.subjectCode, deps);
        } else {
            r.mailNick = deps.maybeUpper(String(r.mailNick || '').replace(/[^A-Za-z0-9-]/g, ''));
        }
    });
    resolveDuplicateNicks(rows);
}

/**
 * @param {Array} rows
 * @param {{ normStr: (v: unknown) => string, subjectForSlug: (line: string) => string, normSubjectCode: (s: string) => string, looksLikeSubjectCode: (s: string) => boolean }} deps
 */
function serializePreviewRowsToLines(rows, deps) {
    return rows.map((r) => {
        const dn = (r.displayName || '').trim();
        const mn = String(r.mailNick || '').trim();
        const subj = deps.normStr(r.subjectName || subjectForSlug(dn));
        const code = deps.normStr(r.subjectCode || '');
        const hasCode = deps.looksLikeSubjectCode(code);
        if (hasCode && subj) {
            if (r.mailNickExplicit && mn) return `${deps.normSubjectCode(code)};${subj};${mn}`;
            return `${deps.normSubjectCode(code)};${subj}`;
        }
        if (r.mailNickExplicit && mn) return dn + ';' + mn;
        return dn;
    });
}

/**
 * @param {{ text: string, previousArgeRows: Array, deps: { getPrefix: () => string, maybeUpper: (s: string) => string } }} ctx
 */
function syncRowsFromInputPreservingOwners(ctx) {
    const { text, previousArgeRows, deps } = ctx;
    const { parsed, errors } = parseArgeInputLines(text, deps);
    if (errors.length) {
        return { ok: false, errors };
    }
    if (!parsed.length) {
        return { ok: false, errors: ['Bitte mindestens eine ARGE-Zeile eintragen.'] };
    }
    const rows = parsed.map((r) => ({ ...r }));
    resolveDuplicateNicks(rows);
    const ownerByKey = new Map(previousArgeRows.map((r) => [r.displayName.toLowerCase(), r.owner]));
    const memberLinesByKey = new Map(previousArgeRows.map((r) => [r.displayName.toLowerCase(), r.memberLines || '']));
    const codeByKey = new Map(previousArgeRows.map((r) => [r.displayName.toLowerCase(), normStr(r.subjectCode || '')]));
    const subjNameByKey = new Map(previousArgeRows.map((r) => [r.displayName.toLowerCase(), normStr(r.subjectName || '')]));
    const merged = rows.map((r) => ({
        displayName: r.displayName,
        mailNick: r.mailNick,
        owner: ownerByKey.get(r.displayName.toLowerCase()) || '',
        memberLines: memberLinesByKey.get(r.displayName.toLowerCase()) || '',
        description: r.description,
        subjectName: normStr(r.subjectName) || subjNameByKey.get(r.displayName.toLowerCase()) || normStr(subjectForSlug(r.displayName)),
        subjectCode: normStr(r.subjectCode) || codeByKey.get(r.displayName.toLowerCase()) || ''
    }));
    return { ok: true, rows: merged };
}

export {
normStr,
normSubjectKey,
buildSubjectCodeFromName,
getTenantSubjectsByNameKey,
getSubjectCodeForDisplayName,
toNickBaseFromName,
subjectForSlug,
looksLikeSubjectCode,
normSubjectCode,
parseArgeLine,
displayNameFromSubjectLine,
resolveDuplicateNicks,
buildMailNickname,
parseArgeInputLines,
recomputeArgePreviewMailNicks,
serializePreviewRowsToLines,
syncRowsFromInputPreservingOwners
};
