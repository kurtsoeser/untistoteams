import { normStr } from '../../shared/utils/strings.js';

function csvEscapeField(v, delimiter) {
    const d = delimiter || ';';
    const s = String(v ?? '');
    const mustQuote = s.includes('"') || s.includes('\n') || s.includes('\r') || s.includes(d);
    if (!mustQuote) return s;
    return '"' + s.replace(/"/g, '""') + '"';
}

function toCsv(rows, delimiter) {
    const d = delimiter || ';';
    return rows
        .map((r) => r.map((c) => csvEscapeField(c, d)).join(d))
        .join('\r\n');
}

function parseCsvLine(line, delimiter) {
    const d = delimiter || ';';
    const out = [];
    let cur = '';
    let inQ = false;
    for (let i = 0; i < line.length; i++) {
        const ch = line[i];
        if (inQ) {
            if (ch === '"') {
                const next = line[i + 1];
                if (next === '"') {
                    cur += '"';
                    i++;
                } else {
                    inQ = false;
                }
            } else {
                cur += ch;
            }
        } else {
            if (ch === '"') inQ = true;
            else if (ch === d) {
                out.push(cur);
                cur = '';
            } else {
                cur += ch;
            }
        }
    }
    out.push(cur);
    return out.map((x) => normStr(x));
}

function detectCsvDelimiter(text) {
    const head = String(text || '')
        .split(/\r\n|\n|\r/)
        .slice(0, 5)
        .join('\n');
    const semis = (head.match(/;/g) || []).length;
    const commas = (head.match(/,/g) || []).length;
    return semis >= commas ? ';' : ',';
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

/**
 * @param {{
 *   domain: string,
 *   argeRows: Array,
 *   argePreviewRows: Array,
 *   syncArgePreviewFromTextarea: () => void,
 *   normStr: (v: unknown) => string,
 *   subjectForSlug: (line: string) => string,
 *   getSubjectCodeForDisplayName: (displayName: string, fallbackSubjectCode: string) => string
 * }} ctx
 */
export function exportArgeCsv(ctx) {
    const domain = ctx.domain;
    const argeRows = ctx.argeRows || [];
    const argePreviewRows = ctx.argePreviewRows || [];
    const syncArgePreviewFromTextarea = ctx.syncArgePreviewFromTextarea;
    const norm = ctx.normStr;
    const subjectForSlug = ctx.subjectForSlug;
    const getSubjectCodeForDisplayName = ctx.getSubjectCodeForDisplayName;

    const ownerByKey = new Map(argeRows.map((r) => [norm(r.displayName).toLowerCase(), norm(r.owner || '')]));
    const memByKey = new Map(argeRows.map((r) => [norm(r.displayName).toLowerCase(), String(r.memberLines ?? '')]));
    const codeByKey = new Map(argeRows.map((r) => [norm(r.displayName).toLowerCase(), norm(r.subjectCode || '')]));
    const subjByKey = new Map(
        argeRows.map((r) => [norm(r.displayName).toLowerCase(), norm(r.subjectName || subjectForSlug(r.displayName))])
    );

    if (!argePreviewRows.length) syncArgePreviewFromTextarea();

    const dataRows = argePreviewRows.map((r) => {
        const dn = norm(r.displayName);
        const k = dn.toLowerCase();
        const subjectName = norm(r.subjectName || subjByKey.get(k) || subjectForSlug(dn));
        const subjectCode = norm(r.subjectCode || codeByKey.get(k) || getSubjectCodeForDisplayName(dn));
        const owner = ownerByKey.get(k) || '';
        const members = (memByKey.get(k) || '').replace(/\r\n|\n|\r/g, '|');
        const email = norm(r.mailNick) ? `${norm(r.mailNick)}@${domain}` : '';
        return [subjectCode, subjectName, dn, norm(r.mailNick), email, owner, members];
    });

    const csv = toCsv(
        [['subjectCode', 'subjectName', 'displayName', 'mailNick', 'groupEmail', 'ownerUpn', 'membersUpnPipe'], ...dataRows],
        ';'
    );
    const filename = `arge-export-${new Date().toISOString().slice(0, 10)}.csv`;
    return { csv, filename };
}

/**
 * @param {string} text
 * @param {{
 *   normStr: (v: unknown) => string,
 *   displayNameFromSubjectLine: (line: string) => string,
 *   subjectForSlug: (line: string) => string,
 *   normSubjectCode: (s: string) => string
 * }} deps
 * @returns {Array<{ displayName: string, subjectCode: string, subjectName: string, mailNick: string, owner: string, memberLines: string }>}
 */
export function parseArgeCsvToRows(text, deps) {
    const norm = deps.normStr;
    const displayNameFromSubjectLine = deps.displayNameFromSubjectLine;
    const subjectForSlug = deps.subjectForSlug;
    const normSubjectCode = deps.normSubjectCode;

    const raw = String(text || '');
    const delimiter = detectCsvDelimiter(raw);
    const lines = raw.split(/\r\n|\n|\r/).filter((l) => norm(l));
    if (!lines.length) return null;

    const first = parseCsvLine(lines[0], delimiter).map((x) => normHeaderKey(x));
    const hasHeader = first.some((h) =>
        ['subjectcode', 'subjectname', 'displayname', 'mailnick', 'ownerupn', 'members', 'membersupnpipe'].includes(h)
    );

    const idx = (key) => first.indexOf(key);
    const get = (row, key, altIdx) => {
        const i = hasHeader ? idx(key) : altIdx;
        if (i == null || i < 0) return '';
        return norm(row[i] ?? '');
    };

    const outRows = [];
    const start = hasHeader ? 1 : 0;
    for (let i = start; i < lines.length; i++) {
        const row = parseCsvLine(lines[i], delimiter);
        const subjectCode = get(row, 'subjectcode', 0);
        const subjectName = get(row, 'subjectname', 1);
        const displayName = get(row, 'displayname', 2) || (subjectName ? displayNameFromSubjectLine(subjectName) : '');
        const mailNick = get(row, 'mailnick', 3);
        const owner = get(row, 'ownerupn', 5);
        const membersPipe = get(row, 'membersupnpipe', 6) || get(row, 'members', 6);
        const memberLines = membersPipe ? membersPipe.split('|').map((x) => norm(x)).filter(Boolean).join('\n') : '';
        if (!subjectName && !displayName) continue;
        outRows.push({
            displayName,
            subjectCode: subjectCode ? normSubjectCode(subjectCode) : '',
            subjectName: subjectName || norm(subjectForSlug(displayName)),
            mailNick,
            owner,
            memberLines
        });
    }
    return outRows;
}
