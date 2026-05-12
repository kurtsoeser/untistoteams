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
 *   jgRows: Array,
 *   jgPreviewRows: Array,
 *   syncJgPreviewRowsFromTextarea: () => void,
 *   normStr: (v: unknown) => string,
 *   jgM365DisplayName: (row: object) => string
 * }} ctx
 */
export function exportJgCsv(ctx) {
    const domain = ctx.domain;
    const jgRows = ctx.jgRows || [];
    let preview = ctx.jgPreviewRows || [];
    const syncJgPreviewRowsFromTextarea = ctx.syncJgPreviewRowsFromTextarea;
    const norm = ctx.normStr;
    const jgM365DisplayName = ctx.jgM365DisplayName;

    if (!preview.length) syncJgPreviewRowsFromTextarea();

    const ownerByKlasse = new Map(jgRows.map((r) => [norm(r.klasse), norm(r.owner || '')]));
    const memByKlasse = new Map(jgRows.map((r) => [norm(r.klasse), String(r.memberLines ?? '')]));
    const nameByKlasse = new Map(jgRows.map((r) => [norm(r.klasse), norm(r.displayName || '')]));

    const dataRows = preview.map((r) => {
        const klasse = norm(r.klasse);
        const y = norm(r.jahr || '');
        const year = /^\d{4}$/.test(y) ? y : '';
        const className = norm(r.displayName || nameByKlasse.get(klasse) || '');
        const displayName = className || jgM365DisplayName(r);
        const mailNick = norm(r.mailNick);
        const email = mailNick ? `${mailNick}@${domain}` : '';
        const owner = ownerByKlasse.get(klasse) || '';
        const members = (memByKlasse.get(klasse) || '').replace(/\r\n|\n|\r/g, '|');
        return [klasse, year, className, displayName, mailNick, email, owner, members];
    });

    const csv = toCsv(
        [
            ['klasse', 'abschlussjahr', 'className', 'displayName', 'mailNick', 'groupEmail', 'ownerUpn', 'membersUpnPipe'],
            ...dataRows
        ],
        ';'
    );
    const filename = `jahrgang-export-${new Date().toISOString().slice(0, 10)}.csv`;
    return { csv, filename };
}

/**
 * @param {string} text
 * @param {{ normStr: (v: unknown) => string }} deps
 */
export function parseJgCsvToRows(text, deps) {
    const norm = deps.normStr;
    const raw = String(text || '');
    const delimiter = detectCsvDelimiter(raw);
    const lines = raw.split(/\r\n|\n|\r/).filter((l) => norm(l));
    if (!lines.length) return null;

    const first = parseCsvLine(lines[0], delimiter).map((x) => normHeaderKey(x));
    const hasHeader = first.some((h) =>
        ['klasse', 'abschlussjahr', 'classname', 'displayname', 'mailnick', 'ownerupn', 'membersupnpipe'].includes(h)
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
        const klasse = get(row, 'klasse', 0);
        const year = get(row, 'abschlussjahr', 1) || get(row, 'jahr', 1);
        const className = get(row, 'classname', 2);
        const displayName = get(row, 'displayname', 3);
        const owner = get(row, 'ownerupn', 6);
        const membersPipe = get(row, 'membersupnpipe', 7) || get(row, 'members', 7);
        const memberLines = membersPipe ? membersPipe.split('|').map((x) => norm(x)).filter(Boolean).join('\n') : '';
        if (!klasse) continue;
        outRows.push({ klasse, jahr: year, displayName: className || displayName, owner, memberLines });
    }
    return outRows;
}
