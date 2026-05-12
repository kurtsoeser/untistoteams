import { normStr } from '../../shared/utils/strings.js';

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

function looksLikeNick(s) {
    const t = normStr(s);
    if (!t) return false;
    return /^[A-Za-z0-9-]{1,64}$/.test(t);
}

function maybeUpper(s, deps) {
    return deps && deps.maybeUpper ? deps.maybeUpper(s) : s;
}

function buildMailNickname(displayName, deps) {
    const base = toNickBaseFromName(displayName);
    const pre = deps && deps.getPrefix ? normStr(deps.getPrefix()) : '';
    const combined = pre ? pre + '-' + base : base;
    return maybeUpper(combined, deps).replace(/[^A-Za-z0-9-]/g, '').slice(0, 64);
}

function normTypeToken(t) {
    const s = normStr(t).toLowerCase();
    if (!s) return '';
    if (s === 'team' || s === 'teams') return 'team';
    if (s === 'gruppe' || s === 'group' || s === 'groups') return 'group';
    return '';
}

function resolveDuplicateNicks(rows) {
    const seen = new Map();
    rows.forEach((r) => {
        const base = r.mailNick;
        let candidate = base;
        let n = 2;
        while (seen.has(candidate)) {
            candidate = (base + '-' + n).slice(0, 64);
            n++;
        }
        r.mailNick = candidate;
        seen.set(candidate, true);
    });
}

/**
 * Eingabeformate (pro Zeile):
 * - "Projektteam Nachhaltigkeit"
 * - "Projektteam Nachhaltigkeit;projektteam-nachhaltigkeit"
 * - "Projektteam Nachhaltigkeit;projektteam-nachhaltigkeit;Team"
 * - "Projektteam Nachhaltigkeit;;Gruppe" (Nickname automatisch)
 */
function parseWtgLine(line, deps) {
    const parts = String(line || '')
        .split(/[;\t|]/)
        .map((x) => normStr(x))
        .filter((x, i, arr) => !(i === 1 && arr.length >= 2 && x === '')); // keep empty nick only if explicit ";;" not present; simplified
    const raw = normStr(line);
    if (!raw || raw.startsWith('#')) return null;

    // tolerate ";;" by splitting without filter:
    const rawParts = String(line || '').split(/[;\t|]/).map((x) => normStr(x));
    const displayName = normStr(rawParts[0]);
    if (!displayName) return null;
    const explicitNick = normStr(rawParts[1] ?? '');
    const typeTok = normTypeToken(rawParts[2] ?? '');
    const isTeam = typeTok === 'team';
    const isGroup = typeTok === 'group' || typeTok === '';

    let mailNickExplicit = false;
    let mailNick = '';
    if (explicitNick) {
        mailNick = maybeUpper(explicitNick.replace(/[^A-Za-z0-9-]/g, ''), deps).slice(0, 64);
        mailNickExplicit = true;
    } else {
        mailNick = buildMailNickname(displayName, deps);
    }
    if (!looksLikeNick(mailNick)) return null;

    return {
        displayName,
        mailNick,
        mailNickExplicit,
        kind: isTeam ? 'team' : 'group',
        owner: '',
        memberLines: ''
    };
}

function parseWtgInputLines(text, deps) {
    const lines = String(text || '').split(/\r\n|\n|\r/);
    const errors = [];
    const parsed = [];
    const seenName = new Set();
    lines.forEach((line, idx) => {
        const t = normStr(line);
        if (!t || t.startsWith('#')) return;
        const row = parseWtgLine(line, deps);
        if (!row) {
            errors.push('Zeile ' + (idx + 1) + ': Format nicht erkannt.');
            return;
        }
        const key = row.displayName.toLowerCase();
        if (seenName.has(key)) return;
        seenName.add(key);
        parsed.push(row);
    });
    resolveDuplicateNicks(parsed);
    return { parsed, errors };
}

function recomputeWtgPreviewMailNicks(rows, deps) {
    rows.forEach((r) => {
        r.displayName = normStr(r.displayName);
        if (!r.mailNickExplicit) {
            r.mailNick = buildMailNickname(r.displayName, deps);
        } else {
            r.mailNick = maybeUpper(String(r.mailNick || '').replace(/[^A-Za-z0-9-]/g, ''), deps).slice(0, 64);
        }
        if (r.kind !== 'team') r.kind = 'group';
    });
    resolveDuplicateNicks(rows);
}

function serializePreviewRowsToLines(rows) {
    return rows.map((r) => {
        const dn = normStr(r.displayName);
        const nick = normStr(r.mailNick);
        const kind = r.kind === 'team' ? 'Team' : 'Gruppe';
        if (r.mailNickExplicit && nick) return `${dn};${nick};${kind}`;
        return `${dn};;${kind}`;
    });
}

export {
normStr,
toNickBaseFromName,
looksLikeNick,
buildMailNickname,
parseWtgLine,
parseWtgInputLines,
recomputeWtgPreviewMailNicks,
serializePreviewRowsToLines
};

