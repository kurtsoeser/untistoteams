/**
 * @param {{
 *   jgCurrentStep: number,
 *   jgPrefix: string,
 *   jgDefaultYear: string,
 *   jgSuffixUpper: boolean,
 *   jgCreateTeams: boolean,
 *   jgExchangeSmtp: boolean,
 *   jgClassLines: string,
 *   rows: Array,
 *   normStr: (v: unknown) => string
 * }} ctx
 */
export function buildJgStateSnapshot(ctx) {
    const norm = ctx.normStr;
    return {
        kind: 'ms365-jahrgang-export-v1',
        exportedAt: new Date().toISOString(),
        jgCurrentStep: ctx.jgCurrentStep,
        jgPrefix: ctx.jgPrefix,
        jgDefaultYear: ctx.jgDefaultYear,
        jgSuffixUpper: ctx.jgSuffixUpper,
        jgCreateTeams: ctx.jgCreateTeams,
        jgExchangeSmtp: ctx.jgExchangeSmtp,
        jgClassLines: ctx.jgClassLines,
        rows: (ctx.rows || []).map((r) => ({
            klasse: norm(r.klasse),
            jahr: norm(r.jahr),
            displayName: norm(r.displayName || ''),
            suffix: norm(r.suffix || ''),
            mailNick: norm(r.mailNick || ''),
            owner: norm(r.owner || ''),
            memberLines: String(r.memberLines ?? '')
        }))
    };
}

/**
 * @param {object} obj
 * @param {{
 *   normStr: (v: unknown) => string,
 *   syncJgPreviewRowsFromTextarea: () => void,
 *   getPrefix: () => string,
 *   getJgDefaultAbschlussjahr: () => string,
 *   buildMailNickname: (prefix: string, year: string, suffix: string) => string,
 *   resolveDuplicateNicks: (rows: Array) => void,
 *   setJgRows: (rows: Array) => void,
 *   getJgPreviewRows: () => Array,
 *   rebuildJgOwnerTableFromRows: () => void,
 *   rebuildJgMembersTableFromRows: () => void,
 *   scheduleJgPreviewRowsOnly: () => void,
 *   refreshJgScriptIfStep5: () => void,
 *   updatePrefixExample: () => void,
 *   showToast: (msg: string) => void
 * }} ctx
 */
export function applyJgImportedState(obj, ctx) {
    const o = obj && typeof obj === 'object' ? obj : {};
    const rows = Array.isArray(o.jgRows) ? o.jgRows : Array.isArray(o.rows) ? o.rows : [];
    const lines = o.jgClassLines !== undefined ? String(o.jgClassLines || '') : '';

    const norm = ctx.normStr;

    const prefixEl = document.getElementById('jgPrefix');
    const defYearEl = document.getElementById('jgDefaultYear');
    const upperEl = document.getElementById('jgSuffixUpper');
    if (prefixEl && o.jgPrefix !== undefined) prefixEl.value = String(o.jgPrefix || 'jg');
    if (defYearEl && o.jgDefaultYear !== undefined) defYearEl.value = String(o.jgDefaultYear || '2030');
    if (upperEl && o.jgSuffixUpper !== undefined) upperEl.checked = !!o.jgSuffixUpper;
    const teamsEl = document.getElementById('jgCreateTeams');
    if (teamsEl && o.jgCreateTeams !== undefined) teamsEl.checked = !!o.jgCreateTeams;
    const exoEl = document.getElementById('jgExchangeSmtp');
    if (exoEl && o.jgExchangeSmtp !== undefined) exoEl.checked = !!o.jgExchangeSmtp;

    const ta = document.getElementById('jgClassLines');
    if (ta) ta.value = lines;
    if (ta && !norm(ta.value) && rows.length) {
        const reconstructed = rows
            .map((r) => {
                const k = norm(r.klasse || r.class || r.name);
                const y = norm(r.jahr || r.year);
                if (!k) return '';
                if (/^\d{4}$/.test(y)) return `${k};${y}`;
                return k;
            })
            .filter(Boolean);
        ta.value = reconstructed.join('\n');
    }

    const ownerByKlasse = new Map();
    const memByKlasse = new Map();
    rows.forEach((r) => {
        const k = norm(r.klasse || r.class || r.name);
        if (!k) return;
        ownerByKlasse.set(k, norm(r.owner || r.ownerUpn || ''));
        memByKlasse.set(k, String(r.memberLines ?? r.members ?? ''));
    });

    const nameByKlasse = new Map();
    rows.forEach((r) => {
        const k = norm(r.klasse || r.class || r.code);
        if (!k) return;
        const dn = norm(r.displayName || r.className || r.klassenname || r.name || '');
        if (dn) nameByKlasse.set(k, dn);
    });

    ctx.syncJgPreviewRowsFromTextarea();
    const prefix = ctx.getPrefix();
    const jgPreviewRows = ctx.getJgPreviewRows();
    const nextRows = jgPreviewRows.map((r) => {
        const m = norm(r.klasse).match(/^(\d+)([A-Za-z]+)$/);
        const y = norm(r.jahr || '');
        const year = /^\d{4}$/.test(y) ? y : ctx.getJgDefaultAbschlussjahr();
        const suffix = m ? m[2] : norm(r.suffix || '');
        const klasseTrim = norm(r.klasse);
        return {
            klasse: klasseTrim,
            jahr: year,
            suffix,
            mailNick: ctx.buildMailNickname(prefix, year, suffix),
            displayName: norm(r.displayName || nameByKlasse.get(klasseTrim) || ''),
            owner: ownerByKlasse.get(klasseTrim) || '',
            memberLines: memByKlasse.get(klasseTrim) || ''
        };
    });
    ctx.resolveDuplicateNicks(nextRows);
    ctx.setJgRows(nextRows);

    ctx.rebuildJgOwnerTableFromRows();
    ctx.rebuildJgMembersTableFromRows();
    ctx.scheduleJgPreviewRowsOnly();
    ctx.refreshJgScriptIfStep5();
    ctx.updatePrefixExample();
    ctx.showToast('Jahrgang: Import übernommen.');
}
