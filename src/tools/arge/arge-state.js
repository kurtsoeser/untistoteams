/**
 * @param {{
 *   argeCurrentStep: number,
 *   argeDefaultPrefix: string,
 *   argeUpperNick: boolean,
 *   argeCreateTeams: boolean,
 *   argeExchangeSmtp: boolean,
 *   argeAdminAsOwner: boolean,
 *   argeLines: string,
 *   rows: Array,
 *   normStr: (v: unknown) => string,
 *   subjectForSlug: (line: string) => string
 * }} ctx
 */
export function buildStateSnapshot(ctx) {
    const norm = ctx.normStr;
    const subj = ctx.subjectForSlug;
    return {
        kind: 'ms365-arge-export-v1',
        exportedAt: new Date().toISOString(),
        argeCurrentStep: ctx.argeCurrentStep,
        argeDefaultPrefix: ctx.argeDefaultPrefix,
        argeUpperNick: ctx.argeUpperNick,
        argeCreateTeams: ctx.argeCreateTeams,
        argeExchangeSmtp: ctx.argeExchangeSmtp,
        argeAdminAsOwner: ctx.argeAdminAsOwner,
        argeLines: ctx.argeLines,
        rows: (ctx.rows || []).map((r) => ({
            displayName: norm(r.displayName),
            subjectCode: norm(r.subjectCode || ''),
            subjectName: norm(r.subjectName || subj(r.displayName)),
            mailNick: norm(r.mailNick),
            owner: norm(r.owner || ''),
            memberLines: String(r.memberLines ?? ''),
            description: norm(r.description || '')
        }))
    };
}

/**
 * @param {object} obj
 * @param {{
 *   normStr: (v: unknown) => string,
 *   subjectForSlug: (line: string) => string,
 *   looksLikeSubjectCode: (s: string) => boolean,
 *   normSubjectCode: (s: string) => string,
 *   displayNameFromSubjectLine: (line: string) => string,
 *   syncArgePreviewFromTextarea: () => void,
 *   setArgeRows: (rows: Array) => void,
 *   rebuildArgeOwnerTableFromRows: () => void,
 *   rebuildArgeMembersTableFromRows: () => void,
 *   scheduleArgePreviewRowsOnly: () => void,
 *   refreshArgeScriptIfStep5: () => void,
 *   showToast: (msg: string) => void,
 *   getArgePreviewRows: () => Array
 * }} ctx
 */
export function applyImportedState(obj, ctx) {
    const o = obj && typeof obj === 'object' ? obj : {};
    const rows = Array.isArray(o.argeRows) ? o.argeRows : Array.isArray(o.rows) ? o.rows : [];
    const lines = o.argeLines !== undefined ? String(o.argeLines || '') : '';

    const norm = ctx.normStr;
    const subjectForSlug = ctx.subjectForSlug;
    const looksLikeSubjectCode = ctx.looksLikeSubjectCode;
    const normSubjectCode = ctx.normSubjectCode;
    const displayNameFromSubjectLine = ctx.displayNameFromSubjectLine;

    const prefixEl = document.getElementById('argeDefaultPrefix');
    const upperEl = document.getElementById('argeUpperNick');
    if (prefixEl && o.argeDefaultPrefix !== undefined) prefixEl.value = String(o.argeDefaultPrefix || '');
    if (upperEl && o.argeUpperNick !== undefined) upperEl.checked = !!o.argeUpperNick;
    const teamsEl = document.getElementById('argeCreateTeams');
    if (teamsEl && o.argeCreateTeams !== undefined) teamsEl.checked = !!o.argeCreateTeams;
    const exoEl = document.getElementById('argeExchangeSmtp');
    if (exoEl && o.argeExchangeSmtp !== undefined) exoEl.checked = !!o.argeExchangeSmtp;
    const adminEl = document.getElementById('argeAdminAsOwner');
    if (adminEl && o.argeAdminAsOwner !== undefined) adminEl.checked = !!o.argeAdminAsOwner;

    const ta = document.getElementById('argeLines');
    if (ta) ta.value = lines;

    if (ta && !norm(ta.value) && rows.length) {
        const reconstructed = rows
            .map((r) => {
                const code = norm(r.subjectCode || r.code);
                const name = norm(r.subjectName || r.name || subjectForSlug(r.displayName));
                const nick = norm(r.mailNick || '');
                if (looksLikeSubjectCode(code) && name && nick) return `${normSubjectCode(code)};${name};${nick}`;
                if (looksLikeSubjectCode(code) && name) return `${normSubjectCode(code)};${name}`;
                const dn = norm(r.displayName || displayNameFromSubjectLine(name));
                if (dn && nick) return `${dn};${nick}`;
                return name || dn;
            })
            .filter(Boolean);
        ta.value = reconstructed.join('\n');
    }

    const ownerByKey = new Map();
    const memByKey = new Map();
    const codeByKey = new Map();
    const subjByKey = new Map();
    rows.forEach((r) => {
        const dn = norm(r.displayName || displayNameFromSubjectLine(r.subjectName || r.name || ''));
        if (!dn) return;
        const k = dn.toLowerCase();
        ownerByKey.set(k, norm(r.owner || ''));
        memByKey.set(k, String(r.memberLines ?? r.members ?? ''));
        codeByKey.set(k, norm(r.subjectCode || r.code || ''));
        subjByKey.set(k, norm(r.subjectName || r.name || subjectForSlug(dn)));
    });

    ctx.syncArgePreviewFromTextarea();
    const argePreviewRows = ctx.getArgePreviewRows();
    const nextRows = argePreviewRows.map((r) => {
        const k = norm(r.displayName).toLowerCase();
        return {
            displayName: norm(r.displayName),
            mailNick: norm(r.mailNick),
            owner: ownerByKey.get(k) || '',
            memberLines: memByKey.get(k) || '',
            description: 'ARGE-Gruppe: ' + norm(r.displayName),
            subjectName: norm(r.subjectName || subjByKey.get(k) || subjectForSlug(r.displayName)),
            subjectCode: norm(r.subjectCode || codeByKey.get(k) || '')
        };
    });
    ctx.setArgeRows(nextRows);

    ctx.rebuildArgeOwnerTableFromRows();
    ctx.rebuildArgeMembersTableFromRows();
    ctx.scheduleArgePreviewRowsOnly();
    ctx.refreshArgeScriptIfStep5();
    ctx.showToast('ARGE: Import übernommen.');
}
