export function buildWtgStateSnapshot(ctx) {
    const norm = ctx.normStr;
    return {
        kind: 'ms365-wtg-export-v1',
        exportedAt: new Date().toISOString(),
        wtgCurrentStep: ctx.wtgCurrentStep,
        wtgDefaultPrefix: ctx.wtgDefaultPrefix,
        wtgUpperNick: ctx.wtgUpperNick,
        wtgAdminAsOwner: ctx.wtgAdminAsOwner,
        wtgDefaultVisibilityPrivate: ctx.wtgDefaultVisibilityPrivate,
        wtgLines: ctx.wtgLines,
        rows: (ctx.rows || []).map((r) => ({
            displayName: norm(r.displayName),
            mailNick: norm(r.mailNick),
            mailNickExplicit: !!r.mailNickExplicit,
            kind: r.kind === 'team' ? 'team' : 'group',
            owner: norm(r.owner || ''),
            memberLines: String(r.memberLines ?? '')
        }))
    };
}

export function applyWtgImportedState(obj, ctx) {
    const o = obj && typeof obj === 'object' ? obj : {};
    const rows = Array.isArray(o.rows) ? o.rows : Array.isArray(o.wtgRows) ? o.wtgRows : [];
    const lines = o.wtgLines !== undefined ? String(o.wtgLines || '') : '';

    const prefixEl = document.getElementById('wtgDefaultPrefix');
    const upperEl = document.getElementById('wtgUpperNick');
    if (prefixEl && o.wtgDefaultPrefix !== undefined) prefixEl.value = String(o.wtgDefaultPrefix || '');
    if (upperEl && o.wtgUpperNick !== undefined) upperEl.checked = !!o.wtgUpperNick;
    const adminEl = document.getElementById('wtgAdminAsOwner');
    if (adminEl && o.wtgAdminAsOwner !== undefined) adminEl.checked = !!o.wtgAdminAsOwner;
    const visEl = document.getElementById('wtgDefaultVisibilityPrivate');
    if (visEl && o.wtgDefaultVisibilityPrivate !== undefined) visEl.checked = !!o.wtgDefaultVisibilityPrivate;

    const ta = document.getElementById('wtgLines');
    if (ta) ta.value = lines;

    const ownerByKey = new Map();
    const memByKey = new Map();
    const kindByKey = new Map();
    rows.forEach((r) => {
        const dn = ctx.normStr(r.displayName || '');
        if (!dn) return;
        const k = dn.toLowerCase();
        ownerByKey.set(k, ctx.normStr(r.owner || ''));
        memByKey.set(k, String(r.memberLines ?? r.members ?? ''));
        kindByKey.set(k, r.kind === 'team' ? 'team' : 'group');
    });

    ctx.syncWtgPreviewFromTextarea();
    const previewRows = ctx.getWtgPreviewRows();
    const nextRows = previewRows.map((r) => {
        const k = ctx.normStr(r.displayName).toLowerCase();
        return {
            displayName: ctx.normStr(r.displayName),
            mailNick: ctx.normStr(r.mailNick),
            mailNickExplicit: !!r.mailNickExplicit,
            kind: kindByKey.get(k) || (r.kind === 'team' ? 'team' : 'group'),
            owner: ownerByKey.get(k) || '',
            memberLines: memByKey.get(k) || ''
        };
    });
    ctx.setWtgRows(nextRows);
    ctx.rebuildWtgOwnerTableFromRows();
    ctx.rebuildWtgMembersTableFromRows();
    ctx.scheduleWtgPreviewRowsOnly();
    ctx.showToast('Weitere Teams & Gruppen: Import übernommen.');
}

