
/**
 * Reine Logik für Teamnamen-Muster (Drag-and-drop-Builder / Generierung).
 * Kein DOM – nur Datenstrukturen und Zusammensetzung.
 */
function defaultTeamNamePattern() {
    return [
        { type: 'yearPrefix' },
        { type: 'text', value: ' | ' },
        { type: 'klasse' },
        { type: 'text', value: ' | ' },
        { type: 'fach' }
    ];
}

function normalizePattern(pattern) {
    const arr = Array.isArray(pattern) ? pattern : [];
    const out = [];
    arr.forEach((p) => {
        if (!p || typeof p !== 'object') return;
        const type = String(p.type || '').trim();
        if (!type) return;
        if (type === 'text') {
            out.push({ type: 'text', value: String(p.value ?? '') });
        } else if (type === 'yearPrefix' || type === 'klasse' || type === 'fach' || type === 'gruppe') {
            out.push({ type });
        }
    });
    return out.length ? out : defaultTeamNamePattern();
}

function buildTeamNameFromPattern(pattern, ctx) {
    const parts = [];
    normalizePattern(pattern).forEach((p) => {
        if (p.type === 'text') parts.push(String(p.value ?? ''));
        else if (p.type === 'yearPrefix') parts.push(String(ctx.yearPrefix ?? ''));
        else if (p.type === 'klasse') parts.push(String(ctx.klasse ?? ''));
        else if (p.type === 'fach') parts.push(String(ctx.fach ?? ''));
        else if (p.type === 'gruppe') parts.push(String(ctx.gruppe ?? ''));
    });
    return parts.join('');
}

function tokenLabel(t) {
    if (t.type === 'yearPrefix') return 'Schuljahr';
    if (t.type === 'klasse') return 'Klasse';
    if (t.type === 'fach') return 'Fach';
    if (t.type === 'gruppe') return 'Gruppe';
    if (t.type === 'text') return `Text`;
    return t.type;
}

window.ms365KursteamTeamNames = {
    defaultTeamNamePattern,
    normalizePattern,
    buildTeamNameFromPattern,
    tokenLabel
};
