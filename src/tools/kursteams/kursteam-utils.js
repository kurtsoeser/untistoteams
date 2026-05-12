
// Gemeinsamer Namespace für den Kursteam-Modus (kein ES-Module/Bundler nötig).
const ns = (window.ms365Kursteam = window.ms365Kursteam || {});

ns.STORAGE_KEY = 'webuntis-teams-creator-state-v1';
ns.INVALID_CHARS_REPLACE = /[\\%&*+\/=?{}|<>();:,\[\]"öäü]/g;
ns.INVALID_CHARS_TEST = /[\\%&*+\/=?{}|<>();:,\[\]"öäü]/;

ns.escapeHtml = function escapeHtml(text) {
    const d = document.createElement('div');
    d.textContent = text;
    return d.innerHTML;
};

ns.attrEscape = function attrEscape(text) {
    return String(text ?? '').replace(/&/g, '&amp;').replace(/"/g, '&quot;');
};

ns.csvEscapeField = function csvEscapeField(value) {
    const s = String(value ?? '');
    if (/[",\r\n]/.test(s)) {
        return '"' + s.replace(/"/g, '""') + '"';
    }
    return s;
};

ns.buildCsvRow = function buildCsvRow(cols) {
    return cols.map(ns.csvEscapeField).join(',') + '\r\n';
};

ns.psEscapeSingle = function psEscapeSingle(s) {
    return String(s ?? '').replace(/'/g, "''");
};

ns.downloadBlob = function downloadBlob(filename, text, mime) {
    const blob = new Blob([text], { type: mime || 'text/plain;charset=utf-8' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
    URL.revokeObjectURL(a.href);
};

ns.sanitizeGruppeForMail = function sanitizeGruppeForMail(g) {
    if (!g || !String(g).trim()) return '';
    let s = String(g).replace(/[_\s]+/g, '-').replace(/-+/g, '-');
    s = s.replace(/^[^A-Za-z0-9]+|[^A-Za-z0-9]+$/g, '');
    return s;
};

ns.buildGruppenmailBase = function buildGruppenmailBase(yearPrefix, klasseForName, fach, gruppe) {
    const km = String(klasseForName).replace(/\s+/g, '-');
    const fm = String(fach).replace(/\s+/g, '-');
    let base = `${yearPrefix}-${km}-${fm}`;
    const gs = gruppe ? ns.sanitizeGruppeForMail(gruppe) : '';
    if (gs) base += '-' + gs;
    return base.replace(/\s+/g, '-');
};

ns.resolveDuplicateGruppenmails = function resolveDuplicateGruppenmails(teams) {
    const seen = new Map();
    let adjusted = 0;
    teams.forEach(team => {
        const base = team.gruppenmail;
        let candidate = base;
        let n = 2;
        while (seen.has(candidate)) {
            candidate = base + '-' + n;
            n++;
        }
        if (candidate !== base) {
            team.gruppenmail = candidate;
            team.mailNicknameAdjusted = true;
            adjusted++;
        } else {
            team.mailNicknameAdjusted = false;
        }
        seen.set(candidate, true);
    });
    return adjusted;
};

ns.combineClassNames = function combineClassNames(classString) {
    const classes = classString.split(',').map(c => c.trim());
    if (classes.length === 0) return classString;
    const firstClass = classes[0];
    const jahrgang = firstClass.match(/^\d+/);
    if (!jahrgang) return classString;
    const buchstaben = classes
        .map(c => {
            const match = c.match(/\d+([A-Z]+)/i);
            return match ? match[1].toUpperCase() : '';
        })
        .filter(b => b.length > 0);
    const uniqueBuchstaben = [...new Set(buchstaben.join('').split(''))].join('');
    return jahrgang[0] + uniqueBuchstaben;
};

