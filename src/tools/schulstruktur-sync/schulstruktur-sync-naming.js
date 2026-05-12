/**
 * Naming-/Mail-Nick-/Schuljahr-Helfer für „Schulstruktur-Sync".
 *
 * Aus `schulstruktur-sync.js` 1:1 ausgelagert. Alle Funktionen sind **rein**
 * (keine DOM-/Window-/Storage-Zugriffe) und damit ohne Browser-Mock testbar.
 * Lediglich `currentSchoolYearLabel` und `generateGraphTempPassword` nutzen
 * `new Date()` bzw. `Math.random()` – das ist für Tests deterministisch genug
 * (Längen, Zeichenklassen, Format prüfbar).
 *
 * Schichten:
 *  - Schuljahr-Helfer:    parseSchoolYearStartYear, currentSchoolYearLabel,
 *                         nextSchoolYearLabel, gradeFromGraduationYear,
 *                         replaceLeadingNumber
 *  - Nick-Normalisierung: normNickPart, normNickPrefixLower, maybeUpperByFlag
 *  - Builder:             buildKursteamNameFromTemplate,
 *                         buildKursteamMailNickFromTemplate, buildJgMailNick,
 *                         buildArgeMailNick, buildMailNickFromLabel,
 *                         mailNicknameFromUpn
 *  - Misc:                generateGraphTempPassword
 */

// ────────── Schuljahr-Helfer ──────────

/**
 * Extrahiert das Anfangsjahr eines Labels wie `"2025/26"` oder `"2025/2026"`.
 * @param {string} label
 * @returns {number} Jahr oder `NaN`, wenn das Label nicht passt.
 */
export function parseSchoolYearStartYear(label) {
    const m = String(label || '').trim().match(/^(\d{4})\s*\/\s*(\d{2}|\d{4})/);
    return m ? parseInt(m[1], 10) : NaN;
}

/** Liefert das Schuljahr-Label des aktuellen Datums (`"YYYY/YY"`). */
export function currentSchoolYearLabel() {
    const y = new Date().getFullYear();
    return String(y) + '/' + String(y + 1).slice(2);
}

/**
 * Nächstes Schuljahr zu `cur`. Bei ungültigem `cur` wird das aktuelle
 * Schuljahr gemäß `currentSchoolYearLabel` geliefert.
 * @param {string} cur
 */
export function nextSchoolYearLabel(cur) {
    const y = parseSchoolYearStartYear(cur);
    if (!isFinite(y)) return currentSchoolYearLabel();
    return String(y + 1) + '/' + String(y + 2).slice(2);
}

/**
 * Berechnet die aktuelle Schulstufe aus Abschlussjahr und Schuljahr.
 *
 * @param {string|number} gradYear        4-stelliges Abschlussjahr.
 * @param {string} schoolYearLabel        z. B. `"2025/26"`.
 * @param {number} maxStufen              Anzahl Schulstufen (1..12, Default 5).
 * @returns {number} Schulstufe oder `NaN` bei ungültigen Eingaben.
 */
export function gradeFromGraduationYear(gradYear, schoolYearLabel, maxStufen) {
    const gy = String(gradYear || '').trim();
    const sy = parseSchoolYearStartYear(schoolYearLabel);
    const gyi = /^\d{4}$/.test(gy) ? parseInt(gy, 10) : NaN;
    if (!isFinite(gyi) || !isFinite(sy)) return NaN;
    const ms = isFinite(maxStufen) ? Math.max(1, Math.min(12, Math.round(maxStufen))) : 5;
    /*
     * Abschlussjahr = Ende der höchsten Stufe. In Schuljahr sy/sy+1 gilt:
     *   grade = (maxStufen + 1) - (abschlussjahr - sy)
     */
    return (ms + 1) - (gyi - sy);
}

/**
 * Ersetzt die führende 1–2-stellige Zahl in `label` durch `nextGrade`
 * (z. B. `"1A"` → `"2A"`). Labels ohne führende Zahl bleiben unverändert.
 *
 * @param {string} label
 * @param {number} nextGrade
 */
export function replaceLeadingNumber(label, nextGrade) {
    const s = String(label || '').trim();
    if (!s) return s;
    const g = String(Math.round(nextGrade));
    if (/^\d{1,2}/.test(s)) return s.replace(/^\d{1,2}/, g);
    return s;
}

// ────────── Nick-Normalisierung ──────────

/** Erlaubt nur `[A-Za-z0-9-]`, behält Groß-/Kleinschreibung. */
export function normNickPart(s) {
    return String(s || '').trim().replace(/[^A-Za-z0-9-]/g, '');
}

/**
 * Normalisiert ein Präfix als lower-case, `[a-z0-9]`-only.
 * Liefert `fallback` (selbst normalisiert), wenn `s` leer wird.
 */
export function normNickPrefixLower(s, fallback) {
    const t = String(s || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
    return t || String(fallback || '').trim() || '';
}

/** Liefert `s` in UPPER- oder lower-case, je nach `upper`-Flag. */
export function maybeUpperByFlag(s, upper) {
    const v = String(s || '');
    return upper ? v.toUpperCase() : v.toLowerCase();
}

// ────────── Builder ──────────

/**
 * Setzt `{yearPrefix} | {klasse} | {fach}`-Platzhalter im Template ein.
 * Default-Template wird verwendet, wenn `tpl` leer ist.
 */
export function buildKursteamNameFromTemplate(tpl, ctx) {
    const template = String(tpl || '').trim() || '{yearPrefix} | {klasse} | {fach}';
    return template
        .replaceAll('{yearPrefix}', String(ctx.yearPrefix || ''))
        .replaceAll('{klasse}', String(ctx.klasse || ''))
        .replaceAll('{fach}', String(ctx.fach || ''))
        .replaceAll('{gruppe}', String(ctx.gruppe || ''));
}

/**
 * Wie {@link buildKursteamNameFromTemplate}, aber liefert einen
 * Mail-Nick-tauglichen Slug (`a-z0-9-`).
 */
export function buildKursteamMailNickFromTemplate(tpl, ctx) {
    const template = String(tpl || '').trim() || 'kt-{yearPrefix}-{klasse}-{fach}';
    const raw = template
        .replaceAll('{yearPrefix}', String(ctx.yearPrefix || ''))
        .replaceAll('{klasse}', String(ctx.klasse || ''))
        .replaceAll('{fach}', String(ctx.fach || ''))
        .replaceAll('{gruppe}', String(ctx.gruppe || ''));
    return String(raw || '')
        .trim()
        .toLowerCase()
        .replace(/\s+/g, '-')
        .replace(/[^a-z0-9-]/g, '')
        .replace(/-+/g, '-')
        .replace(/^-|-$/g, '');
}

/**
 * Mail-Nick für Jahrgang. Schema: `<prefix><year>[-<suffix>]`,
 * z. B. `jg2025-A`.
 */
export function buildJgMailNick(schema, year, suffix) {
    const prefix = normNickPrefixLower(schema?.jgPrefix, 'jg');
    const suf = maybeUpperByFlag(normNickPart(suffix), !!schema?.jgUpper);
    const y = String(year || '').trim().replace(/[^0-9]/g, '').slice(0, 4);
    const sep = suf ? '-' : '';
    return (prefix + y + sep + suf).replace(/[^A-Za-z0-9-]/g, '');
}

/**
 * Mail-Nick für Arbeitsgemeinschaft. Schema: `<prefix>[-<shortCode>]`,
 * z. B. `arge-FUSS`.
 */
export function buildArgeMailNick(schema, shortCode) {
    const prefix = normNickPrefixLower(schema?.argePrefix, 'arge');
    const code = maybeUpperByFlag(normNickPart(shortCode), !!schema?.argeUpper);
    const sep = code ? '-' : '';
    return (prefix + sep + code).replace(/[^A-Za-z0-9-]/g, '');
}

/**
 * Generischer Slug aus einem Bezeichner: lower, `[a-z0-9-]` only, Bindestriche
 * kollabieren, keine Lead-/Trail-Bindestriche.
 */
export function buildMailNickFromLabel(label) {
    return String(label || '')
        .trim()
        .toLowerCase()
        .replace(/\s+/g, '-')
        .replace(/[^a-z0-9-]/g, '')
        .replace(/-+/g, '-')
        .replace(/^-|-$/g, '');
}

/**
 * Leitet einen Mail-Nick aus dem Local-Part eines UPN/Mail-Adresse ab.
 * Fallback: zufälliger Wert `u<hex>`. Max. 64 Zeichen.
 */
export function mailNicknameFromUpn(upn) {
    const u = String(upn || '').trim().toLowerCase();
    const local = (u.split('@')[0] || '').trim();
    let nick = buildMailNickFromLabel(local);
    if (!nick) nick = 'u' + String(Math.random()).toString(16).slice(2, 10);
    if (nick.length > 64) nick = nick.slice(0, 64);
    return nick;
}

// ────────── Misc ──────────

/**
 * Generiert ein temporäres Graph-API-Passwort:
 * 14 alphanumerische Zeichen + 1 Sonderzeichen + Garantie-Suffix `1aA`
 * (erfüllt Standard-Komplexitätsregeln).
 */
export function generateGraphTempPassword() {
    const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789';
    const sym = '!@#$%';
    let s = '';
    for (let i = 0; i < 14; i++) s += chars.charAt(Math.floor(Math.random() * chars.length));
    return s + sym.charAt(Math.floor(Math.random() * sym.length)) + '1aA';
}
