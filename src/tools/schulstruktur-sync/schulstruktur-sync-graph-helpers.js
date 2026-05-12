/**
 * Pure Graph-API-Helfer für „Schulstruktur-Sync".
 *
 * Aus `schulstruktur-sync.js` 1:1 ausgelagert (Phase 2 Schnitt 9).
 *
 * Reine Funktionen ohne externe Abhängigkeiten – keine `window.*` Lookups,
 * kein DOM, keine Netzwerk-/MSAL-Aufrufe.
 *
 *  - {@link isInteractionRequired}: MSAL-Fehler-Klassifizierung.
 *  - {@link isGraphDuplicateRefError}: Graph-Fehlertext deutet auf
 *    „bereits vorhanden" (z. B. Member/Owner schon zugeordnet).
 *  - {@link graphErrorLooksLikeNotFound}: Graph-Fehlertext deutet auf
 *    `404 / ResourceNotFound`.
 *  - {@link parseTeamsOperationPathFromLocation}: extrahiert den Polling-Pfad
 *    `/teams/{id}/operations/{op}` aus einem `Location:` Header.
 *  - {@link groupIsTeam}: `g.resourceProvisioningOptions` enthält `'Team'`.
 *  - {@link personLabel}: einheitliches Anzeige-Label für Personen
 *    (DisplayName + UPN/Mail).
 *  - {@link odataEscape}: maskiert Hochkommas für OData-Filterklauseln.
 *  - {@link directoryObjectRef}: baut einen `@odata.id`-Verweis zusammen.
 *  - {@link sleep}: Async-Pause via `setTimeout`.
 */

/**
 * Klassifiziert einen MSAL-Fehler als „User-Interaction nötig".
 *
 * @param {{ name?: string, errorCode?: string, message?: string } | null | undefined} e
 * @returns {boolean}
 */
export function isInteractionRequired(e) {
    return !!(
        e &&
        (e.name === 'InteractionRequiredAuthError' ||
            e.errorCode === 'interaction_required' ||
            (typeof e.message === 'string' && e.message.indexOf('interaction_required') !== -1))
    );
}

/**
 * Async-Pause via Promise + setTimeout.
 *
 * @param {number} ms
 * @returns {Promise<void>}
 */
export function sleep(ms) {
    return new Promise((r) => setTimeout(r, ms));
}

/**
 * Extrahiert aus einem `Location:` (oder `Content-Location:`) Header den
 * relativen Pfad `/teams/{teamId}/operations/{opId}`, sodass damit per Graph
 * gepollt werden kann. Unterstützt sowohl Full-URLs als auch beide
 * Teams-Pfadvarianten (REST und OData).
 *
 * @param {string | null | undefined} locationHeader
 * @returns {string | null}
 */
export function parseTeamsOperationPathFromLocation(locationHeader) {
    if (!locationHeader) return null;
    let loc = String(locationHeader).trim();
    if (loc.indexOf('http') === 0) {
        try {
            const u = new URL(loc);
            loc = u.pathname.replace(/^\/v1\.0/i, '');
        } catch {
            return null;
        }
    }
    const m = loc.match(/\/teams\/([^/]+)\/operations\/([^/?\s]+)/i);
    if (m) return '/teams/' + m[1] + '/operations/' + m[2];
    const m2 = loc.match(/teams\('([^']+)'\)\/operations\('([^']+)'\)/i);
    if (m2) return '/teams/' + m2[1] + '/operations/' + m2[2];
    return null;
}

/**
 * Liefert `true`, wenn eine Unified-Group eine Teams-Provisionierung hat.
 *
 * @param {{ resourceProvisioningOptions?: string[] } | null | undefined} g
 * @returns {boolean}
 */
export function groupIsTeam(g) {
    const opts = g && g.resourceProvisioningOptions;
    return Array.isArray(opts) && opts.indexOf('Team') !== -1;
}

/**
 * Heuristik: Graph-Error-Message deutet auf „nicht gefunden" hin.
 *
 * @param {unknown} err
 * @returns {boolean}
 */
export function graphErrorLooksLikeNotFound(err) {
    const s = String((err && err.message) || err || '');
    return /\b404\b/i.test(s) || /ResourceNotFound|ItemNotFound|not found|Request_ResourceNotFound/i.test(s);
}

/**
 * Einheitliches Anzeige-Label für eine Person: `displayName (UPN)` oder
 * Fallback auf UPN/Mail/Id.
 *
 * @param {{ displayName?: string, userPrincipalName?: string, mail?: string, id?: string|number } | null | undefined} p
 * @returns {string}
 */
export function personLabel(p) {
    if (!p || typeof p !== 'object') return '';
    const dn = p.displayName ? String(p.displayName).trim() : '';
    const upn = p.userPrincipalName || p.mail ? String(p.userPrincipalName || p.mail).trim() : '';
    if (dn && upn && dn !== upn) return dn + ' (' + upn + ')';
    return dn || upn || (p.id ? String(p.id) : '');
}

/**
 * Maskiert Hochkommas für eingebettete Werte in OData-Filterklauseln (Graph).
 *
 * @param {unknown} s
 * @returns {string}
 */
export function odataEscape(s) {
    return String(s).replace(/'/g, "''");
}

/**
 * Baut einen Graph-`@odata.id`-Verweis auf ein DirectoryObject zusammen.
 *
 * @param {string} id
 * @returns {string}
 */
export function directoryObjectRef(id) {
    return 'https://graph.microsoft.com/v1.0/directoryObjects/' + id;
}

/**
 * Heuristik: Graph-Error-Message deutet auf „Beziehung existiert bereits"
 * (Member/Owner-Add scheitert mit `One or more added object references already exist`).
 *
 * @param {unknown} err
 * @returns {boolean}
 */
export function isGraphDuplicateRefError(err) {
    const msg = String(err && err.message ? err.message : err);
    return /already exist/i.test(msg) || /already exists/i.test(msg);
}
