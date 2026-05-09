/**
 * Dynamischer Import von @azure/msal-browser (Version 3.26.1), gleiche CDN-Reihenfolge wie zuvor
 * in den einzelnen Aufrufern: zuerst esm.sh, bei Fehler jsdelivr +esm.
 * @returns {Promise<typeof import('@azure/msal-browser')>}
 */
export async function loadMsalBrowser() {
    try {
        return await import('https://esm.sh/@azure/msal-browser@3.26.1');
    } catch {
        return await import('https://cdn.jsdelivr.net/npm/@azure/msal-browser@3.26.1/+esm');
    }
}
