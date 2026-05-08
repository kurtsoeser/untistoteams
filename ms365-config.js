/**
 * Tragen Sie unten Ihre Anwendungs-ID (Client) aus der Entra-App-Registrierung ein.
 * Ausführliche Schritte: siehe ms365-config.example.js (Kommentarblock oben).
 */
window.MS365_MSAL_CONFIG = {
    clientId: 'e1d877c3-004c-4040-8c3b-81a59e0c7050',
    authority: 'https://login.microsoftonline.com/organizations',
    redirectUri: (function () {
        if (typeof window === 'undefined') return '';
        try {
            const origin = window.location.origin;
            const host = (window.location.hostname || '').toLowerCase();
            const isLocal =
                host === 'localhost' ||
                host === '127.0.0.1' ||
                host === '::1' ||
                host.endsWith('.localhost');

            function basePathForThisHost() {
                // Ziel: bei GitHub Pages Project Pages (…/repo/…) automatisch den Repo-Pfad mitnehmen.
                // Beispiele:
                // - /ms365-schultools/tools/schulstruktur-sync.html  -> /ms365-schultools
                // - /ms365-schultools/index.html                   -> /ms365-schultools
                // - /tools/arge.html                               -> (root)
                const p = String(window.location.pathname || '/');
                const noQuery = p.split('?')[0].split('#')[0];
                // Wenn wir in /tools/… sind, ist alles davor die "Basis"
                const iTools = noQuery.toLowerCase().indexOf('/tools/');
                if (iTools !== -1) {
                    const base = noQuery.slice(0, iTools);
                    return base.endsWith('/') ? base.slice(0, -1) : base;
                }
                // Sonst: Ordner der aktuellen Datei; bei /index.html oder /ms365-schooltool.html ist das bereits die Basis
                const lastSlash = noQuery.lastIndexOf('/');
                if (lastSlash <= 0) return '';
                const base = noQuery.slice(0, lastSlash);
                return base.endsWith('/') ? base.slice(0, -1) : base;
            }

            const base = isLocal ? '' : basePathForThisHost();
            // Immer stabile Redirect-Seite verwenden (keine Tool-Unterseite),
            // damit Entra nur 1 Redirect-URI pro Umgebung braucht.
            return origin + (base ? base : '') + '/ms365-schooltool.html';
        } catch {
            return window.location.href.split('#')[0];
        }
    })()
};

(function () {
    if (typeof document === 'undefined') return;
    function injectSiteCredit() {
        if (document.getElementById('siteCreditKurtrocks')) return;
        const p = document.createElement('p');
        p.id = 'siteCreditKurtrocks';
        p.className = 'site-credit-row';
        const a = document.createElement('a');
        a.className = 'site-credit-link';
        a.href = 'https://www.kurtrocks.com/';
        a.target = '_blank';
        a.rel = 'noopener noreferrer';
        const icon = document.createElement('i');
        icon.className = 'bi bi-info-circle';
        icon.setAttribute('aria-hidden', 'true');
        a.appendChild(icon);
        a.appendChild(document.createTextNode('kurtrocks.com'));
        p.appendChild(a);
        document.body.appendChild(p);
    }
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', injectSiteCredit);
    else injectSiteCredit();
})();
