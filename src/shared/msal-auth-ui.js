(function () {
    'use strict';

    const MSAL_LOADER_IMPORT = (function () {
        const needle = 'msal-auth-ui.js';
        const rel = './msal-loader.js';
        const scripts = document.getElementsByTagName('script');
        for (let i = scripts.length - 1; i >= 0; i--) {
            const src = scripts[i].src || '';
            if (src.indexOf(needle) !== -1) {
                try {
                    return new URL(rel, src).href;
                } catch (_) {}
            }
        }
        try {
            return new URL('src/shared/msal-loader.js', document.baseURI).href;
        } catch (_) {
            return 'src/shared/msal-loader.js';
        }
    })();

    const DEFAULT_SCOPES = ['https://graph.microsoft.com/User.Read'];
    const POST_LOGIN_KEY = 'ms365-post-login-url';

    let msalMod = null;
    let pca = null;
    let initPromise = null;

    function $(sel, root) {
        return (root || document).querySelector(sel);
    }

    function resolveMsalConfig() {
        let cfg = window.MS365_MSAL_CONFIG;
        if (!cfg) cfg = {};
        let id = String(cfg.clientId || '').trim();
        if (!id) {
            const meta = document.querySelector('meta[name="ms365-graph-client-id"]');
            const fromMeta = meta && meta.getAttribute('content') ? meta.getAttribute('content').trim() : '';
            if (fromMeta) id = fromMeta;
        }
        if (!id) throw new Error('Keine clientId: ms365-config.js fehlt/leer oder blockiert.');
        return {
            clientId: id,
            authority: cfg.authority || 'https://login.microsoftonline.com/organizations',
            redirectUri: (cfg.redirectUri || window.location.href.split('#')[0]).trim()
        };
    }

    async function loadMsal() {
        if (msalMod) return msalMod;
        const loader = await import(/* @vite-ignore */ MSAL_LOADER_IMPORT);
        if (typeof loader.loadMsalBrowser !== 'function') {
            throw new Error('MSAL-Loader: loadMsalBrowser fehlt.');
        }
        msalMod = await loader.loadMsalBrowser();
        return msalMod;
    }

    function isInteractionRequired(e) {
        if (!e) return false;
        if (e.name === 'InteractionRequiredAuthError') return true;
        const code = String(e.errorCode || '').toLowerCase();
        if (
            code === 'interaction_required' ||
            code === 'consent_required' ||
            code === 'login_required' ||
            code === 'invalid_grant' ||
            code === 'no_account_in_silent_request' ||
            code === 'no_tokens_found' ||
            code === 'monitor_window_timeout' ||
            code === 'native_account_unavailable'
        ) {
            return true;
        }
        const msg = String((e && e.message) || '').toLowerCase();
        return (
            msg.indexOf('interaction_required') !== -1 ||
            msg.indexOf('consent_required') !== -1 ||
            msg.indexOf('login_required') !== -1 ||
            msg.indexOf('invalid_grant') !== -1 ||
            msg.indexOf('aadsts65001') !== -1 || // Consent fehlt
            msg.indexOf('aadsts50058') !== -1 || // Sitzung verloren
            msg.indexOf('aadsts70008') !== -1 || // Refresh-Token abgelaufen
            msg.indexOf('aadsts50173') !== -1 || // Refresh-Token widerrufen
            msg.indexOf('aadsts50076') !== -1 || // MFA nötig
            msg.indexOf('aadsts50079') !== -1 || // MFA registration nötig
            msg.indexOf('aadsts700084') !== -1 || // Cookie hash mismatch
            msg.indexOf('token contains an invalid signature') !== -1
        );
    }

    async function ensurePca() {
        if (initPromise) return initPromise;
        initPromise = (async () => {
            const m = await loadMsal();
            const PublicClientApplication = m.PublicClientApplication || (m.default && m.default.PublicClientApplication);
            if (!PublicClientApplication) throw new Error('MSAL: PublicClientApplication nicht gefunden.');
            const cfg = resolveMsalConfig();
            pca = new PublicClientApplication({
                auth: { clientId: cfg.clientId, authority: cfg.authority, redirectUri: cfg.redirectUri },
                // localStorage statt sessionStorage: ermöglicht Single-Sign-On zwischen Browser-Tabs
                // (Microsoft 365 Anmeldung wird übernommen, wenn der Benutzer bereits in einem
                // anderen Tab/Modul angemeldet ist).
                cache: { cacheLocation: 'localStorage', storeAuthStateInCookie: true }
            });
            await pca.initialize();
            await pca.handleRedirectPromise();

            const accounts = pca.getAllAccounts();
            if (accounts && accounts[0] && typeof pca.setActiveAccount === 'function') {
                pca.setActiveAccount(accounts[0]);
            }
            return pca;
        })();
        return initPromise;
    }

    /**
     * Versucht eine unsichtbare Single-Sign-On-Anmeldung über die bestehende
     * Microsoft-365-Browser-Sitzung (Hidden Iframe an login.microsoftonline.com).
     * Funktioniert, wenn der Benutzer in einem anderen Tab/Fenster bereits angemeldet ist
     * und Third-Party-Cookies für Microsoft erlaubt sind.
     * Wirft NICHT bei Fehlschlag (z. B. wenn kein Account vorhanden / Cookies blockiert).
     */
    async function trySsoSilent(scopes) {
        if (!pca) return null;
        try {
            const req = {
                scopes: Array.isArray(scopes) && scopes.length ? scopes : DEFAULT_SCOPES
            };
            const result = await pca.ssoSilent(req);
            if (result && result.account && typeof pca.setActiveAccount === 'function') {
                pca.setActiveAccount(result.account);
            }
            return result;
        } catch {
            return null;
        }
    }

    function getAccount() {
        if (!pca) return null;
        const a = typeof pca.getActiveAccount === 'function' ? pca.getActiveAccount() : null;
        if (a) return a;
        const all = pca.getAllAccounts();
        return all && all[0] ? all[0] : null;
    }

    function accountLabel(a) {
        if (!a) return '';
        const u = a.username ? String(a.username) : '';
        const n = a.name ? String(a.name) : '';
        if (n && u && n !== u) return n + ' (' + u + ')';
        return n || u || '';
    }

    /**
     * Anmeldung per Redirect.
     * - Ohne opts.prompt: Microsoft entscheidet selbst (nutzt bestehende Browser-Session,
     *   zeigt Account-Auswahl nur falls nötig). So funktioniert SSO mit anderen MS-365-Tabs.
     * - Mit opts.prompt === 'select_account': erzwingt Account-Auswahl (z. B. zum Konto wechseln).
     */
    async function login(scopes, opts) {
        const instance = await ensurePca();
        try {
            sessionStorage.setItem(POST_LOGIN_KEY, window.location.href);
        } catch {
            // ignore
        }
        const req = {
            scopes: Array.isArray(scopes) && scopes.length ? scopes : DEFAULT_SCOPES,
            redirectStartPage: window.location.href
        };
        if (opts && typeof opts.prompt === 'string' && opts.prompt) {
            req.prompt = opts.prompt;
        }
        await instance.loginRedirect(req);
        // redirect -> no further code
    }

    async function switchAccount(scopes) {
        return login(scopes, { prompt: 'select_account' });
    }

    async function logout() {
        const instance = await ensurePca();
        const a = getAccount();
        try {
            sessionStorage.setItem(POST_LOGIN_KEY, window.location.href);
        } catch {
            // ignore
        }
        await instance.logoutRedirect({ account: a || undefined, postLogoutRedirectUri: window.location.href.split('#')[0] });
    }

    function looksLikeBrokenCache(e) {
        if (!e) return false;
        const msg = String((e && e.message) || '').toLowerCase();
        return (
            msg.indexOf('token contains an invalid signature') !== -1 ||
            msg.indexOf('invalid_grant') !== -1 ||
            msg.indexOf('aadsts70008') !== -1 ||
            msg.indexOf('aadsts50173') !== -1 ||
            msg.indexOf('aadsts700084') !== -1
        );
    }

    async function clearMsalCache(instance) {
        try {
            const accounts = instance && typeof instance.getAllAccounts === 'function' ? instance.getAllAccounts() : [];
            if (typeof instance.clearCache === 'function') {
                try {
                    await instance.clearCache();
                } catch {
                    // ignore – wir versuchen es danach noch manuell
                }
            }
            (accounts || []).forEach((acc) => {
                if (acc && typeof instance.logoutSilent === 'function') {
                    instance.logoutSilent({ account: acc }).catch(() => {});
                }
            });
        } catch {
            // ignore
        }
        try {
            const removeIf = (store, predicate) => {
                const keys = [];
                for (let i = 0; i < store.length; i++) {
                    const k = store.key(i);
                    if (k && predicate(k)) keys.push(k);
                }
                keys.forEach((k) => {
                    try {
                        store.removeItem(k);
                    } catch {
                        // ignore
                    }
                });
            };
            const isMsalKey = (k) =>
                k.indexOf('msal.') === 0 ||
                k.indexOf('msal-') === 0 ||
                k.indexOf('login.microsoftonline.com') !== -1 ||
                k.indexOf('login.windows.net') !== -1 ||
                /[-.]?msal[-.]/i.test(k);
            removeIf(localStorage, isMsalKey);
            removeIf(sessionStorage, isMsalKey);
        } catch {
            // ignore
        }
    }

    async function acquireToken(scopes) {
        const instance = await ensurePca();
        let accounts = instance.getAllAccounts();
        if (!accounts.length) {
            await login(scopes);
            throw new Error('Weiterleitung zur Anmeldung …');
        }
        const a = getAccount() || accounts[0];
        const req = { scopes: Array.isArray(scopes) && scopes.length ? scopes : DEFAULT_SCOPES, account: a };
        try {
            return (await instance.acquireTokenSilent(req)).accessToken;
        } catch (e) {
            // Bei kaputtem/abgelaufenem MSAL-Cache („Token contains an invalid signature",
            // invalid_grant, AADSTS70008/50173/700084 etc.) den lokalen Cache leeren,
            // damit der frische Login-Redirect tatsächlich frische Tokens holt.
            if (looksLikeBrokenCache(e)) {
                try {
                    await clearMsalCache(instance);
                } catch {
                    // ignore
                }
            }
            if (isInteractionRequired(e) || looksLikeBrokenCache(e)) {
                try {
                    sessionStorage.setItem(POST_LOGIN_KEY, window.location.href);
                } catch {
                    // ignore
                }
                const redirectReq = { ...req, redirectStartPage: window.location.href };
                // Beim "Cache broken" zusätzlich Consent erzwingen, damit der Tenant
                // den User korrekt neu authentifiziert.
                if (looksLikeBrokenCache(e)) {
                    redirectReq.prompt = 'select_account';
                }
                await instance.acquireTokenRedirect(redirectReq);
                throw new Error('Weiterleitung zur Anmeldung …');
            }
            throw e;
        }
    }

    function ensureHeaderWidget() {
        const header = $('.header') || $('header');
        if (!header) return;
        if ($('#ms365AuthWidget', header)) return;

        // Prefer toolbar right side; otherwise append to header.
        const toolbar = $('.toolbar', header) || header;
        const wrap = document.createElement('div');
        wrap.id = 'ms365AuthWidget';
        wrap.style.display = 'flex';
        wrap.style.gap = '10px';
        wrap.style.alignItems = 'center';
        wrap.style.justifyContent = 'flex-end';
        wrap.style.flexWrap = 'wrap';
        wrap.style.marginLeft = 'auto';

        const badge = document.createElement('div');
        badge.id = 'ms365AuthBadge';
        badge.style.display = 'inline-flex';
        badge.style.alignItems = 'center';
        badge.style.gap = '8px';
        badge.style.padding = '8px 10px';
        badge.style.borderRadius = '999px';
        badge.style.border = '1px solid rgba(94, 114, 228, 0.22)';
        badge.style.background = 'rgba(255,255,255,0.86)';
        badge.style.fontWeight = '900';
        badge.style.color = '#32325d';
        badge.style.fontSize = '0.92em';
        badge.innerHTML = '<i class="bi bi-person-check"></i><span id="ms365AuthBadgeText">–</span>';

        const btn = document.createElement('button');
        btn.id = 'ms365AuthBtn';
        btn.type = 'button';
        btn.className = 'btn';
        btn.style.margin = '0';
        btn.style.padding = '10px 14px';
        btn.style.borderRadius = '10px';
        btn.innerHTML = '<i class="bi bi-box-arrow-in-right"></i>Anmelden';

        const actions = document.createElement('div');
        actions.id = 'ms365AuthActions';
        actions.style.display = 'flex';
        actions.style.gap = '10px';
        actions.style.alignItems = 'center';
        actions.style.flexWrap = 'wrap';

        wrap.appendChild(badge);
        wrap.appendChild(actions);
        wrap.appendChild(btn);

        // Wenn wir auf dem Dashboard-Header (kein .toolbar) sind, absolut oben rechts platzieren.
        const hasToolbar = !!$('.toolbar', header);
        if (!hasToolbar && header.tagName && header.tagName.toLowerCase() === 'header') {
            try {
                header.style.position = header.style.position || 'relative';
            } catch {
                // ignore
            }
            wrap.style.position = 'absolute';
            wrap.style.top = '14px';
            wrap.style.right = '14px';
            wrap.style.zIndex = '5';
            wrap.style.marginLeft = '0';
            header.appendChild(wrap);
            return;
        }

        function isFlexLikeRow(el) {
            if (!el || el.nodeType !== 1) return false;
            try {
                const inline = el.style && el.style.display;
                if (inline === 'flex' || inline === 'inline-flex') return true;
            } catch (_) {}
            try {
                const d = window.getComputedStyle(el).display;
                return d === 'flex' || d === 'inline-flex';
            } catch (_) {
                return false;
            }
        }

        // If toolbar has a flex row wrapper, insert into that (inline oder per CSS).
        const first = toolbar.firstElementChild;
        const row = isFlexLikeRow(first) ? first : toolbar;
        row.appendChild(wrap);

        try {
            window.dispatchEvent(new CustomEvent('ms365-auth-widget-ready'));
        } catch {
            // ignore
        }
    }

    async function forceFreshLogin() {
        try {
            const instance = await ensurePca();
            await clearMsalCache(instance);
        } catch {
            // ignore – wir versuchen den Redirect trotzdem
        }
        try {
            return await login(DEFAULT_SCOPES, { prompt: 'select_account' });
        } catch {
            // bei Redirect ohnehin kein weiterer Code mehr
        }
    }

    function ensureSwitchAccountButton() {
        const widget = document.getElementById('ms365AuthWidget');
        if (!widget) return null;
        let extra = document.getElementById('ms365AuthSwitchBtn');
        if (extra) return extra;
        const a = getAccount();
        if (!a) return null;
        extra = document.createElement('button');
        extra.id = 'ms365AuthSwitchBtn';
        extra.type = 'button';
        extra.className = 'btn';
        extra.title = 'Konto wechseln / Anmeldung zurücksetzen (löscht den lokalen MSAL-Cache und meldet frisch an)';
        extra.style.margin = '0';
        extra.style.padding = '8px 12px';
        extra.style.borderRadius = '10px';
        extra.style.background = 'rgba(255,255,255,0.85)';
        extra.style.color = '#32325d';
        extra.style.border = '1px solid rgba(94, 114, 228, 0.35)';
        extra.style.fontSize = '0.88em';
        extra.style.fontWeight = '700';
        extra.innerHTML = '<i class="bi bi-arrow-repeat"></i>Konto wechseln';
        extra.onclick = () => forceFreshLogin().catch(() => {});
        const badge = document.getElementById('ms365AuthBadge');
        const mainBtn = document.getElementById('ms365AuthBtn');
        if (badge && mainBtn && badge.parentElement) {
            badge.parentElement.insertBefore(extra, mainBtn);
        } else {
            widget.appendChild(extra);
        }
        return extra;
    }

    function setWidgetState() {
        const badgeText = document.getElementById('ms365AuthBadgeText');
        const btn = document.getElementById('ms365AuthBtn');
        const a = getAccount();
        if (badgeText) {
            badgeText.textContent = a ? 'Angemeldet' : 'Nicht angemeldet';
            if (a) badgeText.textContent = accountLabel(a);
        }
        if (btn) {
            if (a) {
                btn.innerHTML = '<i class="bi bi-box-arrow-right"></i>Abmelden';
                btn.onclick = () => logout().catch(() => {});
            } else {
                btn.innerHTML = '<i class="bi bi-box-arrow-in-right"></i>Anmelden';
                btn.onclick = () => login(DEFAULT_SCOPES).catch(() => {});
            }
        }
        const switchBtn = document.getElementById('ms365AuthSwitchBtn');
        if (a) {
            ensureSwitchAccountButton();
        } else if (switchBtn && switchBtn.parentElement) {
            switchBtn.parentElement.removeChild(switchBtn);
        }
    }

    async function init() {
        if (typeof document === 'undefined') return;
        ensureHeaderWidget();
        try {
            // In case widget existed already, still notify listeners.
            window.dispatchEvent(new CustomEvent('ms365-auth-widget-ready'));
        } catch {
            // ignore
        }
        try {
            await ensurePca();
        } catch {
            // ignore (widget still renders)
        }
        // Wenn lokal noch kein Account im Cache ist, einmalig SSO Silent versuchen.
        // Damit wird die Microsoft-365-Anmeldung übernommen, wenn der Benutzer
        // in einem anderen Tab/Fenster (z. B. Outlook, Teams Web, anderes Modul) bereits
        // angemeldet ist – ohne sichtbaren Redirect.
        try {
            if (pca && !getAccount()) {
                await trySsoSilent(DEFAULT_SCOPES);
            }
        } catch {
            // ignore
        }
        setWidgetState();
    }

    // Public API for tools
    window.ms365AuthEnsureInitialized = ensurePca;
    window.ms365AuthGetActionSlot = function () {
        try {
            return document.getElementById('ms365AuthActions');
        } catch {
            return null;
        }
    };
    window.ms365AuthGetAccountLabel = function () {
        try {
            return accountLabel(getAccount());
        } catch {
            return '';
        }
    };
    window.ms365AuthIsLoggedIn = function () {
        try {
            return !!getAccount();
        } catch {
            return false;
        }
    };
    window.ms365AuthLogin = login;
    window.ms365AuthSwitchAccount = switchAccount;
    window.ms365AuthLogout = logout;
    window.ms365AuthAcquireToken = acquireToken;

    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
    else init();
})();

