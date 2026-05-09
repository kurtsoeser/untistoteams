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
        return (
            e &&
            (e.name === 'InteractionRequiredAuthError' ||
                e.errorCode === 'interaction_required' ||
                (typeof e.message === 'string' && e.message.indexOf('interaction_required') !== -1))
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
                cache: { cacheLocation: 'sessionStorage', storeAuthStateInCookie: true }
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

    async function login(scopes) {
        const instance = await ensurePca();
        try {
            sessionStorage.setItem(POST_LOGIN_KEY, window.location.href);
        } catch {
            // ignore
        }
        await instance.loginRedirect({
            scopes: Array.isArray(scopes) && scopes.length ? scopes : DEFAULT_SCOPES,
            prompt: 'select_account',
            redirectStartPage: window.location.href
        });
        // redirect -> no further code
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
            if (isInteractionRequired(e)) {
                try {
                    sessionStorage.setItem(POST_LOGIN_KEY, window.location.href);
                } catch {
                    // ignore
                }
                await instance.acquireTokenRedirect({ ...req, redirectStartPage: window.location.href });
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
        setWidgetState();
        try {
            // Falls irgendwo vorher ein Post-Login-Link gesetzt war und diese Seite NICHT ms365-schooltool ist,
            // dann bleibt der Schlüssel stehen. Wir löschen ihn hier zur Sicherheit nicht.
        } catch {
            // ignore
        }
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
    window.ms365AuthLogout = logout;
    window.ms365AuthAcquireToken = acquireToken;

    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
    else init();
})();

