(function (global) {
    'use strict';

    async function getGraphToken(scopes) {
        if (typeof global.ms365AuthAcquireToken === 'function') {
            return await global.ms365AuthAcquireToken(scopes);
        }
        throw new Error('Bitte oben rechts anmelden (MSAL-Widget nicht verfügbar).');
    }

    function sleep(ms) {
        return new Promise(function (r) {
            setTimeout(r, ms);
        });
    }

    function graphBase(version) {
        const v = version === 'beta' ? 'beta' : 'v1.0';
        return 'https://graph.microsoft.com/' + v;
    }

    async function graphRequest(method, pathOrUrl, token, body, version) {
        let url = pathOrUrl;
        if (url.indexOf('http') !== 0) {
            url = graphBase(version) + (pathOrUrl.indexOf('/') === 0 ? pathOrUrl : '/' + pathOrUrl);
        }
        let attempt = 0;
        while (true) {
            const headers = { Authorization: 'Bearer ' + token };
            if (body !== undefined) {
                headers['Content-Type'] = 'application/json';
            }
            const res = await fetch(url, {
                method: method,
                headers: headers,
                body: body !== undefined ? JSON.stringify(body) : undefined
            });
            if (res.status === 429 && attempt < 8) {
                const ra = parseInt(res.headers.get('Retry-After') || '5', 10);
                await sleep((isNaN(ra) ? 5 : ra) * 1000);
                attempt++;
                continue;
            }
            return res;
        }
    }

    async function graphJson(method, pathOrUrl, token, body, version) {
        const res = await graphRequest(method, pathOrUrl, token, body, version);
        const text = await res.text();
        let data = null;
        if (text) {
            try {
                data = JSON.parse(text);
            } catch {
                data = text;
            }
        }
        if (!res.ok) {
            const msg =
                typeof data === 'object' && data && data.error
                    ? (data.error.message || JSON.stringify(data.error))
                    : text || String(res.status);
            const err = new Error(method + ' ' + pathOrUrl + ': ' + msg);
            err.status = res.status;
            err.payload = data;
            throw err;
        }
        return data || {};
    }

    async function getSharePointHostname(token) {
        const sitesRead = [
            'https://graph.microsoft.com/User.Read',
            'https://graph.microsoft.com/Sites.Read.All'
        ];
        let t = token;
        if (!t) t = await getGraphToken(sitesRead);
        const root = await graphJson('GET', '/sites/root', t, undefined, 'v1.0');
        const w = root && root.webUrl ? String(root.webUrl) : '';
        if (!w) return '';
        try {
            return new URL(w).hostname;
        } catch {
            return '';
        }
    }

    /**
     * @param {string} operationUrl Vollständige URL aus dem Location-Header (202)
     * @param {string} token Graph-Zugriffstoken
     */
    async function pollRichLongRunningOperation(operationUrl, token) {
        const max = 45;
        for (let i = 0; i < max; i++) {
            const res = await fetch(operationUrl, {
                method: 'GET',
                headers: { Authorization: 'Bearer ' + token }
            });
            const text = await res.text();
            let data = null;
            if (text) {
                try {
                    data = JSON.parse(text);
                } catch {
                    data = { raw: text };
                }
            }
            if (!res.ok) {
                throw new Error('Vorgang: HTTP ' + res.status + ' – ' + (text || ''));
            }
            const status = data && (data.status || data.Status);
            const s = String(status || '').toLowerCase();
            if (s === 'succeeded' || s === 'completed' || s === 'complete') {
                return data;
            }
            if (s === 'failed' || s === 'cancelled' || s === 'canceled') {
                throw new Error('Vorgang fehlgeschlagen: ' + (data && (data.error || data.resourceId) ? JSON.stringify(data.error || data) : JSON.stringify(data)));
            }
            /* notStarted, running, waiting … → weiter pollen */
            await sleep(2000);
        }
        throw new Error('Timeout: Die Site-Erstellung dauert ungewöhnlich lange. Bitte im SharePoint Admin Center prüfen.');
    }

    /**
     * SharePoint REST: RequestDigest, dann RegisterHubSite (kann im Browser an CORS scheitern).
     * @param {string} siteWebUrl z. B. https://tenant.sharepoint.com/sites/Intranet
     * @param {string} spoToken Zugriffstoken mit Audience https://tenant.sharepoint.com/
     */
    async function registerHubSiteViaSpoRest(siteWebUrl, spoToken) {
        const origin = String(siteWebUrl || '').replace(/\/+$/, '');
        if (!origin || !spoToken) throw new Error('Site-URL oder SharePoint-Token fehlt.');

        const ctxRes = await fetch(origin + '/_api/contextinfo', {
            method: 'POST',
            headers: {
                Accept: 'application/json;odata=nometadata',
                'Content-Type': 'application/json;odata=nometadata;charset=utf-8',
                Authorization: 'Bearer ' + spoToken
            }
        });
        const ctxText = await ctxRes.text();
        if (!ctxRes.ok) {
            throw new Error('contextinfo: ' + ctxRes.status + ' ' + (ctxText || ''));
        }
        let ctxJson = null;
        try {
            ctxJson = JSON.parse(ctxText);
        } catch {
            throw new Error('contextinfo: keine JSON-Antwort');
        }
        const digest =
            (ctxJson && ctxJson.FormDigestValue) ||
            (ctxJson && ctxJson.d && ctxJson.d.GetContextWebInformation && ctxJson.d.GetContextWebInformation.FormDigestValue) ||
            '';
        if (!digest) throw new Error('Kein FormDigestValue erhalten.');

        const hubRes = await fetch(origin + '/_api/site/RegisterHubSite', {
            method: 'POST',
            headers: {
                Accept: 'application/json;odata=nometadata',
                'Content-Type': 'application/json;odata=nometadata;charset=utf-8',
                Authorization: 'Bearer ' + spoToken,
                'X-RequestDigest': digest
            },
            body: ''
        });
        const hubText = await hubRes.text();
        if (!hubRes.ok) {
            throw new Error('RegisterHubSite: ' + hubRes.status + ' ' + (hubText || ''));
        }
        let hubJson = null;
        try {
            hubJson = JSON.parse(hubText);
        } catch {
            hubJson = { raw: hubText };
        }
        return hubJson;
    }

    /**
     * @returns {{ host: string, serverRelativeUrl: string } | null}
     */
    function parseSharePointWebUrl(input) {
        let raw = String(input || '').trim();
        if (!raw) return null;
        if (!/^https?:\/\//i.test(raw)) raw = 'https://' + raw;
        let u;
        try {
            u = new URL(raw);
        } catch {
            return null;
        }
        const host = String(u.hostname || '')
            .trim()
            .toLowerCase();
        if (!host) return null;
        let pathname = u.pathname || '';
        if (pathname.length > 1 && pathname.endsWith('/')) pathname = pathname.slice(0, -1);
        if (!pathname || pathname === '') pathname = '/';
        return { host: host, serverRelativeUrl: pathname };
    }

    /**
     * GET /sites/{hostname}:/{path} – liefert Site-Ressource inkl. id.
     */
    async function resolveSiteFromWebUrl(token, webUrlInput) {
        const parts = parseSharePointWebUrl(webUrlInput);
        if (!parts) throw new Error('Ungültige SharePoint-Website-URL.');
        const hostPath = parts.host + ':' + parts.serverRelativeUrl;
        const seg = encodeURIComponent(hostPath);
        return await graphJson('GET', '/sites/' + seg, token, undefined, 'v1.0');
    }

    function graphPathSite(siteId) {
        return '/sites/' + encodeURIComponent(siteId);
    }

    global.ms365SpoGraph = {
        getGraphToken: getGraphToken,
        sleep: sleep,
        graphRequest: graphRequest,
        graphJson: graphJson,
        getSharePointHostname: getSharePointHostname,
        pollRichLongRunningOperation: pollRichLongRunningOperation,
        registerHubSiteViaSpoRest: registerHubSiteViaSpoRest,
        graphBase: graphBase,
        parseSharePointWebUrl: parseSharePointWebUrl,
        resolveSiteFromWebUrl: resolveSiteFromWebUrl,
        graphPathSite: graphPathSite
    };
})(typeof window !== 'undefined' ? window : globalThis);
