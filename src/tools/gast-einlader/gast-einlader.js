const STORAGE_KEY = 'ms365-gast-einlader-policy-v1';

// Vorlagen-/Rollen-IDs in Entra (Microsoft Graph)
// "Guest Inviter" / "Gasteinladende:r":
const GUEST_INVITER_ROLE_TEMPLATE_ID = '95e79109-95c0-4d8e-aee3-d01accf2d47b';

const GRAPH_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Group.ReadWrite.All',
    'https://graph.microsoft.com/RoleManagement.ReadWrite.Directory',
    'https://graph.microsoft.com/Policy.ReadWrite.Authorization'
];

let giCurrentStep = 1;

function toast(msg) {
    if (typeof window.ms365ToastOrAlert === 'function') {
        window.ms365ToastOrAlert(msg);
    } else if (typeof window.ms365ShowToast === 'function') {
        window.ms365ShowToast(msg);
    } else {
        window.alert(msg);
    }
}

async function getGraphToken() {
    if (typeof window.ms365AuthEnsureInitialized === 'function') {
        try {
            await window.ms365AuthEnsureInitialized();
        } catch {
            /* ignore */
        }
    }
    if (typeof window.ms365AuthAcquireToken !== 'function') {
        throw new Error('Microsoft-Anmeldung nicht verfügbar (msal-auth-ui.js fehlt).');
    }
    return window.ms365AuthAcquireToken(GRAPH_SCOPES);
}

function sleep(ms) {
    return new Promise(function (r) {
        setTimeout(r, ms);
    });
}

async function graphRequest(method, path, token, body, extraHeaders) {
    const url = path.indexOf('http') === 0 ? path : 'https://graph.microsoft.com/v1.0' + path;
    let attempt = 0;
    while (true) {
        const headers = { Authorization: 'Bearer ' + token };
        if (extraHeaders && typeof extraHeaders === 'object') Object.assign(headers, extraHeaders);
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

async function graphJson(method, path, token, body, extraHeaders) {
    const res = await graphRequest(method, path, token, body, extraHeaders);
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
                ? JSON.stringify(data.error)
                : text || String(res.status);
        throw new Error(method + ' ' + path + ': ' + msg);
    }
    return data || {};
}

async function graphCollect(path, token) {
    let url = path;
    const out = [];
    while (url) {
        const data = await graphJson('GET', url, token, undefined);
        if (Array.isArray(data.value)) {
            data.value.forEach(function (v) {
                out.push(v);
            });
        }
        url = data['@odata.nextLink'] || '';
    }
    return out;
}

function appendLog(msg, kind) {
    const el = document.getElementById('giOnlineLog');
    if (!el) return;
    const line = document.createElement('div');
    line.textContent = new Date().toLocaleTimeString() + '  ' + msg;
    if (kind === 'err') line.style.color = '#b00020';
    else if (kind === 'ok') line.style.color = '#0d8050';
    else if (kind === 'warn') line.style.color = '#856404';
    else line.style.color = '#212529';
    el.appendChild(line);
    el.scrollTop = el.scrollHeight;
}

function clearLog() {
    const el = document.getElementById('giOnlineLog');
    if (el) el.replaceChildren();
}

function guidLooksValid(s) {
    return /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(
        String(s || '').trim()
    );
}

function sanitizeMailNickname(name) {
    let n = String(name || '')
        .replace(/[^0-9a-zA-Z]/g, '')
        .slice(0, 60);
    if (!n) n = 'gasteinlader';
    return n.toLowerCase();
}

// ----- Graph: Authorization Policy -----

async function getAuthorizationPolicy(token) {
    return graphJson('GET', '/policies/authorizationPolicy', token, undefined);
}

async function setAllowInvitesFrom(token, value) {
    return graphJson('PATCH', '/policies/authorizationPolicy', token, { allowInvitesFrom: value });
}

// ----- Graph: Guest Inviter Rolle -----

async function activateGuestInviterRole(token) {
    // Versuche, die Rolle (sofern noch nicht aktiviert) per directoryRoles zu aktivieren.
    // /directoryRoles enthält nur aktivierte Rollen; die Aktivierung erfolgt klassisch über POST.
    try {
        return await graphJson('POST', '/directoryRoles', token, {
            roleTemplateId: GUEST_INVITER_ROLE_TEMPLATE_ID
        });
    } catch (e) {
        // Wenn die Rolle bereits aktiviert ist, kommt 409/Request_BadRequest – das ist okay.
        if (/already exist|conflict|409/i.test(String(e.message))) {
            return null;
        }
        // Wenn der Tenant moderne RBAC nutzt, ist eine "Aktivierung" nicht nötig.
        return null;
    }
}

async function getGuestInviterRoleObject(token) {
    // Erst aktivierte Rolle suchen (älterer Pfad)
    try {
        const filter =
            "roleTemplateId eq '" + GUEST_INVITER_ROLE_TEMPLATE_ID + "'";
        const data = await graphJson(
            'GET',
            '/directoryRoles?$filter=' + encodeURIComponent(filter),
            token,
            undefined
        );
        const row = (data.value || [])[0];
        if (row) return { kind: 'directoryRole', id: row.id, displayName: row.displayName };
    } catch {
        /* ignore */
    }
    // Versuche Aktivierung
    try {
        const created = await activateGuestInviterRole(token);
        if (created && created.id) {
            return { kind: 'directoryRole', id: created.id, displayName: created.displayName || 'Guest Inviter' };
        }
    } catch {
        /* ignore */
    }
    // Erneuter Versuch nach Aktivierung
    try {
        const filter =
            "roleTemplateId eq '" + GUEST_INVITER_ROLE_TEMPLATE_ID + "'";
        const data = await graphJson(
            'GET',
            '/directoryRoles?$filter=' + encodeURIComponent(filter),
            token,
            undefined
        );
        const row = (data.value || [])[0];
        if (row) return { kind: 'directoryRole', id: row.id, displayName: row.displayName };
    } catch {
        /* ignore */
    }
    throw new Error('Rolle „Gasteinladende:r“ konnte nicht ermittelt/aktiviert werden.');
}

async function getRoleMembers(token, role) {
    return graphCollect(
        '/directoryRoles/' +
            encodeURIComponent(role.id) +
            '/members?$select=id,displayName,userPrincipalName',
        token
    );
}

async function addRoleMember(token, role, principalId) {
    const body = {
        '@odata.id': 'https://graph.microsoft.com/v1.0/directoryObjects/' + encodeURIComponent(principalId)
    };
    return graphJson(
        'POST',
        '/directoryRoles/' + encodeURIComponent(role.id) + '/members/$ref',
        token,
        body
    );
}

async function removeRoleMember(token, role, principalId) {
    return graphJson(
        'DELETE',
        '/directoryRoles/' +
            encodeURIComponent(role.id) +
            '/members/' +
            encodeURIComponent(principalId) +
            '/$ref',
        token,
        undefined
    );
}

// ----- Graph: Gruppen -----

async function getGroupById(token, id) {
    // $select inkl. isAssignableToRole (rollenfähige Gruppe)
    return graphJson(
        'GET',
        '/groups/' +
            encodeURIComponent(id) +
            '?$select=id,displayName,description,mailEnabled,securityEnabled,groupTypes,mailNickname,isAssignableToRole',
        token,
        undefined
    );
}

async function findSecurityGroupByDisplayName(token, displayName) {
    const esc = String(displayName).replace(/'/g, "''");
    const filter = "displayName eq '" + esc + "' and securityEnabled eq true";
    const path = '/groups?$filter=' + encodeURIComponent(filter) + '&$top=25';
    const data = await graphJson('GET', path, token, undefined);
    return data.value || [];
}

async function createRoleAssignableSecurityGroup(token, displayName) {
    let nick = sanitizeMailNickname(displayName);
    const body = {
        displayName: String(displayName).trim(),
        description:
            'Mitglieder dieser Gruppe dürfen im Mandanten Gäste einladen (Entra: Rolle „Gasteinladende:r“ + authorizationPolicy.allowInvitesFrom=adminsAndGuestInviters). Rollenfähige Sicherheitsgruppe (isAssignableToRole=true).',
        mailEnabled: false,
        mailNickname: nick,
        securityEnabled: true,
        isAssignableToRole: true,
        groupTypes: []
    };
    try {
        return await graphJson('POST', '/groups', token, body);
    } catch (e) {
        if (
            String(e.message || e).indexOf('mailNickname') !== -1 ||
            /409|conflict/i.test(String(e.message))
        ) {
            body.mailNickname = nick + '-' + Math.random().toString(36).slice(2, 8);
            return await graphJson('POST', '/groups', token, body);
        }
        throw e;
    }
}

async function getGroupMembers(token, groupId) {
    return graphCollect(
        '/groups/' +
            encodeURIComponent(groupId) +
            '/members?$select=id,displayName,userPrincipalName',
        token
    );
}

async function getGroupMemberCount(token, groupId) {
    try {
        const data = await graphJson(
            'GET',
            '/groups/' + encodeURIComponent(groupId) + '/members/$count',
            token,
            undefined
        );
        if (typeof data === 'number') return data;
        const n = parseInt(String(data), 10);
        if (!isNaN(n)) return n;
    } catch {
        /* ignore */
    }
    try {
        const arr = await getGroupMembers(token, groupId);
        return arr.length;
    } catch {
        return -1;
    }
}

async function getGroupOwners(token, groupId) {
    return graphCollect(
        '/groups/' +
            encodeURIComponent(groupId) +
            '/owners?$select=id,displayName,userPrincipalName,mail',
        token
    );
}

function userRefUrl(userId) {
    return 'https://graph.microsoft.com/v1.0/directoryObjects/' + encodeURIComponent(userId);
}

async function addGroupMember(token, groupId, principalId) {
    return graphJson(
        'POST',
        '/groups/' + encodeURIComponent(groupId) + '/members/$ref',
        token,
        { '@odata.id': userRefUrl(principalId) }
    );
}

async function removeGroupMember(token, groupId, principalId) {
    return graphJson(
        'DELETE',
        '/groups/' +
            encodeURIComponent(groupId) +
            '/members/' +
            encodeURIComponent(principalId) +
            '/$ref',
        token,
        undefined
    );
}

async function addGroupOwner(token, groupId, principalId) {
    return graphJson(
        'POST',
        '/groups/' + encodeURIComponent(groupId) + '/owners/$ref',
        token,
        { '@odata.id': userRefUrl(principalId) }
    );
}

async function removeGroupOwner(token, groupId, principalId) {
    return graphJson(
        'DELETE',
        '/groups/' +
            encodeURIComponent(groupId) +
            '/owners/' +
            encodeURIComponent(principalId) +
            '/$ref',
        token,
        undefined
    );
}

function odataEscape(s) {
    return String(s || '').replace(/'/g, "''");
}

/** Sucht eine Person für die Owner-/Member-Verwaltung.
 *  – Enthält die Eingabe ein '@', wird gezielt per UPN/Mail aufgelöst.
 *  – Sonst per Volltextsuche ($search) und Fallback per startswith(displayName).
 *  Gibt bis zu maxResults Treffer zurück (id, displayName, userPrincipalName, mail).
 */
async function searchUsersForGroup(token, queryRaw, maxResults) {
    const q = String(queryRaw || '').trim();
    const limit = Math.max(1, Math.min(20, maxResults || 6));
    if (!q) return [];
    const select = 'id,displayName,userPrincipalName,mail,accountEnabled';

    if (q.indexOf('@') !== -1) {
        const esc = odataEscape(q);
        const filter = "(mail eq '" + esc + "' or userPrincipalName eq '" + esc + "')";
        const path =
            '/users?$filter=' +
            encodeURIComponent(filter) +
            '&$select=' +
            encodeURIComponent(select) +
            '&$top=' +
            limit;
        try {
            const data = await graphJson('GET', path, token, undefined);
            return Array.isArray(data.value) ? data.value : [];
        } catch {
            /* unten weiter mit Suche */
        }
    }

    try {
        const phrase = '"' + q.replace(/"/g, '') + '"';
        const aqs =
            '(displayName:' + phrase + ' OR userPrincipalName:' + phrase + ' OR mail:' + phrase + ')';
        const path =
            '/users?$search=' +
            encodeURIComponent('"' + aqs + '"') +
            '&$select=' +
            encodeURIComponent(select) +
            '&$top=' +
            limit;
        const data = await graphJson('GET', path, token, undefined, { ConsistencyLevel: 'eventual' });
        const list = Array.isArray(data.value) ? data.value : [];
        if (list.length) return list;
    } catch {
        /* Fallback unten */
    }

    try {
        const esc = odataEscape(q);
        const filter =
            "startswith(displayName,'" +
            esc +
            "') or startswith(userPrincipalName,'" +
            esc +
            "') or startswith(mail,'" +
            esc +
            "')";
        const path =
            '/users?$filter=' +
            encodeURIComponent(filter) +
            '&$select=' +
            encodeURIComponent(select) +
            '&$top=' +
            limit;
        const data = await graphJson('GET', path, token, undefined);
        return Array.isArray(data.value) ? data.value : [];
    } catch {
        return [];
    }
}

async function getCurrentUserId(token) {
    try {
        const me = await graphJson('GET', '/me?$select=id,userPrincipalName', token, undefined);
        return me && me.id ? String(me.id) : '';
    } catch {
        return '';
    }
}

// ----- UI: gewählte Gruppe -----

function entraGroupBladeUrl(tab, groupId) {
    return (
        'https://entra.microsoft.com/#view/Microsoft_AAD_IAM/GroupDetailsMenuBlade/~/' +
        tab +
        '/groupId/' +
        encodeURIComponent(groupId)
    );
}

function entraGuestInviterRoleBladeUrl() {
    return (
        'https://entra.microsoft.com/#view/Microsoft_AAD_IAM/RoleMenuBlade/~/AdminAssignments/roleTemplateId/' +
        encodeURIComponent(GUEST_INVITER_ROLE_TEMPLATE_ID)
    );
}

function setResolvedGroup(id, displayName) {
    const hid = document.getElementById('giResolvedObjectId');
    const out = document.getElementById('giResolvedSummary');
    const linkOv = document.getElementById('giLinkOverview');
    const linkOwn = document.getElementById('giLinkOwners');
    const linkMem = document.getElementById('giLinkMembers');
    const linkRole = document.getElementById('giLinkRole');
    if (hid) hid.value = id || '';
    if (out) {
        out.style.display = id ? 'block' : 'none';
        out.textContent = id
            ? 'Ausgewählte Gruppe: ' + (displayName ? displayName + ' · ' : '') + 'Object-ID ' + id
            : '';
    }
    if (linkOv) {
        linkOv.style.display = id ? 'inline' : 'none';
        if (id) linkOv.href = entraGroupBladeUrl('Overview', id);
    }
    if (linkOwn) {
        linkOwn.style.display = id ? 'inline' : 'none';
        if (id) linkOwn.href = entraGroupBladeUrl('Owners', id);
    }
    if (linkMem) {
        linkMem.style.display = id ? 'inline' : 'none';
        if (id) linkMem.href = entraGroupBladeUrl('Members', id);
    }
    if (linkRole) {
        linkRole.style.display = 'inline';
        linkRole.href = entraGuestInviterRoleBladeUrl();
    }
}

function describeGraphGroupKind(g) {
    if (!g) return '–';
    const types = g.groupTypes || [];
    if (types.indexOf('Unified') !== -1) return 'Microsoft 365-Gruppe';
    if (g.securityEnabled && g.mailEnabled) return 'E-Mail-Sicherheitsgruppe';
    if (g.securityEnabled && !g.mailEnabled) return 'Sicherheitsgruppe';
    if (g.mailEnabled) return 'Mail-aktivierte Gruppe';
    return 'Gruppe';
}

async function loadGroupDetailsIntoStep2() {
    const hid = document.getElementById('giResolvedObjectId');
    const gid = hid && hid.value ? hid.value.trim() : '';
    const ph = document.getElementById('giDetailsPlaceholder');
    const dl = document.getElementById('giDetailsDl');
    if (!ph || !dl) return;
    if (!guidLooksValid(gid)) {
        ph.style.display = 'block';
        ph.textContent = 'Keine gültige Gruppen-Object-ID – bitte zuerst Schritt 1 abschließen.';
        dl.style.display = 'none';
        return;
    }
    ph.style.display = 'block';
    ph.textContent = 'Lese Gruppe aus Microsoft Graph …';
    dl.style.display = 'none';
    try {
        const token = await getGraphToken();
        const g = await getGroupById(token, gid);
        const setTxt = function (id, text) {
            const n = document.getElementById(id);
            if (n) n.textContent = text != null && String(text) !== '' ? String(text) : '–';
        };
        const count = await getGroupMemberCount(token, gid);
        setTxt('giDetDisplayName', g.displayName);
        setTxt('giDetType', describeGraphGroupKind(g));
        setTxt(
            'giDetRoleAssignable',
            g.isAssignableToRole === true ? 'Ja (isAssignableToRole=true)' : 'Nein – Fallback nötig'
        );
        setTxt('giDetMailEnabled', g.mailEnabled ? 'Ja' : 'Nein');
        setTxt('giDetDescription', g.description);
        setTxt('giDetMemberCount', count >= 0 ? String(count) : '–');
        setTxt('giDetId', g.id);
        ph.style.display = 'none';
        dl.style.display = 'block';

        // Buttons je nach Eignung
        const apply = document.getElementById('giBtnApply');
        if (apply) {
            if (g.isAssignableToRole === true && g.securityEnabled === true && g.mailEnabled === false) {
                apply.disabled = false;
                apply.title = 'Rollenfähige reine Sicherheitsgruppe – Direktzuweisung möglich.';
            } else {
                apply.disabled = true;
                apply.title =
                    'Diese Gruppe kann der Rolle nicht direkt zugewiesen werden. Bitte Fallback (Mitglieder synchronisieren) verwenden.';
            }
        }
    } catch (e) {
        ph.style.display = 'block';
        ph.textContent = 'Gruppe konnte nicht gelesen werden: ' + (e.message || e);
        dl.style.display = 'none';
    }
}

// ----- Mitglieder- und Besitzer-Verwaltung -----

const giPeopleState = {
    currentUserOid: '',
    ownersLoadedFor: '',
    membersLoadedFor: ''
};

function escHtml(s) {
    return String(s == null ? '' : s)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function principalTypeLabel(odataType) {
    const t = String(odataType || '').toLowerCase();
    if (t.indexOf('user') !== -1) return { label: 'Benutzer:in', cls: 'is-user' };
    if (t.indexOf('group') !== -1) return { label: 'Gruppe (verschachtelt)', cls: 'is-group' };
    if (t.indexOf('serviceprincipal') !== -1) return { label: 'App / Service-Principal', cls: 'is-group' };
    return { label: '–', cls: '' };
}

function renderPeopleRows(tbody, rows, kind, currentUserOid) {
    if (!tbody) return;
    if (!rows.length) {
        tbody.innerHTML =
            '<tr><td colspan="3" class="muted" style="padding:14px 10px;">' +
            (kind === 'owner' ? 'Keine Besitzer:innen vorhanden.' : 'Keine Mitglieder vorhanden.') +
            '</td></tr>';
        return;
    }
    const frag = document.createDocumentFragment();
    for (const p of rows) {
        const tr = document.createElement('tr');
        const isSelf = currentUserOid && currentUserOid === p.id;
        if (isSelf) tr.classList.add('gi-row-self');
        const typ = principalTypeLabel(p['@odata.type']);
        const isUser = typ.cls === 'is-user';
        const pillHtml =
            typ.cls === 'is-user'
                ? ''
                : '<span class="gi-type-pill ' + typ.cls + '">' + escHtml(typ.label) + '</span>';
        tr.innerHTML =
            '<td><strong>' +
            escHtml(p.displayName || '–') +
            '</strong>' +
            pillHtml +
            (isSelf ? ' <span class="muted" style="font-size:0.85em;">(Sie selbst)</span>' : '') +
            '</td>' +
            '<td>' +
            escHtml(p.userPrincipalName || p.mail || '') +
            '</td>' +
            '<td>' +
            '<button type="button" class="btn btn-danger btn-mini" data-gi-act="remove" data-gi-kind="' +
            kind +
            '" data-gi-id="' +
            escHtml(p.id) +
            '" data-gi-label="' +
            escHtml(p.displayName || p.userPrincipalName || p.id) +
            '"' +
            (kind === 'owner' && isSelf && isUser
                ? ' title="Achtung: Sie würden sich selbst aus den Besitzern entfernen."'
                : '') +
            '><i class="bi bi-x-circle"></i>Entfernen</button>' +
            '</td>';
        frag.appendChild(tr);
    }
    tbody.replaceChildren(frag);
}

function setPeopleLoading(tbody, label) {
    if (!tbody) return;
    tbody.innerHTML =
        '<tr><td colspan="3" class="muted" style="padding:14px 10px;">' +
        escHtml(label || 'Wird geladen …') +
        '</td></tr>';
}

function setPeopleError(tbody, msg) {
    if (!tbody) return;
    tbody.innerHTML =
        '<tr><td colspan="3" style="padding:14px 10px;color:#9b2c2c;line-height:1.45;">' +
        escHtml(msg || 'Fehler beim Laden.') +
        '</td></tr>';
}

function getResolvedGroupId() {
    const hid = document.getElementById('giResolvedObjectId');
    const gid = hid && hid.value ? hid.value.trim() : '';
    return guidLooksValid(gid) ? gid : '';
}

async function refreshOwnersList(silent) {
    const tbody = document.getElementById('giOwnersTbody');
    const countEl = document.getElementById('giOwnersCount');
    const gid = getResolvedGroupId();
    if (!gid) {
        if (tbody) tbody.innerHTML = '<tr><td colspan="3" class="muted" style="padding:14px 10px;">Noch keine Gruppe gewählt.</td></tr>';
        if (countEl) countEl.textContent = '';
        giPeopleState.ownersLoadedFor = '';
        return;
    }
    if (!silent) setPeopleLoading(tbody, 'Besitzer:innen werden geladen …');
    try {
        const token = await getGraphToken();
        if (!giPeopleState.currentUserOid) {
            giPeopleState.currentUserOid = await getCurrentUserId(token);
        }
        const list = await getGroupOwners(token, gid);
        list.sort(function (a, b) {
            return String(a.displayName || '').localeCompare(String(b.displayName || ''), 'de', { sensitivity: 'base' });
        });
        renderPeopleRows(tbody, list, 'owner', giPeopleState.currentUserOid);
        if (countEl) countEl.textContent = list.length + ' Besitzer' + (list.length === 1 ? '' : ':innen');
        giPeopleState.ownersLoadedFor = gid;
    } catch (e) {
        setPeopleError(tbody, 'Besitzer konnten nicht geladen werden: ' + (e && e.message ? e.message : String(e)));
        if (countEl) countEl.textContent = '';
    }
}

async function refreshMembersList(silent) {
    const tbody = document.getElementById('giMembersTbody');
    const countEl = document.getElementById('giMembersCount');
    const gid = getResolvedGroupId();
    if (!gid) {
        if (tbody) tbody.innerHTML = '<tr><td colspan="3" class="muted" style="padding:14px 10px;">Noch keine Gruppe gewählt.</td></tr>';
        if (countEl) countEl.textContent = '';
        giPeopleState.membersLoadedFor = '';
        return;
    }
    if (!silent) setPeopleLoading(tbody, 'Mitglieder werden geladen …');
    try {
        const token = await getGraphToken();
        if (!giPeopleState.currentUserOid) {
            giPeopleState.currentUserOid = await getCurrentUserId(token);
        }
        const list = await getGroupMembers(token, gid);
        list.sort(function (a, b) {
            return String(a.displayName || '').localeCompare(String(b.displayName || ''), 'de', { sensitivity: 'base' });
        });
        renderPeopleRows(tbody, list, 'member', giPeopleState.currentUserOid);
        if (countEl) countEl.textContent = list.length + ' Mitglied' + (list.length === 1 ? '' : 'er');
        giPeopleState.membersLoadedFor = gid;
    } catch (e) {
        setPeopleError(tbody, 'Mitglieder konnten nicht geladen werden: ' + (e && e.message ? e.message : String(e)));
        if (countEl) countEl.textContent = '';
    }
}

function renderSuggestions(kind, candidates) {
    const box = document.getElementById(
        kind === 'owner' ? 'giOwnerSuggestions' : 'giMemberSuggestions'
    );
    if (!box) return;
    if (!candidates || !candidates.length) {
        box.innerHTML = '<div class="gi-sugg-empty">Keine passenden Personen im Verzeichnis gefunden.</div>';
        box.setAttribute('data-open', '1');
        return;
    }
    const frag = document.createDocumentFragment();
    for (const c of candidates) {
        const row = document.createElement('div');
        row.className = 'gi-sugg-row';
        const upn = c.userPrincipalName || c.mail || '';
        row.innerHTML =
            '<div style="min-width:0;flex:1;">' +
            '<strong>' +
            escHtml(c.displayName || '–') +
            '</strong>' +
            (upn ? ' <span class="muted">&lt;' + escHtml(upn) + '&gt;</span>' : '') +
            '</div>' +
            '<button type="button" data-gi-pick="' +
            escHtml(c.id) +
            '" data-gi-pick-label="' +
            escHtml(c.displayName || upn || c.id) +
            '">Auswählen</button>';
        frag.appendChild(row);
    }
    box.replaceChildren(frag);
    box.setAttribute('data-open', '1');
}

function closeSuggestions(kind) {
    const box = document.getElementById(
        kind === 'owner' ? 'giOwnerSuggestions' : 'giMemberSuggestions'
    );
    if (!box) return;
    box.removeAttribute('data-open');
    box.innerHTML = '';
}

async function addPrincipalToGroup(kind, principalId, label) {
    const gid = getResolvedGroupId();
    if (!gid) {
        toast('Bitte zuerst in Schritt 1 eine Sicherheitsgruppe wählen oder anlegen.');
        return false;
    }
    clearLog();
    appendLog(
        (kind === 'owner' ? 'Besitzer' : 'Mitglied') + ' wird hinzugefügt: ' + (label || principalId),
        ''
    );
    try {
        const token = await getGraphToken();
        if (kind === 'owner') await addGroupOwner(token, gid, principalId);
        else await addGroupMember(token, gid, principalId);
        appendLog('Erfolgreich hinzugefügt.', 'ok');
        toast((kind === 'owner' ? 'Besitzer' : 'Mitglied') + ' hinzugefügt.');
        return true;
    } catch (e) {
        const msg = String(e && e.message ? e.message : e);
        if (/already exist|already a member|conflict|409/i.test(msg)) {
            appendLog(
                'Hinweis: Person ist bereits ' + (kind === 'owner' ? 'Besitzer:in' : 'Mitglied') + ' (übersprungen).',
                'warn'
            );
            toast('Bereits ' + (kind === 'owner' ? 'Besitzer:in' : 'Mitglied') + '.');
            return true;
        }
        appendLog('Fehler: ' + msg, 'err');
        toast('Fehler: ' + msg);
        return false;
    }
}

async function removePrincipalFromGroup(kind, principalId, label) {
    const gid = getResolvedGroupId();
    if (!gid) return false;
    const isSelfOwner =
        kind === 'owner' && giPeopleState.currentUserOid && giPeopleState.currentUserOid === principalId;
    const confirmText = isSelfOwner
        ? 'Soll Ihre eigene Besitzer-Berechtigung an der Sicherheitsgruppe „' +
          (label || principalId) +
          '“ wirklich entfernt werden? Sie könnten danach die Gruppe nicht mehr verwalten.'
        : (kind === 'owner' ? 'Besitzer:in „' : 'Mitglied „') +
          (label || principalId) +
          '“ aus der Gruppe entfernen?';
    const ok =
        typeof window.ms365AppDialogConfirm === 'function'
            ? await window.ms365AppDialogConfirm(confirmText, {
                  title: kind === 'owner' ? 'Besitzer entfernen' : 'Mitglied entfernen',
                  okText: 'Entfernen',
                  danger: true
              })
            : window.confirm(confirmText);
    if (!ok) return false;
    clearLog();
    appendLog(
        (kind === 'owner' ? 'Besitzer' : 'Mitglied') + ' wird entfernt: ' + (label || principalId),
        ''
    );
    try {
        const token = await getGraphToken();
        if (kind === 'owner') await removeGroupOwner(token, gid, principalId);
        else await removeGroupMember(token, gid, principalId);
        appendLog('Erfolgreich entfernt.', 'ok');
        toast((kind === 'owner' ? 'Besitzer' : 'Mitglied') + ' entfernt.');
        return true;
    } catch (e) {
        const msg = String(e && e.message ? e.message : e);
        appendLog('Fehler beim Entfernen: ' + msg, 'err');
        toast('Fehler: ' + msg);
        return false;
    }
}

async function onAddPersonClick(kind) {
    const inp = document.getElementById(kind === 'owner' ? 'giOwnerAddInput' : 'giMemberAddInput');
    const q = inp && inp.value ? inp.value.trim() : '';
    if (!q) {
        toast('Bitte UPN, E-Mail oder Anzeigename eintippen.');
        inp && inp.focus();
        return;
    }
    if (!getResolvedGroupId()) {
        toast('Bitte zuerst in Schritt 1 eine Sicherheitsgruppe wählen oder anlegen.');
        return;
    }
    const btn = document.getElementById(kind === 'owner' ? 'giBtnAddOwner' : 'giBtnAddMember');
    if (btn) btn.disabled = true;
    try {
        const token = await getGraphToken();
        const candidates = await searchUsersForGroup(token, q, 6);
        if (!candidates.length) {
            renderSuggestions(kind, []);
            return;
        }
        if (candidates.length === 1) {
            const c = candidates[0];
            const ok = await addPrincipalToGroup(kind, c.id, c.displayName || c.userPrincipalName || c.id);
            if (ok) {
                if (inp) inp.value = '';
                closeSuggestions(kind);
                if (kind === 'owner') await refreshOwnersList(true);
                else await refreshMembersList(true);
            }
            return;
        }
        renderSuggestions(kind, candidates);
    } catch (e) {
        const msg = String(e && e.message ? e.message : e);
        appendLog('Fehler bei der Personensuche: ' + msg, 'err');
        toast('Fehler: ' + msg);
    } finally {
        if (btn) btn.disabled = false;
    }
}

function bindPeopleManagement() {
    const btnAddOwner = document.getElementById('giBtnAddOwner');
    const btnAddMember = document.getElementById('giBtnAddMember');
    const btnRefreshOwners = document.getElementById('giBtnRefreshOwners');
    const btnRefreshMembers = document.getElementById('giBtnRefreshMembers');
    const inpOwner = document.getElementById('giOwnerAddInput');
    const inpMember = document.getElementById('giMemberAddInput');
    const tbodyOwners = document.getElementById('giOwnersTbody');
    const tbodyMembers = document.getElementById('giMembersTbody');
    const suggOwner = document.getElementById('giOwnerSuggestions');
    const suggMember = document.getElementById('giMemberSuggestions');

    if (btnAddOwner && btnAddOwner.dataset.giBound !== '1') {
        btnAddOwner.dataset.giBound = '1';
        btnAddOwner.addEventListener('click', function () { onAddPersonClick('owner'); });
    }
    if (btnAddMember && btnAddMember.dataset.giBound !== '1') {
        btnAddMember.dataset.giBound = '1';
        btnAddMember.addEventListener('click', function () { onAddPersonClick('member'); });
    }
    if (btnRefreshOwners && btnRefreshOwners.dataset.giBound !== '1') {
        btnRefreshOwners.dataset.giBound = '1';
        btnRefreshOwners.addEventListener('click', function () { refreshOwnersList(false); });
    }
    if (btnRefreshMembers && btnRefreshMembers.dataset.giBound !== '1') {
        btnRefreshMembers.dataset.giBound = '1';
        btnRefreshMembers.addEventListener('click', function () { refreshMembersList(false); });
    }
    [inpOwner, inpMember].forEach(function (inp) {
        if (!inp || inp.dataset.giBound === '1') return;
        inp.dataset.giBound = '1';
        inp.addEventListener('keydown', function (e) {
            if (e.key === 'Enter') {
                e.preventDefault();
                onAddPersonClick(inp.id === 'giOwnerAddInput' ? 'owner' : 'member');
            }
        });
    });

    function onSuggestionClick(kind) {
        return async function (e) {
            const btn = e.target.closest('button[data-gi-pick]');
            if (!btn) return;
            const id = btn.getAttribute('data-gi-pick') || '';
            const label = btn.getAttribute('data-gi-pick-label') || id;
            if (!id) return;
            btn.disabled = true;
            const ok = await addPrincipalToGroup(kind, id, label);
            if (ok) {
                const inp = document.getElementById(kind === 'owner' ? 'giOwnerAddInput' : 'giMemberAddInput');
                if (inp) inp.value = '';
                closeSuggestions(kind);
                if (kind === 'owner') await refreshOwnersList(true);
                else await refreshMembersList(true);
            } else {
                btn.disabled = false;
            }
        };
    }
    if (suggOwner && suggOwner.dataset.giBound !== '1') {
        suggOwner.dataset.giBound = '1';
        suggOwner.addEventListener('click', onSuggestionClick('owner'));
    }
    if (suggMember && suggMember.dataset.giBound !== '1') {
        suggMember.dataset.giBound = '1';
        suggMember.addEventListener('click', onSuggestionClick('member'));
    }

    function onTableClick(kind) {
        return async function (e) {
            const btn = e.target.closest('button[data-gi-act="remove"]');
            if (!btn) return;
            const id = btn.getAttribute('data-gi-id') || '';
            const label = btn.getAttribute('data-gi-label') || id;
            if (!id) return;
            btn.disabled = true;
            const ok = await removePrincipalFromGroup(kind, id, label);
            if (ok) {
                if (kind === 'owner') await refreshOwnersList(true);
                else await refreshMembersList(true);
            } else {
                btn.disabled = false;
            }
        };
    }
    if (tbodyOwners && tbodyOwners.dataset.giBound !== '1') {
        tbodyOwners.dataset.giBound = '1';
        tbodyOwners.addEventListener('click', onTableClick('owner'));
    }
    if (tbodyMembers && tbodyMembers.dataset.giBound !== '1') {
        tbodyMembers.dataset.giBound = '1';
        tbodyMembers.addEventListener('click', onTableClick('member'));
    }
}

// ----- Schritt-Navigation -----

function giStepNum(el) {
    const raw = el.getAttribute('data-gi-step');
    const n = parseFloat(String(raw || '').trim());
    return Number.isFinite(n) ? n : NaN;
}

function goToGiStep(step) {
    giCurrentStep = step;
    document.querySelectorAll('.gi-step-content').forEach(function (el) {
        el.classList.toggle('active', giStepNum(el) === step);
    });
    document.querySelectorAll('.gi-steps .step').forEach(function (el) {
        const s = giStepNum(el);
        el.classList.toggle('active', s === step);
        el.classList.toggle('completed', s < step);
    });
    if (typeof window.ms365ApplyStepProgress === 'function') {
        window.ms365ApplyStepProgress(document.querySelector('.gi-steps'), step, [1, 2]);
    }
    if (step === 2) {
        loadGroupDetailsIntoStep2().catch(function () {});
        refreshStatusIntoUi().catch(function () {});
        refreshOwnersList(false).catch(function () {});
        refreshMembersList(false).catch(function () {});
    }
}

// ----- Aktionen -----

async function onCreateGroupClick() {
    const nameInp = document.getElementById('giInputDisplayName');
    const name = nameInp && nameInp.value ? nameInp.value.trim() : '';
    if (!name) {
        toast('Bitte einen Anzeigenamen für die Sicherheitsgruppe eintragen.');
        return;
    }
    const btn = document.getElementById('giBtnCreateGroup');
    if (btn) btn.disabled = true;
    clearLog();
    appendLog('Rollenfähige Sicherheitsgruppe wird angelegt …', '');
    try {
        const token = await getGraphToken();
        const g = await createRoleAssignableSecurityGroup(token, name);
        setResolvedGroup(g.id, g.displayName || name);
        appendLog(
            'Sicherheitsgruppe erstellt: ' + (g.displayName || name) + ' (' + g.id + ').',
            'ok'
        );
        try {
            const me = await graphJson('GET', '/me?$select=id', token, undefined);
            if (me && me.id) {
                try {
                    await addGroupOwner(token, g.id, me.id);
                    appendLog('Sie wurden als Besitzer:in der neuen Gruppe gesetzt.', 'ok');
                } catch (eo) {
                    if (!/already exist|conflict|409/i.test(String(eo.message || eo))) {
                        appendLog('Hinweis: Besitzer-Eintrag konnte nicht gesetzt werden: ' + (eo.message || eo), 'warn');
                    }
                }
            }
        } catch {
            /* /me optional */
        }
        toast('Rollenfähige Sicherheitsgruppe angelegt.');
    } catch (e) {
        appendLog('Fehler: ' + (e.message || e), 'err');
        toast('Fehler: ' + (e.message || e));
    } finally {
        if (btn) btn.disabled = false;
    }
}

async function onFindGroupClick() {
    const oidInp = document.getElementById('giInputObjectId');
    const nameInp = document.getElementById('giInputDisplayName');
    const rawId = oidInp && oidInp.value ? oidInp.value.trim() : '';
    const name = nameInp && nameInp.value ? nameInp.value.trim() : '';

    const btn = document.getElementById('giBtnFindGroup');
    if (btn) btn.disabled = true;
    clearLog();
    try {
        const token = await getGraphToken();
        if (guidLooksValid(rawId)) {
            const g = await getGroupById(token, rawId);
            if (!g.securityEnabled) {
                throw new Error(
                    'Diese Gruppe ist keine Sicherheitsgruppe (securityEnabled=false). Bitte eine Sicherheits- oder mail-aktivierte Sicherheitsgruppe verwenden.'
                );
            }
            setResolvedGroup(g.id, g.displayName);
            appendLog('Gruppe per Object-ID geladen: ' + g.displayName + ' (' + g.id + ').', 'ok');
            if (!g.isAssignableToRole) {
                appendLog(
                    'Hinweis: Gruppe ist nicht rollenfähig – Direktzuweisung der Rolle scheitert. Fallback in Schritt 2 verwenden (Mitglieder einzeln zur Rolle hinzufügen).',
                    'warn'
                );
            }
            toast('Gruppe gefunden.');
        } else if (name) {
            const list = await findSecurityGroupByDisplayName(token, name);
            if (!list.length) {
                throw new Error(
                    'Keine Sicherheitsgruppe mit diesem Anzeigenamen gefunden (securityEnabled=true).'
                );
            }
            if (list.length > 1) {
                appendLog('Mehrere Treffer – bitte Object-ID aus dem Admin Center kopieren.', 'warn');
            }
            const g = list[0];
            setResolvedGroup(g.id, g.displayName);
            appendLog('Gruppe gefunden: ' + g.displayName + ' (' + g.id + ').', 'ok');
            toast('Gruppe gefunden.');
        } else {
            throw new Error('Bitte gültige Object-ID (GUID) eintragen oder Anzeigenamen zur Suche.');
        }
    } catch (e) {
        appendLog('Fehler: ' + (e.message || e), 'err');
        toast('Fehler: ' + (e.message || e));
    } finally {
        if (btn) btn.disabled = false;
    }
}

async function onApplyClick() {
    const hid = document.getElementById('giResolvedObjectId');
    const gid = hid && hid.value ? hid.value.trim() : '';
    if (!guidLooksValid(gid)) {
        toast('Bitte zuerst in Schritt 1 eine Sicherheitsgruppe wählen oder anlegen.');
        return;
    }
    const btn = document.getElementById('giBtnApply');
    if (btn) btn.disabled = true;
    clearLog();
    appendLog('Wende Richtlinie und Rollenzuweisung an …', '');
    try {
        const token = await getGraphToken();
        const g = await getGroupById(token, gid);
        if (!g.isAssignableToRole) {
            throw new Error(
                'Gruppe ist nicht rollenfähig (isAssignableToRole=false). Bitte den Fallback unten verwenden (Mitglieder einzeln zur Rolle hinzufügen) oder eine neue rollenfähige Sicherheitsgruppe anlegen.'
            );
        }

        const role = await getGuestInviterRoleObject(token);
        appendLog('Rolle „Gasteinladende:r“ identifiziert (' + role.id + ').', '');

        // Member ($ref) hinzufügen – ignore "already exists"
        try {
            await addRoleMember(token, role, gid);
            appendLog('Gruppe der Rolle „Gasteinladende:r“ zugewiesen.', 'ok');
        } catch (e) {
            if (/added object references already exist|already exists|conflict|409/i.test(String(e.message))) {
                appendLog('Gruppe ist bereits Trägerin der Rolle (übersprungen).', '');
            } else {
                throw e;
            }
        }

        await setAllowInvitesFrom(token, 'adminsAndGuestInviters');
        appendLog(
            'authorizationPolicy.allowInvitesFrom auf „adminsAndGuestInviters“ gesetzt.',
            'ok'
        );

        toast('Richtlinie und Rolle gespeichert.');
        await refreshStatusIntoUi(token);
    } catch (e) {
        appendLog('Fehler: ' + (e.message || e), 'err');
        toast('Fehler: ' + (e.message || e));
    } finally {
        if (btn) btn.disabled = false;
    }
}

async function onSyncMembersClick() {
    const hid = document.getElementById('giResolvedObjectId');
    const gid = hid && hid.value ? hid.value.trim() : '';
    if (!guidLooksValid(gid)) {
        toast('Bitte zuerst in Schritt 1 eine Sicherheitsgruppe wählen oder anlegen.');
        return;
    }
    const cleanup = !!document.getElementById('giChkCleanup') && document.getElementById('giChkCleanup').checked;
    const btn = document.getElementById('giBtnSyncMembers');
    if (btn) btn.disabled = true;
    clearLog();
    appendLog('Synchronisiere Gruppenmitglieder mit der Rolle „Gasteinladende:r“ …', '');
    try {
        const token = await getGraphToken();

        // Auch hier Policy setzen, damit der gewünschte Effekt eintritt.
        await setAllowInvitesFrom(token, 'adminsAndGuestInviters');
        appendLog('authorizationPolicy.allowInvitesFrom = adminsAndGuestInviters.', 'ok');

        const role = await getGuestInviterRoleObject(token);
        appendLog('Rolle „Gasteinladende:r“ identifiziert (' + role.id + ').', '');

        const [members, roleMembers] = await Promise.all([
            getGroupMembers(token, gid),
            getRoleMembers(token, role)
        ]);
        const memberIds = new Set(members.map(function (m) { return m.id; }).filter(Boolean));
        const roleIds = new Set(roleMembers.map(function (m) { return m.id; }).filter(Boolean));

        // Nur Benutzer (keine Gruppen-Mitglieder) hinzufügen, weil Verschachtelung nicht funktioniert,
        // wenn die Quelle keine rollenfähige Gruppe ist.
        const memberUsers = members.filter(function (m) {
            const t = String(m['@odata.type'] || '').toLowerCase();
            return t.indexOf('user') !== -1;
        });
        const toAdd = memberUsers.filter(function (m) { return !roleIds.has(m.id); });
        const toRemove = cleanup
            ? roleMembers.filter(function (m) {
                  const t = String(m['@odata.type'] || '').toLowerCase();
                  if (t.indexOf('user') === -1) return false; // Nur Benutzer entfernen
                  return !memberIds.has(m.id);
              })
            : [];

        appendLog('Gruppe: ' + members.length + ' Mitglieder · Rolle: ' + roleMembers.length + ' Träger:innen.', '');
        appendLog('Hinzuzufügen: ' + toAdd.length + (cleanup ? ' · zu entfernen: ' + toRemove.length : ''), '');

        let added = 0;
        for (const m of toAdd) {
            try {
                await addRoleMember(token, role, m.id);
                added++;
                appendLog('+ ' + (m.displayName || m.userPrincipalName || m.id), 'ok');
            } catch (e) {
                if (/already exist|conflict|409/i.test(String(e.message))) {
                    appendLog('= ' + (m.displayName || m.userPrincipalName || m.id) + ' (bereits Träger:in)', '');
                } else {
                    appendLog('Fehler beim Hinzufügen ' + (m.displayName || m.id) + ': ' + e.message, 'err');
                }
            }
        }

        let removed = 0;
        for (const m of toRemove) {
            try {
                await removeRoleMember(token, role, m.id);
                removed++;
                appendLog('− ' + (m.displayName || m.userPrincipalName || m.id), 'warn');
            } catch (e) {
                appendLog('Fehler beim Entfernen ' + (m.displayName || m.id) + ': ' + e.message, 'err');
            }
        }

        appendLog('Synchronisierung fertig. Hinzugefügt: ' + added + ', entfernt: ' + removed + '.', 'ok');
        toast('Mitglieder synchronisiert.');
        await refreshStatusIntoUi(token);
    } catch (e) {
        appendLog('Fehler: ' + (e.message || e), 'err');
        toast('Fehler: ' + (e.message || e));
    } finally {
        if (btn) btn.disabled = false;
    }
}

async function onResetPolicyClick() {
    const ok =
        typeof window.ms365AppDialogConfirm === 'function'
            ? await window.ms365AppDialogConfirm(
                  'Wirklich allowInvitesFrom auf „everyone“ zurücksetzen? Damit dürfen wieder alle internen Benutzer:innen Gäste einladen.',
                  { title: 'Einstellung zurücksetzen', okText: 'Zurücksetzen', danger: true }
              )
            : window.confirm(
                  'Wirklich allowInvitesFrom auf „everyone“ zurücksetzen? Damit dürfen wieder alle internen Benutzer:innen Gäste einladen.'
              );
    if (!ok) return;
    const btn = document.getElementById('giBtnResetPolicy');
    if (btn) btn.disabled = true;
    clearLog();
    try {
        const token = await getGraphToken();
        await setAllowInvitesFrom(token, 'everyone');
        appendLog('authorizationPolicy.allowInvitesFrom = everyone (Standard).', 'ok');
        toast('Einstellung zurückgesetzt.');
        await refreshStatusIntoUi(token);
    } catch (e) {
        appendLog('Fehler: ' + (e.message || e), 'err');
        toast('Fehler: ' + (e.message || e));
    } finally {
        if (btn) btn.disabled = false;
    }
}

async function onRemoveGroupFromRoleClick() {
    const hid = document.getElementById('giResolvedObjectId');
    const gid = hid && hid.value ? hid.value.trim() : '';
    if (!guidLooksValid(gid)) {
        toast('Keine Gruppe gewählt.');
        return;
    }
    const ok =
        typeof window.ms365AppDialogConfirm === 'function'
            ? await window.ms365AppDialogConfirm(
                  'Soll die ausgewählte Gruppe aus der Rolle „Gasteinladende:r“ entfernt werden?',
                  { title: 'Aus Rolle entfernen', okText: 'Entfernen', danger: true }
              )
            : window.confirm(
                  'Soll die ausgewählte Gruppe aus der Rolle „Gasteinladende:r“ entfernt werden?'
              );
    if (!ok) return;
    const btn = document.getElementById('giBtnRemoveGroupFromRole');
    if (btn) btn.disabled = true;
    clearLog();
    try {
        const token = await getGraphToken();
        const role = await getGuestInviterRoleObject(token);
        await removeRoleMember(token, role, gid);
        appendLog('Gruppe wurde aus der Rolle entfernt.', 'ok');
        toast('Gruppe entfernt.');
        await refreshStatusIntoUi(token);
    } catch (e) {
        appendLog('Fehler: ' + (e.message || e), 'err');
        toast('Fehler: ' + (e.message || e));
    } finally {
        if (btn) btn.disabled = false;
    }
}

async function refreshStatusIntoUi(tokenIn) {
    const box = document.getElementById('giStatusPre');
    const token = tokenIn || (await getGraphToken());
    let text = '';
    try {
        const ap = await getAuthorizationPolicy(token);
        text += 'authorizationPolicy.allowInvitesFrom: ' + (ap.allowInvitesFrom || '(unbekannt)') + '\n';
    } catch (e) {
        text += 'authorizationPolicy: Lesefehler – ' + (e.message || e) + '\n';
    }

    try {
        const role = await getGuestInviterRoleObject(token);
        text += 'Rolle „Gasteinladende:r“: ' + role.displayName + ' (' + role.id + ')\n';
        const members = await getRoleMembers(token, role);
        text += 'Rollenträger:innen (gesamt): ' + members.length + '\n';
        const sample = members.slice(0, 20).map(function (m) {
            const t = String(m['@odata.type'] || '').replace('#microsoft.graph.', '');
            return '  · [' + t + '] ' + (m.displayName || '(ohne Name)') + (m.userPrincipalName ? ' <' + m.userPrincipalName + '>' : '');
        });
        if (sample.length) text += sample.join('\n') + (members.length > sample.length ? '\n  …' : '') + '\n';

        const hid = document.getElementById('giResolvedObjectId');
        const gid = hid && hid.value ? hid.value.trim() : '';
        if (guidLooksValid(gid)) {
            const inRole = members.some(function (m) { return m.id === gid; });
            text +=
                'Aktuelle Gruppe (' +
                gid +
                ') als Rollenträgerin: ' +
                (inRole ? 'JA' : 'NEIN') +
                '\n';
        }
    } catch (e) {
        text += 'Rolle: Lesefehler – ' + (e.message || e) + '\n';
    }

    if (box) box.textContent = text || '(kein Status)';
}

async function onRefreshStatusClick() {
    const btn = document.getElementById('giBtnRefreshStatus');
    if (btn) btn.disabled = true;
    clearLog();
    try {
        await refreshStatusIntoUi();
        appendLog('Status gelesen.', 'ok');
    } catch (e) {
        appendLog('Fehler: ' + (e.message || e), 'err');
        toast('Fehler: ' + (e.message || e));
    } finally {
        if (btn) btn.disabled = false;
    }
}

async function onLoginClick() {
    const btn = document.getElementById('giBtnLogin');
    if (btn) btn.disabled = true;
    try {
        await getGraphToken();
        toast('Angemeldet – Sie können den Status lesen oder die Einschränkung setzen.');
        appendLog(
            'Anmeldung OK (benötigt Policy.ReadWrite.Authorization, RoleManagement.ReadWrite.Directory, Group.ReadWrite.All).',
            'ok'
        );
    } catch (e) {
        toast('Anmeldung: ' + (e.message || e));
    } finally {
        if (btn) btn.disabled = false;
    }
}

// ----- Zwischenstand -----

function saveState() {
    try {
        const hid = document.getElementById('giResolvedObjectId');
        const dn = document.getElementById('giInputDisplayName');
        const oid = document.getElementById('giInputObjectId');
        const state = {
            giCurrentStep: giCurrentStep,
            groupDisplayName: dn ? dn.value : '',
            groupObjectIdRaw: oid ? oid.value : '',
            resolvedObjectId: hid ? hid.value : ''
        };
        localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
        toast('Gast-Einlader: Zwischenstand gespeichert.');
    } catch (e) {
        toast('Speichern fehlgeschlagen: ' + (e.message || e));
    }
}

function loadState() {
    try {
        const raw = localStorage.getItem(STORAGE_KEY);
        if (!raw) {
            toast('Kein gespeicherter Stand.');
            return;
        }
        const state = JSON.parse(raw);
        const dn = document.getElementById('giInputDisplayName');
        const oid = document.getElementById('giInputObjectId');
        if (dn && state.groupDisplayName !== undefined) dn.value = state.groupDisplayName;
        if (oid && state.groupObjectIdRaw !== undefined) oid.value = state.groupObjectIdRaw;
        if (state.resolvedObjectId && guidLooksValid(state.resolvedObjectId)) {
            setResolvedGroup(state.resolvedObjectId, '');
        }
        if (typeof state.giCurrentStep === 'number' && state.giCurrentStep >= 1) {
            let st = state.giCurrentStep;
            if (st > 2) st = 2;
            goToGiStep(st);
        }
        toast('Gast-Einlader: Stand geladen.');
    } catch (e) {
        toast('Laden fehlgeschlagen: ' + (e.message || e));
    }
}

function clearState() {
    try {
        localStorage.removeItem(STORAGE_KEY);
        toast('Gast-Einlader: lokaler Speicher geleert.');
    } catch (e) {
        toast('Fehler: ' + (e.message || e));
    }
}

function init() {
    goToGiStep(1);
    const dn = document.getElementById('giInputDisplayName');
    if (dn && !dn.value) dn.value = 'GastEinlader';

    const $ = function (id) { return document.getElementById(id); };
    $('giBtnCreateGroup') && $('giBtnCreateGroup').addEventListener('click', onCreateGroupClick);
    $('giBtnFindGroup') && $('giBtnFindGroup').addEventListener('click', onFindGroupClick);
    $('giBtnNext1') &&
        $('giBtnNext1').addEventListener('click', function () {
            const hid = document.getElementById('giResolvedObjectId');
            const oidInp = document.getElementById('giInputObjectId');
            const raw = oidInp && oidInp.value ? oidInp.value.trim() : '';
            if (guidLooksValid(raw)) {
                setResolvedGroup(raw, '');
            }
            const gid = hid && hid.value ? hid.value.trim() : '';
            if (!guidLooksValid(gid)) {
                toast(
                    'Bitte zuerst eine Sicherheitsgruppe anlegen, suchen oder eine gültige Object-ID (GUID) eintragen.'
                );
                return;
            }
            goToGiStep(2);
        });
    $('giBtnBack2') &&
        $('giBtnBack2').addEventListener('click', function () {
            goToGiStep(1);
        });

    $('giBtnLogin') && $('giBtnLogin').addEventListener('click', onLoginClick);
    $('giBtnRefreshStatus') && $('giBtnRefreshStatus').addEventListener('click', onRefreshStatusClick);
    $('giBtnApply') && $('giBtnApply').addEventListener('click', onApplyClick);
    $('giBtnSyncMembers') && $('giBtnSyncMembers').addEventListener('click', onSyncMembersClick);
    $('giBtnResetPolicy') && $('giBtnResetPolicy').addEventListener('click', onResetPolicyClick);
    $('giBtnRemoveGroupFromRole') &&
        $('giBtnRemoveGroupFromRole').addEventListener('click', onRemoveGroupFromRoleClick);

    $('btnSaveState') && $('btnSaveState').addEventListener('click', saveState);
    $('btnLoadState') && $('btnLoadState').addEventListener('click', loadState);
    $('btnClearStorage') && $('btnClearStorage').addEventListener('click', clearState);

    bindPeopleManagement();
}

window.ms365SaveGastEinlader = saveState;
window.ms365LoadGastEinlader = loadState;
window.ms365ClearGastEinlader = clearState;
window.ms365GiGraphLogin = onLoginClick;
window.ms365GiGraphApply = onApplyClick;

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
} else {
    init();
}
