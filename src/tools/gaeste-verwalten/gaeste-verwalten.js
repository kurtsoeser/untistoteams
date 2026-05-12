(function () {
    'use strict';

    /**
     * Tab-Steuerung für das zusammengeführte Werkzeug „Gäste verwalten“.
     * Die fachliche Logik der vier Unterbereiche stammt unverändert aus
     *   - src/tools/gast-zugaenge/gast-zugaenge.js  (Tabs 1–3: Teams, Audit, Tenant-Gäste)
     *   - src/tools/gast-einlader/gast-einlader.js  (Tab 4: Berechtigung)
     * Diese Datei kümmert sich nur um das Umschalten der Panels, das
     * Persistieren des zuletzt aktiven Tabs (Browser-Tab-Sitzung) und um
     * Deep-Links via URL-Hash (z. B. „…/gaeste-verwalten.html#einlader“).
     */

    const STORAGE_KEY = 'ms365-gaeste-verwalten-active-tab-v1';
    const VALID_TABS = ['tenant', 'teams', 'audit', 'einlader'];
    const DEFAULT_TAB = 'tenant';

    function normalizeTab(value) {
        const v = String(value || '').trim().toLowerCase();
        return VALID_TABS.indexOf(v) === -1 ? '' : v;
    }

    function readHashTab() {
        const raw = String(window.location.hash || '').replace(/^#/, '');
        if (!raw) return '';
        const params = new URLSearchParams(raw);
        const candidate = params.get('tab') || params.get('bereich') || raw;
        return normalizeTab(candidate);
    }

    function readStoredTab() {
        try {
            return normalizeTab(sessionStorage.getItem(STORAGE_KEY));
        } catch {
            return '';
        }
    }

    function storeTab(name) {
        try {
            sessionStorage.setItem(STORAGE_KEY, name);
        } catch {
            /* privates Fenster, Quota – egal */
        }
    }

    function activate(name) {
        const target = normalizeTab(name) || DEFAULT_TAB;
        const tabs = document.querySelectorAll('.gv-tab[data-gv-tab]');
        const panels = document.querySelectorAll('.gv-panel[role="tabpanel"]');
        tabs.forEach(function (btn) {
            const selected = btn.getAttribute('data-gv-tab') === target;
            btn.setAttribute('aria-selected', selected ? 'true' : 'false');
            btn.tabIndex = selected ? 0 : -1;
        });
        panels.forEach(function (panel) {
            const id = panel.getAttribute('id') || '';
            const matches = id === 'gv-panel-' + target;
            panel.classList.toggle('is-active', matches);
            if (matches) panel.removeAttribute('hidden');
            else panel.setAttribute('hidden', '');
        });
        storeTab(target);
    }

    function focusTab(name) {
        const btn = document.querySelector('.gv-tab[data-gv-tab="' + name + '"]');
        if (btn && typeof btn.focus === 'function') btn.focus();
    }

    function onKeydown(e) {
        const key = e.key;
        if (key !== 'ArrowLeft' && key !== 'ArrowRight' && key !== 'Home' && key !== 'End') return;
        const list = Array.prototype.slice.call(document.querySelectorAll('.gv-tab[data-gv-tab]'));
        if (!list.length) return;
        const current = document.activeElement;
        const idx = list.indexOf(current);
        if (idx === -1) return;
        e.preventDefault();
        let nextIdx = idx;
        if (key === 'ArrowLeft') nextIdx = (idx - 1 + list.length) % list.length;
        else if (key === 'ArrowRight') nextIdx = (idx + 1) % list.length;
        else if (key === 'Home') nextIdx = 0;
        else if (key === 'End') nextIdx = list.length - 1;
        const next = list[nextIdx];
        const name = next.getAttribute('data-gv-tab') || DEFAULT_TAB;
        activate(name);
        focusTab(name);
    }

    function init() {
        const tabList = document.querySelector('.gv-tablist');
        if (!tabList) return;

        document.querySelectorAll('.gv-tab[data-gv-tab]').forEach(function (btn) {
            btn.addEventListener('click', function () {
                const name = btn.getAttribute('data-gv-tab') || DEFAULT_TAB;
                activate(name);
                focusTab(name);
            });
        });
        tabList.addEventListener('keydown', onKeydown);

        window.addEventListener('hashchange', function () {
            const fromHash = readHashTab();
            if (fromHash) activate(fromHash);
        });

        const initial = readHashTab() || readStoredTab() || DEFAULT_TAB;
        activate(initial);
    }

    window.ms365GaesteVerwaltenSetTab = function (name) {
        activate(name);
    };

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
