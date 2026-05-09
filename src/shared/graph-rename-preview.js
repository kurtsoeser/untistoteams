(function () {
    'use strict';

    function escapeRe(s) {
        return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }

    /**
     * Erwartetes Muster: "<Präfix> <Stufe><Kürzel>" z. B. "Klasse 1A", "Klasse 10HAK"
     * @param {string} displayName
     * @param {string} prefix z. B. "Klasse"
     * @returns {string|null}
     */
    function computeNewDisplayNamePlusOne(displayName, prefix) {
        const p = String(prefix || '').trim();
        if (!p) return null;
        const re = new RegExp('^' + escapeRe(p) + '\\s+(\\d+)([A-Za-z0-9\\-]*)$', 'i');
        const m = String(displayName || '').trim().match(re);
        if (!m) return null;
        const current = parseInt(m[1], 10);
        if (!isFinite(current)) return null;
        const next = current + 1;
        return p + ' ' + String(next) + (m[2] || '');
    }

    window.ms365GraphRenamePreview = {
        computeNewDisplayNamePlusOne,
        escapeRe
    };
})();
