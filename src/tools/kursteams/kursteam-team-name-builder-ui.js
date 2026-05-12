
const KT = window.ms365KursteamTeamNames;
window.ms365AssertModules({ KT }, 'kursteam-team-name-builder-ui.js');

/**
 * @param {object} ns window.ms365Kursteam
 */
function mount(ns) {
    function getPatternFromBuilder() {
        const zone = document.getElementById('teamNameBuilder');
        if (!zone) return KT.normalizePattern(ns.teamNamePattern);
        const tokens = [];
        zone.querySelectorAll('[data-token-type]').forEach((el) => {
            const type = String(el.getAttribute('data-token-type') || '');
            if (type === 'text') tokens.push({ type: 'text', value: String(el.getAttribute('data-token-value') || '') });
            else tokens.push({ type });
        });
        return KT.normalizePattern(tokens);
    }

    function setPreviewFromPattern(pattern) {
        const el = document.getElementById('teamNamePreview');
        if (!el) return;
        const yearPrefix = document.getElementById('yearPrefix')?.value || 'SJ26';
        const preview = KT.buildTeamNameFromPattern(pattern, { yearPrefix, klasse: '1AK', fach: 'D', gruppe: 'G1' });
        el.textContent = 'Vorschau: ' + preview;
    }

    function wireBuilderDnD(zone) {
        let dragEl = null;
        zone.addEventListener('dragstart', (e) => {
            const target = e.target && e.target.closest ? e.target.closest('.name-chip') : null;
            if (!target) return;
            dragEl = target;
            target.classList.add('dragging');
            e.dataTransfer.effectAllowed = 'move';
        });
        zone.addEventListener('dragend', () => {
            if (dragEl) dragEl.classList.remove('dragging');
            dragEl = null;
        });
        zone.addEventListener('dragover', (e) => {
            e.preventDefault();
            const over = e.target && e.target.closest ? e.target.closest('.name-chip') : null;
            if (!dragEl || !over || over === dragEl) return;
            const rect = over.getBoundingClientRect();
            const after = e.clientX > rect.left + rect.width / 2;
            if (after) over.after(dragEl);
            else over.before(dragEl);
        });
        zone.addEventListener('drop', (e) => {
            e.preventDefault();
            ns.teamNamePattern = getPatternFromBuilder();
            setPreviewFromPattern(ns.teamNamePattern);
        });
    }

    function addChip(zone, token) {
        const chip = document.createElement('span');
        chip.className = 'name-chip';
        chip.draggable = true;
        chip.setAttribute('data-token-type', token.type);
        if (token.type === 'text') chip.setAttribute('data-token-value', String(token.value ?? ''));

        const txt = document.createElement('span');
        if (token.type === 'text') {
            const v = String(token.value ?? '');
            txt.textContent = v === '' ? '(leer)' : v;
        } else {
            txt.textContent = KT.tokenLabel(token);
        }

        const x = document.createElement('button');
        x.type = 'button';
        x.className = 'chip-x';
        x.textContent = '✕';
        x.title = 'Baustein entfernen';
        x.addEventListener('click', () => {
            chip.remove();
            ns.teamNamePattern = getPatternFromBuilder();
            setPreviewFromPattern(ns.teamNamePattern);
        });

        chip.append(txt, x);
        zone.appendChild(chip);
    }

    let nameBuilderWired = false;
    function wireNameBuilderOnce() {
        if (nameBuilderWired) return;
        nameBuilderWired = true;

        const zone = document.getElementById('teamNameBuilder');
        if (!zone) return;

        wireBuilderDnD(zone);

        const btnSep = document.getElementById('teamNameAddSep');
        if (btnSep) {
            btnSep.addEventListener('click', () => {
                const v = document.getElementById('teamNameSepValue')?.value;
                addChip(zone, { type: 'text', value: String(v ?? '') });
                ns.teamNamePattern = getPatternFromBuilder();
                setPreviewFromPattern(ns.teamNamePattern);
            });
        }
        const btnText = document.getElementById('teamNameAddText');
        if (btnText) {
            btnText.addEventListener('click', () => {
                const v = document.getElementById('teamNameTextValue')?.value;
                addChip(zone, { type: 'text', value: String(v ?? '') });
                ns.teamNamePattern = getPatternFromBuilder();
                setPreviewFromPattern(ns.teamNamePattern);
            });
        }
        const btnReset = document.getElementById('teamNameResetDefault');
        if (btnReset) {
            btnReset.addEventListener('click', () => {
                ns.teamNamePattern = KT.defaultTeamNamePattern();
                ns.renderTeamNameBuilder();
            });
        }

        document.querySelectorAll('[data-teamname-token]').forEach((btn) => {
            btn.addEventListener('click', () => {
                const type = String(btn.getAttribute('data-teamname-token') || '').trim();
                if (!type) return;
                addChip(zone, { type });
                ns.teamNamePattern = getPatternFromBuilder();
                setPreviewFromPattern(ns.teamNamePattern);
            });
        });

        const yp = document.getElementById('yearPrefix');
        if (yp) yp.addEventListener('input', () => setPreviewFromPattern(getPatternFromBuilder()));
    }

    ns.getPatternFromBuilder = getPatternFromBuilder;
    ns.renderTeamNameBuilder = function renderTeamNameBuilder() {
        const zone = document.getElementById('teamNameBuilder');
        if (!zone) return;
        wireNameBuilderOnce();
        const pattern = KT.normalizePattern(ns.teamNamePattern);
        zone.replaceChildren();
        pattern.forEach((t) => addChip(zone, t));
        setPreviewFromPattern(pattern);
    };
}

window.ms365KursteamTeamNameBuilderUI = {
    mount
};
