
const KS = window.ms365KursteamSubjectFilterLogic;
window.ms365AssertModules({ KS }, 'kursteam-subject-filter-ui.js');

/**
 * Fachfilter-Liste, Suche, Schnellbuttons (Schritt Bereinigung).
 * @param {object} ns window.ms365Kursteam
 */
function mount(ns) {
    ns.parseExcludeSubjectsFromInput = function parseExcludeSubjectsFromInput() {
        const el = document.getElementById('excludeSubjects');
        if (!el) return [];
        return KS.parseExcludeSubjectsFromString(el.value);
    };

    ns.setExcludeSubjectsInput = function setExcludeSubjectsInput(tokens) {
        const el = document.getElementById('excludeSubjects');
        if (!el) return;
        el.value = KS.uniqSortedSubjectTokens(tokens).join(',');
    };

    function collectAvailableSubjectsFromRawData() {
        return KS.collectSubjectsFromRows(ns.rawData);
    }

    function updateSubjectFilterSummary(available, excluded) {
        const el = document.getElementById('subjectFilterSummary');
        if (!el) return;
        el.textContent = KS.subjectFilterSummaryText(available.length, excluded.length);
    }

    function applySearchToSubjectList(query) {
        const q = KS.normalizeSubjectToken(query);
        const list = document.getElementById('subjectFilterList');
        if (!list) return;
        Array.from(list.querySelectorAll('[data-subject]')).forEach((node) => {
            const subj = String(node.getAttribute('data-subject') || '');
            node.style.display = !q || subj.includes(q) ? '' : 'none';
        });
    }

    function wireSubjectFilterEventsOnce() {
        if (wireSubjectFilterEventsOnce._wired) return;
        wireSubjectFilterEventsOnce._wired = true;

        const search = document.getElementById('subjectFilterSearch');
        if (search) {
            search.addEventListener('input', () => applySearchToSubjectList(search.value));
        }

        const btnNone = document.getElementById('subjectFilterExcludeNone');
        if (btnNone) {
            btnNone.addEventListener('click', () => {
                ns.setExcludeSubjectsInput([]);
                ns.refreshSubjectFilterUI();
            });
        }

        const btnAll = document.getElementById('subjectFilterExcludeAll');
        if (btnAll) {
            btnAll.addEventListener('click', () => {
                ns.setExcludeSubjectsInput(collectAvailableSubjectsFromRawData());
                ns.refreshSubjectFilterUI();
            });
        }

        const btnDefault = document.getElementById('subjectFilterResetDefault');
        if (btnDefault) {
            btnDefault.addEventListener('click', () => {
                ns.setExcludeSubjectsInput(['ORD', 'DIR', 'KV']);
                ns.refreshSubjectFilterUI();
            });
        }

        const input = document.getElementById('excludeSubjects');
        if (input) {
            input.addEventListener('input', () => ns.refreshSubjectFilterUI());
        }
    }

    ns.refreshSubjectFilterUI = function refreshSubjectFilterUI() {
        const list = document.getElementById('subjectFilterList');
        const search = document.getElementById('subjectFilterSearch');
        if (!list) return;

        wireSubjectFilterEventsOnce();

        const available = collectAvailableSubjectsFromRawData();
        const excluded = new Set(ns.parseExcludeSubjectsFromInput());

        list.replaceChildren();
        available.forEach((subj) => {
            const label = document.createElement('label');
            label.className = 'subject-filter-item';
            label.setAttribute('data-subject', subj);

            const cb = document.createElement('input');
            cb.type = 'checkbox';
            cb.checked = excluded.has(subj);
            cb.addEventListener('change', () => {
                const current = new Set(ns.parseExcludeSubjectsFromInput());
                if (cb.checked) current.add(subj);
                else current.delete(subj);
                ns.setExcludeSubjectsInput(Array.from(current));
                updateSubjectFilterSummary(available, Array.from(current));
            });

            const text = document.createElement('code');
            text.textContent = subj;

            label.append(cb, text);
            list.appendChild(label);
        });

        updateSubjectFilterSummary(available, Array.from(excluded));
        if (search) applySearchToSubjectList(search.value);
    };
}

window.ms365KursteamSubjectFilterUI = {
    mount
};
