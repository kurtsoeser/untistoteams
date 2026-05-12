import { describe, it, expect } from 'vitest';
import {
    normStr,
    normCode,
    normEmail,
    normHeaderKey,
    escapeHtml,
    attrEscape,
    compareDe
} from '../src/shared/utils/strings.js';
import { safeJsonParse } from '../src/shared/utils/json.js';
import { loadJson, saveJson, removeKey } from '../src/shared/utils/storage.js';
import { dlgAlert, dlgConfirm, dlgPrompt } from '../src/shared/utils/dialog.js';
import { getEl } from '../src/shared/utils/dom.js';

describe('strings.js', () => {
    it('normStr: trim, tolerant gegen null/undefined', () => {
        expect(normStr('  x  ')).toBe('x');
        expect(normStr(null)).toBe('');
        expect(normStr(undefined)).toBe('');
        expect(normStr(42)).toBe('42');
    });

    it('normCode: trim + UPPER', () => {
        expect(normCode('  ab12  ')).toBe('AB12');
        expect(normCode(null)).toBe('');
    });

    it('normEmail: trim + lower', () => {
        expect(normEmail('  Foo@Bar.AT  ')).toBe('foo@bar.at');
    });

    it('normHeaderKey: trim, lower, Umlaute auflösen, Sonderzeichen weg', () => {
        expect(normHeaderKey(' Schüler-Nr ')).toBe('schuelernr');
        expect(normHeaderKey('Größe (in cm)')).toBe('groesseincm');
        expect(normHeaderKey('ABC 123')).toBe('abc123');
        expect(normHeaderKey('Maß')).toBe('mass');
    });

    it('escapeHtml: ersetzt &, <, >, ", \'', () => {
        expect(escapeHtml('<b>"a"&\'b\'</b>')).toBe(
            '&lt;b&gt;&quot;a&quot;&amp;&#039;b&#039;&lt;/b&gt;'
        );
        expect(escapeHtml(null)).toBe('');
    });

    it('attrEscape: nur & und "', () => {
        expect(attrEscape('a"b&c')).toBe('a&quot;b&amp;c');
        expect(attrEscape(null)).toBe('');
    });

    it('compareDe: deutsche Sortierung, case-insensitive', () => {
        const arr = ['Banane', 'apfel', 'Äpfel'];
        const sorted = [...arr].sort(compareDe);
        expect(sorted[0].toLowerCase()).toMatch(/^(äpfel|apfel)$/);
        expect(sorted[sorted.length - 1]).toBe('Banane');
    });
});

describe('json.js', () => {
    it('safeJsonParse: gültiges JSON', () => {
        expect(safeJsonParse('{"a":1}')).toEqual({ a: 1 });
        expect(safeJsonParse('[1,2,3]')).toEqual([1, 2, 3]);
    });

    it('safeJsonParse: ungültiges JSON → fallback', () => {
        expect(safeJsonParse('not json')).toBeNull();
        expect(safeJsonParse('not json', { x: 1 })).toEqual({ x: 1 });
        expect(safeJsonParse(null, [])).toBeNull();
        // null wird zu "null" als String → ergibt null durch JSON.parse,
        // aber das ist gültiges JSON, daher null erwartet (nicht fallback).
    });
});

describe('storage.js (ohne window)', () => {
    it('loadJson liefert fallback, wenn kein window', () => {
        expect(loadJson('any-key', { a: 1 })).toEqual({ a: 1 });
    });

    it('saveJson liefert false, wenn kein window', () => {
        expect(saveJson('any-key', { a: 1 })).toBe(false);
    });

    it('removeKey liefert false, wenn kein window', () => {
        expect(removeKey('any-key')).toBe(false);
    });
});

describe('storage.js (mit Mock-localStorage)', () => {
    function withMockWindow(fn) {
        const mem = new Map();
        const store = {
            getItem: (k) => (mem.has(k) ? mem.get(k) : null),
            setItem: (k, v) => mem.set(k, String(v)),
            removeItem: (k) => mem.delete(k)
        };
        const prev = globalThis.window;
        globalThis.window = { localStorage: store };
        try {
            fn(mem);
        } finally {
            globalThis.window = prev;
        }
    }

    it('saveJson + loadJson Roundtrip', () => {
        withMockWindow(() => {
            expect(saveJson('k', { hello: 'welt' })).toBe(true);
            expect(loadJson('k')).toEqual({ hello: 'welt' });
        });
    });

    it('loadJson liefert fallback für fehlenden Key', () => {
        withMockWindow(() => {
            expect(loadJson('missing', 'def')).toBe('def');
        });
    });

    it('removeKey entfernt den Eintrag', () => {
        withMockWindow((mem) => {
            saveJson('k', 1);
            expect(mem.has('k')).toBe(true);
            expect(removeKey('k')).toBe(true);
            expect(mem.has('k')).toBe(false);
        });
    });
});

describe('dialog.js (ohne window)', () => {
    it('dlgAlert resolved auch ohne window', async () => {
        await expect(dlgAlert('x')).resolves.toBeUndefined();
    });

    it('dlgConfirm resolved zu false ohne window', async () => {
        await expect(dlgConfirm('x')).resolves.toBe(false);
    });

    it('dlgPrompt resolved zu null ohne window', async () => {
        await expect(dlgPrompt('x', 'def')).resolves.toBeNull();
    });
});

describe('dom.js (ohne document)', () => {
    it('getEl liefert null ohne document', () => {
        expect(getEl('any')).toBeNull();
    });
});
