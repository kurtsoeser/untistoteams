import js from '@eslint/js';
import globals from 'globals';

const vitestGlobals = {
    describe: 'readonly',
    it: 'readonly',
    expect: 'readonly',
    beforeEach: 'readonly',
    afterEach: 'readonly',
    beforeAll: 'readonly',
    afterAll: 'readonly',
    vi: 'readonly'
};

export default [
    {
        ignores: ['dist/**', 'node_modules/**', 'coverage/**']
    },
    js.configs.recommended,
    {
        languageOptions: {
            ecmaVersion: 2022,
            sourceType: 'module',
            globals: {
                ...globals.browser,
                ...globals.node,
                /** SheetJS aus CDN (xlsx.full.min.js), mehrere Tools */
                XLSX: 'readonly'
            }
        },
        rules: {
            'no-unused-vars': [
                'warn',
                {
                    argsIgnorePattern: '^_',
                    varsIgnorePattern: '^_',
                    caughtErrors: 'none'
                }
            ],
            'no-empty': ['warn', { allowEmptyCatch: true }],
            'no-irregular-whitespace': 'off',
            'no-useless-escape': 'off'
        }
    },
    {
        files: ['tests/**/*.{js,mjs}', '**/*.test.{js,mjs}'],
        languageOptions: {
            globals: {
                ...globals.node,
                ...vitestGlobals
            }
        }
    }
];
