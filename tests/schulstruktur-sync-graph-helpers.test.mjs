import { describe, it, expect } from 'vitest';
import {
    isInteractionRequired,
    sleep,
    parseTeamsOperationPathFromLocation,
    groupIsTeam,
    graphErrorLooksLikeNotFound,
    personLabel,
    odataEscape,
    directoryObjectRef,
    isGraphDuplicateRefError
} from '../src/tools/schulstruktur-sync/schulstruktur-sync-graph-helpers.js';

describe('isInteractionRequired', () => {
    it('matcht auf MSAL Error-Klasse', () => {
        expect(isInteractionRequired({ name: 'InteractionRequiredAuthError' })).toBe(true);
    });

    it('matcht auf errorCode', () => {
        expect(isInteractionRequired({ errorCode: 'interaction_required' })).toBe(true);
    });

    it('matcht auf Substring in message', () => {
        expect(isInteractionRequired({ message: 'AADSTS50058 interaction_required: ...' })).toBe(true);
    });

    it('liefert false für unverwandte Fehler', () => {
        expect(isInteractionRequired({ message: 'Network error' })).toBe(false);
        expect(isInteractionRequired({ name: 'OtherError' })).toBe(false);
    });

    it('liefert false für null/undefined', () => {
        expect(isInteractionRequired(null)).toBe(false);
        expect(isInteractionRequired(undefined)).toBe(false);
    });
});

describe('sleep', () => {
    it('wartet mindestens die angegebene Zeit', async () => {
        const start = Date.now();
        await sleep(50);
        expect(Date.now() - start).toBeGreaterThanOrEqual(40);
    });

    it('liefert ein Promise', () => {
        expect(sleep(0)).toBeInstanceOf(Promise);
    });
});

describe('parseTeamsOperationPathFromLocation', () => {
    it('REST-Pfad (/teams/{id}/operations/{op})', () => {
        const path = parseTeamsOperationPathFromLocation('/teams/abc-123/operations/op-456');
        expect(path).toBe('/teams/abc-123/operations/op-456');
    });

    it('Full URL wird auf Pfad reduziert (mit v1.0)', () => {
        const path = parseTeamsOperationPathFromLocation(
            'https://graph.microsoft.com/v1.0/teams/abc-123/operations/op-456'
        );
        expect(path).toBe('/teams/abc-123/operations/op-456');
    });

    it('OData-Pfad (teams(\'…\')/operations(\'…\'))', () => {
        const path = parseTeamsOperationPathFromLocation("/teams('abc-123')/operations('op-456')");
        expect(path).toBe('/teams/abc-123/operations/op-456');
    });

    it('liefert null bei null/undefined/leer', () => {
        expect(parseTeamsOperationPathFromLocation(null)).toBeNull();
        expect(parseTeamsOperationPathFromLocation(undefined)).toBeNull();
        expect(parseTeamsOperationPathFromLocation('')).toBeNull();
    });

    it('liefert null bei nicht parsbarer URL', () => {
        expect(parseTeamsOperationPathFromLocation('http://[invalid')).toBeNull();
    });

    it('liefert null bei Pfad, der nicht passt', () => {
        expect(parseTeamsOperationPathFromLocation('/groups/abc/owners')).toBeNull();
    });

    it('Query-String wird abgeschnitten', () => {
        const path = parseTeamsOperationPathFromLocation('/teams/abc/operations/xyz?api-version=1.0');
        expect(path).toBe('/teams/abc/operations/xyz');
    });
});

describe('groupIsTeam', () => {
    it('true bei resourceProvisioningOptions inkl. "Team"', () => {
        expect(groupIsTeam({ resourceProvisioningOptions: ['Team'] })).toBe(true);
        expect(groupIsTeam({ resourceProvisioningOptions: ['Other', 'Team'] })).toBe(true);
    });

    it('false ohne "Team"', () => {
        expect(groupIsTeam({ resourceProvisioningOptions: ['Other'] })).toBe(false);
        expect(groupIsTeam({ resourceProvisioningOptions: [] })).toBe(false);
    });

    it('false ohne Property / null / non-Array', () => {
        expect(groupIsTeam({})).toBe(false);
        expect(groupIsTeam(null)).toBe(false);
        expect(groupIsTeam({ resourceProvisioningOptions: 'Team' })).toBe(false);
    });
});

describe('graphErrorLooksLikeNotFound', () => {
    it('matcht 404 in der message', () => {
        expect(graphErrorLooksLikeNotFound(new Error('HTTP 404 Not Found'))).toBe(true);
    });

    it('matcht ResourceNotFound', () => {
        expect(graphErrorLooksLikeNotFound({ message: 'Code: ResourceNotFound' })).toBe(true);
    });

    it('matcht Request_ResourceNotFound', () => {
        expect(graphErrorLooksLikeNotFound({ message: 'Request_ResourceNotFound' })).toBe(true);
    });

    it('matcht ItemNotFound (case-insensitive)', () => {
        expect(graphErrorLooksLikeNotFound({ message: 'itemnotfound' })).toBe(true);
    });

    it('matcht "not found" Phrase', () => {
        expect(graphErrorLooksLikeNotFound('Object not found')).toBe(true);
    });

    it('liefert false bei anderen Fehlern', () => {
        expect(graphErrorLooksLikeNotFound(new Error('HTTP 500'))).toBe(false);
        expect(graphErrorLooksLikeNotFound({ message: 'Forbidden' })).toBe(false);
    });

    it('robust gegen null/undefined', () => {
        expect(graphErrorLooksLikeNotFound(null)).toBe(false);
        expect(graphErrorLooksLikeNotFound(undefined)).toBe(false);
    });
});

describe('personLabel', () => {
    it('DisplayName + UPN', () => {
        expect(personLabel({ displayName: 'Max Mustermann', userPrincipalName: 'max@x.at' }))
            .toBe('Max Mustermann (max@x.at)');
    });

    it('nur UPN, wenn kein DisplayName', () => {
        expect(personLabel({ userPrincipalName: 'a@b.c' })).toBe('a@b.c');
    });

    it('Mail-Fallback, wenn kein UPN', () => {
        expect(personLabel({ displayName: 'Max', mail: 'm@x.at' })).toBe('Max (m@x.at)');
    });

    it('kein Suffix, wenn DisplayName == UPN', () => {
        expect(personLabel({ displayName: 'max@x.at', userPrincipalName: 'max@x.at' })).toBe('max@x.at');
    });

    it('Id-Fallback, wenn nichts anderes vorhanden', () => {
        expect(personLabel({ id: 'u1' })).toBe('u1');
    });

    it('Leerstring bei null/non-object', () => {
        expect(personLabel(null)).toBe('');
        expect(personLabel('string')).toBe('');
    });
});

describe('odataEscape', () => {
    it('verdoppelt Hochkommas', () => {
        expect(odataEscape("O'Neill")).toBe("O''Neill");
        expect(odataEscape("'foo'")).toBe("''foo''");
    });

    it('lässt anderes unverändert', () => {
        expect(odataEscape('hallo')).toBe('hallo');
        expect(odataEscape('foo bar')).toBe('foo bar');
    });

    it('konvertiert non-strings via String()', () => {
        expect(odataEscape(42)).toBe('42');
    });
});

describe('directoryObjectRef', () => {
    it('baut den vollen Graph-Pfad', () => {
        expect(directoryObjectRef('user-id-1')).toBe('https://graph.microsoft.com/v1.0/directoryObjects/user-id-1');
    });
});

describe('isGraphDuplicateRefError', () => {
    it('matcht "already exist" (case-insensitive)', () => {
        expect(isGraphDuplicateRefError(new Error('One or more added object references already exist'))).toBe(true);
        expect(isGraphDuplicateRefError({ message: 'ALREADY EXISTS' })).toBe(true);
    });

    it('false bei unverwandten Fehlern', () => {
        expect(isGraphDuplicateRefError(new Error('Not found'))).toBe(false);
        expect(isGraphDuplicateRefError({ message: 'Forbidden' })).toBe(false);
    });

    it('robust gegen null/undefined/string', () => {
        expect(isGraphDuplicateRefError(null)).toBe(false);
        expect(isGraphDuplicateRefError('already exist')).toBe(true);
    });
});
