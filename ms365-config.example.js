/**
 * Diese Datei ist die Vorlage; tragen Sie die Client-ID in ms365-config.js ein (siehe gleicher Inhalt).
 *
 * === Entra ID: App-Registrierung (einmalig in IHREM Mandanten) ===
 *
 * 1) https://entra.microsoft.com → Identität → Anwendungen → App-Registrierungen → „Neue Registrierung“
 * 2) Name: z. B. „MS365 Schul-Tool ARGE“
 *    „Unterstützte Kontotypen“: „Konten in einem beliebigen Organisationsverzeichnis (beliebiger Microsoft Entra ID-Mandant – Multimandanten)“
 * 3) „Weiterleitungs-URI“: Plattform „Einzelseitenanwendung (SPA)“
 *    URI exakt so eintragen wie Ihre Seite im Browser, z. B.:
 *    https://IHRE-DOMAIN.tld/ms365-schooltool.html
 *    (lokal testen: http://localhost:PORT/ms365-schooltool.html – dieselbe URI auch in Entra eintragen)
 * 4) Registrieren → auf der Übersichtsseite „Anwendungs-ID (Client)“ kopieren → unten bei clientId einfügen
 * 5) API-Berechtigungen → Berechtigung hinzufügen → Microsoft Graph → Delegierte Berechtigungen:
 *    - Group.ReadWrite.All
 *    - Directory.ReadWrite.All (Modus „Gruppenerstellung“: Verzeichniseinstellung Group.Unified setzen/entfernen)
 *    - Team.Create (Kursteams: POST /teams mit Template educationClass)
 *    - EduRoster.ReadWrite (wird für POST /education/classes mitangefordert; laut Microsoft oft nur App-Only – dann scheitert Kursteam im Browser, siehe Kursteam-Anlage.cmd)
 *    - User.Read.All
 *    - User.ReadWrite.All (Schulstruktur-Sync: Benutzer per Graph aus Typ „Person“ anlegen)
 *    - User.Read (Profil des angemeldeten Benutzers / für GET /me)
 *    - Group.Read.All (optional, lesende Reports: „Leere Gruppen“, „Gast-Zugänge“-Teamliste; sonst deckt Group.ReadWrite.All)
 *    - AuditLog.Read.All (optional, Werkzeug „Gast-Zugänge“: B2B-Einladungen aus Verzeichnis-Audit; GET /invitations existiert in Graph nicht)
 *    - TeamSettings.ReadWrite.All („Teams archivieren“ und Schulstruktur-Sync → Tenant-Details → Team-Archiv per Update: POST …/teams/{id}/archive|unarchive)
 *    - SharePointTenantSettings.Read.All / SharePointTenantSettings.ReadWrite.All (Werkzeuge „SharePoint – Websiteerstellung“ und „SharePoint – Mandanten-Freigaben“)
 *    - Sites.Read.All, Sites.Create.All (Sites.Read.All u. a. für „Hostname per Graph“ in „Intranet & Hub“ und „Mandanten-Freigaben“; Create nur Intranet)
 *    - Sites.ReadWrite.All (SharePoint-Listen „Lehrerliste“ / „Schultermine“: Site auflösen, Liste und Spalten anlegen, ggf. Zeilen schreiben)
 *    - Office 365 SharePoint Online → Sites.FullControl.All (delegiert, optional): Hub-Registrierung per SharePoint-REST aus dem Browser; sonst PowerShell-Fallback im Tool
 *    → „Administratorzustimmung für [Organisation] erteilen“ (Global Admin o. ä.)
 * 6) Unter „Authentifizierung“ prüfen: implizite Genehmigung ist NICHT nötig; SPA + Redirect-URI reicht.
 *
 * Schul-Admins legen KEINE eigene App an – sie öffnen nur Ihre URL und melden sich an (Zustimmung ggf. einmal pro Mandant).
 */
window.MS365_MSAL_CONFIG = {
    /** Anwendungs-ID (Client) aus der App-Registrierung */
    clientId: '',
    /** Multimandanten: alle Organisationskonten */
    authority: 'https://login.microsoftonline.com/organizations',
    /**
     * Muss EXAKT einer „Weiterleitungs-URI“ in der App-Registrierung entsprechen.
     * Standard: aktuelle Seiten-URL ohne Hash (Anker).
     */
    redirectUri: typeof window !== 'undefined' ? window.location.href.split('#')[0] : ''
};
