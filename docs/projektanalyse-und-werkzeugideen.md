# MS365-Schulverwaltung (ms365-schultools) – Projektanalyse und Ideen

*Stand: interne Codebasis-Analyse (Repository MS365schule), ohne Live-Mandant.*

## Kurzfassung

Das Projekt ist eine **statische Web-App** (HTML, CSS, Vanilla-JS) mit **Microsoft-Graph-Anbindung über MSAL im Browser**, ergänzt um **PowerShell-Skripte** dort, wo Graph im delegierten Modus an Grenzen stößt (v. a. Kursteams/Education). Zentrale Schwerpunkte: **Schul-Stammdaten lokal im Browser**, **Teams und Microsoft-365-Gruppen** planen/anlegen/pflegen, **Abgleich** zwischen SOLL-Struktur und Tenant, **Mail/Exchange-Themen** (Verteiler, Postfächer, Gruppenerstellungsrichtlinien).

Das Gefühl „noch nicht ganz rund“ passt zur Architektur: **viele eigenständige Tool-Seiten**, teils **überlappende Graph-/MSAL-Schichten** (`graph-unified-groups.js`, `kursteam-graph.js`, `msal-auth-ui.js` …) und ein **sehr mächtiges** Modul **Schulstruktur-Sync** mit vielen Modi – das wirkt eher wie eine **Werkzeugkiste** als wie ein durchgängiger Assistent.

---

## Technischer Überblick

| Aspekt | Befund |
|--------|--------|
| Build | Vite bündelt mehrere HTML-Einstiegspunkte (`index.html`, `tenant.html`, `tools/*.html` …). |
| Konfiguration | `ms365-config.js` / `ms365-config.example.js` (Client-ID, Redirect, Hinweise zu Entra-Berechtigungen). |
| Daten lokal | `app-data-v2.js` (Schema v3): Einrichtungsassistent, Zuordnungen, Schuljahresdaten u. a. |
| Tenant-Einstellungen | `tenant-settings-core.js` / `tenant-settings-ui.js` – Stammdaten, Klassen-Anzeigenamen, Listen. |
| Graph | Delegierte Scopes u. a. in `graph-unified-groups.js` und tool-spezifischen `*-graph.js`; `ms365-config.example.js` listet erweiterte Scopes (z. B. `Directory.ReadWrite.All`, `TeamSettings.ReadWrite.All`, `EduRoster.ReadWrite`). |
| Deep-Link | `ms365-schooltool.html` verarbeitet Redirect und leitet per `?mode=` zu Tools weiter. |
| Dashboard | `index.html`: Kategorien-Chips, Suche, Favoriten/Sortierung per `localStorage`. |

---

## Inventar: sichtbare Werkzeuge (Dashboard & nahe Umgebung)

- **Schul-Grundeinstellungen** (`tenant.html`) – lokale Stammdaten, Klassen-Tab, Statistiken auf dem Start-Dashboard.
- **Geführte Einrichtung** (`einrichtung.html`) / **Experten: Struktur** (`ersteinrichtung.html`) – Onboarding und strukturierte Planung.
- **Lehrerinnen & Schülerinnen** – Sammelgruppen (`tools/schueler-lehrer-gruppen.html`).
- **Jahrgangsgruppen** (`tools/jahrgang.html`).
- **Kursteams** – WebUntis-Import, Teams mit EDU-Vorlage, Online vs. PowerShell (`tools/kursteams.html`).
- **Weitere Teams & Gruppen** – generische Gruppen/Teams-Pipeline (`tools/weitere-teams-gruppen*.html`).
- **Personen-Verwaltung** – Benutzer:innen und Gruppenmitgliedschaften (`tools/personen-verwaltung.html`).
- **MS 365 Gruppenverwaltung / Schulstruktur-Sync** – große Oberfläche inkl. Archiv, Details, Graph-Aktionen (`tools/schulstruktur-sync.html`).
- **Abgleichen** – Einstieg per Redirect in Schulstruktur-Sync mit Modus `match` (`tools/abgleichen.html`).
- **Freigegebene Postfächer** (v1 Lesen/Heuristik) (`tools/postfaecher.html`).
- **Verteilerlisten & mail-aktivierte Sicherheitsgruppen** – Graph + Exchange-Online-PowerShell-Skripte (`tools/verteilerlisten.html`).
- **Gruppenerstellung einschränken** – Verzeichniseinstellung `Group.Unified` (`tools/gruppenerstellung.html`).
- **Archiv** – z. B. ARGE-Tool, Teams-Archiv, ältere Klassen-Umbenennen-Wizard-Seiten unter `tools/archiv/`.

**Hinweis:** Klassen-Anzeigenamen sind in den **Einstellungen** verankert (`tenant.html#classes`), nicht als eigenständige Dashboard-Kachel – das kann je nach Erwartung „Lücken“ erzeugen.

---

## Stärken (was schon gut zur Schule passt)

1. **Klare Domäne**: Gruppen, Kursteams, Jahrgang, SLG – typische österreichische/deutsche Schul-IT-Szenarien.
2. **Datenschutz-freundlicher Ansatz**: Stammdaten primär **lokal im Browser**, keine zwingende Server-Backend-Pflicht für den Kern.
3. **Pragmatische Hybrid-Strategie**: Wo Graph im SPA nicht reicht, **PowerShell** als Ausweg (Kursteams, Exchange) – realistisch für Schul-IT.
4. **Struktur-Regeln** (`structure-rules.js`): explizites Modell (Jahrgang → Klasse → Kursteam/Gruppe …) – gute Basis für spätere Automatisierung und Validierung.
5. **Dashboard-UX**: Suche, Kategorien, Pin/Sortierung – solide Basis für wachsende Toolzahl.

---

## Warum es sich „unrund“ anfühlen kann

1. **Mehrere parallele „App-Kerne“**: Einrichtung, Tenant, jedes Tool mit eigenem HTML-Head und teils eigenem Graph-Client – mentales Modell für Nutzer:innen: „Eine App“ vs. „viele Seiten“.
2. **Schulstruktur-Sync als „Schweizer Taschenmesser“**: Viele Funktionen an einem Ort erhöhen Lernkurve und Wartungslast; **Abgleichen** ist bereits ein Alias – das zeigt, dass Nutzer:innen einen **eigenen mentalen Einstieg** wollen.
3. **Versions- und Archiv-Split**: Aktive vs. archivierte Tools (ARGE, Teams-Archiv, alte Wizard-Routen) – sinnvoll für Migration, aber für das Gesamtbild leicht uneinheitlich.
4. **Berechtigungs-Komplexität**: Volle Power der App braucht **breite Admin-Scopes**; Schulen mit restriktiveren Rollen brauchen ggf. **abgestufte Modi** (nur Lesen, nur Skript-Export ohne Online-Ausführung) – sonst wirkt das Produkt mal „zu mächtig“, mal „blockiert“.
5. **Kursteams und Education-API**: dokumentierte Einschränkung bei `EduRoster.ReadWrite` im Browser – fachlich nachvollziehbar, produktseitig aber ein **Bruch** zwischen „klick fertig“ und „CMD ausführen“.

---

## Richtungen, um es „runder“ zu machen (ohne sofort alles umzubauen)

### Produkt / Information Architecture

- **Eine „Shell“-Seite** für alle Tools: gemeinsame Kopfzeile (Mandant, angemeldete Person, Kontext Schuljahr), **Breadcrumb**, **einheitliche Anmelde-Schaltfläche** und Status – Inhalte als „Panels“ oder iframe-freie Teilviews laden (größerer Umbau, hoher Gewinn).
- **„Playbooks“ statt nur Tools**: geführte Checklisten („Schuljahresstart“, „Neue Klasse“, „Lehrerwechsel“), die **bestehende Tools** in sinnvoller Reihenfolge verlinken und Zwischenstände in `app-data-v2` speichern.
- **Lesemodus / Audit-Modus**: gleiche Datenquellen, aber ohne Schreib-Scopes – reduziert Angst und erleichtert Piloten.

### Technik / Wartbarkeit

- **Eine Graph-Zugriffsschicht** pro Mandant-Session (Token-Cache, `fetch`-Wrapper, Pagination, Throttling) statt mehrerer nahezu kopierter MSAL-Blöcke – langfristig weniger Bugs bei Token-Renewal.
- **Feature-Flags** in der Konfiguration: welche Kacheln/Scopes für welche Rolle – optional pro Schul-Typ (AHS, VS, …).

### UX-Kleinigkeiten

- **Konsistente „Zurück zum Dashboard“-Position** und einheitliche **Fehler-/Hinweis-Komponente** (Toast, Dialog) auf allen Tool-Seiten.
- **Kontext anzeigen**: aktuelles Schuljahr, Domain, ob Daten „nur lokal“ vs. „Tenant“ betroffen sind – senkt Support-Fragen.

---

## Ideen für weitere praktische Werkzeuge

Die Liste ist bewusst **nah an Microsoft 365 für Schulen** und an eurer bestehenden Datenbasis (Klassen, Fächer, Gruppen). Priorisierung sollte über **Häufigkeit im Schulalltag** und **Graph-/Policy-Zulässigkeit** erfolgen.

### A) Direkt anschlussfähig (ähnliche Patterns wie heute)

| Idee | Nutzen | Anmerkung |
|------|--------|-----------|
| **Namenskonvention-Prüfer** | Alle Gruppen/Teams gegen Regeln aus Tenant-Einstellungen prüfen (Präfix, Jahr, verbotene Zeichen) | Nutzt vorhandene Validierungslogik ähnlich `sanitizeUnifiedGroupMailNickname` / Struktur-Export; Ergebnis als Report + CSV. |
| **„Leerer Gruppen“-Report** | Gruppen ohne Owner, ohne Mitglieder, ohne Nutzung (Letzte Aktivität nur eingeschränkt möglich) | Fokus Owner/Mitglied = leicht mit Graph; Aktivität ggf. nur über Teams/Reports API begrenzt. |
| **Gast-Zugänge / Externe** (Lesend) | Übersicht Teams mit Gästen, Einladungen ausstehend | Wichtig für DSGVO-Audits; meist lesende Scopes + klare UI-Hinweise. |
| **Teams-Richtlinien / Messaging (Lesend)** | Welche Policies gelten für Kursteam vs. Staff? | Oft nur lesend für Schul-IT hilfreich; Schreiben ist tenant-weit heikel. |
| **SharePoint-Klassenwebs** | Pro Kursteam/Gruppe Team-Site-URL und Bibliotheken anzeigen; „fehlende“ Sites markieren | Graph `group` → `team` → `site`; eher Übersicht als Massenänderung. |
| **OneDrive-Quota & Freigaben-Report** | Speicherwarnung für Lehrkräfte vor Prüfungen | Braucht zusätzliche Scopes (`Files.Read.All` o. ä.) – eher optionaler „Advanced“-Bereich. |
| **Lizenz- und Zuweisungs-Übersicht** | Wer hat A1/A3, Teams-Lizenz, Exchange? | `subscribedSkus`, `assignedLicenses` – sehr gefragt in Schulen; Daten sensibel, UI klar trennen. |

### B) Schuljahres- und Organisationslogik (baucht auf `app-data-v2` + Struktur)

| Idee | Nutzen |
|------|--------|
| **„Schuljahreswechsel-Assistent“** | Archivieren/umbenennen/duplizieren von Strukturzeilen; Verknüpfung zu Klassen-Anzeigenamen und Kursteam-Listen. |
| **Kohorten / Abschlussjahrgang** | Automatisch Gruppen für Maturajahrgang, Elternschaft-Jahrgang (wenn Datenmodell erweitert). |
| **Fachschafts-Gruppen aus Tenant-Fächerliste** | Einmalig Fach → Unified Group mit konsistentem `mailNickname`; ähnlich Wizard-Teilen zu ARGE/Fach-Präfixen. |

### C) Kommunikation & Kalender (mittlerer Aufwand)

- **Raum- und Gerätereservierung** (Bookings, Räume aus `place`/`room`): Stundenplannahe Use Cases; oft eher **Lesen + Link** in Admin Center statt voller Write.
- **Verteiler „Schuljahres-Snapshot“**: Mitgliedschaft aus Klassenlisten **exportieren** (du habt schon Verteilerlisten-PS) – reiner Export reduziert Risiko.

### D) Sicherheit & Compliance (vorsichtig positionieren)

- **Anmelde-Risiko / MFA-Status** (nur für berechtigte Leserollen): `signInActivity`, `authenticationMethods` – stark nachgefragt, aber **sehr sensibel**; als separates „Security“-Tool mit extra Consent.
- **Gruppen-Ablauf / Verwaltung durch Besitzer** erinnern (nur Hinweise + Links zu Entra-Einstellungen).

### E) Integrationen außerhalb Graph (Dokumentation + Export)

- **School Data Sync (SDS)** – **Kein Ersatz** für eure App, aber: Assistent „CSV für SDS“ aus eurer Strukturtabelle – viele Schulen pflegen parallel SDS für Insights/License.
- **Lernmanagementsystem**: nur **Deep-Links** oder CSV-Konventionen (Moodle Kurs-ID ↔ Gruppe) – geringer Scope, hoher Alltagsnutzen.

### F) „Meta-Werkzeuge“ (machen die Kiste runder, ohne neue Graph-Monster)

- **Aktionsprotokoll lokal**: Wer hat wann welches Bulk-Skript erzeugt (rein Browser, JSON-Export) – erhöht Vertrauen bei Admins.
- **Konfigurations-Backup**: `tenant` + `app-data-v2` **exportieren/importieren** (Versionierung, Schulwechsel) – oft mehr wert als ein weiteres Graph-Feature.

---

## Empfohlene nächste Schritte (pragmatisch)

1. **Ein Playbook „Schuljahresstart“** (nur Text + Links + gespeicherte Checklisten in `app-data-v2`) – schnell, hohe wahrgenommene Rundung.
2. **Namenskonvention-Prüfer / Report** – baut auf vorhandenen Daten und Graph-Reads auf, klare Nutzenstory.
3. **Lizenzübersicht (lesend)** – extrem gefragt, Scope-Karte im Hilfe-Dokument ergänzen.
4. Langfristig: **eine gemeinsame App-Shell** oder zumindest **gemeinsames Layout-Skript** für alle `tools/*.html`, um Kopfzeile, Auth und Navigation zu vereinheitlichen.

---

## Literatur / interne Referenzen im Repo

- Berechtigungs-Leitfaden: `ms365-config.example.js`
- Strukturmodell: `src/tools/schulstruktur-sync/structure-rules.js`
- Lokales Datenmodell: `src/shared/app-data-v2.js`
- Dashboard-Katalog: `index.html` (Kacheln und `data-cluster`)

---

*Dieses Dokument wurde als Arbeitsgrundlage aus der Code-Struktur abgeleitet; Prioritäten sollten mit echten Schul-IT-Prozessen validiert werden.*
