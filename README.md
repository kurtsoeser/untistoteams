# MS365-Schulverwaltung

Einfache **reine Browser-App** (ohne Server): **Dashboard** (`index.html`), **Schul-Grundeinstellungen** (`tenant.html`) und mehrere Werkzeuge unter `tools/` – u. a. **Kursteams** (`tools/kursteams.html`, Logik in `src/tools/kursteams/`), **Jahrgangsgruppen** und **ARGEs** (Archiv). Exportdaten (z. B. WebUntis CSV/Excel) werden aufbereitet; Anlage per **Microsoft Graph** im Browser oder per **PowerShell**-Skripten/CMD-Downloads.

## Datenschutz

Es werden **keine Daten** an einen Server gesendet. Verarbeitung erfolgt **lokal im Browser**. Optional können Sie pro Modus (**Kursteams**, **Jahrgangsgruppen**, **ARGEs**) einen getrennten Zwischenstand im **lokalen Speicher** dieses Browsers sichern (Schaltflächen oben in der App).

## Nutzung

1. Repository klonen oder die Dateien herunterladen.
2. **`index.html`** (Dashboard mit allen Werkzeugen) im Browser öffnen **oder** ein statisches Hosting nutzen (z. B. GitHub Pages).
3. **`ms365-schooltool.html`** ist typischerweise die **Redirect-URI** der Entra-Anwendung: Nach der Anmeldung verarbeitet die Seite MSAL und leitet per **`?mode=…`** (z. B. `schulstruktur`, `slg`) auf die passende Datei unter `tools/` weiter; ohne Parameter geht es zum Dashboard.

## Entwicklung & Hosting (Vite + GitHub Pages)

Dieses Repo ist eine **statische Multi-Page-App**. Für lokale Entwicklung und für GitHub Pages kann Vite verwendet werden.

### Lokal starten

```bash
npm install
npm run dev
```

### Build (für GitHub Pages)

```bash
npm run build
```

`npm run build` führt `vite build` und anschließend `node scripts/copy-static.mjs` aus: gebündelte HTML-Einträge in `dist/` plus Kopie u. a. von `src/**`, Root-`app.css`, `ms365-schooltool.html` und weiteren Root-Dateien aus `scripts/copy-static.mjs`. **Sync/Commit ersetzt keinen Build** – vor dem Deploy lokal oder in CI ausführen.

Für **GitHub Pages Project Pages** (URL `https://<user>.github.io/<repo>/`) setzt der Workflow automatisch `VITE_BASE="/<repo>/"`.
Lokal wird standardmäßig `/` verwendet.

### Tests (optional)

Unit-Tests für **reine JS-Logik** (ohne Browser-DOM), u. a. Kursteams-Helfer und weitere Module:

```bash
npm install
npm test
```

Während der Entwicklung: `npm run test:watch`

Auf GitHub läuft bei Push/PR ein Workflow (`.github/workflows/ci.yml`): `npm ci`, `npm test`, `npm run build`.

### Kursteams: Skript-Reihenfolge

Die Datei **`tools/kursteams.html`** lädt die Schritte über Module unter **`src/tools/kursteams/`** (siehe **HTML-Kommentar** zur Ladereihenfolge). Zusätzlich prüft `src/shared/ms365-module-guard.js` (`window.ms365AssertModules`) beim Start, ob abhängige Module geladen sind.

### Hinweis zu `ms365-config.js`

`ms365-config.js` ist eine **Runtime-Konfiguration** (Client-ID/Redirect-URI) und wird als normale Datei mit deployed.
Wenn Sie die Client-ID nicht öffentlich im Repo haben möchten, löschen Sie `ms365-config.js` aus dem Repo und legen Sie sie
in Ihrem Deployment (Pages/Hosting) separat ab – dann muss die Datei weiterhin im Root erreichbar sein.

### Head-Minimum für neue HTML-Seiten (statische MPA)

Wir **duplizieren** bewusst die gleichen Head-Zeilen in vielen Dateien (kein gemeinsamer Build-Partial). Für neue Seiten orientieren Sie sich an einer bestehenden Seite **derselben Ordner-Ebene**:

| Seite liegt in … | `app.css` |
|------------------|-----------|
| Repository-Root (`index.html`, `tenant.html`, …) | `href="app.css"` |
| `tools/*.html` | `href="../app.css"` |
| `tools/archiv/*.html` | `href="../../app.css"` |

**Bootstrap Icons** (einheitliche Version, wie in den anderen Seiten):  
`https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css`

**`<meta name="ms365-graph-client-id" content="…">`** nur dort, wo die Seite **Microsoft Graph / MSAL** im Browser nutzt (wie bei den bestehenden Tool-Seiten). Wert wie in den Nachbar-HTMLs oder konsistent mit `ms365-config.js`. Öffentliche Infoseiten ohne Graph brauchen dieses Meta nicht.

**Minimale Weiterleitungs-HTMLs** (nur Refresh/Script, z. B. unter `tools/`) können ohne `app.css`, ohne Icons und ohne Graph-Meta auskommen.

### Modus A: Kursteams

1. Fächerliste exportieren und durch die Schritte in der App führen.
2. Entweder `neueteams.csv` + kurzes Skript aus der Anleitung **oder** die Datei **Kursteam-Anlage.cmd** herunterladen (enthält das PowerShell-Skript eingebettet) und per Doppelklick starten (Anmeldung interaktiv für MFA oder optional per `Get-Credential`).

### Modus B: Jahrgangsgruppen (Microsoft 365-Gruppen)

1. Domain und Präfix (z. B. `jg`) einstellen; Klassenzeilen im Format `1AK;2030` (Klasse;Abschlussjahr) eintragen.
2. Besitzer-UPNs pro Klasse eintragen.
3. Das generierte **Microsoft Graph PowerShell**-Skript kopieren **oder** **Jahrgangsgruppen-Anlage.cmd** herunterladen und starten (interaktive Graph-Anmeldung, MFA möglich).

Dokumentation: [New-MgGroup](https://learn.microsoft.com/powershell/module/microsoft.graph.groups/new-mggroup), [Gruppe erstellen (Graph)](https://learn.microsoft.com/graph/api/group-post-groups).

### Modus C: ARGEs (Microsoft 365-Gruppen)

1. **Domain** und optional **Präfix** für das Mail-Nickname festlegen (Präfix + aus dem Fachnamen erzeugter Teil, z. B. Fach `Deutsch` → `arge-deutsch`). Optional: Mail-Nickname in Großbuchstaben.
2. **Fächer / Bezeichnungen** eintragen: **eine Zeile pro ARGE** – z. B. aus Excel kopieren (`Deutsch`, `Mathematik`, `ARGE BB`, …). Daraus erzeugt die App den **Anzeigenamen** der Gruppe (`ARGE …`, sofern die Zeile nicht schon mit `ARGE ` beginnt) und das **Mail-Nickname** (Umlaute und Sonderzeichen werden für den technischen Teil normalisiert). Unter dem Eingabefeld zeigt eine **Live-Vorschau** als Tabelle: Anzeigename, Fach technisch, Mail-Nickname und die **E-Mail-Adresse** (`MailNickname@Domain`).
3. **Optional (fortgeschritten):** Weiterhin pro Zeile `Anzeigename;MailNickname` möglich – dann gilt der rechte Teil als festes Mail-Nickname (z. B. `ARGE BB;ARGEBB`).
4. Besitzer-UPNs pro ARGE eintragen.
5. Script kopieren oder **ARGE-Gruppen-Anlage.cmd** herunterladen und starten (interaktive Graph-Anmeldung, MFA möglich).

### Erwartete Spalten (flexibel)

Die App erkennt u. a.: **Klasse(n)**, **Fach**, **Lehrer**, **Schülergruppe** (je nach Export unterschiedlich benannt).

### Dateien (Auswahl)

| Pfad | Beschreibung |
|--------|----------------|
| `index.html` | **Dashboard**: Kacheln zu allen Werkzeugen, Suche, Favoriten, Links zu Einstellungen und Einrichtung |
| `tenant.html` | **Schul-Grundeinstellungen**: Stammdaten, Tab „Klassen“, lokale Daten; nutzt u. a. `src/shared/msal-auth-ui.js` und `src/shared/tenant-settings-*.js` |
| `einrichtung.html` | **Geführte Einrichtung** (Setup-Wizard; `src/shared/setup-wizard.js` u. a.) |
| `ms365-schooltool.html` | **Redirect-/Login-Einstieg** für Entra: MSAL `handleRedirectPromise`, Rücksprung nach `sessionStorage`, Weiterleitung per `?mode=…` auf `tools/…` oder `index.html` |
| `tools/kursteams.html` | Kursteams-Oberfläche; Logik in **`src/tools/kursteams/`** (u. a. `kursteam-graph.js`, `kursteam-ui.js`, …) |
| `tools/jahrgang.html` | Jahrgangsgruppen-Assistent; Skripte unter **`src/tools/jahrgang/`** |
| `tools/archiv/arge.html` | ARGE-Assistent (Archiv); Skripte unter **`src/tools/arge/`** |
| `src/shared/` | Gemeinsame Module (z. B. `app-data-v2.js`, `tenant-settings-*.js`, `graph-unified-groups.js`, `msal-loader.js`, `polyglot-cmd.js` für CMD-Downloads) |
| `app.js` | **Kompatibilitäts-Stub** (no-op): frühere Bookmarks/Deployments, die noch `app.js` laden, brechen nicht; die Kursteams-Logik liegt unter `src/tools/kursteams/` |
| `app.css` | Gemeinsames Layout/Styling für die meisten Seiten |
| `ms365-config.js` / `ms365-config.example.js` | MSAL-/Graph-Runtime-Konfiguration (siehe Abschnitt unten) |

### Windows: Downloads und Sicherheit

Aus dem Browser heruntergeladene Dateien können mit **Mark of the Web** (Zone) markiert sein. Windows kann dann melden, dass die Datei „wegen Internetsicherheitseinstellungen“ nicht geöffnet werden darf – das betrifft auch **eine einzelne** `.cmd` im Ordner **Downloads**.

**So gehen Sie vor (privat / Einzelplatz):**

1. **Rechtsklick** auf die heruntergeladene `.cmd` → **Eigenschaften**.
2. Unten **Zulassen** (engl. **Unblock**) aktivieren → **OK**.
3. Datei **erneut** per Doppelklick starten.

**Alternative:** In **PowerShell** (als Benutzer reicht meist):

```powershell
Unblock-File -LiteralPath "$env:USERPROFILE\Downloads\ARGE-Gruppen-Anlage.cmd"
```

(Pfad und Dateiname anpassen.)

Bei **SmartScreen** („Windows hat den PC geschützt“): **Weitere Informationen** → **Trotzdem ausführen** – nur wenn Sie die Datei aus dieser App selbst erzeugt haben.

Die App liefert **eine** `.cmd` mit eingebettetem PowerShell (keine separate `.ps1` im Download-Ordner). Eine **100 % warnungsfreie** Ausführung aus dem Internet ohne Signatur oder IT-Freigabe kann Windows nicht garantieren.

## GitHub Pages

Repository auf **Pages** schalten (Branch `main`, Ordner `/`). Mit dem Repo-Namen **`ms365-schultools`** ist die App typischerweise unter:

- **Dashboard:** [https://kurtsoeser.github.io/ms365-schultools/](https://kurtsoeser.github.io/ms365-schultools/) (`index.html`)
- **Redirect-/Login-URI (Entra):** [https://kurtsoeser.github.io/ms365-schultools/ms365-schooltool.html](https://kurtsoeser.github.io/ms365-schultools/ms365-schooltool.html)

Klonen per Git (nach tatsächlichem Repo-Namen auf GitHub):

```bash
git clone https://github.com/kurtsoeser/ms365-schultools.git
```

### Lokales Git nach Umbenennung des Repos auf GitHub

Wenn Sie das Repository auf GitHub umbenannt haben, passen Sie die **Remote-URL** in Ihrem geklonten Ordner an (einmalig):

```bash
git remote set-url origin https://github.com/kurtsoeser/ms365-schultools.git
git remote -v
```

**Hinweis:** Bei kostenlosem GitHub ist Pages für **private** Repos oft nicht verfügbar; öffentliches Repo oder GitHub Pro nötig.

## Lizenz

Keine Lizenz gesetzt – ergänzen Sie bei Bedarf eine `LICENSE` im Repository.
