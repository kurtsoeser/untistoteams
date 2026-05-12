# Code-Architektur & Refactoring-Leitplanken

Dieses Dokument beschreibt die **Soll-Architektur** der Tool- und Shared-Module
unter `src/` und das Vorgehen, mit dem wir die aktuell teils sehr großen
Monolith-Dateien (`schulstruktur-sync.js`, `setup-wizard.js`,
`tenant-settings-ui.js`, …) Stück für Stück in lesbare, testbare Teile
zerlegen.

> **Goldene Regel:** Refactorings sind **reine Move-/Extract-Operationen**.
> Verhalten darf sich dabei **nicht** ändern. Funktionale Verbesserungen
> kommen erst in einem separaten Schritt **nach** dem Split.

---

## 1. Größen-Leitplanken

| Bereich | Soll | Hartes Limit |
|---|---|---|
| Tool-/Shared-Modul (`.js`) | 100–400 Zeilen | < 600 Zeilen |
| Test-Dateien (`.test.mjs`) | beliebig | beliebig |

ESLint warnt ab **650 Zeilen** (Blanks/Kommentare ausgenommen). Die Warnung
ist absichtlich „Warn" und nicht „Error", damit der Lint-Build der
historisch gewachsenen Dateien grün bleibt – sie zeigt aber transparent,
welche Files als nächstes anstehen.

---

## 2. Ordnerstruktur

```
src/
├─ shared/
│  ├─ utils/                ← schichtfreie, kleine Helpers (ESM)
│  │  ├─ strings.js         ← normStr, normCode, normEmail, escapeHtml, …
│  │  ├─ json.js            ← safeJsonParse
│  │  ├─ dialog.js          ← dlgAlert/Confirm/Prompt-Wrapper
│  │  ├─ storage.js         ← loadJson/saveJson/removeKey
│  │  ├─ dom.js             ← getEl, …
│  │  └─ index.js           ← Barrel-Re-Export
│  ├─ app-dialog.js         ← konkrete Dialog-Implementierung (UI)
│  ├─ graph-unified-groups.js
│  ├─ msal-auth-ui.js
│  └─ …
└─ tools/
   └─ <tool>/
      ├─ <tool>.js          ← Entry/Wiring, DOMContentLoaded, Event-Binding
      ├─ <tool>-ui.js       ← Render-Funktionen, HTML/DOM-Aufbau
      ├─ <tool>-state.js    ← In-Memory-State, Defaults, Migrationen
      ├─ <tool>-storage.js  ← localStorage-Keys + Load/Save
      ├─ <tool>-graph.js    ← Graph-API-Calls (Fetch, Scopes)
      ├─ <tool>-logic.js    ← reine Geschäftslogik (testbar ohne DOM!)
      ├─ <tool>-csv.js      ← CSV-Im-/Export
      ├─ <tool>-parse.js    ← Eingabe-Parser
      └─ <tool>-import/-export.js
```

Nicht jedes Tool braucht alle Suffixe. Der **kleinste sinnvolle** Split
ist meist: `<tool>.js` + `<tool>-graph.js` + `<tool>-state.js`.

---

## 3. Modulform: ESM bevorzugen

**Neuer Code = ES-Module** (`export`/`import`). Vorteile:

- Statische Importgraph → bessere Tree-Shaking-/Build-Optimierung durch Vite
- Echte Tests ohne DOM (Vitest)
- IDE-Navigation und -Refactoring funktionieren

**Legacy = IIFE.** Bestehende `(function () { 'use strict'; … })()`-Files
werden nur dann auf ESM umgestellt, wenn das Tool ohnehin aufgeteilt
wird. Bei der Umstellung müssen die zugehörigen HTML-Files ihre
`<script src="…">`-Tags auf `<script src="…" type="module">` migrieren.

### Zugriff auf gemeinsame Helpers

Aus ESM-Modulen:

```javascript
import { normStr, escapeHtml } from '../../shared/utils/strings.js';
import { safeJsonParse } from '../../shared/utils/json.js';
import { dlgAlert, dlgConfirm } from '../../shared/utils/dialog.js';
```

Aus Legacy-IIFE-Files: vorerst **die lokalen Definitionen behalten**.
Die Umstellung erfolgt zusammen mit der ESM-Migration des Tools.

---

## 4. Schicht-Trennung in einem Tool

| Schicht | Darf nutzen | Darf NICHT nutzen |
|---|---|---|
| **Logic** (`-logic.js`) | nur reine JS-Funktionen, eigene Helpers | DOM, `window`, `fetch`, `localStorage` |
| **State** (`-state.js`) | Logic, Helpers | DOM |
| **Storage** (`-storage.js`) | `localStorage`, Helpers | DOM, Graph |
| **Graph** (`-graph.js`) | `fetch`, MSAL, Helpers | DOM |
| **UI** (`-ui.js`) | DOM, State, Logic, Helpers | direkte Graph-Calls (über Entry/Wiring orchestrieren) |
| **Entry** (`<tool>.js`) | alles | – |

Vorteil: Logic + State sind **ohne DOM** testbar (Vitest).

---

## 5. Refactoring-Workflow

Pro großem File:

1. **Inventarisieren**: Funktionen nach Schicht gruppieren
   (Grep nach `function …(`)
2. **Extract**: Funktion(en) 1:1 in neue Datei mit `export` ziehen,
   Imports in der alten Datei hinzufügen
3. **Smoke-Test**: HTML-Seite einmal aufrufen
   (`npm run dev`), Build prüfen (`npm run build`)
4. **Tests anlegen** für Logic/State, sobald getrennt
5. **Erst dann** kosmetische Verbesserungen

> **Wichtig:** Pro Commit / PR möglichst **nur ein** Tool oder **ein**
> Helper-Modul anfassen. Reviews werden sonst unmöglich.

---

## 6. Aktueller Stand (Stand: Phase 2-Pilot + ESM-Adoption + 4 Multi-File-Tools, 2026-05-12)

### Schulstruktur-Sync (Pilot, **abgeschlossen**)

`schulstruktur-sync.js`: **7138 → 5464 Zeilen (−23,5 %)**, 13 neue Module:

| Modul | Z. | Verantwortlichkeit |
|---|---:|---|
| `-tree.js` | 497 | Strukturbaum, Icon-Logik |
| `-graph-layout.js` | 441 | Organigramm-Layout + Render |
| `-match.js` | 307 | Tenant-Abgleich-Heuristiken |
| `-demo.js` | 268 | Demo-Daten, Default-Schemata |
| `-io.js` | 206 | CSV-Export, PowerShell-Generierung |
| `-state.js` | 202 | Persistenz |
| `-naming.js` | 188 | Mail-Nicknames, Schuljahr |
| `-stats.js` | 179 | Statistik, Filter, Datumsformat |
| `-mapping.js` | 156 | Match-Link-Updates, Select-Value, Dirty-Check |
| `-render.js` | 154 | DOM-Render (Stats, Filter, Hint) |
| `-anlegen.js` | 143 | Tenant-Defaults & Vorschläge |
| `-graph-helpers.js` | 134 | Pure Graph-API-Helfer |
| `-notify.js` | 60 | Toast, Progress-Bar |

**Test-Bilanz:** 24 → 343 Tests (+319) in < 1 s.

### ESM-Adoption von Shared Utils

Diese Tools wurden bereits auf ESM umgestellt und nutzen `shared/utils/*`:

| Tool | Vorher → Nachher | Adoptierte Helfer |
|---|---|---|
| `postfaecher.js` | 718 → 705 Z. | `safeJsonParse`, `getEl`, `compareDe` |
| `leere-gruppen-report.js` | 640 → 638 Z. | `getEl` |
| `gruppenerstellung-policy.js` | 658 → 651 Z. | `dlgConfirm` |
| `klassen-umbenennen.js` | 1111 → 1085 Z. | (keine Duplikate, nur ESM-konform) |
| `gast-zugaenge.js` | 1366 → 1321 Z. | `getEl`, `compareDe`, `escapeHtml` |
| `gast-einlader.js` | 1388 → 1373 Z. | (keine Duplikate, nur ESM-konform) |

### ESM-Migration Multi-File-Tools

| Tool | Files | Bundle | Status |
|---|---:|---|---|
| `arge/` | 6 | `arge-*.js` 55 KB / 15 KB gzip | **fertig** |
| `jahrgang/` | 5 | `jahrgang-*.js` 50 KB / 14 KB gzip | **fertig** |
| `weitere-teams-gruppen/` | 4 (5 HTML-Entries) | `wtg-*.js` 25 KB / 8 KB gzip | **fertig** |
| `kursteams/` | 20 | `kursteams-*.js` 103 KB / 27 KB gzip | **fertig** |

**Pattern für Multi-File-Tools** (gleich für alle 4):

- Alle Submodule auf ESM (kein IIFE-Wrapper)
- Pure Logic-Files: lokale Duplikate von `normStr` etc. durch `shared/utils`-Imports ersetzt
- **Cross-Refs zwischen Submodulen** wahlweise:
  - **echter ESM-Import** (sauber, z. B. `arge/`, `jahrgang/`)
  - **`window.ms365XXX`-Loose-Coupling beibehalten** (semantisch
    identisch, kein zirkulärer Import nötig, z. B. `kursteams/`
    mit seinem geteilten `ns = window.ms365Kursteam`-Namespace)
- **Entry-Point** (z. B. `arge.js`, `jahrgang.js`, `wtg.js`, `kursteam-teams.js`):
  Side-Effect-Imports aller Submodule am Anfang in topologischer Reihenfolge.
  Externe `window.ms365*`-Bridges für inline-`onclick`-Handler in HTML
  (z. B. `ms365ArgeGraphLogin/Run`, `ms365JahrgangGraphLogin/Run`,
  `ms365WtgGraphLogin/Run`, `ms365KursteamMembersGraphLogin/Run/FetchGraphTeams`)
  bleiben erhalten, bis die HTML-Seiten selbst migriert sind.
- **HTML**: alle N `<script src="..." defer>` werden durch **einen**
  `<script type="module" src="entry.js">` ersetzt.
  Bei `kursteams.html` bleibt zusätzlich `ms365-module-guard.js` als
  `<script defer>` davor, weil es einen Pre-ESM-Check
  (`window.ms365AssertModules`) bereitstellt.

**Spezialfälle**:
- `kursteam-graph.js` hatte eine eigenwillige
  `<script>`-Tag-Lookup-Konstruktion für den `msal-loader.js`-Pfad, die
  bei gehashten Vite-Bundles bricht. Ersetzt durch einen statischen
  `await import('../../shared/msal-loader.js')` (Vite versteht das,
  packt `msal-loader.js` als separaten Chunk).

**Build-Erkenntnis**: Vite zieht shared Module (`dom-…js`, `strings-…js`,
`dialog-…js`, `json-…js`) als eigene Chunks raus → echtes Tree-Shaking. Bei
Multi-File-Tools wie `arge` zieht Vite alle 6 Submodule **gemeinsam** in
einen Tool-Bundle, mit `modulepreload` für die shared Chunks.

### Übrige Monolithen

| Tool | Hauptdatei (Zeilen) | Schon gesplittet? | Modul-Typ |
|---|---:|---|---|
| setup-wizard | 3485 | teilweise (`-admin-model.js`) | ESM |
| tenant-settings-ui | 2450 | teilweise (`-core.js`) | IIFE |
| gast-einlader | 1388 | nein | IIFE |
| schueler-lehrer-gruppen | 1377 | teilweise (`slg-…`) | IIFE |
| gast-zugaenge | 1366 | nein | IIFE |
| personen-verwaltung | 1208 | nein | IIFE |
| klassen-umbenennen | 1111 | nein | IIFE |
| jahrgang | 1110 | ja (`-csv`, `-graph`, `-state`, `-standalone-ps1`) | IIFE |
| arge | 1068 | ja (`-csv`, `-graph`, `-state`, `-standalone-ps1`, `-parse`) | IIFE |
| kursteams | 997 (kursteam-members) | sehr gut (22 Module) | IIFE |

`kursteams/` ist die **Referenz-Implementierung** für gut gesplittete
Tools. Die ESM-Migration steht für die meisten Tools noch aus.

---

## 7. Migrationsfahrplan (Phasen)

- **Phase 0** _(abgeschlossen)_ – Standards, ESLint-Warn, dieses Dokument
- **Phase 1** _(abgeschlossen)_ – Zentrale Helpers extrahieren
  (`src/shared/utils/`)
- **Phase 2-Pilot** _(abgeschlossen)_ – `schulstruktur-sync.js` in 13
  thematische Module zerlegt; Pure-Logik mit DI testbar gemacht
  (z. B. `buildKursteamCsv(rows, …, resolveKlasseFach)`).
- **Phase 1b** _(in Arbeit)_ – Restliche ESM-Tools auf zentrale Helpers
  umstellen. Reihenfolge: `arge` → `jahrgang` → `gast-einlader` →
  `personen-verwaltung` → `klassen-umbenennen` → `gast-zugaenge`.
- **Phase 2** – Top-5-Monolithen analog aufsplitten
  (setup-wizard, tenant-settings-ui, gast-einlader,
  schueler-lehrer-gruppen, personen-verwaltung). Konkrete Reihenfolge:
  1. `gast-einlader` (klar abgegrenzte Pure-Logik, gute
     Vorlage für mittelgroße Tools)
  2. `personen-verwaltung` (analog Schulstruktur-Sync, mittlere Größe)
  3. `schueler-lehrer-gruppen` (bereits Teil-Split, Rest finishen)
  4. `tenant-settings-ui` (IIFE → ESM, größerer Schritt)
  5. `setup-wizard` (IIFE → ESM, mit HTML-`type="module"`-Migration)
- **Phase 3** – Mittelgewichte (500–1100 Z.) konsequent nach Schema
- **Phase 4** – Restliche Legacy-IIFE-Tools auf ESM heben
  (HTML-Migration auf `type="module"`)

Jede Phase ist in sich abgeschlossen und einzeln releasebar.

---

## 8. Lessons Learned aus dem Pilot

1. **Lazy `window`-Lookups als Migrationsbrücke**: Solange IIFE-Module
   noch `window.ms365…` exportieren, kapselt eine kleine `getXxx()`-Funktion
   im neuen ESM-Modul den Zugriff (`getInferRootForType()`,
   `getTenantSettingsLoad()`, `getRulesResolver()`). Damit kann ein Modul
   **schon heute** ESM sein, obwohl seine Dependencies noch nicht migriert
   sind.
2. **Dependency Injection statt Globals**: Tool-spezifische Helper als
   **Parameter** durchreichen (`buildKursteamCsv(…, resolveKlasseFach)`,
   `computeTenantCreateSuggestionPure(…, resolveKlasseFach)`,
   `renderGraphView(…, normRoleKey)`). Pure-Tests brauchen kein
   `window`-Mock.
3. **Pure-Version + Wrapper-Pattern**: Im Modul liegt die pure Funktion
   (`applyMatchLinkUpdate(currentLinks, …)`), im Hauptfile bleibt der
   schmale Wrapper, der State lädt/speichert. So bleibt **die Logik**
   testbar, **die UI-Anbindung** isoliert.
4. **Move first, refactor later**: Erst 1:1 verschieben, danach – wenn
   überhaupt – Verbesserungen. Sonst sind PRs nicht reviewbar.
5. **ESLint `max-lines` als Frühwarnsystem**: Funktioniert in der Praxis
   sehr gut. Auf `warn` statt `error` lassen, damit der Build grün bleibt
   und das Team selbst entscheidet, was als nächstes anstehen sollte.

---

## 9. ESM-Migration: Schritt-für-Schritt-Rezept

Für **Einzeldatei-Tools** (z. B. `postfaecher.js`, `leere-gruppen-report.js`,
`gruppenerstellung-policy.js`) ist die Migration in 5 Schritten möglich:

1. **Imports an die Spitze** statt lokaler Duplikate:

   ```javascript
   import { safeJsonParse } from '../../shared/utils/json.js';
   import { getEl } from '../../shared/utils/dom.js';
   import { compareDe } from '../../shared/utils/strings.js';
   import { dlgConfirm } from '../../shared/utils/dialog.js';
   ```

2. **IIFE-Wrapper entfernen**: `(function () { 'use strict'; ` am Anfang,
   `})();` am Ende. Ggf. mit einer Massenoperation Einrückung um 4 Spaces
   reduzieren, dabei aber **vorsichtig** sein, dass Funktion-Bodies und
   verschachtelte Blöcke korrekt eingerückt bleiben (das PowerShell-Skript
   greift nur am Top-Level, innere Indents bleiben proportional).

3. **HTML-Tag umstellen**:

   ```html
   <!-- vorher -->
   <script src="../src/tools/<tool>/<tool>.js" defer></script>
   <!-- nachher -->
   <script type="module" src="../src/tools/<tool>/<tool>.js"></script>
   ```

4. **Build prüfen**: `npm run build`. Es muss ein eigener Bundle für das
   Tool entstehen, und die `shared/utils/*` werden als kleine Chunks
   herausgezogen (z. B. `dom-…js`, `strings-…js`, `dialog-…js`).

5. **Smoke-Test**: Tool-Seite manuell aufrufen, Klicks/Workflows
   durchgehen. Tests laufen lassen (`npm test`).

Für **Multi-File-Tools** (z. B. `arge` mit `-parse`, `-state`, `-csv`,
`-graph`, `-standalone-ps1`) müssen **alle** Submodule **gleichzeitig** auf
ESM umgestellt werden, da der HTML-Tag nur **einen** ESM-Entry hat und der
Rest per `import` gezogen wird. Cross-Refs (`window.ms365ArgeParse.normStr`)
werden durch echte `import`-Statements ersetzt. Größerer Brocken pro Tool,
aber gut planbar.

---

## 10. Nächste Schritte (Roadmap, Stand 2026-05-12)

### Kurzfristig (1–2 weitere Sessions)
- ESM-Migration **weiterer Einzeldatei-Tools**:
  `personen-verwaltung.js` (1208 Z.),
  `organisations-assistent.js`,
  `teams-archiv.js`,
  `verteilerlisten.js`.
  Pro Tool ca. 30 Minuten, lokale Duplikate (`normStr`, `escapeHtml`,
  `safeJsonParse`, `loadJson`, `dlgConfirm`, …) entfernen.

### Mittelfristig (3–6 Sessions)
- ESM-Migration **der verbleibenden Multi-File-Tools**:
  1. ~~`arge`~~ **fertig** (6 Files)
  2. ~~`jahrgang`~~ **fertig** (5 Files)
  3. ~~`weitere-teams-gruppen`~~ **fertig** (4 Files, 5 HTML-Entries)
  4. ~~`kursteams`~~ **fertig** (20 Files, größter Brocken)
  5. `schueler-lehrer-gruppen` (2 Files: `schueler-lehrer-gruppen.js` +
     `slg-gruppenverwaltung.js`)
  6. `schulstruktur-sync` (2 Files: `schulstruktur-sync.js` +
     `structure-rules.js`) — Hauptdatei ist bereits ESM-ähnlich
     (16 thematische Module), nur der IIFE-Wrapper muss raus
  7. `leere-gruppen-report` (2 Files: `leere-gruppen-report.js` ESM +
     `leere-gruppen-core.js` IIFE)
  8. `sharepoint/` (5 Files, alle IIFE, jeweils unabhängig)
- **Splitting** der noch unsplitteten Tools nach Schulstruktur-Sync-Pattern:
  `gast-einlader`, `schueler-lehrer-gruppen` (Rest),
  `personen-verwaltung`, `klassen-umbenennen` (ggf. parallel zur
  Migration).

### Langfristig
- IIFE-Migration der **shared Module** auf ESM:
  - `tenant-settings-ui.js` (2450 Z.) – größter Hebel, aber komplex
  - `tenant-settings-core.js` (314 Z.) – Vorstufe
  - `app-data-v2.js` (638 Z.)
  - `graph-unified-groups.js` (448 Z.)
  - `msal-auth-ui.js` (485 Z.) – nutzt schon `type="module"`!
- Splitting des **3485-Zeilen-`setup-wizard.js`** (ESM, aber Monolith).

### Definition of Done für „migrationsfrei"
- Alle Tool-Files unter 650 Zeilen
- Alle Tool-Files ESM
- Alle HTML-Tags `type="module"`
- Lokale Duplikate von `normStr`/`escapeHtml`/`safeJsonParse`/
  `dlgConfirm` etc. = **0**
- `npm run build` ≤ 750 ms, `npm test` ≤ 1 s
