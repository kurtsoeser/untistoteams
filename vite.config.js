import { defineConfig } from 'vite';
import { resolve } from 'node:path';

function withTrailingSlash(value) {
  if (!value) return '/';
  return value.endsWith('/') ? value : `${value}/`;
}

export default defineConfig(() => {
  // For GitHub Pages Project Pages set VITE_BASE="/<repo-name>/"
  const base = withTrailingSlash(process.env.VITE_BASE || '/');

  return {
    base,
    build: {
      outDir: 'dist',
      emptyOutDir: true,
      rollupOptions: {
        input: {
          index: resolve(__dirname, 'index.html'),
          schooltool: resolve(__dirname, 'ms365-schooltool.html'),
          tenant: resolve(__dirname, 'tenant.html'),
          help: resolve(__dirname, 'hilfe.html'),
          kursteams: resolve(__dirname, 'tools/kursteams.html'),
          jahrgang: resolve(__dirname, 'tools/jahrgang.html'),
          arge: resolve(__dirname, 'tools/arge.html'),
          gruppenerstellung: resolve(__dirname, 'tools/gruppenerstellung.html'),
          teamsArchiv: resolve(__dirname, 'tools/teams-archiv.html'),
          gruppenUebersicht: resolve(__dirname, 'tools/gruppen-uebersicht.html'),
          klassenUmbenennen: resolve(__dirname, 'tools/klassen-umbenennen.html'),
          schuelerLehrerGruppen: resolve(__dirname, 'tools/schueler-lehrer-gruppen.html'),
          weitereTeamsGruppen: resolve(__dirname, 'tools/weitere-teams-gruppen.html'),
          weitereTeamsGruppenBesitzer: resolve(__dirname, 'tools/weitere-teams-gruppen-2-besitzer.html'),
          weitereTeamsGruppenMitglieder: resolve(__dirname, 'tools/weitere-teams-gruppen-3-mitglieder.html'),
          weitereTeamsGruppenEinstellungen: resolve(__dirname, 'tools/weitere-teams-gruppen-4-einstellungen.html'),
          weitereTeamsGruppenAusfuehren: resolve(__dirname, 'tools/weitere-teams-gruppen-5-ausfuehren.html'),
          schulstrukturSync: resolve(__dirname, 'tools/schulstruktur-sync.html')
        }
      }
    }
  };
});

