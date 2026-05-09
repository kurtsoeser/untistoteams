import fs from 'node:fs/promises';
import path from 'node:path';

const projectRoot = path.resolve(process.cwd());
const distRoot = path.resolve(projectRoot, 'dist');

/**
 * Ergänzt den Vite-`dist/`-Output um alles, was die MPA beim statischen Hosting
 * noch braucht, Vite aber nicht (oder nicht vollständig) mitschreibt:
 * - `src/**`: HTML verweist auf `src/shared/...` und Tool-Module ohne Bundling
 *   (`<script src="...">` ohne `type="module"` bleiben externe Pfade).
 * - Root-`app.css`: u. a. `ms365-schooltool.html` nutzt weiter `href="app.css"`
 *   (nicht die gehashte `/assets/app-*.css` aus anderen Entry-HTMLs).
 * - `ms365-config*.js` / `app.js`: gleicher Grund — nicht in den Rollup-Bundle
 *   gezogen, müssen neben den HTML-Dateien liegen.
 * - `ms365-schooltool.html`: kanonische Datei aus dem Repo-Root nach `dist/`
 *   kopieren (Redirect/MSAL/mode-Map), damit der Build nicht von einer älteren
 *   Vorlage abweicht.
 * - `README.md`: optional für Deployment-Artefakte.
 */

async function exists(p) {
  try {
    await fs.access(p);
    return true;
  } catch {
    return false;
  }
}

async function copyFileToDist(relFile) {
  const src = path.resolve(projectRoot, relFile);
  const dst = path.resolve(distRoot, relFile);
  if (!(await exists(src))) return;
  await fs.mkdir(path.dirname(dst), { recursive: true });
  await fs.copyFile(src, dst);
}

async function copyDirToDist(relDir) {
  const src = path.resolve(projectRoot, relDir);
  const dst = path.resolve(distRoot, relDir);
  if (!(await exists(src))) return;
  await fs.mkdir(path.dirname(dst), { recursive: true });
  await fs.cp(src, dst, { recursive: true });
}

async function main() {
  if (!(await exists(distRoot))) {
    throw new Error('dist/ fehlt – zuerst `vite build` ausführen.');
  }

  // Nur Dateien, die im Repo-Root existieren (Shared-Module liegen unter `src/`
  // und werden mit `copyDirToDist('src')` mitkopiert).
  const rootFiles = [
    'ms365-schooltool.html',
    'app.css',
    'app.js',
    'ms365-config.js',
    'ms365-config.example.js',
    'README.md'
  ];

  for (const f of rootFiles) await copyFileToDist(f);

  await copyDirToDist('src');
}

await main();
