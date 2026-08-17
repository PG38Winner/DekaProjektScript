/**
 * Baut die Browser-Fassung nach dist/.
 *
 * Ergebnis sind zwei Dateien, die zusammen in einem Ordner liegen muessen:
 *   dist/maskierer.html   die Oberflaeche
 *   dist/maskierer.js     der gesamte Code samt Abhaengigkeiten
 *
 * Warum nicht alles in einer HTML-Datei: Der Code enthaelt in Zeichenketten
 * und regulaeren Ausdruecken die Folgen "<!--" und "-->". In ein script-Element
 * eingebettet bringen sie den HTML-Parser aus dem Tritt; maskiert man sie mit
 * einem Gegenstrich, lehnt Chrome die dadurch veraenderten regulaeren
 * Ausdruecke ab (in Node laufen sie noch, der Fehler faellt also erst im
 * Browser auf). Der Umweg ueber "eval" mit base64 waere die Alternative - das
 * ist fuer eine Sicherheitspruefung der schlechtere Tausch.
 *
 * Es wird nichts nachgeladen: die zweite Datei liegt daneben, keine Adresse
 * zeigt ins Netz.
 */

import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { build } from 'esbuild';

const HTML = 'dist/maskierer.html';
const SCRIPT = 'dist/maskierer.js';

const result = await build({
  entryPoints: ['src/web/app.js'],
  bundle: true,
  platform: 'browser',
  format: 'iife',
  target: 'es2020',
  minify: true,
  write: false,
  legalComments: 'none',
  define: { 'process.env.NODE_ENV': '"production"', global: 'globalThis' },
  // Einige Abhaengigkeiten erwarten Node-Bausteine. Die Browser-Ersatzteile
  // sind Standardpakete; ohne sie liesse sich nicht buendeln.
  alias: {
    stream: 'stream-browserify',
    events: 'events',
    buffer: 'buffer',
  },
  inject: ['src/web/shim.js'],
});

const script = result.outputFiles[0].text;
const template = await readFile('src/web/index.html', 'utf8');

if (!template.includes('<!--SCRIPT-->')) {
  throw new Error('In src/web/index.html fehlt die Stelle <!--SCRIPT-->.');
}

await mkdir('dist', { recursive: true });
await writeFile(SCRIPT, script, 'utf8');
await writeFile(
  HTML,
  template.replace('<!--SCRIPT-->', '<script src="maskierer.js" defer></script>'),
  'utf8',
);

const size = (Buffer.byteLength(script, 'utf8') / 1048576).toFixed(1);
console.log(`${HTML} und ${SCRIPT} geschrieben (${size} MB)`);
