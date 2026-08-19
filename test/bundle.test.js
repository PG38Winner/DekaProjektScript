/**
 * Prueft das eigenstaendige Buendel aus `dist/`.
 *
 * Das Buendel ist die Fassung, die auf Rechnern ohne Internetzugang laeuft -
 * es muss sich deshalb genauso verhalten wie der Quellcode und ohne
 * `node_modules` auskommen. Ist es nicht gebaut, werden die Tests
 * uebersprungen statt zu scheitern (`npm run build` erzeugt es).
 */

import { test, before, after, describe } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, rm, mkdir, copyFile } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

import { createFixture, createMacroFixture, readSheet, readMacroState, VBA_CONTENT } from './fixture.js';

const run = promisify(execFile);
const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const BUNDLE = path.join(projectRoot, 'dist', 'maskierer.cjs');
const SOURCE = path.join(projectRoot, 'src', 'cli.js');

let workdir;

// Muss beim Laden der Datei feststehen: die Optionen von describe() werden
// ausgewertet, bevor before() laeuft.
const bundleExists = existsSync(BUNDLE);

before(async () => {
  workdir = await mkdtemp(path.join(tmpdir(), 'anonymizer-bundle-'));
});

after(async () => {
  await rm(workdir, { recursive: true, force: true });
});

/** Vergleichbare Darstellung eines Blattes. */
function fingerprint(rows) {
  return rows.map((row) => [
    row.Vorname, row.Nachname, String(row['E-Mail']), row.Telefon,
    row.Strasse, row.PLZ, row.Ort, row.Umsatz, row.Aktiv,
    row.Geburtsdatum instanceof Date ? row.Geburtsdatum.toISOString() : row.Geburtsdatum,
  ].join('|'));
}

describe('Eigenstaendiges Buendel (dist/)', { skip: !bundleExists && 'dist/ nicht gebaut - "npm run build" ausfuehren' }, () => {
  test('verhaelt sich wie der Quellcode', async () => {
    const fromSource = path.join(workdir, 'quelle.xlsx');
    const fromBundle = path.join(workdir, 'buendel.xlsx');
    await createFixture(fromSource);
    await createFixture(fromBundle);

    await run(process.execPath, [SOURCE, fromSource, '--seed', '99', '--in-place', '--yes']);
    await run(process.execPath, [BUNDLE, fromBundle, '--seed', '99', '--in-place', '--yes']);

    assert.deepEqual(
      fingerprint(await readSheet(fromBundle)),
      fingerprint(await readSheet(fromSource)),
      'Buendel und Quellcode muessen bei gleichem Seed dieselben Werte erzeugen',
    );
  });

  test('laeuft ohne node_modules', async () => {
    // Eigenes Verzeichnis ausserhalb des Projekts: dort ist keine Abhaengigkeit
    // aufloesbar, das Buendel muss alles mitbringen.
    const isolated = path.join(workdir, 'isoliert');
    await mkdir(isolated, { recursive: true });

    const copiedBundle = path.join(isolated, 'maskierer.cjs');
    const data = path.join(isolated, 'daten.xlsx');
    await copyFile(BUNDLE, copiedBundle);
    await createFixture(data);

    const { stdout } = await run(process.execPath, [copiedBundle, data, '--seed', '1', '--in-place', '--yes'], {
      cwd: isolated,
    });

    assert.match(stdout, /Maskierte Zellen gesamt:\s+\d+/);

    const rows = await readSheet(data);
    assert.notEqual(rows[0].Nachname, 'Müller');
  });

  test('zeigt mit --list alle Blaetter und Spalten an', async () => {
    const file = path.join(workdir, 'liste.xlsx');
    await createFixture(file);

    const { stdout } = await run(process.execPath, [BUNDLE, file, '--list']);

    for (const sheet of ['Kunden', 'Bestellungen', 'Hinweise']) {
      assert.match(stdout, new RegExp(`Arbeitsblatt: ${sheet}`), `${sheet} fehlt in der Ausgabe`);
    }
    for (const [column, kind] of [['E-Mail', 'email'], ['Geburtsdatum', 'date'], ['Ort', 'city']]) {
      assert.match(stdout, new RegExp(`${column}\\s+${kind}`), `${column} fehlt in der Ausgabe`);
    }
    assert.match(stdout, /Summe\s+wird uebersprungen\s+\[Formel\]/);
  });

  test('erhaelt auch im Buendel die Makros', async () => {
    const file = path.join(workdir, 'makro.xlsm');
    await createMacroFixture(file);

    const { stdout } = await run(process.execPath, [BUNDLE, file, '--seed', '1', '--in-place', '--yes']);
    assert.match(stdout, /Makros:\s+uebernommen/);

    const state = await readMacroState(file);
    assert.equal(state.vba, VBA_CONTENT);
  });

  test('meldet Fehler mit Exit-Code 1', async () => {
    const missing = path.join(workdir, 'gibtsnicht.xlsx');
    await assert.rejects(
      () => run(process.execPath, [BUNDLE, missing]),
      (error) => error.code === 1 && /nicht gefunden/.test(error.stderr),
    );
  });
});
