/**
 * Prueft den sparsamen Leser fuer `--list`.
 *
 * Zwei Eigenschaften muessen stimmen: Er muss dasselbe erkennen wie das
 * vollstaendige Einlesen, und er muss mit Dateien zurechtkommen, an denen
 * dieses scheitert.
 */

import { test, before, after, describe } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import ExcelJS from 'exceljs';

import { inspectWorkbook } from '../src/excel/inspect.js';
import { createFixture } from './fixture.js';

let workdir;
let large;

before(async () => {
  workdir = await mkdtemp(path.join(tmpdir(), 'anonymizer-scan-'));
  large = path.join(workdir, 'gross.xlsx');
  await createLargeWorkbook(large, 20000);
});

after(async () => {
  await rm(workdir, { recursive: true, force: true });
});

/** Schreibt eine Arbeitsmappe im Datenstrom, damit der Test selbst sparsam bleibt. */
async function createLargeWorkbook(file, rows) {
  const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: file });
  const sheet = workbook.addWorksheet('Gross');

  sheet.addRow(['KundenID', 'Vorname', 'Nachname', 'E-Mail', 'Geburtsdatum', 'Umsatz']).commit();
  for (let index = 0; index < rows; index += 1) {
    const row = sheet.addRow([
      1000 + index, 'Anna', 'Müller', `a${index}@firma.de`,
      new Date(1990, 0, 1 + (index % 365)), 99.5 + index,
    ]);
    row.getCell(5).numFmt = 'dd.mm.yyyy';
    row.commit();
  }

  await sheet.commit();
  await workbook.commit();
}

/** Vergleichbare Darstellung einer Analyse. */
function fingerprint(sheets) {
  return sheets.map((sheet) => [
    sheet.name,
    sheet.rowCount,
    sheet.columns.map((column) => `${column.header}=${column.kind}${column.readOnly ? '[ro]' : ''}`).join(','),
  ].join(' | '));
}

describe('Sparsamer Leser fuer --list', () => {
  test('erkennt dasselbe wie das vollstaendige Einlesen', async () => {
    const file = path.join(workdir, 'vergleich.xlsx');
    await createFixture(file);

    assert.deepEqual(
      fingerprint(await inspectWorkbook(file)),
      fingerprint(await inspectWorkbook(file, { full: true })),
    );
  });

  test('listet eine grosse Datei auf', async () => {
    const sheets = await inspectWorkbook(large);

    assert.equal(sheets.length, 1);
    assert.equal(sheets[0].name, 'Gross');
    assert.deepEqual(
      sheets[0].columns.map((column) => column.header),
      ['KundenID', 'Vorname', 'Nachname', 'E-Mail', 'Geburtsdatum', 'Umsatz'],
    );
  });

  test('erkennt Datumsspalten am Zahlenformat', async () => {
    const sheets = await inspectWorkbook(large);
    const kinds = Object.fromEntries(sheets[0].columns.map((c) => [c.header, c.kind]));

    // Ein Datum steht in der Datei als Zahl; nur das Zahlenformat verraet es.
    assert.equal(kinds.Geburtsdatum, 'date');
    assert.equal(kinds.Umsatz, 'decimal');
    assert.equal(kinds['E-Mail'], 'email');
    assert.equal(kinds.KundenID, 'integer');
  });

  test('bleibt beim Lesen sparsam', async () => {
    // Der Speicherbedarf darf an der Dateigroesse haengen, nicht an der
    // Zeilenzahl - sonst scheitert die Anzeige an grossen Arbeitsmappen.
    const before = process.memoryUsage().heapUsed;
    await inspectWorkbook(large);
    const used = (process.memoryUsage().heapUsed - before) / 1048576;

    assert.ok(used < 100, `Zuwachs von ${used.toFixed(0)} MB ist zu viel fuer 20.000 Zeilen`);
  });

  test('greift bei einem Fehler auf das vollstaendige Einlesen zurueck', async () => {
    // Keine Excel-Datei: der sparsame Weg scheitert schon am Archiv. Dass die
    // Meldung vom vollstaendigen Weg stammt, belegt, dass der Rueckfall lief -
    // der sparsame Weg ist eine Beschleunigung, keine Bedingung.
    const file = path.join(workdir, 'keine-arbeitsmappe.xlsx');
    await writeFile(file, 'das ist keine Arbeitsmappe');

    await assert.rejects(
      () => inspectWorkbook(file),
      /Datei konnte nicht gelesen werden/,
    );
  });

  test('kommt mit leeren Blaettern zurecht', async () => {
    const file = path.join(workdir, 'leer.xlsx');
    const workbook = new ExcelJS.Workbook();
    workbook.addWorksheet('Leer');
    const gefuellt = workbook.addWorksheet('Gefuellt');
    gefuellt.addRow(['Name', 'Ort']);
    gefuellt.addRow(['Anna', 'Berlin']);
    await workbook.xlsx.writeFile(file);

    const sheets = await inspectWorkbook(file);
    assert.deepEqual(sheets.map((s) => s.name), ['Leer', 'Gefuellt']);
    assert.equal(sheets[0].columns.length, 0);
    assert.deepEqual(sheets[1].columns.map((c) => c.header), ['Name', 'Ort']);
  });
});
