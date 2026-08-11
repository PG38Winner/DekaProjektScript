import { test, before, after, describe } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, rm, readdir } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';

import { anonymizeWorkbook, inspectWorkbook } from '../src/anonymize.js';
import { createFixture, readSheet, ROWS } from './fixture.js';

let workdir;
let source;

before(async () => {
  workdir = await mkdtemp(path.join(tmpdir(), 'anonymizer-'));
  source = path.join(workdir, 'kunden.xlsx');
  await createFixture(source);
});

after(async () => {
  await rm(workdir, { recursive: true, force: true });
});

/** Legt fuer jeden Test eine frische Kopie der Beispieldatei an. */
async function freshFile(name) {
  const file = path.join(workdir, name);
  await createFixture(file);
  return file;
}

describe('inspectWorkbook', () => {
  test('erkennt Blaetter, Spalten und Inhaltsarten', async () => {
    const sheets = await inspectWorkbook(source);

    assert.deepEqual(sheets.map((s) => s.name), ['Kunden', 'Bestellungen', 'Hinweise']);

    const kunden = sheets[0];
    assert.equal(kunden.rowCount, ROWS.length);

    const kinds = Object.fromEntries(kunden.columns.map((c) => [c.header, c.kind]));
    assert.equal(kinds['E-Mail'], 'email');
    assert.equal(kinds.Telefon, 'phone');
    assert.equal(kinds.Vorname, 'firstName');
    assert.equal(kinds.Nachname, 'lastName');
    assert.equal(kinds.PLZ, 'zip');
    assert.equal(kinds.Ort, 'city');
    assert.equal(kinds.Strasse, 'street');
    assert.equal(kinds.Geburtsdatum, 'date');
    assert.equal(kinds.Umsatz, 'decimal');
    assert.equal(kinds.Aktiv, 'boolean');
    assert.equal(kinds.KundenID, 'integer');
  });

  test('markiert Formelspalten als schreibgeschuetzt', async () => {
    const [kunden] = await inspectWorkbook(source);
    const summe = kunden.columns.find((c) => c.header === 'Summe');
    assert.equal(summe.readOnly, true);
  });
});

describe('anonymizeWorkbook', () => {
  test('ersetzt nicht ausgenommene Spalten und laesst --keep unveraendert', async () => {
    const file = await freshFile('keep.xlsx');
    await anonymizeWorkbook({ file, keep: ['KundenID'], backup: false, seed: 42 });

    const rows = await readSheet(file);
    assert.equal(rows.length, ROWS.length);

    rows.forEach((row, index) => {
      assert.equal(row.KundenID, ROWS[index][0], 'KundenID muss erhalten bleiben');
      assert.notEqual(row.Vorname, ROWS[index][1]);
      assert.notEqual(row.Nachname, ROWS[index][2]);
      assert.notEqual(row.Ort, ROWS[index][7]);
    });
  });

  test('behaelt die Datentypen bei', async () => {
    const file = await freshFile('typen.xlsx');
    await anonymizeWorkbook({ file, backup: false, seed: 7 });

    const rows = await readSheet(file);
    for (const row of rows) {
      assert.equal(typeof row.KundenID, 'number');
      assert.equal(typeof row.Umsatz, 'number');
      assert.equal(typeof row.Aktiv, 'boolean');
      assert.ok(row.Geburtsdatum instanceof Date, 'Datum bleibt ein Datum');
      assert.match(String(row['E-Mail']), /^[^\s@]+@example\.com$/);
      assert.match(String(row.Telefon), /^\+49 /);
    }
  });

  test('behaelt Zahlen in aehnlicher Groessenordnung', async () => {
    const file = await freshFile('groesse.xlsx');
    await anonymizeWorkbook({ file, backup: false, seed: 3 });

    const rows = await readSheet(file);
    rows.forEach((row, index) => {
      const original = ROWS[index][9];
      assert.ok(
        row.Umsatz >= original / 2 - 1 && row.Umsatz <= original * 2 + 1,
        `Umsatz ${row.Umsatz} liegt nicht in der Groessenordnung von ${original}`,
      );
    });
  });

  test('laesst Formelspalten unangetastet', async () => {
    const file = await freshFile('formel.xlsx');
    await anonymizeWorkbook({ file, backup: false, seed: 11 });

    const rows = await readSheet(file);
    for (const row of rows) {
      assert.ok(row.Summe && typeof row.Summe === 'object' && 'formula' in row.Summe,
        'Formel muss erhalten bleiben');
    }
  });

  test('erhaelt Zellformate', async () => {
    const file = await freshFile('format.xlsx');
    await anonymizeWorkbook({ file, backup: false, seed: 5 });

    const rows = await readSheet(file);
    for (const row of rows) {
      assert.equal(row.__row.getCell(9).numFmt, 'dd.mm.yyyy');
      assert.equal(row.__row.getCell(10).numFmt, '#,##0.00 "EUR"');
    }
  });

  test('erhaelt leere Werte', async () => {
    const file = await freshFile('leer.xlsx');
    await anonymizeWorkbook({ file, backup: false, seed: 9 });

    const rows = await readSheet(file);
    const empty = rows[3].Bemerkung;
    assert.ok(empty === null || empty === undefined || empty === '',
      `Leerwert wurde ueberschrieben: ${JSON.stringify(empty)}`);
  });

  test('vergibt fuer gleiche Werte denselben Ersatz', async () => {
    const file = await freshFile('konsistent.xlsx');
    await anonymizeWorkbook({ file, backup: false, seed: 13 });

    const rows = await readSheet(file);
    // Zeile 1 und 5 der Beispieldatei enthalten dieselbe Person.
    assert.equal(rows[0].Vorname, rows[4].Vorname);
    assert.equal(rows[0].Nachname, rows[4].Nachname);
    assert.equal(String(rows[0]['E-Mail']), String(rows[4]['E-Mail']));
  });

  test('--no-consistent trennt gleiche Werte auf', async () => {
    const file = await freshFile('inkonsistent.xlsx');
    await anonymizeWorkbook({ file, backup: false, seed: 13, consistent: false });

    const rows = await readSheet(file);
    assert.notEqual(
      `${rows[0].Vorname}|${rows[0].Nachname}`,
      `${rows[4].Vorname}|${rows[4].Nachname}`,
    );
  });

  test('liefert mit gleichem Seed reproduzierbare Ergebnisse', async () => {
    const first = await freshFile('seed-a.xlsx');
    const second = await freshFile('seed-b.xlsx');
    await anonymizeWorkbook({ file: first, backup: false, seed: 2024 });
    await anonymizeWorkbook({ file: second, backup: false, seed: 2024 });

    const rowsA = await readSheet(first);
    const rowsB = await readSheet(second);
    assert.deepEqual(rowsA.map((r) => r.Nachname), rowsB.map((r) => r.Nachname));
  });

  test('legt standardmaessig eine Sicherungskopie an', async () => {
    const file = await freshFile('backup.xlsx');
    const report = await anonymizeWorkbook({ file, seed: 1 });

    assert.ok(report.backup, 'Bericht muss die Sicherungskopie nennen');
    const files = await readdir(workdir);
    assert.ok(files.some((name) => name.startsWith('backup.backup-')));

    // Die Sicherung enthaelt noch die Originalwerte.
    const backupRows = await readSheet(report.backup);
    assert.equal(backupRows[0].Vorname, ROWS[0][1]);
  });

  test('schreibt mit --out in eine neue Datei', async () => {
    const file = await freshFile('quelle.xlsx');
    const out = path.join(workdir, 'ziel.xlsx');
    const report = await anonymizeWorkbook({ file, out, backup: true, seed: 1 });

    assert.equal(report.output, out);
    assert.equal(report.backup, null, 'ohne Ueberschreiben ist keine Sicherung noetig');

    const originalRows = await readSheet(file);
    assert.equal(originalRows[0].Vorname, ROWS[0][1], 'Original bleibt unveraendert');

    const outRows = await readSheet(out);
    assert.notEqual(outRows[0].Vorname, ROWS[0][1]);
  });

  test('--dry-run veraendert die Datei nicht', async () => {
    const file = await freshFile('trocken.xlsx');
    const report = await anonymizeWorkbook({ file, dryRun: true, seed: 1 });

    assert.equal(report.output, null);
    assert.ok(report.changed > 0, 'Bericht zeigt trotzdem die geplanten Aenderungen');

    const rows = await readSheet(file);
    assert.equal(rows[0].Vorname, ROWS[0][1]);
  });

  test('--sheet beschraenkt auf das gewaehlte Arbeitsblatt', async () => {
    const file = await freshFile('blatt.xlsx');
    await anonymizeWorkbook({ file, sheet: 'Bestellungen', keep: ['BestellID'], backup: false, seed: 1 });

    const bestellungen = await readSheet(file, 'Bestellungen');
    assert.equal(bestellungen[0].BestellID, 1);
    assert.notEqual(bestellungen[0].KundenID, 1001);

    const kunden = await readSheet(file, 'Kunden');
    assert.equal(kunden[0].Vorname, ROWS[0][1], 'anderes Blatt bleibt unberuehrt');
  });

  test('verarbeitet ohne --sheet ALLE Arbeitsblaetter', async () => {
    const file = await freshFile('alle-blaetter.xlsx');
    const report = await anonymizeWorkbook({ file, backup: false, seed: 4 });

    assert.deepEqual(
      report.sheets.map((s) => s.name),
      ['Kunden', 'Bestellungen', 'Hinweise'],
      'der Bericht muss alle Blaetter nennen',
    );

    const kunden = await readSheet(file, 'Kunden');
    assert.notEqual(kunden[0].Nachname, ROWS[0][2], 'erstes Blatt anonymisiert');

    const bestellungen = await readSheet(file, 'Bestellungen');
    assert.notEqual(bestellungen[0].BestellID, 1, 'zweites Blatt ebenfalls anonymisiert');
    assert.notEqual(bestellungen[0].Betrag, 99.9);
  });

  test('haelt Verknuepfungen zwischen Blaettern zusammen', async () => {
    const file = await freshFile('verknuepfung.xlsx');
    await anonymizeWorkbook({ file, backup: false, seed: 21 });

    const kunden = await readSheet(file, 'Kunden');
    const bestellungen = await readSheet(file, 'Bestellungen');

    // KundenID 1001 steht in Zeile 1 der Kunden und in den Bestellungen 1 und 3.
    assert.equal(
      bestellungen[0].KundenID, kunden[0].KundenID,
      'dieselbe KundenID muss blattuebergreifend denselben Ersatzwert erhalten',
    );
    assert.equal(bestellungen[2].KundenID, kunden[0].KundenID);

    // KundenID 1002 steht in Zeile 2 der Kunden und in Bestellung 2.
    assert.equal(bestellungen[1].KundenID, kunden[1].KundenID);
    assert.notEqual(bestellungen[0].KundenID, bestellungen[1].KundenID,
      'verschiedene IDs bleiben verschieden');
  });

  test('ueberspringt Blaetter ohne Kopfzeile, statt abzubrechen', async () => {
    const file = await freshFile('ohne-kopfzeile.xlsx');
    const report = await anonymizeWorkbook({ file, backup: false, seed: 5 });

    const hinweise = report.sheets.find((s) => s.name === 'Hinweise');
    assert.match(hinweise.skipped, /keine Kopfzeile/);
    assert.equal(hinweise.changed, 0);
    assert.ok(report.changed > 0, 'die uebrigen Blaetter werden trotzdem verarbeitet');
  });

  test('akzeptiert --keep fuer eine Spalte, die nur ein Blatt hat', async () => {
    const file = await freshFile('keep-anderes-blatt.xlsx');
    await anonymizeWorkbook({ file, keep: ['BestellID'], backup: false, seed: 6 });

    const bestellungen = await readSheet(file, 'Bestellungen');
    assert.equal(bestellungen[0].BestellID, 1, 'BestellID bleibt erhalten');

    const kunden = await readSheet(file, 'Kunden');
    assert.notEqual(kunden[0].Nachname, ROWS[0][2], 'Kunden werden weiterhin anonymisiert');
  });

  test('meldet unbekannte Spalten in --keep', async () => {
    const file = await freshFile('unbekannt.xlsx');
    await assert.rejects(
      () => anonymizeWorkbook({ file, keep: ['Gibtsnicht'], backup: false }),
      /Unbekannte Spalte/,
    );
  });

  test('meldet ein unbekanntes Arbeitsblatt', async () => {
    await assert.rejects(
      () => anonymizeWorkbook({ file: source, sheet: 'Fehlt', dryRun: true }),
      /nicht gefunden/,
    );
  });

  test('meldet eine fehlende Datei', async () => {
    await assert.rejects(
      () => anonymizeWorkbook({ file: path.join(workdir, 'weg.xlsx'), dryRun: true }),
      /nicht gefunden/,
    );
  });
});
