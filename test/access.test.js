/**
 * Prueft den Access-Weg.
 *
 * plan.js arbeitet auf einer schlichten Beschreibung aus Tabellen, Spalten und
 * Zeilen - genau der Form, die read.js aus mdb-reader erzeugt. Dadurch laesst
 * sich der gesamte Ablauf bis zur geschriebenen Arbeitsmappe pruefen; nur das
 * Einlesen einer echten .accdb bleibt aussen vor.
 */

import { test, describe, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import ExcelJS from 'exceljs';

import { planAnonymization, chooseKeyColumn } from '../src/access/plan.js';
import { writeWorkbook, sheetNameFor } from '../src/access/export.js';

let workdir;

before(async () => {
  workdir = await mkdtemp(path.join(tmpdir(), 'anonymizer-access-'));
});

after(async () => {
  await rm(workdir, { recursive: true, force: true });
});

function column(name, type, extra = {}) {
  return {
    name,
    type,
    size: 255,
    nullable: true,
    autoNumber: false,
    readOnly: false,
    readOnlyReason: null,
    ...extra,
  };
}

/** Tabelle "Kunden" mit Autowert-Schluessel und gemischten Typen. */
function kundenTable(rows) {
  return {
    name: 'Kunden',
    rowCount: rows.length,
    columns: [
      column('KundenID', 'Long', { autoNumber: true, readOnly: true, readOnlyReason: 'Autowert' }),
      column('Vorname', 'Text'),
      column('Nachname', 'Text'),
      column('EMail', 'Text'),
      column('Geburtsdatum', 'DateTime'),
      column('Umsatz', 'Currency'),
      column('Aktiv', 'Boolean'),
      column('Anhang', 'OLE', { readOnly: true, readOnlyReason: 'Typ OLE' }),
    ],
    rows,
  };
}

const KUNDEN_ROWS = [
  {
    KundenID: 1, Vorname: 'Anna', Nachname: 'Müller', EMail: 'anna@firma.de',
    Geburtsdatum: new Date('1985-03-14'), Umsatz: 1499.99, Aktiv: true,
  },
  {
    KundenID: 2, Vorname: 'Bernd', Nachname: 'Schmidt', EMail: 'bernd@firma.de',
    Geburtsdatum: new Date('1972-11-02'), Umsatz: 240.5, Aktiv: false,
  },
  {
    KundenID: 3, Vorname: 'Anna', Nachname: 'Müller', EMail: 'anna@firma.de',
    Geburtsdatum: new Date('1985-03-14'), Umsatz: 55, Aktiv: true,
  },
];

/** Tabelle "Bestellungen", ueber KundenID mit den Kunden verknuepft. */
function bestellungenTable() {
  return {
    name: 'Bestellungen',
    rowCount: 3,
    columns: [
      column('BestellNr', 'Long', { autoNumber: true, readOnly: true, readOnlyReason: 'Autowert' }),
      column('KundenID', 'Long'),
      column('Betrag', 'Currency'),
    ],
    rows: [
      { BestellNr: 10, KundenID: 1, Betrag: 99.9 },
      { BestellNr: 11, KundenID: 2, Betrag: 12.5 },
      { BestellNr: 12, KundenID: 1, Betrag: 44 },
    ],
  };
}

/** Liest ein geschriebenes Blatt als Objekte zurueck. */
async function readSheet(file, name) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(file);
  const sheet = workbook.getWorksheet(name);

  const headers = [];
  sheet.getRow(1).eachCell((cell, index) => { headers[index] = String(cell.value); });

  const rows = [];
  sheet.eachRow((row, number) => {
    if (number === 1) return;
    const record = {};
    headers.forEach((header, index) => { if (header) record[header] = row.getCell(index).value; });
    rows.push(record);
  });
  return rows;
}

describe('Schluesselspalte finden', () => {
  test('bevorzugt den Autowert', () => {
    const table = kundenTable(KUNDEN_ROWS);
    assert.equal(chooseKeyColumn(table, table.columns).name, 'KundenID');
  });

  test('nimmt sonst eine Spalte mit durchgaengig eindeutigen Werten', () => {
    const table = {
      name: 'T',
      rowCount: 2,
      columns: [column('Kennung', 'Text'), column('Ort', 'Text')],
      rows: [{ Kennung: 'A', Ort: 'Berlin' }, { Kennung: 'B', Ort: 'Berlin' }],
    };
    assert.equal(chooseKeyColumn(table, table.columns).name, 'Kennung');
  });

  test('bevorzugt bei mehreren eindeutigen Spalten einen Schluesselnamen', () => {
    const table = {
      name: 'T',
      rowCount: 2,
      columns: [column('Bezeichnung', 'Text'), column('ArtikelNr', 'Text')],
      rows: [
        { Bezeichnung: 'Schraube', ArtikelNr: 'A-1' },
        { Bezeichnung: 'Mutter', ArtikelNr: 'A-2' },
      ],
    };
    assert.equal(chooseKeyColumn(table, table.columns).name, 'ArtikelNr');
  });

  test('verwirft Spalten mit Luecken', () => {
    const table = {
      name: 'T',
      rowCount: 2,
      columns: [column('Kennung', 'Text')],
      rows: [{ Kennung: 'A' }, { Kennung: null }],
    };
    assert.equal(chooseKeyColumn(table, table.columns), null);
  });
});

describe('Anonymisieren der Tabellen', () => {
  test('laesst Schluessel- und Autowertspalten unveraendert', () => {
    const { report, result } = planAnonymization([kundenTable(KUNDEN_ROWS)], { seed: 1 });

    const status = Object.fromEntries(report.sheets[0].columns.map((c) => [c.header, c.status]));
    assert.equal(status.KundenID, 'unveraendert (Schluessel)');
    assert.equal(status.Vorname, 'anonymisiert');

    result.tables[0].rows.forEach((row, index) => {
      assert.equal(row.KundenID, KUNDEN_ROWS[index].KundenID);
    });
  });

  test('ueberspringt Spaltentypen, die sich nicht ersetzen lassen', () => {
    const { report } = planAnonymization([kundenTable(KUNDEN_ROWS)], { seed: 1 });
    const anhang = report.sheets[0].columns.find((c) => c.header === 'Anhang');
    assert.equal(anhang.status, 'uebersprungen (Typ OLE)');
  });

  test('liefert vollstaendige Datensaetze, nicht nur die Aenderungen', () => {
    const { result } = planAnonymization([kundenTable(KUNDEN_ROWS)], { seed: 2 });
    const table = result.tables[0];

    assert.equal(table.rows.length, KUNDEN_ROWS.length);
    assert.deepEqual(
      table.columns.map((c) => c.name),
      ['KundenID', 'Vorname', 'Nachname', 'EMail', 'Geburtsdatum', 'Umsatz', 'Aktiv', 'Anhang'],
    );
  });

  test('behaelt die Datentypen bei', () => {
    const { result } = planAnonymization([kundenTable(KUNDEN_ROWS)], { seed: 3 });
    const [first] = result.tables[0].rows;

    assert.equal(typeof first.Vorname, 'string');
    assert.equal(typeof first.Umsatz, 'number');
    assert.equal(typeof first.Aktiv, 'boolean');
    assert.ok(first.Geburtsdatum instanceof Date);
    assert.match(first.EMail, /@example\.com$/);
  });

  test('vergibt fuer gleiche Werte denselben Ersatz', () => {
    const { result } = planAnonymization([kundenTable(KUNDEN_ROWS)], { seed: 4 });
    const [a, , c] = result.tables[0].rows;

    // Datensatz 1 und 3 sind dieselbe Person.
    assert.equal(a.Vorname, c.Vorname);
    assert.equal(a.Nachname, c.Nachname);
    assert.equal(a.EMail, c.EMail);
  });

  test('haelt Verknuepfungen zwischen Tabellen zusammen', () => {
    const { report, result } = planAnonymization(
      [kundenTable(KUNDEN_ROWS), bestellungenTable()],
      { seed: 5 },
    );

    // KundenID ist in "Kunden" Schluessel; in "Bestellungen" verweist die
    // gleichnamige Spalte darauf und darf deshalb nicht ersetzt werden.
    const bestellungen = result.tables.find((t) => t.name === 'Bestellungen');
    assert.deepEqual(bestellungen.rows.map((r) => r.KundenID), [1, 2, 1]);

    const status = report.sheets
      .find((s) => s.name === 'Bestellungen').columns
      .find((c) => c.header === 'KundenID').status;
    assert.equal(status, 'unveraendert (Verknuepfung)');
  });

  test('laesst NULL- und Leerwerte stehen', () => {
    const rows = [
      { KundenID: 1, Vorname: 'Anna', Nachname: null, EMail: '', Umsatz: 10, Aktiv: true },
      { KundenID: 2, Vorname: 'Bernd', Nachname: 'Schmidt', EMail: 'b@x.de', Umsatz: 20, Aktiv: false },
    ];
    const { result } = planAnonymization([kundenTable(rows)], { seed: 6 });
    const [first] = result.tables[0].rows;

    assert.equal(first.Nachname, null, 'NULL bleibt NULL');
    assert.equal(first.EMail, '', 'Leerwert bleibt leer');
    assert.notEqual(first.Vorname, 'Anna');
  });

  test('anonymisiert auch Tabellen ohne Schluessel', () => {
    // Frueher wurden sie uebersprungen, weil kein Datensatz ansteuerbar war.
    // Geschrieben wird jetzt eine neue Datei - dann gibt es nichts anzusteuern.
    const table = {
      name: 'Protokoll',
      rowCount: 2,
      columns: [column('Ort', 'Text'), column('Notiz', 'Text')],
      rows: [{ Ort: 'Berlin', Notiz: 'aaa' }, { Ort: 'Berlin', Notiz: 'bbb' }],
    };

    const { report, result } = planAnonymization([table], { seed: 8 });

    assert.equal(report.sheets[0].skipped, null);
    assert.equal(result.tables[0].rows.length, 2);
    assert.notEqual(result.tables[0].rows[0].Ort, 'Berlin');
  });

  test('beachtet --keep je Tabelle', () => {
    const { result } = planAnonymization(
      [kundenTable(KUNDEN_ROWS), bestellungenTable()],
      { keep: ['Kunden:Nachname,Aktiv'], seed: 7 },
    );

    result.tables[0].rows.forEach((row, index) => {
      assert.equal(row.Nachname, KUNDEN_ROWS[index].Nachname);
      assert.equal(row.Aktiv, KUNDEN_ROWS[index].Aktiv);
    });
  });

  test('liefert mit gleichem Seed dasselbe Ergebnis', () => {
    const a = planAnonymization([kundenTable(KUNDEN_ROWS)], { seed: 99 });
    const b = planAnonymization([kundenTable(KUNDEN_ROWS)], { seed: 99 });
    assert.deepEqual(
      a.result.tables[0].rows.map((r) => r.Nachname),
      b.result.tables[0].rows.map((r) => r.Nachname),
    );
  });

  test('meldet eine unbekannte Tabelle', () => {
    assert.throws(
      () => planAnonymization([kundenTable(KUNDEN_ROWS)], { table: 'Gibtsnicht' }),
      /nicht gefunden/,
    );
  });

  test('verarbeitet mit --table nur die genannte Tabelle', () => {
    const { result } = planAnonymization(
      [kundenTable(KUNDEN_ROWS), bestellungenTable()],
      { table: 'Bestellungen', seed: 10 },
    );
    assert.deepEqual(result.tables.map((t) => t.name), ['Bestellungen']);
  });
});

describe('Schreiben als Arbeitsmappe', () => {
  test('schreibt je Tabelle ein Blatt mit Kopfzeile', async () => {
    const { result } = planAnonymization(
      [kundenTable(KUNDEN_ROWS), bestellungenTable()],
      { seed: 12 },
    );
    const file = path.join(workdir, 'ergebnis.xlsx');
    await writeWorkbook(file, result);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(file);
    assert.deepEqual(workbook.worksheets.map((s) => s.name), ['Kunden', 'Bestellungen']);

    const kunden = await readSheet(file, 'Kunden');
    assert.equal(kunden.length, KUNDEN_ROWS.length);
    assert.equal(kunden[0].KundenID, 1);
    assert.notEqual(kunden[0].Vorname, 'Anna');
  });

  test('behaelt die Datentypen in der Datei', async () => {
    const { result } = planAnonymization([kundenTable(KUNDEN_ROWS)], { seed: 13 });
    const file = path.join(workdir, 'typen.xlsx');
    await writeWorkbook(file, result);

    const [first] = await readSheet(file, 'Kunden');
    assert.equal(typeof first.Umsatz, 'number');
    assert.equal(typeof first.Aktiv, 'boolean');
    assert.ok(first.Geburtsdatum instanceof Date, 'Datum bleibt ein Datum');
  });

  test('kuerzt zu lange Tabellennamen und meldet es', async () => {
    const long = 'Eine-sehr-lange-Tabellenbezeichnung-die-Excel-nicht-mag';
    const table = {
      name: long,
      rowCount: 1,
      columns: [column('Ort', 'Text')],
      rows: [{ Ort: 'Berlin' }],
    };

    const { result } = planAnonymization([table], { seed: 14 });
    const file = path.join(workdir, 'langer-name.xlsx');
    const notes = await writeWorkbook(file, result);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(file);
    assert.ok(workbook.worksheets[0].name.length <= 31);
    assert.ok(notes.some((note) => note.includes(long)));
  });

  test('vergibt eindeutige Blattnamen', () => {
    const used = new Set();
    assert.equal(sheetNameFor('Kunden', used), 'Kunden');
    assert.equal(sheetNameFor('Kunden', used), 'Kunden_2');
    assert.equal(sheetNameFor('Kunden', used), 'Kunden_3');
  });

  test('ersetzt Zeichen, die Excel in Blattnamen verbietet', () => {
    assert.equal(sheetNameFor('Kunden/Adressen', new Set()), 'Kunden_Adressen');
    assert.equal(sheetNameFor('Liste[2024]', new Set()), 'Liste_2024_');
  });
});
