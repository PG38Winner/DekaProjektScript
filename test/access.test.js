/**
 * Prueft die Access-Logik ohne Datenbank.
 *
 * plan.js arbeitet bewusst auf einer schlichten Beschreibung aus Tabellen,
 * Spalten und Zeilen - genau der Form, die read.js aus mdb-reader erzeugt.
 * Dadurch laesst sich alles pruefen, was entscheidet, WAS geaendert wird,
 * obwohl das Schreiben selbst Windows und ACE-OLEDB braucht.
 */

import { test, describe } from 'node:test';
import assert from 'node:assert/strict';

import { planAnonymization, chooseKeyColumn } from '../src/access/plan.js';
import { serializePlan, parseResult } from '../src/access/write.js';

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

describe('Schluesselspalte waehlen', () => {
  test('bevorzugt den Autowert', () => {
    const table = kundenTable(KUNDEN_ROWS);
    const key = chooseKeyColumn(table, table.columns);
    assert.equal(key.name, 'KundenID');
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

  test('findet ohne eindeutige Spalte keinen Schluessel', () => {
    const table = {
      name: 'T',
      rowCount: 2,
      columns: [column('Ort', 'Text')],
      rows: [{ Ort: 'Berlin' }, { Ort: 'Berlin' }],
    };
    assert.equal(chooseKeyColumn(table, table.columns), null);
  });
});

describe('Aenderungsplan', () => {
  test('laesst Schluessel- und Autowertspalten unveraendert', () => {
    const { report, plan } = planAnonymization([kundenTable(KUNDEN_ROWS)], { seed: 1 });

    const status = Object.fromEntries(
      report.sheets[0].columns.map((c) => [c.header, c.status]),
    );
    assert.equal(status.KundenID, 'unveraendert (Schluessel)');
    assert.equal(status.Vorname, 'anonymisiert');

    for (const update of plan.tables[0].updates) {
      assert.ok(!('KundenID' in update.v), 'der Schluessel darf nie geaendert werden');
    }
  });

  test('ueberspringt Spaltentypen, die sich nicht ersetzen lassen', () => {
    const { report } = planAnonymization([kundenTable(KUNDEN_ROWS)], { seed: 1 });
    const anhang = report.sheets[0].columns.find((c) => c.header === 'Anhang');
    assert.equal(anhang.status, 'uebersprungen (Typ OLE)');
  });

  test('steuert jeden Datensatz ueber seinen Schluessel an', () => {
    const { plan } = planAnonymization([kundenTable(KUNDEN_ROWS)], { seed: 2 });
    const table = plan.tables[0];

    assert.equal(table.key, 'KundenID');
    assert.deepEqual(table.updates.map((u) => u.k), [1, 2, 3]);
    assert.equal(table.types.Geburtsdatum, 'DateTime');
    assert.equal(table.types.Umsatz, 'Currency');
  });

  test('behaelt die Datentypen der Ersatzwerte bei', () => {
    const { plan } = planAnonymization([kundenTable(KUNDEN_ROWS)], { seed: 3 });
    const [first] = plan.tables[0].updates;

    assert.equal(typeof first.v.Vorname, 'string');
    assert.equal(typeof first.v.Umsatz, 'number');
    assert.equal(typeof first.v.Aktiv, 'boolean');
    assert.ok(first.v.Geburtsdatum instanceof Date);
    assert.match(first.v.EMail, /@example\.com$/);
  });

  test('vergibt fuer gleiche Werte denselben Ersatz', () => {
    const { plan } = planAnonymization([kundenTable(KUNDEN_ROWS)], { seed: 4 });
    const [a, , c] = plan.tables[0].updates;

    // Datensatz 1 und 3 sind dieselbe Person.
    assert.equal(a.v.Vorname, c.v.Vorname);
    assert.equal(a.v.Nachname, c.v.Nachname);
    assert.equal(a.v.EMail, c.v.EMail);
  });

  test('haelt Verknuepfungen zwischen Tabellen zusammen', () => {
    const { plan } = planAnonymization(
      [kundenTable(KUNDEN_ROWS), bestellungenTable()],
      { seed: 5 },
    );

    // KundenID ist in "Kunden" Schluessel (bleibt), in "Bestellungen" ein
    // gewoehnliches Feld - dort darf es nicht auf einen anderen Kunden zeigen.
    const bestellungen = plan.tables.find((t) => t.name === 'Bestellungen');
    const changedIds = bestellungen.updates.filter((u) => 'KundenID' in u.v);
    assert.equal(changedIds.length, 0,
      'ein Fremdschluessel auf eine unveraenderte Schluesselspalte muss stehen bleiben');
  });

  test('laesst NULL- und Leerwerte stehen', () => {
    const rows = [
      { KundenID: 1, Vorname: 'Anna', Nachname: null, EMail: '', Umsatz: 10, Aktiv: true },
      { KundenID: 2, Vorname: 'Bernd', Nachname: 'Schmidt', EMail: 'b@x.de', Umsatz: 20, Aktiv: false },
    ];
    const { plan } = planAnonymization([kundenTable(rows)], { seed: 6 });
    const [first] = plan.tables[0].updates;

    assert.ok(!('Nachname' in first.v), 'NULL bleibt NULL');
    assert.ok(!('EMail' in first.v), 'Leerwert bleibt leer');
    assert.ok('Vorname' in first.v);
  });

  test('beachtet --keep je Tabelle', () => {
    const { report, plan } = planAnonymization(
      [kundenTable(KUNDEN_ROWS), bestellungenTable()],
      { keep: ['Kunden:Nachname,Aktiv'], seed: 7 },
    );

    const kunden = report.sheets.find((s) => s.name === 'Kunden');
    const status = Object.fromEntries(kunden.columns.map((c) => [c.header, c.status]));
    assert.equal(status.Nachname, 'unveraendert');
    assert.equal(status.Aktiv, 'unveraendert');
    assert.equal(status.Vorname, 'anonymisiert');

    for (const update of plan.tables[0].updates) {
      assert.ok(!('Nachname' in update.v));
      assert.ok(!('Aktiv' in update.v));
    }
  });

  test('ruehrt eine Tabelle ohne Schluessel nicht an und sagt es', () => {
    const table = {
      name: 'Protokoll',
      rowCount: 2,
      columns: [column('Ort', 'Text'), column('Text', 'Text')],
      rows: [{ Ort: 'Berlin', Text: 'a' }, { Ort: 'Berlin', Text: 'a' }],
    };

    const { report, plan } = planAnonymization([table], { seed: 8 });

    assert.equal(plan.tables.length, 0, 'ohne Schluessel darf nichts geschrieben werden');
    assert.match(report.sheets[0].skipped, /Schluesselspalte/);
    assert.ok(report.warnings.some((w) => w.includes('Protokoll')));
  });

  test('liefert mit gleichem Seed dasselbe Ergebnis', () => {
    const a = planAnonymization([kundenTable(KUNDEN_ROWS)], { seed: 99 });
    const b = planAnonymization([kundenTable(KUNDEN_ROWS)], { seed: 99 });
    assert.deepEqual(
      a.plan.tables[0].updates.map((u) => u.v.Nachname),
      b.plan.tables[0].updates.map((u) => u.v.Nachname),
    );
  });

  test('meldet eine unbekannte Tabelle', () => {
    assert.throws(
      () => planAnonymization([kundenTable(KUNDEN_ROWS)], { table: 'Gibtsnicht' }),
      /nicht gefunden/,
    );
  });

  test('meldet eine unbekannte Spalte in --keep', () => {
    assert.throws(
      () => planAnonymization([kundenTable(KUNDEN_ROWS)], { keep: ['Kunden:Gibtsnicht'] }),
      /gibt es im Blatt "Kunden" nicht/,
    );
  });

  test('verarbeitet mit --table nur die genannte Tabelle', () => {
    const { plan } = planAnonymization(
      [kundenTable(KUNDEN_ROWS), bestellungenTable()],
      { table: 'Bestellungen', seed: 10 },
    );
    assert.deepEqual(plan.tables.map((t) => t.name), ['Bestellungen']);
  });
});

describe('Uebergabe an PowerShell', () => {
  test('wandelt Datumswerte in ISO-Zeichenketten', () => {
    const { plan } = planAnonymization([kundenTable(KUNDEN_ROWS)], { seed: 11 });
    const serialized = serializePlan(plan);
    const [first] = serialized.tables[0].updates;

    assert.equal(typeof first.v.Geburtsdatum, 'string');
    assert.match(first.v.Geburtsdatum, /^\d{4}-\d{2}-\d{2}T/);
    assert.equal(typeof first.v.Umsatz, 'number', 'Zahlen bleiben Zahlen');
    assert.ok(JSON.parse(JSON.stringify(serialized)), 'muss sich als JSON uebergeben lassen');
  });

  test('liest das Ergebnis aus der Ausgabe', () => {
    const result = parseResult('{"tables":[{"name":"Kunden","updated":3,"unmatched":0}],"errors":[]}');
    assert.equal(result.tables[0].updated, 3);
    assert.deepEqual(result.errors, []);
  });

  test('vertraegt PowerShells Einzelobjekt statt Feld', () => {
    const result = parseResult('{"tables":{"name":"Kunden","updated":1,"unmatched":0},"errors":"Fehler"}');
    assert.equal(result.tables.length, 1);
    assert.deepEqual(result.errors, ['Fehler']);
  });

  test('meldet unbrauchbare Ausgabe deutlich', () => {
    assert.throws(() => parseResult(''), /Unerwartete Ausgabe/);
  });
});
