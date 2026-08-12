import { test, before, after, describe } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';

import { anonymizeWorkbook } from '../src/anonymize.js';
import { parseKeep, isKept, splitList } from '../src/keep.js';
import { createFixture, readSheet, ROWS } from './fixture.js';

let workdir;

before(async () => {
  workdir = await mkdtemp(path.join(tmpdir(), 'anonymizer-keep-'));
});

after(async () => {
  await rm(workdir, { recursive: true, force: true });
});

async function freshFile(name) {
  const file = path.join(workdir, name);
  await createFixture(file);
  return file;
}

describe('--keep zerlegen', () => {
  test('ohne Blattnamen gilt die Angabe fuer jedes Blatt', () => {
    const rules = parseKeep(['KundenID']);
    assert.equal(isKept(rules, 'Kunden', 'KundenID'), true);
    assert.equal(isKept(rules, 'Bestellungen', 'KundenID'), true);
    assert.equal(isKept(rules, 'Kunden', 'Vorname'), false);
  });

  test('mit Blattnamen gilt sie nur dort', () => {
    const rules = parseKeep(['Kunden:KundenID']);
    assert.equal(isKept(rules, 'Kunden', 'KundenID'), true);
    assert.equal(isKept(rules, 'Bestellungen', 'KundenID'), false);
  });

  test('beliebig viele Blaetter lassen sich angeben', () => {
    const rules = parseKeep(['A:Eins', 'B:Zwei', 'C:Drei', 'D:Vier']);
    assert.equal(isKept(rules, 'A', 'Eins'), true);
    assert.equal(isKept(rules, 'B', 'Zwei'), true);
    assert.equal(isKept(rules, 'C', 'Drei'), true);
    assert.equal(isKept(rules, 'D', 'Vier'), true);
    assert.equal(isKept(rules, 'A', 'Zwei'), false);
  });

  test('mehrere Spalten je Blatt', () => {
    const rules = parseKeep(['Kunden:ID', 'Kunden:Nummer']);
    assert.equal(isKept(rules, 'Kunden', 'ID'), true);
    assert.equal(isKept(rules, 'Kunden', 'Nummer'), true);
  });

  test('Gross- und Kleinschreibung spielt keine Rolle', () => {
    const rules = parseKeep(['kunden:kundenid']);
    assert.equal(isKept(rules, 'KUNDEN', 'KundenID'), true);
  });

  test('trennt am ersten Doppelpunkt', () => {
    const rules = parseKeep(['Kunden:Anteil:Prozent']);
    assert.equal(isKept(rules, 'Kunden', 'Anteil:Prozent'), true);
  });

  test('weist unvollstaendige Angaben zurueck', () => {
    assert.throws(() => parseKeep([':ID']), /Ungueltige --keep-Angabe/);
    assert.throws(() => parseKeep(['Kunden:']), /Ungueltige --keep-Angabe/);
  });

  test('zerlegt Kommalisten und achtet auf maskierte Kommas', () => {
    assert.deepEqual(splitList('A,B , C'), ['A', 'B', 'C']);
    assert.deepEqual(splitList('Kunden:ID,Bestellungen:Nr'), ['Kunden:ID', 'Bestellungen:Nr']);
    assert.deepEqual(splitList('Nachname\\, Vorname,Ort'), ['Nachname, Vorname', 'Ort']);
  });
});

describe('--keep je Arbeitsblatt anwenden', () => {
  test('haelt in jedem Blatt genau die genannte Spalte', async () => {
    const file = await freshFile('je-blatt.xlsx');
    await anonymizeWorkbook({
      file,
      keep: ['Kunden:KundenID', 'Bestellungen:BestellID'],
      backup: false,
      seed: 4,
    });

    const kunden = await readSheet(file, 'Kunden');
    const bestellungen = await readSheet(file, 'Bestellungen');

    assert.equal(kunden[0].KundenID, ROWS[0][0], 'KundenID im Blatt Kunden bleibt');
    assert.equal(bestellungen[0].BestellID, 1, 'BestellID im Blatt Bestellungen bleibt');

    // KundenID war nur fuer das Blatt "Kunden" ausgenommen.
    assert.notEqual(bestellungen[0].KundenID, 1001,
      'KundenID im Blatt Bestellungen wird anonymisiert');
  });

  test('mischt blattweite und blattbezogene Angaben', async () => {
    const file = await freshFile('gemischt.xlsx');
    await anonymizeWorkbook({
      file,
      keep: ['KundenID', 'Bestellungen:Betrag'],
      backup: false,
      seed: 8,
    });

    const kunden = await readSheet(file, 'Kunden');
    const bestellungen = await readSheet(file, 'Bestellungen');

    assert.equal(kunden[0].KundenID, ROWS[0][0]);
    assert.equal(bestellungen[0].KundenID, 1001, 'blattweite Angabe wirkt in beiden Blaettern');
    assert.equal(bestellungen[0].Betrag, 99.9, 'Betrag nur in Bestellungen ausgenommen');
    assert.notEqual(kunden[0].Umsatz, ROWS[0][9]);
  });

  test('meldet ein unbekanntes Arbeitsblatt', async () => {
    const file = await freshFile('falsches-blatt.xlsx');
    await assert.rejects(
      () => anonymizeWorkbook({ file, keep: ['Gibtsnicht:ID'], backup: false }),
      /wird nicht verarbeitet/,
    );
  });

  test('meldet eine Spalte, die es in diesem Blatt nicht gibt', async () => {
    const file = await freshFile('falsche-spalte.xlsx');
    await assert.rejects(
      () => anonymizeWorkbook({ file, keep: ['Bestellungen:Vorname'], backup: false }),
      /gibt es im Blatt "Bestellungen" nicht/,
    );
  });

  test('meldet eine Spalte, die in keinem Blatt vorkommt', async () => {
    const file = await freshFile('unbekannt.xlsx');
    await assert.rejects(
      () => anonymizeWorkbook({ file, keep: ['Gibtsnicht'], backup: false }),
      /kommt in keinem verarbeiteten Blatt vor/,
    );
  });
});
