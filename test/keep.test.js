import { test, before, after, describe } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';

import { anonymizeWorkbook } from '../src/excel/anonymize.js';
import { parseKeep, isKept, splitList } from '../src/core/keep.js';
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

  test('mehrere Spalten je Blatt ueber getrennte Angaben', () => {
    const rules = parseKeep(['Kunden:ID', 'Kunden:Nummer']);
    assert.equal(isKept(rules, 'Kunden', 'ID'), true);
    assert.equal(isKept(rules, 'Kunden', 'Nummer'), true);
  });

  test('das Blatt-Praefix gilt fuer die folgenden Spalten weiter', () => {
    const rules = parseKeep(['Kunden:ID,Name,Ort']);

    for (const column of ['ID', 'Name', 'Ort']) {
      assert.equal(isKept(rules, 'Kunden', column), true, `${column} sollte in Kunden bleiben`);
      assert.equal(
        isKept(rules, 'Bestellungen', column), false,
        `${column} darf NICHT in anderen Blaettern gelten`,
      );
    }
  });

  test('ein neues Praefix wechselt das Blatt', () => {
    const rules = parseKeep(['Kunden:ID,Name,Artikel:Nr,Preis']);

    assert.equal(isKept(rules, 'Kunden', 'ID'), true);
    assert.equal(isKept(rules, 'Kunden', 'Name'), true);
    assert.equal(isKept(rules, 'Kunden', 'Nr'), false);
    assert.equal(isKept(rules, 'Artikel', 'Nr'), true);
    assert.equal(isKept(rules, 'Artikel', 'Preis'), true);
    assert.equal(isKept(rules, 'Artikel', 'Name'), false);
  });

  test('ohne Praefix am Anfang gelten die Spalten ueberall', () => {
    const rules = parseKeep(['ID,Nummer,Kunden:Name']);

    assert.equal(isKept(rules, 'Irgendeins', 'ID'), true);
    assert.equal(isKept(rules, 'Irgendeins', 'Nummer'), true);
    assert.equal(isKept(rules, 'Irgendeins', 'Name'), false);
    assert.equal(isKept(rules, 'Kunden', 'Name'), true);
  });

  test('"*" schaltet zurueck auf alle Blaetter', () => {
    const rules = parseKeep(['Kunden:Name,*:ID,Nummer']);

    assert.equal(isKept(rules, 'Kunden', 'Name'), true);
    assert.equal(isKept(rules, 'Bestellungen', 'Name'), false);
    assert.equal(isKept(rules, 'Bestellungen', 'ID'), true);
    assert.equal(isKept(rules, 'Bestellungen', 'Nummer'), true, '"*" gilt auch fuer Folgespalten');
  });

  test('jede --keep-Angabe beginnt wieder bei "alle Blaetter"', () => {
    const rules = parseKeep(['Kunden:ID,Name', 'Nummer']);

    assert.equal(isKept(rules, 'Kunden', 'Name'), true);
    assert.equal(isKept(rules, 'Bestellungen', 'Name'), false);
    assert.equal(isKept(rules, 'Bestellungen', 'Nummer'), true,
      'ein Praefix wirkt nicht ueber die Angabe hinaus');
  });

  test('beliebig viele Blaetter in einer einzigen Angabe', () => {
    const rules = parseKeep(['A:a1,a2,B:b1,C:c1,c2,c3,D:d1']);

    assert.equal(isKept(rules, 'A', 'a2'), true);
    assert.equal(isKept(rules, 'B', 'b1'), true);
    assert.equal(isKept(rules, 'C', 'c3'), true);
    assert.equal(isKept(rules, 'D', 'd1'), true);
    assert.equal(isKept(rules, 'C', 'b1'), false);
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

  test('haelt mehrere Spalten je Blatt in einer Angabe', async () => {
    const file = await freshFile('mehrere-spalten.xlsx');
    await anonymizeWorkbook({
      file,
      keep: ['Kunden:KundenID,Nachname,Ort', 'Bestellungen:BestellID,Betrag'],
      backup: false,
      seed: 12,
    });

    const kunden = await readSheet(file, 'Kunden');
    assert.equal(kunden[0].KundenID, ROWS[0][0]);
    assert.equal(kunden[0].Nachname, ROWS[0][2]);
    assert.equal(kunden[0].Ort, ROWS[0][7]);
    assert.notEqual(kunden[0].Vorname, ROWS[0][1], 'nicht genannte Spalten werden ersetzt');

    const bestellungen = await readSheet(file, 'Bestellungen');
    assert.equal(bestellungen[0].BestellID, 1);
    assert.equal(bestellungen[0].Betrag, 99.9);
    assert.notEqual(bestellungen[0].KundenID, 1001,
      'KundenID war nur fuer das Blatt Kunden genannt');
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
