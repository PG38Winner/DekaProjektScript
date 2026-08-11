import { test, before, after, describe } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';

import { anonymizeWorkbook } from '../src/anonymize.js';
import {
  createMacroFixture, readMacroState, readSheet, ROWS,
  VBA_CONTENT, VBA_SIGNATURE_CONTENT, WORKBOOK_CODE_NAME, SHEET_CODE_NAME,
} from './fixture.js';

let workdir;

before(async () => {
  workdir = await mkdtemp(path.join(tmpdir(), 'anonymizer-macros-'));
});

after(async () => {
  await rm(workdir, { recursive: true, force: true });
});

async function freshMacroFile(name) {
  const file = path.join(workdir, name);
  await createMacroFixture(file);
  return file;
}

describe('Makro-Arbeitsmappen (.xlsm)', () => {
  test('erhaelt das VBA-Projekt', async () => {
    const file = await freshMacroFile('makro.xlsm');
    const report = await anonymizeWorkbook({ file, backup: false, seed: 1 });

    assert.equal(report.macrosPreserved, true);

    const state = await readMacroState(file);
    assert.equal(state.vba, VBA_CONTENT, 'vbaProject.bin muss unveraendert erhalten bleiben');
  });

  test('erhaelt die Signatur des VBA-Projekts', async () => {
    const file = await freshMacroFile('signatur.xlsm');
    await anonymizeWorkbook({ file, backup: false, seed: 1 });

    const state = await readMacroState(file);
    assert.equal(state.signature, VBA_SIGNATURE_CONTENT);
    assert.match(state.contentTypes, /PartName="\/xl\/vbaProjectSignature\.bin"/);
  });

  test('anonymisiert die Daten trotzdem', async () => {
    const file = await freshMacroFile('makro-daten.xlsm');
    await anonymizeWorkbook({ file, keep: ['KundenID'], backup: false, seed: 2 });

    const rows = await readSheet(file);
    assert.equal(rows.length, ROWS.length);
    assert.equal(rows[0].KundenID, ROWS[0][0]);
    assert.notEqual(rows[0].Nachname, ROWS[0][2]);
  });

  test('setzt den makrofaehigen Arbeitsmappentyp', async () => {
    const file = await freshMacroFile('typ.xlsm');
    await anonymizeWorkbook({ file, backup: false, seed: 3 });

    const { contentTypes } = await readMacroState(file);
    assert.match(contentTypes, /application\/vnd\.ms-excel\.sheet\.macroEnabled\.main\+xml/);
    assert.doesNotMatch(
      contentTypes,
      /spreadsheetml\.sheet\.main\+xml/,
      'der Nicht-Makro-Typ darf nicht zurueckbleiben',
    );
    assert.match(contentTypes, /PartName="\/xl\/vbaProject\.bin"/);
    assert.match(contentTypes, /application\/vnd\.ms-office\.vbaProject/);
  });

  test('traegt den Verweis auf das VBA-Projekt ein', async () => {
    const file = await freshMacroFile('rels.xlsm');
    await anonymizeWorkbook({ file, backup: false, seed: 4 });

    const { workbookRels } = await readMacroState(file);
    assert.match(workbookRels, /relationships\/vbaProject/);
    assert.match(workbookRels, /Target="vbaProject\.bin"/);

    // Keine doppelt vergebenen Beziehungs-Ids.
    const ids = [...workbookRels.matchAll(/Id="(rId\d+)"/g)].map((m) => m[1]);
    assert.equal(ids.length, new Set(ids).size, `doppelte rIds: ${ids.join(', ')}`);
  });

  test('erhaelt die Codenamen von Arbeitsmappe und Blatt', async () => {
    const file = await freshMacroFile('codenamen.xlsm');
    await anonymizeWorkbook({ file, backup: false, seed: 5 });

    const state = await readMacroState(file);
    assert.match(state.workbook, new RegExp(`codeName="${WORKBOOK_CODE_NAME}"`));
    assert.match(state.sheet1, new RegExp(`codeName="${SHEET_CODE_NAME}"`));

    // sheetPr muss das erste Kindelement von worksheet sein.
    assert.match(state.sheet1, /<worksheet\b[^>]*>\s*<sheetPr\b/);
  });

  test('meldet Bestandteile, die nicht uebernommen werden konnten', async () => {
    const file = await freshMacroFile('verlust.xlsm');
    const report = await anonymizeWorkbook({ file, backup: false, seed: 6 });

    const warning = report.warnings.find((text) => text.includes('activeX'));
    assert.ok(warning, `ActiveX-Verlust muss gemeldet werden, erhalten: ${JSON.stringify(report.warnings)}`);
  });

  test('warnt, wenn das Ziel keine .xlsm-Datei ist', async () => {
    const file = await freshMacroFile('ziel.xlsm');
    const out = path.join(workdir, 'ziel.xlsx');
    const report = await anonymizeWorkbook({ file, out, backup: false, seed: 7 });

    assert.equal(report.macrosPreserved, false);
    assert.ok(
      report.warnings.some((text) => text.includes('.xlsm')),
      'Verlust der Makros muss gemeldet werden',
    );
  });

  test('laesst Dateien ohne Makros unveraendert behandeln', async () => {
    const { createFixture } = await import('./fixture.js');
    const file = path.join(workdir, 'ohne-makro.xlsx');
    await createFixture(file);

    const report = await anonymizeWorkbook({ file, backup: false, seed: 8 });
    assert.equal(report.macrosPreserved, false);
    assert.deepEqual(report.warnings, []);
  });

  test('bleibt ueber mehrere Laeufe stabil', async () => {
    const file = await freshMacroFile('mehrfach.xlsm');
    await anonymizeWorkbook({ file, backup: false, seed: 9 });
    await anonymizeWorkbook({ file, backup: false, seed: 10 });

    const state = await readMacroState(file);
    assert.equal(state.vba, VBA_CONTENT);

    // Der zweite Lauf darf weder Typen noch Beziehungen verdoppeln.
    const overrides = state.contentTypes.match(/PartName="\/xl\/vbaProject\.bin"/g);
    assert.equal(overrides.length, 1);
    const vbaRels = state.workbookRels.match(/relationships\/vbaProject/g);
    assert.equal(vbaRels.length, 1);
  });
});
