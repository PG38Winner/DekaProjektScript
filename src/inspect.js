/**
 * Schnelle Struktur-Analyse einer Arbeitsmappe fuer `--list`.
 *
 * Zum Anzeigen von Blaettern, Spalten und Inhaltsarten muss die Datei nicht
 * vollstaendig eingelesen werden. Statt `Workbook.xlsx.readFile()`, das jede
 * Zelle in den Speicher holt, kommen hier zwei sparsame Wege zum Einsatz:
 *
 *   - der Datenstrom-Leser von exceljs, der nach den ersten Datenzeilen
 *     abgebrochen wird - mehr braucht die Typerkennung nicht;
 *   - der `dimension`-Eintrag im Kopf jedes Blattes fuer die Zeilenzahl, der
 *     sich lesen laesst, ohne das Blatt zu entpacken.
 */

import ExcelJS from 'exceljs';

import { classifyColumn } from './classify.js';
import { toPlainValue, isReadOnlyValue } from './cells.js';
import {
  openWorkbookZip, readText, findInPartHead, listSheets, resolveRelationship, sheetPartPath,
} from './ooxml.js';

/** Wie viele Datenzeilen fuer die Typerkennung herangezogen werden. */
export const SAMPLE_SIZE = 200;

/**
 * Liest die Struktur der Datei aus: Arbeitsblaetter, Spaltennamen, erkannte
 * Inhaltsart. Entspricht dem "Laden"-Schritt der frueheren Oberflaeche.
 *
 * @param {string} file Pfad zur .xlsx/.xlsm-Datei.
 * @returns {Promise<Array<{name: string, rowCount: number|null, columns: Array}>>}
 */
export async function inspectWorkbook(file, { fast = false } = {}) {
  const rowCounts = await readRowCounts(file);

  // Standard ist der vollstaendige, belastbare Weg. Der Datenstrom-Leser ist
  // um ein Vielfaches schneller, hat sich aber an einer echten Arbeitsmappe als
  // unzuverlaessig erwiesen (leere Anzeige), solange die Ursache nicht geklaert
  // ist. Er laesst sich mit --fast anfordern.
  if (!fast) return inspectByFullRead(file, rowCounts);

  try {
    return await inspectByStream(file, rowCounts);
  } catch (error) {
    // Der Datenstrom-Leser von exceljs kommt mit unkomprimiert abgelegten
    // Archiveintraegen nicht zurecht (er scheitert dann beim Aufloesen der
    // Blattliste). Solche Dateien gibt es, deshalb der langsamere, aber
    // belastbare Weg als Rueckfall - ein Absturz waere die schlechteste
    // Antwort auf eine lesbare Datei.
    if (!isStreamLimitation(error)) throw error;
    return inspectByFullRead(file, rowCounts);
  }
}

async function inspectByStream(file, rowCounts) {
  const sheets = [];
  const reader = new ExcelJS.stream.xlsx.WorkbookReader(file, {
    worksheets: 'emit',
    entries: 'emit',
    sharedStrings: 'cache',
    // Die Formatvorlagen werden gebraucht: ob eine Zahl ein Datum ist, steht
    // nicht im Wert, sondern im Zahlenformat der Zelle.
    styles: 'cache',
  });

  for await (const worksheet of reader) {
    sheets.push({
      name: worksheet.name,
      rowCount: rowCounts.get(worksheet.name) ?? null,
      columns: await scanSheet(worksheet),
    });
  }

  // Der schnelle Weg darf nie zu einer leeren Anzeige fuehren. Liefert er kein
  // Blatt, keinen Namen oder nirgends eine Spalte, obwohl die Datei Daten
  // enthaelt, wird vollstaendig gelesen - lieber langsam und richtig.
  if (!sheets.length) throw new StreamLimitation('kein Arbeitsblatt im Datenstrom gefunden');
  if (sheets.some((sheet) => !sheet.name)) {
    throw new StreamLimitation('Blattname im Datenstrom nicht lesbar');
  }

  const foundColumns = sheets.some((sheet) => sheet.columns.length);
  const mayHaveData = sheets.some((sheet) => sheet.rowCount === null || sheet.rowCount > 0);
  if (!foundColumns && mayHaveData) {
    throw new StreamLimitation('keine Spalten im Datenstrom erkannt');
  }

  return sheets;
}

/** Belastbarer Rueckfall: liest die Datei vollstaendig ein. */
async function inspectByFullRead(file, rowCounts) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(file);

  return workbook.worksheets.map((sheet) => ({
    name: sheet.name,
    rowCount: rowCounts.get(sheet.name) ?? Math.max(sheet.rowCount - 1, 0),
    columns: readColumns(sheet),
  }));
}

class StreamLimitation extends Error {}

function isStreamLimitation(error) {
  return error instanceof StreamLimitation
    || /reading 'sheets'|Cannot read properties of undefined/.test(error?.message ?? '');
}

/**
 * Liest Kopfzeile und Beispielwerte eines Blattes aus dem Datenstrom und
 * bricht ab, sobald genug Zeilen gesehen wurden.
 */
async function scanSheet(worksheet) {
  let headers = null;
  const samples = new Map();
  const counts = new Map();
  let dataRows = 0;

  for await (const row of worksheet) {
    if (row.number === 1) {
      headers = row.values;
      continue;
    }
    if (!headers) continue;

    collectRow(row, samples, counts);
    dataRows += 1;
    if (dataRows >= SAMPLE_SIZE) break;
  }

  if (!headers) return [];

  const columns = [];
  headers.forEach((header, index) => {
    if (header === null || header === undefined) return;
    const name = String(toPlainValue(header) ?? '').trim();
    if (!name) return;

    const count = counts.get(index) ?? { nonEmpty: 0, readOnly: 0 };
    columns.push({
      header: name,
      index,
      kind: classifyColumn(name, samples.get(index) ?? []),
      readOnly: count.nonEmpty > 0 && count.readOnly === count.nonEmpty,
    });
  });

  return columns;
}

/** Verteilt die Werte einer Zeile auf die Sammelbehaelter der Spalten. */
function collectRow(row, samples, counts) {
  row.eachCell({ includeEmpty: false }, (cell, index) => {
    const value = toPlainValue(cell.value);
    if (value === null || value === '') return;

    let count = counts.get(index);
    if (!count) {
      count = { nonEmpty: 0, readOnly: 0 };
      counts.set(index, count);
    }
    count.nonEmpty += 1;

    if (isReadOnlyValue(cell.value)) {
      count.readOnly += 1;
      return;
    }

    let list = samples.get(index);
    if (!list) {
      list = [];
      samples.set(index, list);
    }
    list.push(value);
  });
}

/**
 * Liest die Zeilenzahl jedes Blattes aus dessen `dimension`-Eintrag.
 * Fehlt die Angabe, bleibt die Zeilenzahl offen (`null`) - lieber keine Zahl
 * als eine erfundene.
 */
async function readRowCounts(file) {
  const counts = new Map();

  try {
    const zip = await openWorkbookZip(file);
    const workbookXml = await readText(zip, 'xl/workbook.xml');
    const rels = await readText(zip, 'xl/_rels/workbook.xml.rels');

    for (const sheet of listSheets(workbookXml)) {
      const target = resolveRelationship(rels, sheet.rid);
      if (!target) continue;

      const match = await findInPartHead(
        zip, sheetPartPath(target), /<dimension\b[^>]*\bref="([^"]*)"/,
      );
      const lastRow = match ? lastRowOfRange(match[1]) : null;
      if (lastRow !== null) counts.set(sheet.name, Math.max(lastRow - 1, 0));
    }
  } catch {
    // Ohne die Angabe laeuft die Analyse weiter, nur ohne Zeilenzahl.
  }

  return counts;
}

/** "A1:L50001" -> 50001 */
function lastRowOfRange(ref) {
  const end = String(ref).split(':').pop();
  const match = /(\d+)$/.exec(end ?? '');
  return match ? Number(match[1]) : null;
}

/**
 * Liest die Kopfzeile und leitet fuer jede Spalte die Inhaltsart ab.
 *
 * Die Beispielwerte aller Spalten werden in **einem** Durchlauf ueber die
 * ersten Datenzeilen gesammelt. Frueher lief je Spalte ein eigener Durchlauf
 * ueber das gesamte Blatt - bei vielen Spalten vervielfachte das die Arbeit.
 */
export function readColumns(sheet) {
  const headerRow = sheet.getRow(1);
  const headers = new Map();

  headerRow.eachCell({ includeEmpty: false }, (cell, colNumber) => {
    const header = toPlainValue(cell.value);
    if (header === null || String(header).trim() === '') return;
    headers.set(colNumber, String(header).trim());
  });

  if (!headers.size) return [];

  const { samples, counts } = collectSamples(sheet, headers);

  return [...headers].map(([index, header]) => {
    const count = counts.get(index) ?? { nonEmpty: 0, readOnly: 0 };
    return {
      header,
      index,
      kind: classifyColumn(header, samples.get(index) ?? []),
      // Eine Spalte gilt als schreibgeschuetzt, wenn sie ausschliesslich aus
      // Formeln besteht - wie die berechneten Felder der Access-Variante.
      readOnly: count.nonEmpty > 0 && count.readOnly === count.nonEmpty,
    };
  });
}

/** Sammelt Beispielwerte aller Spalten in einem Durchlauf ueber die Zeilen. */
function collectSamples(sheet, headers) {
  const samples = new Map();
  const counts = new Map();
  let dataRows = 0;

  const lastRow = Math.min(sheet.rowCount, SAMPLE_SIZE + 1);
  for (let rowNumber = 2; rowNumber <= lastRow && dataRows < SAMPLE_SIZE; rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    if (!row) continue;
    dataRows += 1;

    for (const colNumber of headers.keys()) {
      const raw = row.getCell(colNumber).value;
      const value = toPlainValue(raw);
      if (value === null || value === '') continue;

      let count = counts.get(colNumber);
      if (!count) {
        count = { nonEmpty: 0, readOnly: 0 };
        counts.set(colNumber, count);
      }
      count.nonEmpty += 1;

      if (isReadOnlyValue(raw)) {
        count.readOnly += 1;
        continue;
      }

      let list = samples.get(colNumber);
      if (!list) {
        list = [];
        samples.set(colNumber, list);
      }
      list.push(value);
    }
  }

  return { samples, counts };
}
