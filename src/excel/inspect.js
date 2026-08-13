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

import { classifyColumn } from '../core/classify.js';
import { scanWorkbook } from './scan.js';
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
export async function inspectWorkbook(file, { full = false } = {}) {
  const rowCounts = await readRowCounts(file);

  // Standard ist der eigene, sparsame Leser: er kommt mit beliebig grossen
  // Dateien zurecht, weil er nur die ersten Zeilen liest. Der vollstaendige
  // Weg laedt jede Zelle in den Speicher und scheitert an grossen Dateien -
  // er bleibt als Rueckfall und laesst sich mit --full erzwingen.
  if (full) return inspectByFullRead(file, rowCounts);

  try {
    return await inspectByScan(file, rowCounts);
  } catch {
    // Jeder Fehler des sparsamen Wegs fuehrt zum vollstaendigen Einlesen: er
    // ist eine Beschleunigung, keine Bedingung. Eine leere Anzeige oder ein
    // Abbruch waeren die schlechteste Antwort auf eine lesbare Datei - lieber
    // langsam und richtig. Scheitert auch der vollstaendige Weg, meldet
    // dessen Fehler das eigentliche Problem.
    return inspectByFullRead(file, rowCounts);
  }
}

/**
 * Sparsamer Weg: liest nur Kopfzeile und die ersten Zeilen je Blatt.
 * Liefert er nichts Brauchbares, wird der Fehler gemeldet und der Aufrufer
 * faellt auf das vollstaendige Einlesen zurueck - eine leere Anzeige waere
 * die schlechteste Antwort auf eine lesbare Datei.
 */
async function inspectByScan(file, rowCounts) {
  const sheets = await scanWorkbook(file, SAMPLE_SIZE);

  if (!sheets.length) throw new ScanLimitation('kein Arbeitsblatt gefunden');
  if (sheets.some((sheet) => !sheet.name)) throw new ScanLimitation('Blattname nicht lesbar');

  const mapped = sheets.map((sheet) => ({
    name: sheet.name,
    rowCount: rowCounts.get(sheet.name) ?? sheet.rowCount,
    columns: sheet.columns.map((column) => ({
      header: column.header,
      index: column.index,
      kind: classifyColumn(column.header, column.values),
      readOnly: column.readOnly,
      note: column.readOnly ? 'Formel' : null,
    })),
  }));

  const foundColumns = mapped.some((sheet) => sheet.columns.length);
  const mayHaveData = mapped.some((sheet) => sheet.rowCount === null || sheet.rowCount > 0);
  if (!foundColumns && mayHaveData) throw new ScanLimitation('keine Spalten erkannt');

  return mapped;
}


/** Belastbarer Rueckfall: liest die Datei vollstaendig ein. */
async function inspectByFullRead(file, rowCounts) {
  const workbook = new ExcelJS.Workbook();
  try {
    await workbook.xlsx.readFile(file);
  } catch (error) {
    throw new Error(`Datei konnte nicht gelesen werden: ${error.message}`);
  }

  return workbook.worksheets.map((sheet) => ({
    name: sheet.name,
    rowCount: rowCounts.get(sheet.name) ?? Math.max(sheet.rowCount - 1, 0),
    columns: readColumns(sheet),
  }));
}

/** Signalisiert, dass der sparsame Weg nichts Brauchbares geliefert hat. */
class ScanLimitation extends Error {}



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
      // Formeln besteht - wie die berechneten Felder in Access.
      readOnly: count.nonEmpty > 0 && count.readOnly === count.nonEmpty,
      note: count.nonEmpty > 0 && count.readOnly === count.nonEmpty ? 'Formel' : null,
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
