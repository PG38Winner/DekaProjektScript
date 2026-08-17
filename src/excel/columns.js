/**
 * Liest Kopfzeile und Beispielwerte einer bereits geladenen Arbeitsmappe.
 *
 * Ohne Datei- und Betriebssystemzugriff, damit sowohl die Kommandozeile als
 * auch die Browser-Fassung dieselbe Spaltenerkennung verwenden.
 */

import { classifyColumn } from '../core/classify.js';
import { toPlainValue, isReadOnlyValue } from './cells.js';

/** Wie viele Datenzeilen fuer die Typerkennung herangezogen werden. */
export const SAMPLE_SIZE = 200;

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
