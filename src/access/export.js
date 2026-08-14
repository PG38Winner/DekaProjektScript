/**
 * Schreibt die anonymisierten Tabellen einer Access-Datenbank als
 * Excel-Arbeitsmappe - je Tabelle ein Arbeitsblatt.
 *
 * Warum nicht zurueck in die .accdb: In eine Access-Datenbank schreiben kann
 * nur die Microsoft Access Database Engine (ACE-OLEDB). Aus reinem JavaScript
 * ist das nicht moeglich - saemtliche Pakete lesen entweder nur oder rufen
 * ACE ueber einen Fremdprozess auf. Der Umweg ueber Excel bleibt damit der
 * einzige Weg, der ohne Fremdprozesse und auf jedem System funktioniert.
 *
 * Nebenwirkung, und zwar eine gute: Die Quelldatenbank wird nie angefasst.
 */

import ExcelJS from 'exceljs';

/** Zeichen, die Excel in Blattnamen nicht zulaesst. */
const FORBIDDEN_IN_SHEET_NAME = /[\\/?*[\]:]/g;
const MAX_SHEET_NAME = 31;

/** Access-Typen, die als Datum geschrieben werden. */
const DATE_TYPES = new Set(['DateTime', 'DateTimeExtended']);

/**
 * @param {string} file    Zieldatei (.xlsx).
 * @param {object} result  Rueckgabe von planAnonymization().
 * @returns {Promise<string[]>} Hinweise, etwa zu umbenannten Blaettern.
 */
export async function writeWorkbook(file, result) {
  const workbook = new ExcelJS.Workbook();
  const notes = [];
  const used = new Set();

  for (const table of result.tables) {
    const name = sheetNameFor(table.name, used);
    if (name !== table.name) {
      notes.push(`Tabelle "${table.name}" heisst als Arbeitsblatt "${name}" - `
        + 'Excel erlaubt keine laengeren oder anders geschriebenen Blattnamen.');
    }

    const sheet = workbook.addWorksheet(name);
    const columns = table.columns.map((column) => column.name);
    sheet.addRow(columns);
    sheet.getRow(1).font = { bold: true };

    const dateColumns = table.columns
      .map((column, index) => (DATE_TYPES.has(column.type) ? index + 1 : null))
      .filter((index) => index !== null);

    for (const row of table.rows) {
      const added = sheet.addRow(columns.map((column) => row[column] ?? null));
      for (const index of dateColumns) added.getCell(index).numFmt = 'dd.mm.yyyy';
    }
  }

  if (!workbook.worksheets.length) workbook.addWorksheet('Leer');

  await workbook.xlsx.writeFile(file);
  return notes;
}

/**
 * Excel begrenzt Blattnamen auf 31 Zeichen und verbietet einige Zeichen.
 * Namen, die dadurch gleich wuerden, werden durchnummeriert.
 */
export function sheetNameFor(tableName, used = new Set()) {
  let name = String(tableName).replace(FORBIDDEN_IN_SHEET_NAME, '_').slice(0, MAX_SHEET_NAME).trim();
  if (!name) name = 'Tabelle';

  if (!used.has(name.toLowerCase())) {
    used.add(name.toLowerCase());
    return name;
  }

  for (let suffix = 2; ; suffix += 1) {
    const candidate = `${name.slice(0, MAX_SHEET_NAME - String(suffix).length - 1)}_${suffix}`;
    if (used.has(candidate.toLowerCase())) continue;
    used.add(candidate.toLowerCase());
    return candidate;
  }
}
