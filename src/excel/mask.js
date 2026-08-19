/**
 * Maskiert eine bereits geladene Excel-Arbeitsmappe im Arbeitsspeicher.
 *
 * Bewusst ohne Datei- oder Betriebssystemzugriff: dieselbe Logik laeuft damit
 * in Node.js (Kommandozeile) und im Browser (Einzeldatei-Anwendung). Wer die
 * Datei liest und schreibt, entscheidet der jeweilige Aufrufer.
 */

import { KINDS } from '../core/classify.js';
import { createGenerator } from '../core/generators.js';
import { parseKeep, isKept, validateKeep } from '../core/keep.js';
import { readColumns } from './columns.js';
import { toPlainValue, isReadOnlyValue } from './cells.js';

/**
 * @param {import('exceljs').Workbook} workbook  Geladene Arbeitsmappe; wird veraendert.
 * @param {object} options
 * @param {string} [options.sheet]        Nur dieses Blatt; Standard: alle.
 * @param {string[]} [options.keep]       --keep-Angaben.
 * @param {number} [options.seed]
 * @param {boolean} [options.consistent]
 * @returns {{report: object, warnings: string[]}}
 */
export function maskWorkbook(workbook, { sheet: sheetName, keep = [], clear = [], seed, consistent = true } = {}) {
  const selected = selectSheets(workbook, sheetName);

  // Erst alle Blaetter einlesen, dann pruefen: --keep darf sich auf eine Spalte
  // beziehen, die nur in einem der Blaetter vorkommt.
  const prepared = selected.map((sheet) => ({ sheet, columns: readColumns(sheet) }));
  const keepRules = parseKeep(keep);
  const clearRules = parseKeep(clear);
  validateKeep(keepRules, prepared);
  validateKeep(clearRules, prepared);

  const generate = createGenerator({ seed, consistent });
  const report = { sheets: [], rows: 0, changed: 0 };

  for (const { sheet, columns } of prepared) {
    const entry = {
      name: sheet.name,
      rows: Math.max(sheet.rowCount - 1, 0),
      columns: [],
      changed: 0,
      skipped: null,
    };

    // Blaetter ohne Kopfzeile (Deckblatt, Notizen) werden uebergangen, statt
    // den ganzen Lauf abzubrechen.
    if (!columns.length) {
      entry.skipped = 'keine Kopfzeile mit Spaltennamen';
      report.sheets.push(entry);
      continue;
    }

    const todo = [];
    for (const column of columns) {
      const kept = isKept(keepRules, sheet.name, column.header);
      const cleared = !kept && isKept(clearRules, sheet.name, column.header);
      const columnEntry = {
        header: column.header,
        kind: column.kind,
        status: kept ? 'unveraendert'
          : column.readOnly ? 'uebersprungen (Formel)'
            : cleared ? 'geleert' : 'anonymisiert',
        changed: 0,
      };

      if (!kept && !column.readOnly && (cleared || column.kind !== KINDS.EMPTY)) {
        todo.push({ column, entry: columnEntry, clear: cleared });
      } else if (!kept && !column.readOnly && column.kind === KINDS.EMPTY) {
        columnEntry.status = 'uebersprungen (leer)';
      }

      entry.columns.push(columnEntry);
    }

    entry.changed = anonymizeSheet(sheet, todo, generate);
    report.changed += entry.changed;
    report.rows += entry.rows;
    report.sheets.push(entry);
  }

  return { report, warnings: [] };
}

function normalizeHeader(name) {
  return String(name).trim().toLowerCase();
}

/**
 * Ersetzt die Werte aller vorgemerkten Spalten eines Blattes in einem
 * Durchlauf ueber die Zeilen und gibt die Anzahl geaenderter Zellen zurueck.
 *
 * Ein Durchlauf je Spalte waere bei breiten Blaettern ein Vielfaches an
 * Arbeit - die Zeilen werden deshalb genau einmal besucht.
 */
function anonymizeSheet(sheet, todo, generate) {
  if (!todo.length) return 0;

  // Der Schluessel ist bewusst blattuebergreifend: eine Spalte gleichen Namens
  // und gleicher Art verknuepft in der Regel zwei Blaetter (KundenID in
  // "Kunden" und in "Bestellungen"). Nur mit gemeinsamem Schluessel bleibt die
  // Verknuepfung nach der Anonymisierung bestehen.
  const targets = todo.map(({ column, entry, clear }) => ({
    index: column.index,
    kind: column.kind,
    key: `${column.kind}:${normalizeHeader(column.header)}`,
    clear,
    entry,
  }));

  let changed = 0;

  sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return; // Kopfzeile

    for (const target of targets) {
      const cell = row.getCell(target.index);
      const original = toPlainValue(cell.value);

      // NULL- und Leerwerte bleiben erhalten, Formeln werden nicht angetastet.
      if (original === null || isReadOnlyValue(cell.value)) continue;

      // Geleerte Spalten bekommen keinen Ersatzwert - der sicherste Umgang
      // mit Freitext, in dem alles Moegliche stehen kann.
      if (target.clear) cell.value = null;
      else writeCell(cell, generate(target.kind, original, target.key));
      target.entry.changed += 1;
      changed += 1;
    }
  });

  return changed;
}

/**
 * Schreibt den Ersatzwert und erhaelt dabei die Zellformatierung. exceljs haelt
 * `style` getrennt vom Wert, das Zahlen-/Datumsformat bleibt also bestehen.
 */
function writeCell(cell, replacement) {
  const previous = cell.value;

  // Verlinkte Zellen (z. B. mailto:) behalten ihre Struktur, damit der Link
  // nicht auf die echte Adresse zeigt.
  if (previous && typeof previous === 'object' && 'hyperlink' in previous) {
    const text = String(replacement);
    cell.value = {
      text,
      hyperlink: /^[^\s@]+@[^\s@]+$/.test(text) ? `mailto:${text}` : text,
    };
    return;
  }

  cell.value = replacement;
}

/** Ohne Namen alle Arbeitsblaetter, sonst genau das genannte. */
function selectSheets(workbook, sheetName) {
  if (!workbook.worksheets.length) throw new Error('Die Datei enthaelt kein Arbeitsblatt.');

  if (!sheetName) return workbook.worksheets;

  const sheet = workbook.getWorksheet(sheetName);
  if (!sheet) {
    const available = workbook.worksheets.map((s) => s.name).join(', ');
    throw new Error(`Arbeitsblatt "${sheetName}" nicht gefunden. Vorhanden: ${available}`);
  }
  return [sheet];
}
