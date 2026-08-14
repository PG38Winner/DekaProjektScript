/**
 * Liest eine Excel-Datei, ersetzt die Werte der ausgewaehlten Spalten und
 * schreibt die Datei zurueck. Arbeitet ausschliesslich ueber die Datei selbst -
 * ein installiertes Microsoft Excel wird nicht benoetigt.
 */

import { access, constants, copyFile } from 'node:fs/promises';
import path from 'node:path';
import ExcelJS from 'exceljs';

import { KINDS } from '../core/classify.js';
import { createGenerator } from '../core/generators.js';
import { readMacroParts, restoreMacroParts } from './macros.js';
import { readColumns } from './inspect.js';
import { parseKeep, isKept, validateKeep } from '../core/keep.js';
import { toPlainValue, isReadOnlyValue } from './cells.js';

// `--list` kommt ohne vollstaendiges Einlesen der Datei aus und liegt deshalb
// in einem eigenen Modul.
export { inspectWorkbook } from './inspect.js';

/**
 * Anonymisiert eine Arbeitsmappe.
 *
 * Ohne `sheet` werden **alle** Arbeitsblaetter verarbeitet - dasselbe, was
 * `--list` anzeigt. Mit `sheet` bleibt es bei dem einen genannten Blatt.
 *
 * @param {object} options
 * @param {string} options.file             Eingabedatei.
 * @param {string} [options.sheet]          Blattname; Standard: alle Blaetter.
 * @param {string[]} [options.keep]         Spalten, die unveraendert bleiben.
 * @param {string} [options.out]            Zieldatei; Standard: Eingabedatei.
 * @param {boolean} [options.backup]        Sicherungskopie anlegen (Standard: true).
 * @param {boolean} [options.dryRun]        Nur analysieren, nichts schreiben.
 * @param {number} [options.seed]           Fester Startwert fuer reproduzierbare Laeufe.
 * @param {boolean} [options.consistent]    Gleicher Wert -> gleicher Ersatz (Standard: true).
 * @returns {Promise<object>} Bericht ueber den Lauf.
 */
export async function anonymizeWorkbook({
  file,
  sheet: sheetName,
  keep = [],
  out,
  backup = true,
  dryRun = false,
  seed,
  consistent = true,
}) {
  const workbook = await readWorkbook(file);
  const selected = selectSheets(workbook, sheetName);

  // Erst alle Blaetter einlesen, dann pruefen: --keep darf sich auf eine Spalte
  // beziehen, die nur in einem der Blaetter vorkommt.
  const prepared = selected.map((sheet) => ({ sheet, columns: readColumns(sheet) }));
  const keepRules = parseKeep(keep);
  validateKeep(keepRules, prepared);

  const generate = createGenerator({ seed, consistent });
  const report = {
    sheets: [],
    rows: 0,
    changed: 0,
    output: null,
    backup: null,
    warnings: [],
    macrosPreserved: false,
  };

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
      const columnEntry = {
        header: column.header,
        kind: column.kind,
        status: kept ? 'unveraendert' : column.readOnly ? 'uebersprungen (Formel)' : 'anonymisiert',
        changed: 0,
      };

      if (!kept && !column.readOnly && column.kind !== KINDS.EMPTY) {
        todo.push({ column, entry: columnEntry });
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

  if (dryRun) return report;

  const target = out ? path.resolve(out) : path.resolve(file);

  // Nur sichern, wenn die Originaldatei ueberschrieben wird - schreiben wir in
  // eine neue Datei, bleibt das Original ohnehin unangetastet.
  if (backup && target === path.resolve(file)) {
    report.backup = await createBackup(file);
  }

  // Makros vor dem Schreiben sichern - exceljs verwirft sie beim Erzeugen der
  // neuen Datei und sie muessen anschliessend wieder eingefuegt werden.
  const macro = await readMacroParts(file);

  await workbook.xlsx.writeFile(target);
  report.output = target;

  if (macro) await handleMacros(macro, target, report);

  return report;
}

/**
 * Fuegt ein gesichertes VBA-Projekt wieder ein und meldet, was dabei trotzdem
 * auf der Strecke bleibt.
 */
async function handleMacros(macro, target, report) {
  if (path.extname(target).toLowerCase() !== '.xlsm') {
    report.warnings.push(
      `Die Quelldatei enthaelt Makros, die Zieldatei "${path.basename(target)}" ist aber keine ` +
      '.xlsm-Datei. Makros koennen nur in .xlsm gespeichert werden und gehen hier verloren.',
    );
    return;
  }

  const lost = await restoreMacroParts(target, macro);
  report.macrosPreserved = true;

  if (lost.length) {
    report.warnings.push(
      `Diese Bestandteile der Originaldatei konnten nicht uebernommen werden: ${lost.join(', ')}`,
    );
  }
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
  const targets = todo.map(({ column, entry }) => ({
    index: column.index,
    kind: column.kind,
    key: `${column.kind}:${normalizeHeader(column.header)}`,
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

      writeCell(cell, generate(target.kind, original, target.key));
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


async function readWorkbook(file) {
  // exceljs meldet eine fehlende Datei nur als Text ohne ENOENT-Code, deshalb
  // pruefen wir vorher selbst - so bekommt der Aufrufer eine klare Meldung.
  try {
    await access(file, constants.R_OK);
  } catch {
    throw new Error(`Datei nicht gefunden oder nicht lesbar: ${file}`);
  }

  const workbook = new ExcelJS.Workbook();
  try {
    await workbook.xlsx.readFile(file);
  } catch (error) {
    throw new Error(`Datei konnte nicht gelesen werden: ${error.message}`);
  }
  return workbook;
}

/** Legt "datei.backup-<zeitstempel>.xlsx" neben der Originaldatei an. */
async function createBackup(file) {
  const resolved = path.resolve(file);
  const extension = path.extname(resolved);
  const stamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
  const target = path.join(
    path.dirname(resolved),
    `${path.basename(resolved, extension)}.backup-${stamp}${extension}`,
  );

  await copyFile(resolved, target);
  return target;
}

function normalizeHeader(name) {
  return String(name).trim().toLowerCase();
}
