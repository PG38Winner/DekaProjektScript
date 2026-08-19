/**
 * Liest eine Excel-Datei, ersetzt die Werte der ausgewaehlten Spalten und
 * schreibt die Datei zurueck. Arbeitet ausschliesslich ueber die Datei selbst -
 * ein installiertes Microsoft Excel wird nicht benoetigt.
 */

import { access, constants, copyFile } from 'node:fs/promises';
import path from 'node:path';
import ExcelJS from 'exceljs';

import { readMacroParts, restoreMacroParts } from './macros.js';
import { maskWorkbook } from './mask.js';

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
  clear = [],
  out,
  backup = true,
  dryRun = false,
  seed,
  consistent = true,
}) {
  const workbook = await readWorkbook(file);
  const { report, warnings } = maskWorkbook(workbook, {
    sheet: sheetName, keep, clear, seed, consistent,
  });
  report.output = null;
  report.backup = null;
  report.macrosPreserved = false;
  report.warnings = warnings;

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
