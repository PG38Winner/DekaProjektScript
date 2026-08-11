/**
 * Liest eine Excel-Datei, ersetzt die Werte der ausgewaehlten Spalten und
 * schreibt die Datei zurueck. Arbeitet ausschliesslich ueber die Datei selbst -
 * ein installiertes Microsoft Excel wird nicht benoetigt.
 */

import { access, constants, copyFile } from 'node:fs/promises';
import path from 'node:path';
import ExcelJS from 'exceljs';

import { classifyColumn, KINDS } from './classify.js';
import { createGenerator } from './generators.js';
import { readMacroParts, restoreMacroParts } from './macros.js';

/** Wie viele Werte pro Spalte fuer die Typerkennung herangezogen werden. */
const SAMPLE_SIZE = 200;

/**
 * Liest die Struktur der Datei aus: Arbeitsblaetter, Spaltennamen, erkannte
 * Inhaltsart. Entspricht dem "Laden"-Schritt der grafischen Oberflaeche.
 *
 * @param {string} file Pfad zur .xlsx/.xlsm-Datei.
 */
export async function inspectWorkbook(file) {
  const workbook = await readWorkbook(file);

  return workbook.worksheets.map((sheet) => {
    const columns = readColumns(sheet);
    return {
      name: sheet.name,
      rowCount: Math.max(sheet.rowCount - 1, 0), // ohne Kopfzeile
      columns: columns.map((column) => ({
        header: column.header,
        index: column.index,
        kind: column.kind,
        readOnly: column.readOnly,
      })),
    };
  });
}

/**
 * Anonymisiert ein Arbeitsblatt.
 *
 * @param {object} options
 * @param {string} options.file             Eingabedatei.
 * @param {string} [options.sheet]          Blattname; Standard: erstes Blatt.
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
  const sheet = selectSheet(workbook, sheetName);
  const columns = readColumns(sheet);

  if (!columns.length) {
    throw new Error(`Arbeitsblatt "${sheet.name}" enthaelt keine Kopfzeile mit Spaltennamen.`);
  }

  const keepSet = new Set(keep.map(normalizeHeader));
  const unknownKeep = keep.filter(
    (name) => !columns.some((column) => normalizeHeader(column.header) === normalizeHeader(name)),
  );
  if (unknownKeep.length) {
    throw new Error(
      `Unbekannte Spalte(n) in --keep: ${unknownKeep.join(', ')}\n` +
      `Vorhanden: ${columns.map((c) => c.header).join(', ')}`,
    );
  }

  const generate = createGenerator({ seed, consistent });
  const report = {
    sheet: sheet.name,
    columns: [],
    rows: 0,
    changed: 0,
    output: null,
    backup: null,
    warnings: [],
    macrosPreserved: false,
  };

  for (const column of columns) {
    const kept = keepSet.has(normalizeHeader(column.header));
    const entry = {
      header: column.header,
      kind: column.kind,
      status: kept ? 'unveraendert' : column.readOnly ? 'uebersprungen (Formel)' : 'anonymisiert',
      changed: 0,
    };

    if (!kept && !column.readOnly && column.kind !== KINDS.EMPTY) {
      entry.changed = anonymizeColumn(sheet, column, generate);
      report.changed += entry.changed;
    } else if (!kept && !column.readOnly && column.kind === KINDS.EMPTY) {
      entry.status = 'uebersprungen (leer)';
    }

    report.columns.push(entry);
  }

  report.rows = Math.max(sheet.rowCount - 1, 0);

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

/** Ersetzt die Werte einer Spalte; gibt die Anzahl geaenderter Zellen zurueck. */
function anonymizeColumn(sheet, column, generate) {
  const columnKey = `${sheet.name}!${column.index}`;
  let changed = 0;

  sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return; // Kopfzeile

    const cell = row.getCell(column.index);
    const original = toPlainValue(cell.value);

    // NULL- und Leerwerte bleiben erhalten, Formeln werden nicht angetastet.
    if (original === null || isReadOnlyValue(cell.value)) return;

    const replacement = generate(column.kind, original, columnKey);
    writeCell(cell, replacement);
    changed += 1;
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

/** Liest die Kopfzeile und leitet fuer jede Spalte die Inhaltsart ab. */
function readColumns(sheet) {
  const headerRow = sheet.getRow(1);
  const columns = [];

  headerRow.eachCell({ includeEmpty: false }, (cell, colNumber) => {
    const header = toPlainValue(cell.value);
    if (header === null || String(header).trim() === '') return;

    const { samples, readOnly } = collectSamples(sheet, colNumber);
    columns.push({
      header: String(header).trim(),
      index: colNumber,
      kind: classifyColumn(String(header), samples),
      readOnly,
    });
  });

  return columns;
}

/**
 * Sammelt Beispielwerte einer Spalte fuer die Typerkennung und stellt fest, ob
 * die Spalte beschreibbar ist. Eine Spalte gilt als schreibgeschuetzt, wenn sie
 * ausschliesslich aus Formeln besteht - genau wie die berechneten Felder, die
 * die Access-Variante ueberspringt.
 */
function collectSamples(sheet, colNumber) {
  const samples = [];
  let nonEmpty = 0;
  let readOnlyCount = 0;

  sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1 || samples.length >= SAMPLE_SIZE) return;

    const raw = row.getCell(colNumber).value;
    const value = toPlainValue(raw);
    if (value === null || value === '') return;

    nonEmpty += 1;
    if (isReadOnlyValue(raw)) readOnlyCount += 1;
    else samples.push(value);
  });

  return { samples, readOnly: nonEmpty > 0 && readOnlyCount === nonEmpty };
}

/** Formeln und Fehlerwerte werden nicht ueberschrieben. */
function isReadOnlyValue(raw) {
  if (!raw || typeof raw !== 'object') return false;
  return 'formula' in raw || 'sharedFormula' in raw || 'error' in raw;
}

/**
 * Reduziert die verschiedenen exceljs-Zellrepraesentationen auf einen einfachen
 * JavaScript-Wert. Gibt `null` fuer leere Zellen zurueck.
 */
function toPlainValue(raw) {
  if (raw === null || raw === undefined) return null;
  if (raw instanceof Date) return raw;
  if (typeof raw !== 'object') return raw;

  if ('richText' in raw) return raw.richText.map((part) => part.text).join('');
  if ('text' in raw) return raw.text;
  if ('formula' in raw || 'sharedFormula' in raw) return raw.result ?? null;
  if ('error' in raw) return null;

  return String(raw);
}

function selectSheet(workbook, sheetName) {
  if (!sheetName) {
    const first = workbook.worksheets[0];
    if (!first) throw new Error('Die Datei enthaelt kein Arbeitsblatt.');
    return first;
  }

  const sheet = workbook.getWorksheet(sheetName);
  if (!sheet) {
    const available = workbook.worksheets.map((s) => s.name).join(', ');
    throw new Error(`Arbeitsblatt "${sheetName}" nicht gefunden. Vorhanden: ${available}`);
  }
  return sheet;
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
