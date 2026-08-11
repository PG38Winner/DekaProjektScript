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
  validateKeep(keep, prepared);

  const keepSet = new Set(keep.map(normalizeHeader));
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

    for (const column of columns) {
      const kept = keepSet.has(normalizeHeader(column.header));
      const columnEntry = {
        header: column.header,
        kind: column.kind,
        status: kept ? 'unveraendert' : column.readOnly ? 'uebersprungen (Formel)' : 'anonymisiert',
        changed: 0,
      };

      if (!kept && !column.readOnly && column.kind !== KINDS.EMPTY) {
        columnEntry.changed = anonymizeColumn(sheet, column, generate);
        entry.changed += columnEntry.changed;
      } else if (!kept && !column.readOnly && column.kind === KINDS.EMPTY) {
        columnEntry.status = 'uebersprungen (leer)';
      }

      entry.columns.push(columnEntry);
    }

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

/** Ersetzt die Werte einer Spalte; gibt die Anzahl geaenderter Zellen zurueck. */
function anonymizeColumn(sheet, column, generate) {
  // Der Schluessel ist bewusst blattuebergreifend: eine Spalte gleichen Namens
  // und gleicher Art verknuepft in der Regel zwei Blaetter (KundenID in
  // "Kunden" und in "Bestellungen"). Nur mit gemeinsamem Schluessel bleibt die
  // Verknuepfung nach der Anonymisierung bestehen.
  const columnKey = `${column.kind}:${normalizeHeader(column.header)}`;
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

/**
 * Prueft die --keep-Namen gegen alle zu verarbeitenden Blaetter. Ein Name muss
 * in mindestens einem Blatt vorkommen; wo er vorkommt, wirkt er.
 */
function validateKeep(keep, prepared) {
  if (!keep.length) return;

  const known = new Set(
    prepared.flatMap(({ columns }) => columns.map((column) => normalizeHeader(column.header))),
  );
  const unknown = keep.filter((name) => !known.has(normalizeHeader(name)));
  if (!unknown.length) return;

  const available = prepared
    .filter(({ columns }) => columns.length)
    .map(({ sheet, columns }) => `  ${sheet.name}: ${columns.map((c) => c.header).join(', ')}`)
    .join('\n');

  throw new Error(
    `Unbekannte Spalte(n) in --keep: ${unknown.join(', ')}\nVorhanden:\n${available}`,
  );
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
