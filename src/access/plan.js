/**
 * Ermittelt, was in einer Access-Datenbank ersetzt wird, und erzeugt daraus
 * die anonymisierten Tabellen.
 *
 * Dieses Modul kennt Access nicht - es arbeitet auf einer schlichten
 * Beschreibung aus Tabellen, Spalten und Zeilen. Dadurch laesst sich die
 * gesamte Entscheidungslogik ohne Datenbank pruefen.
 */

import { classifyColumn, KINDS } from '../core/classify.js';
import { createGenerator } from '../core/generators.js';
import { parseKeep, isKept, validateKeep } from '../core/keep.js';

/** Wie viele Werte je Spalte in die Typerkennung eingehen. */
const SAMPLE_SIZE = 200;

/**
 * Spaltennamen, die bevorzugt als Schluessel dienen, wenn es keinen Autowert
 * gibt. Nur eine Vorliebe: gewaehlt wird ohnehin nur unter Spalten, deren
 * Werte nachweislich eindeutig sind.
 */
const KEY_NAME_HINT = /(id|nr|nummer|schluessel|key)$/i;

/**
 * @param {Array} tables    Rueckgabe von readDatabase().
 * @param {object} options
 * @param {string[]} [options.keep]      --keep-Angaben.
 * @param {string} [options.table]       Nur diese Tabelle.
 * @param {number} [options.seed]
 * @param {boolean} [options.consistent]
 * @returns {{report: object, result: {tables: Array}}}
 */
export function planAnonymization(tables, {
  keep = [],
  clear = [],
  table: tableName,
  seed,
  consistent = true,
} = {}) {
  const selected = selectTables(tables, tableName);
  const prepared = selected.map(prepareTable);

  const shape = prepared.map(({ table, columns }) => ({
    sheet: { name: table.name },
    columns: columns.map((column) => ({ header: column.name })),
  }));
  const keepRules = parseKeep(keep);
  const clearRules = parseKeep(clear);
  validateKeep(keepRules, shape);
  validateKeep(clearRules, shape);

  // Erst die Schluessel aller Tabellen bestimmen. Eine Spalte, die anderswo
  // Schluessel ist, verweist in der Regel genau dorthin (Fremdschluessel).
  // Wuerde sie ersetzt, zeigte sie ins Leere - der Schluessel selbst bleibt ja
  // unveraendert. Solche Spalten bleiben deshalb ebenfalls stehen.
  const keys = new Map(prepared.map(({ table, columns }) => [
    table.name, chooseKeyColumn(table, columns),
  ]));
  const keyNames = new Set(
    [...keys.values()].filter(Boolean).map((column) => column.name.toLowerCase()),
  );

  const generate = createGenerator({ seed, consistent });
  const report = {
    label: 'Tabelle',
    sheets: [],
    rows: 0,
    changed: 0,
    output: null,
    backup: null,
    warnings: [],
  };
  const result = { tables: [] };

  for (const { table, columns } of prepared) {
    const entry = {
      name: table.name,
      rows: table.rowCount,
      columns: [],
      changed: 0,
      skipped: null,
    };

    if (!columns.length) {
      entry.skipped = 'keine Spalten';
      report.sheets.push(entry);
      continue;
    }

    const key = keys.get(table.name);
    const targets = [];

    for (const column of columns) {
      const status = columnStatus(column, key, keyNames, keepRules, clearRules, table.name);
      entry.columns.push({
        header: column.name,
        kind: column.kind,
        status: status.text,
        changed: 0,
      });
      if (status.anonymize) targets.push({ column, entry: entry.columns.at(-1), clear: status.clear });
    }

    const { rows, changed } = buildRows(table, columns, targets, generate);
    entry.changed = changed;
    report.changed += changed;
    report.rows += table.rowCount;
    report.sheets.push(entry);

    result.tables.push({
      name: table.name,
      columns: columns.map((column) => ({ name: column.name, type: column.type })),
      rows,
    });
  }

  return { report, result };
}

/** Leitet fuer jede Spalte die Inhaltsart aus Name und Werten ab. */
function prepareTable(table) {
  const columns = table.columns.map((column) => {
    const samples = [];
    for (const row of table.rows) {
      if (samples.length >= SAMPLE_SIZE) break;
      const value = row[column.name];
      if (value === null || value === undefined || value === '') continue;
      samples.push(value);
    }

    return { ...column, kind: classifyColumn(column.name, samples) };
  });

  return { table, columns };
}

function columnStatus(column, key, keyNames, keepRules, clearRules, tableName) {
  if (key && column.name === key.name) {
    return { anonymize: false, text: 'unveraendert (Schluessel)' };
  }
  if (keyNames.has(column.name.toLowerCase())) {
    return { anonymize: false, text: 'unveraendert (Verknuepfung)' };
  }
  if (isKept(keepRules, tableName, column.name)) {
    return { anonymize: false, text: 'unveraendert' };
  }
  if (column.readOnly) {
    return { anonymize: false, text: `uebersprungen (${column.readOnlyReason})` };
  }
  if (isKept(clearRules, tableName, column.name)) {
    return { anonymize: true, clear: true, text: 'geleert' };
  }
  if (column.kind === KINDS.EMPTY) {
    return { anonymize: false, text: 'uebersprungen (leer)' };
  }
  return { anonymize: true, text: 'anonymisiert' };
}

/**
 * Findet die Spalte, die die Datensaetze der Tabelle identifiziert.
 *
 * Bevorzugt wird der Autowert - in Access ist das praktisch immer der
 * Primaerschluessel. Sonst kommt jede Spalte in Frage, deren Werte
 * durchgaengig vorhanden und eindeutig sind; geprueft wird das an den
 * gelesenen Daten und nicht am Schema.
 *
 * Die Schluesselspalte bleibt unveraendert, damit die Beziehungen zwischen
 * den Tabellen erhalten bleiben. Findet sich keine, ist auch nichts zu
 * schuetzen - dann werden alle Spalten anonymisiert.
 */
export function chooseKeyColumn(table, columns) {
  const autoNumber = columns.find((column) => column.autoNumber);
  if (autoNumber) return autoNumber;

  const unique = columns.filter((column) => !column.readOnly && isUnique(table.rows, column.name));
  if (!unique.length) return null;

  return unique.find((column) => KEY_NAME_HINT.test(column.name)) ?? unique[0];
}

function isUnique(rows, name) {
  if (!rows.length) return false;

  const seen = new Set();
  for (const row of rows) {
    const value = row[name];
    if (value === null || value === undefined || value === '') return false;

    const key = value instanceof Date ? `d:${value.getTime()}` : `v:${String(value)}`;
    if (seen.has(key)) return false;
    seen.add(key);
  }
  return true;
}

/**
 * Baut die vollstaendigen Datensaetze der anonymisierten Tabelle: unberuehrte
 * Spalten unveraendert, die uebrigen ersetzt.
 */
function buildRows(table, columns, targets, generate) {
  const replace = new Map(targets.map(({ column, entry, clear }) => [column.name, { column, entry, clear }]));
  const rows = [];
  let changed = 0;

  for (const source of table.rows) {
    const row = {};

    for (const column of columns) {
      const original = source[column.name];
      const target = replace.get(column.name);

      // Unberuehrte Spalten sowie NULL- und Leerwerte werden uebernommen.
      if (!target || original === null || original === undefined || original === '') {
        row[column.name] = original ?? null;
        continue;
      }

      // Geleerte Spalten bekommen keinen Ersatzwert - der sicherste Umgang
      // mit Freitext, in dem alles Moegliche stehen kann.
      row[column.name] = target.clear ? null : generate(
        target.column.kind, original, `${target.column.kind}:${column.name.toLowerCase()}`,
      );
      target.entry.changed += 1;
      changed += 1;
    }

    rows.push(row);
  }

  return { rows, changed };
}

function selectTables(tables, tableName) {
  if (!tables.length) throw new Error('Die Datenbank enthaelt keine Tabelle.');
  if (!tableName) return tables;

  const found = tables.find((table) => table.name.toLowerCase() === tableName.toLowerCase());
  if (!found) {
    throw new Error(
      `Tabelle "${tableName}" nicht gefunden. Vorhanden: ${tables.map((t) => t.name).join(', ')}`,
    );
  }
  return [found];
}
