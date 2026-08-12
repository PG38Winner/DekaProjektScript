/**
 * Ermittelt, was in einer Access-Datenbank ersetzt wird, und erzeugt daraus
 * einen Aenderungsplan.
 *
 * Dieses Modul kennt weder Access noch COM - es arbeitet auf einer schlichten
 * Beschreibung aus Tabellen, Spalten und Zeilen. Dadurch laesst sich die
 * gesamte Entscheidungslogik ohne Datenbank pruefen, obwohl das eigentliche
 * Schreiben nur unter Windows moeglich ist.
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
 * @returns {{report: object, plan: object}}
 */
export function planAnonymization(tables, {
  keep = [],
  table: tableName,
  seed,
  consistent = true,
} = {}) {
  const selected = selectTables(tables, tableName);
  const prepared = selected.map(prepareTable);

  const keepRules = parseKeep(keep);
  validateKeep(keepRules, prepared.map(({ table, columns }) => ({
    sheet: { name: table.name },
    columns: columns.map((column) => ({ header: column.name })),
  })));

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
  const plan = { tables: [] };

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
      const status = columnStatus(column, key, keyNames, keepRules, table.name);
      entry.columns.push({
        header: column.name,
        kind: column.kind,
        status: status.text,
        changed: 0,
      });
      if (status.anonymize) targets.push({ column, entry: entry.columns.at(-1) });
    }

    if (!targets.length) {
      report.sheets.push(entry);
      continue;
    }

    if (!key) {
      entry.skipped = 'keine eindeutige Schluesselspalte gefunden';
      report.warnings.push(
        `Tabelle "${table.name}" hat keine Spalte mit durchgaengig eindeutigen Werten. ` +
        'Ohne sie laesst sich kein Datensatz sicher ansteuern - die Tabelle bleibt unveraendert.',
      );
      report.sheets.push(entry);
      continue;
    }

    const updates = buildUpdates(table, targets, key, generate);
    entry.changed = updates.reduce((sum, update) => sum + Object.keys(update.v).length, 0);
    report.changed += entry.changed;
    report.rows += table.rowCount;
    report.sheets.push(entry);

    if (updates.length) {
      plan.tables.push({
        name: table.name,
        key: key.name,
        keyType: key.type,
        types: Object.fromEntries(targets.map(({ column }) => [column.name, column.type])),
        updates,
      });
    }
  }

  return { report, plan };
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

function columnStatus(column, key, keyNames, keepRules, tableName) {
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
  if (column.kind === KINDS.EMPTY) {
    return { anonymize: false, text: 'uebersprungen (leer)' };
  }
  return { anonymize: true, text: 'anonymisiert' };
}

/**
 * Waehlt die Spalte, ueber die ein Datensatz spaeter angesteuert wird.
 *
 * Bevorzugt wird der Autowert - in Access ist das praktisch immer der
 * Primaerschluessel. Sonst kommt jede Spalte in Frage, deren Werte
 * durchgaengig vorhanden und eindeutig sind; die Eindeutigkeit wird an den
 * gelesenen Daten geprueft und nicht dem Schema geglaubt, denn nur sie
 * entscheidet, ob ein UPDATE genau einen Datensatz trifft.
 *
 * Die Schluesselspalte bleibt immer unveraendert - andernfalls waeren die
 * Datensaetze nach dem ersten Schreibvorgang nicht mehr auffindbar, und
 * Beziehungen zu anderen Tabellen wuerden zerreissen.
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

/** Erzeugt je Datensatz die neuen Werte, angesteuert ueber den Schluessel. */
function buildUpdates(table, targets, key, generate) {
  const updates = [];

  for (const row of table.rows) {
    const values = {};

    for (const { column, entry } of targets) {
      const original = row[column.name];
      // NULL- und Leerwerte bleiben erhalten.
      if (original === null || original === undefined || original === '') continue;

      values[column.name] = generate(column.kind, original, `${column.kind}:${column.name.toLowerCase()}`);
      entry.changed += 1;
    }

    if (Object.keys(values).length) {
      updates.push({ k: row[key.name], v: values });
    }
  }

  return updates;
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
