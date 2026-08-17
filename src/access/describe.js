/**
 * Wertet eine bereits gelesene Access-Datenbank aus.
 *
 * Ohne Dateizugriff, damit die Browser-Fassung dieselbe Auswertung nutzt.
 * Das Ergebnis ist eine schlichte Beschreibung aus Tabellen, Spalten und
 * Zeilen - genau die Form, die plan.js erwartet.
 */

import MDBReader from 'mdb-reader';

/**
 * Spaltentypen, deren Inhalt nicht sinnvoll ersetzt werden kann: eingebettete
 * Dateien, Bilder, Anlagen und Mehrfachwertfelder. Sie bleiben unangetastet.
 */
const UNSUPPORTED_TYPES = new Set(['OLE', 'Binary', 'Complex', 'RepID']);

/**
 * Wie readDatabase, aber auf einem bereits gelesenen Puffer - so kann auch die
 * Browser-Fassung dieselbe Auswertung nutzen.
 *
 * @param {Buffer|Uint8Array} buffer
 */
export function describeDatabase(buffer, { password, rowLimit } = {}) {
  const reader = openDatabase(buffer, password);

  return reader.getTableNames({ normalTables: true, systemTables: false, linkedTables: false })
    .map((name) => {
      const table = reader.getTable(name);
      const columns = table.getColumns().map(describeColumn);

      return {
        name,
        rowCount: table.rowCount,
        columns,
        rows: readRows(table, columns, rowLimit),
      };
    });
}

function openDatabase(buffer, password) {
  try {
    return new MDBReader(buffer, password ? { password } : undefined);
  } catch (error) {
    throw new Error(
      `Datenbank konnte nicht gelesen werden: ${error.message}\n` +
      'Ist die Datei kennwortgeschuetzt, das Kennwort mit --password angeben.',
    );
  }
}

/**
 * Uebersetzt eine Spaltenbeschreibung von mdb-reader in unsere Form und
 * entscheidet, ob die Spalte beschreibbar ist.
 *
 * Autowert-Spalten (`autoLong`) und automatische GUIDs (`autoUUID`) vergibt
 * Access selbst - ihre Werte sind Verwaltungsdaten und bleiben stehen.
 */
function describeColumn(column) {
  const unsupported = UNSUPPORTED_TYPES.has(column.type);
  const automatic = Boolean(column.autoLong || column.autoUUID);

  return {
    name: column.name,
    type: column.type,
    size: column.size,
    nullable: column.nullable,
    autoNumber: Boolean(column.autoLong),
    readOnly: unsupported || automatic,
    readOnlyReason: automatic ? 'Autowert' : unsupported ? `Typ ${column.type}` : null,
  };
}

/** Liest die Zeilen ohne die Spalten, die ohnehin nicht angefasst werden. */
function readRows(table, columns, rowLimit) {
  const readable = columns
    .filter((column) => !UNSUPPORTED_TYPES.has(column.type))
    .map((column) => column.name);

  if (!readable.length) return [];

  return table.getData({
    columns: readable,
    ...(rowLimit === undefined ? {} : { rowLimit }),
  });
}
