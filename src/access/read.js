/**
 * Liest Access-Datenbanken (.accdb / .mdb).
 *
 * Gelesen wird mit `mdb-reader` - reines JavaScript, ohne ACE-OLEDB und ohne
 * Windows. Dadurch funktionieren `--list` und `--dry-run` auf jedem System.
 * Zum **Schreiben** ist weiterhin Windows mit ACE-OLEDB noetig, siehe write.js.
 *
 * Das Ergebnis hat dieselbe Form, die plan.js erwartet - eine schlichte
 * Beschreibung aus Tabellen, Spalten und Zeilen. So laesst sich die gesamte
 * Anonymisierungslogik ohne Datenbank pruefen.
 */

import { readFile } from 'node:fs/promises';
import MDBReader from 'mdb-reader';

/**
 * Spaltentypen, deren Inhalt nicht sinnvoll ersetzt werden kann: eingebettete
 * Dateien, Bilder, Anlagen und Mehrfachwertfelder. Sie bleiben unangetastet.
 */
const UNSUPPORTED_TYPES = new Set(['OLE', 'Binary', 'Complex', 'RepID']);

/**
 * @param {string} file Pfad zur .accdb/.mdb-Datei.
 * @param {object} [options]
 * @param {string} [options.password] Datenbankkennwort, falls gesetzt.
 * @param {number} [options.rowLimit] Nur die ersten n Zeilen lesen (fuer --list).
 * @returns {Promise<Array<{name: string, rowCount: number, columns: Array, rows: Array}>>}
 */
export async function readDatabase(file, { password, rowLimit } = {}) {
  let buffer;
  try {
    buffer = await readFile(file);
  } catch {
    throw new Error(`Datei nicht gefunden oder nicht lesbar: ${file}`);
  }

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
 * Access selbst - sie lassen sich nicht aktualisieren. Das entspricht der
 * Erkennung ueber COLUMN_FLAGS, die das fruehere PowerShell-Skript nutzte.
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
