/**
 * Liest Access-Datenbanken (.accdb / .mdb).
 *
 * Gelesen wird mit `mdb-reader` - reines JavaScript, ohne Zusatzsoftware und
 * ohne Windows. Das Ergebnis wird als Excel-Arbeitsmappe geschrieben, die
 * Datenbank selbst bleibt unveraendert (siehe export.js).
 *
 * Das Ergebnis hat dieselbe Form, die plan.js erwartet - eine schlichte
 * Beschreibung aus Tabellen, Spalten und Zeilen. So laesst sich die gesamte
 * Anonymisierungslogik ohne Datenbank pruefen.
 */

import { readFile } from 'node:fs/promises';

import { describeDatabase } from './describe.js';

export { describeDatabase };


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

  return describeDatabase(buffer, { password, rowLimit });
}
