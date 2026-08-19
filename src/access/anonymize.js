/**
 * Anonymisiert Access-Datenbanken (.accdb / .mdb).
 *
 * Ablauf, vollstaendig in Node.js und ohne Fremdprozesse:
 *   1. Lesen mit mdb-reader (reines JavaScript)
 *   2. Planen in plan.js
 *   3. Schreiben als Excel-Arbeitsmappe, je Tabelle ein Blatt
 *
 * Die Quelldatenbank wird dabei **nie veraendert**. In eine .accdb
 * zurueckzuschreiben kann nur die Microsoft Access Database Engine, und die
 * laesst sich aus reinem JavaScript nicht ansprechen - siehe export.js.
 */

import path from 'node:path';

import { readDatabase } from './read.js';
import { planAnonymization } from './plan.js';
import { writeWorkbook } from './export.js';

/**
 * Liest die Struktur der Datenbank: Tabellen, Spalten, erkannte Inhaltsart.
 *
 * @param {string} file
 * @param {object} [options]
 * @param {string} [options.password]
 */
export async function inspectDatabase(file, { password } = {}) {
  // Fuer die Typerkennung genuegen wenige Zeilen - das haelt --list auch bei
  // grossen Datenbanken schnell.
  const tables = await readDatabase(file, { password, rowLimit: 200 });
  const { report } = planAnonymization(tables, {});

  return tables.map((table) => {
    const planned = report.sheets.find((entry) => entry.name === table.name);
    return {
      name: table.name,
      rowCount: table.rowCount,
      columns: (planned?.columns ?? []).map((column) => ({
        header: column.header,
        kind: column.kind,
        readOnly: column.status.startsWith('uebersprungen'),
        note: noteFor(column.status),
      })),
    };
  });
}

function noteFor(status) {
  if (status.includes('(Schluessel)')) return 'Schluessel';
  if (status.includes('(Verknuepfung)')) return 'Verknuepfung';
  const skipped = /^uebersprungen \((.+)\)$/.exec(status);
  return skipped ? skipped[1] : null;
}

/**
 * @param {object} options
 * @param {string} options.file          Eingabedatenbank.
 * @param {string} [options.table]       Nur diese Tabelle; Standard: alle.
 * @param {string[]} [options.keep]      --keep-Angaben.
 * @param {string} [options.out]         Zieldatei (.xlsx); Standard: neben der Datenbank.
 * @param {boolean} [options.dryRun]     Nur planen, nichts schreiben.
 * @param {number} [options.seed]
 * @param {boolean} [options.consistent]
 * @param {string} [options.password]
 * @returns {Promise<object>} Bericht ueber den Lauf.
 */
export async function anonymizeDatabase({
  file,
  table,
  keep = [],
  clear = [],
  out,
  dryRun = false,
  seed,
  consistent = true,
  password,
}) {
  const tables = await readDatabase(file, { password });
  const { report, result } = planAnonymization(tables, { keep, clear, table, seed, consistent });

  if (dryRun) return report;

  const target = out ? path.resolve(out) : defaultTarget(file);
  report.warnings.push(...await writeWorkbook(target, result));
  report.output = target;

  return report;
}

/** "kunden.accdb" -> "kunden.maskiert.xlsx" */
function defaultTarget(file) {
  const resolved = path.resolve(file);
  const extension = path.extname(resolved);
  return path.join(
    path.dirname(resolved),
    `${path.basename(resolved, extension)}.maskiert.xlsx`,
  );
}
