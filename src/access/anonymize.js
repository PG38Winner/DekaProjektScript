/**
 * Anonymisiert Access-Datenbanken (.accdb / .mdb).
 *
 * Ablauf:
 *   1. Lesen mit mdb-reader (reines JavaScript, ueberall lauffaehig)
 *   2. Planen in plan.js (kennt weder Access noch COM, vollstaendig geprueft)
 *   3. Schreiben ueber PowerShell/ACE-OLEDB (nur Windows)
 *
 * Die ersten beiden Schritte genuegen fuer `--list` und `--dry-run`; wer nur
 * sehen will, was passieren wuerde, braucht kein Windows.
 */

import { copyFile } from 'node:fs/promises';
import path from 'node:path';

import { readDatabase } from './read.js';
import { planAnonymization } from './plan.js';
import { applyPlan } from './write.js';

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
        note: column.status.startsWith('unveraendert (Schluessel)') ? 'Schluessel' : null,
      })),
    };
  });
}

/**
 * @param {object} options
 * @param {string} options.file          Eingabedatenbank.
 * @param {string} [options.table]       Nur diese Tabelle; Standard: alle.
 * @param {string[]} [options.keep]      --keep-Angaben.
 * @param {string} [options.out]         Ergebnis in eine Kopie schreiben.
 * @param {boolean} [options.backup]     Sicherungskopie anlegen (Standard: true).
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
  out,
  backup = true,
  dryRun = false,
  seed,
  consistent = true,
  password,
}) {
  const tables = await readDatabase(file, { password });
  const { report, plan } = planAnonymization(tables, { keep, table, seed, consistent });

  if (dryRun) return report;

  // Anders als bei Excel wird nicht neu geschrieben, sondern in der Datei
  // geaendert. Fuer --out wird deshalb zuerst kopiert und dann die Kopie
  // bearbeitet - das Original bleibt so garantiert unberuehrt.
  const target = out ? path.resolve(out) : path.resolve(file);
  if (out) await copyFile(path.resolve(file), target);
  else if (backup) report.backup = await createBackup(file);

  if (!plan.tables.length) {
    report.output = target;
    report.warnings.push('Es gab nichts zu aendern.');
    return report;
  }

  const result = await applyPlan(target, plan);
  report.output = target;
  report.warnings.push(...result.errors);

  for (const applied of result.tables) {
    const entry = report.sheets.find((sheet) => sheet.name === applied.name);
    if (entry) entry.applied = applied.updated;

    if (applied.unmatched > 0) {
      report.warnings.push(
        `In "${applied.name}" wurden ${applied.unmatched} geplante Datensaetze nicht ` +
        'wiedergefunden. Wurde die Datenbank zwischenzeitlich veraendert?',
      );
    }
  }

  return report;
}

/** Legt "datenbank.backup-<zeitstempel>.accdb" neben der Originaldatei an. */
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
