/**
 * Schreibt einen Aenderungsplan in eine Access-Datenbank.
 *
 * Zum Schreiben fuehrt kein Weg an ACE-OLEDB vorbei, und das gibt es nur
 * unter Windows. Statt eines nativen Node-Moduls (das sich weder buendeln noch
 * ohne Uebersetzer installieren liesse) wird PowerShell aufgerufen, das auf
 * jedem Windows vorhanden ist.
 *
 * Das Skript geht ueber die Standardeingabe an "powershell -Command -": So
 * wird es nicht als Skriptdatei ausgefuehrt, die Ausfuehrungsrichtlinie greift
 * also nicht und es braucht kein "Bypass".
 */

import { spawn } from 'node:child_process';
import { writeFile, rm, mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';

import { APPLY_SCRIPT } from './apply-script.js';

/**
 * @param {string} database Pfad zur Datenbank, die geaendert wird.
 * @param {object} plan     Rueckgabe von planAnonymization().
 * @returns {Promise<{tables: Array, errors: string[]}>}
 */
export async function applyPlan(database, plan) {
  if (process.platform !== 'win32') {
    throw new Error(
      'Das Schreiben in Access-Datenbanken ist nur unter Windows moeglich, weil dafuer die\n' +
      'Microsoft Access Database Engine (ACE-OLEDB) benoetigt wird.\n' +
      'Lesen, --list und --dry-run funktionieren auf jedem System.',
    );
  }

  const workdir = await mkdtemp(path.join(tmpdir(), 'anonymizer-access-'));
  const planFile = path.join(workdir, 'plan.json');

  try {
    await writeFile(planFile, JSON.stringify(serializePlan(plan)), 'utf8');
    const output = await runPowerShell({
      ANONYMIZER_DB: path.resolve(database),
      ANONYMIZER_PLAN: planFile,
    });
    return parseResult(output);
  } finally {
    await rm(workdir, { recursive: true, force: true });
  }
}

/**
 * Datumswerte werden als ISO-Zeichenkette uebergeben - JSON kennt keinen
 * Datumstyp. Die Spaltentypen stehen im Plan, PowerShell wandelt anhand
 * derer zurueck.
 */
export function serializePlan(plan) {
  return {
    tables: plan.tables.map((table) => ({
      ...table,
      updates: table.updates.map((update) => ({
        k: serializeValue(update.k),
        v: Object.fromEntries(
          Object.entries(update.v).map(([column, value]) => [column, serializeValue(value)]),
        ),
      })),
    })),
  };
}

function serializeValue(value) {
  return value instanceof Date ? value.toISOString() : value;
}

function runPowerShell(env) {
  return new Promise((resolve, reject) => {
    const child = spawn('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', '-'], {
      env: { ...process.env, ...env },
      windowsHide: true,
    });

    let stdout = '';
    let stderr = '';
    child.stdout.on('data', (chunk) => { stdout += chunk; });
    child.stderr.on('data', (chunk) => { stderr += chunk; });

    child.on('error', (error) => reject(
      new Error(`PowerShell konnte nicht gestartet werden: ${error.message}`),
    ));

    child.on('close', (code) => {
      if (code === 0) return resolve(stdout);
      reject(new Error(`Das Schreiben ist fehlgeschlagen:\n${(stderr || stdout).trim()}`));
    });

    child.stdin.write(APPLY_SCRIPT);
    child.stdin.end();
  });
}

/**
 * Das Skript gibt das Ergebnis als letzte Zeile aus. Davor koennen Meldungen
 * von PowerShell stehen, deshalb wird von hinten nach der ersten Zeile
 * gesucht, die sich als JSON lesen laesst.
 */
export function parseResult(output) {
  const text = String(output).trim();
  if (!text) throw new Error('Unerwartete Ausgabe beim Schreiben: (leer)');

  const lines = text.split(/\r?\n/).map((line) => line.trim()).filter(Boolean);

  for (let index = lines.length - 1; index >= 0; index -= 1) {
    if (!lines[index].startsWith('{')) continue;

    try {
      const parsed = JSON.parse(lines[index]);
      return {
        tables: toArray(parsed.tables),
        errors: toArray(parsed.errors).map(String),
      };
    } catch {
      // naechste Zeile versuchen
    }
  }

  throw new Error(`Unerwartete Ausgabe beim Schreiben:\n${text}`);
}

/** PowerShell macht aus einelementigen Feldern ein einzelnes Objekt. */
function toArray(value) {
  if (value === null || value === undefined) return [];
  return Array.isArray(value) ? value : [value];
}
