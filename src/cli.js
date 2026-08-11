#!/usr/bin/env node
/**
 * Kommandozeilen-Oberflaeche des Excel-Anonymisierers.
 *
 *   node src/cli.js daten.xlsx --list
 *   node src/cli.js daten.xlsx --keep "KundenID,Bestellnummer"
 *   node src/cli.js daten.xlsx --sheet Kunden --out anonym.xlsx --no-backup
 */

import { parseArgs } from 'node:util';
import { anonymizeWorkbook, inspectWorkbook } from './anonymize.js';

const USAGE = `
Excel-Anonymisierer - verschleiert personenbezogene Daten in .xlsx/.xlsm

Aufruf:
  anonymize-xlsx <datei> [optionen]

Optionen:
  --list                Blaetter und Spalten anzeigen (nichts veraendern)
  --sheet <name>        Nur dieses Arbeitsblatt; Standard: ALLE Blaetter
  --keep <a,b,c>        Spalten, die UNVERAENDERT bleiben (z. B. Schluessel, IDs)
  --out <datei>         Ergebnis in neue Datei schreiben statt zu ueberschreiben
  --no-backup           Keine Sicherungskopie anlegen
  --no-consistent       Gleiche Werte muessen nicht denselben Ersatz erhalten
  --seed <zahl>         Fester Startwert - erzeugt reproduzierbare Ergebnisse
  --dry-run             Nur anzeigen, was passieren wuerde
  -h, --help            Diese Hilfe

Merksatz:
  In --keep genannt = bleibt unveraendert - alle anderen Spalten werden verschleiert.
`.trim();

const OPTIONS = {
  list: { type: 'boolean', default: false },
  sheet: { type: 'string' },
  keep: { type: 'string' },
  out: { type: 'string' },
  backup: { type: 'boolean', default: true },
  consistent: { type: 'boolean', default: true },
  seed: { type: 'string' },
  'dry-run': { type: 'boolean', default: false },
  help: { type: 'boolean', short: 'h', default: false },
};

async function main() {
  let parsed;
  try {
    parsed = parseArgs({ options: OPTIONS, allowPositionals: true, allowNegative: true });
  } catch (error) {
    fail(`${error.message}\n\n${USAGE}`);
  }

  const { values, positionals } = parsed;

  if (values.help || positionals.length === 0) {
    console.log(USAGE);
    process.exit(values.help ? 0 : 1);
  }

  if (positionals.length > 1) {
    fail(`Es wird genau eine Datei erwartet, erhalten: ${positionals.join(', ')}`);
  }

  const file = positionals[0];

  if (values.list) {
    await printStructure(file);
    return;
  }

  const seed = parseSeed(values.seed);
  const keep = values.keep
    ? values.keep.split(',').map((s) => s.trim()).filter(Boolean)
    : [];

  const report = await anonymizeWorkbook({
    file,
    sheet: values.sheet,
    keep,
    out: values.out,
    backup: values.backup,
    dryRun: values['dry-run'],
    consistent: values.consistent,
    seed,
  });

  printReport(report, values['dry-run']);
}

async function printStructure(file) {
  const sheets = await inspectWorkbook(file);

  for (const sheet of sheets) {
    // Die Zeilenzahl stammt aus dem Kopf des Blattes; fehlt sie dort, wird sie
    // nicht eigens ermittelt - das wuerde die Datei komplett einlesen.
    const rows = sheet.rowCount === null ? 'Zeilenzahl unbekannt' : `${sheet.rowCount} Datenzeilen`;
    console.log(`\nArbeitsblatt: ${sheet.name}  (${rows})`);
    if (!sheet.columns.length) {
      console.log('  (keine Kopfzeile mit Spaltennamen gefunden)');
      continue;
    }
    for (const column of sheet.columns) {
      const kind = column.readOnly ? 'Formel - wird uebersprungen' : column.kind;
      console.log(`  ${column.header.padEnd(28)} ${kind}`);
    }
  }
  console.log('');
}

function printReport(report, dryRun) {
  for (const sheet of report.sheets) {
    console.log(`\nArbeitsblatt: ${sheet.name}  (${sheet.rows} Datenzeilen)`);

    if (sheet.skipped) {
      console.log(`  uebersprungen - ${sheet.skipped}`);
      continue;
    }

    console.log('');
    for (const column of sheet.columns) {
      const count = column.changed ? `${column.changed} Zellen` : '';
      console.log(
        `  ${column.header.padEnd(28)} ${column.status.padEnd(22)} ${column.kind.padEnd(10)} ${count}`,
      );
    }
  }

  console.log(`\n  Geaenderte Zellen gesamt: ${report.changed}`);
  if (report.backup) console.log(`  Sicherungskopie:          ${report.backup}`);
  if (report.macrosPreserved) console.log('  Makros:                   uebernommen');
  if (report.output) console.log(`  Geschrieben:              ${report.output}`);

  for (const warning of report.warnings ?? []) {
    console.warn(`\n  WARNUNG: ${warning}`);
  }

  if (dryRun) console.log('\n  --dry-run: Es wurde nichts geschrieben.');
  console.log('');
}

function parseSeed(raw) {
  if (raw === undefined) return undefined;
  const seed = Number(raw);
  if (!Number.isFinite(seed)) fail(`--seed erwartet eine Zahl, erhalten: ${raw}`);
  return seed;
}

function fail(message) {
  console.error(`Fehler: ${message}`);
  process.exit(1);
}

main().catch((error) => {
  console.error(`Fehler: ${error.message}`);
  process.exit(1);
});
