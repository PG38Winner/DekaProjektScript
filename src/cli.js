#!/usr/bin/env node
/**
 * Kommandozeilen-Oberflaeche des Excel-Anonymisierers.
 *
 *   node src/cli.js daten.xlsx --list
 *   node src/cli.js daten.xlsx --keep "KundenID,Bestellnummer"
 *   node src/cli.js daten.xlsx --sheet Kunden --out anonym.xlsx --no-backup
 */

import { parseArgs } from 'node:util';
import path from 'node:path';

import { anonymizeWorkbook, inspectWorkbook } from './excel/anonymize.js';
import { anonymizeDatabase, inspectDatabase } from './access/anonymize.js';

/** Waehlt die Engine anhand der Dateiendung. */
const ENGINES = {
  '.xlsx': 'excel', '.xlsm': 'excel',
  '.accdb': 'access', '.mdb': 'access',
};

function engineFor(file) {
  const extension = path.extname(file).toLowerCase();
  const engine = ENGINES[extension];
  if (engine) return engine;

  throw new Error(
    `Nicht unterstuetzte Dateiendung "${extension || '(keine)'}".\n` +
    'Unterstuetzt werden: .xlsx, .xlsm (Excel) sowie .accdb, .mdb (Access).',
  );
}

const USAGE = `
Daten-Anonymisierer - verschleiert personenbezogene Daten
in Excel-Dateien (.xlsx/.xlsm) und Access-Datenbanken (.accdb/.mdb).
Die Engine wird an der Dateiendung erkannt.

Aufruf:
  anonymisieren <datei> [optionen]

Optionen:
  --list                Blaetter/Tabellen und Spalten anzeigen (nichts aendern)
  --sheet <name>        Nur dieses Blatt bzw. diese Tabelle (auch: --table);
                        Standard: ALLE
  --keep <angabe>       Spalten, die UNVERAENDERT bleiben.
                          --keep "KundenID"          gilt in JEDEM Blatt
                          --keep "Kunden:KundenID"   nur in Kunden
                        Ein Blatt-Praefix gilt fuer alle folgenden Spalten der
                        Angabe, bis ein neues Praefix kommt - so lassen sich je
                        Blatt beliebig viele Spalten nennen:
                          --keep "Kunden:ID,Name,Ort"
                          --keep "Kunden:ID,Name,Artikel:Nr,Preis"
                          --keep "*:ID"              wieder fuer jedes Blatt
                        Mehrfach angebbar; jede Angabe beginnt neu:
                          --keep "Kunden:ID,Name" --keep "Artikel:Nr"
  --out <datei>         Zieldatei. Bei Excel statt Ueberschreiben; bei Access
                        die zu schreibende Arbeitsmappe (Standard:
                        <datenbank>.anonymisiert.xlsx)
  --no-backup           Keine Sicherungskopie anlegen
  --no-consistent       Gleiche Werte muessen nicht denselben Ersatz erhalten
  --seed <zahl>         Fester Startwert - erzeugt reproduzierbare Ergebnisse
  --dry-run             Nur anzeigen, was passieren wuerde
  --full                --list liest die Datei vollstaendig (nur Excel).
                        Standard ist der sparsame Weg, der auch mit sehr
                        grossen Dateien zurechtkommt.
  --password <wort>     Kennwort der Access-Datenbank
  -h, --help            Diese Hilfe

Access:
  Laeuft vollstaendig in Node.js, auf jedem System und ohne Zusatzsoftware.
  Das Ergebnis wird als Excel-Arbeitsmappe geschrieben, je Tabelle ein Blatt -
  die Datenbank selbst bleibt unveraendert. In eine .accdb zurueckzuschreiben
  kann nur die Microsoft Access Database Engine; aus reinem JavaScript geht
  das nicht.
  Schluessel- und Verknuepfungsspalten werden erkannt und bleiben stehen.

Merksatz:
  In --keep genannt = bleibt unveraendert - alle anderen Spalten werden verschleiert.

Beispiele:
  anonymisieren daten.xlsx --list
  anonymisieren daten.xlsx --keep "Kunden:KundenID,Nachname" --keep "Artikel:Nr"
  anonymisieren daten.accdb --dry-run
  anonymisieren daten.accdb --keep "Kunden:Kundennummer" --out anonym.xlsx
`.trim();

const OPTIONS = {
  list: { type: 'boolean', default: false },
  sheet: { type: 'string' },
  table: { type: 'string' },
  password: { type: 'string' },
  keep: { type: 'string', multiple: true },
  out: { type: 'string' },
  backup: { type: 'boolean', default: true },
  consistent: { type: 'boolean', default: true },
  seed: { type: 'string' },
  'dry-run': { type: 'boolean', default: false },
  full: { type: 'boolean', default: false },
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
  const engine = engineFor(file);
  const section = values.sheet ?? values.table;

  if (values.list) {
    await printStructure(engine, file, values);
    return;
  }

  const seed = parseSeed(values.seed);
  // Roh weiterreichen: das Zerlegen bleibt in keep.js, weil ein Blatt-Praefix
  // innerhalb einer Angabe fuer die folgenden Spalten weitergilt.
  const keep = values.keep ?? [];

  const common = {
    file,
    keep,
    out: values.out,
    backup: values.backup,
    dryRun: values['dry-run'],
    consistent: values.consistent,
    seed,
  };

  const report = engine === 'access'
    ? await anonymizeDatabase({ ...common, table: section, password: values.password })
    : await anonymizeWorkbook({ ...common, sheet: section });

  printReport(report, values['dry-run']);
}

async function printStructure(engine, file, values) {
  const sheets = engine === 'access'
    ? await inspectDatabase(file, { password: values.password })
    : await inspectWorkbook(file, { full: values.full });

  const label = engine === 'access' ? 'Tabelle' : 'Arbeitsblatt';

  for (const sheet of sheets) {
    // Die Zeilenzahl stammt aus dem Kopf des Blattes; fehlt sie dort, wird sie
    // nicht eigens ermittelt - das wuerde die Datei komplett einlesen.
    const rows = sheet.rowCount === null ? 'Zeilenzahl unbekannt' : `${sheet.rowCount} Datenzeilen`;
    console.log(`\n${label}: ${sheet.name}  (${rows})`);
    if (!sheet.columns.length) {
      console.log('  (keine Spalten gefunden)');
      continue;
    }
    for (const column of sheet.columns) {
      const kind = column.readOnly ? 'wird uebersprungen' : column.kind;
      const note = column.note ? `  [${column.note}]` : '';
      console.log(`  ${column.header.padEnd(28)} ${kind}${note}`);
    }
  }
  console.log('');
}

function printReport(report, dryRun) {
  const label = report.label ?? 'Arbeitsblatt';

  for (const sheet of report.sheets) {
    console.log(`\n${label}: ${sheet.name}  (${sheet.rows} Datenzeilen)`);

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
