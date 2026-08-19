#!/usr/bin/env node
/**
 * Kommandozeilen-Oberflaeche des Daten-Maskierers.
 *
 *   node src/cli.js daten.xlsx --list
 *   node src/cli.js daten.xlsx --keep "Kunden:KundenID"
 *   node src/cli.js daten.accdb --dry-run
 */

import { parseArgs } from 'node:util';
import path from 'node:path';

import { anonymizeWorkbook, inspectWorkbook } from './excel/anonymize.js';
import { anonymizeDatabase, inspectDatabase } from './access/anonymize.js';

/** Fassung; bei jeder Freigabe zusammen mit package.json anheben. */
const VERSION = '1.0.0';

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
Daten-Maskierer - ersetzt personenbezogene Daten durch erfundene Werte
in Excel-Dateien (.xlsx/.xlsm) und Access-Datenbanken (.accdb/.mdb).
Die Engine wird an der Dateiendung erkannt.

WICHTIG: Das Ergebnis ist eine MASKIERUNG, keine zertifizierte Anonymisierung.
Ob sie fuer den jeweiligen Zweck ausreicht, muss fachlich beurteilt werden -
siehe Abschnitt "Grenzen der Maskierung" in der README.

Aufruf:
  maskieren <datei> [optionen]

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
  --clear <angabe>      Spalten, die GELEERT werden - der Inhalt wird nicht
                        ersetzt, sondern entfernt. Gedacht fuer Freitextfelder,
                        in denen alles Moegliche stehen kann. Schreibweise wie
                        bei --keep:
                          --clear "Kunden:Bemerkung,Notiz"
  --out <datei>         Zieldatei. Standard: <name>.maskiert.<endung>
                        neben der Quelldatei.
  --in-place            Die Quelldatei SELBST ueberschreiben (nur Excel).
                        Zeigt vorher, was geaendert wuerde, und verlangt eine
                        ausdrueckliche Bestaetigung. Eine Sicherungskopie wird
                        immer angelegt und laesst sich nicht abschalten.
  --yes                 Bestaetigung zu --in-place vorab erteilen (fuer
                        Aufrufe ohne Terminal)
  --no-consistent       Gleiche Werte muessen nicht denselben Ersatz erhalten
  --seed <zahl>         Fester Startwert - erzeugt reproduzierbare Ergebnisse
  --dry-run             Nur anzeigen, was passieren wuerde
  --full                --list liest die Datei vollstaendig (nur Excel).
                        Standard ist der sparsame Weg, der auch mit sehr
                        grossen Dateien zurechtkommt.
  --password <wort>     Kennwort der Access-Datenbank
  --version             Fassung und Pruefsumme ausgeben
  -h, --help            Diese Hilfe

Access:
  Die Datenbank wird ausschliesslich GELESEN. Das Ergebnis wird als
  Excel-Arbeitsmappe geschrieben, je Tabelle ein Blatt. Es gibt keinen Weg,
  in dem die .accdb veraendert wuerde.
  Schluessel- und Verknuepfungsspalten werden erkannt und bleiben stehen.

Merksatz:
  In --keep genannt = bleibt unveraendert - alle anderen Spalten werden verschleiert.

Beispiele:
  maskieren daten.xlsx --list
  maskieren daten.xlsx --keep "Kunden:KundenID,Nachname" --keep "Artikel:Nr"
  maskieren daten.xlsx --clear "Kunden:Bemerkung"
  maskieren daten.accdb --dry-run
  maskieren daten.xlsx --in-place            (ueberschreibt, legt Sicherung an)
`.trim();

const OPTIONS = {
  list: { type: 'boolean', default: false },
  sheet: { type: 'string' },
  table: { type: 'string' },
  password: { type: 'string' },
  'in-place': { type: 'boolean', default: false },
  yes: { type: 'boolean', default: false },
  version: { type: 'boolean', default: false },
  keep: { type: 'string', multiple: true },
  clear: { type: 'string', multiple: true },
  out: { type: 'string' },
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

  if (values.version) {
    await printVersion();
    return;
  }

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

  if (values['in-place'] && engine === 'access') {
    fail('--in-place gibt es fuer Access nicht: die Datenbank wird nur gelesen.');
  }

  // Standard ist die neue Datei. Die Quelldatei wird nur auf ausdrueckliche
  // Anweisung ueberschrieben - und dann immer mit Sicherungskopie.
  const common = {
    file,
    keep,
    clear: values.clear ?? [],
    out: values['in-place'] ? undefined : (values.out ?? defaultTarget(file)),
    backup: true,
    dryRun: values['dry-run'],
    consistent: values.consistent,
    seed,
  };

  const run = (options) => (engine === 'access'
    ? anonymizeDatabase({ ...options, table: section, password: values.password })
    : anonymizeWorkbook({ ...options, sheet: section }));

  // Vor dem Ueberschreiben der Quelldatei: erst zeigen, was passieren wuerde,
  // dann ausdruecklich bestaetigen lassen. Ein unbeabsichtigter Lauf auf einer
  // Produktivdatei laesst sich sonst nicht mehr rueckgaengig machen.
  if (values['in-place'] && !values['dry-run']) {
    printReport(await run({ ...common, dryRun: true }), true);
    await confirmOverwrite(file, values.yes);
  }

  printReport(await run(common), values['dry-run']);
}

/**
 * Verlangt eine ausdrueckliche Bestaetigung. Ohne Terminal (Aufruf aus einem
 * Skript) muss --yes angegeben werden - stillschweigend ueberschrieben wird
 * nie.
 */
async function confirmOverwrite(file, alreadyConfirmed) {
  if (alreadyConfirmed) return;

  if (!process.stdin.isTTY) {
    fail(
      `"${path.basename(file)}" soll ueberschrieben werden, es ist aber keine Eingabe moeglich.\n`
      + 'Bei einem Aufruf ohne Terminal die Bestaetigung mit --yes angeben.',
    );
  }

  const { createInterface } = await import('node:readline/promises');
  const rl = createInterface({ input: process.stdin, output: process.stdout });

  console.log(
    `\n  Die Datei "${path.basename(file)}" wird UEBERSCHRIEBEN.`
    + '\n  Eine Sicherungskopie wird daneben angelegt.'
    + '\n  Bitte bestaetigen, dass die oben als unveraendert genannten Spalten'
    + '\n  keine personenbezogenen Daten enthalten.',
  );

  const answer = await rl.question('\n  Zum Fortfahren JA eingeben: ');
  rl.close();

  if (answer.trim().toUpperCase() !== 'JA') {
    console.log('\n  Abgebrochen. Es wurde nichts geaendert.\n');
    process.exit(1);
  }
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

  printUnchanged(report, label);

  console.log(`\n  Maskierte Zellen gesamt:  ${report.changed}`);
  if (report.backup) console.log(`  Sicherungskopie:          ${report.backup}`);
  if (report.macrosPreserved) console.log('  Makros:                   uebernommen');
  if (report.output) console.log(`  Geschrieben:              ${report.output}`);

  for (const warning of report.warnings ?? []) {
    console.warn(`\n  WARNUNG: ${warning}`);
  }

  console.log(
    '\n  HINWEIS: Dies ist eine Maskierung, keine zertifizierte Anonymisierung.'
    + '\n  Freitextfelder koennen personenbezogene Angaben enthalten, die nicht als'
    + '\n  solche erkannt werden, und Kombinationen aus Datum, Ort, Betrag oder'
    + '\n  seltenen Merkmalen koennen eine Re-Identifikation ermoeglichen.',
  );

  if (dryRun) console.log('\n  --dry-run: Es wurde nichts geschrieben.');
  console.log('');
}

/**
 * Listet auf, was NICHT maskiert wurde.
 *
 * Der haeufigste Bedienfehler ist die Verwechslung der Richtung: --keep nimmt
 * Spalten von der Maskierung AUS. Wer das umgekehrt versteht, laesst genau die
 * personenbezogenen Spalten stehen. Deshalb stehen sie am Ende noch einmal
 * ausdruecklich beisammen.
 */
function printUnchanged(report, label) {
  const untouched = report.sheets
    .filter((sheet) => !sheet.skipped)
    .map((sheet) => ({
      name: sheet.name,
      columns: sheet.columns
        .filter((column) => column.status.startsWith('unveraendert'))
        .map((column) => column.header),
    }))
    .filter((sheet) => sheet.columns.length);

  if (!untouched.length) return;

  console.log('\n  ACHTUNG - diese Spalten enthalten weiterhin die Originaldaten:');
  for (const sheet of untouched) {
    console.log(`    ${label} ${sheet.name}: ${sheet.columns.join(', ')}`);
  }
  console.log('    Bitte pruefen, dass darin keine personenbezogenen Angaben stehen.');
}

/**
 * Gibt Fassung und Pruefsumme der laufenden Datei aus. Damit laesst sich
 * festhalten, welcher Stand freigegeben wurde und ob spaeter derselbe laeuft.
 */
async function printVersion() {
  const { createHash } = await import('node:crypto');
  const { readFile } = await import('node:fs/promises');
  const running = process.argv[1];

  console.log(`Daten-Maskierer ${VERSION}`);
  console.log(`Node.js         ${process.version}`);

  try {
    const hash = createHash('sha256').update(await readFile(running)).digest('hex');
    console.log(`Datei           ${running}`);
    console.log(`SHA-256         ${hash}`);
  } catch {
    console.log(`Datei           ${running} (Pruefsumme nicht ermittelbar)`);
  }
}

/** "kunden.xlsx" -> "kunden.maskiert.xlsx" (Endung bleibt, wegen Makros). */
function defaultTarget(file) {
  const resolved = path.resolve(file);
  const extension = path.extname(resolved);
  return path.join(
    path.dirname(resolved),
    `${path.basename(resolved, extension)}.maskiert${extension}`,
  );
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
