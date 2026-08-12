/**
 * Zerlegt die --keep-Angaben.
 *
 * Ein Blattname vor dem Doppelpunkt legt fest, fuer welches Arbeitsblatt die
 * folgenden Spalten gelten - und zwar so lange, bis ein neues Praefix kommt.
 * Damit lassen sich je Blatt beliebig viele Spalten in einer Angabe nennen:
 *
 *   --keep "KundenID"                          gilt in JEDEM Blatt
 *   --keep "Kunden:ID,Name,Ort"                drei Spalten nur im Blatt Kunden
 *   --keep "Kunden:ID,Name,Artikel:Nr,Preis"   wechselt beim naechsten Praefix
 *   --keep "ID,Kunden:Name"                    ID ueberall, Name nur in Kunden
 *   --keep "Kunden:Name,*:ID"                  "*" schaltet zurueck auf ueberall
 *
 * Jede einzelne --keep-Angabe beginnt wieder bei "gilt ueberall"; ein Praefix
 * wirkt also nie ueber die Angabe hinaus, in der es steht.
 *
 * Der Doppelpunkt eignet sich als Trenner, weil Excel ihn in Blattnamen nicht
 * zulaesst. Spaltennamen duerfen ihn enthalten - getrennt wird am ersten.
 */

/** Praefix, das ausdruecklich wieder alle Blaetter meint. */
const ALL_SHEETS = '*';

/** @typedef {{global: Set<string>, perSheet: Map<string, Set<string>>, raw: string[]}} KeepRules */

/**
 * @param {string[]} entries Rohe Angaben - je Element eine --keep-Angabe,
 *                           die ihrerseits eine Kommaliste sein darf.
 * @returns {KeepRules}
 */
export function parseKeep(entries) {
  const rules = { global: new Set(), perSheet: new Map(), raw: [] };

  for (const entry of entries) {
    const text = String(entry).trim();
    if (!text) continue;
    rules.raw.push(text);

    // Gilt nur innerhalb dieser Angabe; null bedeutet "alle Blaetter".
    let currentSheet = null;

    for (const part of splitList(text)) {
      const separator = part.indexOf(':');

      if (separator === -1) {
        addColumn(rules, currentSheet, part, text);
        continue;
      }

      const prefix = part.slice(0, separator).trim();
      const column = part.slice(separator + 1).trim();
      if (!prefix || !column) {
        throw new Error(
          `Ungueltige --keep-Angabe: "${part}"\n` +
          'Erwartet wird "Spalte", "Blattname:Spalte" oder "*:Spalte".',
        );
      }

      currentSheet = prefix === ALL_SHEETS ? null : normalize(prefix);
      addColumn(rules, currentSheet, column, text);
    }
  }

  return rules;
}

function addColumn(rules, sheet, column, context) {
  const name = normalize(column);
  if (!name) throw new Error(`Ungueltige --keep-Angabe: "${context}" enthaelt einen leeren Namen.`);

  if (sheet === null) {
    rules.global.add(name);
    return;
  }

  let columns = rules.perSheet.get(sheet);
  if (!columns) {
    columns = new Set();
    rules.perSheet.set(sheet, columns);
  }
  columns.add(name);
}

/** Zerlegt eine Kommaliste; erlaubt maskierte Kommas mit \, im Spaltennamen. */
export function splitList(value) {
  return String(value)
    .split(/(?<!\\),/)
    .map((part) => part.replace(/\\,/g, ',').trim())
    .filter(Boolean);
}

/** Bleibt diese Spalte dieses Blattes unveraendert? */
export function isKept(rules, sheetName, header) {
  const column = normalize(header);
  if (rules.global.has(column)) return true;

  const columns = rules.perSheet.get(normalize(sheetName));
  return Boolean(columns && columns.has(column));
}

/**
 * Prueft die Angaben gegen die tatsaechlich vorhandenen Blaetter und Spalten.
 * Ein Tippfehler soll auffallen, statt stillschweigend wirkungslos zu bleiben.
 *
 * @param {KeepRules} rules
 * @param {Array<{sheet: {name: string}, columns: Array<{header: string}>}>} prepared
 */
export function validateKeep(rules, prepared) {
  const sheetsByName = new Map(
    prepared.map(({ sheet, columns }) => [
      normalize(sheet.name),
      { name: sheet.name, columns: columns.map((column) => column.header) },
    ]),
  );

  const problems = [];

  for (const column of rules.global) {
    const found = prepared.some(({ columns }) =>
      columns.some((candidate) => normalize(candidate.header) === column));
    if (!found) problems.push(`Spalte "${column}" kommt in keinem verarbeiteten Blatt vor.`);
  }

  for (const [sheetName, columns] of rules.perSheet) {
    const sheet = sheetsByName.get(sheetName);
    if (!sheet) {
      problems.push(
        `Arbeitsblatt "${sheetName}" wird nicht verarbeitet. ` +
        `Verarbeitet werden: ${[...sheetsByName.values()].map((s) => s.name).join(', ')}`,
      );
      continue;
    }

    for (const column of columns) {
      const found = sheet.columns.some((header) => normalize(header) === column);
      if (!found) {
        problems.push(
          `Spalte "${column}" gibt es im Blatt "${sheet.name}" nicht. ` +
          `Vorhanden: ${sheet.columns.join(', ')}`,
        );
      }
    }
  }

  if (problems.length) {
    throw new Error(`Fehlerhafte --keep-Angabe:\n  ${problems.join('\n  ')}`);
  }
}

function normalize(value) {
  return String(value).trim().toLowerCase();
}
