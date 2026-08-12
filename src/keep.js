/**
 * Zerlegt die --keep-Angaben.
 *
 * Zwei Schreibweisen, beliebig kombinierbar und fuer beliebig viele Blaetter:
 *
 *   --keep "KundenID"                    gilt in JEDEM Arbeitsblatt
 *   --keep "Kunden:KundenID"             gilt nur im Blatt "Kunden"
 *   --keep "Kunden:ID,Bestellungen:Nr"   je Blatt eigene Spalten
 *
 * Der Doppelpunkt eignet sich als Trenner, weil Excel ihn in Blattnamen nicht
 * zulaesst. Spaltennamen duerfen ihn enthalten - getrennt wird am ersten.
 */

/** @typedef {{global: Set<string>, perSheet: Map<string, Set<string>>, raw: string[]}} KeepRules */

/**
 * @param {string[]} entries Rohe Angaben, z. B. ["Kunden:ID", "Ort"].
 * @returns {KeepRules}
 */
export function parseKeep(entries) {
  const rules = { global: new Set(), perSheet: new Map(), raw: [] };

  for (const entry of entries) {
    const text = String(entry).trim();
    if (!text) continue;
    rules.raw.push(text);

    const separator = text.indexOf(':');
    if (separator === -1) {
      rules.global.add(normalize(text));
      continue;
    }

    const sheet = normalize(text.slice(0, separator));
    const column = normalize(text.slice(separator + 1));
    if (!sheet || !column) {
      throw new Error(
        `Ungueltige --keep-Angabe: "${text}"\n` +
        'Erwartet wird "Spalte" oder "Blattname:Spalte".',
      );
    }

    let columns = rules.perSheet.get(sheet);
    if (!columns) {
      columns = new Set();
      rules.perSheet.set(sheet, columns);
    }
    columns.add(column);
  }

  return rules;
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
