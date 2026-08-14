/**
 * Liest Kopfzeile und Beispielwerte einer Arbeitsmappe, ohne sie vollstaendig
 * zu laden.
 *
 * Warum eigener Code statt einer Bibliothek: `Workbook.xlsx.readFile()` holt
 * jede Zelle in den Speicher - bei grossen Dateien Gigabytes, bis der Lauf
 * scheitert. Der Datenstrom-Leser von exceljs ist zwar sparsam, hat sich an
 * echten Arbeitsmappen aber als unzuverlaessig erwiesen.
 *
 * Dieser Leser oeffnet die Teile des Archivs als Datenstrom und bricht ab,
 * sobald genug Zeilen gesehen wurden. Der Speicherbedarf haengt damit an der
 * Zahl der Beispielzeilen, nicht an der Dateigroesse.
 *
 * Gelesen werden nur die Teile, die fuer die Anzeige noetig sind:
 *   xl/workbook.xml       Blattnamen und Reihenfolge
 *   xl/worksheets/*.xml   die ersten Zeilen je Blatt
 *   xl/sharedStrings.xml  die darin verwendeten Texte
 *   xl/styles.xml         um Datumszellen von Zahlen zu unterscheiden
 */

import {
  openWorkbookZip, readText, listSheets, resolveRelationship, sheetPartPath, decodeXml,
} from './ooxml.js';

/** Zahlenformate, die Excel fest fuer Datum und Uhrzeit vergibt. */
const BUILTIN_DATE_FORMATS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47]);

/**
 * @param {string} file
 * @param {number} sampleRows Wie viele Datenzeilen je Blatt gelesen werden.
 * @returns {Promise<Array<{name, rowCount, columns}>>} columns: Map-artige Liste
 */
export async function scanWorkbook(file, sampleRows) {
  const zip = await openWorkbookZip(file);

  const workbookXml = await readText(zip, 'xl/workbook.xml');
  const rels = await readText(zip, 'xl/_rels/workbook.xml.rels');
  const sheetEntries = listSheets(workbookXml);
  if (!sheetEntries.length) throw new Error('Keine Arbeitsblaetter in der Datei gefunden.');

  const dateStyles = await readDateStyles(zip);

  const sheets = [];
  let maxStringIndex = -1;

  for (const entry of sheetEntries) {
    const target = resolveRelationship(rels, entry.rid);
    if (!target) continue;

    const scanned = await scanSheetPart(zip, sheetPartPath(target), sampleRows);
    maxStringIndex = Math.max(maxStringIndex, scanned.maxStringIndex);
    sheets.push({ name: entry.name, ...scanned });
  }

  // Erst jetzt steht fest, wie weit die Texttabelle gelesen werden muss.
  const strings = maxStringIndex >= 0 ? await readSharedStrings(zip, maxStringIndex) : [];

  return sheets.map((sheet) => ({
    name: sheet.name,
    rowCount: sheet.lastRow === null ? null : Math.max(sheet.lastRow - 1, 0),
    columns: buildColumns(sheet, strings, dateStyles),
  }));
}

/**
 * Liest ein Arbeitsblatt bis zur gewuenschten Zahl an Datenzeilen.
 * Zurueckgegeben werden die Rohzellen - Texte werden erst spaeter aufgeloest,
 * weil die Texttabelle im Archiv hinter den Blaettern liegen kann.
 */
async function scanSheetPart(zip, path, sampleRows) {
  const header = new Map();
  const samples = new Map();
  const readOnly = new Map();
  let lastRow = null;
  let dataRows = 0;
  let maxStringIndex = -1;

  await streamPart(zip, path, (chunk, stop) => {
    if (lastRow === null) {
      const dimension = /<dimension\b[^>]*\bref="([^"]*)"/.exec(chunk.text);
      if (dimension) lastRow = lastRowOfRange(dimension[1]);
    }

    for (const row of takeRows(chunk)) {
      if (row.number === 1) {
        for (const cell of parseCells(row.body)) {
          header.set(cell.column, cell);
          if (cell.type === 's') maxStringIndex = Math.max(maxStringIndex, Number(cell.value));
        }
        continue;
      }

      if (dataRows >= sampleRows) return stop();

      for (const cell of parseCells(row.body)) {
        if (cell.value === null && !cell.inline) continue;

        const counts = readOnly.get(cell.column) ?? { nonEmpty: 0, formula: 0 };
        counts.nonEmpty += 1;
        if (cell.formula) counts.formula += 1;
        readOnly.set(cell.column, counts);

        if (cell.formula) continue;

        const list = samples.get(cell.column) ?? [];
        if (list.length < sampleRows) list.push(cell);
        samples.set(cell.column, list);

        if (cell.type === 's') maxStringIndex = Math.max(maxStringIndex, Number(cell.value));
      }

      dataRows += 1;
    }
  });

  return { header, samples, readOnly, lastRow, maxStringIndex };
}

/** Loest Texte und Datumsformate auf und bildet daraus die Spaltenliste. */
function buildColumns(sheet, strings, dateStyles) {
  const columns = [];

  for (const [index, headerCell] of [...sheet.header].sort((a, b) => a[0] - b[0])) {
    const name = String(resolveValue(headerCell, strings, dateStyles) ?? '').trim();
    if (!name) continue;

    const counts = sheet.readOnly.get(index) ?? { nonEmpty: 0, formula: 0 };
    const values = (sheet.samples.get(index) ?? [])
      .map((cell) => resolveValue(cell, strings, dateStyles))
      .filter((value) => value !== null && value !== '');

    columns.push({
      header: name,
      index,
      values,
      readOnly: counts.nonEmpty > 0 && counts.formula === counts.nonEmpty,
    });
  }

  return columns;
}

/** Macht aus einer Rohzelle den Wert, den die Typerkennung erwartet. */
function resolveValue(cell, strings, dateStyles) {
  if (cell.inline !== null) return cell.inline;
  if (cell.value === null) return null;

  switch (cell.type) {
    case 's':
      return strings[Number(cell.value)] ?? null;
    case 'str':
    case 'e':
      return cell.value;
    case 'b':
      return cell.value === '1';
    default: {
      const numeric = Number(cell.value);
      if (!Number.isFinite(numeric)) return cell.value;
      return dateStyles.has(cell.style) ? excelSerialToDate(numeric) : numeric;
    }
  }
}

/**
 * Excel zaehlt Tage ab dem 30.12.1899. Der bewusst falsche 29.02.1900 im
 * Kalender von Excel liegt vor allen praktisch vorkommenden Daten und wird
 * durch diesen Bezugspunkt mit abgedeckt.
 */
function excelSerialToDate(serial) {
  return new Date(Math.round((serial - 25569) * 86400 * 1000));
}

/** Ermittelt, welche Formatvorlagen ein Datum darstellen. */
async function readDateStyles(zip) {
  const xml = await readText(zip, 'xl/styles.xml');
  const dateStyles = new Set();
  if (!xml) return dateStyles;

  // Eigene Formate gelten als Datum, wenn sie Tages-, Monats- oder
  // Jahresplatzhalter enthalten - Textbausteine in Anfuehrungszeichen zaehlen
  // dabei nicht mit.
  const custom = new Map();
  for (const match of xml.matchAll(/<numFmt\b[^>]*numFmtId="(\d+)"[^>]*formatCode="([^"]*)"/g)) {
    const code = decodeXml(match[2]).replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '');
    custom.set(Number(match[1]), /[ymdhs]/i.test(code));
  }

  const cellXfs = /<cellXfs\b[\s\S]*?<\/cellXfs>/.exec(xml);
  if (!cellXfs) return dateStyles;

  let styleIndex = 0;
  for (const match of cellXfs[0].matchAll(/<xf\b[^>]*>/g)) {
    const numFmtId = Number(/\bnumFmtId="(\d+)"/.exec(match[0])?.[1] ?? 0);
    if (BUILTIN_DATE_FORMATS.has(numFmtId) || custom.get(numFmtId)) dateStyles.add(styleIndex);
    styleIndex += 1;
  }

  return dateStyles;
}

/** Liest die Texttabelle nur so weit, wie sie tatsaechlich gebraucht wird. */
async function readSharedStrings(zip, maxIndex) {
  const strings = [];

  await streamPart(zip, 'xl/sharedStrings.xml', (chunk, stop) => {
    const pattern = /<si\b[^>]*>([\s\S]*?)<\/si>|<si\b[^>]*\/>/g;
    let match;
    let consumed = 0;

    while ((match = pattern.exec(chunk.text)) !== null) {
      strings.push(match[1] === undefined ? '' : textOf(match[1]));
      consumed = match.index + match[0].length;
      if (strings.length > maxIndex) break;
    }

    chunk.keepFrom(consumed);
    if (strings.length > maxIndex) stop();
  });

  return strings;
}

/** Setzt die Textbausteine einer Zeichenfolge zusammen. */
function textOf(xml) {
  let text = '';
  for (const match of xml.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)) text += decodeXml(match[1]);
  return text;
}

/** Schneidet vollstaendige <row>-Bloecke aus dem bisher gelesenen Text. */
function takeRows(chunk) {
  const rows = [];
  const pattern = /<row\b([^>]*)(?:\/>|>([\s\S]*?)<\/row>)/g;
  let match;
  let consumed = 0;

  while ((match = pattern.exec(chunk.text)) !== null) {
    const number = Number(/\br="(\d+)"/.exec(match[1])?.[1] ?? 0);
    rows.push({ number, body: match[2] ?? '' });
    consumed = match.index + match[0].length;
  }

  chunk.keepFrom(consumed);
  return rows;
}

/** Zerlegt die Zellen einer Zeile. */
function* parseCells(body) {
  const pattern = /<c\b([^>]*)(?:\/>|>([\s\S]*?)<\/c>)/g;
  let match;

  while ((match = pattern.exec(body)) !== null) {
    const attributes = match[1];
    const content = match[2] ?? '';
    const reference = /\br="([A-Z]+)\d+"/.exec(attributes);
    if (!reference) continue;

    const inline = /<is\b[^>]*>([\s\S]*?)<\/is>/.exec(content);
    const value = /<v\b[^>]*>([\s\S]*?)<\/v>/.exec(content);

    yield {
      column: columnIndex(reference[1]),
      type: /\bt="([^"]*)"/.exec(attributes)?.[1] ?? 'n',
      style: Number(/\bs="(\d+)"/.exec(attributes)?.[1] ?? -1),
      formula: /<f\b/.test(content),
      inline: inline ? textOf(inline[1]) : null,
      value: value ? decodeXml(value[1]) : null,
    };
  }
}

/** "AB" -> 28 */
function columnIndex(letters) {
  let index = 0;
  for (const character of letters) index = index * 26 + (character.charCodeAt(0) - 64);
  return index;
}

function lastRowOfRange(ref) {
  const end = String(ref).split(':').pop();
  const match = /(\d+)$/.exec(end ?? '');
  return match ? Number(match[1]) : null;
}

/**
 * Liest einen Archivteil stueckweise. Der Rueckruf bekommt den bisher
 * gesammelten Text und gibt ueber `keepFrom` zurueck, was davon bereits
 * verarbeitet wurde - so bleibt der Puffer klein, auch wenn der Teil
 * mehrere hundert Megabyte gross ist.
 */
function streamPart(zip, path, onChunk) {
  const entry = zip.file(path);
  if (!entry) return Promise.resolve();

  return new Promise((resolve, reject) => {
    const stream = entry.nodeStream('nodebuffer');
    let buffer = '';
    let done = false;

    const finish = () => {
      if (done) return;
      done = true;
      stream.destroy();
      resolve();
    };

    const handle = () => {
      const chunk = {
        text: buffer,
        keepFrom(position) { buffer = position > 0 ? buffer.slice(position) : buffer; },
      };
      onChunk(chunk, finish);
    };

    stream.on('data', (data) => {
      if (done) return;
      buffer += data.toString('utf8');
      handle();
    });
    stream.on('end', () => {
      if (!done) handle();
      finish();
    });
    stream.on('error', (error) => {
      if (done) return;
      done = true;
      reject(error);
    });
  });
}
