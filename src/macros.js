/**
 * Erhaelt das VBA-Projekt von Makro-Arbeitsmappen (.xlsm).
 *
 * Hintergrund: exceljs kennt keine Makros. Liest man eine .xlsm-Datei ein und
 * schreibt sie zurueck, fehlt in der Ausgabe `xl/vbaProject.bin` - die Makros
 * sind ersatzlos verschwunden.
 *
 * Dieses Modul sichert die makro-relevanten Teile vor dem Schreiben aus der
 * Originaldatei und fuegt sie danach wieder in die erzeugte Datei ein. Eine
 * .xlsx/.xlsm-Datei ist ein ZIP-Archiv aus XML-Teilen, deshalb genuegt es,
 * die fehlenden Teile zu ergaenzen und die Verweise darauf zu reparieren:
 *
 *   1. `xl/vbaProject.bin`            - das Projekt selbst
 *   2. `[Content_Types].xml`          - Typ des Projekts + makrofaehiger Arbeitsmappentyp
 *   3. `xl/_rels/workbook.xml.rels`   - Verweis der Arbeitsmappe auf das Projekt
 *   4. `codeName`-Attribute           - damit VBA seine Blaetter wiederfindet
 */

import { readFile, writeFile } from 'node:fs/promises';
import JSZip from 'jszip';

const VBA_PART = 'xl/vbaProject.bin';
const VBA_SIGNATURE_PART = 'xl/vbaProjectSignature.bin';
const VBA_CONTENT_TYPE = 'application/vnd.ms-office.vbaProject';
const VBA_RELATIONSHIP = 'http://schemas.microsoft.com/office/2006/relationships/vbaProject';
const WORKBOOK_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml';
const MACRO_WORKBOOK_TYPE = 'application/vnd.ms-excel.sheet.macroEnabled.main+xml';

/**
 * Teile, deren Verlust unkritisch ist: Excel berechnet sie neu oder sie haben
 * keine inhaltliche Bedeutung. Alles andere wird gemeldet, damit ein Verlust
 * nie unbemerkt bleibt.
 */
const HARMLESS_LOSSES = [
  /^xl\/calcChain\.xml$/,
  /^docProps\/thumbnail\./,
  /^xl\/printerSettings\//,
  /\.rels$/,
  /^\[Content_Types\]\.xml$/,
  /\/$/,
];

/**
 * Liest die makro-relevanten Teile einer Arbeitsmappe.
 *
 * @param {string} file Pfad zur Quelldatei.
 * @returns {Promise<object|null>} `null`, wenn die Datei kein VBA-Projekt enthaelt.
 */
export async function readMacroParts(file) {
  const zip = await JSZip.loadAsync(await readFile(file));
  const parts = Object.keys(zip.files).filter((name) => !zip.files[name].dir);

  const vbaEntry = zip.file(VBA_PART);
  if (!vbaEntry) return null;

  const workbookXml = await readText(zip, 'xl/workbook.xml');

  return {
    vba: await vbaEntry.async('nodebuffer'),
    signature: zip.file(VBA_SIGNATURE_PART)
      ? await zip.file(VBA_SIGNATURE_PART).async('nodebuffer')
      : null,
    workbookCodeName: readAttribute(workbookXml, 'workbookPr', 'codeName'),
    sheetCodeNames: await readSheetCodeNames(zip, workbookXml),
    originalParts: parts,
  };
}

/**
 * Fuegt die gesicherten Teile wieder in die geschriebene Datei ein.
 *
 * @param {string} file  Pfad zur erzeugten Datei (wird ueberschrieben).
 * @param {object} macro Rueckgabewert von readMacroParts().
 * @returns {Promise<string[]>} Teile der Originaldatei, die trotzdem fehlen.
 */
export async function restoreMacroParts(file, macro) {
  const zip = await JSZip.loadAsync(await readFile(file));

  zip.file(VBA_PART, macro.vba);
  if (macro.signature) zip.file(VBA_SIGNATURE_PART, macro.signature);

  zip.file('[Content_Types].xml', patchContentTypes(
    await readText(zip, '[Content_Types].xml'),
    Boolean(macro.signature),
  ));

  zip.file('xl/_rels/workbook.xml.rels', patchWorkbookRels(
    await readText(zip, 'xl/_rels/workbook.xml.rels'),
  ));

  const workbookXml = await readText(zip, 'xl/workbook.xml');
  if (macro.workbookCodeName) {
    zip.file('xl/workbook.xml', setAttribute(
      workbookXml, 'workbookPr', 'codeName', macro.workbookCodeName,
    ));
  }

  await restoreSheetCodeNames(zip, workbookXml, macro.sheetCodeNames);

  const rebuilt = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    // Excel erwartet den Typ-Katalog als ersten Eintrag im Archiv.
    mimeType: undefined,
  });
  await writeFile(file, rebuilt);

  const present = new Set(Object.keys(zip.files));
  return macro.originalParts.filter(
    (name) => !present.has(name) && !HARMLESS_LOSSES.some((re) => re.test(name)),
  );
}

/** Ordnet jedem Arbeitsblatt seinen VBA-Codenamen zu (Blattname -> codeName). */
async function readSheetCodeNames(zip, workbookXml) {
  const codeNames = new Map();
  const rels = await readText(zip, 'xl/_rels/workbook.xml.rels');

  for (const sheet of listSheets(workbookXml)) {
    const target = resolveRelationship(rels, sheet.rid);
    if (!target) continue;

    const sheetXml = await readText(zip, `xl/${target.replace(/^\/?xl\//, '')}`);
    const codeName = readAttribute(sheetXml, 'sheetPr', 'codeName');
    if (codeName) codeNames.set(sheet.name, codeName);
  }

  return codeNames;
}

/**
 * Schreibt die Codenamen in die neue Datei zurueck. Die Zuordnung erfolgt ueber
 * den Blattnamen - Bloecke, deren Name sich nicht wiederfindet, werden
 * ausgelassen, statt einen falschen Codenamen zu setzen.
 */
async function restoreSheetCodeNames(zip, workbookXml, codeNames) {
  if (!codeNames.size) return;

  const rels = await readText(zip, 'xl/_rels/workbook.xml.rels');

  for (const sheet of listSheets(workbookXml)) {
    const codeName = codeNames.get(sheet.name);
    if (!codeName) continue;

    const target = resolveRelationship(rels, sheet.rid);
    if (!target) continue;

    const path = `xl/${target.replace(/^\/?xl\//, '')}`;
    const sheetXml = await readText(zip, path);
    if (!sheetXml) continue;

    zip.file(path, setSheetCodeName(sheetXml, codeName));
  }
}

/** Ergaenzt Typ des VBA-Projekts und stellt den Arbeitsmappentyp auf makrofaehig. */
function patchContentTypes(xml, hasSignature) {
  let patched = xml.replace(WORKBOOK_TYPE, MACRO_WORKBOOK_TYPE);

  // Override statt Default: ein Default fuer "bin" wuerde mit den
  // Druckereinstellungen kollidieren, die dieselbe Endung verwenden.
  const overrides = [`<Override PartName="/${VBA_PART}" ContentType="${VBA_CONTENT_TYPE}"/>`];
  if (hasSignature) {
    overrides.push(
      `<Override PartName="/${VBA_SIGNATURE_PART}" ` +
      'ContentType="application/vnd.ms-office.vbaProjectSignature"/>',
    );
  }

  for (const override of overrides) {
    if (patched.includes(override)) continue;
    patched = patched.replace('</Types>', `${override}</Types>`);
  }

  return patched;
}

/** Traegt den Verweis der Arbeitsmappe auf das VBA-Projekt ein. */
function patchWorkbookRels(xml) {
  if (xml.includes(VBA_RELATIONSHIP)) return xml;

  const used = new Set([...xml.matchAll(/Id="(rId\d+)"/g)].map((match) => match[1]));
  let next = 1;
  while (used.has(`rId${next}`)) next += 1;

  const relationship =
    `<Relationship Id="rId${next}" Type="${VBA_RELATIONSHIP}" Target="vbaProject.bin"/>`;
  return xml.replace('</Relationships>', `${relationship}</Relationships>`);
}

/** Liefert Blattname und Beziehungs-Id aller Arbeitsblaetter in Dokumentreihenfolge. */
function listSheets(workbookXml) {
  if (!workbookXml) return [];

  return [...workbookXml.matchAll(/<sheet\b[^>]*\/?>/g)]
    .map((match) => {
      const tag = match[0];
      const name = /\bname="([^"]*)"/.exec(tag);
      const rid = /\br:id="([^"]*)"/.exec(tag);
      return name && rid ? { name: decodeXml(name[1]), rid: rid[1] } : null;
    })
    .filter(Boolean);
}

function resolveRelationship(relsXml, rid) {
  if (!relsXml) return null;
  const pattern = new RegExp(`<Relationship\\b[^>]*Id="${rid}"[^>]*>`);
  const tag = pattern.exec(relsXml);
  if (!tag) return null;

  const target = /\bTarget="([^"]*)"/.exec(tag[0]);
  return target ? target[1] : null;
}

/**
 * `sheetPr` muss das erste Kindelement von `worksheet` sein - deshalb wird ein
 * fehlendes Element direkt hinter dem oeffnenden Tag eingefuegt.
 */
function setSheetCodeName(xml, codeName) {
  if (/<sheetPr\b/.test(xml)) return setAttribute(xml, 'sheetPr', 'codeName', codeName);

  return xml.replace(/(<worksheet\b[^>]*>)/, `$1<sheetPr codeName="${escapeXml(codeName)}"/>`);
}

function readAttribute(xml, tagName, attribute) {
  if (!xml) return null;
  const tag = new RegExp(`<${tagName}\\b[^>]*>`).exec(xml);
  if (!tag) return null;

  const value = new RegExp(`\\b${attribute}="([^"]*)"`).exec(tag[0]);
  return value ? decodeXml(value[1]) : null;
}

/** Setzt ein Attribut; legt das Element an, falls es noch nicht existiert. */
function setAttribute(xml, tagName, attribute, value) {
  const escaped = escapeXml(value);
  const tag = new RegExp(`<${tagName}\\b[^>]*?(/?)>`).exec(xml);

  if (!tag) {
    // workbookPr gehoert direkt hinter das oeffnende workbook-Element.
    return xml.replace(/(<workbook\b[^>]*>)/, `$1<${tagName} ${attribute}="${escaped}"/>`);
  }

  const existing = new RegExp(`\\b${attribute}="[^"]*"`);
  const replacement = existing.test(tag[0])
    ? tag[0].replace(existing, `${attribute}="${escaped}"`)
    : tag[0].replace(/(\/?)>$/, ` ${attribute}="${escaped}"$1>`);

  return xml.replace(tag[0], replacement);
}

async function readText(zip, path) {
  const entry = zip.file(path);
  return entry ? entry.async('string') : null;
}

function escapeXml(value) {
  return String(value)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function decodeXml(value) {
  return String(value)
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
}
