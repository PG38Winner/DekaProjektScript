/**
 * Gemeinsame Helfer fuer den direkten Zugriff auf den Aufbau einer
 * Excel-Datei. Eine .xlsx/.xlsm-Datei ist ein ZIP-Archiv aus XML-Teilen;
 * einige Angaben (Blattnamen, Zeilenzahl) stehen dort weit vorne und lassen
 * sich lesen, ohne die ganze Datei auszuwerten.
 */

import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';

export async function openWorkbookZip(file) {
  return JSZip.loadAsync(await readFile(file));
}

export async function readText(zip, path) {
  const entry = zip.file(path);
  return entry ? entry.async('string') : null;
}

/**
 * Liest nur den Anfang eines Teils. Fuer Angaben im Kopf eines Blattes ist das
 * um ein Vielfaches schneller, als die oft zweistellig megabytegrosse XML-Datei
 * vollstaendig zu entpacken.
 *
 * @param {import('jszip')} zip
 * @param {string} path      Pfad im Archiv.
 * @param {RegExp} pattern   Muster, bei dessen Fund abgebrochen wird.
 * @param {number} [maxBytes] Obergrenze, falls das Muster nicht auftaucht.
 * @returns {Promise<RegExpExecArray|null>}
 */
export function findInPartHead(zip, path, pattern, maxBytes = 65536) {
  const entry = zip.file(path);
  if (!entry) return Promise.resolve(null);

  return new Promise((resolve, reject) => {
    let buffer = '';
    let settled = false;
    const stream = entry.nodeStream('nodebuffer');

    const finish = (result) => {
      if (settled) return;
      settled = true;
      stream.destroy();
      resolve(result);
    };

    stream.on('data', (chunk) => {
      buffer += chunk.toString('utf8');
      const match = pattern.exec(buffer);
      if (match) finish(match);
      else if (buffer.length > maxBytes) finish(null);
    });
    stream.on('end', () => finish(null));
    stream.on('error', (error) => {
      if (settled) return;
      settled = true;
      reject(error);
    });
  });
}

/** Blattname und Beziehungs-Id aller Arbeitsblaetter in Dokumentreihenfolge. */
export function listSheets(workbookXml) {
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

export function resolveRelationship(relsXml, rid) {
  if (!relsXml) return null;
  const pattern = new RegExp(`<Relationship\\b[^>]*Id="${rid}"[^>]*>`);
  const tag = pattern.exec(relsXml);
  if (!tag) return null;

  const target = /\bTarget="([^"]*)"/.exec(tag[0]);
  return target ? target[1] : null;
}

/** Wandelt ein Beziehungsziel in einen Pfad innerhalb des Archivs. */
export function sheetPartPath(target) {
  return `xl/${String(target).replace(/^\/?xl\//, '')}`;
}

export function decodeXml(value) {
  return String(value)
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
}

export function escapeXml(value) {
  return String(value)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
