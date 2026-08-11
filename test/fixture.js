/** Erzeugt eine Beispiel-Arbeitsmappe fuer die Tests. */

import { readFile, writeFile, rm } from 'node:fs/promises';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';

export const HEADERS = [
  'KundenID', 'Vorname', 'Nachname', 'E-Mail', 'Telefon',
  'Strasse', 'PLZ', 'Ort', 'Geburtsdatum', 'Umsatz', 'Aktiv', 'Bemerkung', 'Summe',
];

export const ROWS = [
  [1001, 'Anna', 'Müller', 'anna.mueller@firma.de', '+49 30 1234567',
    'Hauptstr. 12', '10115', 'Berlin', new Date('1985-03-14'), 1499.99, true, 'Stammkunde seit 2010'],
  [1002, 'Bernd', 'Schmidt', 'b.schmidt@firma.de', '089 9876543',
    'Ringweg 4', '80331', 'München', new Date('1972-11-02'), 240.5, false, 'Zahlt per Rechnung'],
  [1003, 'Clara', 'Weber', 'clara.weber@firma.de', '+49 221 555000',
    'Domplatz 1', '50667', 'Köln', new Date('1990-07-21'), 8750, true, 'Rabattstaffel B'],
  [1004, 'David', 'Fischer', 'd.fischer@firma.de', '0711 445566',
    'Marktgasse 9', '70173', 'Stuttgart', new Date('1968-01-30'), 55.25, true, ''],
  [1005, 'Anna', 'Müller', 'anna.mueller@firma.de', '+49 30 1234567',
    'Hauptstr. 12', '10115', 'Berlin', new Date('1985-03-14'), 310, false, null],
];

/**
 * Schreibt eine Testdatei mit gemischten Datentypen, einer Formelspalte und
 * einem zweiten Arbeitsblatt.
 */
export async function createFixture(file) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Kunden');

  sheet.addRow(HEADERS);
  ROWS.forEach((row, index) => {
    const added = sheet.addRow(row);
    // Spalte "Summe" ist berechnet und muss unangetastet bleiben.
    added.getCell(13).value = { formula: `J${index + 2}*1.19`, result: row[9] * 1.19 };
    added.getCell(9).numFmt = 'dd.mm.yyyy';
    added.getCell(10).numFmt = '#,##0.00 "EUR"';
  });

  const second = workbook.addWorksheet('Bestellungen');
  second.addRow(['BestellID', 'KundenID', 'Betrag']);
  second.addRow([1, 1001, 99.9]);
  second.addRow([2, 1002, 12.5]);

  await workbook.xlsx.writeFile(file);
  return file;
}

export const VBA_CONTENT = 'FAKE-VBA-PROJECT-BINARY';
export const VBA_SIGNATURE_CONTENT = 'FAKE-VBA-SIGNATURE';
export const WORKBOOK_CODE_NAME = 'DiesesWorkbook';
export const SHEET_CODE_NAME = 'TabelleKunden';

/**
 * Baut aus der Beispieldatei eine Makro-Arbeitsmappe, wie Excel sie schreibt:
 * mit VBA-Projekt, makrofaehigem Arbeitsmappentyp, Verweis in den Beziehungen
 * und Codenamen fuer Arbeitsmappe und Blatt. Zusaetzlich ein ActiveX-Teil, den
 * exceljs nicht kennt - damit laesst sich die Verlustmeldung pruefen.
 */
export async function createMacroFixture(file) {
  const temporary = `${file}.tmp.xlsx`;
  await createFixture(temporary);

  const zip = await JSZip.loadAsync(await readFile(temporary));

  zip.file('xl/vbaProject.bin', VBA_CONTENT);
  zip.file('xl/vbaProjectSignature.bin', VBA_SIGNATURE_CONTENT);
  zip.file('xl/activeX/activeX1.xml', '<ax:ocx xmlns:ax="http://example.invalid/ax"/>');

  const types = await zip.file('[Content_Types].xml').async('string');
  zip.file('[Content_Types].xml', types
    .replace(
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
      'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
    )
    .replace('</Types>',
      '<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>'
      + '</Types>'));

  const rels = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  zip.file('xl/_rels/workbook.xml.rels', rels.replace('</Relationships>',
    '<Relationship Id="rId999" '
    + 'Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" '
    + 'Target="vbaProject.bin"/></Relationships>'));

  const workbook = await zip.file('xl/workbook.xml').async('string');
  zip.file('xl/workbook.xml', workbook.replace(
    /(<workbook\b[^>]*>)/,
    `$1<workbookPr codeName="${WORKBOOK_CODE_NAME}"/>`,
  ));

  const sheet = await zip.file('xl/worksheets/sheet1.xml').async('string');
  zip.file('xl/worksheets/sheet1.xml', sheet.replace(
    /(<worksheet\b[^>]*>)/,
    `$1<sheetPr codeName="${SHEET_CODE_NAME}"/>`,
  ));

  await writeFile(file, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
  await rm(temporary, { force: true });
  return file;
}

/** Liest die fuer Makros relevanten Teile einer erzeugten Datei aus. */
export async function readMacroState(file) {
  const zip = await JSZip.loadAsync(await readFile(file));
  const text = async (path) => (zip.file(path) ? zip.file(path).async('string') : null);

  return {
    parts: Object.keys(zip.files).filter((name) => !zip.files[name].dir),
    vba: await text('xl/vbaProject.bin'),
    signature: await text('xl/vbaProjectSignature.bin'),
    contentTypes: await text('[Content_Types].xml'),
    workbookRels: await text('xl/_rels/workbook.xml.rels'),
    workbook: await text('xl/workbook.xml'),
    sheet1: await text('xl/worksheets/sheet1.xml'),
  };
}

/** Liest ein Blatt als Array von Objekten zurueck. */
export async function readSheet(file, sheetName = 'Kunden') {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(file);
  const sheet = workbook.getWorksheet(sheetName);

  const headers = [];
  sheet.getRow(1).eachCell((cell, index) => { headers[index] = String(cell.value); });

  const rows = [];
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const record = {};
    headers.forEach((header, index) => {
      if (header) record[header] = row.getCell(index).value;
    });
    record.__row = row;
    rows.push(record);
  });

  return rows;
}
