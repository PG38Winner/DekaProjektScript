/** Erzeugt eine Beispiel-Arbeitsmappe fuer die Tests. */

import ExcelJS from 'exceljs';

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
