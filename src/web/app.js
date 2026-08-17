/**
 * Browser-Fassung des Maskierers.
 *
 * Warum es diese Fassung gibt: Auf verwalteten Firmenrechnern ist jede
 * Programmdatei freigabepflichtig - eine mitgelieferte node.exe genauso wie
 * frueher ein PowerShell-Skript. Der Browser ist dagegen bereits freigegeben.
 * Diese Anwendung ist eine einzelne HTML-Datei: kein Programm, keine
 * Installation, kein Interpreter.
 *
 * Die Datei wird ausschliesslich im Browser verarbeitet. Es gibt in dieser
 * Anwendung keinen Netzwerkzugriff - die Daten verlassen den Rechner nicht.
 *
 * Fachlich laeuft alles ueber dieselben Bausteine wie die Kommandozeile:
 * core/classify, core/generators, core/keep, excel/mask und access/plan.
 */

import ExcelJS from 'exceljs';
import MDBReader from 'mdb-reader';

import { maskWorkbook } from '../excel/mask.js';
import { describeDatabase } from '../access/describe.js';
import { planAnonymization } from '../access/plan.js';
import { fillWorkbook } from '../access/export.js';

const VERSION = '1.0.0';

const state = {
  file: null,
  kind: null,     // 'excel' | 'access'
  buffer: null,
  sections: [],   // { name, columns: [{ header, kind, fixed, fixedReason }] }
};

const el = (id) => document.getElementById(id);

/** Startet die Anwendung, sobald die Seite steht. */
export function start() {
  // Belegt, dass der Code geladen wurde - die HTML-Datei prueft das.
  window.__maskierer = VERSION;
  el('version').textContent = `Fassung ${VERSION}`;

  const input = el('datei');
  input.addEventListener('change', () => {
    if (input.files.length) void loadFile(input.files[0]);
  });

  const zone = el('ablage');
  zone.addEventListener('dragover', (event) => {
    event.preventDefault();
    zone.classList.add('aktiv');
  });
  zone.addEventListener('dragleave', () => zone.classList.remove('aktiv'));
  zone.addEventListener('drop', (event) => {
    event.preventDefault();
    zone.classList.remove('aktiv');
    if (event.dataTransfer.files.length) void loadFile(event.dataTransfer.files[0]);
  });

  el('starten').addEventListener('click', () => void maskiere());
}

async function loadFile(file) {
  reset();
  const extension = file.name.toLowerCase().split('.').pop();
  const kind = ['xlsx', 'xlsm'].includes(extension) ? 'excel'
    : ['accdb', 'mdb'].includes(extension) ? 'access' : null;

  if (!kind) {
    zeigeFehler(`Nicht unterstuetzte Dateiendung ".${extension}". `
      + 'Erwartet werden .xlsx, .xlsm, .accdb oder .mdb.');
    return;
  }

  melde(`"${file.name}" wird gelesen …`);

  try {
    state.file = file;
    state.kind = kind;
    state.buffer = await file.arrayBuffer();
    state.sections = kind === 'excel'
      ? await leseArbeitsmappe(state.buffer)
      : leseDatenbank(state.buffer);

    zeigeSpalten();
    melde(`"${file.name}" gelesen. Bitte pruefen, welche Spalten unveraendert bleiben sollen.`);
  } catch (error) {
    zeigeFehler(error.message);
  }
}

/** Liest Blattnamen und Spalten einer Arbeitsmappe. */
async function leseArbeitsmappe(buffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  // Die Spaltenerkennung steckt in maskWorkbook; ein Probelauf auf einer
  // Kopie liefert dieselbe Einschaetzung, ohne etwas zu behalten.
  const { report } = maskWorkbook(workbook, { seed: 0 });

  return report.sheets.map((sheet) => ({
    name: sheet.name,
    skipped: sheet.skipped,
    columns: sheet.columns.map((column) => ({
      header: column.header,
      kind: column.kind,
      fixed: column.status.startsWith('uebersprungen'),
      fixedReason: column.status,
    })),
  }));
}

/** Liest Tabellen und Spalten einer Access-Datenbank. */
function leseDatenbank(buffer) {
  const tables = describeDatabase(new Uint8Array(buffer), { rowLimit: 200 });
  const { report } = planAnonymization(tables, {});

  return report.sheets.map((sheet) => ({
    name: sheet.name,
    skipped: sheet.skipped,
    columns: sheet.columns.map((column) => ({
      header: column.header,
      kind: column.kind,
      fixed: !column.status.startsWith('anonymisiert'),
      fixedReason: column.status,
    })),
  }));
}

/** Baut die Spaltenliste mit Ankreuzfeldern auf. */
function zeigeSpalten() {
  const ziel = el('spalten');
  ziel.innerHTML = '';

  const bezeichnung = state.kind === 'access' ? 'Tabelle' : 'Arbeitsblatt';

  for (const section of state.sections) {
    const block = document.createElement('section');
    const titel = document.createElement('h3');
    titel.textContent = `${bezeichnung}: ${section.name}`;
    block.append(titel);

    if (section.skipped) {
      const hinweis = document.createElement('p');
      hinweis.className = 'still';
      hinweis.textContent = `uebersprungen - ${section.skipped}`;
      block.append(hinweis);
      ziel.append(block);
      continue;
    }

    for (const column of section.columns) {
      const zeile = document.createElement('label');
      zeile.className = 'spalte';

      const box = document.createElement('input');
      box.type = 'checkbox';
      box.dataset.sheet = section.name;
      box.dataset.column = column.header;
      box.disabled = column.fixed;
      box.checked = column.fixed;

      const name = document.createElement('span');
      name.className = 'name';
      name.textContent = column.header;

      const art = document.createElement('span');
      art.className = 'art';
      art.textContent = column.fixed ? column.fixedReason : column.kind;

      zeile.append(box, name, art);
      block.append(zeile);
    }

    ziel.append(block);
  }

  el('auswahl').hidden = false;
}

/** Sammelt die angekreuzten Spalten als --keep-Angaben. */
function sammleKeep() {
  return [...document.querySelectorAll('#spalten input[type=checkbox]')]
    .filter((box) => box.checked && !box.disabled)
    .map((box) => `${box.dataset.sheet}:${box.dataset.column}`);
}

async function maskiere() {
  const keep = sammleKeep();
  melde('Wird maskiert …');
  el('starten').disabled = true;

  try {
    const { blob, name, report } = state.kind === 'excel'
      ? await maskiereArbeitsmappe(keep)
      : await maskiereDatenbank(keep);

    zeigeBericht(report, keep);
    speichere(blob, name);
    melde(`Fertig. "${name}" wurde heruntergeladen.`);
  } catch (error) {
    zeigeFehler(error.message);
  } finally {
    el('starten').disabled = false;
  }
}

async function maskiereArbeitsmappe(keep) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(state.buffer);

  const { report } = maskWorkbook(workbook, { keep });
  const buffer = await workbook.xlsx.writeBuffer();

  return {
    blob: new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    }),
    name: zielName(state.file.name, extensionOf(state.file.name)),
    report,
  };
}

async function maskiereDatenbank(keep) {
  const tables = describeDatabase(new Uint8Array(state.buffer));
  const { report, result } = planAnonymization(tables, { keep });

  const workbook = new ExcelJS.Workbook();
  fillWorkbook(workbook, result);
  const buffer = await workbook.xlsx.writeBuffer();

  return {
    blob: new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    }),
    name: zielName(state.file.name, 'xlsx'),
    report,
  };
}

function extensionOf(name) {
  return name.toLowerCase().split('.').pop();
}

function zielName(name, extension) {
  const ohne = name.replace(/\.[^.]+$/, '');
  return `${ohne}.anonymisiert.${extension}`;
}

/** Loest den Download aus. Die Datei bleibt dabei auf diesem Rechner. */
function speichere(blob, name) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = name;
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 10000);
}

/** Zeigt den Bericht und nennt ausdruecklich die unveraenderten Spalten. */
function zeigeBericht(report, keep) {
  const ziel = el('bericht');
  ziel.innerHTML = '';

  const kopf = document.createElement('p');
  kopf.innerHTML = `<strong>${report.changed}</strong> Zellen maskiert.`;
  ziel.append(kopf);

  const unveraendert = report.sheets
    .filter((sheet) => !sheet.skipped)
    .map((sheet) => ({
      name: sheet.name,
      columns: sheet.columns
        .filter((column) => column.status.startsWith('unveraendert'))
        .map((column) => column.header),
    }))
    .filter((sheet) => sheet.columns.length);

  if (unveraendert.length) {
    const warnung = document.createElement('div');
    warnung.className = 'warnung';
    warnung.innerHTML = '<strong>Diese Spalten enthalten weiterhin die '
      + 'Originaldaten:</strong>';

    const liste = document.createElement('ul');
    for (const sheet of unveraendert) {
      const punkt = document.createElement('li');
      punkt.textContent = `${sheet.name}: ${sheet.columns.join(', ')}`;
      liste.append(punkt);
    }
    warnung.append(liste);

    const pruefen = document.createElement('p');
    pruefen.textContent = 'Bitte pruefen, dass darin keine personenbezogenen '
      + 'Angaben stehen.';
    warnung.append(pruefen);
    ziel.append(warnung);
  }

  ziel.hidden = false;
  void keep;
}

function melde(text) {
  el('meldung').textContent = text;
  el('meldung').className = 'meldung';
}

function zeigeFehler(text) {
  el('meldung').textContent = `Fehler: ${text}`;
  el('meldung').className = 'meldung fehler';
}

function reset() {
  el('auswahl').hidden = true;
  el('bericht').hidden = true;
  el('spalten').innerHTML = '';
}

if (typeof document !== 'undefined') {
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', start);
  } else {
    start();
  }
}
