/**
 * Erzeugt die Ersatzwerte. Die Ersetzung ist typ- und groessenordnungserhaltend:
 * aus einer Zahl wird eine Zahl aehnlicher Groesse, aus einem Datum ein Datum in
 * der Naehe, aus einem Text ein Text aehnlicher Laenge. So bleiben Spaltentypen,
 * Zellformate und Auswertungen der Datei benutzbar.
 */

import { fakerDE as faker } from '@faker-js/faker';
import { KINDS } from './classify.js';

const MS_PER_DAY = 24 * 60 * 60 * 1000;

/**
 * @param {object} [options]
 * @param {number} [options.seed]        Fester Startwert - gleicher Seed erzeugt
 *                                       reproduzierbar dieselben Ersatzwerte.
 * @param {boolean} [options.consistent] Gleicher Eingabewert erhaelt innerhalb
 *                                       einer Spalte denselben Ersatzwert
 *                                       (Standard: true). Damit bleiben
 *                                       Gruppierungen und Verknuepfungen ueber
 *                                       die Spalte hinweg erhalten.
 */
export function createGenerator({ seed, consistent = true } = {}) {
  if (seed !== undefined) faker.seed(seed);

  /** @type {Map<string, Map<string, unknown>>} Spalte -> Originalwert -> Ersatz */
  const memo = new Map();

  return function generate(kind, original, columnKey) {
    if (!consistent) return build(kind, original);

    let column = memo.get(columnKey);
    if (!column) {
      column = new Map();
      memo.set(columnKey, column);
    }

    // Date-Objekte und Zahlen brauchen einen stabilen Map-Schluessel.
    const key = original instanceof Date ? `d:${original.getTime()}` : `v:${String(original)}`;
    if (column.has(key)) return column.get(key);

    const replacement = build(kind, original);
    column.set(key, replacement);
    return replacement;
  };
}

function build(kind, original) {
  switch (kind) {
    case KINDS.EMAIL:
      return buildEmail();
    case KINDS.PHONE:
      return `+49 ${faker.number.int({ min: 30, max: 999 })} ${faker.string.numeric({ length: { min: 6, max: 8 } })}`;
    case KINDS.FIRST_NAME:
      return faker.person.firstName();
    case KINDS.LAST_NAME:
      return faker.person.lastName();
    case KINDS.FULL_NAME:
      return faker.person.fullName();
    case KINDS.STREET:
      return faker.location.streetAddress();
    case KINDS.CITY:
      return faker.location.city();
    case KINDS.COUNTRY:
      return faker.location.country();
    case KINDS.COMPANY:
      return faker.company.name();
    case KINDS.IBAN:
      return faker.finance.iban({ countryCode: 'DE' });
    case KINDS.ZIP:
      return buildZip(original);
    case KINDS.INTEGER:
      return buildInteger(original);
    case KINDS.DECIMAL:
      return buildDecimal(original);
    case KINDS.DATE:
      return buildDate(original);
    case KINDS.BOOLEAN:
      return faker.datatype.boolean();
    case KINDS.TEXT:
    default:
      return buildText(original);
  }
}

function buildEmail() {
  const first = faker.person.firstName();
  const last = faker.person.lastName();
  const normalize = (s) => s
    .toLowerCase()
    .replace(/ä/g, 'ae').replace(/ö/g, 'oe').replace(/ü/g, 'ue').replace(/ß/g, 'ss')
    .replace(/[^a-z0-9]/g, '');
  return `${normalize(first)}.${normalize(last)}@example.com`;
}

/** Behaelt die Stellenzahl bei, damit fuehrende Nullen und Textformate passen. */
function buildZip(original) {
  const digits = String(original ?? '').replace(/\D/g, '').length || 5;
  const zip = faker.string.numeric({ length: digits, allowLeadingZeros: false });
  return typeof original === 'number' ? Number(zip) : zip;
}

/** Zufallszahl in derselben Groessenordnung wie das Original. */
function buildInteger(original) {
  const value = Number(original);
  if (!Number.isFinite(value)) return faker.number.int({ min: 1, max: 999 });

  const magnitude = Math.max(Math.abs(value), 1);
  const min = Math.floor(magnitude / 2);
  const max = Math.ceil(magnitude * 2);
  const replacement = faker.number.int({ min, max });
  return value < 0 ? -replacement : replacement;
}

/** Wie buildInteger, behaelt aber die Anzahl der Nachkommastellen bei. */
function buildDecimal(original) {
  const value = Number(original);
  if (!Number.isFinite(value)) return faker.number.float({ min: 1, max: 999, fractionDigits: 2 });

  const decimals = decimalPlaces(value);
  const magnitude = Math.max(Math.abs(value), 1);
  const replacement = faker.number.float({
    min: magnitude / 2,
    max: magnitude * 2,
    fractionDigits: decimals,
  });
  return value < 0 ? -replacement : replacement;
}

function decimalPlaces(value) {
  const text = String(value);
  const dot = text.indexOf('.');
  if (dot === -1) return 0;
  return Math.min(text.length - dot - 1, 6);
}

/** Verschiebt das Datum um bis zu ~1 Jahr, behaelt die Uhrzeit bei. */
function buildDate(original) {
  const date = original instanceof Date ? original : new Date(original);
  if (Number.isNaN(date.getTime())) return faker.date.past();

  const shift = faker.number.int({ min: -365, max: 365 });
  return new Date(date.getTime() + shift * MS_PER_DAY);
}

/**
 * Erfundener Fuelltext. Die Laenge orientiert sich am Original und wird darauf
 * begrenzt, damit Spaltenbreiten und Feldlaengen nicht gesprengt werden.
 */
function buildText(original) {
  const target = String(original ?? '').length;
  if (target === 0) return '';
  if (target <= 3) return faker.string.alpha({ length: target });

  let text = '';
  while (text.length < target) {
    text += (text ? ' ' : '') + faker.word.sample();
  }
  return text.slice(0, target).trim();
}
