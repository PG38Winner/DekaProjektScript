/**
 * Erkennt, welche Art von Inhalt in einer Spalte steht, damit die Ersetzung
 * typ- und inhaltsabhaengig erfolgen kann (siehe Tabelle in der README).
 *
 * Die Erkennung laeuft in zwei Stufen:
 *   1. Kopfzeile  - der Spaltenname ist das staerkste Signal ("E-Mail", "PLZ").
 *   2. Inhalt     - erkennt Muster in den Werten (Mail-Adressen, Telefonnummern)
 *                   und faellt sonst auf den Datentyp zurueck.
 */

export const KINDS = {
  EMAIL: 'email',
  PHONE: 'phone',
  FIRST_NAME: 'firstName',
  LAST_NAME: 'lastName',
  FULL_NAME: 'fullName',
  STREET: 'street',
  CITY: 'city',
  ZIP: 'zip',
  COUNTRY: 'country',
  COMPANY: 'company',
  IBAN: 'iban',
  TEXT: 'text',
  INTEGER: 'integer',
  DECIMAL: 'decimal',
  DATE: 'date',
  BOOLEAN: 'boolean',
  EMPTY: 'empty',
};

/**
 * Kopfzeilen-Muster. Reihenfolge ist bedeutsam: das erste Muster, das passt,
 * gewinnt. Spezifische Muster stehen deshalb vor allgemeinen - "nachname" muss
 * vor dem blossen "name" geprueft werden.
 */
const HEADER_PATTERNS = [
  [KINDS.EMAIL, /(e[-_ ]?mail|mail[-_ ]?adresse)/i],
  [KINDS.PHONE, /(telefon|telefonnr|tel\b|handy|mobil|fax|phone)/i],
  [KINDS.ZIP, /(plz|postleitzahl|zip|postal)/i],
  [KINDS.FIRST_NAME, /(vorname|rufname|first[-_ ]?name)/i],
  [KINDS.LAST_NAME, /(nachname|familienname|zuname|nachnahme|last[-_ ]?name|surname)/i],
  [KINDS.COMPANY, /(firma|unternehmen|company|arbeitgeber)/i],
  [KINDS.STREET, /(strasse|straße|str\.|adresse|anschrift|street|address)/i],
  [KINDS.CITY, /(\bort\b|wohnort|stadt|city|gemeinde)/i],
  [KINDS.COUNTRY, /(land|country|staat)/i],
  [KINDS.IBAN, /(iban|kontonummer|bankverbindung)/i],
  // Bewusst am Ende: faengt "Name", "Kundenname", "Mitarbeitername" ab, nachdem
  // Vor-/Nachname bereits ihre Chance hatten.
  [KINDS.FULL_NAME, /name/i],
];

const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[a-z]{2,}$/i;
// Telefonnummern: mind. 6 Ziffern, dazu die uebliche Interpunktion.
const PHONE_RE = /^[+(]?[\d][\d\s()/.-]{5,}$/;
const ZIP_RE = /^\d{4,5}$/;
const IBAN_RE = /^[A-Z]{2}\d{2}[A-Z0-9\s]{10,30}$/i;

/**
 * Bestimmt die Art einer Spalte aus Kopfzeile und Beispielwerten.
 *
 * @param {string} header      Spaltenname aus der ersten Zeile.
 * @param {unknown[]} samples  Werte der Spalte (NULL/leer bereits gefiltert).
 * @returns {string} eine der KINDS-Konstanten.
 */
export function classifyColumn(header, samples) {
  const byHeader = classifyByHeader(header);
  const byValue = classifyByValue(samples);

  if (!byHeader) return byValue;

  // Der Inhalt schlaegt die Kopfzeile, wenn er eindeutig ein anderes Format
  // zeigt - eine Spalte "Kontakt" mit lauter Mail-Adressen bleibt sonst Text.
  if ((byValue === KINDS.EMAIL || byValue === KINDS.PHONE) && byHeader === KINDS.FULL_NAME) {
    return byValue;
  }

  // Eine als Text erkannte Namensspalte, die in Wahrheit Zahlen oder Daten
  // enthaelt, muss numerisch behandelt werden - sonst schreiben wir Text in
  // eine Zahlenspalte.
  const headerIsTextual = TEXTUAL_KINDS.has(byHeader);
  const valueIsNumeric = byValue === KINDS.INTEGER || byValue === KINDS.DECIMAL ||
    byValue === KINDS.DATE || byValue === KINDS.BOOLEAN;
  if (headerIsTextual && valueIsNumeric && byHeader !== KINDS.ZIP) {
    return byValue;
  }

  return byHeader;
}

const TEXTUAL_KINDS = new Set([
  KINDS.EMAIL, KINDS.PHONE, KINDS.FIRST_NAME, KINDS.LAST_NAME, KINDS.FULL_NAME,
  KINDS.STREET, KINDS.CITY, KINDS.ZIP, KINDS.COUNTRY, KINDS.COMPANY,
  KINDS.IBAN, KINDS.TEXT,
]);

function classifyByHeader(header) {
  if (typeof header !== 'string') return null;
  const name = header.trim();
  if (!name) return null;

  for (const [kind, pattern] of HEADER_PATTERNS) {
    if (pattern.test(name)) return kind;
  }
  return null;
}

/**
 * Leitet die Art aus den Werten ab. Ein Muster muss auf die Mehrheit der
 * Beispiele passen, damit ein einzelner Ausreisser die Spalte nicht umdeutet.
 */
function classifyByValue(samples) {
  if (!samples.length) return KINDS.EMPTY;

  const majority = Math.ceil(samples.length * 0.6);
  const matches = (predicate) => samples.filter(predicate).length >= majority;

  if (matches((v) => v instanceof Date)) return KINDS.DATE;
  if (matches((v) => typeof v === 'boolean')) return KINDS.BOOLEAN;

  if (matches((v) => typeof v === 'number')) {
    return matches((v) => typeof v === 'number' && Number.isInteger(v))
      ? KINDS.INTEGER
      : KINDS.DECIMAL;
  }

  const strings = samples.map((v) => String(v).trim());
  const strMatches = (re) => strings.filter((s) => re.test(s)).length >= majority;

  if (strMatches(EMAIL_RE)) return KINDS.EMAIL;
  if (strMatches(IBAN_RE)) return KINDS.IBAN;
  if (strMatches(ZIP_RE)) return KINDS.ZIP;
  if (strMatches(PHONE_RE)) return KINDS.PHONE;

  return KINDS.TEXT;
}
