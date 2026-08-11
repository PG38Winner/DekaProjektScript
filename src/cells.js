/**
 * Umgang mit den verschiedenen Formen, in denen exceljs den Inhalt einer Zelle
 * liefert - als einfacher Wert, als Datum oder als Objekt fuer Formeln,
 * formatierten Text, Verweise und Fehlerwerte.
 */

/**
 * Reduziert eine Zelle auf einen einfachen JavaScript-Wert.
 * @returns {unknown} `null` fuer leere Zellen.
 */
export function toPlainValue(raw) {
  if (raw === null || raw === undefined) return null;
  if (raw instanceof Date) return raw;
  if (typeof raw !== 'object') return raw;

  if ('richText' in raw) return raw.richText.map((part) => part.text).join('');
  if ('text' in raw) return raw.text;
  if ('formula' in raw || 'sharedFormula' in raw) return raw.result ?? null;
  if ('error' in raw) return null;

  return String(raw);
}

/** Formeln und Fehlerwerte werden nicht ueberschrieben. */
export function isReadOnlyValue(raw) {
  if (!raw || typeof raw !== 'object') return false;
  return 'formula' in raw || 'sharedFormula' in raw || 'error' in raw;
}
