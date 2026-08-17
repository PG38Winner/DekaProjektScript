/**
 * Stellt im Browser die wenigen Node-Bausteine bereit, die einzelne
 * Abhaengigkeiten voraussetzen (Buffer, process). Beides sind reine
 * Hilfsobjekte - es wird nichts vom Betriebssystem angesprochen.
 */

import { Buffer } from 'buffer';

const process = { env: {}, browser: true, version: '', nextTick: (fn, ...args) => setTimeout(() => fn(...args), 0) };

export { Buffer, process };
