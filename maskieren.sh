#!/bin/sh
# ---------------------------------------------------------------------------
#  Daten-Maskierer - Start ohne Installation (Linux/macOS).
#
#  Aufruf:  ./maskieren.sh ~/daten/kunden.xlsx --list
# ---------------------------------------------------------------------------
set -eu

ROOT=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
BUNDLE="$ROOT/dist/maskierer.cjs"

if [ ! -f "$BUNDLE" ]; then
    echo "Fehler: \"$BUNDLE\" wurde nicht gefunden." >&2
    echo "Bitte das Repository vollstaendig auschecken." >&2
    exit 1
fi

if ! command -v node >/dev/null 2>&1; then
    echo "Fehler: Keine Node.js-Laufzeit gefunden." >&2
    echo "Node.js ab Version 20 installieren." >&2
    echo "Die mitgelieferte ZIP-Datei enthaelt nur die Windows-Laufzeit." >&2
    exit 1
fi

exec node "$BUNDLE" "$@"
