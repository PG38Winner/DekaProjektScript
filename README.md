# Excel-Daten Anonymisierer

Kommandozeilen-Werkzeug zum **Verschleiern (Anonymisieren) personenbezogener
Daten** in Excel-Dateien (`.xlsx` / `.xlsm`).

Es arbeitet direkt auf der Datei – **ein installiertes Microsoft Excel wird
nicht benötigt**. Dadurch läuft es unter Windows, Linux und macOS, auf Servern
und in CI-Pipelines.

> **Hinweis zur Projektgeschichte:** Frühere Stände enthielten zusätzlich das
> PowerShell-Skript `Anonymize-AccessDb.ps1` mit grafischer Oberfläche und
> Access-Unterstützung. Es wurde in Commit `84e2f66` entfernt. **Access
> (`.accdb`/`.mdb`) wird derzeit nicht unterstützt** – dafür wäre weiterhin der
> ACE-OLEDB-Provider nötig. Über `git show 84e2f66^:Anonymize-AccessDb.ps1`
> lässt sich die alte Fassung bei Bedarf wieder herausholen.

## Einrichten

### Ohne Internetzugang (für den Zielrechner)

**Es ist nichts zu installieren und nichts zu entpacken.** Repository
herunterladen (grüner Knopf *Code → Download ZIP*), entpacken, Eingabe­auf­for­de­rung
im entpackten Ordner öffnen – fertig:

```
anonymisieren.cmd "C:\Daten\kunden.xlsx" --list
```

Alles Nötige liegt gebrauchsfertig im Repository:

| Bestandteil | Zweck |
|-------------|-------|
| `node-v24.19.0-win-x64\node.exe` | Node.js-Laufzeit für Windows |
| `dist\anonymize-xlsx.cjs` | das Werkzeug samt aller Abhängigkeiten in einer Datei |
| `anonymisieren.cmd` | Startskript – findet die Laufzeit selbst |

Ist auf dem Rechner bereits Node.js ab Version 20 installiert, verwendet das
Startskript dieses.

**Befehle direkt eintippen:** `node-umgebung.cmd` öffnet eine
Eingabeaufforderung, in der die mitgelieferte Laufzeit im Suchpfad liegt.
Darin funktioniert `node` ohne Installation:

```
node dist\anonymize-xlsx.cjs "C:\Daten\kunden.xlsx" --list
```

Unter Linux/macOS leistet `./anonymisieren.sh` dasselbe, setzt dort aber ein
installiertes Node.js voraus – die mitgelieferte Laufzeit ist eine
Windows-Version.

### Mit Internetzugang (für die Weiterentwicklung)

```bash
npm install          # Abhängigkeiten
npm test             # Tests
npm run build        # dist/anonymize-xlsx.cjs neu erzeugen
```

`npm run build` ist nach jeder Änderung an `src/` nötig, damit das Bündel den
Quellcode wieder abbildet. Ein Test wacht darüber: er vergleicht die Ausgabe
von Bündel und Quellcode bei gleichem Startwert.

## Verwenden

Die Beispiele verwenden `node src/cli.js`; mit dem Bündel entsprechend
`node dist/anonymize-xlsx.cjs` oder `anonymisieren.cmd`.

```bash
# Blätter und erkannte Spaltentypen anzeigen (verändert nichts)
node src/cli.js daten.xlsx --list

# Anonymisieren, Schlüsselspalten ausnehmen
node src/cli.js daten.xlsx --keep "Kunden:KundenID" --keep "Bestellungen:BestellID"

# In eine neue Datei schreiben, bestimmtes Blatt, reproduzierbar
node src/cli.js daten.xlsx --sheet Kunden --out anonym.xlsx --seed 42
```

| Option | Wirkung |
|--------|---------|
| `--list` | Blätter, Spalten und erkannte Inhaltsart anzeigen |
| `--sheet <name>` | nur dieses Arbeitsblatt (Standard: **alle** Blätter) |
| `--keep <angabe>` | Spalten, die **unverändert** bleiben – siehe unten |
| `--out <datei>` | Ergebnis in neue Datei statt Überschreiben |
| `--no-backup` | Keine Sicherungskopie anlegen |
| `--no-consistent` | Gleiche Werte dürfen unterschiedliche Ersatzwerte erhalten |
| `--seed <zahl>` | Fester Startwert für reproduzierbare Läufe |
| `--dry-run` | Nur anzeigen, was passieren würde |
| `--fast` | `--list` beschleunigen (liest die Datei nur teilweise) |
| `-h`, `--help` | Hilfe anzeigen |

> **Merksatz:**
> **In `--keep` genannt = bleibt unverändert** · **alles andere wird verschleiert**

Die erste Zeile eines Arbeitsblatts gilt als **Kopfzeile** mit den Spaltennamen.
Spaltennamen in `--keep` werden ohne Rücksicht auf Groß-/Kleinschreibung
verglichen; ein unbekannter Name bricht den Lauf ab, statt ihn stillschweigend
zu ignorieren.

### Geschwindigkeit

`--list` liest die Datei standardmäßig **vollständig** ein. Das ist der
belastbare Weg und funktioniert mit jeder lesbaren Datei.

Mit **`--fast`** wird stattdessen nur der Anfang gelesen: die ersten 200
Datenzeilen je Blatt – mehr braucht die Typerkennung nicht – und die Zeilenzahl
aus dem Kopf des Blattes. Gemessen an einer Datei mit 200.000 Zeilen ×
30 Spalten (29 MB):

| | Standard | mit `--fast` |
|---|---|---|
| Dauer | 27,9 s | 1,4 s |
| Arbeitsspeicher | 2,4 GB | gering |

> **`--fast` ist bewusst nicht der Standard.** In einer echten Arbeitsmappe
> lieferte der schnelle Weg eine leere Anzeige; die Ursache ist noch nicht
> geklärt. Erkennt er nichts Brauchbares, schaltet er selbsttätig auf den
> vollständigen Weg um – verlassen sollte man sich darauf aber erst, wenn er
> sich an den eigenen Dateien bewährt hat.

Das **Anonymisieren** muss die Datei zwangsläufig ganz einlesen und wieder
schreiben – dort bleibt es bei der Dauer, die Dateigröße und Excel-Format
vorgeben.

### Mehrere Arbeitsblätter

Ohne `--sheet` werden **alle** Arbeitsblätter der Datei verarbeitet – also
genau die, die `--list` anzeigt. Mit `--sheet Kunden` bleibt es bei dem einen
genannten Blatt.

Blätter ohne Kopfzeile (Deckblatt, Notizen) werden übersprungen und im Bericht
als solche ausgewiesen; der Lauf bricht deswegen nicht ab.

### `--keep` je Arbeitsblatt

Ohne Blattnamen gilt eine Angabe in **jedem** Arbeitsblatt:

```bash
--keep "KundenID"
```

Mit vorangestelltem Blattnamen gilt sie nur dort – und das Präfix bleibt für
alle **folgenden** Spalten derselben Angabe wirksam. So lassen sich je Tabelle
beliebig viele Spalten nennen:

```bash
# Drei Spalten, alle nur im Blatt "Kunden"
--keep "Kunden:KundenID,Nachname,Ort"
```

Ein neues Präfix wechselt das Blatt – beliebig oft, für beliebig viele
Tabellen in einer einzigen Angabe:

```bash
--keep "Kunden:KundenID,Nachname,Artikel:ArtikelNr,Preis,Rechnungen:RechnungsNr"
```

Gleichwertig und oft übersichtlicher, mehrfach angegeben:

```bash
node src/cli.js daten.xlsx \
  --keep "Kunden:KundenID,Nachname,Ort" \
  --keep "Bestellungen:BestellID,Betrag" \
  --keep "Artikel:ArtikelNr,Preis"
```

Jede `--keep`-Angabe beginnt wieder bei „gilt überall"; ein Präfix wirkt also
nie über die Angabe hinaus, in der es steht. Innerhalb einer Angabe schaltet
`*:` ausdrücklich zurück auf alle Blätter:

```bash
--keep "Kunden:Nachname,*:KundenID"   # Nachname nur in Kunden, KundenID überall
```

#### Feinheiten

- **Groß-/Kleinschreibung** spielt weder bei Blatt- noch bei Spaltennamen eine
  Rolle.
- Der **Doppelpunkt** eignet sich als Trenner, weil Excel ihn in Blattnamen
  nicht zulässt. Getrennt wird am **ersten** Doppelpunkt, ein Spaltenname darf
  also selbst welche enthalten.
- Enthält ein Spaltenname ein **Komma**, wird es als `\,` maskiert – oder man
  gibt `--keep` einfach mehrfach an.
- Ein **Tippfehler** bricht den Lauf ab, statt wirkungslos zu bleiben:

```
Fehlerhafte --keep-Angabe:
  Spalte "kundeid" gibt es im Blatt "Kunden" nicht. Vorhanden: KundenID, Vorname, ...
  Arbeitsblatt "Kundne" wird nicht verarbeitet. Verarbeitet werden: Kunden, Bestellungen
```

Was tatsächlich stehen bleibt, zeigt `--dry-run` vor dem echten Lauf:

```
Arbeitsblatt: Kunden
  KundenID     unveraendert     integer
  Vorname      anonymisiert     firstName   5 Zellen
  Nachname     unveraendert     lastName
  Ort          unveraendert     city
```

## Art der Anonymisierung

Die Ersetzung erfolgt typ- und inhaltsabhängig. Die Inhaltsart wird aus dem
**Spaltennamen** und dem **Zelleninhalt** abgeleitet:

| Inhalt / Typ | Ersetzung |
|--------------|-----------|
| E-Mail | erfundene Adresse `vorname.nachname@example.com` |
| Telefonnummer | zufällige Nummer (`+49 …`) |
| Vor-/Nachname, Straße, Ort, Land, Firma | passender Zufallswert (deutsche Namens-/Ortsdaten) |
| PLZ | Zufallszahl mit gleicher Stellenzahl |
| IBAN | erfundene deutsche IBAN |
| sonstiger Text | erfundener Fülltext, auf die Länge des Originals begrenzt |
| Ganzzahl | Zufallszahl in ähnlicher Größenordnung |
| Dezimal / Währung | Zufallswert in ähnlicher Größenordnung, gleiche Nachkommastellen |
| Datum | um bis zu ±365 Tage verschoben, Uhrzeit bleibt |
| Ja/Nein | zufälliger Wahrheitswert |

Erhalten bleiben dabei:

- **`NULL`- und Leerwerte** – sie werden nicht überschrieben.
- **Datentypen** – aus einer Zahl wird eine Zahl, aus einem Datum ein Datum.
- **Zellformate** – Datums- und Währungsformate bleiben bestehen.
- **Formelspalten** – sie werden übersprungen, die Formel bleibt stehen.
- **Makros** – siehe unten.

### Konsistente Ersetzung

Standardmäßig erhält derselbe Ausgangswert in Spalten gleichen Namens
**denselben** Ersatzwert – auch **über Arbeitsblätter hinweg**. Taucht „Anna
Müller" fünfmal auf, wird daraus fünfmal dieselbe erfundene Person; und steht
`KundenID` sowohl im Blatt *Kunden* als auch im Blatt *Bestellungen*, zeigen
die Bestellungen nach dem Lauf weiterhin auf denselben Kunden. Dadurch bleiben
Gruppierungen, Zählungen und Verknüpfungen auswertbar.

Mit `--no-consistent` erhält jede Zelle einen eigenen Zufallswert.

Mit `--seed <zahl>` wird derselbe Lauf reproduzierbar – nützlich für Tests.
Ohne Seed ist jeder Lauf anders.

## Makro-Arbeitsmappen (`.xlsm`)

Das VBA-Projekt wird **übernommen**. Die verwendete Bibliothek `exceljs` kennt
selbst keine Makros und würde sie beim Zurückschreiben verwerfen; das Werkzeug
sichert deshalb vor dem Schreiben `xl/vbaProject.bin` samt Signatur und
Codenamen und fügt sie danach wieder ein – einschließlich der Verweise in
`[Content_Types].xml` und den Beziehungen und des makrofähigen
Arbeitsmappentyps.

Erhalten bleiben: das VBA-Projekt, dessen Signatur (falls vorhanden), der
Codename der Arbeitsmappe und die Codenamen der Arbeitsblätter (damit VBA seine
Blätter wiederfindet).

**Nicht** erhalten bleiben Bestandteile, die `exceljs` grundsätzlich nicht
unterstützt – etwa ActiveX-Elemente, Formularsteuerelemente, Diagramme und
Pivot-Tabellen. Solche Verluste werden **gemeldet**:

```
WARNUNG: Diese Bestandteile der Originaldatei konnten nicht uebernommen
werden: xl/activeX/activeX1.xml
```

Schreibt man eine Makro-Datei per `--out` in eine `.xlsx`-Datei, gehen die
Makros verloren – auch darauf wird hingewiesen.

> Die Wiederherstellung ist gegen den Dateiaufbau geprüft (VBA-Projekt,
> Inhaltstypen, Beziehungen, Codenamen, ZIP-Integrität), **nicht** gegen ein
> echtes Microsoft Excel. Vor dem Einsatz an Produktivdaten bitte einmal mit
> einer eigenen Makro-Datei testen.

## Tests

```bash
npm test
```

60 Tests zu Typerkennung, Werterhaltung, Dateibehandlung, Makro-Erhalt und
Gleichlauf von Bündel und Quellcode. Die Bündel-Tests werden übersprungen,
solange `dist/` nicht gebaut ist.

## Grenzen

- Nur **Excel** (`.xlsx` / `.xlsm`). **Kein Access** (`.accdb`/`.mdb`).
- Das alte **`.xls`-Format** (BIFF) wird nicht gelesen; vorher in `.xlsx`
  umwandeln.
- **Keine grafische Oberfläche** – bisher reine Kommandozeile.
- Nicht unterstützte Arbeitsmappen-Bestandteile gehen verloren (siehe oben);
  sie werden aber immer gemeldet.
- `npm audit` meldet einen mittelschweren Hinweis auf `uuid` (transitive
  Abhängigkeit von `exceljs`). Er betrifft die UUID-Varianten v3/v5/v6 mit
  eigenem Puffer, die hier nicht verwendet werden; die einzige angebotene
  „Behebung" wäre ein Downgrade auf `exceljs` 3.x.

## Hinweise

- Die Änderungen sind **endgültig**. Wird die Originaldatei überschrieben, legt
  das Werkzeug standardmäßig eine **Sicherungskopie**
  `datei.backup-<zeitstempel>.xlsx` daneben (abschaltbar mit `--no-backup`).
  Mit `--out` bleibt das Original ohnehin unangetastet.
- Am besten **immer zuerst an einer Kopie** testen, oder `--dry-run` verwenden.
- Die Datei darf während der Verarbeitung **nicht** in Excel geöffnet sein.
- Anonymisierung ist kein Ersatz für eine Risikobewertung: Bleiben genug
  unveränderte Spalten stehen, können Datensätze weiterhin
  re-identifizierbar sein. Bei `--keep` sparsam sein.

## Aufbau

| Datei | Inhalt |
|-------|--------|
| `anonymisieren.cmd` | Start unter Windows |
| `node-umgebung.cmd` | Eingabeaufforderung mit `node`/`npm` im Suchpfad |
| `anonymisieren.sh` | Start unter Linux/macOS |
| `node-v24.19.0-win-x64/node.exe` | Node.js-Laufzeit für Windows |
| `dist/anonymize-xlsx.cjs` | Eigenständiges Bündel, erzeugt mit `npm run build` |
| `src/cli.js` | Kommandozeile, Ausgabe der Berichte |
| `src/anonymize.js` | Arbeitsmappe lesen, ersetzen, schreiben |
| `src/inspect.js` | schnelle Struktur-Analyse für `--list` |
| `src/ooxml.js` | direkter Zugriff auf den Dateiaufbau |
| `src/cells.js` | Umgang mit den Zellformen von exceljs |
| `src/keep.js` | Zerlegen und Anwenden der `--keep`-Angaben |
| `src/classify.js` | Erkennung der Inhaltsart je Spalte |
| `src/generators.js` | Erzeugung der Ersatzwerte |
| `src/macros.js` | Erhalt des VBA-Projekts bei `.xlsm` |
| `test/` | Tests und Beispieldateien |

## Mitgelieferte Node.js-Laufzeit

Im Repository liegt die offizielle, portable Node.js-Laufzeit für Windows –
**bereits entpackt** als `node-v24.19.0-win-x64\node.exe`. Sie läuft ohne
Installation und ohne Administrator­rechte.

**Reduziert auf `node.exe`:** Aus der Original-Auslieferung wurden `npm`,
`npx` und `corepack` entfernt. Zum Ausführen des Werkzeugs werden sie nicht
gebraucht, und ihre tief verschachtelten Ordner sprengten beim Entpacken unter
Windows die Pfadlängengrenze von 260 Zeichen (`Fehler 0x80010135: Pfad zu
lang`). Der längste Pfad im Repository ist dadurch von 125 auf 30 Zeichen
gesunken. Für die Weiterentwicklung ein vollständiges Node.js installieren.

- Version: **v24.19.0** (LTS „Krypton")
- Quelle: <https://nodejs.org/dist/v24.19.0/node-v24.19.0-win-x64.zip>

### Herkunft nachprüfen

Das Archiv wurde beim Herunterladen gegen die offizielle `SHASUMS256.txt`
geprüft und anschließend entpackt; das Archiv selbst liegt nicht mehr bei.
Die Prüfsummen der Kette:

| Gegenstand | SHA-256 |
|------------|---------|
| Archiv von nodejs.org | `57f71ab3652e797d84acddc79c81cc9ff1c6ddb2a1974cdb83f00fee9bff4c73` |
| daraus entpackte `node.exe` | `3602f2bb1a10f2cbab4c36886218a33c1ab3db87290e73b033c46c77147d0237` |

Die zweite Zeile lässt sich jederzeit gegen die eingecheckte Datei prüfen:

```powershell
Get-FileHash node-v24.19.0-win-x64\node.exe -Algorithm SHA256
```

Wer der Kette nicht traut, lädt das Archiv selbst von nodejs.org, prüft es
gegen `SHASUMS256.txt` und vergleicht die entpackte `node.exe`.

### Falls das Entpacken scheitert

Meldet Windows beim Entpacken `Fehler 0x80010135: Pfad zu lang`, ist der
Zielpfad zu tief. Der Explorer legt beim *Alle extrahieren* standardmäßig einen
Unterordner mit dem Namen des Archivs an – und das GitHub-ZIP enthält bereits
einen gleichnamigen Ordner, sodass der Name doppelt im Pfad steht. Abhilfe:

- beim Entpacken einen **kurzen Zielpfad** wählen, z. B. `C:\deka`
- oder mit **7-Zip** entpacken, das die Grenze nicht kennt

### Größe des Repositorys

Die Laufzeit macht das Repository groß: **rund 92 MB**, davon allein 89 MB
`node.exe`. Der Download dauert entsprechend – dafür ist auf dem
Zielrechner kein einziger Installationsschritt nötig.

Das ZIP-Archiv war zwischenzeitlich ebenfalls eingecheckt und wurde entfernt,
da es neben dem entpackten Ordner keinen Zweck mehr erfüllte. Es steckt
weiterhin in der Git-Historie – ein vollständiger `git clone` überträgt es also
mit. Wer nur den aktuellen Stand braucht, spart das:

| Bezugsweg | Gesamt | davon `.git` |
|-----------|--------|--------------|
| *Code → Download ZIP* auf GitHub | 92 MB, keine Historie | – |
| `git clone --depth 1` | 126 MB | 34 MB |
| `git clone` (vollständig) | 164 MB | 72 MB |

```bash
git clone --depth 1 https://github.com/PG38Winner/DekaProjektScript.git
```

Für den Zielrechner ist der ZIP-Download der einfachste Weg – er enthält keine
Historie und setzt kein Git voraus.

> Die Windows-Startskripte (`anonymisieren.cmd`, `node-umgebung.cmd`) sind unter
> Windows **nicht** erprobt – diese Entwicklungsumgebung ist Linux. Getestet
> sind der Linux-Start (`anonymisieren.sh`), das Bündel, die Anonymisierung und
> die Unversehrtheit der eingecheckten `node.exe`. Bitte den ersten
> Windows-Aufruf einmal beobachten.
