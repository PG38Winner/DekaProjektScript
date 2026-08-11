# Access-/Excel-Daten Anonymisierer (`Anonymize-AccessDb.ps1`)

PowerShell-Script mit grafischer Oberfläche zum **Verschleiern (Anonymisieren)
personenbezogener Daten** in

- **Microsoft-Access-Datenbanken** (`.accdb` / `.mdb`) und
- **Excel-Dateien** (`.xlsx` / `.xlsm` / `.xlsb` / `.xls`).

Der Dateityp wird automatisch an der Endung erkannt.

## Funktionsweise

1. Access- **oder** Excel-Datei über den Datei-Dialog auswählen und **Laden** klicken.
2. **Tabelle** (Access) bzw. **Arbeitsblatt** (Excel) aus dem Dropdown wählen –
   die Spalten erscheinen als Checkbox-Liste. Bei Excel ist die **erste Zeile**
   die Kopfzeile mit den Spaltennamen.
3. Spalten **ankreuzen**, die **unverändert** bleiben sollen
   (z. B. Primärschlüssel, IDs, Referenzen).
4. **Nicht angekreuzte** Spalten werden **anonymisiert**.
5. **Verschleiern starten** klicken.

> **Merksatz:**
> **Angekreuzt = bleibt unverändert** · **Nicht angekreuzt = wird verschleiert**

## Art der Anonymisierung

Die Ersetzung erfolgt typ- und inhaltsabhängig:

| Inhalt / Typ            | Ersetzung                                    |
|-------------------------|----------------------------------------------|
| E-Mail                  | erfundene Adresse `vorname.nachname@example.com` |
| Telefonnummer           | zufällige Nummer (`+49 …`)                    |
| Vor-/Nachname, Adresse, Ort, PLZ | passender Zufallswert aus Namens-/Ortspools |
| sonstiger Text          | erfundener Fülltext (Länge wird begrenzt)     |
| Ganzzahl                | Zufallszahl in ähnlicher Größenordnung        |
| Dezimal / Währung       | Zufallswert in ähnlicher Größenordnung        |
| Datum                   | um zufällige Tage verschoben                  |
| Ja/Nein                 | zufälliger Wahrheitswert                      |

`NULL`- und Leerwerte bleiben erhalten. Bei Access werden Autowert-/RowID-Spalten
und nicht beschreibbare Felder automatisch übersprungen.

## Zwei Engines

| Dateityp | Zugriff | Schreiben |
|----------|---------|-----------|
| Access (`.accdb`/`.mdb`) | ADODB / ACE-OLEDB | editierbarer Server-Cursor (schreibt sofort in die Datei) |
| Excel (`.xlsx`/`.xls` …)  | Excel-COM-Automation | Spaltenweise, danach `Speichern` |

## Voraussetzungen

- **Windows** mit **Windows PowerShell 5.1** (für WinForms/COM).
- **Für Access:** **Microsoft Access Database Engine** (ACE-OLEDB-Provider).
  - Für `.accdb`: `Microsoft.ACE.OLEDB.16.0` bzw. `.12.0`.
  - Für alte `.mdb`: ggf. `Microsoft.Jet.OLEDB.4.0` (nur 32-Bit).
  - **Wichtig:** Die Bit-Version (32/64) des Providers muss zur PowerShell-
    Bit-Version passen. Bei Problemen `powershell.exe` aus
    `C:\Windows\SysWOW64\WindowsPowerShell\v1.0\` (32-Bit) verwenden.
- **Für Excel:** installiertes **Microsoft Excel** (COM-Automation).

## Ausführen

Empfohlen (signiertes Skript, keine gelockerte Richtlinie nötig):

```powershell
powershell -File .\Anonymize-AccessDb.ps1
```

Falls die Ausführungsrichtlinie unsignierte lokale Skripte blockiert, den
sauberen Weg wählen — **nicht** `Bypass`, sondern für den aktuellen Benutzer:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

Am besten das Skript **digital signieren** (Authenticode) und in der
Sicherheitssoftware freigeben lassen – siehe Abschnitt „Fehlalarm der
Sicherheitssoftware".

## Fehlalarm der Sicherheitssoftware

Das Skript enthält **keinen Schadcode**. Weil es jedoch viele Datensätze
massenhaft überschreibt (das ähnelt für Verhaltensheuristiken Ransomware) und
PowerShell + COM verwendet, kann es einen **Fehlalarm** auslösen. Der saubere,
transparente Umgang damit:

1. **Skript digital signieren** (Authenticode-Code-Signing-Zertifikat):
   ```powershell
   Set-AuthenticodeSignature -FilePath .\Anonymize-AccessDb.ps1 `
       -Certificate $cert -TimeStampServer "http://timestamp.digicert.com"
   ```
2. **In der Sicherheitssoftware freigeben** (Allow-List / Ausschluss anhand
   des Datei-Hashes oder Zertifikats) – über die zuständige IT/den Admin.
3. **Fehlalarm an den Hersteller melden** (z. B. Microsoft Defender:
   „Submit a file for analysis"), damit die Erkennung generell korrigiert wird.

## Node.js-Variante: Excel-Anonymisierer (CLI)

Neben dem PowerShell-Skript liegt eine Node.js-Umsetzung für **Excel-Dateien**
im Ordner `src/`. Sie arbeitet direkt auf der Datei – **ein installiertes
Microsoft Excel wird nicht benötigt**, und sie läuft dadurch auch auf Servern,
in CI-Pipelines und unter Linux/macOS.

### Installation

```bash
npm install
```

### Verwenden

```bash
# Blätter und erkannte Spaltentypen anzeigen (verändert nichts)
node src/cli.js daten.xlsx --list

# Anonymisieren, Schlüsselspalten ausnehmen
node src/cli.js daten.xlsx --keep "KundenID,Bestellnummer"

# In eine neue Datei schreiben, bestimmtes Blatt, reproduzierbar
node src/cli.js daten.xlsx --sheet Kunden --out anonym.xlsx --seed 42
```

| Option | Wirkung |
|--------|---------|
| `--list` | Blätter, Spalten und erkannte Inhaltsart anzeigen |
| `--sheet <name>` | Arbeitsblatt wählen (Standard: erstes Blatt) |
| `--keep <a,b,c>` | Spalten, die **unverändert** bleiben |
| `--out <datei>` | Ergebnis in neue Datei statt Überschreiben |
| `--no-backup` | Keine Sicherungskopie anlegen |
| `--no-consistent` | Gleiche Werte dürfen unterschiedliche Ersatzwerte erhalten |
| `--seed <zahl>` | Fester Startwert für reproduzierbare Läufe |
| `--dry-run` | Nur anzeigen, was passieren würde |

> **Merksatz (wie in der GUI):**
> **In `--keep` genannt = bleibt unverändert** · **alles andere wird verschleiert**

### Verhalten

- **Typerkennung** aus Spaltenname *und* Inhalt: E-Mail, Telefon, Vor-/Nachname,
  Straße, PLZ, Ort, Land, Firma, IBAN, Text, Ganzzahl, Dezimal, Datum, Ja/Nein.
- **Typ- und größenordnungserhaltend**: Zahlen bleiben Zahlen ähnlicher Größe,
  Datumswerte werden um bis zu ±365 Tage verschoben, Texte behalten ihre Länge.
- **Zellformate bleiben erhalten** (Datums- und Währungsformate).
- **`NULL`/Leerwerte bleiben erhalten**, **Formelspalten werden übersprungen** –
  analog zu den nicht beschreibbaren Feldern der Access-Variante.
- **Konsistente Ersetzung** (Standard): derselbe Ausgangswert erhält innerhalb
  einer Spalte denselben Ersatzwert. Dadurch bleiben Gruppierungen und
  Verknüpfungen über die Spalte hinweg auswertbar. Abschaltbar mit
  `--no-consistent`.
- **Sicherungskopie** wird standardmäßig angelegt, wenn die Originaldatei
  überschrieben wird.

### Tests

```bash
npm test
```

### Grenzen

- Nur **Excel** (`.xlsx`/`.xlsm`). Access (`.accdb`/`.mdb`) bleibt beim
  PowerShell-Skript – zum Schreiben von Access ist weiterhin ACE-OLEDB nötig.
- Das alte `.xls`-Format (BIFF) wird nicht gelesen; vorher in `.xlsx` umwandeln.
- Noch ohne grafische Oberfläche.
- `npm audit` meldet einen mittelschweren Hinweis auf `uuid` (transitive
  Abhängigkeit von `exceljs`). Er betrifft die UUID-Varianten v3/v5/v6 mit
  eigenem Puffer, die hier nicht verwendet werden; die einzige angebotene
  „Behebung" wäre ein Downgrade auf `exceljs` 3.x.

## Mitgelieferte Node.js-Laufzeit

Im Repository liegt die offizielle, portable Node.js-Laufzeit für Windows:

- Datei: `node-v24.19.0-win-x64.zip` (ca. 36 MB)
- Version: **v24.19.0** (LTS „Krypton")
- Quelle: <https://nodejs.org/dist/v24.19.0/node-v24.19.0-win-x64.zip>
- SHA-256: `57f71ab3652e797d84acddc79c81cc9ff1c6ddb2a1974cdb83f00fee9bff4c73`
  (geprüft gegen die offizielle `SHASUMS256.txt`)

Zum Verwenden das ZIP entpacken – `node.exe` und `npm.cmd` liegen darin
direkt im Ordner `node-v24.19.0-win-x64\` und laufen ohne Installation.

## Hinweise

- Die Änderungen sind **endgültig**. Vor der Ausführung wird standardmäßig eine
  **Backup-Kopie** der Datei angelegt (Option in der GUI abschaltbar).
- Am besten **immer zuerst an einer Kopie** testen.
- Die Datei darf während der Verarbeitung **nicht** in Access bzw. Excel
  geöffnet sein.
- **Excel-Datumsfelder:** Werte werden über die schnelle `Value2`-Schnittstelle
  gelesen; Datumszellen werden dabei wie Zahlen behandelt (der Wert wird
  verschleiert, das Zellformat bleibt erhalten).
