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

```powershell
# ggf. Ausführungsrichtlinie für die aktuelle Sitzung lockern
powershell -ExecutionPolicy Bypass -File .\Anonymize-AccessDb.ps1
```

## Hinweise

- Die Änderungen sind **endgültig**. Vor der Ausführung wird standardmäßig eine
  **Backup-Kopie** der Datei angelegt (Option in der GUI abschaltbar).
- Am besten **immer zuerst an einer Kopie** testen.
- Die Datei darf während der Verarbeitung **nicht** in Access bzw. Excel
  geöffnet sein.
- **Excel-Datumsfelder:** Werte werden über die schnelle `Value2`-Schnittstelle
  gelesen; Datumszellen werden dabei wie Zahlen behandelt (der Wert wird
  verschleiert, das Zellformat bleibt erhalten).
