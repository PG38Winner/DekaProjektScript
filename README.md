# Access-/Excel-Daten Maskierung (`Anonymize-AccessDb.ps1`)

PowerShell-Script mit grafischer Oberfläche zum **Maskieren personenbezogener
Daten** (Ersetzen durch erfundene Test-Werte) in

- **Microsoft-Access-Datenbanken** (`.accdb` / `.mdb`) und
- **Excel-Dateien** (`.xlsx` / `.xlsm` / `.xlsb` / `.xls`).

Der Dateityp wird automatisch an der Endung erkannt.

> **Hinweis (Datenschutz):** Dies ist eine **Maskierung / Testdaten-Ersetzung**,
> **keine** formal validierte Anonymisierung im Sinne der DSGVO (keine
> Re-Identifikations-Risikoanalyse, kein geprüftes Verfahren, nicht-kryptografischer
> Zufall). Für Testdaten geeignet – nicht als regulatorisch belastbare
> Anonymisierung/Pseudonymisierung.

## Funktionsweise

1. Access- **oder** Excel-Datei über den Datei-Dialog auswählen und **Laden** klicken.
2. **Tabelle** (Access) bzw. **Arbeitsblatt** (Excel) aus dem Dropdown wählen –
   die Spalten erscheinen als Checkbox-Liste. Bei Excel ist die **erste Zeile**
   die Kopfzeile mit den Spaltennamen (mit Spaltennummer, damit gleiche
   Überschriften unterscheidbar bleiben).
3. Spalten **ankreuzen**, die **unverändert** bleiben sollen
   (z. B. Primärschlüssel, IDs, Referenzen, Join-Spalten).
4. **Nicht angekreuzte** Spalten werden **maskiert**.
5. Optionen wählen (siehe unten) und **Maskierung starten** klicken.

> **Merksatz:**
> **Angekreuzt = bleibt unverändert** · **Nicht angekreuzt = wird maskiert**

## Optionen: Arbeitskopie vs. direkte Änderung

- **Original nicht ändern (Ausgabe in Kopie)** – *Standard, empfohlen.* Es wird
  eine Kopie `…_maskiert_<Zeitstempel>.<ext>` erzeugt und **nur diese** verändert;
  das Original bleibt unangetastet (kein Risiko einer teilweise veränderten
  Originaldatei).
- Ist diese Option **aus**, wird die Datei **direkt** verändert – dann optional
  mit **Backup-Kopie** des Originals. Bei Access läuft die direkte Änderung in
  einer **Transaktion** (Rollback bei Fehler, sofern der Provider das unterstützt).

## Art der Maskierung

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

**Konsistente Ersetzung (pro Spalte):** Gleiche Originalwerte innerhalb
**derselben Spalte** werden immer gleich ersetzt (z. B. „Max Müller" → immer
derselbe Fake-Name). Der Schlüssel ist `Spalte|Originalwert`, damit derselbe
Wert in unterschiedlichen Spalten (z. B. Nachname „Berlin" vs. Stadt „Berlin")
nicht denselben, fachlich falschen Ersatz erhält.

**Excel-Besonderheiten:** Spalten mit **Formeln** werden übersprungen (Formeln
bleiben erhalten). Excel wird **sichtbar** gestartet, damit das Verhalten
nachvollziehbar ist.

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

## Hinweise

- Die Änderungen sind **endgültig**. Vor der Ausführung wird standardmäßig eine
  **Backup-Kopie** der Datei angelegt (Option in der GUI abschaltbar).
- Am besten **immer zuerst an einer Kopie** testen.
- Die Datei darf während der Verarbeitung **nicht** in Access bzw. Excel
  geöffnet sein.
- **Backup enthält Originaldaten:** Die automatische Sicherungskopie enthält
  weiterhin die **echten personenbezogenen Daten**. Sie muss genauso geschützt
  und nach Freigabe **gelöscht** werden. Wer keine Kopie mit Originaldaten will,
  deaktiviert die Backup-Option.
- **Excel-Datumsfelder:** Werte werden über `.Value` gelesen, sodass
  Datumszellen als echtes Datum erkannt und (wie bei Access) um zufällige Tage
  verschoben werden.
- **Zufall:** Die Ersatzwerte werden mit `System.Random` erzeugt – geeignet für
  Test-/Demodaten, **nicht** für kryptografisch starke Pseudonymisierung.
