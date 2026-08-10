# Access-Datenbank Anonymisierer (`Anonymize-AccessDb.ps1`)

PowerShell-Script mit grafischer Oberfläche zum **Verschleiern (Anonymisieren)
personenbezogener Daten** in Microsoft-Access-Datenbanken (`.accdb` / `.mdb`).

## Funktionsweise

1. Access-Datei über den Datei-Dialog auswählen und **Laden** klicken.
2. Tabelle aus dem Dropdown wählen – die Spalten erscheinen als Checkbox-Liste.
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

`NULL`- und Leerwerte bleiben erhalten. Autowert-/RowID-Spalten und nicht
beschreibbare Felder werden automatisch übersprungen.

## Voraussetzungen

- **Windows** mit **Windows PowerShell 5.1** (für WinForms/COM).
- **Microsoft Access Database Engine** (ACE-OLEDB-Provider).
  - Für `.accdb`: `Microsoft.ACE.OLEDB.16.0` bzw. `.12.0`.
  - Für alte `.mdb`: ggf. `Microsoft.Jet.OLEDB.4.0` (nur 32-Bit).
  - **Wichtig:** Die Bit-Version (32/64) des Providers muss zur PowerShell-
    Bit-Version passen. Bei Problemen `powershell.exe` aus
    `C:\Windows\SysWOW64\WindowsPowerShell\v1.0\` (32-Bit) verwenden.

## Ausführen

```powershell
# ggf. Ausführungsrichtlinie für die aktuelle Sitzung lockern
powershell -ExecutionPolicy Bypass -File .\Anonymize-AccessDb.ps1
```

## Hinweise

- Die Änderungen sind **endgültig**. Vor der Ausführung wird standardmäßig eine
  **Backup-Kopie** der Datenbank angelegt (Option in der GUI abschaltbar).
- Am besten **immer zuerst an einer Kopie** testen.
- Die Datenbank darf während der Verarbeitung **nicht** in Access geöffnet sein.
