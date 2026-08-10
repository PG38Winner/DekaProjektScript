<#
.SYNOPSIS
    Liest eine Microsoft-Access-Datenbank (.accdb / .mdb) und anonymisiert
    ("verschleiert") ausgewaehlte Spalten mit personenbezogenen Daten.

.DESCRIPTION
    Das Script oeffnet eine grafische Oberflaeche (WinForms):

      1. Access-Datei auswaehlen  (Datei-Dialog)
      2. Tabelle auswaehlen        (Dropdown)
      3. Spalten werden in einer Checkbox-Liste angezeigt

    Logik der Auswahl:
      * ANGEKREUZT   = Spalte bleibt UNVERAENDERT  (z. B. Schluessel, Referenzen)
      * NICHT angekreuzt = Spalte wird VERSCHLEIERT / anonymisiert

    Die Anonymisierung erfolgt typ-abhaengig:
      * Text        -> je nach Inhalt E-Mail / Telefon / Name / generischer Text
      * Zahl        -> zufaelliger Wert in aehnlicher Groessenordnung
      * Datum       -> um zufaellige Tage verschoben
      * Ja/Nein     -> zufaelliger Boolescher Wert

    Vor jeder Aenderung wird automatisch eine Backup-Kopie der Datenbank
    angelegt (kann in der GUI deaktiviert werden).

.REQUIREMENTS
    * Windows mit Windows PowerShell 5.1 (WinForms / COM).
    * "Microsoft Access Database Engine" (ACE OLEDB Provider).
      Wichtig: Die Bit-Version (32/64) des Providers muss zur PowerShell-
      Bit-Version passen. Fuer .accdb i. d. R. Microsoft.ACE.OLEDB.12.0
      oder .16.0. Fuer alte .mdb ggf. Microsoft.Jet.OLEDB.4.0 (nur 32-Bit).

.NOTES
    Nur auf Kopien / mit Backup ausfuehren. Die Aenderungen sind endgueltig.
#>

# ------------------------------------------------------------------------
# Voraussetzungen laden
# ------------------------------------------------------------------------
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Zufallsgenerator (einmalig)
$script:Rnd = [System.Random]::new()

# ------------------------------------------------------------------------
# Daten-Pools fuer realistisch wirkende, aber erfundene Werte
# ------------------------------------------------------------------------
$script:Vornamen = @('Anna','Lena','Marie','Sophie','Laura','Julia','Sarah','Lisa',
                     'Max','Paul','Leon','Felix','Jonas','Lukas','Tim','Jan',
                     'Nina','Emma','Mia','Ben','Finn','Noah','Elena','David')
$script:Nachnamen = @('Mueller','Schmidt','Schneider','Fischer','Weber','Meyer','Wagner',
                      'Becker','Schulz','Hoffmann','Koch','Bauer','Richter','Klein',
                      'Wolf','Schroeder','Neumann','Braun','Zimmermann','Krueger')
$script:Strassen = @('Hauptstrasse','Bahnhofstrasse','Gartenweg','Lindenallee','Schulstrasse',
                     'Bergstrasse','Kirchweg',' Amselweg','Rosenweg','Feldstrasse')
$script:Staedte  = @('Berlin','Hamburg','Muenchen','Koeln','Frankfurt','Stuttgart',
                     'Duesseldorf','Leipzig','Dresden','Hannover','Bremen','Essen')
$script:Woerter  = @('alpha','beta','gamma','delta','omega','sigma','lima','kilo',
                     'echo','tango','viktor','xray','yankee','zulu','nova','orbit')

# ------------------------------------------------------------------------
# ADODB DataType / FieldAttribute Konstanten (Auszug)
# ------------------------------------------------------------------------
$adText     = @(200,201,202,203,129,130,131,141)   # varchar/char/wchar/longtext
$adInteger  = @(2,3,16,17,18,19,20,21)             # (un)signed int/tinyint/bigint
$adDecimal  = @(4,5,6,14,131)                        # single/double/currency/numeric
$adDate     = @(7,133,134,135)                       # date/timestamp
$adBoolean  = @(11)
$adGuid     = @(72)

$adFldUpdatable = 4      # Feld ist beschreibbar
$adFldRowID     = 256    # Autowert / RowID  -> nicht anfassen

# ------------------------------------------------------------------------
# Hilfsfunktionen zur Anonymisierung
# ------------------------------------------------------------------------
function Get-RandomFrom { param([string[]]$Pool) $Pool[$script:Rnd.Next(0,$Pool.Count)] }

function Test-LooksLikeEmail { param([string]$Value) $Value -match '^[^@\s]+@[^@\s]+\.[^@\s]+$' }

function Test-LooksLikePhone {
    param([string]$Value)
    # Ueberwiegend Ziffern, +, /, -, Leerzeichen, Klammern und mind. 5 Ziffern
    ($Value -match '^[\d\s\+\-\/\(\)]+$') -and (($Value -replace '\D','').Length -ge 5)
}

function New-FakeEmail {
    $v = (Get-RandomFrom $script:Vornamen).ToLower()
    $n = (Get-RandomFrom $script:Nachnamen).ToLower()
    "$v.$n$($script:Rnd.Next(1,999))@example.com"
}

function New-FakePhone {
    "+49 {0} {1}" -f $script:Rnd.Next(150,179), $script:Rnd.Next(1000000,9999999)
}

function New-FakeText {
    param([int]$MaxLen)
    # Realistisch wirkender Fuelltext aus dem Wort-Pool
    $parts = @()
    for ($i=0; $i -lt $script:Rnd.Next(1,4); $i++) { $parts += Get-RandomFrom $script:Woerter }
    $text = ($parts -join '-') + $script:Rnd.Next(10,99)
    if ($MaxLen -gt 0 -and $text.Length -gt $MaxLen) { $text = $text.Substring(0,$MaxLen) }
    $text
}

# Liefert einen anonymisierten Wert passend zum Feldtyp / Inhalt
function Get-ObfuscatedValue {
    param(
        $OriginalValue,
        [int]$DataType,
        [int]$MaxLen,
        [string]$ColumnName
    )

    # NULL / leer unveraendert lassen (Struktur bleibt erhalten)
    if ($null -eq $OriginalValue -or $OriginalValue -is [System.DBNull]) { return [System.DBNull]::Value }
    $strVal = [string]$OriginalValue
    if ($strVal.Trim().Length -eq 0) { return $OriginalValue }

    # -------- Text ----------------------------------------------------
    if ($adText -contains $DataType) {
        if (Test-LooksLikeEmail $strVal) { return (New-FakeEmail) }
        if (Test-LooksLikePhone $strVal) { $p = New-FakePhone; if ($MaxLen -gt 0 -and $p.Length -gt $MaxLen) { $p = $p.Substring(0,$MaxLen) }; return $p }

        # Spaltenname als Hinweis nutzen
        $lc = $ColumnName.ToLower()
        switch -Regex ($lc) {
            'vorname|firstname|first_name'          { $v = Get-RandomFrom $script:Vornamen;  if ($MaxLen -gt 0 -and $v.Length -gt $MaxLen) { $v=$v.Substring(0,$MaxLen) }; return $v }
            'nachname|lastname|last_name|name'      { $v = Get-RandomFrom $script:Nachnamen; if ($MaxLen -gt 0 -and $v.Length -gt $MaxLen) { $v=$v.Substring(0,$MaxLen) }; return $v }
            'strasse|street|adresse|address'        { $v = "{0} {1}" -f (Get-RandomFrom $script:Strassen), $script:Rnd.Next(1,199); if ($MaxLen -gt 0 -and $v.Length -gt $MaxLen) { $v=$v.Substring(0,$MaxLen) }; return $v }
            'ort|stadt|city'                        { $v = Get-RandomFrom $script:Staedte;   if ($MaxLen -gt 0 -and $v.Length -gt $MaxLen) { $v=$v.Substring(0,$MaxLen) }; return $v }
            'plz|zip|postal'                        { return ("{0:D5}" -f $script:Rnd.Next(1000,99999)) }
            default                                 { return (New-FakeText -MaxLen $MaxLen) }
        }
    }

    # -------- Ganzzahl ------------------------------------------------
    if ($adInteger -contains $DataType) {
        $orig = 0
        [void][int64]::TryParse($strVal, [ref]$orig)
        $magnitude = [Math]::Max(10, [Math]::Abs($orig))
        return [int64]$script:Rnd.Next(0, [int][Math]::Min([int]::MaxValue, ($magnitude*2)+10))
    }

    # -------- Dezimal / Waehrung -------------------------------------
    if ($adDecimal -contains $DataType) {
        $orig = 0.0
        [void][double]::TryParse($strVal, [ref]$orig)
        $magnitude = [Math]::Max(10.0, [Math]::Abs($orig))
        return [Math]::Round(($script:Rnd.NextDouble() * $magnitude * 2), 2)
    }

    # -------- Datum ---------------------------------------------------
    if ($adDate -contains $DataType) {
        $d = [datetime]::Now
        if ([datetime]::TryParse($strVal, [ref]$d)) {
            return $d.AddDays($script:Rnd.Next(-365,365))
        }
        return (Get-Date).AddDays(-$script:Rnd.Next(0,3650))
    }

    # -------- Ja/Nein ------------------------------------------------
    if ($adBoolean -contains $DataType) { return [bool]($script:Rnd.Next(0,2)) }

    # -------- GUID ---------------------------------------------------
    if ($adGuid -contains $DataType) { return [guid]::NewGuid().ToString('B') }

    # -------- Fallback ------------------------------------------------
    return (New-FakeText -MaxLen $MaxLen)
}

# ------------------------------------------------------------------------
# Access-Verbindung / Provider ermitteln
# ------------------------------------------------------------------------
function New-AccessConnection {
    param([string]$Path)

    $providers = @('Microsoft.ACE.OLEDB.16.0','Microsoft.ACE.OLEDB.12.0')
    if ([System.IO.Path]::GetExtension($Path).ToLower() -eq '.mdb') {
        $providers += 'Microsoft.Jet.OLEDB.4.0'
    }

    foreach ($p in $providers) {
        try {
            $conn = New-Object -ComObject ADODB.Connection
            $conn.Open("Provider=$p;Data Source=$Path;Persist Security Info=False;")
            return $conn
        } catch {
            if ($conn) { try { $conn.Close() } catch {} }
        }
    }
    throw "Kein passender Access-OLEDB-Provider gefunden. Bitte 'Microsoft Access Database Engine' installieren (Bit-Version muss zu PowerShell passen)."
}

function Get-AccessTables {
    param($Conn)
    $tables = @()
    $adSchemaTables = 20
    $rs = $Conn.OpenSchema($adSchemaTables)
    while (-not $rs.EOF) {
        $type = [string]$rs.Fields.Item('TABLE_TYPE').Value
        if ($type -eq 'TABLE') { $tables += [string]$rs.Fields.Item('TABLE_NAME').Value }
        $rs.MoveNext()
    }
    $rs.Close()
    $tables | Sort-Object
}

function Get-AccessColumns {
    param($Conn, [string]$Table)
    $cols = @()
    $rs = New-Object -ComObject ADODB.Recordset
    $rs.Open("SELECT * FROM [$Table] WHERE 1=0", $Conn, 0, 1)   # ForwardOnly / ReadOnly
    foreach ($f in $rs.Fields) { $cols += $f.Name }
    $rs.Close()
    $cols
}

# ------------------------------------------------------------------------
# Kernfunktion: ausgewaehlte (nicht angekreuzte) Spalten verschleiern
# ------------------------------------------------------------------------
function Invoke-Obfuscation {
    param(
        [string]$Path,
        [string]$Table,
        [string[]]$ColumnsToObfuscate,     # die NICHT angekreuzten
        [scriptblock]$Log
    )

    $conn = New-AccessConnection -Path $Path
    try {
        $adOpenKeyset    = 1
        $adLockOptimistic = 3
        $rs = New-Object -ComObject ADODB.Recordset
        $rs.CursorLocation = 3   # adUseClient (stabileres Update)
        $rs.Open("SELECT * FROM [$Table]", $conn, $adOpenKeyset, $adLockOptimistic)

        # Feld-Metadaten cachen
        $meta = @{}
        foreach ($f in $rs.Fields) {
            $meta[$f.Name] = @{
                Type       = [int]$f.Type
                MaxLen     = [int]$f.DefinedSize
                Attributes = [int]$f.Attributes
            }
        }

        # Nur echte Autowert-/RowID-Spalten ausschliessen. Das Attribut
        # 'adFldUpdatable' wird von ACE oft als 'adFldUnknownUpdatable' (8)
        # gemeldet - dann NICHT vorab aussortieren, sondern das Schreiben
        # versuchen und nur bei echtem Fehler ueberspringen.
        $targets = @()
        foreach ($c in $ColumnsToObfuscate) {
            if (-not $meta.ContainsKey($c)) { continue }
            $attr = $meta[$c].Attributes
            if (($attr -band $adFldRowID) -ne 0) { & $Log "  '$c' uebersprungen (Autowert/RowID)."; continue }
            $targets += $c
        }

        if ($targets.Count -eq 0) { & $Log "Keine beschreibbaren Spalten zum Anonymisieren."; $rs.Close(); return 0 }

        & $Log ("Anonymisiere Spalten: " + ($targets -join ', '))

        $rowCount = 0
        while (-not $rs.EOF) {
            foreach ($c in $targets) {
                $field = $rs.Fields.Item($c)
                $new = Get-ObfuscatedValue -OriginalValue $field.Value -DataType $meta[$c].Type -MaxLen $meta[$c].MaxLen -ColumnName $c
                try { $field.Value = $new } catch { & $Log "  Zeile $rowCount, Spalte '$c': Wert konnte nicht gesetzt werden ($($_.Exception.Message))." }
            }
            $rs.Update()
            $rowCount++
            $rs.MoveNext()
        }
        $rs.Close()
        & $Log "$rowCount Datensaetze verarbeitet."
        return $rowCount
    }
    finally {
        try { $conn.Close() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($conn) | Out-Null
    }
}

# ========================================================================
#  GRAFISCHE OBERFLAECHE
# ========================================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Access-Datenbank anonymisieren / verschleiern"
$form.Size = New-Object System.Drawing.Size(640, 620)
$form.StartPosition = 'CenterScreen'
$form.MinimumSize = New-Object System.Drawing.Size(560, 560)

# --- Datei-Auswahl ---
$lblFile = New-Object System.Windows.Forms.Label
$lblFile.Text = "Access-Datenbank (.accdb / .mdb):"
$lblFile.Location = '15,15'; $lblFile.AutoSize = $true
$form.Controls.Add($lblFile)

$txtFile = New-Object System.Windows.Forms.TextBox
$txtFile.Location = '15,38'; $txtFile.Size = '480,23'
$txtFile.Anchor = 'Top,Left,Right'
$form.Controls.Add($txtFile)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = "Durchsuchen..."; $btnBrowse.Location = '505,37'; $btnBrowse.Size = '105,25'
$btnBrowse.Anchor = 'Top,Right'
$form.Controls.Add($btnBrowse)

# --- Tabellen-Auswahl ---
$lblTable = New-Object System.Windows.Forms.Label
$lblTable.Text = "Tabelle:"; $lblTable.Location = '15,74'; $lblTable.AutoSize = $true
$form.Controls.Add($lblTable)

$cmbTable = New-Object System.Windows.Forms.ComboBox
$cmbTable.Location = '15,96'; $cmbTable.Size = '480,23'
$cmbTable.DropDownStyle = 'DropDownList'; $cmbTable.Anchor = 'Top,Left,Right'
$form.Controls.Add($cmbTable)

$btnLoad = New-Object System.Windows.Forms.Button
$btnLoad.Text = "Laden"; $btnLoad.Location = '505,95'; $btnLoad.Size = '105,25'
$btnLoad.Anchor = 'Top,Right'
$form.Controls.Add($btnLoad)

# --- Hinweis ---
$lblHint = New-Object System.Windows.Forms.Label
$lblHint.Text = "Angekreuzt = bleibt UNVERAENDERT   |   NICHT angekreuzt = wird anonymisiert"
$lblHint.Location = '15,130'; $lblHint.AutoSize = $true
$lblHint.Font = New-Object System.Drawing.Font($lblHint.Font, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblHint)

# --- Spalten-Checkbox-Liste ---
$clbColumns = New-Object System.Windows.Forms.CheckedListBox
$clbColumns.Location = '15,155'; $clbColumns.Size = '595,220'
$clbColumns.CheckOnClick = $true
$clbColumns.Anchor = 'Top,Bottom,Left,Right'
$form.Controls.Add($clbColumns)

# --- Alle an/aus ---
$btnAll = New-Object System.Windows.Forms.Button
$btnAll.Text = "Alle ankreuzen"; $btnAll.Location = '15,382'; $btnAll.Size = '130,25'
$btnAll.Anchor = 'Bottom,Left'
$form.Controls.Add($btnAll)

$btnNone = New-Object System.Windows.Forms.Button
$btnNone.Text = "Alle abwaehlen"; $btnNone.Location = '150,382'; $btnNone.Size = '130,25'
$btnNone.Anchor = 'Bottom,Left'
$form.Controls.Add($btnNone)

# --- Backup-Option ---
$chkBackup = New-Object System.Windows.Forms.CheckBox
$chkBackup.Text = "Vor Aenderung Backup-Kopie erstellen"
$chkBackup.Location = '300,384'; $chkBackup.AutoSize = $true; $chkBackup.Checked = $true
$chkBackup.Anchor = 'Bottom,Right'
$form.Controls.Add($chkBackup)

# --- Start-Button ---
$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Verschleiern starten"; $btnRun.Location = '15,415'; $btnRun.Size = '595,32'
$btnRun.Anchor = 'Bottom,Left,Right'
$btnRun.BackColor = [System.Drawing.Color]::FromArgb(220,53,69)
$btnRun.ForeColor = [System.Drawing.Color]::White
$btnRun.Font = New-Object System.Drawing.Font($btnRun.Font, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnRun)

# --- Log ---
$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = '15,455'; $txtLog.Size = '595,110'
$txtLog.Multiline = $true; $txtLog.ScrollBars = 'Vertical'; $txtLog.ReadOnly = $true
$txtLog.Anchor = 'Bottom,Left,Right'
$form.Controls.Add($txtLog)

# --- Log-Helfer ---
$Log = {
    param([string]$Message)
    $ts = (Get-Date).ToString('HH:mm:ss')
    $txtLog.AppendText("[$ts] $Message`r`n")
    $txtLog.SelectionStart = $txtLog.Text.Length
    $txtLog.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

# ------------------------------------------------------------------------
# Event-Handler
# ------------------------------------------------------------------------
$btnBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "Access-Datenbanken (*.accdb;*.mdb)|*.accdb;*.mdb|Alle Dateien (*.*)|*.*"
    if ($dlg.ShowDialog() -eq 'OK') {
        $txtFile.Text = $dlg.FileName
        $cmbTable.Items.Clear()
        $clbColumns.Items.Clear()
    }
})

$btnLoad.Add_Click({
    $path = $txtFile.Text.Trim()
    if (-not (Test-Path -LiteralPath $path)) { & $Log "Datei nicht gefunden: $path"; return }

    $cmbTable.Items.Clear(); $clbColumns.Items.Clear()
    try {
        $conn = New-AccessConnection -Path $path
        try {
            $tables = Get-AccessTables -Conn $conn
            foreach ($t in $tables) { [void]$cmbTable.Items.Add($t) }
            & $Log "$($tables.Count) Tabelle(n) geladen."
            if ($cmbTable.Items.Count -gt 0) { $cmbTable.SelectedIndex = 0 }
        } finally { $conn.Close(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($conn) | Out-Null }
    } catch { & $Log "Fehler: $($_.Exception.Message)" }
})

$cmbTable.Add_SelectedIndexChanged({
    $path  = $txtFile.Text.Trim()
    $table = [string]$cmbTable.SelectedItem
    if (-not (Test-Path -LiteralPath $path) -or [string]::IsNullOrEmpty($table)) { return }

    $clbColumns.Items.Clear()
    try {
        $conn = New-AccessConnection -Path $path
        try {
            $cols = Get-AccessColumns -Conn $conn -Table $table
            foreach ($c in $cols) { [void]$clbColumns.Items.Add($c, $false) }  # standardmaessig NICHT angekreuzt
            & $Log "Tabelle '$table': $($cols.Count) Spalte(n) geladen."
        } finally { $conn.Close(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($conn) | Out-Null }
    } catch { & $Log "Fehler beim Laden der Spalten: $($_.Exception.Message)" }
})

$btnAll.Add_Click({  for ($i=0; $i -lt $clbColumns.Items.Count; $i++) { $clbColumns.SetItemChecked($i, $true) } })
$btnNone.Add_Click({ for ($i=0; $i -lt $clbColumns.Items.Count; $i++) { $clbColumns.SetItemChecked($i, $false) } })

$btnRun.Add_Click({
    $path  = $txtFile.Text.Trim()
    $table = [string]$cmbTable.SelectedItem
    if (-not (Test-Path -LiteralPath $path)) { & $Log "Bitte gueltige Datenbank waehlen."; return }
    if ([string]::IsNullOrEmpty($table))     { & $Log "Bitte eine Tabelle waehlen."; return }
    if ($clbColumns.Items.Count -eq 0)       { & $Log "Keine Spalten geladen."; return }

    # NICHT angekreuzte Spalten = zu anonymisieren
    $toObf = @()
    for ($i=0; $i -lt $clbColumns.Items.Count; $i++) {
        if (-not $clbColumns.GetItemChecked($i)) { $toObf += [string]$clbColumns.Items[$i] }
    }
    if ($toObf.Count -eq 0) { & $Log "Alle Spalten sind angekreuzt - es wird nichts veraendert."; return }

    $msg = "Folgende Spalten werden UNWIDERRUFLICH anonymisiert:`n`n" + ($toObf -join ", ") +
           "`n`nTabelle: $table`nFortfahren?"
    if ([System.Windows.Forms.MessageBox]::Show($msg, "Bestaetigung", 'YesNo', 'Warning') -ne 'Yes') {
        & $Log "Abgebrochen."; return
    }

    try {
        if ($chkBackup.Checked) {
            $bak = [System.IO.Path]::Combine(
                        [System.IO.Path]::GetDirectoryName($path),
                        [System.IO.Path]::GetFileNameWithoutExtension($path) +
                        "_backup_" + (Get-Date -Format 'yyyyMMdd_HHmmss') +
                        [System.IO.Path]::GetExtension($path))
            Copy-Item -LiteralPath $path -Destination $bak -Force
            & $Log "Backup erstellt: $bak"
        }

        $form.Cursor = 'WaitCursor'; $btnRun.Enabled = $false
        & $Log "Starte Anonymisierung..."
        $n = Invoke-Obfuscation -Path $path -Table $table -ColumnsToObfuscate $toObf -Log $Log
        & $Log "Fertig. $n Datensatz/-saetze aktualisiert."
        [System.Windows.Forms.MessageBox]::Show("Anonymisierung abgeschlossen.`n$n Datensaetze aktualisiert.", "Fertig", 'OK', 'Information') | Out-Null
    } catch {
        & $Log "FEHLER: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Fehler: $($_.Exception.Message)", "Fehler", 'OK', 'Error') | Out-Null
    } finally {
        $form.Cursor = 'Default'; $btnRun.Enabled = $true
    }
})

# Start
[void]$form.ShowDialog()
