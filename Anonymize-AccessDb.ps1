<#
.SYNOPSIS
    Liest eine Microsoft-Access-Datenbank (.accdb / .mdb) ODER eine Excel-Datei
    (.xlsx / .xlsm / .xlsb / .xls) und anonymisiert ("verschleiert") ausgewaehlte
    Spalten mit personenbezogenen Daten.

.DESCRIPTION
    Das Script oeffnet eine grafische Oberflaeche (WinForms):

      1. Access- oder Excel-Datei auswaehlen  (Datei-Dialog)
      2. Tabelle bzw. Arbeitsblatt auswaehlen  (Dropdown)
      3. Spalten werden in einer Checkbox-Liste angezeigt

    Logik der Auswahl:
      * ANGEKREUZT   = Spalte bleibt UNVERAENDERT  (z. B. Schluessel, Referenzen)
      * NICHT angekreuzt = Spalte wird VERSCHLEIERT / anonymisiert

    Die Anonymisierung erfolgt typ-abhaengig:
      * Text        -> je nach Inhalt E-Mail / Telefon / Name / generischer Text
      * Zahl        -> zufaelliger Wert in aehnlicher Groessenordnung
      * Datum       -> um zufaellige Tage verschoben (Access) bzw. verschleiert (Excel)
      * Ja/Nein     -> zufaelliger Boolescher Wert

    Zwei Engines:
      * Access -> ADODB/ACE-OLEDB, Schreiben ueber editierbaren Server-Cursor.
      * Excel  -> Excel-COM-Automation (Kopfzeile = erste Zeile = Spaltennamen).

    Vor jeder Aenderung wird automatisch eine Backup-Kopie der Datei angelegt
    (kann in der GUI deaktiviert werden).

.REQUIREMENTS
    * Windows mit Windows PowerShell 5.1 (WinForms / COM).
    * Fuer Access: "Microsoft Access Database Engine" (ACE OLEDB Provider).
      Wichtig: Die Bit-Version (32/64) des Providers muss zur PowerShell-
      Bit-Version passen. Fuer .accdb i. d. R. Microsoft.ACE.OLEDB.12.0
      oder .16.0. Fuer alte .mdb ggf. Microsoft.Jet.OLEDB.4.0 (nur 32-Bit).
    * Fuer Excel: installiertes Microsoft Excel (COM-Automation).

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

# Konsistente Ersetzung: gleicher Originalwert -> gleicher Ersatzwert
# (referentielle Zusammenhaenge bleiben erhalten). Gilt fuer Textwerte.
$script:ObfMap = @{}

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
$adText     = @(200,201,202,203,129,130,141)       # varchar/char/wchar/longtext
                                                   # (131 = adNumeric gehoert zu Decimal, NICHT hier)
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

# Kuerzt einen Text auf die (optionale) Maximallaenge.
function Limit-Length { param([string]$Value,[int]$MaxLen) if ($MaxLen -gt 0 -and $Value.Length -gt $MaxLen) { return $Value.Substring(0,$MaxLen) }; $Value }

# Erzeugt einen (rohen) Ersatzwert fuer einen Text - ohne Laengenbegrenzung
# und ohne Cache.
function Get-FakeForString {
    param([string]$OriginalValue, [string]$ColumnName)

    if (Test-LooksLikeEmail $OriginalValue) { return (New-FakeEmail) }
    if (Test-LooksLikePhone $OriginalValue) { return (New-FakePhone) }

    $lc = ([string]$ColumnName).ToLower()
    switch -Regex ($lc) {
        'vorname|firstname|first_name'     { return (Get-RandomFrom $script:Vornamen) }
        'nachname|lastname|last_name|name' { return (Get-RandomFrom $script:Nachnamen) }
        'strasse|street|adresse|address'   { return ("{0} {1}" -f (Get-RandomFrom $script:Strassen), $script:Rnd.Next(1,199)) }
        'ort|stadt|city'                   { return (Get-RandomFrom $script:Staedte) }
        'plz|zip|postal'                   { return ("{0:D5}" -f $script:Rnd.Next(1000,99999)) }
        default                            { return (New-FakeText -MaxLen 0) }
    }
}

# Anonymisiert einen Text-Wert (E-Mail / Telefon / Name / generisch).
# Wird sowohl vom Access- als auch vom Excel-Pfad verwendet.
# Konsistent: gleicher Originalwert -> immer gleicher Ersatzwert (Cache).
function Get-ObfuscatedString {
    param([string]$OriginalValue, [string]$ColumnName, [int]$MaxLen = 0)

    if (-not $script:ObfMap.ContainsKey($OriginalValue)) {
        $script:ObfMap[$OriginalValue] = Get-FakeForString -OriginalValue $OriginalValue -ColumnName $ColumnName
    }
    return (Limit-Length $script:ObfMap[$OriginalValue] $MaxLen)
}

# Liefert einen anonymisierten Wert anhand des .NET-Typs (fuer Excel-Zellen,
# wo keine OLE-DB-Typinformation vorliegt).
function Get-ObfuscatedValueGeneric {
    param($Value, [string]$ColumnName)

    if ($null -eq $Value) { return $Value }

    if ($Value -is [string]) {
        if (([string]$Value).Trim().Length -eq 0) { return $Value }
        return (Get-ObfuscatedString -OriginalValue ([string]$Value) -ColumnName $ColumnName -MaxLen 0)
    }
    if ($Value -is [datetime]) { return ([datetime]$Value).AddDays($script:Rnd.Next(-365,365)) }
    if ($Value -is [bool])     { return [bool]($script:Rnd.Next(0,2)) }
    if ($Value -is [double] -or $Value -is [single] -or $Value -is [decimal] -or
        $Value -is [int] -or $Value -is [int64] -or $Value -is [int16] -or $Value -is [byte]) {
        $d = [double]$Value
        $mag = [Math]::Max(10.0, [Math]::Abs($d))
        if ($d -eq [Math]::Floor($d)) {
            return [double]$script:Rnd.Next(0, [int][Math]::Min([int]::MaxValue, ($mag*2)+10))
        }
        return [Math]::Round(($script:Rnd.NextDouble() * $mag * 2), 2)
    }
    # Fallback: als Text behandeln
    return (Get-ObfuscatedString -OriginalValue ([string]$Value) -ColumnName $ColumnName -MaxLen 0)
}

# Liefert einen anonymisierten Wert passend zum Access-Feldtyp / Inhalt
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
        return (Get-ObfuscatedString -OriginalValue $strVal -ColumnName $ColumnName -MaxLen $MaxLen)
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

# Ermittelt ueber das DB-Schema, welche Spalten SCHREIBGESCHUETZT sind
# (Autowert / AutoNumber, berechnete Felder). Diese koennen von Access nicht
# aktualisiert werden und muessen von der Anonymisierung ausgenommen werden.
# Grundlage: COLUMN_FLAGS (DBCOLUMNFLAGS) aus adSchemaColumns.
function Get-ReadOnlyColumns {
    param($Conn, [string]$Table)

    $DBCOLUMNFLAGS_WRITE        = 0x00000004
    $DBCOLUMNFLAGS_WRITEUNKNOWN = 0x00000008
    $DBCOLUMNFLAGS_ISROWID      = 0x00000100   # Autowert / RowID

    $readonly = @{}
    $adSchemaColumns = 4
    try {
        # Restriktionen: [Katalog, Schema, Tabellenname, Spaltenname]
        $rest = @($null, $null, $Table, $null)
        $rs = $Conn.OpenSchema($adSchemaColumns, $rest)
        while (-not $rs.EOF) {
            $name  = [string]$rs.Fields.Item('COLUMN_NAME').Value
            $flags = 0
            try { $flags = [int64]$rs.Fields.Item('COLUMN_FLAGS').Value } catch {}
            $writable = ((($flags -band $DBCOLUMNFLAGS_WRITE) -ne 0) -or
                         (($flags -band $DBCOLUMNFLAGS_WRITEUNKNOWN) -ne 0)) -and
                        (($flags -band $DBCOLUMNFLAGS_ISROWID) -eq 0)
            if (-not $writable) { $readonly[$name] = $true }
            $rs.MoveNext()
        }
        $rs.Close()
    } catch {
        # Schema nicht verfuegbar -> leere Liste (Fallback auf Laufzeit-Erkennung)
    }
    $readonly
}

# Schreibt die anonymisierten Werte ueber einen editierbaren Server-Cursor.
# adUseServer + Update() schreibt jede Zeile sofort in die .accdb/.mdb.
function Invoke-ObfuscationByCursor {
    param($Conn, [string]$Table, [string[]]$Targets, $Meta, [scriptblock]$Log)

    $adOpenKeyset = 1; $adLockOptimistic = 3; $adUpdate = 0x01000000
    $rs = New-Object -ComObject ADODB.Recordset
    $rs.CursorLocation = 2   # adUseServer -> Update() schreibt sofort in die DB
    $rs.Open("SELECT * FROM [$Table]", $Conn, $adOpenKeyset, $adLockOptimistic)

    if (-not $rs.Supports($adUpdate)) {
        $rs.Close()
        & $Log "Tabelle ist nicht aktualisierbar (kein eindeutiger Index). Keine Aenderung moeglich."
        return 0
    }

    $failed = @{}; $rowCount = 0
    while (-not $rs.EOF) {
        $changed = $false
        foreach ($c in $Targets) {
            if ($failed.ContainsKey($c)) { continue }
            $field = $rs.Fields.Item($c)
            $new = Get-ObfuscatedValue -OriginalValue $field.Value -DataType $Meta[$c].Type -MaxLen $Meta[$c].MaxLen -ColumnName $c
            try { $field.Value = $new; $changed = $true }
            catch { $failed[$c] = $true; & $Log "  '$c' uebersprungen (Feld nicht aktualisierbar)." }
        }
        if ($changed) { try { $rs.Update() } catch { try { $rs.CancelUpdate() } catch {}; & $Log "  Datensatz uebersprungen (Update-Fehler)." } }
        $rowCount++
        $rs.MoveNext()
    }
    $rs.Close()
    & $Log "$rowCount Datensatz/-saetze verarbeitet."
    return $rowCount
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
        # --- Feld-Metadaten ueber ein leeres Recordset ermitteln ---
        $rsMeta = New-Object -ComObject ADODB.Recordset
        $rsMeta.Open("SELECT * FROM [$Table] WHERE 1=0", $conn, 0, 1)
        $meta = @{}
        foreach ($f in $rsMeta.Fields) {
            $meta[$f.Name] = @{
                Type       = [int]$f.Type
                MaxLen     = [int]$f.DefinedSize
                Attributes = [int]$f.Attributes
            }
        }
        $rsMeta.Close()

        # --- Schreibgeschuetzte Spalten (Autowert/berechnet) ausschliessen ---
        $readonly = Get-ReadOnlyColumns -Conn $conn -Table $Table
        $targets = @()
        foreach ($c in $ColumnsToObfuscate) {
            if (-not $meta.ContainsKey($c)) { continue }
            $attr = $meta[$c].Attributes
            if (($attr -band $adFldRowID) -ne 0) { & $Log "  '$c' uebersprungen (Autowert/RowID)."; continue }
            if ($readonly.ContainsKey($c))       { & $Log "  '$c' uebersprungen (nicht aktualisierbar, z.B. Autowert/berechnet)."; continue }
            $targets += $c
        }
        if ($targets.Count -eq 0) { & $Log "Keine beschreibbaren Spalten zum Anonymisieren."; return 0 }

        & $Log ("Anonymisiere Spalten: " + ($targets -join ', '))

        # Schreiben ueber einen editierbaren Server-Cursor (schreibt sofort in
        # die Datei). Bewusst OHNE ADODB.Command/Parameter, um COM-Typkonflikte
        # zu vermeiden.
        return (Invoke-ObfuscationByCursor -Conn $conn -Table $Table -Targets $targets -Meta $meta -Log $Log)
    }
    finally {
        try { $conn.Close() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($conn) | Out-Null
    }
}

# ========================================================================
#  EXCEL-ENGINE (COM-Automation)  -  benoetigt installiertes Microsoft Excel
# ========================================================================

# Bestimmt anhand der Dateiendung den Typ: 'Access', 'Excel' oder 'Unknown'.
function Get-FileKind {
    param([string]$Path)
    switch -Regex ([System.IO.Path]::GetExtension($Path).ToLower()) {
        '\.(accdb|mdb)$'            { return 'Access' }
        '\.(xlsx|xlsm|xlsb|xls)$'   { return 'Excel' }
        default                     { return 'Unknown' }
    }
}

# Startet eine unsichtbare Excel-Instanz.
function New-ExcelApp {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel
}

# Gibt COM-Objekte frei und beendet Excel sauber.
function Close-ExcelApp {
    param($Excel, $Workbook)
    if ($Workbook) { try { $Workbook.Close($false) } catch {} }
    if ($Excel)    { try { $Excel.Quit() } catch {} }
    if ($Workbook) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Workbook) } catch {} }
    if ($Excel)    { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel) } catch {} }
    [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
}

# Liefert die Namen aller Arbeitsblaetter einer Excel-Datei.
function Get-ExcelSheets {
    param([string]$Path)
    $sheets = @()
    $excel = New-ExcelApp; $wb = $null
    try {
        $wb = $excel.Workbooks.Open($Path, 0, $true)   # ReadOnly
        foreach ($ws in $wb.Worksheets) { $sheets += [string]$ws.Name }
    } finally { Close-ExcelApp -Excel $excel -Workbook $wb }
    $sheets
}

# Liefert die Spalten (Kopfzeile = erste Zeile) eines Arbeitsblattes.
# Rueckgabe: Liste von Objekten mit AbsoluteColumn (Spaltennummer) und Name.
function Get-ExcelColumns {
    param([string]$Path, [string]$Sheet)
    $cols = @()
    $excel = New-ExcelApp; $wb = $null
    try {
        $wb = $excel.Workbooks.Open($Path, 0, $true)   # ReadOnly
        $ws = $wb.Worksheets.Item($Sheet)
        $used = $ws.UsedRange
        $firstRow = [int]$used.Row
        $firstCol = [int]$used.Column
        $nCols    = [int]$used.Columns.Count
        for ($i = 0; $i -lt $nCols; $i++) {
            $absCol = $firstCol + $i
            $hdr = $ws.Cells.Item($firstRow, $absCol).Value2
            $name = if ($null -eq $hdr -or ([string]$hdr).Trim().Length -eq 0) { "Spalte $absCol" } else { [string]$hdr }
            $cols += [pscustomobject]@{ AbsoluteColumn = $absCol; Name = $name }
        }
    } finally { Close-ExcelApp -Excel $excel -Workbook $wb }
    $cols
}

# Anonymisiert die angegebenen Spalten eines Arbeitsblattes und speichert.
# $Columns = Liste von Objekten mit AbsoluteColumn und Name.
function Invoke-ObfuscationExcel {
    param([string]$Path, [string]$Sheet, $Columns, [scriptblock]$Log)

    $excel = New-ExcelApp; $wb = $null
    try {
        $wb = $excel.Workbooks.Open($Path)
        $ws = $wb.Worksheets.Item($Sheet)
        $used = $ws.UsedRange
        $firstRow  = [int]$used.Row
        $nRows     = [int]$used.Rows.Count
        $headerRow = $firstRow
        $dataStart = $firstRow + 1
        $lastRow   = $firstRow + $nRows - 1

        if ($lastRow -lt $dataStart) { & $Log "Arbeitsblatt '$Sheet' hat keine Datenzeilen."; return 0 }

        & $Log ("Anonymisiere Spalten: " + (($Columns | ForEach-Object { $_.Name }) -join ', '))

        foreach ($col in $Columns) {
            $absCol = [int]$col.AbsoluteColumn
            $rng = $ws.Range($ws.Cells.Item($dataStart, $absCol), $ws.Cells.Item($lastRow, $absCol))
            # .Value (statt .Value2) liefert Datumszellen als [datetime] und nicht
            # als serielle Zahl -> Datumsfelder werden korrekt als Datum behandelt.
            $vals = $rng.Value

            if ($vals -is [array]) {
                # 2D-Array [1..n, 1..1]
                for ($r = 1; $r -le $vals.GetLength(0); $r++) {
                    $vals[$r,1] = Get-ObfuscatedValueGeneric -Value $vals[$r,1] -ColumnName $col.Name
                }
                $rng.Value = $vals
            } else {
                # Einzelne Datenzelle -> Skalar
                $rng.Value = Get-ObfuscatedValueGeneric -Value $vals -ColumnName $col.Name
            }
            & $Log "  Spalte '$($col.Name)' anonymisiert."
        }

        $wb.Save()
        $rows = $lastRow - $dataStart + 1
        & $Log "$rows Datenzeile(n) verarbeitet."
        return $rows
    } finally {
        Close-ExcelApp -Excel $excel -Workbook $wb
    }
}

# ========================================================================
#  GRAFISCHE OBERFLAECHE
# ========================================================================
# Zwischen Handlern geteilte Zuordnung Excel-Spalten (Position -> Spaltennr.)
$script:ExcelCols = @()

$form = New-Object System.Windows.Forms.Form
$form.Text = "Access-/Excel-Daten anonymisieren / verschleiern"
$form.Size = New-Object System.Drawing.Size(640, 620)
$form.StartPosition = 'CenterScreen'
$form.MinimumSize = New-Object System.Drawing.Size(560, 560)

# --- Datei-Auswahl ---
$lblFile = New-Object System.Windows.Forms.Label
$lblFile.Text = "Access-Datenbank (.accdb / .mdb) oder Excel-Datei (.xlsx / .xls):"
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

# --- Tabellen-/Arbeitsblatt-Auswahl ---
$lblTable = New-Object System.Windows.Forms.Label
$lblTable.Text = "Tabelle / Arbeitsblatt:"; $lblTable.Location = '15,74'; $lblTable.AutoSize = $true
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
$chkBackup.Location = '410,380'; $chkBackup.AutoSize = $true; $chkBackup.Checked = $true
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
    $dlg.Filter = "Access/Excel (*.accdb;*.mdb;*.xlsx;*.xlsm;*.xlsb;*.xls)|*.accdb;*.mdb;*.xlsx;*.xlsm;*.xlsb;*.xls|" +
                  "Access-Datenbanken (*.accdb;*.mdb)|*.accdb;*.mdb|" +
                  "Excel-Dateien (*.xlsx;*.xlsm;*.xlsb;*.xls)|*.xlsx;*.xlsm;*.xlsb;*.xls|" +
                  "Alle Dateien (*.*)|*.*"
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
    $kind = Get-FileKind -Path $path
    try {
        if ($kind -eq 'Excel') {
            $sheets = Get-ExcelSheets -Path $path
            foreach ($s in $sheets) { [void]$cmbTable.Items.Add($s) }
            & $Log "Excel-Datei: $($sheets.Count) Arbeitsblatt/-blaetter geladen."
            if ($cmbTable.Items.Count -gt 0) { $cmbTable.SelectedIndex = 0 }
        }
        elseif ($kind -eq 'Access') {
            $conn = New-AccessConnection -Path $path
            try {
                $tables = Get-AccessTables -Conn $conn
                foreach ($t in $tables) { [void]$cmbTable.Items.Add($t) }
                & $Log "Access-Datenbank: $($tables.Count) Tabelle(n) geladen."
                if ($cmbTable.Items.Count -gt 0) { $cmbTable.SelectedIndex = 0 }
            } finally { $conn.Close(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($conn) | Out-Null }
        }
        else { & $Log "Unbekannter Dateityp. Bitte .accdb/.mdb oder .xlsx/.xls waehlen." }
    } catch { & $Log "Fehler: $($_.Exception.Message)" }
})

$cmbTable.Add_SelectedIndexChanged({
    $path  = $txtFile.Text.Trim()
    $table = [string]$cmbTable.SelectedItem
    if (-not (Test-Path -LiteralPath $path) -or [string]::IsNullOrEmpty($table)) { return }

    $clbColumns.Items.Clear()
    $script:ExcelCols = @()
    $kind = Get-FileKind -Path $path
    try {
        if ($kind -eq 'Excel') {
            $script:ExcelCols = @(Get-ExcelColumns -Path $path -Sheet $table)
            foreach ($c in $script:ExcelCols) { [void]$clbColumns.Items.Add($c.Name, $false) }
            & $Log "Arbeitsblatt '$table': $($script:ExcelCols.Count) Spalte(n) geladen."
        }
        elseif ($kind -eq 'Access') {
            $conn = New-AccessConnection -Path $path
            try {
                $cols     = Get-AccessColumns -Conn $conn -Table $table
                $readonly = Get-ReadOnlyColumns -Conn $conn -Table $table
                $roCount  = 0
                foreach ($c in $cols) {
                    # Schreibgeschuetzte Spalten (Autowert/berechnet) automatisch
                    # ankreuzen -> bleiben unveraendert; alle anderen offen lassen.
                    $isRo = $readonly.ContainsKey($c)
                    [void]$clbColumns.Items.Add($c, $isRo)
                    if ($isRo) { $roCount++ }
                }
                & $Log "Tabelle '$table': $($cols.Count) Spalte(n) geladen ($roCount schreibgeschuetzt, autom. angekreuzt)."
            } finally { $conn.Close(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($conn) | Out-Null }
        }
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

    $kind = Get-FileKind -Path $path

    # NICHT angekreuzte Eintraege = zu anonymisieren (Positionen merken)
    $uncheckedPos = @()
    for ($i=0; $i -lt $clbColumns.Items.Count; $i++) {
        if (-not $clbColumns.GetItemChecked($i)) { $uncheckedPos += $i }
    }
    if ($uncheckedPos.Count -eq 0) { & $Log "Alle Spalten sind angekreuzt - es wird nichts veraendert."; return }

    $namesToObf = $uncheckedPos | ForEach-Object { [string]$clbColumns.Items[$_] }
    $msg = "Folgende Spalten werden UNWIDERRUFLICH anonymisiert:`n`n" + ($namesToObf -join ", ") +
           "`n`nTabelle/Blatt: $table`nFortfahren?"
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
            & $Log "WARNUNG: Das Backup enthaelt weiterhin die ORIGINAL-Personendaten - bitte geschuetzt aufbewahren und nach Freigabe loeschen."
        }

        $form.Cursor = 'WaitCursor'; $btnRun.Enabled = $false
        & $Log "Starte Anonymisierung..."

        if ($kind -eq 'Excel') {
            # Positionen -> Excel-Spalten (Nummer + Name) ueber die gespeicherte Zuordnung
            $cols = $uncheckedPos | ForEach-Object { $script:ExcelCols[$_] }
            $n = Invoke-ObfuscationExcel -Path $path -Sheet $table -Columns $cols -Log $Log
            & $Log "Fertig. $n Datenzeile(n) verarbeitet."
            [System.Windows.Forms.MessageBox]::Show("Anonymisierung abgeschlossen.`n$n Datenzeilen verarbeitet.", "Fertig", 'OK', 'Information') | Out-Null
        }
        else {
            $n = Invoke-Obfuscation -Path $path -Table $table -ColumnsToObfuscate $namesToObf -Log $Log
            & $Log "Fertig. $n Datensatz/-saetze aktualisiert."
            [System.Windows.Forms.MessageBox]::Show("Anonymisierung abgeschlossen.`n$n Datensaetze aktualisiert.", "Fertig", 'OK', 'Information') | Out-Null
        }
    } catch {
        & $Log "FEHLER: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Fehler: $($_.Exception.Message)", "Fehler", 'OK', 'Error') | Out-Null
    } finally {
        $form.Cursor = 'Default'; $btnRun.Enabled = $true
    }
})

# Start
[void]$form.ShowDialog()
