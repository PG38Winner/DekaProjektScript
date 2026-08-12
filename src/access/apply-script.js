/**
 * Das PowerShell-Skript, das den Aenderungsplan in die Access-Datenbank
 * schreibt.
 *
 * Warum als Zeichenkette und nicht als .ps1-Datei:
 *   - das eigenstaendige Buendel (dist/) besteht aus genau einer Datei; eine
 *     danebenliegende .ps1 waere dort nicht vorhanden;
 *   - uebergeben wird es ueber die Standardeingabe an "powershell -Command -".
 *     Das ist kein Skriptaufruf, also greift die Ausfuehrungsrichtlinie nicht
 *     und es braucht kein "Bypass".
 *
 * Uebergabe erfolgt ueber Umgebungsvariablen, damit keine Anfuehrungszeichen
 * in Pfaden zu Problemen fuehren:
 *   ANONYMIZER_DB    - Pfad zur Datenbank
 *   ANONYMIZER_PLAN  - Pfad zur Plandatei (JSON, UTF-8)
 *
 * Geschrieben wird ueber einen editierbaren Server-Cursor. Diesen Weg hatte
 * bereits das fruehere PowerShell-Skript des Projekts erarbeitet (Commit
 * 8afde11) - er schreibt sofort in die Datei und umgeht die Typprobleme, die
 * der Weg ueber ADODB.Command mit sich brachte (b0441b5).
 *
 * Hinweis: Backticks sind hier bewusst vermieden, damit der Text in ein
 * JavaScript-Template passt.
 */

export const APPLY_SCRIPT = String.raw`
$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.Encoding]::UTF8

$dbPath   = $env:ANONYMIZER_DB
$planPath = $env:ANONYMIZER_PLAN
$invariant = [System.Globalization.CultureInfo]::InvariantCulture

function Format-Key($value) {
    if ($null -eq $value) { return '' }
    if ($value -is [datetime]) { return $value.ToString('o', $invariant) }
    if ($value -is [double] -or $value -is [single] -or $value -is [decimal]) {
        return [string]::Format($invariant, '{0}', $value)
    }
    return [string]$value
}

function Convert-Value($raw, [string]$type) {
    if ($null -eq $raw) { return $null }
    switch ($type) {
        'DateTime'         { return [datetime]::Parse([string]$raw, $invariant) }
        'DateTimeExtended' { return [datetime]::Parse([string]$raw, $invariant) }
        'Boolean'          { return [bool]$raw }
        'Byte'             { return [byte]$raw }
        'Integer'          { return [int32]$raw }
        'Long'             { return [int32]$raw }
        'BigInt'           { return [int64]$raw }
        'Currency'         { return [decimal]::Parse([string]$raw, $invariant) }
        'Numeric'          { return [decimal]::Parse([string]$raw, $invariant) }
        'Float'            { return [double]::Parse([string]$raw, $invariant) }
        'Double'           { return [double]::Parse([string]$raw, $invariant) }
        default            { return [string]$raw }
    }
}

function Open-Database([string]$path) {
    $providers = @('Microsoft.ACE.OLEDB.16.0', 'Microsoft.ACE.OLEDB.12.0')
    if ([System.IO.Path]::GetExtension($path).ToLower() -eq '.mdb') {
        $providers += 'Microsoft.Jet.OLEDB.4.0'
    }

    $problems = @()
    foreach ($provider in $providers) {
        try {
            $conn = New-Object -ComObject ADODB.Connection
            $conn.Open("Provider=$provider;Data Source=$path;")
            return $conn
        } catch {
            $problems += "$provider : $($_.Exception.Message)"
        }
    }

    $nl = [Environment]::NewLine
    throw ("Keine Verbindung zur Datenbank moeglich. Die Microsoft Access Database Engine " +
           "muss installiert sein und ihre Bit-Version (32/64) zu der von PowerShell passen." +
           $nl + ($problems -join $nl))
}

$plan = Get-Content -LiteralPath $planPath -Raw -Encoding UTF8 | ConvertFrom-Json
$conn = Open-Database $dbPath

$result = @{ tables = @(); errors = @() }

try {
    foreach ($table in $plan.tables) {
        $byKey = @{}
        foreach ($update in $table.updates) { $byKey[(Format-Key $update.k)] = $update.v }

        $adOpenKeyset = 1; $adLockOptimistic = 3; $adUpdate = 0x01000000
        $rs = New-Object -ComObject ADODB.Recordset
        $rs.CursorLocation = 2   # adUseServer - Update() schreibt sofort in die Datei
        $rs.Open("SELECT * FROM [$($table.name)]", $conn, $adOpenKeyset, $adLockOptimistic)

        if (-not $rs.Supports($adUpdate)) {
            $rs.Close()
            $result.errors += "Tabelle '$($table.name)' ist nicht aktualisierbar (kein eindeutiger Index)."
            continue
        }

        $updated = 0; $missed = 0; $failedColumns = @{}

        while (-not $rs.EOF) {
            $key = Format-Key $rs.Fields.Item($table.key).Value
            $values = $byKey[$key]

            if ($null -ne $values) {
                $changed = $false
                foreach ($property in $values.PSObject.Properties) {
                    if ($failedColumns.ContainsKey($property.Name)) { continue }
                    $type = $table.types.($property.Name)
                    try {
                        $rs.Fields.Item($property.Name).Value = Convert-Value $property.Value $type
                        $changed = $true
                    } catch {
                        $failedColumns[$property.Name] = $true
                        $result.errors += "Spalte '$($property.Name)' in '$($table.name)' nicht beschreibbar: $($_.Exception.Message)"
                    }
                }
                if ($changed) {
                    try { $rs.Update(); $updated++ }
                    catch {
                        try { $rs.CancelUpdate() } catch {}
                        $result.errors += "Datensatz $key in '$($table.name)' konnte nicht gespeichert werden: $($_.Exception.Message)"
                    }
                }
                $byKey.Remove($key)
            } else {
                $missed++
            }

            $rs.MoveNext()
        }

        $rs.Close()
        $result.tables += @{ name = $table.name; updated = $updated; unmatched = $byKey.Count }
    }
} finally {
    try { $conn.Close() } catch {}
}

$result | ConvertTo-Json -Depth 6 -Compress
`;
