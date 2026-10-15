<#
.SYNOPSIS
Compares two collection JSON files (*_Proactiva.json or *_DriverInfo.json) sheet by sheet.

.DESCRIPTION
Used to check that a change to the collection does not change its output: the sheet and
column names, the values and the row order of the raw data are a contract with the annex
checks (see the note at the top of src\collection\ProactivaCollector.psm1).

For every sheet it reports the row count of each side, column differences, rows that are only
on one side (regardless of order), the columns that changed for rows with the same key, and
whether the row order changed. Fields that change on every run (Timestamp, the date inside
"Backup File") are ignored; numbers are compared by value (16384 = 16384.0).

Live data that may genuinely change between two runs (free space, snapshot size, Tools
status...) shows up as differences too: review them by hand.

Exit code: 0 when both files are equivalent, 1 when there are differences.

.EXAMPLE
.\runtime\pwsh.exe -ExecutionPolicy Bypass -File .\src\tools\Compare-CollectionJson.ps1 -Reference .\resultados\2026-10-14\json\a_Proactiva.json -Candidate .\resultados\2026-10-14 (1)\json\b_Proactiva.json
#>
param(
    [Parameter(Mandatory = $true)] [string] $Reference,
    [Parameter(Mandatory = $true)] [string] $Candidate,
    [int] $MaxDifferences = 15
)

# Columns that identify a row, to show field-level differences between rows that changed
$keyColumnsBySheet = @{
    'vCenter'           = @('vCenter Server')
    'ESXi'              = @('vCenter', 'Hostname')
    'ESXi IO'           = @('vCenter', 'Hostname', 'Placa', 'Controlador')
    'VM'                = @('vCenter', 'VM', 'Host')
    'Datastores'        = @('vCenter', 'Hostname', 'Datastore')
    'Standard Switch'   = @('vCenter', 'ESXi', 'Switch', 'PortGroup')
    'VMkernel Adapters' = @('Host', 'Name')
    'Snapshot'          = @('vCenter', 'VM', 'Snapshot', 'Fecha')
    'Partitions'        = @('vCenter', 'VM', 'Disk')
    'Sizing'            = @('vCenter')
    'vDS'               = @('vCenter', 'Name')
    'vNetwork'          = @('vCenter', 'VM', 'Mac')
    'vLicense'          = @('vCenter', 'LicenseKey', 'Name')
    'Certficate'        = @('vCenter', 'Ubicacion', 'Subject', 'Valid Until')
    'PerformanceHealth' = @('vCenter', 'Cluster')
    'BackupActivity'    = @('vCenter', 'StartTime')
    'Falso Positivo'    = @('vCenter', 'Alarm Name')
}

$ignoredColumns = @('Timestamp')

function Get-Rows($report, [string] $sheet) {
    $value = $report.$sheet
    if ($null -eq $value) { return @() }
    return @($value)
}

# Value as a comparable string that keeps the JSON type (a boolean true is not the text "True")
function Format-Value($value) {
    if ($null -eq $value) { return 'null' }
    if ($value -is [bool]) { return "b:$value" }
    if ($value -is [datetime]) { return "d:" + $value.ToUniversalTime().ToString('o') }
    if ($value -is [int] -or $value -is [long] -or $value -is [double] -or $value -is [decimal]) {
        return "n:" + ([double]$value).ToString('R', [Globalization.CultureInfo]::InvariantCulture)
    }
    if ($value -is [array]) { return "[" + (($value | ForEach-Object { Format-Value $_ }) -join ' | ') + "]" }
    if ($value -is [PSCustomObject]) { return "o:" + ($value | ConvertTo-Json -Compress -Depth 10) }
    return "s:$value"
}

function Get-NormalizedRow($row, [string] $column) {
    $value = $row.$column
    if ($column -eq 'Backup File' -and $value -is [string]) {
        $value = $value -replace '_\d{4}-\d{2}-\d{2}_\d{6}(?=\.zip$)', '_<fecha>'
    }
    return Format-Value $value
}

function Get-Columns($rows) {
    if ($rows.Count -eq 0) { return @() }
    return @($rows[0].PSObject.Properties.Name)
}

function Get-RowText($row, [string[]] $columns) {
    return (($columns | Where-Object { $_ -notin $ignoredColumns } | ForEach-Object { "$_=$(Get-NormalizedRow $row $_)" }) -join '; ')
}

function Get-RowKey($row, [string[]] $keyColumns) {
    return (($keyColumns | ForEach-Object { [string]$row.$_ }) -join ' / ')
}

$referenceJson = Get-Content -Path $Reference -Raw | ConvertFrom-Json
$candidateJson = Get-Content -Path $Candidate -Raw | ConvertFrom-Json

Write-Host "Referencia: $Reference ($($referenceJson.DateTime), $($referenceJson.Endpoint))"
Write-Host "Candidato:  $Candidate ($($candidateJson.DateTime), $($candidateJson.Endpoint))"

$referenceSheets = @($referenceJson.Report.PSObject.Properties.Name)
$candidateSheets = @($candidateJson.Report.PSObject.Properties.Name)
$sheets = @($referenceSheets) + @($candidateSheets | Where-Object { $_ -notin $referenceSheets })

$differentSheets = 0

foreach ($sheet in $sheets) {
    $findings = [System.Collections.Generic.List[string]]::new()

    if ($sheet -notin $referenceSheets) { $findings.Add("La hoja solo existe en el candidato.") }
    if ($sheet -notin $candidateSheets) { $findings.Add("La hoja solo existe en la referencia.") }

    $referenceRows = Get-Rows $referenceJson.Report $sheet
    $candidateRows = Get-Rows $candidateJson.Report $sheet

    if ($referenceRows.Count -ne $candidateRows.Count) {
        $findings.Add("Filas: $($referenceRows.Count) en la referencia, $($candidateRows.Count) en el candidato.")
    }

    $referenceColumns = Get-Columns $referenceRows
    $candidateColumns = Get-Columns $candidateRows
    if ($referenceRows.Count -gt 0 -and $candidateRows.Count -gt 0 -and (($referenceColumns -join '|') -ne ($candidateColumns -join '|'))) {
        $findings.Add("Columnas distintas (u otro orden):`n      referencia: $($referenceColumns -join ', ')`n      candidato:  $($candidateColumns -join ', ')")
    }

    # Every column of both sides, so a missing column counts as a difference
    $columns = @($referenceColumns) + @($candidateColumns | Where-Object { $_ -notin $referenceColumns })

    $referenceTexts = @($referenceRows | ForEach-Object { Get-RowText $_ $columns })
    $candidateTexts = @($candidateRows | ForEach-Object { Get-RowText $_ $columns })

    # Multiset difference: rows only on one side, regardless of order
    $pending = @{}
    foreach ($text in $candidateTexts) { $pending[$text] = 1 + [int]$pending[$text] }
    $onlyReference = [System.Collections.Generic.List[int]]::new()
    for ($i = 0; $i -lt $referenceTexts.Count; $i++) {
        if ([int]$pending[$referenceTexts[$i]] -gt 0) { $pending[$referenceTexts[$i]]-- } else { $onlyReference.Add($i) }
    }
    $pending = @{}
    foreach ($text in $referenceTexts) { $pending[$text] = 1 + [int]$pending[$text] }
    $onlyCandidate = [System.Collections.Generic.List[int]]::new()
    for ($i = 0; $i -lt $candidateTexts.Count; $i++) {
        if ([int]$pending[$candidateTexts[$i]] -gt 0) { $pending[$candidateTexts[$i]]-- } else { $onlyCandidate.Add($i) }
    }

    if ($onlyReference.Count -gt 0 -or $onlyCandidate.Count -gt 0) {
        $keyColumns = if ($keyColumnsBySheet.ContainsKey($sheet)) { $keyColumnsBySheet[$sheet] } else { @($columns | Select-Object -First 2) }

        $candidateByKey = @{}
        foreach ($i in $onlyCandidate) {
            $key = Get-RowKey $candidateRows[$i] $keyColumns
            if (-not $candidateByKey.ContainsKey($key)) { $candidateByKey[$key] = [System.Collections.Generic.Queue[int]]::new() }
            $candidateByKey[$key].Enqueue($i)
        }

        $shown = 0
        foreach ($i in $onlyReference) {
            $key = Get-RowKey $referenceRows[$i] $keyColumns
            if ($candidateByKey.ContainsKey($key) -and $candidateByKey[$key].Count -gt 0) {
                $j = $candidateByKey[$key].Dequeue()
                $changes = foreach ($column in $columns) {
                    if ($column -in $ignoredColumns) { continue }
                    $old = Get-NormalizedRow $referenceRows[$i] $column
                    $new = Get-NormalizedRow $candidateRows[$j] $column
                    if ($old -ne $new) { "$column`: $old -> $new" }
                }
                if ($shown -lt $MaxDifferences) { $findings.Add("Cambio en [$key]: $($changes -join '; ')") }
            }
            elseif ($shown -lt $MaxDifferences) {
                $findings.Add("Solo en la referencia: $($referenceTexts[$i])")
            }
            $shown++
        }
        foreach ($queue in $candidateByKey.Values) {
            foreach ($j in $queue) {
                if ($shown -lt $MaxDifferences) { $findings.Add("Solo en el candidato: $($candidateTexts[$j])") }
                $shown++
            }
        }
        if ($shown -gt $MaxDifferences) { $findings.Add("... y $($shown - $MaxDifferences) diferencias mas.") }
    }
    elseif (($referenceTexts -join "`n") -ne ($candidateTexts -join "`n")) {
        $firstMismatch = 0
        while ($referenceTexts[$firstMismatch] -eq $candidateTexts[$firstMismatch]) { $firstMismatch++ }
        $findings.Add("Mismas filas pero en otro orden (primera diferencia en la fila $($firstMismatch + 1)).")
    }

    if ($findings.Count -eq 0) {
        Write-Host ("  OK  {0} ({1} filas)" -f $sheet, $referenceRows.Count) -ForegroundColor Green
    }
    else {
        $differentSheets++
        Write-Host ("  DIF {0}" -f $sheet) -ForegroundColor Yellow
        foreach ($finding in $findings) { Write-Host "      $finding" }
    }
}

Write-Host ""
if ($differentSheets -eq 0) {
    Write-Host "Sin diferencias." -ForegroundColor Green
    exit 0
}
Write-Host "$differentSheets hoja(s) con diferencias." -ForegroundColor Yellow
exit 1
