# Step 3 of the pipeline: run the checks for the visit frequency over the raw Excel and build
# the two deliverables: "Anexo Tecnico - <client> - <month>.xlsx" (one detail sheet per check
# with findings) and "Checklist Proactiva - <client> - <month>.xlsx" (result per checklist ID).

Import-Module (Join-Path $PSScriptRoot '..\shared\ExcelStyle.psm1')

. (Join-Path $PSScriptRoot 'checks\MonthlyChecks.ps1')
. (Join-Path $PSScriptRoot 'checks\QuarterlyChecks.ps1')
. (Join-Path $PSScriptRoot 'checks\SemiannualChecks.ps1')

# Checks per frequency, in the order their sheets appear in the annex. Frequencies are
# cumulative: a quarterly visit also runs the monthly checks, a semiannual one runs all three.
$script:checksByFrequency = [ordered]@{
    Monthly    = @(
        'Test-VCenterSizing', 'Test-PartitionFreeSpace', 'Test-SyslogConfiguration', 'Test-MultipathPolicy',
        'Test-EsxiVersionConsistency', 'Test-HostResourceConsistency', 'Test-VmNicAdapterType', 'Test-EsxiShellTimeout',
        'Test-PowerManagementPolicy', 'Test-VmwareToolsStatus', 'Test-ConnectedIsos', 'Test-VmNicStartConnected',
        'Test-Snapshots', 'Test-AlarmConfiguration'
    )
    Quarterly  = @(
        'Test-EsxiEndOfSupport', 'Test-HardwareCompatibility', 'Test-NtpConfiguration', 'Test-HostDrivers',
        'Test-DnsConfiguration', 'Test-Licenses', 'Test-NiocEnabled', 'Test-StandardSwitchConsistency',
        'Test-VCenterRootPassword', 'Test-VCenterCertificates', 'Test-EsxiCertificates', 'Test-PerformanceStatistics'
    )
    Semiannual = @(
        'Test-VamiBackup', 'Test-VdsBackup'
    )
}

function Get-ChecksForFrequency {
    param([Parameter(Mandatory = $true)] [ValidateSet('Monthly', 'Quarterly', 'Semiannual')] [string] $Frequency)

    $checks = @()
    foreach ($level in $script:checksByFrequency.Keys) {
        $checks += $script:checksByFrequency[$level]
        if ($level -eq $Frequency) { break }
    }
    return $checks
}

# Generates the annex and the checklist. Returns @{ AnnexPath; ChecklistPath }.
function New-ProactivaReport {
    param(
        [Parameter(Mandatory = $true)] [string] $RunFolder,
        [Parameter(Mandatory = $true)] [string] $RawExcelPath,
        [string] $DriverExcelPath,
        [Parameter(Mandatory = $true)] [string] $Client,
        [Parameter(Mandatory = $true)] [string] $Month,
        [Parameter(Mandatory = $true)] [ValidateSet('Monthly', 'Quarterly', 'Semiannual')] [string] $Frequency
    )

    $annexFolder = Join-Path -Path $RunFolder -ChildPath "anexo"
    if (-not (Test-Path $annexFolder)) { New-Item -ItemType Directory -Path $annexFolder -Force | Out-Null }

    $context = [PSCustomObject]@{
        RunFolder           = $RunFolder
        RawExcelPath        = $RawExcelPath
        DriverExcelPath     = $DriverExcelPath
        AnnexPath           = Join-Path -Path $annexFolder -ChildPath "Anexo Tecnico - $Client - $Month.xlsx"
        ChecklistPath       = Join-Path -Path $annexFolder -ChildPath "Checklist Proactiva - $Client - $Month.xlsx"
        ReferenceDataFolder = Join-Path -Path $PSScriptRoot -ChildPath "..\collection\reference-data"
    }

    # Regenerating for the same client/month starts from a clean annex; otherwise Export-Excel
    # would reopen the old file and add to it
    if (Test-Path $context.AnnexPath) { Remove-Item $context.AnnexPath -Force }

    $results = @()
    foreach ($check in (Get-ChecksForFrequency -Frequency $Frequency)) {
        Write-Host "  -> $check" -ForegroundColor DarkGray
        $results += & $check $context
    }

    Update-Checklist -TemplatePath (Join-Path $PSScriptRoot 'templates\checklist-template.xlsx') -OutputPath $context.ChecklistPath -Results $results

    return [PSCustomObject]@{
        AnnexPath     = if (Test-Path $context.AnnexPath) { $context.AnnexPath } else { $null }
        ChecklistPath = $context.ChecklistPath
    }
}

# Copies the template and fills Estado (G) and Detalle (H) on the row of each result's ID
function Update-Checklist {
    param(
        [Parameter(Mandatory = $true)] [string] $TemplatePath,
        [Parameter(Mandatory = $true)] [string] $OutputPath,
        [Parameter(Mandatory = $true)] [object[]] $Results
    )

    Copy-Item -Path $TemplatePath -Destination $OutputPath -Force

    $package = Open-ExcelPackage -Path $OutputPath
    try {
        $sheet = $package.Workbook.Worksheets["Checklist"]

        # Rows that exist only for this client (e.g. VSP-TR-02-NX) are created first, so the
        # lookup by ID below finds them like any template row
        foreach ($result in ($Results | Where-Object { $_.PSObject.Properties['InsertAfter'] })) {
            Add-ChecklistRow -Sheet $sheet -AfterId $result.InsertAfter -Result $result
        }

        foreach ($result in $Results) {
            $idCell = $sheet.Cells["A:A"] | Where-Object { $_.Text -eq $result.ID }
            if (-not $idCell) {
                Write-Warning "El ID '$($result.ID)' no existe en la plantilla de checklist."
                continue
            }

            $row = $idCell.Start.Row
            $sheet.Cells["G$row"].Value = $result.Resultado
            $sheet.Cells["H$row"].Value = $result.Detalle
            Write-Host "Checklist '$($result.ID)': $($result.Resultado)" -ForegroundColor Green
        }
    } finally {
        Close-ExcelPackage $package
    }
}

# Inserts a row right below the row of $AfterId, copying that row's look and Estado dropdown,
# and fills ID / Procedimiento / Chequeo from $Result. EPPlus does not move the dropdowns
# (data validations) when inserting a row, so they are rebuilt shifted by one row.
function Add-ChecklistRow {
    param(
        [Parameter(Mandatory = $true)] $Sheet,
        [Parameter(Mandatory = $true)] [string] $AfterId,
        [Parameter(Mandatory = $true)] $Result
    )

    $afterCell = $Sheet.Cells["A:A"] | Where-Object { $_.Text -eq $AfterId }
    if (-not $afterCell) {
        Write-Warning "No se encontro '$AfterId' en la plantilla para insertar la fila '$($Result.ID)'."
        return
    }
    $source = $afterCell.Start.Row
    $new = $source + 1

    # Dropdowns before the insertion: [address, list formula]
    $validations = @($Sheet.DataValidations | ForEach-Object {
        [PSCustomObject]@{ Address = $_.Address.Address; Formula = $_.Formula.ExcelFormula }
    })
    $sourceFormula = ($validations | Where-Object {
        $_.Address -split ' ' | Where-Object { Test-CellInRange -Range $_ -Column 'G' -Row $source }
    } | Select-Object -First 1).Formula

    $Sheet.DataValidations.Clear()
    $Sheet.InsertRow($new, 1)

    foreach ($validation in $validations) {
        $shifted = @($validation.Address -split ' ' | ForEach-Object { Move-RangeForInsertedRow -Range $_ -InsertedRow $new }) -join ' '
        $list = $Sheet.DataValidations.AddListValidation($shifted)
        $list.Formula.ExcelFormula = $validation.Formula
        $list.AllowBlank = $true
        $list.ShowErrorMessage = $true
    }
    if ($sourceFormula) {
        $list = $Sheet.DataValidations.AddListValidation("G$new")
        $list.Formula.ExcelFormula = $sourceFormula
        $list.AllowBlank = $true
        $list.ShowErrorMessage = $true
    }

    # Same look as the source row, then its own content
    $Sheet.Cells["A${source}:I${source}"].Copy($Sheet.Cells["A${new}:I${new}"])
    $Sheet.Row($new).Height = $Sheet.Row($source).Height
    $Sheet.Cells["A$new"].Value = $Result.ID
    $Sheet.Cells["D$new"].Value = $Result.Procedimiento
    $Sheet.Cells["E$new"].Value = $Result.Chequeo
    foreach ($column in 'G', 'H', 'I') { $Sheet.Cells["$column$new"].Value = $null }
}

# Does a validation range such as "G14:G16" or "G18" cover the given cell?
function Test-CellInRange {
    param([string] $Range, [string] $Column, [int] $Row)

    if ($Range -notmatch '^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$') { return $false }
    $firstColumn = $matches[1]; $firstRow = [int]$matches[2]
    $lastColumn = if ($matches[3]) { $matches[3] } else { $firstColumn }
    $lastRow = if ($matches[4]) { [int]$matches[4] } else { $firstRow }
    return ($Column -ge $firstColumn -and $Column -le $lastColumn -and $Row -ge $firstRow -and $Row -le $lastRow)
}

# A validation range after a row was inserted at $InsertedRow: rows at or below it move down
# one; a range that spans the new row is split around it (the new row gets its own dropdown).
function Move-RangeForInsertedRow {
    param([string] $Range, [int] $InsertedRow)

    if ($Range -notmatch '^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$') { return $Range }
    $firstColumn = $matches[1]; $firstRow = [int]$matches[2]
    $lastColumn = if ($matches[3]) { $matches[3] } else { $firstColumn }
    $lastRow = if ($matches[4]) { [int]$matches[4] } else { $firstRow }

    $format = { param($c1, $r1, $c2, $r2) if ($c1 -eq $c2 -and $r1 -eq $r2) { "$c1$r1" } else { "${c1}${r1}:${c2}${r2}" } }

    if ($lastRow -lt $InsertedRow) { return $Range }
    if ($firstRow -ge $InsertedRow) { return & $format $firstColumn ($firstRow + 1) $lastColumn ($lastRow + 1) }
    return (& $format $firstColumn $firstRow $lastColumn ($InsertedRow - 1)), (& $format $firstColumn ($InsertedRow + 1) $lastColumn ($lastRow + 1)) -join ' '
}

# ---------------------------------------------------------------------------------------
# Helpers used by the checks
# ---------------------------------------------------------------------------------------

# Rows of a raw Excel sheet, or nothing when the file or the sheet does not exist (the
# conversion skips empty sheets, e.g. "vDS" when there are no distributed switches)
function Import-RawSheet {
    param(
        [string] $Path,
        [Parameter(Mandatory = $true)] [string] $WorksheetName
    )

    if (-not $Path -or -not (Test-Path $Path)) { return }
    try {
        Import-Excel -Path $Path -WorksheetName $WorksheetName -ErrorAction Stop
    } catch {}
}

# Appends a sheet to the annex with the Wetcom header style
function Export-AnnexSheet {
    param(
        [Parameter(Mandatory = $true)] [object[]] $Data,
        [Parameter(Mandatory = $true)] [string] $Path,
        [Parameter(Mandatory = $true)] [string] $WorksheetName,
        [string[]] $NoNumberConversion
    )

    $exportParams = @{
        Path          = $Path
        WorksheetName = $WorksheetName
        AutoSize      = $true
        FreezeTopRow  = $true
        PassThru      = $true
    }
    if ($NoNumberConversion) { $exportParams['NoNumberConversion'] = $NoNumberConversion }

    $package = $Data | Export-Excel @exportParams

    # Excel limits sheet names to 31 characters
    $worksheet = $package.Workbook.Worksheets[$WorksheetName.Substring(0, [Math]::Min(31, $WorksheetName.Length))]
    if ($null -ne $worksheet -and $null -ne $worksheet.Dimension) {
        Set-HeaderStyle -Worksheet $worksheet
    } else {
        Write-Warning "No se pudo aplicar estilo a la hoja '$WorksheetName' porque no se encontro o estaba vacia."
    }

    $package.Save()
    $package.Dispose()
}

# ESXi hosts grouped by "vCenter|Cluster" (hosts without cluster are left out).
# Returns a hashtable; iterate its .Values to get each cluster's hosts.
function Get-HostsByCluster {
    param([object[]] $Hosts)

    $groups = @{}
    foreach ($row in $Hosts) {
        if ($null -ne $row -and -not [string]::IsNullOrEmpty($row.Cluster) -and -not [string]::IsNullOrEmpty($row.vCenter)) {
            $key = "$($row.vCenter.Trim())|$($row.Cluster.Trim())"
            if (-not $groups.ContainsKey($key)) { $groups[$key] = @() }
            $groups[$key] += $row
        }
    }
    return $groups
}

# Build -> update release ("ESXi 7.0 U3") lookup, plus every release the IO HCL knows about
function Get-EsxiReleaseCatalog {
    param([Parameter(Mandatory = $true)] [string] $ReferenceDataFolder)

    # The update level comes from releaseLevel ("VMware ESXi 8.0.0 Update 3"): for 8.0 builds
    # updateRelease no longer includes it ("ESXi 8.0"), while the IO HCL does ("ESXi 8.0 U3")
    $releasesByBuild = @{}
    foreach ($release in (Get-Content (Join-Path $ReferenceDataFolder "esxiReleases.json") | ConvertFrom-Json).data.esxiReleases) {
        $releasesByBuild[[string]$release.build] = if ($release.releaseLevel -match 'ESXi (\d+\.\d+)\.\d+ Update (\d+)') {
            "ESXi $($matches[1]) U$($matches[2])"
        } else {
            $release.updateRelease
        }
    }

    $ioHclReleases = (Get-Content (Join-Path $ReferenceDataFolder "vmware-iohcl.json") | ConvertFrom-Json).data.ioDevices.releases | Select-Object -Unique

    return [PSCustomObject]@{
        ReleasesByBuild = $releasesByBuild
        IoHclReleases   = $ioHclReleases
    }
}

# "VMware ESXi 8.0.3 build-24677879" -> "ESXi 8.0 U3". Uses the build list first; for 7.0
# and later the third version digit is the update number, so builds missing from the
# (possibly stale) list can still be resolved.
function Resolve-EsxiUpdateRelease {
    param(
        [string] $FullName,
        [Parameter(Mandatory = $true)] $Catalog
    )

    if ($FullName -match 'build-(\d+)' -and $Catalog.ReleasesByBuild.ContainsKey($matches[1])) {
        return $Catalog.ReleasesByBuild[$matches[1]]
    }

    if ($FullName -match '(\d+)\.(\d+)\.(\d+)' -and [int]$matches[1] -ge 7) {
        $release = "ESXi $($matches[1]).$($matches[2])"
        if ([int]$matches[3] -gt 0) { $release += " U$($matches[3])" }
        return $release
    }

    return $null
}

Export-ModuleMember -Function New-ProactivaReport
