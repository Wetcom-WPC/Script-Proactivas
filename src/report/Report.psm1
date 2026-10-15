# Step 3 of the pipeline: run the checks for the visit frequency over the raw Excel and build
# the two deliverables: "Anexo Tecnico - <client> - <month>.xlsx" (one detail sheet per check
# with findings) and "Checklist Proactiva - <client> - <month>.xlsx" (result per checklist ID).

Import-Module (Join-Path $PSScriptRoot '..\shared\ExcelStyle.psm1')

. (Join-Path $PSScriptRoot 'checks\MonthlyChecks.ps1')
. (Join-Path $PSScriptRoot 'checks\QuarterlyChecks.ps1')
. (Join-Path $PSScriptRoot 'checks\SemiannualChecks.ps1')

# Checks per frequency, in the order their sheets appear in the annex, each with the checklist
# IDs it fills (used to flag those rows when the check itself fails). Frequencies are
# cumulative: a quarterly visit also runs the monthly checks, a semiannual one runs all three.
$script:checksByFrequency = [ordered]@{
    Monthly    = [ordered]@{
        'Test-VCenterSizing'           = @('VSP-ME-01')
        'Test-PartitionFreeSpace'      = @('VSP-ME-02')
        'Test-SyslogConfiguration'     = @('VSP-ME-04')
        'Test-MultipathPolicy'         = @('VSP-ME-05')
        'Test-EsxiVersionConsistency'  = @('VSP-ME-06')
        'Test-HostResourceConsistency' = @('VSP-ME-07')
        'Test-VmNicAdapterType'        = @('VSP-ME-08')
        'Test-EsxiShellTimeout'        = @('VSP-ME-12')
        'Test-PowerManagementPolicy'   = @('VSP-ME-14')
        'Test-VmwareToolsStatus'       = @('VSP-ME-15')
        'Test-ConnectedIsos'           = @('VSP-ME-16')
        'Test-VmNicStartConnected'     = @('VSP-ME-17')
        'Test-Snapshots'               = @('VSP-ME-03', 'VSP-ME-18')
        'Test-AlarmConfiguration'      = @('VSP-ME-13')
    }
    Quarterly  = [ordered]@{
        'Test-EsxiEndOfSupport'            = @('VSP-TR-01')
        'Test-HardwareCompatibility'       = @('VSP-TR-02')
        'Test-NtpConfiguration'            = @('VSP-TR-03')
        'Test-HostDrivers'                 = @('VSP-TR-04')
        'Test-DnsConfiguration'            = @('VSP-TR-05')
        'Test-Licenses'                    = @('VSP-TR-06')
        'Test-NiocEnabled'                 = @('VSP-TR-11')
        'Test-StandardSwitchConsistency'   = @('VSP-TR-12')
        'Test-VCenterRootPassword'         = @('VSP-TR-09')
        'Test-VCenterCertificates'         = @('VSP-TR-07')
        'Test-EsxiCertificates'            = @('VSP-TR-08')
        'Test-PerformanceStatistics'       = @('VSP-TR-10')
    }
    Semiannual = [ordered]@{
        'Test-VamiBackup' = @('VSP-SE-02')
        'Test-VdsBackup'  = @('VSP-SE-01')
    }
}

# Check name -> checklist IDs, for the checks that belong to the frequency
function Get-ChecksForFrequency {
    param([Parameter(Mandatory = $true)] [ValidateSet('Monthly', 'Quarterly', 'Semiannual')] [string] $Frequency)

    $checks = [ordered]@{}
    foreach ($level in $script:checksByFrequency.Keys) {
        foreach ($name in $script:checksByFrequency[$level].Keys) { $checks[$name] = $script:checksByFrequency[$level][$name] }
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

    # One failing check must not take the whole report down: it is logged, its checklist rows
    # are left blank with a note to review them by hand, and the other checks go on
    $results = @()
    $failedChecks = @()
    $checks = Get-ChecksForFrequency -Frequency $Frequency
    foreach ($check in $checks.Keys) {
        Write-Host "  -> $check" -ForegroundColor DarkGray
        try {
            $results += & $check $context
        }
        catch {
            $failedChecks += $check
            Write-Warning "El chequeo '$check' fallo y se omite: $($_.Exception.Message) (linea $($_.InvocationInfo.ScriptLineNumber) de $(Split-Path $_.InvocationInfo.ScriptName -Leaf))"
            foreach ($id in $checks[$check]) {
                $results += [PSCustomObject]@{ ID = $id; Resultado = $null; Detalle = "El chequeo automatico fallo ($($_.Exception.Message)); revisar manualmente. Detalle en el log." }
            }
        }
    }
    if ($failedChecks.Count -gt 0) {
        Write-Host "ATENCION: $($failedChecks.Count) chequeos fallaron y quedaron en blanco en la checklist: $($failedChecks -join ', ')" -ForegroundColor Yellow
    }

    Update-Checklist -TemplatePath (Join-Path $PSScriptRoot 'templates\checklist-template.xlsx') -OutputPath $context.ChecklistPath -Results $results -Frequency $Frequency

    return [PSCustomObject]@{
        AnnexPath     = if (Test-Path $context.AnnexPath) { $context.AnnexPath } else { $null }
        ChecklistPath = $context.ChecklistPath
    }
}

# Recurrences (column C "Recurrencia" of the checklist) that belong to each visit. Rows with
# any other known recurrence are removed, so a monthly checklist only has monthly tasks.
$script:recurrencesByFrequency = @{
    Monthly    = @('Mensual')
    Quarterly  = @('Mensual', 'Trimestral')
    Semiannual = @('Mensual', 'Trimestral', 'Semestral')
}

# Copies the template, keeps only the rows of this visit's frequency, adds the client-specific
# rows (e.g. VSP-TR-02-NX) and fills Estado (G) and Detalle (H) on the row of each result's ID
function Update-Checklist {
    param(
        [Parameter(Mandatory = $true)] [string] $TemplatePath,
        [Parameter(Mandatory = $true)] [string] $OutputPath,
        [Parameter(Mandatory = $true)] [object[]] $Results,
        [Parameter(Mandatory = $true)] [ValidateSet('Monthly', 'Quarterly', 'Semiannual')] [string] $Frequency
    )

    Copy-Item -Path $TemplatePath -Destination $OutputPath -Force

    $package = Open-ExcelPackage -Path $OutputPath
    try {
        $sheet = $package.Workbook.Worksheets["Checklist"]

        # EPPlus does not move the dropdowns (data validations) when rows are deleted or
        # inserted: they are taken out, followed cell by cell while rows change, and put back
        $dropdowns = Get-ChecklistDropdowns -Sheet $sheet
        # One by one: DataValidations.Clear() empties the list but leaves them in the file,
        # which would end up with every dropdown twice
        foreach ($validation in @($sheet.DataValidations)) { [void]$sheet.DataValidations.Remove($validation) }

        Remove-ChecklistRowsOutsideVisit -Sheet $sheet -Frequency $Frequency -Dropdowns $dropdowns

        # Rows that exist only for this client are created before filling, so the lookup by
        # ID below finds them like any template row
        foreach ($result in ($Results | Where-Object { $_.PSObject.Properties['InsertAfter'] })) {
            Add-ChecklistRow -Sheet $sheet -AfterId $result.InsertAfter -Result $result -Dropdowns $dropdowns
        }

        Restore-ChecklistDropdowns -Sheet $sheet -Dropdowns $dropdowns

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

# Deletes the rows whose recurrence does not belong to the visit (bottom-up, so the row
# numbers still to visit do not move). Rows with an unknown recurrence are kept.
function Remove-ChecklistRowsOutsideVisit {
    param(
        [Parameter(Mandatory = $true)] $Sheet,
        [Parameter(Mandatory = $true)] [string] $Frequency,
        [Parameter(Mandatory = $true)] [System.Collections.Generic.List[object]] $Dropdowns
    )

    $allowed = $script:recurrencesByFrequency[$Frequency]
    $known = $script:recurrencesByFrequency['Semiannual']
    $removed = 0

    for ($row = $Sheet.Dimension.End.Row; $row -ge 2; $row--) {
        if (-not $Sheet.Cells["A$row"].Text) { continue }
        $recurrence = $Sheet.Cells["C$row"].Text.Trim()
        if ($recurrence -notin $known -or $recurrence -in $allowed) { continue }

        $Sheet.DeleteRow($row, 1)
        $removed++

        foreach ($dropdown in @($Dropdowns | Where-Object { $_.Row -eq $row })) { [void]$Dropdowns.Remove($dropdown) }
        foreach ($dropdown in $Dropdowns) { if ($dropdown.Row -gt $row) { $dropdown.Row-- } }
    }

    if ($removed -gt 0) {
        Write-Host "Checklist: se quitaron $removed tareas que no corresponden a la proactiva ($($allowed -join ' / '))." -ForegroundColor DarkGray
    }
}

# Inserts a row right below the row of $AfterId with that row's look and Estado dropdown, and
# fills ID / Procedimiento / Chequeo from $Result
function Add-ChecklistRow {
    param(
        [Parameter(Mandatory = $true)] $Sheet,
        [Parameter(Mandatory = $true)] [string] $AfterId,
        [Parameter(Mandatory = $true)] $Result,
        [Parameter(Mandatory = $true)] [System.Collections.Generic.List[object]] $Dropdowns
    )

    $afterCell = $Sheet.Cells["A:A"] | Where-Object { $_.Text -eq $AfterId }
    if (-not $afterCell) {
        Write-Warning "No se encontro '$AfterId' en la plantilla para insertar la fila '$($Result.ID)'."
        return
    }
    $source = $afterCell.Start.Row
    $new = $source + 1

    $Sheet.InsertRow($new, 1)
    foreach ($dropdown in $Dropdowns) { if ($dropdown.Row -ge $new) { $dropdown.Row++ } }
    foreach ($dropdown in @($Dropdowns | Where-Object { $_.Row -eq $source })) {
        $Dropdowns.Add([PSCustomObject]@{ Column = $dropdown.Column; Row = $new; Formula = $dropdown.Formula })
    }

    # Same look as the source row, then its own content
    $Sheet.Cells["A${source}:I${source}"].Copy($Sheet.Cells["A${new}:I${new}"])
    $Sheet.Row($new).Height = $Sheet.Row($source).Height
    $Sheet.Cells["A$new"].Value = $Result.ID
    $Sheet.Cells["D$new"].Value = $Result.Procedimiento
    $Sheet.Cells["E$new"].Value = $Result.Chequeo
    foreach ($column in 'G', 'H', 'I') { $Sheet.Cells["$column$new"].Value = $null }
}

# Every cell that has a dropdown, as { Column; Row; Formula }
function Get-ChecklistDropdowns {
    param([Parameter(Mandatory = $true)] $Sheet)

    $dropdowns = [System.Collections.Generic.List[object]]::new()
    foreach ($validation in $Sheet.DataValidations) {
        foreach ($range in ($validation.Address.Address -split ' ')) {
            if ($range -notmatch '^([A-Z])(\d+)(?::([A-Z])(\d+))?$') {
                Write-Warning "Desplegable con un rango no soportado ('$range'); se omite."
                continue
            }
            $lastColumn = if ($matches[3]) { $matches[3] } else { $matches[1] }
            $lastRow = if ($matches[4]) { [int]$matches[4] } else { [int]$matches[2] }
            foreach ($column in ([int][char]$matches[1])..([int][char]$lastColumn)) {
                foreach ($row in ([int]$matches[2])..$lastRow) {
                    $dropdowns.Add([PSCustomObject]@{ Column = [string][char]$column; Row = $row; Formula = $validation.Formula.ExcelFormula })
                }
            }
        }
    }
    return , $dropdowns
}

# Puts the dropdowns back, one validation per list, with consecutive rows joined in ranges
function Restore-ChecklistDropdowns {
    param(
        [Parameter(Mandatory = $true)] $Sheet,
        [Parameter(Mandatory = $true)] [System.Collections.Generic.List[object]] $Dropdowns
    )

    foreach ($group in ($Dropdowns | Group-Object -Property Formula)) {
        $ranges = foreach ($columnGroup in ($group.Group | Group-Object -Property Column)) {
            $rows = @($columnGroup.Group.Row | Sort-Object -Unique)
            $start = $rows[0]
            for ($i = 0; $i -lt $rows.Count; $i++) {
                if ($i -eq $rows.Count - 1 -or $rows[$i + 1] -ne $rows[$i] + 1) {
                    $column = $columnGroup.Name
                    if ($start -eq $rows[$i]) { "$column$start" } else { "$column${start}:$column$($rows[$i])" }
                    if ($i -lt $rows.Count - 1) { $start = $rows[$i + 1] }
                }
            }
        }

        $list = $Sheet.DataValidations.AddListValidation(($ranges -join ' '))
        $list.Formula.ExcelFormula = $group.Name
        $list.AllowBlank = $true
        $list.ShowErrorMessage = $true
    }
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

# The "ESXi" sheet of the raw Excel without the virtual "hosts" that are not ESXi servers:
# HCX mobility appliances (model "VMware Mobility Platform") are listed as hosts but have no
# configuration data (no advanced settings, DNS, NTP...) and no meaning for the host checks.
# The raw Excel keeps them; only the checks skip them.
function Get-EsxiRows {
    param([Parameter(Mandatory = $true)] $Context)

    Import-RawSheet $Context.RawExcelPath "ESXi" | Where-Object { $_.Model -ne 'VMware Mobility Platform' }
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
