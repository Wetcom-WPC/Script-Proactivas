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
