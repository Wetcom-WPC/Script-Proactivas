<#
.SYNOPSIS
Generates the Proactiva visit deliverables (Anexo Tecnico + Checklist) for one or more vCenters.

.DESCRIPTION
Runs the whole pipeline in a single process:
  1. Collection  - connects to the vCenters and saves the data as JSON (resultados\<run>\json)
  2. Conversion  - turns each JSON into a raw Excel, one sheet per topic (resultados\<run>\excel)
  3. Report      - runs the checks for the visit frequency and writes the annex and the
                   checklist (resultados\<run>\anexo)

Anything not passed as a parameter is asked at the start, so the run needs no attention
after that. Iniciar-Proactiva.bat (repo root) launches this script with the bundled pwsh.

.PARAMETER Frequency
Monthly, Quarterly or Semiannual. Cumulative: Quarterly also runs the monthly checks and
collects host drivers; Semiannual runs everything.

.PARAMETER CredentialPath
Folder with credentials saved as cred-*.xml (Get-Credential | Export-Clixml). When given,
those are the only credentials tried and nothing is prompted. Meant for unattended/test runs.

.PARAMETER ReportOnly
Skips collection and conversion and regenerates the annex/checklist from an existing run
(-RunFolder, or the most recent one). Useful to fix a typo in the client or month.

.PARAMETER Sequential
Collects everything in sequence. By default several vCenters, and the drivers of many hosts,
are collected in parallel (one worker process each). Useful to debug or to compare results.

.PARAMETER SkipAlarmTest
Does not create the temporary "Falso Positivo" alarm in the vCenter (the only change the
collection makes to the environment). The rest of the alarm audit still runs.

.EXAMPLE
.\Start-Proactiva.ps1
.\Start-Proactiva.ps1 -Client "Cliente" -Month "Octubre" -Frequency Quarterly -VCenters vc01.cliente.local
.\Start-Proactiva.ps1 -ReportOnly -Client "Cliente" -Month "Octubre" -Frequency Quarterly
#>
param(
    [string] $Client,
    [string] $Month,
    [ValidateSet('Monthly', 'Quarterly', 'Semiannual')] [string] $Frequency,
    [string[]] $VCenters,
    [string] $CredentialPath,
    [switch] $ReportOnly,
    [string] $RunFolder,
    [switch] $SkipAlarmTest,
    [switch] $Sequential
)

$ErrorActionPreference = 'Continue'
$global:APP_VERSION = "v3.2.2"

$repoRoot = Split-Path -Path $PSScriptRoot -Parent
$runtimeModules = Join-Path $repoRoot 'runtime\Modules'

# Only the bundled modules are visible: a PowerCLI installed on the PC (e.g. 12.5 under
# Program Files\WindowsPowerShell\Modules) would otherwise be resolved by name and fail with
# "Assembly with same name is already loaded". Workers started later inherit this.
$env:PSModulePath = $runtimeModules

$frequencyLabels = [ordered]@{
    Monthly    = 'Mensual'
    Quarterly  = 'Trimestral'
    Semiannual = 'Semestral'
}

Import-Module (Join-Path $PSScriptRoot 'shared\RunFolder.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'shared\Exclusions.psm1') -Force

# Duration as m:ss (shown in the final summary)
function Format-StageDuration([TimeSpan] $Duration) {
    return "{0}:{1:00}" -f [math]::Floor($Duration.TotalMinutes), $Duration.Seconds
}

# ---------------------------------------------------------------------------------------
# Visit details: asked up front so the rest of the run is unattended
# ---------------------------------------------------------------------------------------

Write-Host "==========================================================" -ForegroundColor Green
Write-Host " WPC - PROACTIVAS $global:APP_VERSION" -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Green

while ([string]::IsNullOrWhiteSpace($Client)) {
    $Client = Read-Host "Nombre del cliente"
}
$Client = $Client.Trim()

if ([string]::IsNullOrWhiteSpace($Month)) {
    $currentMonth = (Get-Culture -Name 'es-AR').DateTimeFormat.GetMonthName((Get-Date).Month)
    $currentMonth = $currentMonth.Substring(0, 1).ToUpper() + $currentMonth.Substring(1)
    $Month = Read-Host "Mes del reporte [$currentMonth]"
    if ([string]::IsNullOrWhiteSpace($Month)) { $Month = $currentMonth }
}
$Month = $Month.Trim()

while (-not $Frequency) {
    Write-Host ""
    Write-Host "Tipo de proactiva:"
    Write-Host "  1) Mensual"
    Write-Host "  2) Trimestral  (incluye la mensual + drivers de hosts)"
    Write-Host "  3) Semestral   (incluye la trimestral y la mensual)"
    switch ((Read-Host "Opcion").Trim()) {
        '1' { $Frequency = 'Monthly' }
        '2' { $Frequency = 'Quarterly' }
        '3' { $Frequency = 'Semiannual' }
        default { Write-Host "Opcion no valida." -ForegroundColor Yellow }
    }
}

if (-not $ReportOnly) {
    $VCenters = @($VCenters -join ',' -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    while ($VCenters.Count -eq 0) {
        $VCenters = @((Read-Host "vCenters (separados por coma)") -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }

    $avsVCenters = @($VCenters | Where-Object { Test-IsAvsVCenter $_ })
    if ($avsVCenters.Count -gt 0) {
        Write-Host "Azure VMware Solution ($($avsVCenters -join ', ')): lo gestiona Microsoft, se relevan solo las VMs." -ForegroundColor Yellow
    }
}

# ---------------------------------------------------------------------------------------
# Run folder and log
# ---------------------------------------------------------------------------------------

try {
    $runPath = if ($ReportOnly) { Get-RunFolder -Name $RunFolder } else { New-RunFolder }
} catch {
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}

$logFolder = Join-Path $runPath 'logs'
if (-not (Test-Path $logFolder)) { New-Item -ItemType Directory -Path $logFolder -Force | Out-Null }
Start-Transcript -Path (Join-Path $logFolder ("{0}_proactiva.log" -f (Get-Date -Format 'yyyy-MM-dd_HHmmss'))) | Out-Null

Write-Host ""
Write-Host "Cliente:     $Client"
Write-Host "Mes:         $Month"
Write-Host "Proactiva:   $($frequencyLabels[$Frequency])"
if (-not $ReportOnly) { Write-Host "vCenters:    $($VCenters -join ', ')" }
Write-Host "Resultados:  $runPath"
if ($SkipAlarmTest) { Write-Host "Prueba de alarmas (Falso Positivo): OMITIDA" -ForegroundColor Yellow }

$exitCode = 0
$connectedVCenters = @()
$stageTimings = [ordered]@{}
$stageStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

try {
    Import-Module (Join-Path $runtimeModules 'ImportExcel') -ErrorAction Stop

    if (-not $ReportOnly) {
        # -------------------------------------------------------------------------------
        # 1. Collection
        # -------------------------------------------------------------------------------
        Write-Host "`n[1/3] Recoleccion de datos" -ForegroundColor Cyan

        # Loaded by path so a PowerCLI installed elsewhere on the PC never takes precedence
        foreach ($module in 'VMware.VimAutomation.Sdk', 'VMware.VimAutomation.Common', 'VMware.VimAutomation.Cis.Core', 'VMware.VimAutomation.Core', 'VMware.VimAutomation.Vds') {
            Import-Module (Join-Path $runtimeModules $module) -ErrorAction Stop -WarningAction SilentlyContinue
        }
        Set-PowerCLIConfiguration -Scope Session -ParticipateInCEIP $false -InvalidCertificateAction Ignore -Confirm:$false | Out-Null

        # Background download of the server HCL (thread job). Without it, it downloads when needed.
        Import-Module (Join-Path $runtimeModules 'ThreadJob') -ErrorAction SilentlyContinue

        Import-Module (Join-Path $PSScriptRoot 'collection\Collection.psm1') -Force -WarningAction SilentlyContinue

        $credentials = $null
        if ($CredentialPath) {
            $credentials = @(Get-ChildItem -Path $CredentialPath -Filter 'cred-*.xml' | ForEach-Object { Import-Clixml $_.FullName })
            if ($credentials.Count -eq 0) { throw "No se encontraron archivos cred-*.xml en '$CredentialPath'." }
        }

        $connectedVCenters = @(Connect-VCenters -Servers $VCenters -Credentials $credentials)
        if ($connectedVCenters.Count -eq 0) { throw "No se pudo conectar a ningun vCenter." }

        $jsonFiles = @()
        $jsonFiles += Invoke-ProactivaCollection -VCenters $connectedVCenters -RunFolder $runPath -SkipAlarmTest:$SkipAlarmTest -Sequential:$Sequential
        if ($Frequency -ne 'Monthly') {
            $jsonFiles += Invoke-DriverCollection -VCenters $connectedVCenters -RunFolder $runPath -Sequential:$Sequential
        }

        Disconnect-VCenters -VCenters $connectedVCenters
        $connectedVCenters = @()
        $stageTimings['Recoleccion'] = $stageStopwatch.Elapsed
        $stageStopwatch.Restart()

        # -------------------------------------------------------------------------------
        # 2. Conversion
        # -------------------------------------------------------------------------------
        Write-Host "`n[2/3] Conversion a Excel" -ForegroundColor Cyan
        Import-Module (Join-Path $PSScriptRoot 'conversion\Conversion.psm1') -Force

        foreach ($jsonFile in $jsonFiles) {
            [void](ConvertTo-RawExcel -JsonPath $jsonFile -ExcelFolder (Join-Path $runPath 'excel'))
        }
        $stageTimings['Conversion a Excel'] = $stageStopwatch.Elapsed
        $stageStopwatch.Restart()
    }

    # -----------------------------------------------------------------------------------
    # 3. Report
    # -----------------------------------------------------------------------------------
    Write-Host "`n[3/3] Anexo Tecnico y Checklist" -ForegroundColor Cyan
    Import-Module (Join-Path $PSScriptRoot 'report\Report.psm1') -Force

    $excelFolder = Join-Path $runPath 'excel'
    $rawExcel = Get-ChildItem -Path $excelFolder -Filter '*_Proactiva.xlsx' -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    $driverExcel = Get-ChildItem -Path $excelFolder -Filter '*_DriverInfo.xlsx' -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1

    if (-not $rawExcel) { throw "No se encontro el Excel de Proactiva en '$excelFolder'." }

    $report = New-ProactivaReport -RunFolder $runPath -RawExcelPath $rawExcel.FullName -DriverExcelPath $driverExcel.FullName `
        -Client $Client -Month $Month -Frequency $Frequency
    $stageTimings['Anexo y Checklist'] = $stageStopwatch.Elapsed

    Write-Host "`n==========================================================" -ForegroundColor Green
    Write-Host " Proceso completado" -ForegroundColor Green
    if ($report.AnnexPath) { Write-Host " Anexo:     $($report.AnnexPath)" -ForegroundColor Green }
    Write-Host " Checklist: $($report.ChecklistPath)" -ForegroundColor Green
    if ($Frequency -eq 'Semiannual' -and (Get-ChildItem (Join-Path $runPath 'vds_configuration') -Filter '*.zip' -ErrorAction SilentlyContinue)) {
        Write-Host " Recuerde subir al Drive los .zip de: $(Join-Path $runPath 'vds_configuration')" -ForegroundColor Yellow
    }
    Write-Host " Tiempos:   $(($stageTimings.GetEnumerator() | ForEach-Object { "$($_.Key) $(Format-StageDuration $_.Value)" }) -join ' | ')" -ForegroundColor Green
    Write-Host "==========================================================" -ForegroundColor Green
}
catch {
    Write-Host "`nERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "El detalle completo quedo en el log: $logFolder" -ForegroundColor Red
    $exitCode = 1
}
finally {
    if ($connectedVCenters.Count -gt 0) { Disconnect-VCenters -VCenters $connectedVCenters }
    Stop-Transcript | Out-Null
}

exit $exitCode
