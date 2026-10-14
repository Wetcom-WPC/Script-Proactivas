using module .\ProactivaCollector.psm1

# Step 1 of the pipeline: connect to the vCenters and collect the data into JSON files.

# Connects to every vCenter in $Servers (SOAP + CIS/REST).
#
# Credentials work like a keychain: each credential that connects successfully is kept and
# tried first on the next vCenter, so vCenters that share credentials only ask once.
# With -Credentials (non-interactive runs) only those are tried and nothing is prompted.
#
# Returns one object per connected vCenter: @{ Name; Connection; CisConnection }.
function Connect-VCenters {
    param(
        [Parameter(Mandatory = $true)] [string[]] $Servers,
        [System.Management.Automation.PSCredential[]] $Credentials
    )

    $interactive = -not $Credentials
    $keychain = [System.Collections.ArrayList]@()
    if ($Credentials) { $Credentials | ForEach-Object { [void]$keychain.Add($_) } }

    $connected = @()

    foreach ($server in $Servers) {
        $connection = $null
        $usedCredential = $null
        $lastError = $null

        if ($keychain.Count -gt 0) {
            Write-Host "`nProbando credenciales guardadas en '$server'..." -ForegroundColor Cyan
            foreach ($credential in $keychain) {
                try {
                    $connection = Connect-VIServer -Server $server -Credential $credential -ErrorAction Stop
                    $usedCredential = $credential
                    break
                } catch { $lastError = $_.Exception.Message }
            }
        }

        while (-not $connection -and $interactive) {
            Write-Host "`nConectando a '$server' (se necesita una credencial nueva)..."
            try {
                $credential = Get-Credential -Message "Ingrese credenciales para '$server'"
                if (-not $credential) { throw "Cancelado por el usuario" }

                $connection = Connect-VIServer -Server $server -Credential $credential -ErrorAction Stop
                $usedCredential = $credential
                [void]$keychain.Add($credential)
            }
            catch {
                Write-Warning "-> FALLO la conexion con '$server': $($_.Exception.Message)"
                $choice = Read-Host "Presione 'R' para reintentar, o 'X' para omitir este vCenter"
                if ($choice.ToLower() -eq 'x') { break }
            }
        }

        if (-not $connection) {
            Write-Warning "-> No se pudo conectar a '$server'. Se omite.$(if ($lastError) { " Ultimo error: $lastError" })"
            continue
        }
        Write-Host "-> Conexion con '$server' exitosa." -ForegroundColor Green

        # CIS (REST) is needed for root password, certificates and VAMI backups
        $cisConnection = $null
        try {
            $cisConnection = Connect-CisServer -Server $server -Credential $usedCredential -ErrorAction Stop
            Write-Host "`t-> Conexion CIS (REST) exitosa." -ForegroundColor Green
        } catch {
            Write-Warning "`t-> La conexion CIS (REST) fallo: $($_.Exception.Message)"
        }

        $connected += [PSCustomObject]@{
            Name          = $connection.Name
            Connection    = $connection
            CisConnection = $cisConnection
        }
    }

    return $connected
}

function Disconnect-VCenters {
    param([Parameter(Mandatory = $true)] [object[]] $VCenters)

    foreach ($vcenter in $VCenters) {
        Disconnect-VIServer -Server $vcenter.Connection -Force -Confirm:$false -ErrorAction SilentlyContinue
        if ($vcenter.CisConnection) {
            Disconnect-CisServer -Server $vcenter.CisConnection -Force -Confirm:$false -ErrorAction SilentlyContinue
        }
    }
}

# Collects the Proactiva data (every sheet except "ESXi IO") from all vCenters.
# Returns the path of the JSON file written to <RunFolder>\json.
function Invoke-ProactivaCollection {
    param(
        [Parameter(Mandatory = $true)] [object[]] $VCenters,
        [Parameter(Mandatory = $true)] [string] $RunFolder,
        [switch] $SkipAlarmTest
    )

    $collector = [ProactivaCollector]::new((Join-Path $PSScriptRoot 'reference-data'), $RunFolder)

    foreach ($vcenter in $VCenters) {
        $collector.SetCurrentVCenter($vcenter.Connection, $vcenter.CisConnection)
        $server = $vcenter.Connection
        Write-Host "`nProcesando vCenter: $($vcenter.Name)"
        $timings = [ordered]@{}

        Write-Host "`tObteniendo objetos principales..."
        $hosts = @(Measure-CollectionStage $timings 'Hosts' { Get-VMHost -Server $server })
        $clusters = @(Measure-CollectionStage $timings 'Clusters' { Get-Cluster -Server $server })
        $vdSwitches = @(Measure-CollectionStage $timings 'Lista de vDS' { Get-VDSwitch -Server $server })
        $vmCount = Measure-CollectionStage $timings 'Conteo de VMs' { @(Get-VM -Server $server).Count }
        Write-Host "`tSe encontraron $($hosts.Count) hosts, $($clusters.Count) clusters, $vmCount VMs, $($vdSwitches.Count) vDS."

        # VMs and snapshots host by host (faster than one big query on large inventories)
        $allVms = [System.Collections.ArrayList]@()
        $allSnapshots = [System.Collections.ArrayList]@()

        Write-Host "`tRecolectando VMs y snapshots (puede demorar)..."
        Measure-CollectionStage $timings 'VMs y snapshots' {
            for ($i = 0; $i -lt $hosts.Count; $i++) {
                $vmHost = $hosts[$i]
                Write-Progress -Activity "Recolectando datos de $($vcenter.Name)" -Status "Host $($i+1)/$($hosts.Count): $($vmHost.Name)" -PercentComplete (($i / $hosts.Count) * 100)

                $vmsOnHost = @(Get-VM -Location $vmHost)
                if ($vmsOnHost) {
                    [void]$allVms.AddRange($vmsOnHost)

                    $snapshotsOnHost = @(Get-Snapshot -VM $vmsOnHost)
                    if ($snapshotsOnHost) {
                        [void]$allSnapshots.AddRange($snapshotsOnHost)
                    }
                }
            }
            Write-Progress -Activity "Recoleccion completa" -Completed
        }

        Write-Host "`tProcesando reportes..."
        Measure-CollectionStage $timings 'Prueba de alarmas' { $collector.CollectAlarmTest([bool]$SkipAlarmTest) }
        Measure-CollectionStage $timings 'Sizing del vCenter' { $collector.CollectVCenterSizing($allVms, $hosts) }
        Measure-CollectionStage $timings 'Backups (VAMI)' { $collector.CollectBackupActivity() }
        Measure-CollectionStage $timings 'Estadisticas de performance' { $collector.CollectPerformanceHealth($clusters) }
        Measure-CollectionStage $timings 'Certificados del vCenter' { $collector.CollectCertificates() }
        Measure-CollectionStage $timings 'ESXi' { $collector.CollectEsxi($hosts) }
        Measure-CollectionStage $timings 'Informacion del vCenter' { $collector.CollectVCenterInfo($allVms) }
        Measure-CollectionStage $timings 'VMs y placas de red' { $collector.CollectVms($allVms, $clusters, $allSnapshots) }
        Measure-CollectionStage $timings 'Datastores' { $collector.CollectDatastores($hosts) }
        Measure-CollectionStage $timings 'Standard switches' { $collector.CollectStandardSwitches($hosts) }
        Measure-CollectionStage $timings 'Adaptadores VMkernel' { $collector.CollectKernelAdapters($hosts) }
        Measure-CollectionStage $timings 'Snapshots' { $collector.CollectSnapshots($allSnapshots) }
        Measure-CollectionStage $timings 'Particiones' { $collector.CollectPartitions($allVms) }
        Measure-CollectionStage $timings 'vDS' { $collector.CollectDistributedSwitches($vdSwitches) }
        Measure-CollectionStage $timings 'Licencias' { $collector.CollectLicenses() }

        Write-StageTimings -Title "recoleccion de $($vcenter.Name)" -Timings $timings
    }

    return Save-CollectionJson -Name "Proactiva" -Report $collector.GetReport() -VCenters $VCenters -RunFolder $RunFolder
}

# Collects host drivers/firmware ("ESXi IO" sheet) from all vCenters. Quarterly and semiannual
# visits only. Returns the path of the JSON file written to <RunFolder>\json.
function Invoke-DriverCollection {
    param(
        [Parameter(Mandatory = $true)] [object[]] $VCenters,
        [Parameter(Mandatory = $true)] [string] $RunFolder
    )

    $collector = [ProactivaCollector]::new((Join-Path $PSScriptRoot 'reference-data'), $RunFolder)

    foreach ($vcenter in $VCenters) {
        $collector.SetCurrentVCenter($vcenter.Connection, $vcenter.CisConnection)
        Write-Host "`nProcesando drivers del vCenter: $($vcenter.Name)"

        $timings = [ordered]@{}
        $hosts = @(Measure-CollectionStage $timings 'Hosts' { Get-VMHost -Server $vcenter.Connection })
        Write-Host "`t$($hosts.Count) hosts encontrados"

        Measure-CollectionStage $timings 'Drivers y firmware' { $collector.CollectHostDrivers($hosts) }
        Write-StageTimings -Title "drivers de $($vcenter.Name)" -Timings $timings
    }

    return Save-CollectionJson -Name "DriverInfo" -Report $collector.GetReport() -VCenters $VCenters -RunFolder $RunFolder
}

# Runs one collection stage, records how long it took in $Timings and returns its output
function Measure-CollectionStage {
    param(
        [Parameter(Mandatory = $true)] [System.Collections.Specialized.OrderedDictionary] $Timings,
        [Parameter(Mandatory = $true)] [string] $Name,
        [Parameter(Mandatory = $true)] [scriptblock] $Stage
    )

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        & $Stage
    }
    finally {
        $Timings[$Name] = $stopwatch.Elapsed
    }
}

# Duration as m:ss.f (or s.f under a minute)
function Format-Duration([TimeSpan] $Duration) {
    if ($Duration.TotalMinutes -ge 1) { return "{0}:{1:00.0}" -f [math]::Floor($Duration.TotalMinutes), ($Duration.TotalSeconds % 60) }
    return "{0:0.0} s" -f $Duration.TotalSeconds
}

# Table of the time each stage took, slowest first (ends up in the run log)
function Write-StageTimings {
    param(
        [Parameter(Mandatory = $true)] [string] $Title,
        [Parameter(Mandatory = $true)] [System.Collections.Specialized.OrderedDictionary] $Timings
    )

    $total = [TimeSpan]::Zero
    foreach ($duration in $Timings.Values) { $total += $duration }

    Write-Host "`n`tTiempos de $Title (total $(Format-Duration $total)):" -ForegroundColor Cyan
    foreach ($entry in ($Timings.GetEnumerator() | Sort-Object { $_.Value } -Descending)) {
        $share = if ($total.TotalSeconds -gt 0) { $entry.Value.TotalSeconds / $total.TotalSeconds * 100 } else { 0 }
        Write-Host ("`t  {0,-30} {1,9} {2,5:0}%" -f $entry.Key, (Format-Duration $entry.Value), $share)
    }
}

function Save-CollectionJson {
    param(
        [Parameter(Mandatory = $true)] [string] $Name,
        [Parameter(Mandatory = $true)] $Report,
        [Parameter(Mandatory = $true)] [object[]] $VCenters,
        [Parameter(Mandatory = $true)] [string] $RunFolder
    )

    $content = [PSCustomObject] @{
        Result    = "OK"
        Name      = $Name
        Version   = $global:APP_VERSION
        DateTime  = Get-Date -Format "yyyy-MM-dd HH:mm"
        LocalHost = [System.Environment]::MachineName
        User      = whoami
        Endpoint  = $VCenters.Name -join ", "
        Component = "vcenter"
        Report    = $Report
    }

    $jsonFolder = Join-Path -Path $RunFolder -ChildPath "json"
    if (-not (Test-Path $jsonFolder)) { New-Item -ItemType Directory -Path $jsonFolder -Force | Out-Null }

    $jsonPath = Join-Path -Path $jsonFolder -ChildPath ("{0}_{1}.json" -f (Get-Date -Format 'yyyy-MM-dd HHmmss'), $Name)
    $content | ConvertTo-Json -Depth 99 | Set-Content -Path $jsonPath

    return $jsonPath
}

Export-ModuleMember -Function Connect-VCenters, Disconnect-VCenters, Invoke-ProactivaCollection, Invoke-DriverCollection
