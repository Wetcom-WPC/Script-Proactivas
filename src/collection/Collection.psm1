using module .\ProactivaCollector.psm1

Import-Module (Join-Path $PSScriptRoot '..\shared\Exclusions.psm1')

# Step 1 of the pipeline: connect to the vCenters and collect the data into JSON files.

# Connects to every vCenter in $Servers (SOAP + CIS/REST).
#
# Credentials work like a keychain: each credential that connects successfully is kept and
# tried first on the next vCenter, so vCenters that share credentials only ask once.
# With -Credentials (non-interactive runs) only those are tried and nothing is prompted.
#
# Returns one object per connected vCenter: @{ Name; Connection; CisConnection; Credential }.
# The credential stays in memory only: the parallel workers use it to open their own sessions.
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
            Credential    = $usedCredential
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

# ---------------------------------------------------------------------------------------
# Proactiva data (every sheet except "ESXi IO")
# ---------------------------------------------------------------------------------------

# Collects the Proactiva data (every sheet except "ESXi IO") from all vCenters.
#
# With more than one vCenter they are collected at the same time, one worker process each
# (see Invoke-CollectionWorkers); -Sequential collects them one after the other. A vCenter
# whose worker fails is collected again here, in sequence.
# Returns the path of the JSON file written to <RunFolder>\json.
function Invoke-ProactivaCollection {
    param(
        [Parameter(Mandatory = $true)] [object[]] $VCenters,
        [Parameter(Mandatory = $true)] [string] $RunFolder,
        [switch] $SkipAlarmTest,
        [switch] $Sequential
    )

    $referenceDataFolder = Join-Path $PSScriptRoot 'reference-data'

    # The server HCL (TR-02) downloads in the background while the inventory is read
    Start-ServerHclDownload

    $reports = @{}

    if (-not $Sequential -and $VCenters.Count -gt 1 -and (Test-ParallelCollection)) {
        # Every worker needs the HCL, so it is waited for here and handed over once (the
        # lookup caches are rebuilt by each worker)
        $loadedHcl = Get-HclReferenceData -ReferenceDataFolder $referenceDataFolder
        $hclData = [PSCustomObject]@{ Servers = $loadedHcl.Servers; Source = $loadedHcl.Source; EsxiReleases = $loadedHcl.EsxiReleases }

        Write-Host "`nRecolectando $($VCenters.Count) vCenters en paralelo (el detalle de cada uno se muestra al terminar)..."
        $tasks = foreach ($vcenter in $VCenters) {
            [PSCustomObject]@{
                Label     = $vcenter.Name
                VCenter   = $vcenter
                WithCis   = $null -ne $vcenter.CisConnection
                Command   = 'Invoke-VCenterCollectionWorker'
                Arguments = @{ RunFolder = $RunFolder; SkipAlarmTest = [bool]$SkipAlarmTest; HclData = $hclData }
            }
        }

        $results = @(Invoke-CollectionWorkers -Tasks $tasks)
        for ($i = 0; $i -lt $VCenters.Count; $i++) {
            if ($results[$i].Ok -and $results[$i].Result) { $reports[$i] = $results[$i].Result }
            else { Write-Warning "La recoleccion en paralelo de '$($VCenters[$i].Name)' fallo ($($results[$i].Error)). Se recolecta de nuevo, en secuencia." }
        }
    }

    for ($i = 0; $i -lt $VCenters.Count; $i++) {
        if ($reports.ContainsKey($i)) { continue }
        $collector = [ProactivaCollector]::new($referenceDataFolder, $RunFolder)
        $collector.SetCurrentVCenter($VCenters[$i].Connection, $VCenters[$i].CisConnection)
        Invoke-VCenterCollection -Collector $collector -Server $VCenters[$i].Connection -SkipAlarmTest:$SkipAlarmTest | Out-Null
        $reports[$i] = $collector.GetReport()
    }

    $report = Merge-CollectionReports -Reports @(0..($VCenters.Count - 1) | ForEach-Object { $reports[$_] })
    return Save-CollectionJson -Name "Proactiva" -Report $report -VCenters $VCenters -RunFolder $RunFolder
}

# Collects one vCenter into $Collector (already pointed at it with SetCurrentVCenter) and
# prints how long each stage took.
function Invoke-VCenterCollection {
    param(
        [Parameter(Mandatory = $true)] $Collector,
        [Parameter(Mandatory = $true)] $Server,
        [switch] $SkipAlarmTest
    )

    Write-Host "`nProcesando vCenter: $($Collector.currentVCenter)"
    $timings = [ordered]@{}

    # Azure VMware Solution: Microsoft runs everything except the VMs, so only those (and
    # their snapshots) are collected: no alarm test, no hosts, no appliance, no vDS export
    $isAvs = Test-IsAvsVCenter $Collector.currentVCenter
    if ($isAvs) {
        Write-Host "`tvCenter de Azure VMware Solution: se relevan solo las VMs y sus snapshots (el resto lo gestiona Microsoft)." -ForegroundColor Yellow
    }

    Write-Host "`tObteniendo objetos principales..."
    $hosts = @(Measure-CollectionStage $timings 'Hosts' { Get-VMHost -Server $Server })
    $clusters = @(Measure-CollectionStage $timings 'Clusters' { Get-Cluster -Server $Server })
    $vdSwitches = if ($isAvs) { @() } else { @(Measure-CollectionStage $timings 'Lista de vDS' { Get-VDSwitch -Server $Server }) }
    Measure-CollectionStage $timings 'Ubicacion de los hosts' { $Collector.SetHostLocations($clusters) }
    if (-not $isAvs) {
        Measure-CollectionStage $timings 'Lista de datastores' { $Collector.SetDatastores((Get-Datastore -Server $Server)) }
    }

    # Every VM in one query (with its network cards and ISOs); snapshots only of the VMs
    # that have any
    Write-Host "`tRecolectando VMs y snapshots..."
    $allVms = @(Measure-CollectionStage $timings 'VMs' { Get-VmInventory -Server $Server -Hosts $hosts })
    $allSnapshots = @(Measure-CollectionStage $timings 'Snapshots (lectura)' { Get-VmSnapshots -Server $Server -Vms $allVms })
    Write-Host "`tSe encontraron $($hosts.Count) hosts, $($clusters.Count) clusters, $($allVms.Count) VMs, $($vdSwitches.Count) vDS."

    Write-Host "`tProcesando reportes..."
    if ($isAvs) {
        Measure-CollectionStage $timings 'VMs y placas de red' { $Collector.CollectVms($allVms, $clusters, $allSnapshots) }
        Measure-CollectionStage $timings 'Snapshots' { $Collector.CollectSnapshots($allSnapshots) }
        Write-StageTimings -Title "recoleccion de $($Collector.currentVCenter)" -Timings $timings
        return
    }

    Measure-CollectionStage $timings 'Prueba de alarmas' { $Collector.CollectAlarmTest([bool]$SkipAlarmTest) }
    Measure-CollectionStage $timings 'Sizing del vCenter' { $Collector.CollectVCenterSizing($allVms, $hosts) }
    Measure-CollectionStage $timings 'Backups (VAMI)' { $Collector.CollectBackupActivity() }
    Measure-CollectionStage $timings 'Estadisticas de performance' { $Collector.CollectPerformanceHealth($clusters) }
    Measure-CollectionStage $timings 'Certificados del vCenter' { $Collector.CollectCertificates() }
    Measure-CollectionStage $timings 'ESXi' { $Collector.CollectEsxi($hosts) }
    Measure-CollectionStage $timings 'Informacion del vCenter' { $Collector.CollectVCenterInfo($allVms) }
    Measure-CollectionStage $timings 'VMs y placas de red' { $Collector.CollectVms($allVms, $clusters, $allSnapshots) }
    Measure-CollectionStage $timings 'Datastores' { $Collector.CollectDatastores($hosts) }
    Measure-CollectionStage $timings 'Standard switches' { $Collector.CollectStandardSwitches($hosts) }
    Measure-CollectionStage $timings 'Adaptadores VMkernel' { $Collector.CollectKernelAdapters($hosts) }
    Measure-CollectionStage $timings 'Snapshots' { $Collector.CollectSnapshots($allSnapshots) }
    Measure-CollectionStage $timings 'Particiones' { $Collector.CollectPartitions($allVms) }
    Measure-CollectionStage $timings 'vDS' { $Collector.CollectDistributedSwitches($vdSwitches) }
    Measure-CollectionStage $timings 'Licencias' { $Collector.CollectLicenses() }

    Write-StageTimings -Title "recoleccion de $($Collector.currentVCenter)" -Timings $timings
}

# Worker side of Invoke-ProactivaCollection: collects one vCenter with the worker's own
# connection and returns its report.
function Invoke-VCenterCollectionWorker {
    param(
        [Parameter(Mandatory = $true)] $Connection,
        $CisConnection,
        [Parameter(Mandatory = $true)] [string] $RunFolder,
        [bool] $SkipAlarmTest,
        $HclData
    )

    if ($HclData) {
        Add-Member -InputObject $HclData -NotePropertyName ModelEntries -NotePropertyValue ([System.Collections.Concurrent.ConcurrentDictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)) -Force
        Add-Member -InputObject $HclData -NotePropertyName Releases -NotePropertyValue ([System.Collections.Concurrent.ConcurrentDictionary[string, object]]::new()) -Force
        Set-HclReferenceData -Data $HclData
    }

    $collector = [ProactivaCollector]::new((Join-Path $PSScriptRoot 'reference-data'), $RunFolder)
    $collector.SetCurrentVCenter($Connection, $CisConnection)
    Invoke-VCenterCollection -Collector $collector -Server $Connection -SkipAlarmTest:$SkipAlarmTest | Out-Null
    return $collector.GetReport()
}

# Joins the reports of several vCenters, sheet by sheet, in the order given. Keeps the rule
# CollectDatastores applies within a vCenter: a vSAN datastore whose name was already listed
# (by an earlier vCenter) is listed only once.
function Merge-CollectionReports {
    param([Parameter(Mandatory = $true)] [object[]] $Reports)

    if ($Reports.Count -eq 1) { return $Reports[0] }

    $sheetNames = @($Reports[0].PSObject.Properties.Name)
    $merged = [ordered]@{}
    foreach ($name in $sheetNames) { $merged[$name] = [System.Collections.ArrayList]@() }

    $listedDatastores = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($report in $Reports) {
        foreach ($name in $sheetNames) {
            $rows = @($report.$name)
            if ($name -eq 'Datastores') {
                $rows = @($rows | Where-Object { -not ($_.Policy -eq 'vSAN' -and $listedDatastores.Contains([string]$_.Datastore)) })
                foreach ($row in $rows) {
                    foreach ($datastoreName in @($row.Datastore)) {
                        if ($null -ne $datastoreName) { [void]$listedDatastores.Add([string]$datastoreName) }
                    }
                }
            }
            foreach ($row in $rows) { [void]$merged[$name].Add($row) }
        }
    }
    return [PSCustomObject]$merged
}

# ---------------------------------------------------------------------------------------
# Host drivers ("ESXi IO" sheet)
# ---------------------------------------------------------------------------------------

# Collects host drivers/firmware ("ESXi IO" sheet) from all vCenters. Quarterly and semiannual
# visits only.
#
# esxcli through vCenter is slow and each host is independent, so the hosts of a vCenter are
# split among several worker processes (unless -Sequential, or there are only a few hosts). Hosts
# whose worker fails are collected here, in sequence. Rows keep the order of the hosts.
# Returns the path of the JSON file written to <RunFolder>\json.
function Invoke-DriverCollection {
    param(
        [Parameter(Mandatory = $true)] [object[]] $VCenters,
        [Parameter(Mandatory = $true)] [string] $RunFolder,
        [switch] $Sequential
    )

    $collector = [ProactivaCollector]::new((Join-Path $PSScriptRoot 'reference-data'), $RunFolder)
    $maxWorkers = 4
    $minHostsForParallel = 6

    foreach ($vcenter in $VCenters) {
        # The hosts of Azure VMware Solution are Microsoft's
        if (Test-IsAvsVCenter $vcenter.Name) { continue }

        $collector.SetCurrentVCenter($vcenter.Connection, $vcenter.CisConnection)
        Write-Host "`nProcesando drivers del vCenter: $($vcenter.Name)"

        $timings = [ordered]@{}
        $hosts = @(Measure-CollectionStage $timings 'Hosts' { Get-VMHost -Server $vcenter.Connection })
        Write-Host "`t$($hosts.Count) hosts encontrados"

        $reachableHosts = @($hosts | Where-Object { @("Connected", "Maintenance") -contains $_.ConnectionState })
        if ($Sequential -or $reachableHosts.Count -lt $minHostsForParallel -or -not (Test-ParallelCollection)) {
            Measure-CollectionStage $timings 'Drivers y firmware' { $collector.CollectHostDrivers($hosts) }
        }
        else {
            Measure-CollectionStage $timings 'Drivers y firmware (en paralelo)' {
                # Indexes built once here and shared (read-only) with every worker
                $ioIndex = $collector.GetIoHclIndex()
                $vsanIndex = $collector.GetVsanHclIndex()

                $workerCount = [math]::Min($maxWorkers, [math]::Ceiling($reachableHosts.Count / 2))
                Write-Host "`tProcesando drivers de $($reachableHosts.Count) hosts en $workerCount tareas en paralelo..."

                $tasks = for ($w = 0; $w -lt $workerCount; $w++) {
                    $hostIds = @(for ($i = $w; $i -lt $reachableHosts.Count; $i += $workerCount) { $reachableHosts[$i].Id })
                    [PSCustomObject]@{
                        Label     = "$($vcenter.Name) (tarea $($w + 1) de $workerCount)"
                        VCenter   = $vcenter
                        WithCis   = $false
                        Command   = 'Invoke-HostDriversWorker'
                        Arguments = @{ RunFolder = $RunFolder; HostIds = $hostIds; IoHclIndex = $ioIndex; VsanHclIndex = $vsanIndex }
                        HostIds   = $hostIds
                    }
                }
                $results = @(Invoke-CollectionWorkers -Tasks $tasks)

                # Rows by host; hosts of a failed worker are collected here
                $rowsByHost = @{}
                $failedHostIds = [System.Collections.Generic.List[string]]::new()
                for ($i = 0; $i -lt $tasks.Count; $i++) {
                    if ($results[$i].Ok) {
                        foreach ($row in $results[$i].Result.Rows) {
                            if (-not $rowsByHost.ContainsKey($row.Hostname)) { $rowsByHost[$row.Hostname] = [System.Collections.Generic.List[object]]::new() }
                            $rowsByHost[$row.Hostname].Add($row)
                        }
                    }
                    else {
                        Write-Warning "Una tarea de drivers fallo ($($results[$i].Error)). Sus hosts se procesan de nuevo, en secuencia."
                        $failedHostIds.AddRange([string[]]$tasks[$i].HostIds)
                    }
                }

                if ($failedHostIds.Count -gt 0) {
                    $fallback = [ProactivaCollector]::new((Join-Path $PSScriptRoot 'reference-data'), $RunFolder)
                    $fallback.SetCurrentVCenter($vcenter.Connection, $vcenter.CisConnection)
                    $fallback.ioHclIndex = $ioIndex
                    $fallback.vsanHclIndex = $vsanIndex
                    $fallback.CollectHostDrivers(@($reachableHosts | Where-Object { $failedHostIds -contains $_.Id }))
                    foreach ($row in $fallback.nicReport) {
                        if (-not $rowsByHost.ContainsKey($row.Hostname)) { $rowsByHost[$row.Hostname] = [System.Collections.Generic.List[object]]::new() }
                        $rowsByHost[$row.Hostname].Add($row)
                    }
                }

                foreach ($h in $reachableHosts) {
                    foreach ($row in $rowsByHost[$h.Name]) { [void]$collector.nicReport.Add($row) }
                }
            }
        }
        Write-StageTimings -Title "drivers de $($vcenter.Name)" -Timings $timings
    }

    return Save-CollectionJson -Name "DriverInfo" -Report $collector.GetReport() -VCenters $VCenters -RunFolder $RunFolder
}

# Worker side of Invoke-DriverCollection: drivers of some hosts of one vCenter, with the
# worker's own connection. Returns @{ Rows } in the order of $HostIds.
function Invoke-HostDriversWorker {
    param(
        [Parameter(Mandatory = $true)] $Connection,
        $CisConnection,
        [Parameter(Mandatory = $true)] [string] $RunFolder,
        [Parameter(Mandatory = $true)] [string[]] $HostIds,
        [hashtable] $IoHclIndex,
        [hashtable] $VsanHclIndex
    )

    $collector = [ProactivaCollector]::new((Join-Path $PSScriptRoot 'reference-data'), $RunFolder)
    $collector.SetCurrentVCenter($Connection, $CisConnection)
    if ($IoHclIndex) { $collector.ioHclIndex = $IoHclIndex }
    if ($VsanHclIndex) { $collector.vsanHclIndex = $VsanHclIndex }

    $hostsById = @{}
    foreach ($h in (Get-VMHost -Id $HostIds -Server $Connection)) { $hostsById[$h.Id] = $h }
    $collector.CollectHostDrivers(@($HostIds | ForEach-Object { $hostsById[$_] } | Where-Object { $_ }))

    return [PSCustomObject]@{ Rows = $collector.nicReport }
}

# ---------------------------------------------------------------------------------------
# Parallel workers
# ---------------------------------------------------------------------------------------

# Whether collection can run in parallel workers (background jobs, always available in pwsh)
function Test-ParallelCollection {
    return [bool](Get-Command Start-Job -ErrorAction SilentlyContinue)
}

# Runs each task in its own worker and returns one result per task, in the same order:
# @{ Ok; Result; Error }. A task is @{ Label; VCenter; WithCis; Command; Arguments }.
#
# Each worker is a separate pwsh PROCESS (Start-Job), not a thread: PowerCLI keeps connection
# state per process, and two thread jobs connecting at once break each other ("You have
# modified the global:DefaultVIServer ... system variables"). The worker loads PowerCLI and
# this module, opens its OWN session to the vCenter with the credential that worked in
# Connect-VCenters (SOAP, and CIS when WithCis), runs "<Command> -Connection -CisConnection
# @Arguments" and disconnects that session.
#
# Arguments and results travel as Clixml files in a temporary folder (deleted at the end):
# job output is only serialized one level deep, which would flatten the report rows.
#
# What each worker writes (Write-Host, warnings, errors) is shown when it finishes, in task
# order, so the log stays readable.
function Invoke-CollectionWorkers {
    param(
        [Parameter(Mandatory = $true)] [object[]] $Tasks,
        # Longest a worker may take (counted from when its turn to be waited for comes)
        [int] $TimeoutSeconds = 3 * 60 * 60,
        # Where the workers load PowerCLI from: the bundled runtime (tests point it at stubs)
        [string] $ModulesFolder = (Join-Path $PSScriptRoot '..\..\runtime\Modules')
    )

    $modulesFolder = (Resolve-Path $ModulesFolder).Path
    $collectionModule = Join-Path $PSScriptRoot 'Collection.psm1'
    $exchangeFolder = Join-Path ([System.IO.Path]::GetTempPath()) ("proactiva-" + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $exchangeFolder -Force | Out-Null

    $worker = {
        param($ModulesFolder, $CollectionModule, $Server, $Credential, $WithCis, $Command, $InputPath, $OutputPath)

        # Same policy the launcher gives the main process (the files may carry the "downloaded
        # from the internet" mark); process scope only
        Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
        $ErrorActionPreference = 'Continue'
        $connection = $null
        $cisConnection = $null
        # Everything the work writes, in order, as plain entries the main process can show
        $log = [System.Collections.Generic.List[object]]::new()
        $outcome = $null
        try {
            foreach ($module in 'VMware.VimAutomation.Sdk', 'VMware.VimAutomation.Common', 'VMware.VimAutomation.Cis.Core', 'VMware.VimAutomation.Core', 'VMware.VimAutomation.Vds') {
                Import-Module (Join-Path $ModulesFolder $module) -ErrorAction Stop -WarningAction SilentlyContinue
            }
            Set-PowerCLIConfiguration -Scope Session -ParticipateInCEIP $false -InvalidCertificateAction Ignore -Confirm:$false | Out-Null
            Import-Module $CollectionModule -ErrorAction Stop -WarningAction SilentlyContinue

            $connection = Connect-VIServer -Server $Server -Credential $Credential -ErrorAction Stop
            if ($WithCis) {
                try { $cisConnection = Connect-CisServer -Server $Server -Credential $Credential -ErrorAction Stop }
                catch { $log.Add([PSCustomObject]@{ Kind = 'Warning'; Text = "`t-> La conexion CIS (REST) fallo: $($_.Exception.Message)" }) }
            }

            $arguments = Import-Clixml -Path $InputPath
            $result = foreach ($item in (& $Command -Connection $connection -CisConnection $cisConnection @arguments *>&1)) {
                if ($item -is [System.Management.Automation.InformationRecord]) {
                    $message = $item.MessageData
                    if ($message -is [System.Management.Automation.HostInformationMessage]) {
                        $log.Add([PSCustomObject]@{ Kind = 'Host'; Text = [string]$message.Message; NoNewLine = [bool]$message.NoNewLine; Color = if ($null -ne $message.ForegroundColor) { [string]$message.ForegroundColor } else { $null } })
                    }
                    else { $log.Add([PSCustomObject]@{ Kind = 'Host'; Text = [string]$message; NoNewLine = $false; Color = $null }) }
                }
                elseif ($item -is [System.Management.Automation.WarningRecord]) { $log.Add([PSCustomObject]@{ Kind = 'Warning'; Text = $item.Message }) }
                elseif ($item -is [System.Management.Automation.ErrorRecord]) { $log.Add([PSCustomObject]@{ Kind = 'Error'; Text = $item.ToString() }) }
                elseif ($item -is [System.Management.Automation.VerboseRecord] -or $item -is [System.Management.Automation.DebugRecord]) { }
                else { $item }
            }
            $outcome = [PSCustomObject]@{ Ok = $true; Result = $result; Error = $null; Log = $log }
        }
        catch {
            $outcome = [PSCustomObject]@{ Ok = $false; Result = $null; Error = $_.Exception.Message; Log = $log }
        }
        finally {
            if ($connection) { Disconnect-VIServer -Server $connection -Force -Confirm:$false -ErrorAction SilentlyContinue }
            if ($cisConnection) { Disconnect-CisServer -Server $cisConnection -Force -Confirm:$false -ErrorAction SilentlyContinue }
        }
        $outcome | Export-Clixml -Path $OutputPath -Depth 20
    }

    try {
        # One job per task (none when there is no credential to open the worker's session)
        $jobs = [System.Collections.Generic.List[object]]::new()
        for ($i = 0; $i -lt $Tasks.Count; $i++) {
            $task = $Tasks[$i]
            if (-not $task.VCenter.Credential) { $jobs.Add($null); continue }

            $inputPath = Join-Path $exchangeFolder "input-$i.xml"
            $task.Arguments | Export-Clixml -Path $inputPath -Depth 20
            $jobs.Add((Start-Job -ScriptBlock $worker -ArgumentList @(
                $modulesFolder, $collectionModule, $task.VCenter.Connection.Name, $task.VCenter.Credential,
                $task.WithCis, $task.Command, $inputPath, (Join-Path $exchangeFolder "output-$i.xml")
            )))
        }

        for ($i = 0; $i -lt $Tasks.Count; $i++) {
            $job = $jobs[$i]
            if ($null -eq $job) {
                [PSCustomObject]@{ Ok = $false; Result = $null; Error = "No hay credencial guardada para '$($Tasks[$i].VCenter.Name)'" }
                continue
            }

            # A worker that hangs (network, vCenter not answering) must not hang the whole run:
            # after the timeout it is stopped and its work is redone in sequence by the caller
            if (-not (Wait-Job -Job $job -Timeout $TimeoutSeconds)) {
                Stop-Job -Job $job
                Remove-Job -Job $job -Force
                [PSCustomObject]@{ Ok = $false; Result = $null; Error = "Tiempo agotado ($([math]::Round($TimeoutSeconds / 60)) min)" }
                continue
            }

            Write-Host "`n--- $($Tasks[$i].Label) ---" -ForegroundColor DarkCyan
            $jobErrors = $null
            Receive-Job -Job $job -Wait -AutoRemoveJob -ErrorAction Continue -ErrorVariable jobErrors | Out-Null

            $outputPath = Join-Path $exchangeFolder "output-$i.xml"
            if (Test-Path $outputPath) {
                $outcome = Import-Clixml -Path $outputPath
                Write-WorkerLog $outcome.Log
                [PSCustomObject]@{ Ok = [bool]$outcome.Ok; Result = $outcome.Result; Error = $outcome.Error }
            }
            else {
                [PSCustomObject]@{ Ok = $false; Result = $null; Error = "La tarea no devolvio resultado. $($jobErrors | Select-Object -First 1)" }
            }
        }
    }
    finally {
        Get-Job | Where-Object { $_.State -eq 'Running' -and $jobs -contains $_ } | Stop-Job
        Remove-Item -Path $exchangeFolder -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# Shows what a worker wrote, as it would have looked running here
function Write-WorkerLog($log) {
    foreach ($entry in $log) {
        switch ($entry.Kind) {
            'Host' {
                $hostParams = @{ Object = $entry.Text; NoNewline = [bool]$entry.NoNewLine }
                if ($entry.Color) { $hostParams['ForegroundColor'] = $entry.Color }
                Write-Host @hostParams
            }
            'Warning' { Write-Warning $entry.Text }
            'Error' { Write-Host $entry.Text -ForegroundColor Red }
        }
    }
}

# ---------------------------------------------------------------------------------------
# Timing and output
# ---------------------------------------------------------------------------------------

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
        Write-Host ("`t  {0,-34} {1,9} {2,5:0}%" -f $entry.Key, (Format-Duration $entry.Value), $share)
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

# The *Worker functions are called by the worker processes, which import this module
Export-ModuleMember -Function Connect-VCenters, Disconnect-VCenters, Invoke-ProactivaCollection, Invoke-DriverCollection, Invoke-VCenterCollectionWorker, Invoke-HostDriversWorker
