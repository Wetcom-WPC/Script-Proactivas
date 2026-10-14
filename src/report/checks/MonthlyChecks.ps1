# Monthly checks (VSP-ME-*). Dot-sourced by Report.psm1.
#
# Each check receives the report context, reads one or more raw Excel sheets, writes a detail
# sheet to the annex when it finds something, and returns one checklist result per ID:
#   [PSCustomObject]@{ ID; Resultado; Detalle }
# Annex sheet names, their columns and the checklist texts are what the client receives:
# change them only on purpose.

function Test-VCenterSizing($Context) {
    $sizingWeights = @{
        'tiny'    = 1; 'small'   = 2; 'medium'  = 3
        'large'   = 4; 'x-large' = 5
    }

    # Sizing requirements per vCenter version (newer versions need more RAM for the same
    # tier). Same source the collection uses to classify "Sizing actual".
    $sizingByVersion = Get-Content -Path (Join-Path $Context.ReferenceDataFolder "vcenter-sizing.json") -Raw | ConvertFrom-Json -AsHashtable

    $vmStates = @{}
    foreach ($vm in (Import-RawSheet $Context.RawExcelPath "VM")) {
        if (-not $vmStates.ContainsKey($vm.VM)) { $vmStates.Add($vm.VM, $vm.State) }
    }
    $vcenterRows = Import-RawSheet $Context.RawExcelPath "vCenter"
    $sizingRows = Import-RawSheet $Context.RawExcelPath "Sizing"

    $combined = @()
    foreach ($row in $sizingRows) {
        $state = $vmStates[$row.VM]
        if ($state -ne "PoweredOff") {
            $vcenterInfo = $vcenterRows | Where-Object { $_.'vCenter Server' -eq $row.vCenter } | Select-Object -First 1
            $combined += [PSCustomObject]@{
                State = $state; VM = $row.VM; vCenter = $row.vCenter
                Version = if ($vcenterInfo) { $vcenterInfo.Version } else { "Not Found" }
                Build = if ($vcenterInfo) { $vcenterInfo.Build } else { "Not Found" }
                "Sizing actual" = $row."Sizing actual"
                "Sizing recomendado" = $row."Sizing recomendado"
                "Cantidad de VMs" = $row."Cantidad de VMs"
                "vCPU" = $row."vCPU"; "Memory GB" = $row."Memory GB"
            }
        }
    }

    # One row per vCenter VM: the one with the largest current tier
    $uniqueVCenters = $combined | Group-Object -Property @{ Expression = { $_.vCenter + '|' + $_.VM } } | ForEach-Object {
        $_.Group | Sort-Object -Property @{
            Expression = {
                $sizing = ""; if (-not [string]::IsNullOrEmpty($_."Sizing actual")) { $sizing = $_."Sizing actual".ToLower().Trim() }
                if ($sizingWeights.ContainsKey($sizing)) { return $sizingWeights[$sizing] } else { return 0 }
            }
        } | Select-Object -Last 1
    }

    $findings = @()

    foreach ($vcenter in $uniqueVCenters) {
        $majorVersion = ($vcenter.Version -split '\.')[0]
        $requirements = $sizingByVersion[$majorVersion]

        $currentSizing = ""; if (-not [string]::IsNullOrEmpty($vcenter."Sizing actual")) { $currentSizing = $vcenter."Sizing actual".ToLower().Trim() }

        if ($requirements -and $requirements.ContainsKey($currentSizing)) {
            $required = $requirements[$currentSizing]

            $cores = 0; $memoryGB = 0.0; $vmCount = 0
            [void][int]::TryParse($vcenter.'vCPU', [ref]$cores)
            [void][double]::TryParse(($vcenter.'Memory GB'.ToString()).Replace(',','.'), [ref]$memoryGB)
            [void][int]::TryParse($vcenter.'Cantidad de VMs', [ref]$vmCount)

            if (($cores -lt $required.vcpus) -or ($memoryGB -lt $required.ramGB) -or ($vmCount -gt $required.vmsMax)) {
                $findings += $vcenter
            }
        }
    }

    if ($findings) {
        $annexRows = $findings | Select-Object -Property @(
            'State', 'VM', 'vCenter', 'Version', 'Build',
            'Sizing actual',
            @{ Name = 'VMs';          Expression = { $_.'Cantidad de VMs' } },
            @{ Name = 'Cores';        Expression = { $_.'vCPU' } },
            @{ Name = 'Memoria (GB)'; Expression = { $_.'Memory GB' } }
        )
        Export-AnnexSheet -Data $annexRows -Path $Context.AnnexPath -WorksheetName "Sizing Incorrecto" -NoNumberConversion "VM", "vCenter", "Version"

        return [PSCustomObject]@{ ID = "VSP-ME-01"; Resultado = "No recomendado"; Detalle = "Se encontraron $($findings.Count) vCenters mal dimensionados" }
    }

    return [PSCustomObject]@{ ID = "VSP-ME-01"; Resultado = "Resultado Esperado"; Detalle = "Todos los vCenter Servers se encuentran dimensionados correctamente" }
}

function Test-PartitionFreeSpace($Context) {
    $findings = @()

    foreach ($row in (Import-RawSheet $Context.RawExcelPath "Partitions")) {
        if ($row.PSObject.Properties['Disk'] -and $row.PSObject.Properties['Free %']) {
            $freePercent = 0
            [void][double]::TryParse((([string]$row."Free %").Replace(',', '.') -replace '[^0-9.]'), [ref]$freePercent)
            $diskPath = $row.Disk.Trim()

            if (($freePercent -lt 30) -and ($diskPath -ne "/storage/core" -and $diskPath -ne "/storage/archive")) {
                $findings += [PSCustomObject]@{
                    "VM"         = $row."VM"
                    "Annotation" = $row."Annotation"
                    "Disk"       = $row."Disk"
                    "Free %"     = $row."Free %"
                }
            }
        }
    }

    if ($findings.Count -gt 0) {
        Export-AnnexSheet -Data $findings -Path $Context.AnnexPath -WorksheetName "Espacio en Particiones"
        return [PSCustomObject]@{ ID = "VSP-ME-02"; Resultado = "No recomendado"; Detalle = "Se encontraron $($findings.Count) particiones con espacio libre inferior al 30%" }
    }

    return [PSCustomObject]@{ ID = "VSP-ME-02"; Resultado = "Resultado Esperado"; Detalle = "Todas las particiones de disco analizadas tienen un espacio libre superior al 30%" }
}

function Test-SyslogConfiguration($Context) {
    $findings = @()

    foreach ($row in (Import-RawSheet $Context.RawExcelPath "ESXi")) {
        if ($row.PSObject.Properties['SyslogGlobalLogDir'] -and $row.PSObject.Properties['SyslogGlobalLogHost']) {
            $logDir = $row.SyslogGlobalLogDir
            $logHost = $row.SyslogGlobalLogHost

            $logDirIsWrong = ($logDir -match "/scratch/log" -or $logDir -match "local" -or [string]::IsNullOrEmpty($logDir))
            $logHostIsWrong = ($logHost -match "/scratch/log" -or $logHost -match "local" -or [string]::IsNullOrEmpty($logHost))

            if ($logDirIsWrong -and $logHostIsWrong) {
                $findings += [PSCustomObject]@{
                    "vCenter"             = $row."vCenter"
                    "Hostname"            = $row."Hostname"
                    "Datacenter"          = $row."Datacenter"
                    "Cluster"             = $row."Cluster"
                    "SyslogGlobalLogDir"  = $logDir
                    "SyslogGlobalLogHost" = $logHost
                }
            }
        }
    }

    if ($findings.Count -gt 0) {
        Export-AnnexSheet -Data $findings -Path $Context.AnnexPath -WorksheetName "Syslogs y ubicacion de logs" -NoNumberConversion "Hostname"
        return [PSCustomObject]@{ ID = "VSP-ME-04"; Resultado = "No configurado"; Detalle = "Se encontraron $($findings.Count) hosts sin configuracion de Syslog remoto" }
    }

    return [PSCustomObject]@{ ID = "VSP-ME-04"; Resultado = "Resultado Esperado"; Detalle = "Todos los hosts tienen una configuracion de Syslog valida" }
}

function Test-MultipathPolicy($Context) {
    $findings = @()

    foreach ($row in (Import-RawSheet $Context.RawExcelPath "Datastores")) {
        if (-not ($row.PSObject.Properties['Datastore'] -and $row.PSObject.Properties['Policy'] -and $row.PSObject.Properties['Hostname'])) { continue }

        $datastoreName = $row.Datastore
        $hostShortName = ($row.Hostname -split '\.')[0]

        # A datastore named after its host is the host's local datastore: skip it
        if ($datastoreName -match $hostShortName) { continue }

        $policy = $row.Policy.Trim()
        if ($policy -eq 'vSAN') { continue }

        $isLocal = ($datastoreName -match "local" -or $datastoreName -match "datastore1")
        $recommendedPolicy = ""

        if (-not $isLocal -and $policy -ne "RoundRobin") {
            $recommendedPolicy = "RoundRobin"
        }
        if ($isLocal -and $policy -eq "RoundRobin") {
            $recommendedPolicy = "MRU o Fixed"
        }

        if ($recommendedPolicy) {
            $findings += [PSCustomObject]@{
                "vCenter"             = $row."vCenter"
                "Hostname"            = $row."Hostname"
                "Datastore"           = $datastoreName
                "Policy"              = $policy
                "Policy recomendado"  = $recommendedPolicy
            }
        }
    }

    if ($findings.Count -gt 0) {
        Export-AnnexSheet -Data $findings -Path $Context.AnnexPath -WorksheetName "Politicas de Multipath" -NoNumberConversion "Hostname"
        return [PSCustomObject]@{ ID = "VSP-ME-05"; Resultado = "No recomendado"; Detalle = "Se encontraron $($findings.Count) datastores con politicas de Multipath no recomendadas" }
    }

    return [PSCustomObject]@{ ID = "VSP-ME-05"; Resultado = "Resultado Esperado"; Detalle = "Todos los datastores analizados cumplen con las politicas de Multipath recomendadas" }
}

function Test-EsxiVersionConsistency($Context) {
    $findings = @()

    foreach ($clusterHosts in (Get-HostsByCluster (Import-RawSheet $Context.RawExcelPath "ESXi")).Values) {
        $versions = $clusterHosts | ForEach-Object { if ($_.EsxiVersion) { $_.EsxiVersion.Trim() } } | Select-Object -Unique
        if ($versions.Count -gt 1) {
            $findings += $clusterHosts
        }
    }

    if ($findings.Count -gt 0) {
        $annexRows = $findings | ForEach-Object {
            [PSCustomObject]@{
                "vCenter"     = $_."vCenter"
                "Hostname"    = $_."Hostname"
                "Datacenter"  = $_."Datacenter"
                "Cluster"     = $_."Cluster"
                "EsxiVersion" = $_."EsxiVersion"
            }
        }
        Export-AnnexSheet -Data $annexRows -Path $Context.AnnexPath -WorksheetName "Consistencias de versiones" -NoNumberConversion "Hostname"
        return [PSCustomObject]@{ ID = "VSP-ME-06"; Resultado = "No recomendado"; Detalle = "Se encontraron clusteres con versiones de ESXi inconsistentes" }
    }

    return [PSCustomObject]@{ ID = "VSP-ME-06"; Resultado = "Resultado Esperado"; Detalle = "Todos los clusteres analizados tienen versiones de ESXi consistentes" }
}

function Test-HostResourceConsistency($Context) {
    $memoryToleranceGB = 1
    $cpuSpeedToleranceMHz = 100

    $findings = @()

    foreach ($clusterHosts in (Get-HostsByCluster (Import-RawSheet $Context.RawExcelPath "ESXi")).Values) {
        $memories = @(); $speeds = @(); $models = @()

        foreach ($h in $clusterHosts) {
            $memories += [double](([string]$h.MemoryGB) -replace '[^0-9\.,]', '' -replace ',', '.')
            $speeds += [double](([string]$h.CpuSpeed) -replace '[^0-9\.,]', '' -replace ',', '.')
            $models += $h.CpuModel.ToString().Trim()
        }

        $memoryRange = ($memories | Measure-Object -Maximum).Maximum - ($memories | Measure-Object -Minimum).Minimum
        $speedRange = ($speeds | Measure-Object -Maximum).Maximum - ($speeds | Measure-Object -Minimum).Minimum

        $isInconsistent = (($models | Select-Object -Unique).Count -gt 1) -or
                          ($memoryRange -gt $memoryToleranceGB) -or
                          ($speedRange -gt $cpuSpeedToleranceMHz)

        if ($isInconsistent) {
            $findings += $clusterHosts
        }
    }

    if ($findings.Count -gt 0) {
        $annexRows = $findings | ForEach-Object {
            [PSCustomObject]@{
                "vCenter"  = $_."vCenter"
                "Hostname" = $_."Hostname"
                "Cluster"  = $_."Cluster"
                "MemoryGB" = $_."MemoryGB"
                "CpuModel" = $_."CpuModel"
                "CpuSpeed" = $_."CpuSpeed"
            }
        }
        Export-AnnexSheet -Data $annexRows -Path $Context.AnnexPath -WorksheetName "Consistencias de recursos" -NoNumberConversion "Hostname"
        return [PSCustomObject]@{ ID = "VSP-ME-07"; Resultado = "No recomendado"; Detalle = "Se encontraron clusteres con inconsistencias de recursos de hardware" }
    }

    return [PSCustomObject]@{ ID = "VSP-ME-07"; Resultado = "Resultado Esperado"; Detalle = "Todos los clusteres analizados tienen recursos de hardware consistentes" }
}

# Windows VMs (2008 / Windows 7 onwards) with an E1000/E1000e/Flexible adapter
function Test-VmNicAdapterType($Context) {
    $findings = @()

    foreach ($row in (Import-RawSheet $Context.RawExcelPath "VM")) {
        if ($null -eq $row -or -not $row.PSObject.Properties['SO (vCenter)'] -or -not $row.PSObject.Properties['VM']) { continue }

        $guestOs = $row."SO (vCenter)".Trim()
        $isModernWindows = ($guestOs -like "Microsoft Windows*") -and ($guestOs -notlike "*2003*") -and ($guestOs -notlike "*2000*")
        if (-not $isModernWindows) { continue }

        # Replicas are not reported
        if ($row.VM -match "_replica|_rep|_cont") { continue }

        $hasE1000 = $false
        for ($i = 1; $i -le 10; $i++) {
            $adapterColumn = "Adapter_{0:D2}" -f $i
            if ($row.PSObject.Properties[$adapterColumn]) {
                $adapterType = ([string]$row.$adapterColumn).Trim()
                if ($adapterType -eq "e1000" -or $adapterType -eq "e1000e" -or $adapterType -eq "Flexible") {
                    $hasE1000 = $true
                    break
                }
            }
        }

        if ($hasE1000) {
            $findings += $row
        }
    }

    if ($findings.Count -eq 0) {
        return [PSCustomObject]@{ ID = "VSP-ME-08"; Resultado = "Resultado Esperado"; Detalle = "No se encontraron VMs Windows (excluyendo réplicas) con adaptadores E1000/E1000e." }
    }

    if ($findings.ToolsStatus -contains 'toolsNotInstalled') {
        $result = "VMware Tools no instaladas"; $detail = "Se encontraron $($findings.Count) VMs con adaptadores E1000, y al menos una no tiene VMware Tools instaladas."
    } elseif ($findings | Where-Object { $_.ToolsStatus -ne 'toolsOk' }) {
        $result = "VMware Tools desactualizadas"; $detail = "Se encontraron $($findings.Count) VMs con adaptadores E1000, y al menos una no tiene VMware Tools actualizadas."
    } else {
        $result = "VMware Tools OK"; $detail = "Se encontraron $($findings.Count) VMs con adaptadores E1000, pero todas tienen VMware Tools actualizadas."
    }

    $annexRows = $findings | ForEach-Object {
        [PSCustomObject]@{
            "vCenter"      = $_."vCenter"; "VM" = $_."VM"; "Cluster" = $_."Cluster"; "Host" = $_."Host"
            "State"        = $_."State"; "ToolsStatus" = $_."ToolsStatus"; "SO (vCenter)" = $_."SO (vCenter)"
            "SO (Tools)"   = $_."SO (Tools)"; "Adapter_01" = $_."Adapter_01"; "Adapter_02" = $_."Adapter_02"
            "Adapter_03"   = $_."Adapter_03"; "Adapter_04" = $_."Adapter_04"; "Adapter_05" = $_."Adapter_05"
            "Adapter_06"   = $_."Adapter_06"; "Adapter_07" = $_."Adapter_07"; "Adapter_08" = $_."Adapter_08"
            "Adapter_09"   = $_."Adapter_09"; "Adapter_10" = $_."Adapter_10"
        }
    }
    Export-AnnexSheet -Data $annexRows -Path $Context.AnnexPath -WorksheetName "Hardware Placa de Red" -NoNumberConversion "Host"

    return [PSCustomObject]@{ ID = "VSP-ME-08"; Resultado = $result; Detalle = $detail }
}

# ESXi Shell timeouts: 0 is not recommended, outside 300-1800s is out of the recommended range
function Test-EsxiShellTimeout($Context) {
    $findings = @()
    $foundZero = $false
    $foundOutOfRange = $false

    foreach ($row in (Import-RawSheet $Context.RawExcelPath "ESXi")) {
        if (-not ($row.PSObject.Properties['ESXIShellTimeOut'] -and $row.PSObject.Properties['ESXIShellinteractiveTimeOut'])) { continue }

        $timeout = $row.ESXIShellTimeOut
        $interactiveTimeout = $row.ESXIShellinteractiveTimeOut
        $isWrong = $false

        if ($timeout -eq 0 -or $interactiveTimeout -eq 0) {
            $foundZero = $true
            $isWrong = $true
        } elseif (($timeout -lt 300 -or $timeout -gt 1800) -or ($interactiveTimeout -lt 300 -or $interactiveTimeout -gt 1800)) {
            $foundOutOfRange = $true
            $isWrong = $true
        }

        if ($isWrong) {
            $findings += [PSCustomObject]@{
                "vCenter"                     = $row."vCenter"
                "Hostname"                    = $row."Hostname"
                "Datacenter"                  = $row."Datacenter"
                "Cluster"                     = $row."Cluster"
                "ESXIShellTimeOut"            = $timeout
                "ESXIShellinteractiveTimeOut" = $interactiveTimeout
            }
        }
    }

    $result = "Resultado Esperado"
    $detail = "Todos los hosts tienen la configuracion de TSM dentro del rango recomendado (300-1800s)"

    if ($foundZero) {
        $result = "No recomendado"
        $detail = "Se encontraron hosts con timeout de TSM configurado en 0 segundos"
    } elseif ($foundOutOfRange) {
        $result = "Fuera de valor recomendado"
        $detail = "Se encontraron hosts con timeout de TSM configurado fuera del rango recomendado (300-1800s)"
    }

    if ($findings.Count -gt 0) {
        Export-AnnexSheet -Data $findings -Path $Context.AnnexPath -WorksheetName "Verificacion de TSM" -NoNumberConversion "Hostname"
    }

    return [PSCustomObject]@{ ID = "VSP-ME-12"; Resultado = $result; Detalle = $detail }
}

function Test-PowerManagementPolicy($Context) {
    $findings = @()

    foreach ($row in (Import-RawSheet $Context.RawExcelPath "ESXi")) {
        if ($null -ne $row -and $row.PSObject.Properties['PowerManagement'] -and $row.PowerManagement.Trim() -ne "High performance") {
            $findings += [PSCustomObject]@{
                "vCenter"         = $row."vCenter"
                "Hostname"        = $row."Hostname"
                "Datacenter"      = $row."Datacenter"
                "Cluster"         = $row."Cluster"
                "PowerManagement" = $row."PowerManagement"
            }
        }
    }

    if ($findings.Count -gt 0) {
        Export-AnnexSheet -Data $findings -Path $Context.AnnexPath -WorksheetName "Power Management" -NoNumberConversion "Hostname"
        return [PSCustomObject]@{ ID = "VSP-ME-14"; Resultado = "No recomendado"; Detalle = "Se encontraron $($findings.Count) hosts con una politica de Power Management no recomendada" }
    }

    return [PSCustomObject]@{ ID = "VSP-ME-14"; Resultado = "Resultado Esperado"; Detalle = "Todos los hosts analizados tienen la politica de Power Management configurada en High Performance" }
}

function Test-VmwareToolsStatus($Context) {
    $findings = @(Import-RawSheet $Context.RawExcelPath "VM" | Where-Object {
        $status = $_.ToolsStatus.Trim()
        $state = $_.State.Trim()
        # Powered-off VMs with tools "not running" are fine
        $hasToolsProblem = -not (($status -eq "toolsOk") -or (($state -eq "PoweredOff") -and ($status -eq "toolsNotRunning")))
        # Replicas are not reported
        $isNotReplica = $_.VM -notmatch "(_replica|_rep|_cont)"

        $hasToolsProblem -and $isNotReplica
    })

    if ($findings.Count -eq 0) {
        return [PSCustomObject]@{ ID = "VSP-ME-15"; Resultado = "Resultado Esperado"; Detalle = "Todas las VMs tienen las VMware Tools en estado 'OK'" }
    }

    # A cluster is inconsistent when its affected VMs run more than one Tools version
    $isClusterInconsistent = $false
    foreach ($clusterName in ($findings.Cluster | Select-Object -Unique)) {
        $toolsVersions = $findings | Where-Object { $_.Cluster -eq $clusterName } | Select-Object -ExpandProperty ToolsVersion -Unique
        if ($toolsVersions.Count -gt 1) {
            $isClusterInconsistent = $true
            break
        }
    }

    if ($isClusterInconsistent) {
        $result = "Cluster no consistente"
        $detail = "Se encontraron VMs con problemas de VMware Tools en clusteres con versiones de Tools inconsistentes"
    } else {
        $result = "Cluster consistente"
        $detail = "Se encontraron VMs con problemas de VMware Tools, y las versiones de Tools son consistentes dentro de cada cluster"
    }

    $annexRows = $findings | ForEach-Object {
        [PSCustomObject]@{
            "vCenter"              = $_."vCenter"
            "VM"                   = $_."VM"
            "Cluster"              = $_."Cluster"
            "Host"                 = $_."Host"
            "ConnectionState"      = $_."ConnectionState"
            "State"                = $_."State"
            "ToolsStatus"          = $_."ToolsStatus"
            "ToolsVersion"         = $_."ToolsVersion"
            "ToolsRequiredVersion" = $_."ToolsRequiredVersion"
        }
    }
    Export-AnnexSheet -Data $annexRows -Path $Context.AnnexPath -WorksheetName "VMware Tools" -NoNumberConversion "Host"

    return [PSCustomObject]@{ ID = "VSP-ME-15"; Resultado = $result; Detalle = $detail }
}

function Test-ConnectedIsos($Context) {
    $findings = @()

    foreach ($row in (Import-RawSheet $Context.RawExcelPath "VM")) {
        if (-not $row.PSObject.Properties['IsoConnected']) { continue }

        $isoPath = $row.IsoConnected
        if (-not [string]::IsNullOrEmpty($isoPath) -and $isoPath.Trim() -ne "[]") {
            $findings += [PSCustomObject]@{
                "vCenter"      = $row."vCenter"
                "VM"           = $row."VM"
                "Cluster"      = $row."Cluster"
                "Host"         = $row."Host"
                "IsoConnected" = $row."IsoConnected"
            }
        }
    }

    if ($findings.Count -gt 0) {
        Export-AnnexSheet -Data $findings -Path $Context.AnnexPath -WorksheetName "ISOs Conectadas" -NoNumberConversion "Host"
        return [PSCustomObject]@{ ID = "VSP-ME-16"; Resultado = "No recomendado"; Detalle = "Se encontraron $($findings.Count) VMs con ISOs conectadas" }
    }

    return [PSCustomObject]@{ ID = "VSP-ME-16"; Resultado = "Resultado Esperado"; Detalle = "No se encontraron VMs con ISOs conectadas." }
}

# Connected network cards of powered-on VMs that do not have "Connect At Power On" enabled
function Test-VmNicStartConnected($Context) {
    $findings = @()

    foreach ($row in (Import-RawSheet $Context.RawExcelPath "vNetwork")) {
        if (-not ($row.PSObject.Properties['Status'] -and $row.PSObject.Properties['Connected'] -and $row.PSObject.Properties['StartsConnected'])) { continue }

        if (([string]$row.Status -eq "1") -and ($row.Connected -eq "True") -and ($row.StartsConnected -eq "False")) {
            $findings += [PSCustomObject]@{
                "vCenter"         = $row."vCenter"
                "VM"              = $row."VM"
                "Cluster"         = $row."Cluster"
                "Host"            = $row."Host"
                "Status"          = $row."Status"
                "Mac"             = $row."Mac"
                "Connected"       = $row."Connected"
                "StartsConnected" = $row."StartsConnected"
            }
        }
    }

    if ($findings.Count -gt 0) {
        Export-AnnexSheet -Data $findings -Path $Context.AnnexPath -WorksheetName "Placas No Inician Conectadas" -NoNumberConversion "Host"
        return [PSCustomObject]@{ ID = "VSP-ME-17"; Resultado = "No recomendado"; Detalle = "Se encontraron $($findings.Count) placas de red no configuradas para iniciar conectadas" }
    }

    return [PSCustomObject]@{ ID = "VSP-ME-17"; Resultado = "Resultado Esperado"; Detalle = "Todas las placas de red de VMs estan configuradas para iniciar conectadas." }
}

# Returns two results: VSP-ME-18 (snapshots older than 14 days) and VSP-ME-03 (size of the
# largest old snapshot, limit 300 GB)
function Test-Snapshots($Context) {
    $findings = @()
    $cutoffDate = (Get-Date).AddDays(-14)

    # Replicas and backup restore points are not reported
    $snapshots = Import-RawSheet $Context.RawExcelPath "Snapshot" |
        Where-Object { $_.VM -notmatch "(_replica|_rep|_cont)" } |
        Where-Object { $_.Snapshot -notmatch "Restore Point" }

    foreach ($row in $snapshots) {
        if (-not ($row.PSObject.Properties['Fecha'] -and $row.PSObject.Properties['SizeMB'])) { continue }
        try {
            if ([datetime]$row.Fecha -lt $cutoffDate) {
                $findings += $row
            }
        } catch {
            Write-Warning "No se pudo convertir la fecha '$($row.Fecha)' para la VM '$($row.VM)'. Se omite esta fila."
        }
    }

    if ($findings.Count -gt 0) {
        Export-AnnexSheet -Data $findings -Path $Context.AnnexPath -WorksheetName "Snapshots Antiguos"
        $oldSnapshotsResult = [PSCustomObject]@{ ID = "VSP-ME-18"; Resultado = "No recomendado"; Detalle = "Se encontraron $($findings.Count) snapshots con mas de 14 dias" }
    } else {
        $oldSnapshotsResult = [PSCustomObject]@{ ID = "VSP-ME-18"; Resultado = "Resultado Esperado"; Detalle = "No se encontraron snapshots con mas de 14 dias." }
    }

    $totalSizeGB = [math]::Round((($findings | Measure-Object -Property SizeMB -Sum).Sum) / 1024, 2)
    $largest = $findings | Sort-Object -Property @{ Expression = { [double]$_.SizeMB } } | Select-Object -Last 1

    if ($largest) {
        $largestGB = [math]::Round(([double]$largest.SizeMB / 1024), 2)

        if ([double]$largest.SizeMB -gt (300 * 1024)) {
            $sizeResult = [PSCustomObject]@{ ID = "VSP-ME-03"; Resultado = "No recomendado"; Detalle = "Se encontro un snapshot ('$($largest.Snapshot)') de $($largestGB)GB, que supera los 300GB. Total: $totalSizeGB GB" }
        } else {
            $sizeResult = [PSCustomObject]@{ ID = "VSP-ME-03"; Resultado = "Resultado Esperado"; Detalle = "Ningun snapshot supera los 300GB. El mas grande es de $($largestGB)GB. Total: $totalSizeGB GB" }
        }
    } else {
        $sizeResult = [PSCustomObject]@{ ID = "VSP-ME-03"; Resultado = "Resultado Esperado"; Detalle = "No existen snapshots antiguos." }
    }

    return $sizeResult, $oldSnapshotsResult
}

# "Falso Positivo" test results plus the audit of the expected alarms
function Test-AlarmConfiguration($Context) {
    $rows = @(Import-RawSheet $Context.RawExcelPath "Falso Positivo" | Where-Object { -not [string]::IsNullOrEmpty($_.vCenter) })

    if ($rows.Count -eq 0) {
        return [PSCustomObject]@{ ID = "VSP-ME-13"; Resultado = "Sin acceso"; Detalle = "No se encontraron registros de la prueba de alarmas (hoja vacía o no encontrada)." }
    }

    $findings = @()
    $successCount = 0
    $failureCount = 0
    $causes = @()
    $scriptArguments = @()
    $testSkipped = $false

    foreach ($row in $rows) {
        $result = if ($row.Result) { $row.Result.ToString().Trim() } else { "" }
        $alarmName = if ($row."Alarm Name") { $row."Alarm Name" } else { $row."Alarm Source" }
        $scriptPath = if ($row."Script Path") { $row."Script Path" } else { $row."Alarm Path" }
        if (-not $scriptPath) { $scriptPath = $row."Path Alarma" }

        # Run with -SkipAlarmTest: the "Falso Positivo" test did not run, which is not a failure
        if ($result -like "Omitida*") {
            $testSkipped = $true
            continue
        }

        # Success: OK, SUCCESS, or messages starting with SUCCESS
        if ($result -eq "OK" -or $result -match "^SUCCESS") {
            $successCount++

            # Script arguments (tokens) of the successful ones, to detect inconsistent configs
            if (-not [string]::IsNullOrEmpty($scriptPath) -and $scriptPath -ne "N/A") {
                $parts = $scriptPath -split "\s+"
                if ($parts.Count -gt 1) {
                    $scriptArguments += ($parts[1..($parts.Count - 1)]) -join " "
                }
            }
            continue
        }

        $failureCount++
        switch -Wildcard ($result) {
            "*No Script Configured*" {
                $diagnosis = "Configuración: Alarma sin acción de script"
                if ("Configuración incompleta" -notin $causes) { $causes += "Configuración incompleta" }
            }
            "*Script Action Empty*" {
                $diagnosis = "Configuración: Acción de script vacía"
                if ("Configuración incompleta" -notin $causes) { $causes += "Configuración incompleta" }
            }
            "*Not Found*" {
                $diagnosis = "Alarma no encontrada en vCenter"
                if ("Alarmas faltantes" -notin $causes) { $causes += "Alarmas faltantes" }
            }
            "*No hay hosts*" {
                $diagnosis = "Sin hosts conectados"
                if ("Entorno no disponible" -notin $causes) { $causes += "Entorno no disponible" }
            }
            "*Failed*" {
                $diagnosis = "Fallo funcional (Smoke Test)"
                if ("Fallo funcional" -notin $causes) { $causes += "Fallo funcional" }
            }
            Default {
                $diagnosis = "Error: $result"
                if ("Errores varios" -notin $causes) { $causes += "Errores varios" }
            }
        }

        $findings += [PSCustomObject]@{
            "vCenter"       = $row.vCenter
            "Host"          = $row.Host
            "Alarma"        = $alarmName
            "Script Path"   = $scriptPath
            "Resultado"     = $result
            "Diagnóstico"   = $diagnosis
            "Timestamp"     = $row.Timestamp
        }
    }

    $checklistResult = "Resultado Esperado"
    $detail = "Se verificaron $successCount alarmas/pruebas correctamente."

    $inconsistencyNote = ""
    $uniqueArguments = @($scriptArguments | Select-Object -Unique)
    if ($uniqueArguments.Count -gt 1) {
        $inconsistencyNote = " (Nota: Se detectaron $($uniqueArguments.Count) variantes de tokens/argumentos)."
        if ($failureCount -eq 0) { $detail += $inconsistencyNote }
    }

    if ($failureCount -gt 0) {
        $checklistResult = "No recomendado"
        $detail = "Se detectaron $failureCount anomalías. Causas: $($causes -join ", ").$inconsistencyNote"
    }

    if ($testSkipped) {
        $detail += " La prueba de Falso Positivo no se ejecuto (corrida con -SkipAlarmTest): solo se audito la configuracion de las alarmas."
    }

    if ($findings.Count -gt 0) {
        Export-AnnexSheet -Data $findings -Path $Context.AnnexPath -WorksheetName "Falso Positivo"
    }

    return [PSCustomObject]@{ ID = "VSP-ME-13"; Resultado = $checklistResult; Detalle = $detail }
}
