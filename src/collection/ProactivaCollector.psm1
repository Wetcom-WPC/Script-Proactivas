# Collects everything the Proactiva report needs from one or more vCenters.
#
# One [ProactivaCollector] instance accumulates rows for every vCenter processed; call
# SetCurrentVCenter() before each one, run the Collect* methods, then GetReport() returns
# one property per raw Excel sheet.
#
# The sheet names and column names returned by GetReport() are a contract: the raw Excel
# and the annex checks (src/report/checks) read them by name. Do not rename them.
#
# Every vCenter call here is read-only, EXCEPT CollectAlarmTest(), which creates a temporary
# alarm on a datastore to trigger the client's alarm script (the "Falso Positivo" test) and
# deletes it right after. Pass -SkipAlarmTest to Start-Proactiva.ps1 to skip that part.

. (Join-Path $PSScriptRoot 'HardwareCompatibility.ps1')

function Show-Progress($total, $index) {
    $current = [math]::Round(($index / $total) * 100)
    $whitespace = 3 - $current.ToString().Length
    Write-Host (" " * $whitespace + $current + "%") -NoNewline

    if ($current -lt 100) {
        Write-Host "`b`b`b`b" -NoNewline
    } else {
        Write-Host ""
    }
}

class ProactivaCollector {

    [string] $ReferenceDataFolder
    [string] $RunFolder

    [PSCustomObject[]] $toolsRefJson
    [PSCustomObject] $vCenterSizing
    [PSCustomObject] $vsanHcl = $null
    [PSCustomObject] $ioHclRef = $null
    [hashtable] $ioHclIndex = $null
    [hashtable] $vsanHclIndex = $null

    # Where each host of the current vCenter lives, by host Id ("HostSystem-host-12"), built
    # once per vCenter by SetHostLocations() instead of asking Get-Cluster/Get-Datacenter per object
    [hashtable] $clusterByHostId = @{}
    [hashtable] $datacenterByHostId = @{}
    $currentClusters = $null

    # Every datastore of the current vCenter (Get-Datastore, read once by SetDatastores)
    $currentDatastores = $null

    # Every datastore name already listed in the Datastores sheet (vSAN ones are listed once)
    [System.Collections.Generic.HashSet[string]] $listedDatastores = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

    # Annotations that identify the vCenter Server appliance VM
    [String[]] $vcsaAnnotations = @("VMware vCenter Server Appliance")

    [PSCustomObject] $hostStandardSwitches = [PSCustomObject] @{}
    [PSCustomObject] $toolsReference = [PSCustomObject] @{}

    [String] $currentVCenter = "No vCenter"
    $currentConnection = $null
    $currentCisConnection = $null

    $esxiReport = [System.Collections.ArrayList]@()
    $nicReport = [System.Collections.ArrayList]@()
    $vmReport = [System.Collections.ArrayList]@()
    $datastoreReport = [System.Collections.ArrayList]@()
    $switchReport = [System.Collections.ArrayList]@()
    $kernelAdaptersReport = [System.Collections.ArrayList]@()
    $snapshotReport = [System.Collections.ArrayList]@()
    $partitionReport = [System.Collections.ArrayList]@()
    $sizingReport = [System.Collections.ArrayList]@()
    $vdsReport = [System.Collections.ArrayList]@()
    $vNetworkReport = [System.Collections.ArrayList]@()
    $licenseReport = [System.Collections.ArrayList]@()
    $vCenterReport = [System.Collections.ArrayList]@()
    $certificateReport = [System.Collections.ArrayList]@()
    $performanceHealthReport = [System.Collections.ArrayList]@()
    $backupActivityReport = [System.Collections.ArrayList]@()
    $alarmCheckReport = [System.Collections.ArrayList]@()

    ProactivaCollector([string] $referenceDataFolder, [string] $runFolder) {
        $this.ReferenceDataFolder = $referenceDataFolder
        $this.RunFolder = $runFolder
        $this.toolsRefJson = Get-Content -Path (Join-Path $referenceDataFolder "vmtools-ref.json") -Raw | ConvertFrom-Json
        $this.vCenterSizing = Get-Content -Path (Join-Path $referenceDataFolder "vcenter-sizing.json") -Raw | ConvertFrom-Json
        # vsan-hcl.json (~21 MB) and the IO HCL are only needed for the drivers: loaded on first use
    }

    SetCurrentVCenter($connection, $cisConnection) {
        $this.currentConnection = $connection
        $this.currentCisConnection = $cisConnection
        $this.currentVCenter = $connection.Name
        $this.clusterByHostId = @{}
        $this.datacenterByHostId = @{}
        $this.currentClusters = $null
        $this.currentDatastores = $null
    }

    # Every datastore of the current vCenter, shared by the alarm test and CollectDatastores
    SetDatastores($datastores) {
        $this.currentDatastores = @($datastores)
    }

    # Datastores of the current vCenter: the ones given to SetDatastores, or a fresh query
    [object[]] GetDatastores() {
        if ($null -eq $this.currentDatastores) { $this.currentDatastores = @(Get-Datastore -Server $this.currentVCenter) }
        return $this.currentDatastores
    }

    # Maps every host of the current vCenter to its cluster and datacenter, so the Collect*
    # methods do not need a Get-Cluster / Get-Datacenter call per host, VM or network card
    SetHostLocations($clusters) {
        $this.currentClusters = $clusters
        foreach ($cluster in $clusters) {
            foreach ($hostRef in $cluster.ExtensionData.Host) {
                $this.clusterByHostId["$($hostRef.Type)-$($hostRef.Value)"] = $cluster.Name
            }
        }

        try {
            foreach ($datacenter in (Get-View -ViewType Datacenter -Property Name -Server $this.currentVCenter -ErrorAction Stop)) {
                foreach ($hostView in (Get-View -ViewType HostSystem -Property Name -SearchRoot $datacenter.MoRef -Server $this.currentVCenter -ErrorAction Stop)) {
                    $this.datacenterByHostId["$($hostView.MoRef.Type)-$($hostView.MoRef.Value)"] = $datacenter.Name
                }
            }
        }
        catch {
            # Hosts left out of the map fall back to Get-Datacenter (see GetDatacenterName)
        }
    }

    [object] GetClusterName($vmHost) {
        return $this.clusterByHostId[$vmHost.Id]
    }

    [object] GetDatacenterName($vmHost) {
        if ($this.datacenterByHostId.ContainsKey($vmHost.Id)) { return $this.datacenterByHostId[$vmHost.Id] }
        return ($vmHost | Get-Datacenter).Name
    }

    # Adds a row to the Datastores sheet and remembers its datastore name(s)
    AddDatastoreRow([PSCustomObject] $row) {
        [void]$this.datastoreReport.Add($row)
        foreach ($name in @($row.Datastore)) {
            if ($null -ne $name) { [void]$this.listedDatastores.Add([string]$name) }
        }
    }

    CollectEsxi($hosts) {
        $hclData = Get-HclReferenceData -ReferenceDataFolder $this.ReferenceDataFolder
        Write-Host "`tProcesando ESXi..." -NoNewline

        $serverContext = $this.currentVCenter

        # Advanced settings and certificates of every reachable host in a few bulk queries;
        # the rest of the host configuration comes from the ExtensionData Get-VMHost already read
        $reachableHosts = @($hosts | Where-Object { @("Connected", "Maintenance") -contains $_.ConnectionState })
        $settingsByHostId = $this.GetAdvancedSettings($reachableHosts)
        $certManagerByHostId = $this.GetCertificateManagers($reachableHosts)

        for ($count = 0; $count -lt $hosts.length; $count++) {
            Show-Progress $hosts.length ($count + 1)
            $h = $hosts[$count]
            try {
                $hclResult = Test-HostHcl -VMHost $h -HclData $hclData
            }
            catch {
                $hclResult = [PSCustomObject]@{
                    Model             = "No data available"
                    Supported         = $null
                    SupportedReleases = "No data available"
                    Reference         = "No data available"
                    Note              = "Host cannot be processed"
                }
            }

            if (@("Connected", "Maintenance") -contains $h.ConnectionState) {
                # Name -> value, case-insensitive like the -eq it replaces. Hosts whose bulk query
                # failed are read one by one, as before.
                $advancedSettings = $settingsByHostId[$h.Id]
                if ($null -eq $advancedSettings) {
                    $advancedSettings = New-SettingDictionary ($h | Get-AdvancedSetting)
                }

                $shellTimeout = $advancedSettings['UserVars.ESXIShellTimeOut']
                $interactiveTimeout = $advancedSettings['UserVars.ESXIShellinteractiveTimeOut']
                $syslogLogDir = $advancedSettings['Syslog.global.logDir']
                $syslogLogHost = $advancedSettings['Syslog.global.logHost']

                # Same data Get-VMHostService / Get-VMHostNtpServer read, already in ExtensionData
                $ntpService = $h.ExtensionData.Config.Service.Service | Where-Object { $_.Key -eq "ntpd" }
                $ntpdRunning = if ($ntpService) { $ntpService.Running.ToString() } else { "Not Available" }

                $ntpServers = @($h.ExtensionData.Config.DateTimeInfo.NtpConfig.Server) -join ","

                $certValidTo = $null
                $certStatus = "Unknown"
                $certIssuer = ""

                try {
                    if ($h.ExtensionData.ConfigManager.CertificateManager) {
                        # From the bulk query; one by one only if it failed. -Server makes sure
                        # the certificate manager is read from this vCenter
                        $certManager = $certManagerByHostId[$h.Id]
                        if ($null -eq $certManager) {
                            $certManager = Get-View -Id $h.ExtensionData.ConfigManager.CertificateManager -Server $serverContext -ErrorAction Stop
                        }

                        if ($certManager.CertificateInfo) {
                            $certInfo = $certManager.CertificateInfo

                            $certValidToDate = $certInfo.NotAfter
                            $certValidTo = Get-Date $certValidToDate -Format "yyyy-MM-dd"
                            $certIssuer = $certInfo.Issuer

                            $now = Get-Date
                            if ($now -gt $certValidToDate) {
                                $certStatus = "Expired"
                            } elseif ($now.AddMonths(1) -gt $certValidToDate) {
                                $certStatus = "Expiration imminent"
                            } elseif ($now.AddMonths(2) -gt $certValidToDate) {
                                $certStatus = "Expiration shortly"
                            } elseif ($now.AddMonths(8) -gt $certValidToDate) {
                                $certStatus = "Expiration soon"
                            } else {
                                $certStatus = "Valid"
                            }
                        }
                    }
                }
                catch {
                    # Leave the certificate columns as "Unknown"; the annex reports them as a read error
                }
            }
            else {
                $shellTimeout = $null
                $interactiveTimeout = $null
                $syslogLogDir = $null
                $syslogLogHost = $null
                $ntpdRunning = $null
                $ntpServers = $null

                $certValidTo = $null
                $certStatus = "Unknown"
                $certIssuer = ""
            }

            [void]$this.esxiReport.Add([PSCustomObject] @{
                vCenter                     = $this.currentVCenter
                Hostname                    = $h.Name
                Model                       = $hclResult.Model
                Datacenter                  = $this.GetDatacenterName($h)
                Cluster                     = $this.GetClusterName($h)
                ESXiVersion                 = $h.ExtensionData.Summary.Config.Product.FullName
                ConnectionState             = $h.ConnectionState.ToString()
                MemoryGB                    = [math]::Round($h.MemoryTotalGB, 3)
                CpuModel                    = $h.ProcessorType
                CpuSpeed                    = $h.ExtensionData.summary.hardware.CpuMhz
                DnsServer                   = $h.ExtensionData.Config.Network.DnsConfig.Address -join ","
                NtpServer                   = $ntpServers
                NtpdRunning                 = $ntpdRunning
                "Cert Valid To"             = $certValidTo
                "Cert Status"               = $certStatus
                "Cert Issuer"               = $certIssuer
                ESXIShellTimeOut            = $shellTimeout
                ESXIShellinteractiveTimeOut = $interactiveTimeout
                PowerManagement             = $h.ExtensionData.Hardware.CpuPowerManagementInfo.CurrentPolicy
                SyslogGlobalLogDir          = $syslogLogDir
                SyslogGlobalLogHost         = $syslogLogHost
                # Nutanix hosts: ESXi is managed through Nutanix (AOS/LCM) and what is supported is
                # defined by Nutanix's own compatibility matrix, so the Broadcom HCL result is
                # kept only as reference and the check is left to the client's responsible
                Supported                   = if ($hclResult.IsNutanix) { "Nutanix" } elseif ($null -eq $hclResult.Supported) { "Sin datos" } elseif ($hclResult.Supported) { "True" } else { "False" }
                "Supported Releases"        = $hclResult.SupportedReleases -join ","
                Note                        = if ($hclResult.IsNutanix) { "Host Nutanix: validacion manual contra la Nutanix Compatibility & Interoperability Matrix. Referencia Broadcom: $($hclResult.Note)" } else { $hclResult.Note }
                Reference                   = $hclResult.Reference
            })
        }
    }

    # Advanced settings of the given hosts, by host Id, as name -> value dictionaries. One
    # Get-View per batch of hosts instead of one Get-AdvancedSetting per host; hosts of a batch
    # that fails are left out (the caller reads them one by one).
    [hashtable] GetAdvancedSettings($hosts) {
        $settingsByHostId = @{}
        $hostIdByManager = @{}
        foreach ($h in $hosts) {
            $managerRef = $h.ExtensionData.ConfigManager.AdvancedOption
            if ($managerRef) { $hostIdByManager["$($managerRef.Type)-$($managerRef.Value)"] = @{ HostId = $h.Id; Ref = $managerRef } }
        }

        $managers = @($hostIdByManager.Values)
        $batchSize = 20
        for ($start = 0; $start -lt $managers.Count; $start += $batchSize) {
            $batch = @($managers[$start..([math]::Min($start + $batchSize, $managers.Count) - 1)])
            try {
                foreach ($view in (Get-View -Id @($batch | ForEach-Object { $_.Ref }) -Property Setting -Server $this.currentVCenter -ErrorAction Stop)) {
                    $entry = $hostIdByManager["$($view.MoRef.Type)-$($view.MoRef.Value)"]
                    if ($entry) { $settingsByHostId[$entry.HostId] = New-SettingDictionary $view.Setting }
                }
            }
            catch {
                # Left out: read with Get-AdvancedSetting
            }
        }
        return $settingsByHostId
    }

    # CertificateManager views (with CertificateInfo) of the given hosts, by host Id, in one
    # Get-View. If it fails the result is empty and the caller asks host by host.
    [hashtable] GetCertificateManagers($hosts) {
        $managerByHostId = @{}
        $hostIdByManager = @{}
        foreach ($h in $hosts) {
            $managerRef = $h.ExtensionData.ConfigManager.CertificateManager
            if ($managerRef) { $hostIdByManager["$($managerRef.Type)-$($managerRef.Value)"] = @{ HostId = $h.Id; Ref = $managerRef } }
        }
        if ($hostIdByManager.Count -eq 0) { return $managerByHostId }

        try {
            foreach ($view in (Get-View -Id @($hostIdByManager.Values | ForEach-Object { $_.Ref }) -Property CertificateInfo -Server $this.currentVCenter -ErrorAction Stop)) {
                $entry = $hostIdByManager["$($view.MoRef.Type)-$($view.MoRef.Value)"]
                if ($entry) { $managerByHostId[$entry.HostId] = $view }
            }
        }
        catch {
            # Empty: read host by host
        }
        return $managerByHostId
    }

    # Network cards and storage controllers of every host, matched against the IO HCL.
    # Devices that are not in the HCL are reported too (URL = "No encontrado en HCL"), so the
    # driver check can flag them.
    CollectHostDrivers($hosts) {
        Write-Host "`tProcesando drivers (ESXi IO)..." -NoNewline
        $vmHosts = $hosts | Where-Object { @("Connected", "Maintenance") -contains $_.ConnectionState }
        $devices = @($vmHosts | Get-VMHostPciDevice | Where-Object { $_.DeviceClass -eq "MassStorageController" -or $_.DeviceClass -eq "NetworkController" -or $_.DeviceClass -eq "SerialBusController" })
        $hclIndex = $this.GetIoHclIndex()

        # esxcli and HBA results are per host; cache them instead of asking again for every device
        $esxcliByHost = @{}
        $nicListByHost = @{}
        $vibListByHost = @{}
        $hbasByHost = @{}

        for ($count = 0; $count -lt $devices.length; $count++) {
            Show-Progress $devices.length ($count + 1)
            $device = $devices[$count]
            if ($device.DeviceName -like "*USB*" -or $device.DeviceName -like "*iLO*" -or $device.DeviceName -like "*iDRAC*") {
                continue
            }

            $vid = [String]::Format("{0:x4}", $device.VendorId)
            $did = [String]::Format("{0:x4}", $device.DeviceId)
            $svid = [String]::Format("{0:x4}", $device.SubVendorId)
            $ssid = [String]::Format("{0:x4}", $device.SubDeviceId)

            $hclEntry = $hclIndex["$vid`:$did`:$svid`:$ssid"]
            $hclUrl = if ($hclEntry) { $hclEntry.url } else { "No encontrado en HCL" }
            $hclReleases = if ($hclEntry) { $hclEntry.releases -join "," } else { "" }

            $hostName = $device.VMHost.Name
            if (-not $esxcliByHost.ContainsKey($hostName)) { $esxcliByHost[$hostName] = $device.VMHost | Get-EsxCli -V2 }
            $esxcli = $esxcliByHost[$hostName]

            if ($device.DeviceClass -eq "NetworkController") {
                if (-not $nicListByHost.ContainsKey($hostName)) { $nicListByHost[$hostName] = $esxcli.network.nic.list.Invoke() }
                $nicData = $nicListByHost[$hostName] | Where-Object { $_.PCIDevice -like '*' + $device.Id }
                if (-not $nicData) { continue }

                $nicDetail = $esxcli.network.nic.get.Invoke(@{ nicname = $nicData.Name })

                [void]$this.nicReport.Add([PSCustomObject] @{
                    "vCenter"                 = $this.currentVCenter
                    "Hostname"                = $device.VMHost.name
                    "Placa"                   = $nicData.Name
                    "Controlador"             = $device.DeviceName
                    "Vendor"                  = $device.VendorName
                    "Driver"                  = $nicDetail.DriverInfo.Driver
                    "Version"                 = $nicDetail.DriverInfo.Version
                    "Firmware"                = $nicDetail.DriverInfo.FirmwareVersion
                    "Vid"                     = $vid; "Did" = $did; "Svid" = $svid; "ssid" = $ssid
                    "ESXi Release"            = $device.VMHost.ExtensionData.Summary.Config.Product.FullName
                    "ESXi Supported Releases" = $hclReleases
                    "URL"                     = $hclUrl
                    "vSAN Compatibility"      = $null
                })
            }
            else {
                if (-not $hbasByHost.ContainsKey($hostName)) { $hbasByHost[$hostName] = @($device.VMHost | Get-VMHostHba -ErrorAction SilentlyContinue) }
                $hbaData = $hbasByHost[$hostName] | Where-Object { $_.PCI -like '*' + $device.Id }
                if (-not $hbaData) { continue }

                $vibName = $hbaData.Driver -replace "_", "-"
                if (-not $vibListByHost.ContainsKey($hostName)) { $vibListByHost[$hostName] = $esxcli.software.vib.list.Invoke() }
                $driverVersion = ($vibListByHost[$hostName] | Where-Object { $_.Name -match "^(scsi-|sata-|)$vibName" }).Version -join ", "
                $vsanCompatibility = $this.GetVsanHclIndex()["$vid`:$did`:$svid`:$ssid"]

                [void]$this.nicReport.Add([PSCustomObject] @{
                    "vCenter"                 = $this.currentVCenter
                    "Hostname"                = $device.VMHost.name
                    "Placa"                   = $hbaData.Device
                    "Controlador"             = $device.DeviceName
                    "Vendor"                  = $device.VendorName
                    "Driver"                  = $hbaData.Driver
                    "Version"                 = $driverVersion
                    "Firmware"                = $null
                    "Vid"                     = $vid; "Did" = $did; "Svid" = $svid; "ssid" = $ssid
                    "ESXi Release"            = $device.VMHost.ExtensionData.Summary.Config.Product.FullName
                    "ESXi Supported Releases" = $hclReleases
                    "URL"                     = $hclUrl
                    "vSAN Compatibility"      = $vsanCompatibility -join ","
                })
            }
        }
    }

    # $vms come from Get-VmInventory: their network cards and ISOs are already part of each VM
    CollectVms($vms, $clusters, $snapshots) {
        $vms = $vms | ForEach-Object { $_ }

        Write-Host "`tProcesando placas de red de VMs..." -NoNewline
        $allNetworkAdapters = @($vms | ForEach-Object { $_.NetworkAdapters })
        $this.CollectVmNetwork($allNetworkAdapters)

        # Snapshots grouped by VM once, instead of a Where-Object over all of them for every VM
        $snapshotsByVm = Group-ById $snapshots { param($snapshot) $snapshot.VM.Id }

        Write-Host "`tProcesando VMs..." -NoNewline
        for ($count = 0; $count -lt $vms.Count; $count++) {
            $vm = $vms[$count]
            Show-Progress $vms.Count ($count + 1)
            $networkAdapters = $vm.NetworkAdapters
            $nicCount = $networkAdapters.Count
            $toolsRequiredVersion = $this.GetToolsReference($vm.VMHost)
            [void]$this.vmReport.Add([PSCustomObject]@{
                vCenter              = $this.currentVCenter
                VM                   = $vm.Name
                Cluster              = $this.GetClusterName($vm.VMHost)
                Host                 = $vm.VMHost.name
                ConnectionState      = $vm.ExtensionData.Runtime.ConnectionState.ToString()
                State                = $vm.PowerState.ToString()
                vCPU                 = $vm.NumCpu
                "Memory MB"          = $vm.MemoryMB
                HardwareVersion      = $vm.HardwareVersion
                Snapshots            = if ($snapshotsByVm[$vm.Id]) { $snapshotsByVm[$vm.Id].Count } else { 0 }
                ToolsStatus          = try { $vm.ExtensionData.Guest.ToolsStatus.ToString() } Catch { "VM has not been scanned" };
                ToolsVersion         = $vm.ExtensionData.Guest.ToolsVersion
                ToolsRequiredVersion = $toolsRequiredVersion
                "SO (vCenter)"       = $vm.ExtensionData.Config.GuestFullName
                "SO (Tools)"         = $vm.ExtensionData.Guest.GuestFullName
                IsoConnected         = Get-SingleOrArray $vm.IsoPaths
                Adapter_01           = if ($nicCount -gt 0) { $networkAdapters[0].Type.ToString() } else { "" }
                Adapter_02           = if ($nicCount -gt 1) { $networkAdapters[1].Type.ToString() } else { "" }
                Adapter_03           = if ($nicCount -gt 2) { $networkAdapters[2].Type.ToString() } else { "" }
                Adapter_04           = if ($nicCount -gt 3) { $networkAdapters[3].Type.ToString() } else { "" }
                Adapter_05           = if ($nicCount -gt 4) { $networkAdapters[4].Type.ToString() } else { "" }
                Adapter_06           = if ($nicCount -gt 5) { $networkAdapters[5].Type.ToString() } else { "" }
                Adapter_07           = if ($nicCount -gt 6) { $networkAdapters[6].Type.ToString() } else { "" }
                Adapter_08           = if ($nicCount -gt 7) { $networkAdapters[7].Type.ToString() } else { "" }
                Adapter_09           = if ($nicCount -gt 8) { $networkAdapters[8].Type.ToString() } else { "" }
                Adapter_10           = if ($nicCount -gt 9) { $networkAdapters[9].Type.ToString() } else { "" }
            })
        }
    }

    CollectVmNetwork($allNetworkAdapters) {
        for ($count = 0; $count -lt $allNetworkAdapters.length; $count++) {
            Show-Progress $allNetworkAdapters.length ($count + 1)
            $networkAdapter = $allNetworkAdapters[$count]
            [void]$this.vNetworkReport.Add([PSCustomObject]@{
                vCenter         = $this.currentVCenter
                VM              = $networkAdapter.Parent.Name
                Cluster         = $this.GetClusterName($networkAdapter.Parent.VMHost)
                Host            = $networkAdapter.Parent.VMHost.Name
                Status          = $networkAdapter.Parent.PowerState
                Mac             = $networkAdapter.MacAddress
                Connected       = if ($networkAdapter.ConnectionState.Connected) { "True" } else { "False" }
                StartsConnected = if ($networkAdapter.ConnectionState.StartConnected) { "True" } else { "False" }
            })
        }
    }

    CollectLicenses() {
        Write-Host "`tProcesando licencias..."
        foreach ($licenseManager in (Get-View LicenseManager -Server $this.currentVCenter)) {
            foreach ($license in $licenseManager.Licenses) {
                $licenseProperties = $license.Properties
                $expiryInfo = $licenseProperties | Where-Object { $_.Key -eq 'expirationDate' } | Select-Object -ExpandProperty Value
                if ($license.Name -eq 'Product Evaluation') {
                    $expirationDate = 'Evaluation'
                }
                elseif ($null -eq $expiryInfo) {
                    $expirationDate = 'Never'
                }
                else {
                    $expirationDate = $expiryInfo
                }

                $totalLicenses = if ($license.Total -eq 0) { 'Unlimited' } else { $license.Total }

                $productName = $licenseProperties | Where-Object { $_.Key -eq 'ProductName' } | Select-Object -ExpandProperty Value
                $productVersion = $licenseProperties | Where-Object { $_.Key -eq 'ProductVersion' } | Select-Object -ExpandProperty Value

                [void]$this.licenseReport.Add([PSCustomObject]@{
                    Name           = $license.Name
                    LicenseKey     = $license.LicenseKey
                    ExpirationDate = $expirationDate
                    ProductName    = if ($null -eq ($productName)) { "No product name" } else { $productName }
                    ProductVersion = if ($null -eq ($productName)) { "No product version" } else { $productVersion }
                    EditionKey     = $license.EditionKey
                    Total          = $totalLicenses
                    Used           = $license.Used
                    CostUnit       = $license.CostUnit
                    vCenter        = $this.currentVCenter
                })
            }
        }
    }

    CollectDatastores($hosts) {
        Write-Host "`tProcesando datastores..." -NoNewline

        # Every datastore once, then each host's own from its ExtensionData.Datastore list (the
        # list "$h | Get-Datastore" reads); its LUNs come from ExtensionData too
        $datastoresById = @{}
        foreach ($datastore in $this.GetDatastores()) { $datastoresById[$datastore.Id] = $datastore }

        for ($count = 0; $count -lt $hosts.length; $count++) {
            Show-Progress $hosts.length ($count + 1)
            $h = $hosts[$count]
            $esxName = $h.Name

            $allDatastores = @($h.ExtensionData.Datastore | ForEach-Object { $datastoresById["$($_.Type)-$($_.Value)"] } | Where-Object { $_ })
            $vmfsDatastores = $allDatastores | Where-Object { $_.Type -ne "vsan" }
            $storageDevice = $h.ExtensionData.Config.StorageDevice

            foreach ($lun in $storageDevice.MultipathInfo.Lun) {
                $scsiLun = $storageDevice.ScsiLun | Where-Object { $_.Key -eq $lun.Lun }
                $canonicalName = $scsiLun.CanonicalName
                $datastore = ($vmfsDatastores | Where-Object { ($_.extensiondata.info.vmfs.extent | Select-Object -expand diskname) -like $canonicalName }).name

                if ($null -ne $datastore) {
                    $policy = if ($lun.Policy.Policy -match "_FIXED") { "Fixed" }
                    elseif ($lun.Policy.Policy -match "_MRU") { "MostRecentlyUsed" }
                    elseif ($lun.Policy.Policy -match "_RR") { "RoundRobin" }
                    else { "Unknown" }

                    $this.AddDatastoreRow([PSCustomObject] @{
                        vCenter   = $this.currentVCenter
                        Hostname  = $esxName
                        Datastore = $datastore
                        Policy    = $policy
                    })
                }
            }

            # vSAN datastores have no LUNs nor multipath policies; list them once so the sheet is complete
            $vsanDatastores = $allDatastores | Where-Object { $_.Type -eq "vsan" }
            foreach ($vsanDatastore in $vsanDatastores) {
                if (-not $this.listedDatastores.Contains($vsanDatastore.Name)) {
                    $this.AddDatastoreRow([PSCustomObject] @{
                        vCenter   = $this.currentVCenter
                        Hostname  = $esxName
                        Datastore = $vsanDatastore.Name
                        Policy    = "vSAN"
                    })
                }
            }
        }
    }

    CollectStandardSwitches($hosts) {
        Write-Host "`tProcesando standard switches..." -NoNewline
        for ($count = 0; $count -lt $hosts.length; $count++) {
            Show-Progress $hosts.length ($count + 1)
            $h = $hosts[$count]
            $cluster = $this.GetClusterName($h)
            # Same data Get-VirtualSwitch -Standard / Get-VirtualPortGroup read, already in ExtensionData
            $network = $h.ExtensionData.Config.Network
            foreach ($standardSwitch in $network.Vswitch) {
                foreach ($portGroup in ($network.Portgroup | Where-Object { $_.Spec.VswitchName -eq $standardSwitch.Name })) {
                    if ($null -ne $portGroup.Spec.VlanId) {
                        [void]$this.switchReport.Add([PSCustomObject] @{
                            vCenter   = $this.currentVCenter
                            ESXi      = $h.name
                            Cluster   = $cluster
                            PortGroup = $portGroup.Spec.Name
                            Switch    = $standardSwitch.Name
                            vLAN      = $portGroup.Spec.VlanId
                        })
                    }
                }
            }
        }
    }

    CollectSnapshots($snapshots) {
        $snapshots = $snapshots | ForEach-Object { $_ }

        Write-Host "`tProcesando snapshots..."
        for ($count = 0; $count -lt $snapshots.Count; $count++) {
            Show-Progress $snapshots.Count ($count + 1)
            $snapshot = $snapshots[$count]
            [void]$this.snapshotReport.Add([PSCustomObject] @{
                vCenter  = $this.currentVCenter
                VM       = $snapshot.VM.Name
                Snapshot = $snapshot.Name
                Fecha    = ($snapshot.Created | Get-Date -Format "yyyy-MM-dd HH:mm")
                SizeMB   = [int]$snapshot.SizeMB
            })
        }
    }

    # Disk usage of the vCenter appliance partitions (only VMs annotated as a VCSA)
    CollectPartitions($vms) {
        $vms = $vms | ForEach-Object { $_ }

        Write-Host "`tProcesando particiones..." -NoNewline
        for ($count = 0; $count -lt $vms.Count; $count++) {
            $vm = $vms[$count];
            Show-Progress $vms.Count ($count + 1)
            if ($vm.ExtensionData.Config.Annotation -in $this.vcsaAnnotations) {
                foreach ($partition in $vm.ExtensionData.Guest.Disk) {
                    [void]$this.partitionReport.Add([PSCustomObject]@{
                        vCenter    = $this.currentVCenter
                        VM         = $vm.name
                        Annotation = $vm.ExtensionData.Config.Annotation
                        Disk       = $partition.DiskPath
                        "Free %"   = [math]::Round(($partition.FreeSpace / $partition.Capacity) * 100, 2)
                    })
                }
            }
        }
    }

    CollectKernelAdapters($hosts) {
        Write-Host "`tProcesando adaptadores VMkernel..." -NoNewline
        $kernelAdapters = @(Get-KernelAdapters -Hosts $hosts -Server $this.currentVCenter | Where-Object { $_.Name -match "vmk[0-9]+" })
        for ($count = 0; $count -lt $kernelAdapters.length; $count++) {
            Show-Progress $kernelAdapters.length ($count + 1)
            $kernelAdapter = $kernelAdapters[$count]
            [void]$this.kernelAdaptersReport.Add([PSCustomObject] @{
                Host       = $kernelAdapter.VMHost.Name
                Name       = $kernelAdapter.Name
                IP         = $kernelAdapter.IP
                MTU        = $kernelAdapter.MTU
                PortGroup  = $kernelAdapter.PortGroupName
                Management = $kernelAdapter.ManagementTrafficEnabled
                vMotion    = $kernelAdapter.VMotionEnabled
            })
        }
    }

    CollectVCenterSizing($vms, $hosts) {
        Write-Host "`tProcesando sizing del vCenter..." -NoNewline

        $vcenterName = $this.currentVCenter
        if ([string]::IsNullOrEmpty($vcenterName)) {
            Write-Warning " -> Error: no se pudo determinar el nombre del vCenter."
            return
        }

        # 1. Find the vCenter VM: by name (FQDN or short name), falling back to the VCSA annotation
        $vcenterShortName = ($vcenterName).Split('.')[0]

        $targetVm = $vms | Where-Object {
            $_.Name -eq $vcenterName -or
            $_.Name -eq $vcenterShortName
        } | Select-Object -First 1

        if (-not $targetVm) {
            $targetVm = $vms | Where-Object { $_.ExtensionData.Config.Annotation -in $this.vcsaAnnotations } | Select-Object -First 1
        }

        if (-not $targetVm) {
            [void]$this.sizingReport.Add([PSCustomObject]@{
                vCenter             = $vcenterName
                VM                  = "No encontrada"
                Annotation          = "-"
                vCPU                = "-"
                "Memory GB"         = "-"
                "Cantidad de VMs"   = $vms.Count
                "Cantidad de Hosts" = $hosts.Count
                "Sizing actual"     = "Unknown (VM not found)"
            })
            return
        }

        # 2. Match the sizing tier for this vCenter version (RAM requirements change per
        # version, see reference-data/vcenter-sizing.json)
        $majorVersion = "8" # Sensible fallback when the version cannot be determined
        if ($this.currentConnection -and $this.currentConnection.Version) {
            $majorVersion = ($this.currentConnection.Version -split '\.')[0]
        }

        $versionProfiles = $this.vCenterSizing.$majorVersion
        if (-not $versionProfiles) {
            # Version without a profile (very old or very new): use the newest one available
            $versionProfiles = $this.vCenterSizing.'8'
        }

        $annotation = if ($targetVm.ExtensionData.Config.Annotation) { $targetVm.ExtensionData.Config.Annotation } else { "-" }
        $currentSizing = "Custom / Undefined"

        # The largest tier whose requirements the VM meets (greater than or equal)
        foreach ($tier in @('x-large', 'large', 'medium', 'small', 'tiny')) {
            $tierRequirements = $versionProfiles.$tier
            if (-not $tierRequirements) { continue }

            if (($targetVm.NumCpu -ge $tierRequirements.vcpus) -and ($targetVm.MemoryGB -ge $tierRequirements.ramGB)) {
                $currentSizing = $tierRequirements.label
                break
            }
        }

        [void]$this.sizingReport.Add([PSCustomObject]@{
            vCenter             = $vcenterName
            VM                  = $targetVm.Name
            Annotation          = $annotation
            vCPU                = $targetVm.NumCpu
            "Memory GB"         = [math]::Round($targetVm.MemoryGB, 0)
            "Cantidad de VMs"   = $vms.Count
            "Cantidad de Hosts" = $hosts.Count
            "Sizing actual"     = $currentSizing
        })

        Write-Host " -> OK." -ForegroundColor Green
    }

    # vDS details plus a configuration export (.zip) of each one into <run folder>\vds_configuration.
    # Export-VDSwitch only reads the configuration; it does not change the switch.
    CollectDistributedSwitches($vdSwitches) {
        Write-Host "`tProcesando vDS (y backup de configuracion)..." -NoNewline

        $backupFolder = Join-Path -Path $this.RunFolder -ChildPath "vds_configuration"
        if ($vdSwitches -and -not (Test-Path $backupFolder)) {
            try {
                New-Item -ItemType Directory -Path $backupFolder -Force | Out-Null
            }
            catch {
                Write-Warning "No se pudo crear la carpeta de backup en '$backupFolder'. Los respaldos van a fallar."
            }
        }

        foreach ($vds in $vdSwitches) {
            # File name: <vCenter>_<switch>_<date>.zip
            $vcenterSafeName = $this.currentVCenter -replace '[^a-zA-Z0-9]', '_'
            $timestamp = (Get-Date).ToString('yyyy-MM-dd_HHmmss')
            $vdsSafeName = $vds.Name -replace '[\\/:*?"<>|]', '_'   # characters Windows does not allow in file names
            $fileName = "${vcenterSafeName}_${vdsSafeName}_$timestamp.zip"

            try {
                Export-VDSwitch -VDSwitch $vds -Destination (Join-Path -Path $backupFolder -ChildPath $fileName) -ErrorAction Stop
            }
            catch {
                Write-Warning "Error al exportar el backup del vDS '$($vds.Name)': $($_.Exception.Message)"
                $fileName = "ERROR: $($_.Exception.Message)"
            }

            [void]$this.vdsReport.Add([PSCustomObject]@{
                vCenter        = $this.currentVCenter
                Name           = $vds.Name
                MTU            = $vds.Mtu
                "NIOC Enabled" = if ($vds.ExtensionData.Config.NetworkResourceManagementEnabled) { "True" } else { "False" }
                "Backup File"  = $fileName
            })
        }
        Write-Host " -> OK." -ForegroundColor Green
    }

    # "Falso Positivo" test: creates a temporary alarm on a datastore that runs the same script
    # as an existing alarm, waits for it to fire, and deletes it. This is the only write
    # operation of the whole collection; $skipAlarmTest records the row as skipped instead.
    CollectAlarmTest([bool] $skipAlarmTest) {
        Write-Host "`tProcesando prueba de alarmas (Falso Positivo)..." -NoNewline

        # Every alarm definition, read once for both the test and the audit
        $allAlarms = @(Get-AlarmDefinition -Server $this.currentVCenter -ErrorAction SilentlyContinue)

        [void]$this.alarmCheckReport.Add($this.GetAlarmTestRow($skipAlarmTest, $allAlarms))
        $this.AuditAlarmConfiguration($this.currentVCenter, $this.currentVCenter, $allAlarms)

        Write-Host " -> OK." -ForegroundColor Green
    }

    [PSCustomObject] GetAlarmTestRow([bool] $skipAlarmTest, $allAlarms) {
        $serverContext = $this.currentVCenter
        $vcenterName = $this.currentVCenter

        $alarmName = "Falso Positivo DS $((Get-Date).ToString("yyyyMMdd_HHmmss"))"
        $sourceAlarmName = "N/A"
        $scriptPath = $null

        # 1. Target: the first accessible datastore
        $targetDatastore = $this.GetDatastores() | Where-Object { $_.State -eq "Available" } | Select-Object -First 1

        if (-not $targetDatastore) {
            Write-Warning " -> No hay datastores accesibles."
            return [PSCustomObject]@{
                vCenter = $vcenterName; Host = "N/A"; "Alarm Name" = "N/A"; "Script Path" = "N/A"; Result = "No hay datastores accesibles"; Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            }
        }
        $targetName = $targetDatastore.Name

        # 2. Source alarm: the first alarm definition that has a script action configured
        foreach ($definition in $allAlarms) {
            try {
                $info = $definition.ExtensionData.Info
                if ($info.Action -and $info.Action.Action) {
                    foreach ($triggerAction in $info.Action.Action) {
                        if ($triggerAction.Action -is 'VMware.Vim.RunScriptAction') {
                            $candidateScript = $triggerAction.Action.Script
                            if (-not [string]::IsNullOrWhiteSpace($candidateScript)) {
                                $scriptPath = $candidateScript
                                $sourceAlarmName = $definition.Name
                                break
                            }
                        }
                    }
                }
            } catch {}

            if ($scriptPath) { break }
        }

        if ([string]::IsNullOrEmpty($scriptPath)) {
            return [PSCustomObject]@{
                vCenter = $vcenterName; Host = $targetName; "Alarm Name" = "N/A"; "Script Path" = "N/A"; Result = "Ninguna alarma en vCenter tiene script configurado"; Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            }
        }

        if ($skipAlarmTest) {
            $testResult = "Omitida (-SkipAlarmTest): no se creo la alarma de prueba"
        }
        else {
            $testResult = $this.RunAlarmTest($serverContext, $targetDatastore, $alarmName, $scriptPath)
        }

        return [PSCustomObject]@{
            vCenter         = $vcenterName
            Host            = $targetName
            "Alarm Name"    = $sourceAlarmName
            "Script Path"   = $scriptPath
            Result          = $testResult
            Timestamp       = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        }
    }

    [string] RunAlarmTest($serverContext, $targetDatastore, [string] $alarmName, [string] $scriptPath) {
        # Metric used by the alarm: disk.used.latest
        try {
            $perfManager = Get-View (Get-View ServiceInstance -Server $serverContext).Content.PerfManager -Server $serverContext
            $counterInfo = $perfManager.PerfCounter | Where-Object { $_.GroupInfo.Key -eq "disk" -and $_.NameInfo.Key -eq "used" -and $_.RollupType -eq "latest" } | Select-Object -First 1
            if (-not $counterInfo) { throw "Metric missing" }
            $metricId = [int]$counterInfo.Key
        }
        catch {
            return "ERROR API Métrica"
        }

        try {
            # A. Remove any leftover alarm with the same name
            $existing = Get-AlarmDefinition -Name $alarmName -Entity $targetDatastore -Server $serverContext -ErrorAction SilentlyContinue
            if ($existing) { Remove-AlarmDefinition $existing -Server $serverContext -Confirm:$false }

            # B. Alarm definition
            $spec = New-Object VMware.Vim.AlarmSpec
            $spec.Name = $alarmName
            $spec.Description = "DevOps Smoke Test"
            $spec.Enabled = $true

            # C. Expression: a metric alarm (always true: 10/20 KB thresholds) wrapped in an OR
            $expression = New-Object VMware.Vim.MetricAlarmExpression
            $expression.Operator = "isAbove"
            $expression.Type = "Datastore"
            $expression.Yellow = 10
            $expression.Red = 20

            $expression.Metric = New-Object VMware.Vim.PerfMetricId
            $expression.Metric.CounterId = $metricId
            $expression.Metric.Instance = ""

            $spec.Expression = New-Object VMware.Vim.OrAlarmExpression
            $spec.Expression.Expression += $expression

            # D. Action: run the client's alarm script on green->yellow and yellow->red
            $actionGroup = New-Object VMware.Vim.GroupAlarmAction
            $actionTrigger = New-Object VMware.Vim.AlarmTriggeringAction

            $scriptAction = New-Object VMware.Vim.RunScriptAction
            $scriptAction.Script = $scriptPath
            $actionTrigger.Action = $scriptAction

            $actionTrigger.Green2yellow = $true
            $actionTrigger.Yellow2red = $true

            $toYellow = New-Object VMware.Vim.AlarmTriggeringActionTransitionSpec
            $toYellow.StartState = 'green'
            $toYellow.FinalState = 'yellow'
            $toYellow.Repeats = $false
            $actionTrigger.TransitionSpecs += $toYellow

            $toRed = New-Object VMware.Vim.AlarmTriggeringActionTransitionSpec
            $toRed.StartState = 'yellow'
            $toRed.FinalState = 'red'
            $toRed.Repeats = $false
            $actionTrigger.TransitionSpecs += $toRed

            $actionGroup.Action += $actionTrigger
            $spec.Action = $actionGroup

            $spec.Setting = New-Object VMware.Vim.AlarmSetting
            $spec.Setting.ReportingFrequency = 0
            $spec.Setting.ToleranceRange = 0
            $spec.ActionFrequency = 0

            # E. Create
            $alarmManager = Get-View AlarmManager -Server $serverContext
            $alarmRef = $alarmManager.CreateAlarm($targetDatastore.ExtensionData.MoRef, $spec)

            # F. Wait and check. Get-VIEvent does not support datastores ("Events can be retrieved
            # only for inventory objects"), so the trigger is confirmed from the datastore's
            # triggered alarm state instead. Polled every 2 s (up to 30 s) instead of a fixed wait;
            # once it fires, a few seconds of margin let vCenter run the script action before
            # the alarm is removed.
            $deadline = (Get-Date).AddSeconds(30)
            $fired = $false
            do {
                Start-Sleep -Seconds 2
                $datastoreView = Get-View $targetDatastore.Id -Property TriggeredAlarmState -Server $serverContext
                $triggeredState = $datastoreView.TriggeredAlarmState | Where-Object { $_.Alarm.Value -eq $alarmRef.Value }
                $fired = $triggeredState -and $triggeredState.OverallStatus -eq "red"
            } until ($fired -or (Get-Date) -gt $deadline)

            if ($fired) { Start-Sleep -Seconds 5 }

            $result = if ($fired) {
                "SUCCESS (Falso positivo ejecutado)"
            } else {
                "WARNING: Alarm created but trigger failed"
            }

            # G. Clean up
            (Get-View $alarmRef -Server $serverContext).RemoveAlarm()
            return $result
        }
        catch {
            Write-Warning "`nError en la prueba de alarma: $($_.Exception.Message)"
            $leftover = Get-AlarmDefinition -Name $alarmName -Entity $targetDatastore -Server $serverContext -ErrorAction SilentlyContinue
            if ($leftover) { Remove-AlarmDefinition $leftover -Server $serverContext -Confirm:$false }
            return "ERROR: $($_.Exception.Message)"
        }
    }

    # Checks that the alarms Wetcom expects exist and have a script action configured (read-only)
    AuditAlarmConfiguration($serverContext, $vcenterName, $allAlarms) {

        # 1. Is vSAN enabled on any cluster?
        $isVsanEnabled = $false
        try {
            $clusters = if ($null -ne $this.currentClusters) { $this.currentClusters } else { Get-Cluster -Server $serverContext -ErrorAction SilentlyContinue }
            foreach ($cluster in $clusters) {
                if ($cluster.VsanEnabled) { $isVsanEnabled = $true; break }
            }
        } catch {}

        # 2. Expected alarms. Names match exactly or by prefix (base name + space): since
        # vCenter 8.0 U3g VMware appends a KB 385107 reference to some alarm names
        # ("... KB 385107" or the full URL). Keep aligned with the configuracion-de-alarmas repo.
        $standardAlarms = @(
            "Host battery status", "Host connection failure", "Host hardware fan status",
            "Host hardware power status", "Host hardware system board status", "Host hardware temperature status",
            "Host hardware voltage", "Network connectivity lost", "Network uplink redundancy degraded",
            "Network uplink redundancy lost", "Unmanaged workload detected on SIOC-enabled datastore",
            "Cannot connect to storage",
            "Insufficient vSphere HA failover resources", "vSphere HA failover in progress"
        )

        # Certificate alarms changed within 8.0 Update 3, so each group has an old and a new
        # variant. Whichever exists in the vCenter is audited; it is "Not Found" only when
        # neither does:
        #   - 'Certificate Status' (up to 8.0 U3f) -> replaced in 8.0 U3g by 10 alarms, one
        #     per VECS store (KB 412316).
        #   - 'ESXi Host Certificate Status' (up to 8.0 U3g) -> renamed in 8.0 U3h (KB 440340).
        $certificateAlarms = @(
            @{
                Old = @("Certificate Status")
                New = @(
                    "Certificate(s) in VECS TRUSTED_ROOTS store has expired",
                    "Certificate(s) in VECS TRUSTED_ROOTS store is about to expire",
                    "Data-encipherment certificate in VECS has expired",
                    "Data-encipherment certificate in VECS is about to expire",
                    "MACHINE_SSL_CERT certificate in VECS has expired",
                    "MACHINE_SSL_CERT certificate in VECS is about to expire",
                    "SMS certificate in VECS has expired",
                    "SMS certificate in VECS is about to expire",
                    "Solution user certificate(s) in VECS has expired",
                    "Solution user certificate(s) in VECS is about to expire"
                )
                MissingLabel = "Certificate Status (ni sus reemplazos VECS de 8.0 U3g+)"
            },
            @{
                Old = @("ESXi Host Certificate Status")
                New = @("ESXi Host Certificate is about to expire")
                MissingLabel = "ESXi Host Certificate Status (ni su reemplazo de 8.0 U3h+)"
            }
        )

        $vsanAlarms = @(
            "Errors occurred on the disk(s) of a vSAN host", "vSAN capacity utilization alarm 'What if the most consumed host fails'",
            "vSAN capacity utilization alarm 'Storage space'", "vSAN cluster alarm 'vSAN daemon liveness'",
            "vSAN network alarm 'Hosts with connectivity issues'", "vSAN network alarm 'Network latency check'",
            "vSAN network alarm 'vSAN cluster partition'", "vSAN online health alarm 'vSAN critical alert regarding a potential data inconsistency'",
            "vSAN physical disk alarm 'Component limit'", "vSAN physical disk alarm 'Congestion'",
            "vSAN physical disk alarm 'Operation'", "vSAN physical disk alarm 'Disk capacity'",
            "vSAN stretched cluster alarm 'Witness host not found'", "vSAN stretched cluster alarm for 'Site latency' health",
            "vSAN performance service alarm 'Verbose mode'", "vSAN performance service alarm 'Network diagnostic mode'"
        )

        # $allAlarms: every definition, read once by CollectAlarmTest(); searched in memory
        $findAlarm = {
            param($alarms, $name)
            $exact = @($alarms | Where-Object { $_.Name -eq $name })
            if ($exact.Count -gt 0) { return $exact[0] }
            $prefix = "$name "
            $alarms | Where-Object { $_.Name.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase) } |
                Select-Object -First 1
        }

        # 3. Build the list to check
        $alarmsToCheck = @($standardAlarms)
        $missingGroups = @()
        foreach ($group in $certificateAlarms) {
            $hasOld = @($group.Old | Where-Object { & $findAlarm $allAlarms $_ }).Count -gt 0
            $hasNew = @($group.New | Where-Object { & $findAlarm $allAlarms $_ }).Count -gt 0

            if ($hasOld) { $alarmsToCheck += $group.Old }
            if ($hasNew) { $alarmsToCheck += $group.New }
            if (-not $hasOld -and -not $hasNew) { $missingGroups += $group.MissingLabel }
        }
        if ($isVsanEnabled) {
            $alarmsToCheck += $vsanAlarms
        }

        # A certificate group with none of its variants is genuinely missing
        foreach ($label in $missingGroups) {
            [void]$this.alarmCheckReport.Add([PSCustomObject]@{
                vCenter      = $vcenterName
                Host         = "-"
                "Alarm Name" = $label
                "Script Path"= "N/A"
                Result       = "Not Found"
                Timestamp    = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            })
        }

        # 4. Check each alarm's actions
        foreach ($alarmName in $alarmsToCheck) {
            $status = "Not Found"
            $path = "N/A"

            $definition = & $findAlarm $allAlarms $alarmName

            if ($definition) {
                $status = "No Script Configured"

                try {
                    $info = $definition.ExtensionData.Info
                    if ($info.Action -and $info.Action.Action) {
                        foreach ($triggerAction in $info.Action.Action) {
                            if ($triggerAction.Action -is 'VMware.Vim.RunScriptAction') {
                                $script = $triggerAction.Action.Script
                                if (-not [string]::IsNullOrWhiteSpace($script)) {
                                    $path = $script
                                    $status = "OK"
                                } else {
                                    $status = "Script Action Empty"
                                }
                                break
                            }
                        }
                    }
                } catch {
                    $status = "Error reading actions"
                }
            }

            [void]$this.alarmCheckReport.Add([PSCustomObject]@{
                vCenter      = $vcenterName
                Host         = "-"
                "Alarm Name" = $alarmName
                "Script Path"= $path
                Result       = $status
                Timestamp    = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            })
        }
    }

    CollectVCenterInfo($allVms) {
        Write-Host "`tProcesando informacion general del vCenter..." -NoNewline

        # 1. General data
        $vcenterName = $this.currentConnection.Name
        $vcenterShortName = ($vcenterName).Split('.')[0]
        $vcenterVm = $allVms | Where-Object { $_.Name -eq $vcenterName -or $_.Name -eq $vcenterShortName } | Select-Object -First 1
        if (-not $vcenterVm) {
            # VM named differently from the vCenter (e.g. "CBVCENTER8"): same fallback as the sizing
            $vcenterVm = $allVms | Where-Object { $_.ExtensionData.Config.Annotation -in $this.vcsaAnnotations -and $_.Name -match [regex]::Escape($vcenterShortName) } | Select-Object -First 1
        }
        $vmName = if ($vcenterVm) { $vcenterVm.Name } else { "No Encontrada (Físico/Otro)" }

        # 2. Root password expiration (appliance API through CIS)
        $expirationDate = "N/A"
        $daysRemaining = "N/A"

        try {
            $accountService = $this.GetCisService("com.vmware.appliance.local_accounts")

            if ($accountService) {
                $rootInfo = $accountService.get("root")

                if ($rootInfo.password_expires_at) {
                    $expiration = Get-Date $rootInfo.password_expires_at
                    $expirationDate = $expiration.ToString("yyyy-MM-dd HH:mm")
                    $daysRemaining = ($expiration - (Get-Date)).Days
                }
                else {
                    # No expiration date means the password never expires
                    $expirationDate = "Nunca"
                }
            }
        }
        catch {
            Write-Warning "Error leyendo la cuenta root del vCenter: $($_.Exception.Message)"
        }

        [void]$this.vCenterReport.Add([PSCustomObject]@{
            "vCenter Server"  = $vcenterName
            "VM Name"         = $vmName
            "Version"         = $this.currentConnection.Version
            "Build"           = $this.currentConnection.Build
            "Root User"       = "root"
            "Expiration Date" = $expirationDate
            "Days Remaining"  = $daysRemaining
        })

        Write-Host " -> OK." -ForegroundColor Green
    }

    CollectCertificates() {
        Write-Host "`tProcesando certificados del vCenter..." -NoNewline

        $now = Get-Date
        $soonLimit = $now.AddDays(30)

        # A. Machine SSL
        try {
            $tlsService = $this.GetCisService("com.vmware.vcenter.certificate_management.vcenter.tls")

            if ($tlsService) {
                $tlsCert = $tlsService.get()

                $validTo = Get-Date $tlsCert.valid_to
                $validFrom = Get-Date $tlsCert.valid_from

                $status = "Valid"
                if ($now -gt $validTo) { $status = "Expirado" }
                elseif ($soonLimit -gt $validTo) { $status = "Expira Pronto" }

                [void]$this.certificateReport.Add([PSCustomObject]@{
                    vCenter      = $this.currentVCenter
                    Ubicacion    = "Machine SSL"
                    Subject      = $tlsCert.subject_dn
                    Status       = $status
                    "Valid From" = $validFrom
                    "Valid Until"= $validTo
                    Emisor       = $tlsCert.issuer_dn
                })
            }
        }
        catch {
            Write-Warning "Error obteniendo Machine SSL: $($_.Exception.Message)"
        }

        # B. VMCA root / STS signing certificates
        try {
            $signingService = $this.GetCisService("com.vmware.vcenter.certificate_management.vcenter.signing_certificate")

            if ($signingService) {
                foreach ($pem in $signingService.get().signing_cert_chains.cert_chain) {
                    try {
                        $cert = ConvertFrom-PemCertificate $pem

                        $certType = if ($cert.Subject -match "CN=ssoserverSign") { "STS" } else { "VMCA_ROOT" }

                        $status = "Valid"
                        if ($now -gt $cert.NotAfter) { $status = "Expirado" }
                        elseif ($soonLimit -gt $cert.NotAfter) { $status = "Expira Pronto" }

                        [void]$this.certificateReport.Add([PSCustomObject]@{
                            vCenter      = $this.currentVCenter
                            Ubicacion    = $certType
                            Subject      = $cert.Subject
                            Status       = $status
                            "Valid From" = $cert.NotBefore
                            "Valid Until"= $cert.NotAfter
                            Emisor       = $cert.Issuer
                        })
                    } catch {}
                }
            }
        }
        catch {
            Write-Warning "Error obteniendo la lista de VMCA/STS: $($_.Exception.Message)"
        }

        # C. Trusted root chains
        try {
            $rootService = $this.GetCisService("com.vmware.vcenter.certificate_management.vcenter.trusted_root_chains")

            if ($rootService) {
                foreach ($chainId in $rootService.list().chain) {
                    $rawChain = $rootService.get($chainId).cert_chain.cert_chain
                    $pemChain = if ($rawChain -is [Array]) { $rawChain -join "`n" } else { $rawChain }

                    $pemBlocks = [regex]::Matches($pemChain, "(?ms)-----BEGIN CERTIFICATE-----.*?-----END CERTIFICATE-----")

                    foreach ($pemBlock in $pemBlocks) {
                        try {
                            $rootCert = ConvertFrom-PemCertificate $pemBlock.Value

                            $status = "Valid"
                            if ($now -gt $rootCert.NotAfter) { $status = "Expirado" }
                            elseif ($soonLimit -gt $rootCert.NotAfter) { $status = "Expira Pronto" }

                            [void]$this.certificateReport.Add([PSCustomObject]@{
                                vCenter      = $this.currentVCenter
                                Ubicacion    = "Trusted Root"
                                Subject      = $rootCert.Subject
                                Status       = $status
                                "Valid From" = $rootCert.NotBefore
                                "Valid Until"= $rootCert.NotAfter
                                Emisor       = $rootCert.Issuer
                            })
                        } catch {}
                    }
                }
            }
        }
        catch {
            Write-Warning "Error obteniendo la lista de Trusted Roots: $($_.Exception.Message)"
        }

        Write-Host " -> OK." -ForegroundColor Green
    }

    # Whether vCenter is storing performance statistics: CPU/memory at cluster level and
    # network/disk on one host of each cluster, over the last 2 hours.
    CollectPerformanceHealth($clusters) {
        Write-Host "`tProcesando salud de estadisticas de performance..." -NoNewline

        $clusterMetrics = @("cpu.usage.average", "mem.usage.average")
        $hostMetrics = @("net.usage.average", "disk.usage.average")
        $start = (Get-Date).AddHours(-2)

        if (-not $clusters) {
            Write-Host " -> Sin clusters para evaluar." -ForegroundColor Yellow
            return
        }

        foreach ($cluster in $clusters) {
            $cpuStatus = "Sin Datos"
            $memStatus = "Sin Datos"
            $netStatus = "Sin Datos"
            $diskStatus = "Sin Datos"
            $dbStatus = "Posible Error DB"

            # 1. CPU / memory (cluster level)
            try {
                $clusterStats = Get-Stat -Entity $cluster -Stat $clusterMetrics -Start $start -MaxSamples 1 -ErrorAction SilentlyContinue

                if ($clusterStats) {
                    if ($clusterStats | Where-Object { $_.MetricId -eq "cpu.usage.average" }) { $cpuStatus = "OK" }
                    if ($clusterStats | Where-Object { $_.MetricId -eq "mem.usage.average" }) { $memStatus = "OK" }

                    # If both have data, the statistics service works
                    if ($cpuStatus -eq "OK" -and $memStatus -eq "OK") {
                        $dbStatus = "OK"
                    }
                }
            }
            catch {
                $dbStatus = "Error de Consulta (Cluster)"
            }

            # 2. Network / disk (host level): one host returning data is enough
            if ($dbStatus -eq "OK") {
                $clusterHosts = Get-View -ViewType HostSystem -Property Name -SearchRoot $cluster.Id -Server $this.currentVCenter

                if ($clusterHosts) {
                    try {
                        $hostStats = Get-Stat -Entity $clusterHosts[0].Name -Stat $hostMetrics -Start $start -MaxSamples 1 -Server $this.currentVCenter -ErrorAction Stop

                        if ($hostStats) {
                            if ($hostStats | Where-Object { $_.MetricId -eq "net.usage.average" }) { $netStatus = "OK" }
                            if ($hostStats | Where-Object { $_.MetricId -eq "disk.usage.average" }) { $diskStatus = "OK" }
                        }
                    }
                    catch {
                        Write-Warning "Error consultando metricas de host para $($cluster.Name). Error: $($_.Exception.Message)"
                    }
                }
            }

            [void]$this.performanceHealthReport.Add([PSCustomObject]@{
                vCenter      = $this.currentVCenter
                Cluster      = $cluster.Name
                "Health DB"  = $dbStatus
                "CPU Stats"  = $cpuStatus
                "Mem Stats"  = $memStatus
                "Net Stats"  = $netStatus
                "Disk Stats" = $diskStatus
            })
        }
        Write-Host " -> OK." -ForegroundColor Green
    }

    # Last 7 VAMI backup jobs (appliance API through CIS)
    CollectBackupActivity() {
        Write-Host "`tProcesando actividad de backup (VAMI)..." -NoNewline

        $vcenterName = $this.currentVCenter

        if (-not $this.currentCisConnection) {
            Write-Warning " -> No hay conexion CIS activa para $vcenterName (error de permisos/conexion previa)."
            [void]$this.backupActivityReport.Add([PSCustomObject]@{
                vCenter = $vcenterName; Type = "-"; Status = "Error de Conexión CIS"; "Data Transfer" = "-"; Location = "-"; StartTime = "-"; EndTime = "-"; Duration = "-"
            })
            return
        }

        try {
            $allJobs = @()
            $usingDetails = $false

            # Attempt 1: the details service (richer data)
            try {
                $detailsService = $this.GetCisService("com.vmware.appliance.recovery.backup.job.details")

                if ($detailsService) {
                    $jobsMap = $detailsService.list($null)

                    if ($jobsMap) {
                        foreach ($entry in $jobsMap.GetEnumerator()) {
                            $job = $entry.Value
                            if ($null -eq $job.id) {
                                $job | Add-Member -MemberType NoteProperty -Name "id" -Value $entry.Key -Force
                            }
                            $allJobs += $job
                        }
                        $usingDetails = $true
                    }
                }
            } catch { }

            # Attempt 2: the simple service (status and dates only)
            if ($allJobs.Count -eq 0) {
                $simpleService = $this.GetCisService("com.vmware.appliance.recovery.backup.job")

                if ($simpleService) {
                    # Sometimes the list comes wrapped in a Value property, sometimes not
                    $rawList = $simpleService.list()
                    $jobIds = if ($rawList.Value) { $rawList.Value } else { $rawList }

                    foreach ($jobId in $jobIds) {
                        try {
                            $job = $simpleService.get($jobId)
                            if (!$job.id) { $job | Add-Member -Name "id" -Value $jobId -Force }
                            $allJobs += $job
                        } catch {}
                    }
                }
            }

            if ($allJobs.Count -eq 0) {
                [void]$this.backupActivityReport.Add([PSCustomObject]@{
                    vCenter          = $vcenterName
                    Type             = "-"
                    Status           = "No Configurado"
                    "Data Transfer"  = "-"
                    Location         = "-"
                    StartTime        = "-"
                    EndTime          = "-"
                    Duration         = "-"
                })
                Write-Host " -> Sin historial (marcado como No Configurado)." -ForegroundColor Yellow
                return
            }

            $backupHistory = $allJobs | Sort-Object start_time -Descending | Select-Object -First 7

            foreach ($job in $backupHistory) {
                $duration = "N/A"
                if ($job.end_time -and $job.start_time) {
                    $duration = "{0:hh\:mm\:ss}" -f (New-TimeSpan -Start $job.start_time -End $job.end_time)
                }

                $location = "N/A"; $sizeGB = "N/A"; $type = "N/A"

                if ($usingDetails) {
                    if ($job.location) { $location = $job.location }
                    if ($job.type) { $type = $job.type.ToString() }
                    if ($job.size) {
                        $sizeGB = [math]::Round($job.size / 1GB, 2).ToString() + " GB"
                    }
                }

                [void]$this.backupActivityReport.Add([PSCustomObject]@{
                    vCenter          = $vcenterName
                    Type             = $type
                    Status           = if ($job.state) { $job.state } else { $job.status }
                    "Data Transfer"  = $sizeGB
                    Location         = $location
                    StartTime        = Get-Date $job.start_time -Format "yyyy-MM-dd HH:mm:ss"
                    EndTime          = Get-Date $job.end_time -Format "yyyy-MM-dd HH:mm:ss"
                    Duration         = $duration
                })
            }

            Write-Host " -> OK ($($backupHistory.Count) trabajos procesados)." -ForegroundColor Green
        }
        catch {
            Write-Warning "`nError en actividad de backup: $($_.Exception.Message)"
            [void]$this.backupActivityReport.Add([PSCustomObject]@{
                vCenter = $vcenterName; Type = "-"; Status = "Error API: $($_.Exception.Message)"; "Data Transfer" = "-"; Location = "-"; StartTime = "-"; EndTime = "-"; Duration = "-"
            })
        }
    }

    # A CIS service of the current vCenter, or $null when there is no CIS connection to it
    [object] GetCisService([string] $name) {
        if (-not $this.currentCisConnection) { return $null }
        return Get-CisService -Name $name -Server $this.currentCisConnection -ErrorAction SilentlyContinue | Select-Object -First 1
    }

    # One property per raw Excel sheet (see the contract note at the top of this file)
    [PSCustomObject] GetReport() {
        return [PSCustomObject] @{
            "vCenter"           = $this.vCenterReport;
            "ESXi"              = $this.esxiReport;
            "ESXi IO"           = $this.nicReport;
            "VM"                = $this.vmReport;
            "Datastores"        = $this.datastoreReport;
            "Standard Switch"   = $this.switchReport;
            "VMkernel Adapters" = $this.kernelAdaptersReport;
            "Snapshot"          = $this.snapshotReport;
            "Partitions"        = $this.partitionReport;
            "Sizing"            = $this.sizingReport;
            "vDS"               = $this.vdsReport;
            "vNetwork"          = $this.vNetworkReport;
            "vLicense"          = $this.licenseReport;
            "Certficate"        = $this.certificateReport;
            "PerformanceHealth" = $this.performanceHealthReport;
            "BackupActivity"    = $this.backupActivityReport;
            "Falso Positivo"    = $this.alarmCheckReport
        }
    }

    # IO HCL entries by "vid:did:svid:ssid"; the first entry of the file wins, like the
    # Where-Object | Select-Object -First 1 it replaces
    [hashtable] GetIoHclIndex() {
        if ($null -eq $this.ioHclIndex) {
            $index = @{}
            foreach ($entry in $this.GetIoHclReference().data.ioDevices) {
                $key = "$($entry.vid):$($entry.did):$($entry.svid):$($entry.ssid)"
                if (-not $index.ContainsKey($key)) { $index[$key] = $entry }
            }
            $this.ioHclIndex = $index
        }
        return $this.ioHclIndex
    }

    # vSAN HCL links of every storage controller by "vid:did:svid:ssid", in file order
    [hashtable] GetVsanHclIndex() {
        if ($null -eq $this.vsanHclIndex) {
            if ($null -eq $this.vsanHcl) {
                $this.vsanHcl = Get-Content -Path (Join-Path $this.ReferenceDataFolder "vsan-hcl.json") -Raw | ConvertFrom-Json
            }
            $index = @{}
            foreach ($entry in $this.vsanHcl.data.controller) {
                $key = "$($entry.vid):$($entry.did):$($entry.svid):$($entry.ssid)"
                if (-not $index.ContainsKey($key)) { $index[$key] = [System.Collections.Generic.List[object]]::new() }
                $index[$key].Add($entry.vcglink)
            }
            $this.vsanHclIndex = $index
        }
        return $this.vsanHclIndex
    }

    [PSCustomObject] GetIoHclReference() {
        if ($null -eq $this.ioHclRef) {
            $this.ioHclRef = Get-Content -Path (Join-Path $this.ReferenceDataFolder "vmware-iohcl.json") -Raw | ConvertFrom-Json
        }
        return $this.ioHclRef
    }

    [String] GetToolsReference($esxi) {
        if ($null -eq ($this.toolsReference | Get-Member -Name $esxi.name)) {
            $toolsBuild = ($this.toolsRefJson | Where-Object { $_.esxiBuild -eq $esxi.ExtensionData.Config.Product.build }).toolsBuild
            $this.toolsReference | Add-Member -NotePropertyName $esxi.name -NotePropertyValue $toolsBuild
        }
        return $this.toolsReference.($esxi.name)
    }
}

# VMs of the given hosts read with one Get-View, instead of Get-VM host by host plus
# Get-NetworkAdapter / Get-CDDrive per VM. Each VM is an object with the properties of Get-VM's
# that the collector uses (Id, Name, VMHost, PowerState, NumCpu, MemoryMB, MemoryGB,
# HardwareVersion and ExtensionData, a partial view with only the properties below), plus:
#   - NetworkAdapters: like Get-NetworkAdapter's (ParentId, Parent, Type, MacAddress, ConnectionState)
#   - IsoPaths: the IsoPath of every CD drive backed by an ISO (what Get-CDDrive reports)
#   - HasSnapshots
# Templates are left out, as Get-VM does. Order: hosts as given, then the VMs of each host.
function Get-VmInventory {
    param(
        [Parameter(Mandatory = $true)] $Server,
        [Parameter(Mandatory = $true)] [object[]] $Hosts
    )

    $properties = @(
        'Name', 'Runtime.Host', 'Runtime.PowerState', 'Runtime.ConnectionState',
        'Config.Template', 'Config.Hardware.NumCPU', 'Config.Hardware.MemoryMB', 'Config.Hardware.Device',
        'Config.Version', 'Config.Annotation', 'Config.GuestFullName',
        'Guest.ToolsStatus', 'Guest.ToolsVersion', 'Guest.GuestFullName', 'Guest.Disk', 'Snapshot'
    )
    # PowerCLI's own enum, so PowerState prints as "PoweredOn" and serializes to JSON as a number
    # exactly like Get-VM's
    $powerStateType = 'VMware.VimAutomation.ViCore.Types.V1.Inventory.PowerState' -as [type]

    $hostsById = @{}
    $vmsByHostId = @{}
    foreach ($h in $Hosts) {
        $hostsById[$h.Id] = $h
        $vmsByHostId[$h.Id] = [System.Collections.Generic.List[object]]::new()
    }

    foreach ($view in (Get-View -ViewType VirtualMachine -Property $properties -Server $Server)) {
        if ($view.Config.Template -or $null -eq $view.Runtime.Host) { continue }
        $hostId = "$($view.Runtime.Host.Type)-$($view.Runtime.Host.Value)"
        if (-not $vmsByHostId.ContainsKey($hostId)) { continue }

        $vmId = "$($view.MoRef.Type)-$($view.MoRef.Value)"
        $powerStateName = [string]$view.Runtime.PowerState
        $powerStateName = $powerStateName.Substring(0, 1).ToUpper() + $powerStateName.Substring(1)
        $memoryMB = [decimal]$view.Config.Hardware.MemoryMB

        $vm = [PSCustomObject]@{
            Id              = $vmId
            Name            = $view.Name
            VMHost          = $hostsById[$hostId]
            PowerState      = if ($powerStateType) { [enum]::Parse($powerStateType, $powerStateName) } else { $powerStateName }
            NumCpu          = $view.Config.Hardware.NumCPU
            MemoryMB        = $memoryMB
            MemoryGB        = $memoryMB / 1024
            HardwareVersion = $view.Config.Version
            ExtensionData   = $view
            NetworkAdapters = [System.Collections.Generic.List[object]]::new()
            IsoPaths        = [System.Collections.Generic.List[object]]::new()
            HasSnapshots    = $null -ne $view.Snapshot
        }

        foreach ($device in $view.Config.Hardware.Device) {
            if ($device -is 'VMware.Vim.VirtualEthernetCard') {
                $vm.NetworkAdapters.Add([PSCustomObject]@{
                    ParentId        = $vmId
                    Parent          = $vm
                    Type            = Get-NetworkAdapterType $device
                    MacAddress      = $device.MacAddress
                    ConnectionState = [PSCustomObject]@{
                        Connected      = $device.Connectable.Connected
                        StartConnected = $device.Connectable.StartConnected
                    }
                })
            }
            elseif ($device -is 'VMware.Vim.VirtualCdrom' -and $device.Backing -is 'VMware.Vim.VirtualCdromIsoBackingInfo') {
                $vm.IsoPaths.Add($device.Backing.FileName)
            }
        }

        $vmsByHostId[$hostId].Add($vm)
    }

    foreach ($h in $Hosts) {
        foreach ($vm in $vmsByHostId[$h.Id]) { $vm }
    }
}

# PowerCLI's name for a network card type (NetworkAdapter.Type), from its vSphere device class
function Get-NetworkAdapterType($device) {
    switch ($device.GetType().Name) {
        'VirtualE1000' { return 'e1000' }
        'VirtualE1000e' { return 'e1000e' }
        'VirtualPCNet32' { return 'Flexible' }
        'VirtualVmxnet' { return 'Vmxnet' }
        'VirtualVmxnet2' { return 'EnhancedVmxnet' }
        'VirtualVmxnet3' { return 'Vmxnet3' }
        'VirtualVmxnet3Vrdma' { return 'Vmxnet3Vrdma' }
        'VirtualSriovEthernetCard' { return 'SriovEthernetCard' }
    }
    return 'Unknown'
}

# Snapshots (Get-Snapshot, which works out SizeMB) of the VMs that have any, grouped VM by VM
# in the order of $Vms. VMs without snapshots are not asked.
function Get-VmSnapshots {
    param(
        [Parameter(Mandatory = $true)] $Server,
        [object[]] $Vms
    )

    $withSnapshots = @($Vms | Where-Object { $_.HasSnapshots })
    if ($withSnapshots.Count -eq 0) { return }

    $snapshotsByVm = @{}
    foreach ($snapshot in (Get-VM -Id @($withSnapshots | ForEach-Object { $_.Id }) -Server $Server | Get-Snapshot)) {
        $vmId = $snapshot.VM.Id
        if (-not $snapshotsByVm.ContainsKey($vmId)) { $snapshotsByVm[$vmId] = [System.Collections.Generic.List[object]]::new() }
        $snapshotsByVm[$vmId].Add($snapshot)
    }

    foreach ($vm in $withSnapshots) {
        foreach ($snapshot in $snapshotsByVm[$vm.Id]) { $snapshot }
    }
}

# VMkernel adapters of the hosts with the fields of Get-VMHostNetworkAdapter that the report
# uses, read from each host's ExtensionData instead of one query per host:
#   - Management / vMotion: whether the adapter is selected for that service in
#     Config.VirtualNicManagerInfo;
#   - PortGroupName: the standard port group, or the name of the distributed one (one Get-View
#     of every distributed port group, only if some adapter is on a vDS).
# Hosts without configuration (disconnected) are skipped, as the cmdlet does.
function Get-KernelAdapters {
    param(
        [Parameter(Mandatory = $true)] [object[]] $Hosts,
        [Parameter(Mandatory = $true)] $Server
    )

    $distributedPortGroups = $null

    foreach ($h in $Hosts) {
        $config = $h.ExtensionData.Config
        if ($null -eq $config -or $null -eq $config.Network) { continue }

        # Adapter device ("vmk0") -> services it is selected for ("management", "vmotion", ...)
        $servicesByDevice = @{}
        foreach ($netConfig in $config.VirtualNicManagerInfo.NetConfig) {
            foreach ($candidate in $netConfig.CandidateVnic) {
                if ($netConfig.SelectedVnic -contains $candidate.Key) {
                    if (-not $servicesByDevice.ContainsKey($candidate.Device)) { $servicesByDevice[$candidate.Device] = [System.Collections.Generic.List[string]]::new() }
                    $servicesByDevice[$candidate.Device].Add($netConfig.NicType)
                }
            }
        }

        foreach ($vnic in $config.Network.Vnic) {
            $portGroupName = $vnic.Portgroup
            if ([string]::IsNullOrEmpty($portGroupName) -and $vnic.Spec.DistributedVirtualPort) {
                if ($null -eq $distributedPortGroups) {
                    $distributedPortGroups = @{}
                    foreach ($portGroup in (Get-View -ViewType DistributedVirtualPortgroup -Property Key, Name -Server $Server)) {
                        $distributedPortGroups[$portGroup.Key] = $portGroup.Name
                    }
                }
                $portGroupName = $distributedPortGroups[$vnic.Spec.DistributedVirtualPort.PortgroupKey]
            }

            $services = $servicesByDevice[$vnic.Device]
            [PSCustomObject]@{
                VMHost                   = $h
                Name                     = $vnic.Device
                IP                       = $vnic.Spec.Ip.IpAddress
                MTU                      = $vnic.Spec.Mtu
                PortGroupName            = $portGroupName
                ManagementTrafficEnabled = [bool]($services -contains 'management')
                VMotionEnabled           = [bool]($services -contains 'vmotion')
            }
        }
    }
}

# Advanced settings as a case-insensitive name -> value table, from Get-AdvancedSetting results
# (Name / Value) or OptionManager.Setting (Key / Value). The first one of each name wins.
function New-SettingDictionary($settings) {
    $dictionary = [hashtable]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($setting in $settings) {
        $name = if ($setting.PSObject.Properties['Name']) { $setting.Name } else { $setting.Key }
        if ($null -ne $name -and -not $dictionary.ContainsKey($name)) { $dictionary[$name] = $setting.Value }
    }
    return $dictionary
}

# Groups objects in lists by the Id that $getId returns for each one (its only parameter),
# keeping their order
function Group-ById($items, [scriptblock] $getId) {
    $groups = @{}
    foreach ($item in $items) {
        $id = & $getId $item
        if (-not $groups.ContainsKey($id)) { $groups[$id] = [System.Collections.Generic.List[object]]::new() }
        $groups[$id].Add($item)
    }
    return $groups
}

# What a property access on a pipeline result gives: $null for none, the item for one, an
# array for several (keeps the JSON exactly as the Where-Object version produced it)
function Get-SingleOrArray($items) {
    $list = @($items)
    if ($list.Count -eq 0) { return $null }
    if ($list.Count -eq 1) { return $list[0] }
    return $list
}

function ConvertFrom-PemCertificate([string] $pem) {
    $base64 = $pem -replace "-----BEGIN CERTIFICATE-----", "" -replace "-----END CERTIFICATE-----", "" -replace "\s", ""
    return New-Object System.Security.Cryptography.X509Certificates.X509Certificate2(, [System.Convert]::FromBase64String($base64))
}
