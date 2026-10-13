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
    [PSCustomObject[]] $vsanHcl
    [PSCustomObject] $ioHclRef = $null

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
        $this.toolsRefJson = Get-Content -Path (Join-Path $referenceDataFolder "vmtools-ref.json") | ConvertFrom-Json
        $this.vCenterSizing = Get-Content -Path (Join-Path $referenceDataFolder "vcenter-sizing.json") | ConvertFrom-Json
        $this.vsanHcl = Get-Content -Path (Join-Path $referenceDataFolder "vsan-hcl.json") | ConvertFrom-Json
    }

    SetCurrentVCenter($connection, $cisConnection) {
        $this.currentConnection = $connection
        $this.currentCisConnection = $cisConnection
        $this.currentVCenter = $connection.Name
    }

    CollectEsxi($hosts) {
        $hclData = Get-HclReferenceData -ReferenceDataFolder $this.ReferenceDataFolder
        Write-Host "`tProcesando ESXi..." -NoNewline

        $serverContext = $this.currentVCenter

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
                $advancedSettings = $h | Get-AdvancedSetting

                $shellTimeoutSetting = $advancedSettings | Where-Object { $_.name -eq 'UserVars.ESXIShellTimeOut' }
                $shellTimeout = if ($shellTimeoutSetting) { $shellTimeoutSetting.Value } else { $null }

                $interactiveTimeoutSetting = $advancedSettings | Where-Object { $_.name -eq 'UserVars.ESXIShellinteractiveTimeOut' }
                $interactiveTimeout = if ($interactiveTimeoutSetting) { $interactiveTimeoutSetting.Value } else { $null }

                $logDirSetting = $advancedSettings | Where-Object { $_.name -eq 'Syslog.global.logDir' }
                $syslogLogDir = if ($logDirSetting) { $logDirSetting.Value } else { $null }

                $logHostSetting = $advancedSettings | Where-Object { $_.name -eq 'Syslog.global.logHost' }
                $syslogLogHost = if ($logHostSetting) { $logHostSetting.Value } else { $null }

                $ntpService = $h | Get-VMHostService | Where-Object { $_.key -eq "ntpd" }
                $ntpdRunning = if ($ntpService) { $ntpService.Running.ToString() } else { "Not Available" }

                $ntpServers = ""
                try {
                    $ntpServers = ($h | Get-VMHostNtpServer -ErrorAction Stop) -join ","
                }
                catch {
                    $ntpServers = ""
                }

                $certValidTo = $null
                $certStatus = "Unknown"
                $certIssuer = ""

                try {
                    # -Server makes sure the certificate manager is read from this vCenter
                    if ($h.ExtensionData.ConfigManager.CertificateManager) {
                        $certManager = Get-View -Id $h.ExtensionData.ConfigManager.CertificateManager -Server $serverContext -ErrorAction Stop

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

            $this.esxiReport += [PSCustomObject] @{
                vCenter                     = $this.currentVCenter
                Hostname                    = $h.Name
                Model                       = $hclResult.Model
                Datacenter                  = ($h | Get-Datacenter).Name
                Cluster                     = ($h | Get-Cluster).Name
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
                Supported                   = if ($null -eq $hclResult.Supported) { "Sin datos" } elseif ($hclResult.Supported) { "True" } else { "False" }
                "Supported Releases"        = $hclResult.SupportedReleases -join ","
                Note                        = $hclResult.Note
                Reference                   = $hclResult.Reference
            }
        }
    }

    # Network cards and storage controllers of every host, matched against the IO HCL.
    # Devices that are not in the HCL are reported too (URL = "No encontrado en HCL"), so the
    # driver check can flag them.
    CollectHostDrivers($hosts) {
        Write-Host "`tProcesando drivers (ESXi IO)..." -NoNewline
        $vmHosts = $hosts | Where-Object { @("Connected", "Maintenance") -contains $_.ConnectionState }
        $devices = @($vmHosts | Get-VMHostPciDevice | Where-Object { $_.DeviceClass -eq "MassStorageController" -or $_.DeviceClass -eq "NetworkController" -or $_.DeviceClass -eq "SerialBusController" })
        $hclDevices = $this.GetIoHclReference().data.ioDevices

        # esxcli results are per host; cache them instead of asking again for every device
        $esxcliByHost = @{}
        $nicListByHost = @{}
        $vibListByHost = @{}

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

            $hclEntry = $hclDevices | Where-Object { $_.vid -eq $vid -and $_.did -eq $did -and $_.svid -eq $svid -and $_.ssid -eq $ssid } | Select-Object -First 1
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

                $this.nicReport += [PSCustomObject] @{
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
                }
            }
            else {
                $hbaData = $device.VMHost | Get-VMHostHba -ErrorAction SilentlyContinue | Where-Object { $_.PCI -like '*' + $device.Id }
                if (-not $hbaData) { continue }

                $vibName = $hbaData.Driver -replace "_", "-"
                if (-not $vibListByHost.ContainsKey($hostName)) { $vibListByHost[$hostName] = $esxcli.software.vib.list.Invoke() }
                $driverVersion = ($vibListByHost[$hostName] | Where-Object { $_.Name -match "^(scsi-|sata-|)$vibName" }).Version -join ", "
                $vsanCompatibility = Get-VsanHclLink $vid $did $svid $ssid $this.vsanHcl

                $this.nicReport += [PSCustomObject] @{
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
                }
            }
        }
    }

    CollectVms($vms, $clusters, $snapshots) {
        $vms = $vms | ForEach-Object { $_ }

        Write-Host "`tProcesando placas de red de VMs..." -NoNewline
        $connectedIsos = $vms | Get-CDDrive | Where-Object { $null -ne $_.IsoPath }
        $allNetworkAdapters = $vms | Get-NetworkAdapter
        $this.CollectVmNetwork($allNetworkAdapters)

        Write-Host "`tProcesando VMs..." -NoNewline
        for ($count = 0; $count -lt $vms.Count; $count++) {
            $vm = $vms[$count]
            Show-Progress $vms.Count ($count + 1)
            $networkAdapters = $allNetworkAdapters | Where-Object { $_.ParentId -eq $vm.Id }
            $nicCount = $networkAdapters.length
            $toolsRequiredVersion = $this.GetToolsReference($vm.VMHost)
            $this.vmReport += [PSCustomObject]@{
                vCenter              = $this.currentVCenter
                VM                   = $vm.Name
                Cluster              = ($clusters | Where-Object { $_.ExtensionData.Host -contains $vm.VMHost.Id }).name
                Host                 = $vm.VMHost.name
                ConnectionState      = $vm.ExtensionData.Runtime.ConnectionState.ToString()
                State                = $vm.PowerState.ToString()
                vCPU                 = $vm.NumCpu
                "Memory MB"          = $vm.MemoryMB
                HardwareVersion      = $vm.HardwareVersion
                Snapshots            = ($snapshots | Where-Object { $_.VM -eq $vm }).length
                ToolsStatus          = try { $vm.ExtensionData.Guest.ToolsStatus.ToString() } Catch { "VM has not been scanned" };
                ToolsVersion         = $vm.ExtensionData.Guest.ToolsVersion
                ToolsRequiredVersion = $toolsRequiredVersion
                "SO (vCenter)"       = $vm.ExtensionData.Config.GuestFullName
                "SO (Tools)"         = $vm.ExtensionData.Guest.GuestFullName
                IsoConnected         = ($connectedIsos | Where-Object { $_.ParentId -eq $vm.Id }).IsoPath
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
            }
        }
    }

    CollectVmNetwork($allNetworkAdapters) {
        for ($count = 0; $count -lt $allNetworkAdapters.length; $count++) {
            Show-Progress $allNetworkAdapters.length ($count + 1)
            $networkAdapter = $allNetworkAdapters[$count]
            $this.vNetworkReport += [PSCustomObject]@{
                vCenter         = $this.currentVCenter
                VM              = $networkAdapter.Parent.Name
                Cluster         = ($networkAdapter.Parent.VMHost | Get-Cluster).Name
                Host            = $networkAdapter.Parent.VMHost.Name
                Status          = $networkAdapter.Parent.PowerState
                Mac             = $networkAdapter.MacAddress
                Connected       = if ($networkAdapter.ConnectionState.Connected) { "True" } else { "False" }
                StartsConnected = if ($networkAdapter.ConnectionState.StartConnected) { "True" } else { "False" }
            }
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

                $this.licenseReport += [PSCustomObject]@{
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
                }
            }
        }
    }

    CollectDatastores($hosts) {
        Write-Host "`tProcesando datastores..." -NoNewline

        for ($count = 0; $count -lt $hosts.length; $count++) {
            Show-Progress $hosts.length ($count + 1)
            $h = $hosts[$count]
            $esxName = $h.Name

            $allDatastores = $h | Get-Datastore
            $vmfsDatastores = $allDatastores | Where-Object { $_.Type -ne "vsan" }
            $hostView = Get-View -Id $h.Id -Property Name, Config.StorageDevice -Server $this.currentVCenter

            foreach ($lun in $hostView.Config.StorageDevice.MultipathInfo.Lun) {
                $scsiLun = $hostView.Config.StorageDevice.ScsiLun | Where-Object { $_.Key -eq $lun.Lun }
                $canonicalName = $scsiLun.CanonicalName
                $datastore = ($vmfsDatastores | Where-Object { ($_.extensiondata.info.vmfs.extent | Select-Object -expand diskname) -like $canonicalName }).name

                if ($null -ne $datastore) {
                    $policy = if ($lun.Policy.Policy -match "_FIXED") { "Fixed" }
                    elseif ($lun.Policy.Policy -match "_MRU") { "MostRecentlyUsed" }
                    elseif ($lun.Policy.Policy -match "_RR") { "RoundRobin" }
                    else { "Unknown" }

                    $this.datastoreReport += [PSCustomObject] @{
                        vCenter   = $this.currentVCenter
                        Hostname  = $esxName
                        Datastore = $datastore
                        Policy    = $policy
                    }
                }
            }

            # vSAN datastores have no LUNs nor multipath policies; list them once so the sheet is complete
            $vsanDatastores = $allDatastores | Where-Object { $_.Type -eq "vsan" }
            foreach ($vsanDatastore in $vsanDatastores) {
                if (-not ($this.datastoreReport | Where-Object { $_.Datastore -eq $vsanDatastore.Name })) {
                    $this.datastoreReport += [PSCustomObject] @{
                        vCenter   = $this.currentVCenter
                        Hostname  = $esxName
                        Datastore = $vsanDatastore.Name
                        Policy    = "vSAN"
                    }
                }
            }
        }
    }

    CollectStandardSwitches($hosts) {
        Write-Host "`tProcesando standard switches..." -NoNewline
        for ($count = 0; $count -lt $hosts.length; $count++) {
            Show-Progress $hosts.length ($count + 1)
            $h = $hosts[$count]
            $cluster = ($h | Get-Cluster).Name
            foreach ($standardSwitch in ($h | Get-VirtualSwitch -Standard)) {
                foreach ($portGroup in ($standardSwitch | Get-VirtualPortGroup)) {
                    if ($null -ne $portGroup.vLanId) {
                        $this.switchReport += [PSCustomObject] @{
                            vCenter   = $this.currentVCenter
                            ESXi      = $h.name
                            Cluster   = $cluster
                            PortGroup = $portGroup.Name
                            Switch    = $standardSwitch.Name
                            vLAN      = $portGroup.vLanId
                        }
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
            $this.snapshotReport += [PSCustomObject] @{
                vCenter  = $this.currentVCenter
                VM       = $snapshot.VM.Name
                Snapshot = $snapshot.Name
                Fecha    = ($snapshot.Created | Get-Date -Format "yyyy-MM-dd HH:mm")
                SizeMB   = [int]$snapshot.SizeMB
            }
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
                    $this.partitionReport += [PSCustomObject]@{
                        vCenter    = $this.currentVCenter
                        VM         = $vm.name
                        Annotation = $vm.ExtensionData.Config.Annotation
                        Disk       = $partition.DiskPath
                        "Free %"   = [math]::Round(($partition.FreeSpace / $partition.Capacity) * 100, 2)
                    }
                }
            }
        }
    }

    CollectKernelAdapters($hosts) {
        Write-Host "`tProcesando adaptadores VMkernel..." -NoNewline
        $kernelAdapters = $hosts | Get-VMHostNetworkAdapter -ErrorAction SilentlyContinue | Where-Object { $_.Name -match "vmk[0-9]+" }
        for ($count = 0; $count -lt $kernelAdapters.length; $count++) {
            Show-Progress $kernelAdapters.length ($count + 1)
            $kernelAdapter = $kernelAdapters[$count]
            $this.kernelAdaptersReport += [PSCustomObject] @{
                Host       = $kernelAdapter.VMHost.Name
                Name       = $kernelAdapter.Name
                IP         = $kernelAdapter.IP
                MTU        = $kernelAdapter.MTU
                PortGroup  = $kernelAdapter.PortGroupName
                Management = $kernelAdapter.ManagementTrafficEnabled
                vMotion    = $kernelAdapter.VMotionEnabled
            }
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
            $this.sizingReport += [PSCustomObject]@{
                vCenter             = $vcenterName
                VM                  = "No encontrada"
                Annotation          = "-"
                vCPU                = "-"
                "Memory GB"         = "-"
                "Cantidad de VMs"   = $vms.Count
                "Cantidad de Hosts" = $hosts.Count
                "Sizing actual"     = "Unknown (VM not found)"
            }
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

        $this.sizingReport += [PSCustomObject]@{
            vCenter             = $vcenterName
            VM                  = $targetVm.Name
            Annotation          = $annotation
            vCPU                = $targetVm.NumCpu
            "Memory GB"         = [math]::Round($targetVm.MemoryGB, 0)
            "Cantidad de VMs"   = $vms.Count
            "Cantidad de Hosts" = $hosts.Count
            "Sizing actual"     = $currentSizing
        }

        Write-Host " -> OK." -ForegroundColor Green
    }

    # vDS details plus a configuration export (.zip) of each one into <run folder>\vds_configuration.
    # Export-VDSwitch only reads the configuration; it does not change the switch.
    CollectDistributedSwitches($vdSwitches) {
        Write-Host "`tProcesando vDS (y backup de configuracion)..." -NoNewline

        $backupFolder = Join-Path -Path $this.RunFolder -ChildPath "vds_configuration"
        if (-not (Test-Path $backupFolder)) {
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
            $fileName = "${vcenterSafeName}_$($vds.Name)_$timestamp.zip"

            try {
                Export-VDSwitch -VDSwitch $vds -Destination (Join-Path -Path $backupFolder -ChildPath $fileName) -ErrorAction Stop
            }
            catch {
                Write-Warning "Error al exportar el backup del vDS '$($vds.Name)': $($_.Exception.Message)"
                $fileName = "ERROR: $($_.Exception.Message)"
            }

            $this.vdsReport += [PSCustomObject]@{
                vCenter        = $this.currentVCenter
                Name           = $vds.Name
                MTU            = $vds.Mtu
                "NIOC Enabled" = if ($vds.ExtensionData.Config.NetworkResourceManagementEnabled) { "True" } else { "False" }
                "Backup File"  = $fileName
            }
        }
        Write-Host " -> OK." -ForegroundColor Green
    }

    # "Falso Positivo" test: creates a temporary alarm on a datastore that runs the same script
    # as an existing alarm, waits for it to fire, and deletes it. This is the only write
    # operation of the whole collection; $skipAlarmTest records the row as skipped instead.
    CollectAlarmTest([bool] $skipAlarmTest) {
        Write-Host "`tProcesando prueba de alarmas (Falso Positivo)..." -NoNewline

        $this.alarmCheckReport += $this.GetAlarmTestRow($skipAlarmTest)
        $this.AuditAlarmConfiguration($this.currentVCenter, $this.currentVCenter)

        Write-Host " -> OK." -ForegroundColor Green
    }

    [PSCustomObject] GetAlarmTestRow([bool] $skipAlarmTest) {
        $serverContext = $this.currentVCenter
        $vcenterName = $this.currentVCenter

        $alarmName = "Falso Positivo DS $((Get-Date).ToString("yyyyMMdd_HHmmss"))"
        $sourceAlarmName = "N/A"
        $scriptPath = $null

        # 1. Target: the first accessible datastore
        $targetDatastore = Get-Datastore -Server $serverContext | Where-Object { $_.State -eq "Available" } | Select-Object -First 1

        if (-not $targetDatastore) {
            Write-Warning " -> No hay datastores accesibles."
            return [PSCustomObject]@{
                vCenter = $vcenterName; Host = "N/A"; "Alarm Name" = "N/A"; "Script Path" = "N/A"; Result = "No hay datastores accesibles"; Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            }
        }
        $targetName = $targetDatastore.Name

        # 2. Source alarm: the first alarm definition that has a script action configured
        $allAlarms = Get-AlarmDefinition -Server $serverContext -ErrorAction SilentlyContinue

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
            # triggered alarm state instead.
            Start-Sleep -Seconds 15

            $datastoreView = Get-View $targetDatastore.Id -Property TriggeredAlarmState -Server $serverContext
            $triggeredState = $datastoreView.TriggeredAlarmState | Where-Object { $_.Alarm.Value -eq $alarmRef.Value }
            $result = if ($triggeredState -and $triggeredState.OverallStatus -eq "red") {
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
    AuditAlarmConfiguration($serverContext, $vcenterName) {

        # 1. Is vSAN enabled on any cluster?
        $isVsanEnabled = $false
        try {
            $clusters = Get-Cluster -Server $serverContext -ErrorAction SilentlyContinue
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

        # Read every definition once and search in memory
        $allAlarms = @(Get-AlarmDefinition -Server $serverContext -ErrorAction SilentlyContinue)
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
            $this.alarmCheckReport += [PSCustomObject]@{
                vCenter      = $vcenterName
                Host         = "-"
                "Alarm Name" = $label
                "Script Path"= "N/A"
                Result       = "Not Found"
                Timestamp    = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            }
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

            $this.alarmCheckReport += [PSCustomObject]@{
                vCenter      = $vcenterName
                Host         = "-"
                "Alarm Name" = $alarmName
                "Script Path"= $path
                Result       = $status
                Timestamp    = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            }
        }
    }

    CollectVCenterInfo($allVms) {
        Write-Host "`tProcesando informacion general del vCenter..." -NoNewline

        # 1. General data
        $vcenterName = $this.currentConnection.Name
        $vcenterShortName = ($vcenterName).Split('.')[0]
        $vcenterVm = $allVms | Where-Object { $_.Name -eq $vcenterName -or $_.Name -eq $vcenterShortName } | Select-Object -First 1
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

        $this.vCenterReport += [PSCustomObject]@{
            "vCenter Server"  = $vcenterName
            "VM Name"         = $vmName
            "Version"         = $this.currentConnection.Version
            "Build"           = $this.currentConnection.Build
            "Root User"       = "root"
            "Expiration Date" = $expirationDate
            "Days Remaining"  = $daysRemaining
        }

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

                $this.certificateReport += [PSCustomObject]@{
                    vCenter      = $this.currentVCenter
                    Ubicacion    = "Machine SSL"
                    Subject      = $tlsCert.subject_dn
                    Status       = $status
                    "Valid From" = $validFrom
                    "Valid Until"= $validTo
                    Emisor       = $tlsCert.issuer_dn
                }
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

                        $this.certificateReport += [PSCustomObject]@{
                            vCenter      = $this.currentVCenter
                            Ubicacion    = $certType
                            Subject      = $cert.Subject
                            Status       = $status
                            "Valid From" = $cert.NotBefore
                            "Valid Until"= $cert.NotAfter
                            Emisor       = $cert.Issuer
                        }
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

                            $this.certificateReport += [PSCustomObject]@{
                                vCenter      = $this.currentVCenter
                                Ubicacion    = "Trusted Root"
                                Subject      = $rootCert.Subject
                                Status       = $status
                                "Valid From" = $rootCert.NotBefore
                                "Valid Until"= $rootCert.NotAfter
                                Emisor       = $rootCert.Issuer
                            }
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

            $this.performanceHealthReport += [PSCustomObject]@{
                vCenter      = $this.currentVCenter
                Cluster      = $cluster.Name
                "Health DB"  = $dbStatus
                "CPU Stats"  = $cpuStatus
                "Mem Stats"  = $memStatus
                "Net Stats"  = $netStatus
                "Disk Stats" = $diskStatus
            }
        }
        Write-Host " -> OK." -ForegroundColor Green
    }

    # Last 7 VAMI backup jobs (appliance API through CIS)
    CollectBackupActivity() {
        Write-Host "`tProcesando actividad de backup (VAMI)..." -NoNewline

        $vcenterName = $this.currentVCenter

        if (-not $this.currentCisConnection) {
            Write-Warning " -> No hay conexion CIS activa para $vcenterName (error de permisos/conexion previa)."
            $this.backupActivityReport += [PSCustomObject]@{
                vCenter = $vcenterName; Type = "-"; Status = "Error de Conexión CIS"; "Data Transfer" = "-"; Location = "-"; StartTime = "-"; EndTime = "-"; Duration = "-"
            }
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
                $this.backupActivityReport += [PSCustomObject]@{
                    vCenter          = $vcenterName
                    Type             = "-"
                    Status           = "No Configurado"
                    "Data Transfer"  = "-"
                    Location         = "-"
                    StartTime        = "-"
                    EndTime          = "-"
                    Duration         = "-"
                }
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

                $this.backupActivityReport += [PSCustomObject]@{
                    vCenter          = $vcenterName
                    Type             = $type
                    Status           = if ($job.state) { $job.state } else { $job.status }
                    "Data Transfer"  = $sizeGB
                    Location         = $location
                    StartTime        = Get-Date $job.start_time -Format "yyyy-MM-dd HH:mm:ss"
                    EndTime          = Get-Date $job.end_time -Format "yyyy-MM-dd HH:mm:ss"
                    Duration         = $duration
                }
            }

            Write-Host " -> OK ($($backupHistory.Count) trabajos procesados)." -ForegroundColor Green
        }
        catch {
            Write-Warning "`nError en actividad de backup: $($_.Exception.Message)"
            $this.backupActivityReport += [PSCustomObject]@{
                vCenter = $vcenterName; Type = "-"; Status = "Error API: $($_.Exception.Message)"; "Data Transfer" = "-"; Location = "-"; StartTime = "-"; EndTime = "-"; Duration = "-"
            }
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

    [PSCustomObject] GetIoHclReference() {
        if ($null -eq $this.ioHclRef) {
            $this.ioHclRef = Get-Content -Path (Join-Path $this.ReferenceDataFolder "vmware-iohcl.json") | ConvertFrom-Json
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

function ConvertFrom-PemCertificate([string] $pem) {
    $base64 = $pem -replace "-----BEGIN CERTIFICATE-----", "" -replace "-----END CERTIFICATE-----", "" -replace "\s", ""
    return New-Object System.Security.Cryptography.X509Certificates.X509Certificate2(, [System.Convert]::FromBase64String($base64))
}
