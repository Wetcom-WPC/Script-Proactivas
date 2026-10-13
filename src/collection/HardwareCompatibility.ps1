# Hardware compatibility (HCL) lookups against the JSON files in reference-data\.
# Refresh those files with Update-ReferenceData.ps1; stale files make supported hosts look
# unsupported (their build or model is simply missing from the data).

function Get-HclReferenceData {
    param([Parameter(Mandatory = $true)] [string] $ReferenceDataFolder)

    if (-not $script:hclReferenceData) {
        $script:hclReferenceData = [PSCustomObject]@{
            Servers      = (Get-Content -Path (Join-Path $ReferenceDataFolder "vmware-hcl.json") | ConvertFrom-Json).data.server
            EsxiReleases = (Get-Content -Path (Join-Path $ReferenceDataFolder "esxiReleases.json") | ConvertFrom-Json).data.esxiReleases
        }
    }
    return $script:hclReferenceData
}

# Server model/CPU support for the host's ESXi release.
# Based on Check-HCL by Florian Grehl (www.virten.net), which uses the JSON HCL he maintains.
# Works well with HP and Dell, acceptably with IBM and Cisco. Vendors that use a different
# model string in the HCL and in the BIOS may be reported as unsupported.
function Test-HostHcl {
    param(
        [Parameter(Mandatory = $true)] $VMHost,
        [Parameter(Mandatory = $true)] $HclData
    )

    $hostManufacturer = $VMHost.Manufacturer
    $hostModel = $VMHost.Model
    $hostCpu = $VMHost.ProcessorType

    $result = [PSCustomObject]@{
        VMHost            = $VMHost.Name
        Build             = $VMHost.Build
        ReleaseLevel      = $null
        Manufacturer      = $hostManufacturer
        Model             = $hostModel
        Cpu               = $hostCpu
        Supported         = $null
        SupportedReleases = $null
        Reference         = $null
        Note              = $null
    }

    $release = $HclData.EsxiReleases | Where-Object build -eq $VMHost.Build
    if ($release) {
        $updateRelease = $release.updateRelease
        $result.ReleaseLevel = $updateRelease
    }
    else {
        $updateRelease = $false
        $result.Note = "ESXi Build $($VMHost.Build) not found in database."
    }

    $matchingEntries = @()
    foreach ($server in $HclData.Servers) {
        $modelFound = $false
        $modelPattern = "*" + ($server.model -replace "IBM ", "") + "*"

        # Cisco UCS: match model number and generation separately
        if ($hostModel.StartsWith("UCS") -and $server.model.Contains("UCS")) {
            $ucsModel = $hostModel.Substring(5, 4)
            $ucsGeneration = $null
            if ($hostModel.Length -eq 12) { $ucsGeneration = $hostModel.Substring(10, 2) }
            if ($hostModel.Length -eq 14) { $ucsGeneration = $hostModel.Substring(10, 3) }

            if ($server.model.Contains($ucsModel) -and $ucsGeneration -and $server.model.Contains($ucsGeneration)) {
                $modelFound = $true
            }
        }

        if ($hostManufacturer -eq "HP") {
            if ($hostModel -like $modelPattern -and $server.manufacturer -eq $hostManufacturer) {
                $modelFound = $true
            }
        }
        elseif ($hostModel -like $modelPattern) {
            $modelFound = $true
        }

        if (-not $modelFound) { continue }

        if ($server.cpuSeries -like "Intel Xeon*") {
            if ($hostCpu -like "*Gold*" -or $hostCpu -like "*Silver*" -or $hostCpu -like "*Bronze*" -or $hostCpu -like "*Platinum*") {
                # Scalable Xeons: compare the series (e.g. "6200")
                $cpuNumber = $hostCpu -replace '\D+([0-9]*).*', '$1'
                if ($server.cpuSeries -notlike "*$($cpuNumber.Substring(0,2))00*") {
                    continue
                }
            }
            else {
                $cpuSeriesPattern = "*" + ($server.cpuSeries -replace "Intel Xeon ", "" -replace " Series", "" -replace "00", "??" -replace "xx", "??" -replace "-v", " v") + " @*"
                $hostCpuNormalized = $hostCpu -replace " 0 @", " @" -replace "- ", "-" -replace "  ", " "
                if ($hostCpuNormalized -notlike $cpuSeriesPattern) {
                    continue
                }
            }
        }

        $matchingEntries += [PSCustomObject]@{
            Model    = $server.model
            CPU      = $server.cpuSeries
            Releases = $server.releases
            URL      = $server.url
        }
    }

    if ($matchingEntries.Count -eq 1) {
        $entry = $matchingEntries[0]
        $result.Supported = [bool]($updateRelease -and ($entry.Releases -contains $updateRelease))
        $result.SupportedReleases = $entry.Releases
        $result.Reference = $entry.URL
    }
    elseif ($matchingEntries.Count -gt 1) {
        $result.Note = "More than 1 HCL Entries found."
    }
    else {
        $result.Supported = $false
        $result.Note = "No HCL Entries found."
    }

    return $result
}

# vSAN HCL link(s) for a storage controller, by PCI ids
function Get-VsanHclLink($vid, $did, $svid, $ssid, $vsanHcl) {
    foreach ($entry in $vsanHcl.data.controller) {
        if (($vid -eq $entry.vid) -and ($did -eq $entry.did) -and ($svid -eq $entry.svid) -and ($ssid -eq $entry.ssid)) {
            $entry.vcglink
        }
    }
}
