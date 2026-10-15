# Hardware compatibility (HCL) lookups for servers (TR-02), IO devices (TR-04) and vSAN.
#
# Servers: the Broadcom Compatibility Guide "Systems / Servers" catalog. Downloaded in full on
# every run from the same API the guide's website uses (https://compatibilityguide.broadcom.com,
# no authentication); when that fails (no internet, API changed) it falls back to the CSV in
# reference-data\servers-hcl.csv (same format as the guide's "Export to CSV").
# The rest of reference-data\ is refreshed by hand with Update-ReferenceData.ps1.

$script:broadcomGuideUrl = 'https://compatibilityguide.broadcom.com'

# Downloads the whole server catalog from the Broadcom Compatibility Guide
function Get-ServerHclOnline {
    param([int] $TimeoutSec = 60)

    $body = @{ programId = 'server'; filters = @(); keyword = @(); date = @{} } | ConvertTo-Json -Compress

    # One single page: the API's paging has no stable order, so paging repeats some rows and
    # skips others. If the catalog ever outgrows the page, ask again with its full size.
    $pageSize = 10000
    for ($attempt = 1; $attempt -le 2; $attempt++) {
        $response = Invoke-RestMethod -Method Post -TimeoutSec $TimeoutSec `
            -Uri "$script:broadcomGuideUrl/compguide/programs/viewResults?limit=$pageSize&page=1&sortBy=&sortType=ASC" `
            -Headers @{ 'USER-PERSONA' = 'live' } -ContentType 'application/json' -Body $body -ErrorAction Stop

        if (-not $response.success -or -not $response.data.fieldValues) {
            throw "Respuesta inesperada del Broadcom Compatibility Guide."
        }
        if (@($response.data.fieldValues).Count -ge [int]$response.data.count) { break }
        $pageSize = [int]$response.data.count
    }

    $entries = [System.Collections.Generic.List[object]]::new()
    foreach ($row in $response.data.fieldValues) {
        $entries.Add([PSCustomObject]@{
            manufacturer = ([string]$row.partnerName -replace '\s+', ' ').Trim()
            model        = ([string]$row.model[0].name).Trim()
            cpuSeries    = ($row.cpuSeries.name | ForEach-Object { ([string]$_).Trim() }) -join ', '
            releases     = @($row.supportedReleases.name | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ })
            url          = "$script:broadcomGuideUrl/detail$($row.model[0].url)"
        })
    }

    # A catalog this small means the API changed, not that the HCL shrank
    if ($entries.Count -lt 1000) { throw "El catalogo descargado tiene solo $($entries.Count) servidores." }
    return $entries
}

# Reads a server catalog exported from the guide ("Systems _ Servers_<date>.csv"). The website
# export turns every comma into a line break (even inside names: "Macmini6,2"), so they are
# turned back into commas; "Supported Releases" really is one release per line.
function Import-ServerHclCsv {
    param([Parameter(Mandatory = $true)] [string] $Path)

    Import-Csv -Path $Path | ForEach-Object {
        [PSCustomObject]@{
            manufacturer = ($_.'Partner Name' -replace '\r?\n', ',' -replace '\s+', ' ').Trim()
            model        = ($_.Model -replace '\r?\n', ',').Trim()
            cpuSeries    = ($_.'CPU Series' -replace '\r?\n', ',').Trim()
            releases     = @($_.'Supported Releases' -split '\r?\n' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
            url          = $null
        }
    }
}

# A model name reduced to what identifies the machine, so the HCL spelling and the one the
# host reports in its BIOS compare equal: without the vendor prefix ("Lenovo System x3650 M5"
# vs "System x3650 M5: -[8871AC1]-") and without the word "Certified" ("ThinkAgile VX7531
# Certified Node" vs "ThinkAgile VX7531 Node").
function Get-ModelKey {
    param([string] $Model)

    $key = $Model -replace '^\s*(Lenovo|IBM|HPE|HP|Dell( EMC)?|Cisco|Fujitsu|Hitachi|NEC|Supermicro)\s+', ''
    $key = $key -replace '\bCertified\b', ''
    return (($key -replace '\s+', ' ').Trim())
}

# Reference data for Test-HostHcl, loaded once per run: the server catalog (online, or the CSV
# when the guide is not reachable) and the ESXi build list.
function Get-HclReferenceData {
    param([Parameter(Mandatory = $true)] [string] $ReferenceDataFolder)

    if ($script:hclReferenceData) { return $script:hclReferenceData }

    $csvPath = Join-Path $ReferenceDataFolder 'servers-hcl.csv'
    try {
        Write-Host "`tDescargando la HCL de servidores de Broadcom..." -NoNewline
        $servers = Get-ServerHclOnline
        $source = "Broadcom Compatibility Guide (en linea, $(Get-Date -Format 'yyyy-MM-dd'))"
        Write-Host " OK ($($servers.Count) modelos)." -ForegroundColor Green
    }
    catch {
        Write-Host ""
        Write-Warning "No se pudo consultar el Broadcom Compatibility Guide ($($_.Exception.Message)). Se usa la copia local: $csvPath"
        $servers = @(Import-ServerHclCsv -Path $csvPath)
        $source = "CSV local ($((Get-Item $csvPath).LastWriteTime.ToString('yyyy-MM-dd')))"
    }

    foreach ($server in $servers) {
        Add-Member -InputObject $server -NotePropertyName ModelKey -NotePropertyValue (Get-ModelKey $server.model) -Force
    }

    $script:hclReferenceData = [PSCustomObject]@{
        Servers      = $servers
        Source       = $source
        EsxiReleases = (Get-Content -Path (Join-Path $ReferenceDataFolder "esxiReleases.json") | ConvertFrom-Json).data.esxiReleases
    }
    return $script:hclReferenceData
}

# Host release in the HCL's format ("ESXi 8.0 U3"). The build list gives the update level in
# releaseLevel ("VMware ESXi 8.0.0 Update 3"); from 7.0 on, the third digit of the version is
# the update number, which covers builds newer than the list.
function Resolve-HostRelease {
    param(
        [Parameter(Mandatory = $true)] $VMHost,
        $EsxiReleases
    )

    $release = $EsxiReleases | Where-Object build -eq $VMHost.Build | Select-Object -First 1
    if ($release) {
        if ($release.releaseLevel -match 'ESXi (\d+\.\d+)\.\d+ Update (\d+)') { return "ESXi $($matches[1]) U$($matches[2])" }
        return $release.updateRelease
    }

    if ($VMHost.Version -match '^(\d+)\.(\d+)\.(\d+)' -and [int]$matches[1] -ge 7) {
        $name = "ESXi $($matches[1]).$($matches[2])"
        if ([int]$matches[3] -gt 0) { $name += " U$($matches[3])" }
        return $name
    }
    return $null
}

# Pattern that the HCL's "CPU Series" must match for the host's CPU, or $null when the CPU
# family cannot be derived from its name (then every entry of the model is considered).
function Get-CpuSeriesPattern {
    param([string] $HostCpu)

    # Xeon Scalable: "Gold 6230" -> "6200" ("Gold 6200/5200 (Cascade-Lake-SP...)")
    if ($HostCpu -match '(Platinum|Gold|Silver|Bronze)\s+(\d)(\d)\d\d') { return "*$($matches[2])$($matches[3])00*" }
    # Xeon 6: "Xeon(R) 6740P" -> "6700P" ("Xeon 6500P/6700P (Granite Rapids-SP)")
    if ($HostCpu -match 'Xeon\(R\)\s+(\d)(\d)\d\d([PE])\b') { return "*$($matches[1])$($matches[2])00$($matches[3])*" }
    # AMD EPYC: "EPYC 7543" -> "7003", "EPYC 7F72" -> "7002" ("7002/7Fx2/7Hx2")
    if ($HostCpu -match 'EPYC\s+(\d)\w\w(\d)') { return "*$($matches[1])00$($matches[2])*" }
    return $null
}

# Server model/CPU support for the host's ESXi release.
# Model matching is based on Check-HCL by Florian Grehl (www.virten.net): an HCL entry matches
# when its model name is contained in the host's model. Vendors that use a different model
# string in the HCL and in the BIOS may be reported as not found.
#
# Supported: $true / $false, or $null when it cannot be decided (release unknown, or the model
# is in the HCL but the CPU family could not be matched).
function Test-HostHcl {
    param(
        [Parameter(Mandatory = $true)] $VMHost,
        [Parameter(Mandatory = $true)] $HclData
    )

    $hostModel = [string]$VMHost.Model
    $hostCpu = [string]$VMHost.ProcessorType

    $result = [PSCustomObject]@{
        VMHost            = $VMHost.Name
        Build             = $VMHost.Build
        ReleaseLevel      = Resolve-HostRelease -VMHost $VMHost -EsxiReleases $HclData.EsxiReleases
        Manufacturer      = $VMHost.Manufacturer
        IsNutanix         = ([string]$VMHost.Manufacturer -match 'Nutanix')
        IsVirtual         = ([string]$VMHost.Manufacturer -match '^VMware')   # nested ESXi ("VMware7,1")
        Model             = $hostModel
        Cpu               = $hostCpu
        Supported         = $null
        SupportedReleases = $null
        Reference         = $null
        Note              = $null
    }

    # Nested ESXi hosts run on a virtual machine: there is no server model to look up
    if ($result.IsVirtual) {
        $result.Note = "Host virtual (ESXi anidado): la HCL de servidores no aplica."
        return $result
    }

    $hostKey = Get-ModelKey $hostModel
    $modelEntries = @($HclData.Servers | Where-Object {
        $_.ModelKey -and $hostKey.IndexOf($_.ModelKey, [StringComparison]::OrdinalIgnoreCase) -ge 0
    })
    if ($modelEntries.Count -eq 0) {
        $result.Supported = $false
        $result.Note = "Modelo no encontrado en la HCL. Fuente: $($HclData.Source)"
        return $result
    }

    # Most specific model name first ("DL380 Gen10 Plus" over "DL380 Gen10")
    $longestModel = ($modelEntries | Sort-Object { $_.ModelKey.Length } -Descending | Select-Object -First 1).ModelKey.Length
    $modelEntries = @($modelEntries | Where-Object { $_.ModelKey.Length -eq $longestModel })

    $cpuPattern = Get-CpuSeriesPattern -HostCpu $hostCpu
    $entries = if ($cpuPattern) { @($modelEntries | Where-Object { $_.cpuSeries -like $cpuPattern }) } else { $modelEntries }

    if ($entries.Count -eq 0) {
        $result.Note = "Modelo en la HCL, pero no se encontro su familia de CPU ($hostCpu). Fuente: $($HclData.Source)"
        $result.SupportedReleases = @($modelEntries.releases | Select-Object -Unique)
        $result.Reference = ($modelEntries | Select-Object -First 1).url
        return $result
    }

    $releases = @($entries.releases | Select-Object -Unique)
    $result.SupportedReleases = $releases
    $result.Reference = ($entries | Select-Object -First 1).url

    if (-not $result.ReleaseLevel) {
        $result.Note = "No se pudo determinar la release del build $($VMHost.Build). Fuente: $($HclData.Source)"
    } else {
        $result.Supported = $releases -contains $result.ReleaseLevel
        $result.Note = "Fuente: $($HclData.Source)"
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
