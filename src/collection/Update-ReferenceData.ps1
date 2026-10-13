# Maintenance utility (run by hand, not part of a visit): downloads fresh copies of the
# reference data in reference-data\. Stale data makes the HCL and driver checks report
# supported hardware as unsupported or "Sin datos", so refresh it from time to time.
#
# Each download is independent: one failing source does not stop the others, and a failed
# download keeps the previous file.

$dataFolder = Join-Path $PSScriptRoot 'reference-data'
. (Join-Path $PSScriptRoot 'HardwareCompatibility.ps1')

$downloads = [ordered]@{
    'vmware-iohcl.json' = 'https://www.virten.net/repo/vmware-iohcl.json'   # IO devices (NICs, HBAs) HCL
    'esxiReleases.json' = 'https://www.virten.net/repo/esxiReleases.json'   # ESXi build -> release
    'vsan-hcl.json'     = 'https://partnerweb.vmware.com/service/vsan/all.json'
}

foreach ($fileName in $downloads.Keys) {
    $target = Join-Path $dataFolder $fileName
    $temp = "$target.download"
    Write-Host "Actualizando $fileName..." -NoNewline
    try {
        Invoke-WebRequest -Uri $downloads[$fileName] -OutFile $temp -ErrorAction Stop
        # Only replace the current file when the download is valid JSON with data. virten.net
        # stopped publishing the server HCL after the Broadcom move and serves it empty
        # (totalCount 0): keeping the previous file is better than losing every model.
        $downloaded = Get-Content $temp -Raw | ConvertFrom-Json -ErrorAction Stop
        if ($downloaded.PSObject.Properties['totalCount'] -and $downloaded.totalCount -eq 0) {
            throw "la fuente devolvio una lista vacia; se conserva el archivo anterior"
        }
        Move-Item -Path $temp -Destination $target -Force
        Write-Host " OK" -ForegroundColor Green
    }
    catch {
        Remove-Item $temp -ErrorAction SilentlyContinue
        Write-Host " FALLO: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Server HCL: fallback copy of the Broadcom Compatibility Guide catalog, in the same format as
# the guide's "Export to CSV" (so a CSV exported by hand from the website can replace it)
Write-Host "Actualizando servers-hcl.csv (Broadcom Compatibility Guide)..." -NoNewline
try {
    Get-ServerHclOnline | ForEach-Object {
        [PSCustomObject][ordered]@{
            'Partner Name'       = $_.manufacturer
            'Model'              = $_.model
            'CPU Series'         = $_.cpuSeries
            'Supported Releases' = $_.releases -join "`n"
        }
    } | Export-Csv -Path (Join-Path $dataFolder 'servers-hcl.csv') -NoTypeInformation -Encoding utf8
    Write-Host " OK" -ForegroundColor Green
}
catch {
    Write-Host " FALLO: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "Actualizando referencia de VMware Tools por build de ESXi..." -NoNewline
try {
    $lines = (Invoke-WebRequest -Uri 'https://packages.vmware.com/tools/versions' -ErrorAction Stop).Content -split "`n" |
        Where-Object { $_ -notmatch "^#" }

    # Columns: Tools version code | ESXi version | ESXi build | Tools version | Tools build.
    # "esx/0.0" lines (Tools not bundled with any ESXi) have no ESXi build: skipped.
    $toolsReference = foreach ($line in $lines) {
        $fields = @($line.Split() | Where-Object { $_ })
        if ($fields.Count -lt 4 -or $fields[1] -eq 'esx/0.0') { continue }
        [PSCustomObject] @{
            toolsBuild   = $fields[0]
            esxiVersion  = $fields[1]
            toolsVersion = $fields[3]
            esxiBuild    = $fields[2]
        }
    }
    $toolsReference | ConvertTo-Json | Out-File (Join-Path $dataFolder 'vmtools-ref.json')
    Write-Host " OK" -ForegroundColor Green
}
catch {
    Write-Host " FALLO: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "Listo."
