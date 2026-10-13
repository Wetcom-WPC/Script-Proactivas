# Step 2 of the pipeline: turn a collection JSON into the "raw" Excel (one sheet per
# collected topic, no filtering). The annex checks read these sheets by name.

Import-Module (Join-Path $PSScriptRoot '..\shared\ExcelStyle.psm1')

# Columns that Excel would otherwise turn into numbers (IPs, versions, hex PCI ids like
# "14e4" -> 140000, hostnames that look like numbers)
$script:textColumnsBySheet = @{
    'ESXi'                          = @('NtpServer', 'DnsServer', 'Hostname')
    'ESXi IO'                       = @('Hostname', 'Vid', 'Did', 'Svid', 'ssid', 'Version', 'Firmware')
    'vCenter'                       = @('Version')
    'VM'                            = @('Host')
    'Datastores'                    = @('Hostname')
    'Standard Switch'               = @('ESXi')
    'VMkernel Adapters'             = @('Host', 'IP')
    'vNetwork'                      = @('Host')
    'Falso Positivo'                = @('Host')
    'Compatibilidad de Componentes' = @('Hostname')
}

# Writes <ExcelFolder>\<json name>.xlsx and returns its path. Empty sheets are skipped.
function ConvertTo-RawExcel {
    param(
        [Parameter(Mandatory = $true)] [string] $JsonPath,
        [Parameter(Mandatory = $true)] [string] $ExcelFolder
    )

    if (-not (Test-Path $ExcelFolder)) { New-Item -ItemType Directory -Path $ExcelFolder -Force | Out-Null }

    $jsonName = Split-Path -Path $JsonPath -Leaf
    $excelPath = Join-Path -Path $ExcelFolder -ChildPath ([System.IO.Path]::ChangeExtension($jsonName, ".xlsx"))
    if (Test-Path $excelPath) { Remove-Item $excelPath }

    Write-Host "`nConvirtiendo '$jsonName' a Excel..." -ForegroundColor Yellow
    $report = (Get-Content -Path $JsonPath -Raw | ConvertFrom-Json).Report

    foreach ($sheet in $report.PSObject.Properties) {
        if (-not $sheet.Value -or $sheet.Value.Count -eq 0) { continue }

        Write-Host "  -> Creando hoja: '$($sheet.Name)'..."
        $exportParams = @{
            Path          = $excelPath
            WorksheetName = $sheet.Name
            AutoSize      = $true
            FreezeTopRow  = $true
            AutoFilter    = $true
        }
        if ($script:textColumnsBySheet.ContainsKey($sheet.Name)) {
            $exportParams['NoNumberConversion'] = $script:textColumnsBySheet[$sheet.Name]
        }

        $sheet.Value | Export-Excel @exportParams
    }

    if (Test-Path $excelPath) {
        $package = Open-ExcelPackage -Path $excelPath
        foreach ($worksheet in $package.Workbook.Worksheets) {
            Set-HeaderStyle -Worksheet $worksheet
        }
        Close-ExcelPackage $package
        Write-Host "-> '$(Split-Path $excelPath -Leaf)' generado." -ForegroundColor Green
    }
    else {
        Write-Warning "'$jsonName' no tiene datos; no se genero Excel."
        return $null
    }

    return $excelPath
}

Export-ModuleMember -Function ConvertTo-RawExcel
