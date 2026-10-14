# Quarterly checks (VSP-TR-*). Dot-sourced by Report.psm1. Same contract as MonthlyChecks.ps1.

# General support end dates per ESXi major version
$script:esxiEndOfSupport = @{
    '8.0' = '2027-10-11'
    '7.0' = '2025-10-02'
    '6.7' = '2022-10-15'
    '6.5' = '2022-10-15'
}

function Test-EsxiEndOfSupport($Context) {
    $findings = @()
    $foundExpired = $false
    $foundExpiringSoon = $false

    $today = Get-Date
    $oneYearFromNow = $today.AddYears(1)

    foreach ($row in (Import-RawSheet $Context.RawExcelPath "ESXi")) {
        if (-not $row.PSObject.Properties['ESXiVersion']) { continue }

        $fullVersion = $row.ESXiVersion
        $majorVersion = if ($fullVersion -match '(\d+\.\d+)') { $matches[1] } else { $null }
        if (-not ($majorVersion -and $script:esxiEndOfSupport.ContainsKey($majorVersion))) { continue }

        $endOfSupport = [datetime]$script:esxiEndOfSupport[$majorVersion]
        $isFinding = $false

        if ($endOfSupport -lt $today) {
            $foundExpired = $true
            $isFinding = $true
        } elseif ($endOfSupport -lt $oneYearFromNow) {
            $foundExpiringSoon = $true
            $isFinding = $true
        }

        if ($isFinding) {
            $findings += [PSCustomObject]@{
                "vCenter"      = $row."vCenter"
                "Hostname"     = $row."Hostname"
                "Datacenter"   = $row."Datacenter"
                "Cluster"      = $row."Cluster"
                "Version"      = $fullVersion
                "Fecha de EoS" = $endOfSupport.ToString('yyyy-MM-dd')
            }
        }
    }

    $result = "Resultado Esperado"
    $detail = "Todos los productos tienen soporte vigente por más de un año."

    if ($foundExpired) {
        $result = "Vencido"
        $detail = "Se encontraron productos fuera de soporte"
    } elseif ($foundExpiringSoon) {
        $result = "Vencimiento < 1 Año"
        $detail = "Se encontraron productos con vencimiento de soporte en menos de 365 dias"
    }

    if ($findings.Count -gt 0) {
        Export-AnnexSheet -Data $findings -Path $Context.AnnexPath -WorksheetName "End of Support" -NoNumberConversion "Hostname"
    }

    return [PSCustomObject]@{ ID = "VSP-TR-01"; Resultado = $result; Detalle = $detail }
}

# Server hardware support for the host's ESXi release (HCL lookup done during collection).
# Hosts whose support could not be decided ("Sin datos": model in the HCL but CPU family not
# matched, or unknown release) are not findings; the detail mentions them for manual review.
#
# Nutanix hosts are validated against Nutanix's own compatibility matrix, not Broadcom's HCL:
# they are taken out of TR-02 and get an extra manual checklist row (see below).
function Test-HardwareCompatibility($Context) {
    $findings = @()
    $noDataHosts = @()
    $nutanixHosts = @()
    $totalHosts = 0

    foreach ($row in (Import-RawSheet $Context.RawExcelPath "ESXi")) {
        if (-not $row.PSObject.Properties['Supported']) { continue }
        $totalHosts++

        if ($row.Supported.Trim() -eq "Nutanix") {
            $nutanixHosts += $row.Hostname
        }
        elseif ($row.Supported.Trim() -eq "Sin datos") {
            $noDataHosts += $row.Hostname
        }
        elseif ($row.Supported.Trim() -eq "False") {
            $findings += [PSCustomObject]@{
                "vCenter"                   = $row."vCenter"
                "Hostname"                  = $row."Hostname"
                "Datacenter"                = $row."Datacenter"
                "Cluster"                   = $row."Cluster"
                "Version"                   = $row."ESXiVersion"
                "Version minima soportada"  = $row."Supported Releases"
                "Supported"                 = $row."Supported"
            }
        }
    }

    $noDataNote = if ($noDataHosts.Count -gt 0) { " $($noDataHosts.Count) hosts sin datos suficientes para validarlos (revisar manualmente: $($noDataHosts -join ', '))." } else { "" }

    if ($findings.Count -gt 0) {
        Export-AnnexSheet -Data $findings -Path $Context.AnnexPath -WorksheetName "Compatibilidad de Componentes"
    }

    # Nutanix hosts get their own checklist row (VSP-TR-02-NX, created right below TR-02): they
    # are validated against Nutanix's compatibility matrix by the client's responsible, not
    # against Broadcom's HCL. TR-02 itself only covers the other hosts.
    $nutanixNote = if ($nutanixHosts.Count -gt 0) { " Los $($nutanixHosts.Count) hosts Nutanix se validan aparte (fila VSP-TR-02-NX)." } else { "" }

    if ($findings.Count -gt 0) {
        $result = [PSCustomObject]@{ ID = "VSP-TR-02"; Resultado = "No recomendado"; Detalle = "Se encontraron $($findings.Count) componentes no soportados por la version de ESXi actual.$noDataNote$nutanixNote" }
    }
    elseif ($nutanixHosts.Count -gt 0 -and $noDataHosts.Count -eq 0 -and $totalHosts -eq $nutanixHosts.Count) {
        # Every host is Nutanix: nothing left for the HCL to evaluate
        $result = [PSCustomObject]@{ ID = "VSP-TR-02"; Resultado = $null; Detalle = "Todos los hosts del entorno son Nutanix: no hay hosts para validar contra la HCL de Broadcom.$nutanixNote" }
    }
    else {
        $result = [PSCustomObject]@{ ID = "VSP-TR-02"; Resultado = "Resultado Esperado"; Detalle = "Todos los componentes analizados son compatibles con su versión de ESXi.$noDataNote$nutanixNote" }
    }

    if ($nutanixHosts.Count -eq 0) { return $result }

    $nutanixRow = [PSCustomObject]@{
        ID            = "VSP-TR-02-NX"
        Resultado     = $null
        Detalle       = "Tarea manual a cargo del responsable del cliente: validar los $($nutanixHosts.Count) hosts Nutanix contra la Nutanix Compatibility & Interoperability Matrix (version de AOS vs ESXi). Hosts: $($nutanixHosts -join ', ')."
        InsertAfter   = "VSP-TR-02"
        Procedimiento = "Manual"
        Chequeo       = "Verificación de compatibilidad de componentes - hosts Nutanix (a cargo del responsable del cliente)"
    }
    return $result, $nutanixRow
}

function Test-NtpConfiguration($Context) {
    $findings = @()
    $foundNotConfigured = $false
    $foundServiceStopped = $false
    $foundInconsistency = $false

    foreach ($clusterHosts in (Get-HostsByCluster (Import-RawSheet $Context.RawExcelPath "ESXi")).Values) {
        if ($clusterHosts.Where({ [string]::IsNullOrEmpty($_.NtpServer) })) { $foundNotConfigured = $true }
        if ($clusterHosts.Where({ $_.NtpdRunning -ne $true })) { $foundServiceStopped = $true }
        if (($clusterHosts.NtpServer | Select-Object -Unique).Count -gt 1) { $foundInconsistency = $true }

        $uniqueNtpServers = $clusterHosts | ForEach-Object { if ($_) { $_.NtpServer.Trim() } } | Select-Object -Unique
        $stoppedHosts = $clusterHosts | Where-Object { $_.NtpdRunning -ne $true }

        if ($uniqueNtpServers.Count -gt 1) {
            $findings += $clusterHosts
        }
        elseif ($stoppedHosts.Count -gt 0) {
            $findings += $stoppedHosts
        }
    }

    $result = "Resultado Esperado"
    $detail = "Todos los hosts tienen NTP configurado, en ejecucion y consistente por cluster."

    if ($foundNotConfigured) {
        $result = "No Configurado"
        $detail = "Se encontraron hosts sin un servidor NTP configurado"
    } elseif ($foundServiceStopped) {
        $result = "Configurado y servicio no activo"
        $detail = "Se encontraron hosts con el servicio NTP detenido"
    } elseif ($foundInconsistency) {
        $result = "Cluster no consistente"
        $detail = "Se encontraron inconsistencias en la configuracion de NTP dentro de al menos un cluster"
    }

    if ($findings.Count -gt 0) {
        $annexRows = $findings | Sort-Object -Property Hostname -Unique | ForEach-Object {
            [PSCustomObject]@{
                "vCenter"     = $_."vCenter"
                "Hostname"    = $_."Hostname"
                "Datacenter"  = $_."Datacenter"
                "Cluster"     = $_."Cluster"
                "NtpServer"   = $_."NtpServer"
                "NtpdRunning" = $_."NtpdRunning"
            }
        }
        Export-AnnexSheet -Data $annexRows -Path $Context.AnnexPath -WorksheetName "NTP Problemas" -NoNumberConversion "Hostname", "NtpServer"
    }

    return [PSCustomObject]@{ ID = "VSP-TR-03"; Resultado = $result; Detalle = $detail }
}

# Network cards and storage controllers ("ESXi IO" sheet of the DriverInfo Excel) against the
# IO HCL: is each device listed for its host's ESXi release?
#
# Driver and firmware versions cannot be validated automatically (the reference data only
# lists supported releases per device), so when every device is listed the result is left
# blank for the technician to confirm against the URL of each device (KB VSP-TR-04).
# When the host's release is newer than the reference data, the device is "Sin datos" instead
# of incompatible: refresh the data with src\collection\Update-ReferenceData.ps1.
function Test-HostDrivers($Context) {
    $devices = Import-RawSheet $Context.DriverExcelPath "ESXi IO"

    if (-not $devices) {
        return [PSCustomObject]@{ ID = "VSP-TR-04"; Resultado = $null; Detalle = "No se recolectaron drivers de los hosts (hoja 'ESXi IO' vacía o no encontrada)." }
    }

    $releaseCatalog = Get-EsxiReleaseCatalog -ReferenceDataFolder $Context.ReferenceDataFolder
    $knownReleases = @($releaseCatalog.IoHclReleases)

    $annexRows = @()
    $incompatibleCount = 0
    $notInHclCount = 0
    $noDataCount = 0

    foreach ($device in $devices) {
        $hostRelease = Resolve-EsxiUpdateRelease -FullName $device."ESXi Release" -Catalog $releaseCatalog
        $supportedReleases = @(([string]$device."ESXi Supported Releases") -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ })

        if ($device.URL -eq "No encontrado en HCL") {
            $status = "No encontrada en HCL"
            $notInHclCount++
        } elseif (-not $hostRelease -or $hostRelease -notin $knownReleases) {
            $status = "Sin datos (HCL desactualizada)"
            $noDataCount++
        } elseif ($hostRelease -in $supportedReleases) {
            $status = "Sí"
        } else {
            $status = "No"
            $incompatibleCount++
        }

        $annexRows += [PSCustomObject]@{
            "vCenter"           = $device."vCenter"
            "Hostname"          = $device."Hostname"
            "Version del Host"  = $device."ESXi Release"
            "Placa"             = $device."Placa"
            "Controlador"       = $device."Controlador"
            "Vendor"            = $device."Vendor"
            "Driver"            = $device."Driver"
            "Version"           = $device."Version"
            "Firmware"          = $device."Firmware"
            "Vid"               = $device."Vid"
            "Did"               = $device."Did"
            "Svid"              = $device."Svid"
            "ssid"              = $device."ssid"
            "Release soportada" = $status
            "URL"               = $device."URL"
        }
    }

    Export-AnnexSheet -Data $annexRows -Path $Context.AnnexPath -WorksheetName "Host Drivers" -NoNumberConversion "Hostname", "Vid", "Did", "Svid", "ssid", "Version", "Firmware"

    # Devices missing from the IO HCL are usually onboard chipset controllers (e.g. SATA AHCI)
    # that the HCL does not list: they are reported in the annex but do not decide the result
    $notes = ""
    if ($notInHclCount -gt 0) { $notes += " $notInHclCount placas no figuran en la HCL (suelen ser controladoras del chipset; revisar en el anexo)." }
    if ($noDataCount -gt 0) { $notes += " $noDataCount placas sin datos de compatibilidad para su release (HCL de referencia desactualizada)." }

    if ($incompatibleCount -gt 0) {
        return [PSCustomObject]@{
            ID        = "VSP-TR-04"
            Resultado = "No compatible"
            Detalle   = "Se encontraron $incompatibleCount placas no soportadas en la release de su host. Ver anexo.$notes"
        }
    }

    return [PSCustomObject]@{
        ID        = "VSP-TR-04"
        Resultado = $null
        Detalle   = "Revisar manualmente la version de driver y firmware de cada placa en su URL de la HCL (ver hoja 'Host Drivers' del anexo).$notes"
    }
}

function Test-DnsConfiguration($Context) {
    $findings = @()

    foreach ($clusterHosts in (Get-HostsByCluster (Import-RawSheet $Context.RawExcelPath "ESXi")).Values) {
        $hasEmptyDns = $clusterHosts.Where({ [string]::IsNullOrEmpty($_.DnsServer) }).Count -gt 0
        $isInconsistent = ($clusterHosts.DnsServer.Trim() | Select-Object -Unique).Count -gt 1

        if ($hasEmptyDns -or $isInconsistent) {
            $findings += $clusterHosts
        }
    }

    if ($findings) {
        $annexRows = $findings | ForEach-Object {
            [PSCustomObject]@{
                "vCenter"         = $_."vCenter"
                "Hostname"        = $_."Hostname"
                "Datacenter"      = $_."Datacenter"
                "Cluster"         = $_."Cluster"
                "ConnectionState" = $_."ConnectionState"
                "DnsServer"       = $_."DnsServer"
            }
        }
        Export-AnnexSheet -Data $annexRows -Path $Context.AnnexPath -WorksheetName "Host DNS" -NoNumberConversion "Hostname"

        return [PSCustomObject]@{ ID = "VSP-TR-05"; Resultado = "Cluster No Consistente"; Detalle = "Al menos un host no tiene un servidor DNS configurado, o no hay consistencia en un cluster" }
    }

    return [PSCustomObject]@{ ID = "VSP-TR-05"; Resultado = "Resultado Esperado"; Detalle = "Todos los host tienen servidores DNS configurados y son consistenetes por cluster" }
}

function Test-Licenses($Context) {
    $findings = @()
    $foundExpired = $false
    $foundExpiringSoon = $false
    $foundEvaluation = $false
    $foundOverused = $false
    $foundDuplicated = $false

    $today = Get-Date
    $oneYearFromNow = $today.AddYears(1)

    $licenses = @(Import-RawSheet $Context.RawExcelPath "vLicense")

    # 1. Expiration, usage and evaluation per license
    foreach ($license in $licenses) {
        $expiration = $license.ExpirationDate
        $used = 0
        [void][int]::TryParse([string]$license.Used, [ref]$used)
        $total = 0
        $hasLimit = [int]::TryParse([string]$license.Total, [ref]$total)   # "Unlimited" has no limit
        $isFinding = $false

        if ($hasLimit -and $used -gt $total) {
            $foundOverused = $true
            $isFinding = $true
        }

        if ($expiration -eq "Never") {
            # Perpetual license: never expires
        } elseif ($expiration -eq "Evaluation") {
            # Every vCenter lists a "Product Evaluation" entry; it only matters when something uses it
            if ($used -gt 0) {
                $foundEvaluation = $true
                $isFinding = $true
            }
        } else {
            try {
                $expirationDate = [datetimeoffset]::Parse($expiration).DateTime
                if ($expirationDate -lt $today) {
                    $foundExpired = $true
                    $isFinding = $true
                } elseif ($expirationDate -lt $oneYearFromNow) {
                    $foundExpiringSoon = $true
                    $isFinding = $true
                }
            } catch {}
        }
        if ($isFinding) {
            $findings += $license
        }
    }

    # 2. The same license key registered more than once
    $duplicated = $licenses | Group-Object -Property LicenseKey | Where-Object { $_.Count -gt 1 } | ForEach-Object { $_.Group }
    if ($duplicated.Count -gt 0) {
        $foundDuplicated = $true
        $findings += $duplicated
    }

    # Only values from the checklist dropdown: Resultado Esperado / Vencimiento < 1 Año /
    # Vencido / Excepción. Evaluation mode expires within 60 days, so it counts as "< 1 Año".
    $result = "Resultado Esperado"
    $detail = "No se encontraron licencias vencidas, por vencer ni en modo de evaluación."
    if ($foundExpired) {
        $result = "Vencido"
        $detail = "Se encontraron licencias vencidas. Ver anexo."
    } elseif ($foundEvaluation) {
        $result = "Vencimiento < 1 Año"
        $detail = "Se encontraron licencias en modo de evaluación en uso (vencen a los 60 días). Ver anexo."
    } elseif ($foundExpiringSoon) {
        $result = "Vencimiento < 1 Año"
        $detail = "Se encontraron licencias con vencimiento menor a un año. Ver anexo."
    } elseif ($foundOverused) {
        $result = "Excepción"
        $detail = "Se encontraron licencias cuyo uso excede el límite total. Ver anexo."
    } elseif ($foundDuplicated) {
        $result = "Excepción"
        $detail = "Se encontraron License Keys duplicadas. Ver anexo."
    }

    if ($findings.Count -gt 0) {
        $uniqueFindings = $findings | Group-Object -Property vCenter, LicenseKey, Name | ForEach-Object { $_.Group | Select-Object -First 1 }

        $annexRows = $uniqueFindings | Sort-Object -Property Name, LicenseKey | ForEach-Object {
            [PSCustomObject]@{
                Licencia           = $_.Name
                LicenseKey         = $_.LicenseKey
                Componente         = $_.ProductName
                vCenter            = $_.vCenter
                "En Uso"           = $_.Used
                Total              = $_.Total
                "Fecha Expiracion" = $_.ExpirationDate
            }
        }
        Export-AnnexSheet -Data $annexRows -Path $Context.AnnexPath -WorksheetName "Licencias Problemas"
    }

    return [PSCustomObject]@{ ID = "VSP-TR-06"; Resultado = $result; Detalle = $detail }
}

function Test-VCenterCertificates($Context) {
    $certificates = @(Import-RawSheet $Context.RawExcelPath "Certficate" | Where-Object { -not [string]::IsNullOrEmpty($_.vCenter) })

    if ($certificates.Count -eq 0) {
        return [PSCustomObject]@{ ID = "VSP-TR-07"; Resultado = "Sin acceso"; Detalle = "No se pudo recolectar información de certificados o la hoja 'Certficate' está vacía." }
    }

    $findings = @()
    $today = Get-Date
    $oneYearFromNow = $today.AddYears(1)

    $foundExpired = $false
    $foundExpiringSoon = $false
    $foundErrors = $false

    foreach ($row in $certificates) {
        $validUntil = $row."Valid Until"
        $daysRemaining = $null
        $isFinding = $false

        if ([string]::IsNullOrEmpty($validUntil) -or $validUntil -eq "N/A") {
            $isFinding = $true
            $foundErrors = $true
        }
        else {
            try {
                $expirationDate = Get-Date $validUntil -ErrorAction Stop
                $daysRemaining = ($expirationDate - $today).Days

                if ($expirationDate -lt $today) {
                    $isFinding = $true
                    $foundExpired = $true
                }
                elseif ($expirationDate -lt $oneYearFromNow) {
                    $isFinding = $true
                    $foundExpiringSoon = $true
                }
            }
            catch {
                $isFinding = $true
                $foundErrors = $true
            }
        }

        if ($isFinding) {
            $findings += [PSCustomObject]@{
                "vCenter"       = $row.vCenter
                "Ubicación"     = $row.Ubicacion
                "Subject"       = $row.Subject
                "Emisor"        = $row.Emisor
                "Valid From"    = $row."Valid From"
                "Valid Until"   = $validUntil
                "Días Restantes"= $daysRemaining
            }
        }
    }

    $result = "Resultado Esperado"
    $detail = "Todos los certificados tienen vigencia mayor a 1 año."

    if ($foundExpired) {
        $result = "Vencido"
        $detail = "Se detectaron certificados expirados."
    }
    elseif ($foundErrors) {
        $result = "Sin acceso"
        $detail = "Hubo errores al leer las fechas de algunos certificados."
    }
    elseif ($foundExpiringSoon) {
        $result = "Vencimiento < 1 Año"
        $detail = "Se detectaron certificados próximos a vencer (menos de 1 año)."
    }

    if ($findings.Count -gt 0) {
        Export-AnnexSheet -Data $findings -Path $Context.AnnexPath -WorksheetName "Certificados vCenter"
    }

    return [PSCustomObject]@{ ID = "VSP-TR-07"; Resultado = $result; Detalle = $detail }
}

function Test-EsxiCertificates($Context) {
    $hosts = @(Import-RawSheet $Context.RawExcelPath "ESXi" | Where-Object { -not [string]::IsNullOrEmpty($_.vCenter) })

    if ($hosts.Count -eq 0) {
        return [PSCustomObject]@{ ID = "VSP-TR-08"; Resultado = "Excepción"; Detalle = "No se pudo recolectar información de los Hosts ESXi o la hoja está vacía." }
    }

    $findings = @()
    $today = Get-Date
    $oneYearFromNow = $today.AddYears(1)

    $expiredCount = 0
    $expiringSoonCount = 0
    $nearestExpiration = $null

    foreach ($row in $hosts) {
        $validTo = $row."Cert Valid To"
        $isFinding = $false

        if ([string]::IsNullOrEmpty($validTo) -or $validTo -eq "N/A") {
            $isFinding = $true
        }
        else {
            try {
                $expirationDate = Get-Date $validTo -ErrorAction Stop

                if ($expirationDate -lt $today) {
                    $isFinding = $true
                    $expiredCount++
                }
                elseif ($expirationDate -lt $oneYearFromNow) {
                    $isFinding = $true
                    $expiringSoonCount++
                    if ($null -eq $nearestExpiration -or $expirationDate -lt $nearestExpiration) {
                        $nearestExpiration = $expirationDate
                    }
                }
            }
            catch {
                $isFinding = $true
            }
        }

        if ($isFinding) {
            $findings += [PSCustomObject]@{
                "vCenter"       = $row.vCenter
                "Hostname"      = $row.Hostname
                "Cert Valid To" = $validTo
                "Cert Issuer"   = $row."Cert Issuer"
            }
        }
    }

    $result = "Resultado Esperado"
    $detail = "Todos los certificados de host tienen vigencia mayor a 1 año."

    if ($expiredCount -gt 0) {
        $result = "Vencido"
        $detail = "Se detectaron $expiredCount certificados de host ESXi ya expirados."
    }
    elseif ($expiringSoonCount -gt 0) {
        $result = "Vencimiento < 1 Año"
        $nearestText = if ($nearestExpiration) { $nearestExpiration.ToString("yyyy-MM-dd") } else { "N/A" }
        $detail = "Se detectaron $expiringSoonCount certificados próximos a vencer. El más próximo vence el $nearestText."
    }

    if ($findings.Count -gt 0) {
        Export-AnnexSheet -Data $findings -Path $Context.AnnexPath -WorksheetName "Certificados ESXi"
    }

    return [PSCustomObject]@{ ID = "VSP-TR-08"; Resultado = $result; Detalle = $detail }
}

function Test-VCenterRootPassword($Context) {
    $vcenters = @(Import-RawSheet $Context.RawExcelPath "vCenter" | Where-Object { -not [string]::IsNullOrEmpty($_."vCenter Server") })

    if ($vcenters.Count -eq 0) {
        return [PSCustomObject]@{ ID = "VSP-TR-09"; Resultado = "Sin acceso"; Detalle = "No se pudo recolectar información del VAMI o la hoja 'vCenter' está vacía." }
    }

    $findings = @()
    $today = Get-Date
    $sixMonthsFromNow = $today.AddMonths(6)

    $foundExpired = $false
    $foundExpiringSoon = $false
    $unreadable = @()

    foreach ($row in $vcenters) {
        $expiration = $row."Expiration Date"
        $evaluation = "Desconocido"
        $isFinding = $false

        if ($expiration -eq "Nunca") {
            $evaluation = "Resultado Esperado"
        }
        elseif ($expiration -eq "N/A" -or $row.Status -like "*Error*" -or $row.Status -like "*Unknown*") {
            # The appliance API did not return the root account (often a permissions matter)
            $evaluation = "Error de lectura"
            $isFinding = $true
            $unreadable += $row."vCenter Server"
        }
        else {
            try {
                $expirationDate = Get-Date $expiration -ErrorAction Stop

                if ($expirationDate -lt $today) {
                    $evaluation = "Vencido"
                    $isFinding = $true
                    $foundExpired = $true
                }
                elseif ($expirationDate -lt $sixMonthsFromNow) {
                    $evaluation = "Vencimiento < 6 meses"
                    $isFinding = $true
                    $foundExpiringSoon = $true
                }
                else {
                    $evaluation = "Resultado Esperado"
                }
            }
            catch {
                $evaluation = "Error formato fecha"
                $isFinding = $true
            }
        }

        if ($isFinding) {
            $findings += [PSCustomObject]@{
                "vCenter"         = $row."vCenter Server"
                "Root User"       = $row."Root User"
                "Expiration Date" = $expiration
                "Days Remaining"  = $row."Days Remaining"
                "Status Original" = $row."Status"
                "Evaluación"      = $evaluation
            }
        }
    }

    # Dates decide the result. When no vCenter could be read at all, there is nothing to vouch
    # for: "Sin acceso" instead of a "Resultado Esperado" nobody checked.
    $unreadableNote = if ($unreadable.Count -gt 0) { " No se pudo leer la expiracion en: $($unreadable -join ', ') (revisar manualmente en la VAMI)." } else { "" }
    $result = "Resultado Esperado"
    $detail = "Las contraseñas de root no expiran o vencen en más de 6 meses.$unreadableNote"

    if ($foundExpired) {
        $result = "Vencido"
        $detail = "Se detectaron contraseñas de root ya expiradas.$unreadableNote"
    }
    elseif ($foundExpiringSoon) {
        $result = "Vencimiento < 6 meses"
        $detail = "Se detectaron contraseñas próximas a vencer (menos de 6 meses).$unreadableNote"
    }
    elseif ($unreadable.Count -eq $vcenters.Count) {
        $result = "Sin acceso"
        $detail = "No se pudo leer la expiracion de la password de root en ningun vCenter (la API de la VAMI no devolvio la cuenta; suele requerir permisos de administrador del appliance). Revisar manualmente en la VAMI."
    }

    if ($findings.Count -gt 0) {
        Export-AnnexSheet -Data $findings -Path $Context.AnnexPath -WorksheetName "Vencimiento password de Root"
    }

    return [PSCustomObject]@{ ID = "VSP-TR-09"; Resultado = $result; Detalle = $detail }
}

function Test-PerformanceStatistics($Context) {
    $clusters = @(Import-RawSheet $Context.RawExcelPath "PerformanceHealth" | Where-Object { -not [string]::IsNullOrEmpty($_.vCenter) })

    if ($clusters.Count -eq 0) {
        return [PSCustomObject]@{ ID = "VSP-TR-10"; Resultado = "Excepción"; Detalle = "No se pudo recolectar información de salud de performance o la hoja está vacía." }
    }

    $findings = @()

    foreach ($row in $clusters) {
        if ($row."Health DB" -ne "OK" -or $row."CPU Stats" -ne "OK" -or $row."Mem Stats" -ne "OK" -or
            $row."Net Stats" -ne "OK" -or $row."Disk Stats" -ne "OK") {
            $findings += [PSCustomObject]@{
                "vCenter"     = $row.vCenter
                "Cluster"     = $row.Cluster
                "Estado DB"   = $row."Health DB"
                "CPU"         = $row."CPU Stats"
                "Memoria"     = $row."Mem Stats"
                "Red"         = $row."Net Stats"
                "Disco"       = $row."Disk Stats"
            }
        }
    }

    if ($findings.Count -gt 0) {
        Export-AnnexSheet -Data $findings -Path $Context.AnnexPath -WorksheetName "Salud Performance"
        return [PSCustomObject]@{ ID = "VSP-TR-10"; Resultado = "No recomendado"; Detalle = "Se detectaron $($findings.Count) clústeres con problemas en la recolección de estadísticas." }
    }

    return [PSCustomObject]@{ ID = "VSP-TR-10"; Resultado = "Resultado Esperado"; Detalle = "Todos los clústeres están recolectando y almacenando métricas de rendimiento correctamente." }
}

function Test-NiocEnabled($Context) {
    # No vDS in the environment means no "vDS" sheet in the raw Excel: same as "no vDS"
    $vdSwitches = @(Import-RawSheet $Context.RawExcelPath "vDS")
    $findings = @($vdSwitches | Where-Object { $_."NIOC Enabled" -ne $true })

    if ($findings.Count -gt 0) {
        Export-AnnexSheet -Data $findings -Path $Context.AnnexPath -WorksheetName "NIOC Deshabilitado"
        return [PSCustomObject]@{ ID = "VSP-TR-11"; Resultado = "No recomendado"; Detalle = "Se encontraron $($findings.Count) vDS con NIOC deshabilitado" }
    }

    $detail = if ($vdSwitches.Count -gt 0) { "Todos los vDS encontrados tienen NIOC habilitado" } else { "No se encontraron vDS en el entorno" }
    return [PSCustomObject]@{ ID = "VSP-TR-11"; Resultado = "Resultado Esperado"; Detalle = $detail }
}

# Standard switch port groups must have the same names, switch and VLAN on every host of a cluster
function Test-StandardSwitchConsistency($Context) {
    $findings = @()
    $portGroups = Import-RawSheet $Context.RawExcelPath "Standard Switch" | Where-Object { -not [string]::IsNullOrEmpty($_.Cluster) }

    foreach ($cluster in ($portGroups | Group-Object -Property @{ Expression = { $_.vCenter + '|' + $_.Cluster } })) {
        if (($cluster.Group.ESXi | Select-Object -Unique).Count -le 1) {
            continue
        }

        # Each host's sorted list of port group names; more than one distinct list = inconsistent
        $portGroupLists = $cluster.Group | Group-Object -Property ESXi | ForEach-Object {
            ($_.Group.PortGroup | Select-Object -Unique | Sort-Object) -join ';'
        } | Select-Object -Unique
        $namesAreInconsistent = $portGroupLists.Count -gt 1

        $configIsInconsistent = $false
        if (-not $namesAreInconsistent) {
            foreach ($portGroup in ($cluster.Group | Group-Object -Property PortGroup)) {
                $configurations = $portGroup.Group | ForEach-Object { "$($_.Switch)|$($_.vLAN)" } | Select-Object -Unique
                if ($configurations.Count -gt 1) {
                    $configIsInconsistent = $true
                    break
                }
            }
        }

        if ($namesAreInconsistent -or $configIsInconsistent) {
            $findings += $cluster.Group
        }
    }

    if ($findings) {
        $annexRows = $findings | ForEach-Object {
            [PSCustomObject]@{
                "vCenter"   = $_."vCenter"
                "ESXi"      = $_."ESXi"
                "Cluster"   = $_."Cluster"
                "PortGroup" = $_."PortGroup"
                "Switch"    = $_."Switch"
                "vLAN"      = $_."vLAN"
            }
        }
        Export-AnnexSheet -Data $annexRows -Path $Context.AnnexPath -WorksheetName "VSS PortGroups"
        return [PSCustomObject]@{ ID = "VSP-TR-12"; Resultado = "No recomendado"; Detalle = "Configuracion de PortGroups no es consistente entre hosts del mismo cluster" }
    }

    return [PSCustomObject]@{ ID = "VSP-TR-12"; Resultado = "Resultado Esperado"; Detalle = "La configuracion de PortGroups es consistente entre los host del mismo cluster" }
}
