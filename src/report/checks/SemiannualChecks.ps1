# Semiannual checks (VSP-SE-*). Dot-sourced by Report.psm1. Same contract as MonthlyChecks.ps1.

# The collection exports every vDS configuration to <run folder>\vds_configuration (on every
# visit). This check confirms those exports worked; the .zip files must then be uploaded to Drive.
function Test-VdsBackup($Context) {
    # No vDS in the environment means no "vDS" sheet in the raw Excel
    $vdSwitches = @(Import-RawSheet $Context.RawExcelPath "vDS" | Where-Object { -not [string]::IsNullOrEmpty($_.vCenter) })
    $backupFolder = Join-Path $Context.RunFolder "vds_configuration"

    if ($vdSwitches.Count -eq 0) {
        return [PSCustomObject]@{ ID = "VSP-SE-01"; Resultado = "Resultado Esperado"; Detalle = "No se detectaron Distributed Switches (vDS) en la infraestructura para respaldar." }
    }

    $findings = @()
    $uniqueErrors = @()

    foreach ($row in $vdSwitches) {
        $backupFile = $row."Backup File"

        # The collection writes "ERROR: <message>" instead of the file name when the export fails
        if ($backupFile -like "ERROR*") {
            $errorMessage = $backupFile -replace "ERROR: ", ""
            if ($errorMessage -notin $uniqueErrors) { $uniqueErrors += $errorMessage }

            $findings += [PSCustomObject]@{
                "vCenter"  = $row.vCenter
                "vDS Name" = $row.Name
                "Estado"   = "Fallo de Backup"
                "Detalle"  = $errorMessage
            }
        }
    }

    if ($findings.Count -gt 0) {
        Export-AnnexSheet -Data $findings -Path $Context.AnnexPath -WorksheetName "Backup vDS"

        $errorSummary = if ($uniqueErrors.Count -gt 1) { "Múltiples errores (ver anexo)" } else { $uniqueErrors[0] }
        return [PSCustomObject]@{ ID = "VSP-SE-01"; Resultado = "No recomendado"; Detalle = "Falló el backup de $($findings.Count) vDS. Error: $errorSummary" }
    }

    return [PSCustomObject]@{
        ID        = "VSP-SE-01"
        Resultado = "Resultado Esperado"
        Detalle   = "Se realizaron los backups correctamente. Recuerde subir los archivos. Chequear path: $backupFolder"
    }
}

# VAMI file-based backups: they must be scheduled, succeed, and run every day (gaps > 26h)
function Test-VamiBackup($Context) {
    $jobs = @(Import-RawSheet $Context.RawExcelPath "BackupActivity" | Where-Object { -not [string]::IsNullOrEmpty($_.vCenter) })

    if ($jobs.Count -eq 0) {
        return [PSCustomObject]@{ ID = "VSP-SE-02"; Resultado = "Sin acceso"; Detalle = "No se encontraron registros de actividad de backup." }
    }

    $findings = @()
    $foundFailed = $false
    $foundManual = $false
    $foundGaps = $false
    $affectedVCenters = @()

    foreach ($vcenterGroup in ($jobs | Group-Object vCenter)) {
        $vcenterName = $vcenterGroup.Name
        # Newest first
        $backups = $vcenterGroup.Group | Sort-Object StartTime -Descending
        $hasProblems = $false

        # A. Status and type
        foreach ($backup in $backups) {
            if ($backup.Status -ne "SUCCEEDED") { $hasProblems = $true; $foundFailed = $true }
            if ($backup.Type -ne "SCHEDULED")   { $hasProblems = $true; $foundManual = $true }
        }

        # B. Frequency
        $parsedBackups = @()
        foreach ($backup in $backups) {
            $startDate = $null
            try {
                if ($backup.StartTime -and $backup.StartTime -ne "N/A") {
                    $startDate = Get-Date $backup.StartTime -ErrorAction Stop
                }
            } catch {}

            $parsedBackups += [PSCustomObject]@{
                Original = $backup
                Date     = $startDate
                Note     = ""
            }
        }

        # B1. Gap between now and the latest backup
        if ($parsedBackups.Count -gt 0 -and $parsedBackups[0].Date) {
            $hoursSinceLatest = ((Get-Date) - $parsedBackups[0].Date).TotalHours
            if ($hoursSinceLatest -gt 26) {
                $hasProblems = $true
                $foundGaps = $true
                $parsedBackups[0].Note = "Salto > 24h desde hoy ($([math]::Round($hoursSinceLatest))hs)"
            }
        } else {
            # No valid dates at all
            $hasProblems = $true
        }

        # B2. Gaps between consecutive backups (more than 26h means a missing day)
        for ($i = 0; $i -lt ($parsedBackups.Count - 1); $i++) {
            $current = $parsedBackups[$i]
            $previous = $parsedBackups[$i + 1]

            if ($current.Date -and $previous.Date) {
                $gapHours = ($current.Date - $previous.Date).TotalHours
                if ($gapHours -gt 26) {
                    $hasProblems = $true
                    $foundGaps = $true
                    # Mark the older backup of the pair: the gap comes after it
                    $previous.Note = "Salto > 24h hasta el siguiente backup ($([math]::Round($gapHours))hs)"
                }
            }
        }

        # C. When a vCenter has any problem, its whole history goes to the annex for context
        if ($hasProblems) {
            if ($vcenterName -notin $affectedVCenters) { $affectedVCenters += $vcenterName }

            foreach ($item in $parsedBackups) {
                $backup = $item.Original

                if ($backup.Status -ne "SUCCEEDED") { $evaluation = "Estado Incorrecto ($($backup.Status))" }
                elseif ($backup.Type -ne "SCHEDULED") { $evaluation = "Tipo Incorrecto ($($backup.Type))" }
                elseif ($item.Note) { $evaluation = "Error Frecuencia: $($item.Note)" }
                else { $evaluation = "OK (Contexto)" }

                $findings += [PSCustomObject]@{
                    "vCenter"       = $vcenterName
                    "JobId"         = $backup.JobId
                    "Type"          = $backup.Type
                    "Status"        = $backup.Status
                    "StartTime"     = $backup.StartTime
                    "Location"      = $backup.Location
                    "Evaluación"    = $evaluation
                }
            }
        }
    }

    $result = "Resultado Esperado"
    $detail = "Los backups se ejecutan diariamente, son programados y finalizan exitosamente."

    $problems = @()
    if ($foundFailed) { $problems += "fallos de ejecución" }
    if ($foundManual) { $problems += "ejecuciones manuales" }
    if ($foundGaps)   { $problems += "interrupciones en la frecuencia diaria" }

    if ($problems.Count -gt 0) {
        $result = "No recomendado"
        $detail = "Se detectaron: " + ($problems -join ", ") + ". Afecta a: " + ($affectedVCenters -join ", ")
    }

    if ($findings.Count -gt 0) {
        Export-AnnexSheet -Data $findings -Path $Context.AnnexPath -WorksheetName "Backup VAMI"
    }

    return [PSCustomObject]@{ ID = "VSP-SE-02"; Resultado = $result; Detalle = $detail }
}
