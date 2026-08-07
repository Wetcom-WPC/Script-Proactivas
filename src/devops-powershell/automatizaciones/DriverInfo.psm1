using module ".\lib\proactivas.psm1"

$credenciales = "cHJvYWN0aXZhOlBhc3N3b3JkMTIzJA=="

<#
.Synopsis
Recolección de Drivers de Host
.Component
vcenter
.Role
ui
#>
function Get-DriverInfo($vcenters) {
    Start-DriverCollection($vcenters)
}

function Start-DriverCollection($vcenters) {
    $proactiva = New-Object Proactiva

    foreach ($vcenter in $vcenters.conn) {
        $proactiva.setCurrentVcenter($vcenter)
        Write-Host "Processing vCenter: $vcenter"
        Write-Host "`tGathering Hosts..." -NoNewLine
        $hosts = Get-VMHost -server $vcenter
        Write-Host "`t`t$($hosts.length) Hosts found"

        # Llamamos ÚNICAMENTE a la función que procesa las NICs/Drivers
        $proactiva.processNic($hosts, $vdswitches)
    }

    $file = [PSCustomObject] @{
        Result         = "OK";
        Name           = "DriverInfo"; 
        Version        = $global:APP_VERSION;
        DateTime       = (Get-Date -Format "yyyy-MM-dd HH:mm");
        LocalHost      = [system.environment]::MachineName;
        User           = whoami;
        Endpoint       = $vcenters.host;
        Component      = "vcenter";
        Report         = $proactiva.getReport();
        IdAutomatizacion = $credenciales;
    }

    # --- [BLOQUE CORREGIDO] ---
    # 1. Se construye la ruta completa del archivo JSON y se guarda en una variable.
    $nombreJsonGenerado = (Get-Date).toString('yyyy-MM-dd HHmmss') + "_" + $file.Name + ".json"
    $jsonDir = Join-Path -Path $global:CONFIG.REPORTS_FOLDER -ChildPath "json"
    if (-not (Test-Path $jsonDir)) { New-Item -ItemType Directory -Path $jsonDir -Force | Out-Null }
    $rutaJsonGenerado = Join-Path -Path $jsonDir -ChildPath $nombreJsonGenerado

    # 2. Se guarda el archivo JSON usando esa ruta.
    $file | ConvertTo-Json -Depth 99 | Set-Content -Path $rutaJsonGenerado

    # 3. Se usa la variable global definida en portable.ps1 para escribir en resultado.txt.
    #    Añadimos el nombre del archivo que acabamos de crear a la lista.
    Add-Content -Path $global:RutaArchivoResultado -Value $nombreJsonGenerado
    # --- FIN DEL BLOQUE ---
}