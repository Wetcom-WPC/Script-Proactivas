#Requires -Version 5.1
# Diagnoses (and fixes where it is safe) the usual reasons PowerCLI does not load on a new
# PC, even though the project already bundles the modules in runtime\. It does not reinstall
# PowerCLI; it checks what PowerCLI needs AROUND it to load. Launched by extra\Preparar-Equipo.bat.

$repoRoot = Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$modulesFolder = Join-Path -Path $repoRoot -ChildPath "runtime\Modules"

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host " Diagnostico de requisitos para script-proactivas" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host ""

$foundProblems = $false

# -----------------------------------------------------------------------
# 1. Visual C++ Redistributable x64 (several native PowerCLI components need it)
# -----------------------------------------------------------------------
Write-Host "[1/4] Visual C++ Redistributable (x64)..." -NoNewline
$vcRegistryPath = "HKLM:\SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\X64"
$vcInstalled = (Test-Path $vcRegistryPath) -and ((Get-ItemProperty -Path $vcRegistryPath -ErrorAction SilentlyContinue).Installed -eq 1)

if ($vcInstalled) {
    Write-Host " OK" -ForegroundColor Green
} else {
    Write-Host " FALTA" -ForegroundColor Red
    Write-Host "        Varios componentes nativos que usa PowerCLI (ej. VMware.VimAutomation.Cis.Core)" -ForegroundColor Yellow
    Write-Host "        lo necesitan para cargar. No lo instalamos automaticamente porque requiere" -ForegroundColor Yellow
    Write-Host "        permisos de administrador y reiniciar en algunos casos. Descargalo de:" -ForegroundColor Yellow
    Write-Host "        https://aka.ms/vs/17/release/vc_redist.x64.exe" -ForegroundColor Yellow
    $foundProblems = $true
}

# -----------------------------------------------------------------------
# 2. NuGet package provider (PowerShellGet asks for it the first time a module is installed).
#    Safe to fix automatically.
# -----------------------------------------------------------------------
Write-Host "[2/4] Proveedor de paquetes NuGet..." -NoNewline
$nugetInstalled = Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue

if ($nugetInstalled) {
    Write-Host " OK" -ForegroundColor Green
} else {
    Write-Host " FALTA -> instalando..." -ForegroundColor Yellow
    try {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser -ErrorAction Stop | Out-Null
        Write-Host "        Instalado correctamente." -ForegroundColor Green
    } catch {
        Write-Host "        No se pudo instalar automaticamente: $($_.Exception.Message)" -ForegroundColor Red
        $foundProblems = $true
    }
}

# -----------------------------------------------------------------------
# 3. Another PowerCLI installation on the PC that could be confused with the bundled one
# -----------------------------------------------------------------------
Write-Host "[3/4] Instalaciones de PowerCLI fuera del proyecto..." -NoNewline
$otherPowerCli = Get-Module -ListAvailable -Name "VMware.PowerCLI", "VMware.VimAutomation.Core" -ErrorAction SilentlyContinue |
    Where-Object { $_.ModuleBase -notlike "$modulesFolder*" }

if (-not $otherPowerCli) {
    Write-Host " OK (ninguna)" -ForegroundColor Green
} else {
    Write-Host " ENCONTRADAS" -ForegroundColor Yellow
    Write-Host "        Esta PC tiene PowerCLI instalado fuera de la carpeta runtime\ del proyecto:" -ForegroundColor Yellow
    $otherPowerCli | Select-Object Name, Version, ModuleBase -Unique | ForEach-Object {
        Write-Host "        - $($_.Name) $($_.Version)  ($($_.ModuleBase))" -ForegroundColor Yellow
    }
    Write-Host "        El script ignora esas instalaciones (usa solo los modulos de runtime\Modules)." -ForegroundColor Yellow
    Write-Host "        Si aun asi aparece 'Assembly with same name is already loaded', renombra o" -ForegroundColor Yellow
    Write-Host "        desinstala esas carpetas VMware.* para descartar el conflicto de versiones." -ForegroundColor Yellow
}

# -----------------------------------------------------------------------
# 4. Actually import each bundled PowerCLI module, showing the real errors
# -----------------------------------------------------------------------
Write-Host "[4/4] Carga de los modulos de PowerCLI incluidos..."
$projectModules = @(
    "VMware.VimAutomation.Cis.Core",
    "VMware.VimAutomation.Common",
    "VMware.VimAutomation.Core",
    "VMware.VimAutomation.Vds",
    "VMware.VimAutomation.Sdk"
)

foreach ($moduleName in $projectModules) {
    Write-Host "        - $moduleName..." -NoNewline
    try {
        Import-Module (Join-Path $modulesFolder $moduleName) -ErrorAction Stop
        Write-Host " OK" -ForegroundColor Green
    } catch {
        Write-Host " FALLA" -ForegroundColor Red
        Write-Host "          $($_.Exception.Message)" -ForegroundColor Red
        $foundProblems = $true
    }
}

Write-Host ""
Write-Host "==========================================================" -ForegroundColor Cyan
if ($foundProblems) {
    Write-Host " Hay puntos para revisar (ver arriba en rojo/amarillo)." -ForegroundColor Yellow
    Write-Host " Si despues de instalar el Visual C++ Redistributable el problema sigue," -ForegroundColor Yellow
    Write-Host " copia el mensaje de error completo para que lo revisemos." -ForegroundColor Yellow
} else {
    Write-Host " Todo en orden. Esta PC deberia poder correr Iniciar-Proactiva.bat sin problemas." -ForegroundColor Green
}
Write-Host "==========================================================" -ForegroundColor Cyan
