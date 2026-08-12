#Requires -Version 5.1
# Diagnostica y (donde es seguro) corrige los motivos mas comunes por los que PowerCLI
# no carga en una PC nueva, aunque el proyecto ya trae los modulos incluidos en runtime\.
# No reinstala PowerCLI (ya viene bundleado) -- se enfoca en lo que falta ALREDEDOR
# de PowerCLI para que esos modulos carguen bien.

$baseDir = Split-Path -Path $PSScriptRoot -Parent
$modulesDir = Join-Path -Path $baseDir -ChildPath "runtime\Modules"

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host " Diagnostico de requisitos para script-proactivas" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host ""

$huboProblemas = $false

# -----------------------------------------------------------------------
# 1. Visual C++ Redistributable x64 (varios componentes nativos de PowerCLI lo necesitan)
# -----------------------------------------------------------------------
Write-Host "[1/4] Visual C++ Redistributable (x64)..." -NoNewline
$vcRuta = "HKLM:\SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\X64"
$vcInstalado = (Test-Path $vcRuta) -and ((Get-ItemProperty -Path $vcRuta -ErrorAction SilentlyContinue).Installed -eq 1)

if ($vcInstalado) {
    Write-Host " OK" -ForegroundColor Green
} else {
    Write-Host " FALTA" -ForegroundColor Red
    Write-Host "        Varios componentes nativos que usa PowerCLI (ej. VMware.VimAutomation.Cis.Core)" -ForegroundColor Yellow
    Write-Host "        lo necesitan para cargar. No lo instalamos automaticamente porque requiere" -ForegroundColor Yellow
    Write-Host "        permisos de administrador y reiniciar en algunos casos. Descargalo de:" -ForegroundColor Yellow
    Write-Host "        https://aka.ms/vs/17/release/vc_redist.x64.exe" -ForegroundColor Yellow
    $huboProblemas = $true
}

# -----------------------------------------------------------------------
# 2. Proveedor NuGet (lo pide PowerShellGet la primera vez que se instala un modulo,
#    ej. ImportExcel en el Paso 3 si no esta ya instalado). Esto SI es seguro de resolver solo.
# -----------------------------------------------------------------------
Write-Host "[2/4] Proveedor de paquetes NuGet..." -NoNewline
$nugetInstalado = Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue

if ($nugetInstalado) {
    Write-Host " OK" -ForegroundColor Green
} else {
    Write-Host " FALTA -> instalando..." -ForegroundColor Yellow
    try {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser -ErrorAction Stop | Out-Null
        Write-Host "        Instalado correctamente." -ForegroundColor Green
    } catch {
        Write-Host "        No se pudo instalar automaticamente: $($_.Exception.Message)" -ForegroundColor Red
        $huboProblemas = $true
    }
}

# -----------------------------------------------------------------------
# 3. Otra instalacion de PowerCLI en la PC que pueda pisar/confundirse con la incluida
# -----------------------------------------------------------------------
Write-Host "[3/4] Instalaciones de PowerCLI fuera del proyecto..." -NoNewline
$otrasPowerCLI = Get-Module -ListAvailable -Name "VMware.PowerCLI", "VMware.VimAutomation.Core" -ErrorAction SilentlyContinue |
    Where-Object { $_.ModuleBase -notlike "$modulesDir*" }

if (-not $otrasPowerCLI) {
    Write-Host " OK (ninguna)" -ForegroundColor Green
} else {
    Write-Host " ENCONTRADAS" -ForegroundColor Yellow
    Write-Host "        Esta PC tiene PowerCLI instalado fuera de la carpeta runtime\ del proyecto:" -ForegroundColor Yellow
    $otrasPowerCLI | Select-Object Name, Version, ModuleBase -Unique | ForEach-Object {
        Write-Host "        - $($_.Name) $($_.Version)  ($($_.ModuleBase))" -ForegroundColor Yellow
    }
    Write-Host "        No debería causar problemas (el proyecto carga sus modulos por ruta completa," -ForegroundColor Yellow
    Write-Host "        no por nombre), pero si algo falla mas abajo, puede ser un conflicto de versiones." -ForegroundColor Yellow
}

# -----------------------------------------------------------------------
# 4. Import real de cada modulo de PowerCLI que usa el proyecto, sin ocultar errores
#    (a diferencia de la corrida normal, que los silencia para no ensuciar la consola).
# -----------------------------------------------------------------------
Write-Host "[4/4] Carga de los modulos de PowerCLI incluidos..."
$modulosProyecto = @(
    "VMware.VimAutomation.Cis.Core",
    "VMware.VimAutomation.Common",
    "VMware.VimAutomation.Core",
    "VMware.VimAutomation.Vds",
    "VMware.VimAutomation.Sdk"
)

foreach ($nombreModulo in $modulosProyecto) {
    Write-Host "        - $nombreModulo..." -NoNewline
    try {
        Import-Module (Join-Path $modulesDir $nombreModulo) -ErrorAction Stop
        Write-Host " OK" -ForegroundColor Green
    } catch {
        Write-Host " FALLA" -ForegroundColor Red
        Write-Host "          $($_.Exception.Message)" -ForegroundColor Red
        $huboProblemas = $true
    }
}

Write-Host ""
Write-Host "==========================================================" -ForegroundColor Cyan
if ($huboProblemas) {
    Write-Host " Hay puntos para revisar (ver arriba en rojo/amarillo)." -ForegroundColor Yellow
    Write-Host " Si despues de instalar el Visual C++ Redistributable el problema sigue," -ForegroundColor Yellow
    Write-Host " copia el mensaje de error completo para que lo revisemos." -ForegroundColor Yellow
} else {
    Write-Host " Todo en orden. Esta PC deberia poder correr lanzador.bat sin problemas." -ForegroundColor Green
}
Write-Host "==========================================================" -ForegroundColor Cyan
