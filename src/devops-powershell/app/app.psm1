$CURRENT_FOLDER = Split-Path $script:MyInvocation.MyCommand.Path
If (Test-Path "$CURRENT_FOLDER/config_dev.json") {
	$global:CONFIG = Get-Content $CURRENT_FOLDER/config_dev.json | ConvertFrom-Json
} else {
	$global:CONFIG = Get-Content $CURRENT_FOLDER/config.json | ConvertFrom-Json
}
$global:PLUGINS_MODULES = @()
$global:APP_VERSION = "v2.0.1ovw"

# Los resultados de cada corrida se agrupan en resultados/<carpeta de esta corrida>/ en la raiz
# del proyecto. $global:CARPETA_RESULTADOS lo define portable.ps1 (fecha, o fecha + " (N)" si ya
# se corrio hoy); si el modulo se carga sin pasar por portable.ps1, usamos la fecha de hoy.
if (-not $global:CARPETA_RESULTADOS) {
    $global:CARPETA_RESULTADOS = Get-Date -Format "yyyy-MM-dd"
}
$global:CONFIG.REPORTS_FOLDER = Join-Path -Path $global:CONFIG.REPORTS_FOLDER -ChildPath $global:CARPETA_RESULTADOS
if (-not (Test-Path $global:CONFIG.REPORTS_FOLDER)) {
    New-Item -ItemType Directory -Path $global:CONFIG.REPORTS_FOLDER -Force | Out-Null
}

Import-Module $CURRENT_FOLDER/ui.psm1
Import-Module $CURRENT_FOLDER/plugins.psm1
Import-Module $CURRENT_FOLDER/connections.psm1



function Start-App() {
    Start-Plugins
    Set-Endpoints($global:PLUGINS_MODULES)
    Register-ExecuteEventSubscription('Start-Tasks')
    Register-ClearConnectionsEventSubscription('Disconnect-Endpoints') # Method from connections.psm1
    Start-UI
}

function Start-Tasks {
	$modules = $global:PLUGINS_MODULES | Where-Object {$_.checked -eq $true}
    if(Connect-Endpoints $modules){
		$global:PLUGINS_MODULES | Show-Menu

		foreach ($module in $modules) {
			$params = $global:connections | Where-Object {$_.component -in $module.COMPONENT.Split(";") }
			
			Write-Title($module.Name)
			&($module.Name) $params
		}
	}else{
		Read-Host "Could not connect to all endpoints! Press Enter... "
		exit 1	
	}
}


#function Start-App-NoUI() {
#    # 1. Carga los plugins disponibles (igual que la función original)
#    Start-Plugins
#
#    # 2. Pide los endpoints (servidores y credenciales). Esta parte sigue siendo interactiva.
#    Set-Endpoints($global:PLUGINS_MODULES)
#
#    # 3. Selecciona automáticamente todos los plugins que se encontraron.
#    $global:PLUGINS_MODULES | ForEach-Object { $_.checked = $true }
#
#    Write-Host "`n[MODO AUTOMÁTICO] Todos los módulos seleccionados. Iniciando recolección..." -ForegroundColor Cyan
#    
#    # 4. Ejecuta directamente la recolección de datos, salteando el menú.
#    Start-Tasks
#
#    # 5. Se desconecta de los servidores al finalizar.
#    Disconnect-Endpoints
#
#    Write-Host "[MODO AUTOMÁTICO] Recolección de datos finalizada." -ForegroundColor Green
#}