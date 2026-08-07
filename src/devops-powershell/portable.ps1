param(
    [string]$CarpetaResultados = (Get-Date -Format "yyyy-MM-dd")
)

$directorioSrc = Split-Path -Path $PSScriptRoot -Parent
$directorioRaiz = Split-Path -Path $directorioSrc -Parent

$global:CARPETA_RESULTADOS = $CarpetaResultados

$logDir = Join-Path -Path $directorioRaiz -ChildPath "resultados\$CarpetaResultados\logs"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
Start-Transcript -Path (Join-Path -Path $logDir -ChildPath ("{0}_portable.log" -f (Get-Date -Format "yyyy-MM-dd_HHmmss"))) | Out-Null

$global:RutaArchivoResultado = Join-Path -Path $directorioRaiz -ChildPath "resultado.txt"

if (Test-Path $global:RutaArchivoResultado) {
    Remove-Item $global:RutaArchivoResultado
}


Push-Location (Split-Path -Path $MyInvocation.MyCommand.Definition -Parent)

Import-Module ..\..\runtime\Modules\VMware.VimAutomation.Cis.Core
Import-Module ..\..\runtime\Modules\VMware.VimAutomation.Common
Import-Module ..\..\runtime\Modules\VMware.VimAutomation.Core
Import-Module ..\..\runtime\Modules\VMware.VimAutomation.Vds
Import-Module ..\..\runtime\Modules\VMware.VimAutomation.Sdk
Import-Module ..\..\runtime\Modules\VMware.Sdk.vSphere
Set-PowerCLIConfiguration -Scope User -ParticipateInCEIP $false -confirm:$false
Set-PowerCLIConfiguration -InvalidCertificateAction:Ignore -confirm:$false

Import-Module ./app/app.psm1
Start-App

Stop-Transcript | Out-Null

Pop-Location