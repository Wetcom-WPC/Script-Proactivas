"%~dp0..\runtime\pwsh.exe" -command "Set-ExecutionPolicy -Scope CurrentUser Unrestricted"
@start "Paso 1: Recoleccion de Datos" /wait "%~dp0..\runtime\pwsh.exe" -ExecutionPolicy Bypass -Command "& '%~dp0devops-powershell\portable.ps1' -CarpetaResultados '%~1'"
