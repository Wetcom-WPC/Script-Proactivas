@echo off
setlocal
set "BASE_DIR=%~dp0"

:: Diagnostica (y corrige lo que es seguro corregir solo) los requisitos para que
:: PowerCLI cargue bien en esta PC. Correr esto antes de pedir ayuda por "no funciona".

"%BASE_DIR%runtime\pwsh.exe" -NoProfile -ExecutionPolicy Bypass -File "%BASE_DIR%src\Preparar-Equipo.ps1"

echo.
pause
