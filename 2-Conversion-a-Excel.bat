@echo off
setlocal
set "BASE_DIR=%~dp0"

:: Corre SOLO el Paso 2 (convierte a Excel el/los JSON listados en resultado.txt).
:: Util para reprocesar sin volver a conectarse al vCenter.

"%BASE_DIR%runtime\pwsh.exe" -ExecutionPolicy Bypass -File "%BASE_DIR%src\JSONtoExcels.ps1"

echo.
echo Paso 2 (Conversion a Excel) completado.
pause
