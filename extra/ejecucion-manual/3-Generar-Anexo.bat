@echo off
setlocal
set "BASE_DIR=%~dp0"

:: Corre SOLO el Paso 3 (Anexo Tecnico + Checklist) contra el Excel "Proactiva"
:: mas reciente en resultados\. Util para regenerar el Anexo sin repetir los
:: pasos 1 y 2 (ej. corrigiendo el nombre del cliente o el mes).

"%BASE_DIR%..\..\runtime\pwsh.exe" -ExecutionPolicy Bypass -File "%BASE_DIR%..\..\src\proactivas-auto2.0.ps1"

echo.
echo Paso 3 (Anexo Tecnico y Checklist) completado.
pause
