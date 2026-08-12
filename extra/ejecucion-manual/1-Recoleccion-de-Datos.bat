@echo off
setlocal
set "BASE_DIR=%~dp0"

:: Corre SOLO el Paso 1 (recoleccion interactiva de datos de vSphere).
:: Util para recolectar sin generar Excel/Anexo todavia, o para volver a
:: recolectar sin repetir los pasos 2 y 3.

for /f "usebackq delims=" %%d in (`cmd /c ""%BASE_DIR%..\..\runtime\pwsh.exe" -NoProfile -ExecutionPolicy Bypass -File "%BASE_DIR%..\..\src\Resolver-CarpetaResultados.ps1" -BaseDir "%BASE_DIR%..\..\.""`) do set "CARPETA_RESULTADOS=%%d"

echo Resultados de esta corrida en: resultados\%CARPETA_RESULTADOS%
echo.

call "%BASE_DIR%..\..\src\run.bat" "%CARPETA_RESULTADOS%"

echo.
echo Paso 1 (Recoleccion de Datos) completado.
pause
