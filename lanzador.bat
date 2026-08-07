@echo off
setlocal
set "BASE_DIR=%~dp0"

:: --------------------------------------------------------------------------------
:: PASO 1: Recolección
:: --------------------------------------------------------------------------------
set "RESULT_FILE=%BASE_DIR%resultado.txt"
IF EXIST "%RESULT_FILE%" (
    echo Limpiando archivo de resultado anterior...
    del "%RESULT_FILE%"
)

echo [PASO 1 de 3] - Ejecutando la recoleccion de datos de vSphere...
echo (Se abrira una nueva ventana para este paso)
echo.

call "%BASE_DIR%src\run.bat"

:: Aquí usamos la sintaxis antigua pero segura para comprobar si run.bat falló
IF ERRORLEVEL 1 (
    echo.
    echo El usuario cancelo la operacion. El flujo de trabajo se ha detenido.
    GOTO :end_script
)

:: --------------------------------------------------------------------------------
:: PASO 2: Conversión a Excel
:: --------------------------------------------------------------------------------
echo.
echo [PASO 2 de 3] - Recoleccion finalizada. Ejecutando la conversion a Excel...
echo (Se abrira una nueva ventana para este paso)
echo.

start "Paso 2: Conversion a Excel" /wait "%BASE_DIR%runtime\pwsh.exe" -ExecutionPolicy Bypass -File "%BASE_DIR%src\JSONtoExcels.ps1" 2>nul

IF NOT EXIST "%RESULT_FILE%" (
    echo.
    echo ADVERTENCIA: No se encontró el archivo de resultado. Omitiendo el paso 3.
    GOTO :end_script
)

:: --------------------------------------------------------------------------------
:: PASO 3: Validación y Checklist (MÉTODO BLINDADO)
:: --------------------------------------------------------------------------------
echo.
echo Buscando trigger "_Proactiva" en: "%RESULT_FILE%"

:: EXPLICACIÓN TÉCNICA:
:: findstr devuelve "Exit Code 0" si encuentra el texto, y "1" si no.
:: El operador '&&' ejecuta el bloque siguiente SOLO si el comando anterior dio 0 (éxito).
:: El operador '||' ejecuta el bloque siguiente SOLO si el comando anterior dio 1 (fallo).
:: Esto evita leer la variable %ERRORLEVEL% que podría estar sucia.

findstr /M /I /C:"_Proactiva" "%RESULT_FILE%" >nul && (
    echo.
    echo [PASO 3 de 3] - REPORTE PROACTIVA DETECTADO.
    echo Ejecutando generacion de Anexo y Checklist...
    echo.
    start "Paso 3: Generando Anexo" /wait "%BASE_DIR%runtime\pwsh.exe" -ExecutionPolicy Bypass -File "%BASE_DIR%src\proactivas-auto2.0.ps1" 2>nul
) || (
    echo.
    echo [INFO] No se encontro la marca "_Proactiva" en el resultado.
    echo Se omite la generacion del Anexo tecnico y Checklist.
)

:end_script
echo.
echo Proceso completado.
pause