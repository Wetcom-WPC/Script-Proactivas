@echo off
:: Vuelve a generar el Anexo Tecnico y la Checklist de la ultima corrida (sin reconectarse al vCenter).
:: Util para corregir el cliente, el mes o el tipo de proactiva.
"%~dp0..\runtime\pwsh.exe" -NoProfile -ExecutionPolicy Bypass -File "%~dp0..\src\Start-Proactiva.ps1" -ReportOnly %*
pause
