@echo off
:: Genera el Anexo Tecnico y la Checklist de la Proactiva. Todo el proceso esta en src\Start-Proactiva.ps1;
:: este archivo solo existe para poder arrancarlo con doble clic usando el PowerShell incluido en runtime\.
"%~dp0runtime\pwsh.exe" -NoProfile -ExecutionPolicy Bypass -File "%~dp0src\Start-Proactiva.ps1" %*
pause
