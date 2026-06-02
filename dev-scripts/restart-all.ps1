$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$backendScript = Join-Path $scriptDir "restart-backend.ps1"
$frontendScript = Join-Path $scriptDir "restart-frontend.ps1"

Write-Host "Opening backend and frontend in separate PowerShell windows..." -ForegroundColor Cyan
Start-Process powershell -ArgumentList @("-NoExit", "-ExecutionPolicy", "Bypass", "-File", $backendScript)
Start-Process powershell -ArgumentList @("-NoExit", "-ExecutionPolicy", "Bypass", "-File", $frontendScript)
Write-Host "Done. Close those two windows when you want to stop the servers." -ForegroundColor Green
