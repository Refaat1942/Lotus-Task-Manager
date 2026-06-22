# Lotus Task Manager - VPS upload (PowerShell)
$ErrorActionPreference = "Stop"
$Project = Split-Path $PSScriptRoot -Parent
$Server = "root@187.124.15.14"
$Remote = "/root/taskmanager"

$Scp = @(
    "$env:SystemRoot\System32\OpenSSH\scp.exe",
    "${env:ProgramFiles}\Git\usr\bin\scp.exe",
    "${env:ProgramFiles(x86)}\Git\usr\bin\scp.exe"
) | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $Scp) {
    $cmd = Get-Command scp -ErrorAction SilentlyContinue
    if ($cmd) { $Scp = $cmd.Source }
}

if (-not $Scp) {
    Write-Host "ERROR: scp not found. Install OpenSSH Client or Git for Windows." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

Set-Location $Project
Write-Host "SCP: $Scp" -ForegroundColor Gray
Write-Host "Uploading from: $Project" -ForegroundColor Cyan
Write-Host "Enter VPS password when prompted..." -ForegroundColor Yellow

& $Scp "app.py" "${Server}:${Remote}/"
& $Scp "requirements.txt" "${Server}:${Remote}/"
& $Scp "deploy\update.sh" "${Server}:${Remote}/update.sh"
Get-ChildItem "templates\*.html" | ForEach-Object {
    Write-Host "  $($_.Name)" -ForegroundColor Gray
    & $Scp $_.FullName "${Server}:${Remote}/templates/"
}

Write-Host ""
Write-Host "Done! On VPS run:" -ForegroundColor Green
Write-Host "  sed -i 's/\r$//' /root/taskmanager/update.sh"
Write-Host "  bash /root/taskmanager/update.sh"
Read-Host "Press Enter to exit"
