# Upload Lotus Task Manager to VPS
$Project = Split-Path $PSScriptRoot -Parent
if (-not (Test-Path "$Project\app.py")) {
    Write-Host "ERROR: app.py not found. Open PowerShell in the project folder." -ForegroundColor Red
    exit 1
}

$Server = "root@187.124.15.14"
$Remote = "/root/taskmanager"

Write-Host "Project: $Project" -ForegroundColor Gray
Write-Host "Uploading Lotus Task Manager to VPS..." -ForegroundColor Cyan

scp "$Project\app.py" "${Server}:${Remote}/"
scp "$Project\requirements.txt" "${Server}:${Remote}/"
scp -r "$Project\templates\*" "${Server}:${Remote}/templates/"
scp "$Project\deploy\update.sh" "${Server}:${Remote}/update.sh"

Write-Host ""
Write-Host "Upload done. On VPS run:" -ForegroundColor Green
Write-Host "  chmod +x /root/taskmanager/update.sh && /root/taskmanager/update.sh"
