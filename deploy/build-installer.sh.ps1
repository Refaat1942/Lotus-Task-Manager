$Project = Split-Path $PSScriptRoot -Parent
$zipPath = Join-Path $PSScriptRoot "lotus-deploy.zip"
$outPath = Join-Path $PSScriptRoot "install-on-vps.sh"

& (Join-Path $PSScriptRoot "make-zip.ps1")

$b64 = [Convert]::ToBase64String([IO.File]::ReadAllBytes($zipPath))

$script = @"
#!/bin/bash
# Lotus Task Manager - one-command install (no scp/zip upload needed)
# Paste this entire file on VPS, or: curl -sL URL | bash
set -e
APP_DIR="/root/taskmanager"
mkdir -p /root/backups "`$APP_DIR/templates"
[ -f "`$APP_DIR/instance/taskmanager.db" ] && cp -a "`$APP_DIR/instance/taskmanager.db" /root/backups/taskmanager.db.bak

echo "Extracting Lotus Task Manager files..."
echo '$b64' | base64 -d > /tmp/lotus-deploy.zip
unzip -o /tmp/lotus-deploy.zip -d "`$APP_DIR"
[ -f /root/backups/taskmanager.db.bak ] && cp /root/backups/taskmanager.db.bak "`$APP_DIR/instance/taskmanager.db"

sed -i 's/\r$//' "`$APP_DIR/deploy/update.sh" 2>/dev/null || true
bash "`$APP_DIR/deploy/update.sh"
"@

[System.IO.File]::WriteAllText($outPath, $script.Replace("`r`n", "`n"))
Write-Host "Created: $outPath ($((Get-Item $outPath).Length) bytes)"
