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

cd "`$APP_DIR"
[ -d venv ] || python3 -m venv venv
source venv/bin/activate
pip install -q -r requirements.txt flask-mail apscheduler xlsxwriter werkzeug 2>/dev/null || pip install -q -r requirements.txt

python -c "
from app import app, db, seed_feature_visibility, APP_VERSION
with app.app_context():
    db.create_all()
    seed_feature_visibility()
    print('Installed Lotus Task Manager v' + APP_VERSION)
"

sed -i 's/\r$//' update.sh 2>/dev/null || true
kill `$(lsof -t -i:5000) 2>/dev/null || true
sleep 2
nohup python app.py > app.log 2>&1 &
sleep 3
echo "--- Verify ---"
grep APP_VERSION app.py | head -1
grep ltm-blue templates/layout.html | head -1
lsof -i :5000 || true
echo "Open: http://187.124.15.14:5000/login (look for v badge + blue sidebar)"
"@

[System.IO.File]::WriteAllText($outPath, $script.Replace("`r`n", "`n"))
Write-Host "Created: $outPath ($((Get-Item $outPath).Length) bytes)"
