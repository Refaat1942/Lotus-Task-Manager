# Builds deploy\lotus-deploy.zip with the same folder layout as the repo.
# On the VPS:  unzip -o lotus-deploy.zip -d /root/taskmanager && sudo bash /root/taskmanager/deploy/update.sh
# (Easier: use deploy/deploy.sh, which pulls straight from GitHub.)
$Project = Split-Path $PSScriptRoot -Parent
Set-Location $Project
$out = Join-Path $PSScriptRoot "lotus-deploy.zip"
$stage = Join-Path $env:TEMP "lotus-deploy-stage"

if (Test-Path $stage) { Remove-Item $stage -Recurse -Force }
New-Item -ItemType Directory -Path $stage | Out-Null
foreach ($f in @("app.py", "i18n.py", "requirements.txt")) { Copy-Item $f $stage }
foreach ($d in @("templates", "static\images")) {
    New-Item -ItemType Directory -Path (Join-Path $stage $d) -Force | Out-Null
    Copy-Item "$d\*" (Join-Path $stage $d) -Recurse
}
New-Item -ItemType Directory -Path (Join-Path $stage "deploy") | Out-Null
foreach ($f in @("update.sh", "deploy.sh", "nginx-lotus.conf")) {
    Copy-Item "deploy\$f" (Join-Path $stage "deploy")
}

Compress-Archive -Path (Join-Path $stage "*") -DestinationPath $out -Force
Remove-Item $stage -Recurse -Force
Write-Host "Created: $out"
