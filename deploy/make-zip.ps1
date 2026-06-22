$Project = Split-Path $PSScriptRoot -Parent
Set-Location $Project
$out = Join-Path $PSScriptRoot "lotus-deploy.zip"

$items = @(
    "app.py",
    "requirements.txt",
    "deploy\update.sh"
) + (Get-ChildItem "templates\*.html" | ForEach-Object { $_.FullName })

Compress-Archive -Path $items -DestinationPath $out -Force
Write-Host "Created: $out"
