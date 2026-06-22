@echo off
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0make-zip.ps1"
if exist "%~dp0lotus-deploy.zip" (
    echo.
    echo Upload deploy\lotus-deploy.zip via Hostinger File Manager to /root/
    echo Then on VPS:
    echo   mkdir -p /root/taskmanager/templates
    echo   unzip -o /root/lotus-deploy.zip -d /root/taskmanager/
    echo   sed -i 's/\r$//' /root/taskmanager/update.sh
    echo   bash /root/taskmanager/update.sh
    echo.
    echo Login page must show badge: v2.2.0
    explorer /select,"%~dp0lotus-deploy.zip"
) else (
    echo ZIP creation failed.
)
pause
