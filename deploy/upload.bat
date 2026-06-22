@echo off
cd /d "D:\Refaat\My Projects\Lotus-Task-Manager"
if not exist app.py (
    echo ERROR: app.py not found in project folder.
    pause
    exit /b 1
)

echo Uploading Lotus Task Manager to VPS...
echo.

scp app.py root@187.124.15.14:/root/taskmanager/
scp requirements.txt root@187.124.15.14:/root/taskmanager/
scp -r templates\* root@187.124.15.14:/root/taskmanager/templates/
scp deploy\update.sh root@187.124.15.14:/root/taskmanager/update.sh

echo.
echo Upload done.
echo Fixing line endings on VPS is done automatically if you run update via bash:
echo   sed -i 's/\r$//' /root/taskmanager/update.sh
echo   bash /root/taskmanager/update.sh
echo.
echo Or on VPS run these commands directly:
echo   cd /root/taskmanager ^&^& source venv/bin/activate
echo   pip install -r requirements.txt -q
echo   python -c "from app import app, db, seed_feature_visibility, APP_VERSION; app.app_context().push(); db.create_all(); seed_feature_visibility(); print('OK v'+APP_VERSION)"
echo   kill $(lsof -t -i:5000); sleep 2; nohup python app.py ^> app.log 2^>^&1 ^&
echo.
pause
