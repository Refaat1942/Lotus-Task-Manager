@echo off
setlocal EnableExtensions EnableDelayedExpansion

set "PROJECT=D:\Refaat\My Projects\Lotus-Task-Manager"
set "SERVER=root@187.124.15.14"
set "REMOTE=/root/taskmanager"
set "LOG=%PROJECT%\deploy\upload-log.txt"

:: Find scp.exe (double-click often has no PATH)
set "SCP="
if exist "%SystemRoot%\System32\OpenSSH\scp.exe" set "SCP=%SystemRoot%\System32\OpenSSH\scp.exe"
if not defined SCP if exist "%ProgramFiles%\Git\usr\bin\scp.exe" set "SCP=%ProgramFiles%\Git\usr\bin\scp.exe"
if not defined SCP if exist "%ProgramFiles(x86)%\Git\usr\bin\scp.exe" set "SCP=%ProgramFiles(x86)%\Git\usr\bin\scp.exe"
if not defined SCP for /f "delims=" %%i in ('where scp 2^>nul') do set "SCP=%%i"

cd /d "%PROJECT%" 2>nul
if errorlevel 1 (
    echo ERROR: Project folder not found:
    echo   %PROJECT%
    echo Edit PROJECT path at top of upload.bat if your folder is elsewhere.
    pause
    exit /b 1
)

if not exist "app.py" (
    echo ERROR: app.py not found in %CD%
    pause
    exit /b 1
)

if not defined SCP (
    echo ERROR: scp.exe not found.
    echo.
    echo Install one of these, then run this script again:
    echo   1. Windows OpenSSH: Settings ^> Apps ^> Optional Features ^> OpenSSH Client
    echo   2. Git for Windows: https://git-scm.com/download/win
    echo.
    echo Or use WinSCP / FileZilla to upload files manually.
    pause
    exit /b 1
)

echo ============================================ > "%LOG%"
echo Lotus Task Manager Upload >> "%LOG%"
echo Date: %date% %time% >> "%LOG%"
echo SCP: %SCP% >> "%LOG%"
echo Project: %CD% >> "%LOG%"
echo ============================================ >> "%LOG%"
echo.

echo Using: %SCP%
echo Uploading to %SERVER%:%REMOTE%
echo You will be asked for the VPS root password several times.
echo Log file: %LOG%
echo.

call :upload "app.py" "%REMOTE%/"
if errorlevel 1 goto :failed

call :upload "requirements.txt" "%REMOTE%/"
if errorlevel 1 goto :failed

call :upload "deploy\update.sh" "%REMOTE%/update.sh"
if errorlevel 1 goto :failed

echo Uploading templates...
for %%F in (templates\*.html) do (
    echo   %%F
    "%SCP%" "%%F" "%SERVER%:%REMOTE%/templates/" >> "%LOG%" 2>&1
    if errorlevel 1 goto :failed
)

echo.
echo ============================================
echo SUCCESS - Upload complete!
echo ============================================
echo.
echo Now on VPS run:
echo   sed -i 's/\r$//' /root/taskmanager/update.sh
echo   bash /root/taskmanager/update.sh
echo.
pause
exit /b 0

:upload
"%SCP%" "%~1" "%SERVER%%~2" >> "%LOG%" 2>&1
if errorlevel 1 (
    echo FAILED: %~1
    goto :eof
)
echo OK: %~1
exit /b 0

:failed
echo.
echo ============================================
echo UPLOAD FAILED - see details in:
echo   %LOG%
echo ============================================
type "%LOG%"
echo.
pause
exit /b 1
