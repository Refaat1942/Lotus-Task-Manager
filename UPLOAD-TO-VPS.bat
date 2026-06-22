@echo off
:: Drag-and-drop friendly: uploads from whatever folder this bat lives in (project root)
setlocal
cd /d "%~dp0.."
call "%~dp0upload.bat"
