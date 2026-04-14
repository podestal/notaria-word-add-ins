@echo off
REM Notaria Word Add-in Installer
REM Double-click this file to install the add-in

cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "Install-NotariaAddin.ps1"
pause
