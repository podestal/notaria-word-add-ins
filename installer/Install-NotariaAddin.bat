@echo off
REM Notaria Word Add-in Installer
REM Double-click this file to install the add-in

cd /d "%~dp0"
echo [1/3] Enabling sideload policies...
powershell -NoProfile -ExecutionPolicy Bypass -File "Enable-NotariaSideloadPolicy.ps1"
if errorlevel 1 (
  echo.
  echo Policy setup failed. Installer aborted.
  pause
  exit /b 1
)

echo.
echo [2/3] Installing Notaria add-in...
powershell -NoProfile -ExecutionPolicy Bypass -File "Install-NotariaAddin.ps1"
if errorlevel 1 (
  echo.
  echo Add-in install failed. Installer aborted.
  pause
  exit /b 1
)

echo.
echo [3/3] Applying Office 2013 compatibility setup...
powershell -NoProfile -ExecutionPolicy Bypass -File "Enable-NotariaOffice2013Catalog.ps1"
if errorlevel 1 (
  echo.
  echo Office 2013 compatibility setup failed.
  echo Try running this .bat as Administrator.
  pause
  exit /b 1
)
pause
