# Notaria Word Add-in Installer
# This script installs the Notaria Word Add-in for the current user

Write-Host "================================================" -ForegroundColor Cyan
Write-Host "Notaria Word Add-in Installer" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

# Get the script directory
$scriptDir = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
$manifestFile = "manifest.xml"
$manifestPath = Join-Path $scriptDir $manifestFile

# Validate manifest exists
if (-not (Test-Path $manifestPath)) {
    Write-Host "ERROR: manifest.xml not found in $scriptDir" -ForegroundColor Red
    Write-Host ""
    Write-Host 'Press any key to close...' -ForegroundColor Gray
    $readKeyOptions = [System.Management.Automation.Host.ReadKeyOptions]::NoEcho -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown
    $null = $Host.UI.RawUI.ReadKey($readKeyOptions)
    exit 1
}

Write-Host "Step 1: Copying manifest to AppData..." -ForegroundColor Yellow
$appDataPath = "$env:APPDATA\Microsoft\Office\16.0\Wef"
if (-not (Test-Path $appDataPath)) {
    New-Item -Path $appDataPath -ItemType Directory -Force | Out-Null
}

$destManifest = Join-Path $appDataPath "NotariaAddin.xml"
Copy-Item -Path $manifestPath -Destination $destManifest -Force
Write-Host "  [OK] Manifest copied to: $destManifest" -ForegroundColor Green
Write-Host ""

Write-Host "Step 2: Configuring Office Registry..." -ForegroundColor Yellow

# Enable developer mode
$devPath = "HKCU:\Software\Microsoft\Office\16.0\Wef\Developer"
if (-not (Test-Path $devPath)) {
    New-Item -Path $devPath -Force | Out-Null
}
Set-ItemProperty -Path $devPath -Name "EnableWefDeveloperMode" -Value 1 -Type DWORD -Force
Write-Host '  [OK] Developer mode enabled' -ForegroundColor Green

# Add manifest to shared folder catalog
$catalogPath = "HKCU:\Software\Microsoft\Office\16.0\Wef\Catalog\SharedFolder"
if (-not (Test-Path $catalogPath)) {
    New-Item -Path $catalogPath -Force | Out-Null
}
$manifestFullPath = (Resolve-Path $manifestPath).Path
Set-ItemProperty -Path $catalogPath -Name "Path" -Value $appDataPath -Type String -Force
Write-Host '  [OK] Registry configured' -ForegroundColor Green
Write-Host ""

Write-Host "================================================" -ForegroundColor Green
Write-Host "Installation Complete!" -ForegroundColor Green
Write-Host "================================================" -ForegroundColor Green
Write-Host ""
Write-Host "Next step: Restart Microsoft Word" -ForegroundColor Yellow
Write-Host "The add-in will appear in the HOME tab" -ForegroundColor Yellow
Write-Host ""
Write-Host "Manifest installed to: $destManifest" -ForegroundColor Cyan
Write-Host ""

# Pause before closing
Write-Host 'Press any key to close...' -ForegroundColor Gray
$readKeyOptions = [System.Management.Automation.Host.ReadKeyOptions]::NoEcho -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown
$null = $Host.UI.RawUI.ReadKey($readKeyOptions)
