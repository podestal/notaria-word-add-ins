param(
    [Parameter(Mandatory = $false)]
    [string]$CatalogLocalPath = "C:\OfficeAddins",

    [Parameter(Mandatory = $false)]
    [string]$ShareName = "OfficeAddins",

    [Parameter(Mandatory = $false)]
    [string]$ManifestFileName = "NotariaAddin.xml"
)

# Office 2013 uses a shared-folder catalog flow. This script creates
# a local SMB share and registers it as a trusted catalog for current user.

$office2013Detected =
    (Test-Path "HKCU:\Software\Microsoft\Office\15.0\Word") -or
    (Test-Path "HKLM:\Software\Microsoft\Office\15.0\Word") -or
    (Test-Path "HKLM:\Software\WOW6432Node\Microsoft\Office\15.0\Word")

if (-not $office2013Detected) {
    Write-Host "Office 2013 not detected. Skipping Office 2013 shared-catalog setup." -ForegroundColor Gray
    exit 0
}

Write-Host "Configuring Office 2013 shared-folder catalog..." -ForegroundColor Yellow

$scriptDir = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
$manifestSourcePath = Join-Path $scriptDir "manifest.xml"
$manifestDestPath = Join-Path $CatalogLocalPath $ManifestFileName
$catalogUrl = "\\localhost\$ShareName"
$catalogGuid = "{8B90BCDE-2E8E-4A70-8AE9-8F5AB8A9D1C1}"

if (-not (Test-Path $manifestSourcePath)) {
    Write-Host "ERROR: manifest.xml not found at $manifestSourcePath" -ForegroundColor Red
    exit 1
}

try {
    New-Item -Path $CatalogLocalPath -ItemType Directory -Force | Out-Null
    Copy-Item -Path $manifestSourcePath -Destination $manifestDestPath -Force
} catch {
    Write-Host "ERROR: Failed to create catalog folder or copy manifest." -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}

# Create SMB share using New-SmbShare when available.
$shareCreated = $false
$currentUser = "$env:USERDOMAIN\$env:USERNAME"
if (Get-Command -Name "Get-SmbShare" -ErrorAction SilentlyContinue) {
    try {
        $existingShare = Get-SmbShare -Name $ShareName -ErrorAction SilentlyContinue
        if (-not $existingShare) {
            # Avoid localized built-in account names such as "Everyone"/"Todos".
            New-SmbShare -Name $ShareName -Path $CatalogLocalPath -ReadAccess $currentUser -ErrorAction Stop | Out-Null
        }
        $shareCreated = $true
    } catch {
        $shareCreated = $false
    }
}

# Fallback for older systems.
if (-not $shareCreated) {
    $shareCmd = "net share $ShareName=`"$CatalogLocalPath`" /GRANT:`"$currentUser`",READ"
    cmd /c $shareCmd | Out-Null
    if ($LASTEXITCODE -ne 0 -and -not (Test-Path $catalogUrl)) {
        Write-Host "ERROR: Could not create SMB share '$ShareName'. Run installer as Administrator." -ForegroundColor Red
        exit 1
    }
}

$shareExists = $false
if (Get-Command -Name "Get-SmbShare" -ErrorAction SilentlyContinue) {
    $shareExists = [bool](Get-SmbShare -Name $ShareName -ErrorAction SilentlyContinue)
} else {
    $shareExists = Test-Path $catalogUrl
}
if (-not $shareExists) {
    Write-Host "ERROR: Share '$ShareName' was not created successfully." -ForegroundColor Red
    exit 1
}

try {
    $devKey = "HKCU:\Software\Microsoft\Office\15.0\WEF\Developer"
    $trustedCatalogsKey = "HKCU:\Software\Microsoft\Office\15.0\WEF\TrustedCatalogs"
    $catalogEntryKey = Join-Path $trustedCatalogsKey $catalogGuid

    New-Item -Path $devKey -Force | Out-Null
    New-Item -Path $catalogEntryKey -Force | Out-Null

    Set-ItemProperty -Path $devKey -Name "EnableWefDeveloperMode" -Value 1 -Type DWord
    Set-ItemProperty -Path $trustedCatalogsKey -Name "AllowNetworkCatalog" -Value 1 -Type DWord
    Set-ItemProperty -Path $catalogEntryKey -Name "Id" -Value $catalogGuid -Type String
    Set-ItemProperty -Path $catalogEntryKey -Name "Url" -Value $catalogUrl -Type String
    Set-ItemProperty -Path $catalogEntryKey -Name "Flags" -Value 1 -Type DWord
} catch {
    Write-Host "ERROR: Failed to write Office 2013 catalog registry keys." -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}

Write-Host "  [OK] Office 2013 catalog folder: $CatalogLocalPath" -ForegroundColor Green
Write-Host "  [OK] Office 2013 share: $catalogUrl" -ForegroundColor Green
Write-Host "  [OK] Trusted catalog registered for current user" -ForegroundColor Green
