param(
    [Parameter(Mandatory = $false)]
    [string]$CatalogPath = "$env:APPDATA\Microsoft\Office\16.0\Wef",

    [Parameter(Mandatory = $false)]
    [string]$ManifestFileName = "NotariaAddin.xml"
)

# Enables the minimum registry settings required to sideload Word add-ins
# from a shared-folder style catalog for the current user.

Write-Host "================================================" -ForegroundColor Cyan
Write-Host "Notaria Sideload Policy Setup (Current User)" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

if (-not (Test-Path $CatalogPath)) {
    Write-Host "ERROR: Catalog path does not exist: $CatalogPath" -ForegroundColor Red
    exit 1
}

$manifestPath = Join-Path $CatalogPath $ManifestFileName
if (-not (Test-Path $manifestPath)) {
    Write-Host "WARNING: Manifest not found at: $manifestPath" -ForegroundColor Yellow
    Write-Host "The policy setup can continue, but Word will not show the add-in until the manifest exists in the catalog path." -ForegroundColor Yellow
    Write-Host ""
}

$catalogGuid = "{8B90BCDE-2E8E-4A70-8AE9-8F5AB8A9D1C1}"
$officeVersions = @("16.0", "15.0")

foreach ($version in $officeVersions) {
    $wefRoot = "HKCU:\Software\Microsoft\Office\$version\WEF"
    $developerKey = Join-Path $wefRoot "Developer"
    $trustedCatalogsKey = Join-Path $wefRoot "TrustedCatalogs"
    $catalogEntryKey = Join-Path $trustedCatalogsKey $catalogGuid

    if (-not (Test-Path $wefRoot)) {
        New-Item -Path $wefRoot -Force | Out-Null
    }
    if (-not (Test-Path $developerKey)) {
        New-Item -Path $developerKey -Force | Out-Null
    }
    if (-not (Test-Path $trustedCatalogsKey)) {
        New-Item -Path $trustedCatalogsKey -Force | Out-Null
    }
    if (-not (Test-Path $catalogEntryKey)) {
        New-Item -Path $catalogEntryKey -Force | Out-Null
    }

    # Enable Office add-in developer sideload mode.
    Set-ItemProperty -Path $developerKey -Name "EnableWefDeveloperMode" -Value 1 -Type DWord

    # Allow trusted shared-folder catalogs for this user.
    Set-ItemProperty -Path $trustedCatalogsKey -Name "AllowNetworkCatalog" -Value 1 -Type DWord

    # Register this catalog path as a trusted catalog.
    Set-ItemProperty -Path $catalogEntryKey -Name "Id" -Value $catalogGuid -Type String
    Set-ItemProperty -Path $catalogEntryKey -Name "Url" -Value $CatalogPath -Type String
    Set-ItemProperty -Path $catalogEntryKey -Name "Flags" -Value 1 -Type DWord
}

Write-Host "  [OK] Developer sideload mode enabled (Office 15/16)" -ForegroundColor Green
Write-Host "  [OK] Trusted shared-folder catalog enabled" -ForegroundColor Green
Write-Host "  [OK] Catalog registered: $CatalogPath" -ForegroundColor Green
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Yellow
Write-Host "1) Ensure the manifest exists here: $manifestPath" -ForegroundColor Yellow
Write-Host "2) Close all Word windows completely." -ForegroundColor Yellow
Write-Host "3) Open Word and go to Insert > My Add-ins > Shared Folder." -ForegroundColor Yellow
Write-Host "4) Add the Notaria add-in." -ForegroundColor Yellow
Write-Host ""

$policyRootCU = "HKCU:\Software\Policies\Microsoft\Office\16.0\WEF\TrustedCatalogs"
$policyRootLM = "HKLM:\Software\Policies\Microsoft\Office\16.0\WEF\TrustedCatalogs"
if ((Test-Path $policyRootCU) -or (Test-Path $policyRootLM)) {
    Write-Host "NOTE: Group Policy keys were detected under ...\Policies\...\WEF\TrustedCatalogs." -ForegroundColor Yellow
    Write-Host "If Word still hides Shared Folder add-ins, a domain/local GPO is likely overriding user settings." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Done." -ForegroundColor Cyan
