# install.ps1 — Installs the VSME globalMOO Excel Add-in
# Run once to install; re-run anytime to update the manifest.
# No admin rights required. Code updates from GitHub are automatic.

$ErrorActionPreference  = "Stop"
$ProgressPreference     = "SilentlyContinue"

$manifestUrl  = "https://globalmoo.github.io/gmoo-excel-plugin/manifest.xml"
$addinDir     = "$env:LOCALAPPDATA\globalMOO\vsme-addin"
$manifestPath = "$addinDir\manifest.xml"

# Stable GUID matching the add-in ID in manifest.xml — keeps re-runs idempotent
$catalogId = "{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}"

Write-Host ""
Write-Host "Installing VSME - globalMOO Excel Add-in" -ForegroundColor Cyan
Write-Host "-----------------------------------------"

# 1. Create local folder
if (-not (Test-Path $addinDir)) {
    New-Item -ItemType Directory -Path $addinDir -Force | Out-Null
}

# 2. Download manifest from GitHub Pages
Write-Host "  Downloading manifest..."
try {
    Invoke-WebRequest -Uri $manifestUrl -OutFile $manifestPath -UseBasicParsing
    Write-Host "  Manifest saved to $manifestPath" -ForegroundColor Green
} catch {
    Write-Host "  ERROR: Could not download manifest. Check your internet connection." -ForegroundColor Red
    Write-Host "  $_"
    Read-Host "Press Enter to exit"
    exit 1
}

# 3. Register local folder as a trusted add-in catalog
$catalogsKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs"
if (-not (Test-Path $catalogsKey)) {
    New-Item -Path $catalogsKey -Force | Out-Null
}
$catalogKey = "$catalogsKey\$catalogId"
if (-not (Test-Path $catalogKey)) {
    New-Item -Path $catalogKey -Force | Out-Null
}
Set-ItemProperty -Path $catalogKey -Name "Id"    -Value $catalogId
Set-ItemProperty -Path $catalogKey -Name "Url"   -Value $addinDir
Set-ItemProperty -Path $catalogKey -Name "Flags" -Value 1 -Type DWord

Write-Host "  Trusted catalog registered." -ForegroundColor Green
Write-Host ""
Write-Host "Installation complete!" -ForegroundColor Green
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Yellow
Write-Host "  1. Restart Microsoft Excel"
Write-Host "  2. Home tab -> Add-ins -> More Add-ins -> Shared Folder tab"
Write-Host "  3. Select VSME and click Add"
Write-Host ""
Write-Host "Future code updates deploy automatically - no reinstall needed."
Write-Host "Re-run this script only if you are told a new version requires it."
Write-Host ""
Read-Host "Press Enter to close"
