# install.ps1 — Installs the VSME globalMOO Excel Add-in
# Run once to install; re-run anytime to update the manifest.
# No admin rights required. Code updates from GitHub are automatic.

$ErrorActionPreference = "Stop"
$ProgressPreference    = "SilentlyContinue"

$manifestUrl  = "https://globalmoo.github.io/gmoo-excel-plugin/manifest.xml"
$installDir   = "$env:LOCALAPPDATA\GlobalMOO\ExcelAddin"
$manifestDest = "$installDir\manifest.xml"
$regDeveloper = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"

Write-Host ""
Write-Host "VSME - globalMOO Excel Add-in Installer" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# ── 1. Kill Excel if running ──────────────────────────────────────────────────
$excelProcs = Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue
if ($excelProcs) {
    Write-Host "Excel is currently open. It must be closed to continue." -ForegroundColor Yellow
    $confirm = Read-Host "Close Excel now? Unsaved work will be lost. (y/n)"
    if ($confirm -ne "y") {
        Write-Host "Installation cancelled. Please close Excel and re-run this script." -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
    $excelProcs | Stop-Process -Force
    Start-Sleep -Seconds 2
    Write-Host "  Excel closed." -ForegroundColor Green
}

# ── 2. Create install folder ──────────────────────────────────────────────────
Write-Host "Creating install folder..."
New-Item -ItemType Directory -Path $installDir -Force | Out-Null
Write-Host "  $installDir" -ForegroundColor Green

# ── 3. Download manifest ──────────────────────────────────────────────────────
Write-Host "Downloading manifest..."
try {
    Invoke-WebRequest -Uri $manifestUrl -OutFile $manifestDest -UseBasicParsing
    Write-Host "  manifest.xml downloaded." -ForegroundColor Green
} catch {
    Write-Host ""
    Write-Host "ERROR: Could not download the manifest." -ForegroundColor Red
    Write-Host "  Check your internet connection and try again." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# ── 4. Read add-in ID from manifest ──────────────────────────────────────────
Write-Host "Reading add-in ID..."
$manifestContent = Get-Content $manifestDest -Raw
if ($manifestContent -match '<Id>([^<]+)</Id>') {
    $addinId = $matches[1].Trim()
    Write-Host "  ID: $addinId" -ForegroundColor Green
} else {
    Write-Host "ERROR: Could not find add-in ID in manifest.xml." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# ── 5. Register via WEF\Developer (no admin required) ────────────────────────
Write-Host "Registering add-in..."
if (-not (Test-Path $regDeveloper)) {
    New-Item -Path $regDeveloper -Force | Out-Null
}
New-ItemProperty -Path $regDeveloper -Name $addinId -Value $manifestDest -PropertyType String -Force | Out-Null
Write-Host "  Registered." -ForegroundColor Green

# ── 6. Clean up legacy TrustedCatalogs entry if present ──────────────────────
$oldKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}"
if (Test-Path $oldKey) {
    Remove-Item -Path $oldKey -Force | Out-Null
    Write-Host "  Removed legacy catalog entry." -ForegroundColor Gray
}

# ── 7. Done ───────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "Installation complete!" -ForegroundColor Green
Write-Host ""
Write-Host "Open Excel -- the VSME add-in will appear in your Home ribbon." -ForegroundColor Cyan
Write-Host ""
Write-Host "Code updates deploy automatically. Re-run this script only if" -ForegroundColor Gray
Write-Host "prompted to after a major version update." -ForegroundColor Gray
Write-Host ""
Read-Host "Press Enter to close"
