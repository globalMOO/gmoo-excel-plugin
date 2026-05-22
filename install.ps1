# GlobalMOO GMOO Excel Add-in Installer
# Installs the add-in for the current user without admin rights.
# Requires PowerShell 5.1+ and Excel for Microsoft 365 (desktop).
#
# What the -ApiUrl param does:
#   - Probes TLS trust against that host so we can prompt-and-import an
#     untrusted self-signed or private-CA cert into CurrentUser\Root.
#   - Reminds the user of the URL to enter when adding a Connection.
#   It does NOT pre-configure the Connection inside the add-in. Use the
#   activation deep-link flow for zero-paste setup.
#
# Usage:
#   Interactive (one-liner):
#     irm https://globalmoo.github.io/gmoo-excel-plugin/install.ps1 | iex
#
#   Non-interactive with pre-supplied values:
#     $args = @{ApiUrl='https://10.0.0.5/api/'; CertFile='C:\acme-ca.crt'; NoInteractive=$true}
#     & ([scriptblock]::Create((irm https://globalmoo.github.io/gmoo-excel-plugin/install.ps1))) @args
#
#   Trust the API server cert only (skip the install steps — used by the
#   in-app "Test connection" hand-off so the user isn't booted from Excel):
#     & ([scriptblock]::Create((irm https://globalmoo.github.io/gmoo-excel-plugin/install.ps1))) -ApiUrl 'https://10.0.0.5/api/' -CertOnly

[CmdletBinding()]
param(
    [string]$ApiUrl,
    [string]$CertFile,
    [switch]$NoInteractive,
    [switch]$CertOnly
)

$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$manifestUrl     = "https://globalmoo.github.io/gmoo-excel-plugin/manifest.xml"
$installDir      = Join-Path $env:LOCALAPPDATA "GlobalMOO\ExcelAddin"
$manifestDest    = Join-Path $installDir "manifest.xml"
$regDeveloper    = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
$oldCatalogKey   = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{7B3A2F4C-1E9D-4B8A-A6C5-3D0E2F9B1C7A}"

Write-Host ""
Write-Host "GlobalMOO GMOO Excel Add-in Installer" -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host ""

# ── Helpers ───────────────────────────────────────────────────────────────────

function Get-NormalizedApiUrl {
    param([string]$Url)
    $u = $Url.Trim()
    if (-not $u) { return $null }
    if ($u -notmatch '^https?://') { $u = "https://$u" }
    if ($u -notmatch '/$') { $u = "$u/" }
    try {
        $parsed = [Uri]$u
        if ($parsed.Scheme -ne 'https') {
            Write-Host "ERROR: API URL must use https:// (the add-in is served from HTTPS and browsers block mixed content)." -ForegroundColor Red
            return $null
        }
        return $u
    } catch {
        return $null
    }
}

function Test-TlsTrust {
    param([string]$Url)
    $u = [Uri]$Url
    $apiHost = $u.Host
    $port = if ($u.Port -gt 0) { $u.Port } else { 443 }

    $result = @{
        Reachable = $false
        Trusted   = $false
        Cert      = $null
        Chain     = $null
        ErrorMsg  = $null
    }

    $client = New-Object System.Net.Sockets.TcpClient
    try {
        $connectTask = $client.ConnectAsync($apiHost, $port)
        if (-not $connectTask.Wait(10000)) {
            $result.ErrorMsg = "TCP connect timed out to ${apiHost}:${port}"
            return $result
        }
    } catch {
        $result.ErrorMsg = "TCP connect failed: $($_.Exception.Message)"
        return $result
    }
    $result.Reachable = $true

    $captured = @{ Errors = $null; Cert = $null; Chain = $null }
    $callback = [System.Net.Security.RemoteCertificateValidationCallback]{
        param($sender, $cert, $chain, $errors)
        $captured.Errors = $errors
        if ($cert) {
            $captured.Cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($cert)
        }
        if ($chain -and $chain.ChainElements) {
            $chainCerts = @()
            foreach ($el in $chain.ChainElements) {
                $chainCerts += New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($el.Certificate)
            }
            $captured.Chain = $chainCerts
        }
        return $true
    }

    $stream = New-Object System.Net.Security.SslStream($client.GetStream(), $false, $callback)
    try {
        $stream.AuthenticateAsClient($apiHost)
    } catch {
        $result.ErrorMsg = "TLS handshake failed: $($_.Exception.Message)"
        try { $stream.Close() } catch {}
        try { $client.Close() } catch {}
        return $result
    }
    try { $stream.Close() } catch {}
    try { $client.Close() } catch {}

    $result.Cert  = $captured.Cert
    $result.Chain = $captured.Chain
    $result.Trusted = ($captured.Errors -eq [System.Net.Security.SslPolicyErrors]::None)
    if (-not $result.Trusted) {
        $result.ErrorMsg = "Policy errors: $($captured.Errors)"
    }
    return $result
}

function Format-CertDetails {
    param([System.Security.Cryptography.X509Certificates.X509Certificate2]$Cert)
    $fp = ($Cert.Thumbprint -replace '(..)(?=.)', '$1:')
    return @"
  Subject:     $($Cert.Subject)
  Issuer:      $($Cert.Issuer)
  Valid from:  $($Cert.NotBefore.ToString('yyyy-MM-dd'))
  Valid to:    $($Cert.NotAfter.ToString('yyyy-MM-dd'))
  SHA-1:       $fp
"@
}

function Import-CertToUserRoot {
    param([System.Security.Cryptography.X509Certificates.X509Certificate2]$Cert)
    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root", "CurrentUser")
    $store.Open("ReadWrite")
    $store.Add($Cert)
    $store.Close()
}

function Get-CertFromFile {
    param([string]$Path)
    if (-not (Test-Path $Path)) {
        throw "Certificate file not found: $Path"
    }
    # X509Certificate2 constructor auto-detects PEM vs DER for .crt/.cer/.pem files
    return New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($Path)
}

# ── 1. Resolve API URL ────────────────────────────────────────────────────────
if (-not $ApiUrl) {
    if ($NoInteractive) {
        Write-Host "ERROR: -ApiUrl is required in non-interactive mode." -ForegroundColor Red
        exit 1
    }
    $defaultPrompt = "https://app.globalmoo.com/api/"
    $entered = Read-Host "Local API URL (press Enter for $defaultPrompt)"
    if (-not $entered) { $entered = $defaultPrompt }
    $ApiUrl = $entered
}

$ApiUrl = Get-NormalizedApiUrl $ApiUrl
if (-not $ApiUrl) {
    Write-Host "ERROR: Invalid API URL." -ForegroundColor Red
    exit 1
}
Write-Host "API URL: $ApiUrl" -ForegroundColor Green
Write-Host ""

# ── 2. TLS probe + cert install ───────────────────────────────────────────────
Write-Host "Testing TLS connection..."
$probe = Test-TlsTrust $ApiUrl
if (-not $probe.Reachable) {
    Write-Host "ERROR: Cannot reach the API server." -ForegroundColor Red
    Write-Host "  $($probe.ErrorMsg)" -ForegroundColor Red
    exit 1
}

if ($probe.Trusted) {
    Write-Host "  TLS certificate is already trusted by Windows." -ForegroundColor Green
} else {
    if (-not $probe.Cert -and -not $CertFile) {
        Write-Host "ERROR: TLS handshake failed and no certificate was retrieved." -ForegroundColor Red
        Write-Host "  $($probe.ErrorMsg)" -ForegroundColor Gray
        Write-Host "  Check that the API server is running and speaks HTTPS on the expected port." -ForegroundColor Gray
        exit 1
    }
    Write-Host "  TLS certificate is NOT trusted by Windows." -ForegroundColor Yellow
    Write-Host "  Reason: $($probe.ErrorMsg)" -ForegroundColor Gray
    Write-Host ""

    if ($CertFile) {
        Write-Host "Loading certificate from file: $CertFile"
        $certToImport = Get-CertFromFile $CertFile
        Write-Host "Certificate to be trusted:" -ForegroundColor Cyan
        Write-Host (Format-CertDetails $certToImport)
        Write-Host ""
    } else {
        $chainLen = if ($probe.Chain) { $probe.Chain.Count } else { 0 }
        Write-Host "Server presented a chain of $chainLen certificate(s)." -ForegroundColor Gray
        Write-Host ""

        if ($chainLen -gt 1) {
            $rootCert = $probe.Chain[$chainLen - 1]
            $leafCert = $probe.Chain[0]
            Write-Host "Root of chain (recommended to trust):" -ForegroundColor Cyan
            Write-Host (Format-CertDetails $rootCert)
            Write-Host ""
            Write-Host "Leaf certificate:" -ForegroundColor Cyan
            Write-Host (Format-CertDetails $leafCert)
            Write-Host ""

            if ($NoInteractive) {
                $certToImport = $rootCert
                Write-Host "  Non-interactive: trusting root CA." -ForegroundColor Gray
            } else {
                $choice = Read-Host "Trust [R]oot CA (recommended), [L]eaf only, or [N]o thanks?"
                switch ($choice.ToLower()) {
                    'r'     { $certToImport = $rootCert }
                    'l'     { $certToImport = $leafCert }
                    default { Write-Host "Cancelled." -ForegroundColor Red; exit 1 }
                }
            }
        } else {
            $certToImport = $probe.Cert
            Write-Host "Self-signed certificate:" -ForegroundColor Cyan
            Write-Host (Format-CertDetails $certToImport)
            Write-Host ""
            if (-not $NoInteractive) {
                $confirm = Read-Host "Trust this certificate? (y/n)"
                if ($confirm -ne 'y') { Write-Host "Cancelled." -ForegroundColor Red; exit 1 }
            }
        }
    }

    Write-Host "Importing certificate to CurrentUser Trusted Root..."
    try {
        Import-CertToUserRoot $certToImport
        Write-Host "  Imported." -ForegroundColor Green
    } catch {
        Write-Host "ERROR: Failed to import certificate." -ForegroundColor Red
        Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }

    Write-Host "Re-testing TLS connection..."
    $probe2 = Test-TlsTrust $ApiUrl
    if ($probe2.Trusted) {
        Write-Host "  TLS certificate now trusted." -ForegroundColor Green
    } else {
        Write-Host "WARNING: TLS still not trusted after import." -ForegroundColor Yellow
        Write-Host "  $($probe2.ErrorMsg)" -ForegroundColor Gray
        Write-Host "  The cert was imported, but Windows validation still fails." -ForegroundColor Yellow
        Write-Host "  This usually means the server cert has another issue (expired, hostname mismatch, missing intermediate)." -ForegroundColor Yellow
        if (-not $NoInteractive -and -not $CertOnly) {
            $cont = Read-Host "Continue with install anyway? (y/n)"
            if ($cont -ne 'y') { exit 1 }
        }
    }
}
Write-Host ""

# ── Cert-only mode: skip the install steps and exit ─────────────────────────
if ($CertOnly) {
    Write-Host "Cert-only mode: skipping add-in install steps." -ForegroundColor Cyan
    Write-Host "Done. Click Retry in the Excel task pane." -ForegroundColor Green
    Write-Host ""
    exit 0
}

# ── 3. Kill Excel if running ──────────────────────────────────────────────────
$excelProcs = Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue
if ($excelProcs) {
    Write-Host "Excel is currently open. It must be closed to continue." -ForegroundColor Yellow
    if ($NoInteractive) {
        Write-Host "  Closing Excel (non-interactive)..." -ForegroundColor Yellow
    } else {
        $confirm = Read-Host "Close Excel now? Unsaved work will be lost. (y/n)"
        if ($confirm -ne "y") {
            Write-Host "Installation cancelled. Please close Excel and re-run this script." -ForegroundColor Red
            exit 1
        }
    }
    $excelProcs | Stop-Process -Force
    Start-Sleep -Seconds 2
    Write-Host "  Excel closed." -ForegroundColor Green
}

# ── 4. Create install folder ──────────────────────────────────────────────────
Write-Host "Creating install folder..."
New-Item -ItemType Directory -Path $installDir -Force | Out-Null
Write-Host "  $installDir" -ForegroundColor Green

# ── 5. Download manifest ──────────────────────────────────────────────────────
Write-Host "Downloading manifest..."
try {
    Invoke-WebRequest -Uri $manifestUrl -OutFile $manifestDest -UseBasicParsing
    Write-Host "  manifest.xml downloaded." -ForegroundColor Green
} catch {
    Write-Host ""
    Write-Host "ERROR: Could not download the manifest." -ForegroundColor Red
    Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  URL: $manifestUrl" -ForegroundColor Gray
    exit 1
}

# ── 6. Read add-in ID from manifest ───────────────────────────────────────────
Write-Host "Reading add-in ID from manifest..."
$content = Get-Content $manifestDest -Raw
if ($content -match '<Id>([^<]+)</Id>') {
    $addinId = $matches[1].Trim()
    Write-Host "  ID: $addinId" -ForegroundColor Green
} else {
    Write-Host "ERROR: Could not find add-in ID in manifest.xml." -ForegroundColor Red
    exit 1
}

# ── 7. Register add-in via Developer key ──────────────────────────────────────
Write-Host "Registering add-in..."
if (-not (Test-Path $regDeveloper)) {
    New-Item -Path $regDeveloper -Force | Out-Null
}
New-ItemProperty -Path $regDeveloper -Name $addinId -Value $manifestDest -PropertyType String -Force | Out-Null
Write-Host "  Registry entry written." -ForegroundColor Green

# ── 8. Clean up previous failed attempts (TrustedCatalogs) ────────────────────
if (Test-Path $oldCatalogKey) {
    Remove-Item -Path $oldCatalogKey -Force | Out-Null
    Write-Host "  Removed old TrustedCatalogs entry." -ForegroundColor Gray
}

# ── 9. Done ───────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "Installation complete!" -ForegroundColor Green
Write-Host ""
Write-Host "Next step:" -ForegroundColor Cyan
Write-Host "  Open Excel -- the GlobalMOO GMOO add-in will appear in your Home ribbon."
Write-Host "  In the task pane, add a Connection pointing at: $ApiUrl"
Write-Host ""
Write-Host "If it does not appear, go to:" -ForegroundColor Gray
Write-Host "  Home -> Add-ins -> and look for GlobalMOO GMOO in the panel." -ForegroundColor Gray
Write-Host ""
