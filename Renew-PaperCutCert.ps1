<#
.SYNOPSIS
  Renews the PaperCut SSL certificate using Posh-ACME and ensures the PFX
  always has the password stored in pfxpass.txt.

.NOTES
  Run as Administrator.
  Requires: Posh-ACME, Java keytool (from PaperCut runtime\jre\bin).
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$Fqdn,

    [switch]$UseStaging
)

# --- Paths ---
$pfxPassFile    = 'C:\ProgramData\PaperCut\CertRenew\pfxpass.txt'
$keystorePassFile = 'C:\ProgramData\PaperCut\CertRenew\keystorepass.txt'
$pcDir          = 'C:\Program Files\PaperCut MF\server\custom'
$dstPfx         = Join-Path $pcDir 'MySslExportCert.pfx'
$dstKeystore    = Join-Path $pcDir 'my-ssl-keystore'
$keytoolPath    = 'C:\Program Files\PaperCut MF\runtime\jre\bin\keytool.exe'

# --- Load PFX password ---
if (-not (Test-Path $pfxPassFile)) {
    throw "Missing PFX password file at $pfxPassFile"
}
$pfxPass = Get-Content $pfxPassFile | ConvertTo-SecureString
$plainPw = [System.Net.NetworkCredential]::new('', $pfxPass).Password

# --- Load PaperCut keystore password ---
if (-not (Test-Path $keystorePassFile)) {
    $ksPw = Read-Host "Enter PaperCut keystore password (server.ssl.keystore-password)" -AsSecureString
    $ksPlain = [System.Net.NetworkCredential]::new('', $ksPw).Password
    $ksPlain | Out-File -FilePath $keystorePassFile -Encoding ascii -NoNewline
    Write-Host "Saved encrypted keystore password to $keystorePassFile"
} else {
    $ksPw = Get-Content $keystorePassFile | ConvertTo-SecureString
    $ksPlain = [System.Net.NetworkCredential]::new('', $ksPw).Password
    Write-Host "Using existing keystore password at $keystorePassFile"
}

# --- Import modules ---
Import-Module Posh-ACME -ErrorAction Stop
Import-Module "$PSScriptRoot\Modules\PoshAcmeHelpers.psm1" -Force
Import-Module "$PSScriptRoot\Modules\PaperCutIntegration.psm1" -Force

# --- Pick ACME server ---
if ($UseStaging) {
    Write-Host "Using Let's Encrypt STAGING server"
    Set-PAServer LE_STAGE
} else {
    Write-Host "Using Let's Encrypt PRODUCTION server"
    Set-PAServer LE_PROD
}

# --- Ensure ACME account exists ---
try {
    Get-PAAccount -ErrorAction Stop | Out-Null
} catch {
    New-PAAccount -AcceptTOS -Contact "mailto:admin@$($Fqdn.Split('.')[1..2] -join '.')" | Out-Null
    Write-Host "Created new ACME account"
}

# --- Renew or fetch existing cert ---
$cert = Get-PACertificate $Fqdn -ErrorAction SilentlyContinue
if ($null -eq $cert) {
    Write-Host "No existing cert order. Issuing new..."
    $cert = New-PACertificate $Fqdn -Plugin WebSelfHost -PluginArgs @{} -PfxPass $pfxPass -FriendlyName "PaperCut-$Fqdn"
} else {
    $before = $cert.NotAfter
    if ($before -lt (Get-Date).AddDays(30)) {
        Write-Host "Renewing certificate..."
        $cert = Submit-Renewal $Fqdn
        $after = $cert.NotAfter
        Write-Host "Certificate renewed. New expiry: $after"
    } else {
        Write-Host "No renewal needed. Current expiry: $before"
    }
}

# --- Normalize PFX password for PaperCut ---
Write-Host "🔐 Rebuilding PFX from PEMs for PaperCut..."
$plainPw = [System.Net.NetworkCredential]::new('', $pfxPass).Password
$certDir  = Split-Path $cert.CertFile
$certPem  = Join-Path $certDir "cert.cer"
$keyPem   = Join-Path $certDir "privkey.key"
if (-not (Test-Path $keyPem)) {
    $keyPem = Join-Path $certDir "cert.key"
}
$chainPem = Join-Path $certDir "chain.cer"

if (-not (Test-Path $certPem) -or -not (Test-Path $keyPem)) {
    throw "Missing PEM components in $certDir"
}

$openssl = "C:\Program Files\OpenSSL-Win64\bin\openssl.exe"
if (-not (Test-Path $openssl)) {
    throw "OpenSSL not found at $openssl. Please install or update path."
}

# Bundle PEMs
$tmpPem = [System.IO.Path]::GetTempFileName()
Get-Content $keyPem  | Out-File $tmpPem -Encoding ascii
Get-Content $certPem | Out-File $tmpPem -Append -Encoding ascii
if (Test-Path $chainPem) {
    Get-Content $chainPem | Out-File $tmpPem -Append -Encoding ascii
}

$pwFile = [System.IO.Path]::GetTempFileName()
Set-Content -Path $pwFile -Value $plainPw -NoNewline

& $openssl pkcs12 -export `
    -in $tmpPem `
    -out $dstPfx `
    -password file:$pwFile `
    -name "PaperCut-$Fqdn"

Remove-Item $tmpPem,$pwFile -Force
if ($LASTEXITCODE -ne 0) {
    throw "OpenSSL failed to build PFX (exit $LASTEXITCODE)"
}

Write-Host "✅ Rebuilt PFX at $dstPfx with PaperCut password"

# --- Import into PaperCut keystore ---
Write-Host "Importing keystore $dstPfx to $dstKeystore..."
& $keytoolPath -importkeystore `
    -destkeystore $dstKeystore `
    -deststorepass $ksPlain `
    -srckeystore  $dstPfx `
    -srcstoretype PKCS12 `
    -srcstorepass $plainPw `
    -noprompt

if ($LASTEXITCODE -ne 0) {
    throw "Keytool import failed with exit code $LASTEXITCODE"
}

# --- Restart PaperCut ---
Restart-Service PCAppServer -Force -ErrorAction Stop
Write-Host "🎉 PaperCut SSL certificate updated successfully!"
