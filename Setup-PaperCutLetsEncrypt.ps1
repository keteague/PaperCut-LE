param(
    [Parameter(Mandatory=$true)]
    [string]$Fqdn,

    [switch]$UseStaging
)

# --- Paths ---
$passDir        = 'C:\ProgramData\PaperCut\CertRenew'
$pfxPassFile    = Join-Path $passDir 'pfxpass.txt'
$ksPassFile     = Join-Path $passDir 'keystorepass.txt'
$pcDir          = 'C:\Program Files\PaperCut MF\server\custom'
$dstPfx         = Join-Path $pcDir 'MySslExportCert.pfx'
$dstKeystore    = Join-Path $pcDir 'my-ssl-keystore'
$keytoolPath    = 'C:\Program Files\PaperCut MF\runtime\jre\bin\keytool.exe'

# --- Ensure dir exists ---
if (-not (Test-Path $passDir)) {
    New-Item -Path $passDir -ItemType Directory -Force | Out-Null
}

# --- Load or prompt for PFX password ---
if (Test-Path $pfxPassFile) {
    $plainPfx = Get-Content $pfxPassFile -Raw
    $pfxPass  = ConvertTo-SecureString -String $plainPfx -AsPlainText -Force
    Write-Host "Using existing PFX password at $pfxPassFile"
} else {
    $pfxPass = Read-Host "Enter a new PFX export password" -AsSecureString
    $plainPfx = [System.Net.NetworkCredential]::new('', $pfxPass).Password
    $plainPfx | Out-File -FilePath $pfxPassFile -Encoding ascii -NoNewline
    Write-Host "Saved PFX password to $pfxPassFile"
}

# --- Load or prompt for Keystore password ---
if (Test-Path $ksPassFile) {
    $plainKs = Get-Content $ksPassFile -Raw
    $ksPw    = ConvertTo-SecureString -String $plainKs -AsPlainText -Force
    Write-Host "Using existing keystore password at $ksPassFile"
} else {
    $ksPw = Read-Host "Enter PaperCut keystore password (server.ssl.keystore-password)" -AsSecureString
    $plainKs = [System.Net.NetworkCredential]::new('', $ksPw).Password
    $plainKs | Out-File -FilePath $ksPassFile -Encoding ascii -NoNewline
    Write-Host "Saved keystore password to $ksPassFile"
}

# --- Import modules ---
Import-Module Posh-ACME -ErrorAction Stop
Import-Module "$PSScriptRoot\Modules\PoshAcmeHelpers.psm1" -Force
Import-Module "$PSScriptRoot\Modules\PaperCutIntegration.psm1" -Force

# --- ACME server ---
if ($UseStaging) {
    Write-Host "Using Let's Encrypt STAGING server"
    Set-PAServer LE_STAGE
} else {
    Write-Host "Using Let's Encrypt PRODUCTION server"
    Set-PAServer LE_PROD
}

# --- Ensure account ---
try {
    Get-PAAccount -ErrorAction Stop | Out-Null
} catch {
    New-PAAccount -AcceptTOS -Contact "mailto:admin@example.com" | Out-Null
    Write-Host "Created new ACME account"
}

# --- Issue cert ---
$cert = New-PACertificate $Fqdn -Plugin WebSelfHost -PluginArgs @{} -PfxPass $pfxPass -FriendlyName "PaperCut-$Fqdn"

# --- Normalize and import ---
Copy-Item $cert.PfxFullChain $dstPfx -Force
Import-PaperCutPfx -PfxPath $dstPfx -KeystorePath $dstKeystore -KeytoolPath $keytoolPath -KeystorePassword $plainKs -PfxPassword $plainPfx

Write-Host "🎉 Initial PaperCut certificate setup complete for $Fqdn"
