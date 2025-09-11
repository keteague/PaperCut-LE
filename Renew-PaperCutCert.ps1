param(
    [Parameter(Mandatory=$true)]
    [string]$Fqdn,

    [switch]$UseStaging
)

$passDir        = 'C:\ProgramData\PaperCut\CertRenew'
$pfxPassFile    = Join-Path $passDir 'pfxpass.txt'
$ksPassFile     = Join-Path $passDir 'keystorepass.txt'
$pcDir          = 'C:\Program Files\PaperCut MF\server\custom'
$dstPfx         = Join-Path $pcDir 'MySslExportCert.pfx'
$dstKeystore    = Join-Path $pcDir 'my-ssl-keystore'
$keytoolPath    = 'C:\Program Files\PaperCut MF\runtime\jre\bin\keytool.exe'

# --- Load passwords ---
$plainPfx = Get-Content $pfxPassFile -Raw
$pfxPass  = ConvertTo-SecureString -String $plainPfx -AsPlainText -Force

$plainKs  = Get-Content $ksPassFile -Raw
$ksPw     = ConvertTo-SecureString -String $plainKs -AsPlainText -Force

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

# --- Get/renew cert ---
$cert = Get-PACertificate $Fqdn -ErrorAction SilentlyContinue
if ($null -eq $cert) {
    $cert = New-PACertificate $Fqdn -Plugin WebSelfHost -PluginArgs @{} -PfxPass $pfxPass -FriendlyName "PaperCut-$Fqdn"
} else {
    if ($cert.NotAfter -lt (Get-Date).AddDays(30)) {
        $cert = Submit-Renewal $Fqdn
        Write-Host "Certificate renewed. New expiry: $($cert.NotAfter)"
    } else {
        Write-Host "No renewal needed. Current expiry: $($cert.NotAfter)"
    }
}

# --- Normalize PFX with plain-text password ---
Write-Host "🔐 Normalizing PFX password for PaperCut..."
$src = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 `
    -ArgumentList @($cert.PfxFullChain, $plainPfx, 'Exportable,PersistKeySet')

$bytes = $src.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pkcs12, $plainPfx)
[IO.File]::WriteAllBytes($dstPfx, $bytes)
Write-Host "✅ Re-exported to $dstPfx with PaperCut password"

# --- Import into PaperCut keystore ---
Import-PaperCutPfx -PfxPath $dstPfx -KeystorePath $dstKeystore -KeytoolPath $keytoolPath -KeystorePassword $plainKs -PfxPassword $plainPfx

Write-Host "🎉 PaperCut SSL certificate renewed successfully!"
