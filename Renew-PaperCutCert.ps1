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

# --- Get cert object ---
$cert = Get-ExistingCertificate -Fqdn $Fqdn -UseStaging:$UseStaging

# --- Renew if close to expiry ---
if ($cert.NotAfter -lt (Get-Date).AddDays(30)) {
    $cert = Submit-Renewal $Fqdn
    Write-Host "Certificate renewed. New expiry: $($cert.NotAfter)"
} else {
    Write-Host "No renewal needed. Current expiry: $($cert.NotAfter)"
}

# --- Normalize PFX with PaperCut password ---
Write-Host "🔐 Normalizing PFX password for PaperCut..."
$src = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 `
    -ArgumentList @($cert.PfxFullChain, $plainPfx, 'Exportable,PersistKeySet')

$bytes = $src.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pkcs12, $plainPfx)
[IO.File]::WriteAllBytes($dstPfx, $bytes)

# --- Import into PaperCut keystore ---
Import-PaperCutPfx -PfxPath $dstPfx -KeystorePath $dstKeystore -KeytoolPath $keytoolPath -KeystorePassword $plainKs -PfxPassword $plainPfx

Write-Host "🎉 PaperCut SSL certificate renewed successfully for $Fqdn"
