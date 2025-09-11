param(
    [switch]$UseStaging
)

# === Imports ===
Import-Module Posh-ACME -Force
Import-Module "$PSScriptRoot\Modules\PoshAcmeHelpers.psm1" -Force
Import-Module "$PSScriptRoot\Modules\PaperCutIntegration.psm1" -Force

# === Config ===
$Fqdn         = 'rkb-demo-cli1.rkbtesting.net'
$ContactEmail = 'mailto:admin@rkbtesting.net'

$pcRoot     = 'C:\Program Files\PaperCut MF'
$keytool    = Join-Path $pcRoot 'runtime\jre\bin\keytool.exe'
$ksFile     = Join-Path $pcRoot 'server\custom\my-ssl-keystore'
$propsPath  = Join-Path $pcRoot 'server\server.properties'
$dstPfx     = Join-Path $pcRoot 'server\custom\MySslExportCert.pfx'

# === Passwords ===
$pfxPass  = Get-Content 'C:\ProgramData\PaperCut\CertRenew\pfxpass.txt' | ConvertTo-SecureString
$ksPass   = Get-Content 'C:\ProgramData\PaperCut\CertRenew\keystorepass.txt' | ConvertTo-SecureString
$plainPfx = [System.Net.NetworkCredential]::new('', $pfxPass).Password
$plainKs  = [System.Net.NetworkCredential]::new('', $ksPass).Password

# === Step 1: Check existing certificate ===
$cert = Get-ExistingCertificate -Fqdn $Fqdn -UseStaging:$UseStaging -PfxPass $pfxPass -ContactEmail $ContactEmail

if ($cert.NotAfter -le (Get-Date).AddDays(30)) {
    Write-Host ("Certificate for {0} expires soon ({1}). Renewing..." -f $Fqdn, $cert.NotAfter)
    $cert = New-PACertificate $Fqdn -Plugin WebSelfHost -PluginArgs @{ } -PfxPass $pfxPass -FriendlyName "PaperCut-$Fqdn"
}
else {
    Write-Host ("No renewal needed. Current expiry: {0}" -f $cert.NotAfter)
}

# === Step 2: Ensure PFX available ===
if ($cert.PfxFullChain -and (Test-Path $cert.PfxFullChain)) {
    Write-Host ("Using renewed PFX from cache: {0}" -f $cert.PfxFullChain)
    Copy-Item $cert.PfxFullChain $dstPfx -Force
}
else {
    Write-Warning ("No PFX found in cache, attempting to rebuild for {0}..." -f $Fqdn)

    if (-not (Test-Path $cert.CertFile) -or -not (Test-Path $cert.KeyFile)) {
        throw ("No PEMs available for {0}. Cannot rebuild." -f $Fqdn)
    }

    $rebuilt = Rebuild-PfxFromPem -CertFile $cert.CertFile -KeyFile $cert.KeyFile -ChainFile $cert.ChainFile -OutPfx $dstPfx -Password $pfxPass
    $dstPfx = $rebuilt
}

# === Step 3: Verify PFX ===
if (-not (Test-Path $dstPfx)) {
    throw ("Renewal failed: expected PFX file {0} not found" -f $dstPfx)
}

try {
    $testCert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList @($dstPfx, $plainPfx)
}
catch {
    throw ("Failed to load renewed PFX {0}: {1}" -f $dstPfx, $_.Exception.Message)
}

if (-not $testCert.HasPrivateKey) {
    throw ("Renewed PFX {0} does not contain a private key. Aborting PaperCut import." -f $dstPfx)
}

Write-Host ("Verified renewed PFX at {0} (Subject: {1}, Expiry: {2})" -f $dstPfx, $testCert.Subject, $testCert.NotAfter)

# === Step 4: Import into PaperCut ===
Import-PfxToKeystore -PfxPath $dstPfx -KeytoolPath $keytool -KeystorePath $ksFile -PfxPassword $plainPfx -KeystorePassword $plainKs
Update-ServerProperties -ServerPropsPath $propsPath -KeystorePath $ksFile -KeystorePassword $plainKs
Restart-PaperCut
