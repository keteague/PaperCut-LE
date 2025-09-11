param(
    [switch]$UseStaging
)

# === Dependencies ===
if (-not (Get-Module -ListAvailable -Name Posh-ACME)) {
    Install-Module Posh-ACME -Scope AllUsers -Force
}
Import-Module Posh-ACME -Force

# Adjust if modules live elsewhere
Import-Module "$PSScriptRoot\Modules\PoshAcmeHelpers.psm1" -Force
Import-Module "$PSScriptRoot\Modules\PaperCutIntegration.psm1" -Force
Import-Module "$PSScriptRoot\Modules\Renewal.psm1" -Force

# === Config ===
$Fqdn         = 'papercut.domain.com'
$ContactEmail = 'mailto:admin@domain.com'

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

# === Step 1: Get or issue certificate ===
$cert = Get-ExistingCertificate -Fqdn $Fqdn -UseStaging:$UseStaging -PfxPass $pfxPass -ContactEmail $ContactEmail

if (-not $cert.PfxFullChain -or -not (Test-Path $cert.PfxFullChain)) {
    Write-Warning ("No PFX found in cache for {0}. Attempting to rebuild from PEMs..." -f $Fqdn)

    if (-not $cert.CertFile -or -not (Test-Path $cert.CertFile)) {
        throw ("No certificate PEMs available for {0} and issuance failed. Check ACME challenge/port 80." -f $Fqdn)
    }

    $rebuilt = Rebuild-PfxFromPem `
        -CertFile  $cert.CertFile `
        -KeyFile   $cert.KeyFile `
        -ChainFile $cert.ChainFile `
        -OutPfx    $dstPfx `
        -Password  $pfxPass
    $dstPfx = $rebuilt
}
else {
    Write-Host ("Using PFX from Posh-ACME cache: {0}" -f $cert.PfxFullChain)
    Copy-Item $cert.PfxFullChain $dstPfx -Force
}

# === Step 2: Verify PFX before import (with fallback) ===
if (-not (Test-Path $dstPfx)) {
    throw ("Expected PFX file {0} not found" -f $dstPfx)
}

$testCert = $null
try {
    $testCert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList @($dstPfx, $plainPfx)
}
catch {
    Write-Warning ("Failed to load {0} with current password ({1}). Attempting rebuild from PEMs..." -f $dstPfx, $_.Exception.Message)

    if ($cert.CertFile -and (Test-Path $cert.CertFile) -and $cert.KeyFile -and (Test-Path $cert.KeyFile)) {
        $dstPfx = Rebuild-PfxFromPem `
            -CertFile  $cert.CertFile `
            -KeyFile   $cert.KeyFile `
            -ChainFile $cert.ChainFile `
            -OutPfx    $dstPfx `
            -Password  $pfxPass
        $testCert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList @($dstPfx, $plainPfx)
    }
    else {
        throw ("Unable to rebuild {0} — PEMs missing." -f $dstPfx)
    }
}

if (-not $testCert.HasPrivateKey) {
    throw ("The generated PFX {0} does not contain a private key. Aborting before PaperCut import." -f $dstPfx)
}

Write-Host ("Verified PFX at {0} (Subject: {1}, Expiry: {2})" -f $dstPfx, $testCert.Subject, $testCert.NotAfter)

# === Step 3: Import into PaperCut + update config ===
Import-PfxToKeystore -PfxPath $dstPfx -KeytoolPath $keytool -KeystorePath $ksFile -PfxPassword $plainPfx -KeystorePassword $plainKs
Update-ServerProperties -ServerPropsPath $propsPath -KeystorePath $ksFile -KeystorePassword $plainKs
Restart-PaperCut

# === Step 4: Install renewal task ===
Install-RenewalTask -RenewScriptPath "C:\Scripts\Renew-PaperCutCert.ps1" -Fqdn $Fqdn
