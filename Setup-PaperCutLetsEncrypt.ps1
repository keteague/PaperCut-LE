# Load modules
if (-not (Get-Module -ListAvailable -Name Posh-ACME)) {
    Install-Module Posh-ACME -Scope AllUsers -Force
}
Import-Module Posh-ACME -Force

Import-Module "$PSScriptRoot\Modules\PoshAcmeHelpers.psm1"
Import-Module "$PSScriptRoot\Modules\PaperCutIntegration.psm1"
Import-Module "$PSScriptRoot\Modules\Renewal.psm1"

$Fqdn       = 'papercut.domain.com'
$pcRoot     = 'C:\Program Files\PaperCut MF'
$keytool    = Join-Path $pcRoot 'runtime\jre\bin\keytool.exe'
$ksFile     = Join-Path $pcRoot 'server\custom\my-ssl-keystore'
$propsPath  = Join-Path $pcRoot 'server\server.properties'
$dstPfx     = Join-Path $pcRoot 'server\custom\MySslExportCert.pfx'

# Load passwords from state dir (or prompt)
$pfxPass = Get-Content 'C:\ProgramData\PaperCut\CertRenew\pfxpass.txt' | ConvertTo-SecureString
$ksPass  = Get-Content 'C:\ProgramData\PaperCut\CertRenew\keystorepass.txt' | ConvertTo-SecureString
$plainPfx = [System.Net.NetworkCredential]::new('', $pfxPass).Password
$plainKs = [System.Net.NetworkCredential]::new('', $ksPass).Password

# 1. Get cert (or rebuild from PEMs)
$cert = Get-ExistingCertificate -Fqdn $Fqdn
if (-not $cert.PfxFullChainPath -or -not (Test-Path $cert.PfxFullChainPath)) {
    $pemDir = "C:\Users\kteague\AppData\Local\Posh-ACME\LE_PROD\2650081001\$Fqdn"
    $dstPfx = Rebuild-PfxFromPem -CertFile (Join-Path $pemDir 'cert.cer') -KeyFile (Join-Path $pemDir 'privkey.key') -ChainFile (Join-Path $pemDir 'chain.cer') -OutPfx $dstPfx -Password $pfxPass
} else {
    Copy-Item $cert.PfxFullChainPath $dstPfx -Force
}

# 2. Import into PaperCut + update config
Import-PfxToKeystore -PfxPath $dstPfx -KeytoolPath $keytool -KeystorePath $ksFile -PfxPassword $plainPfx -KeystorePassword $plainKs
Update-ServerProperties -ServerPropsPath $propsPath -KeystorePath $ksFile -KeystorePassword $plainKs
Restart-PaperCut

# 3. Install renewal task
Install-RenewalTask -RenewScriptPath "C:\Scripts\Renew-PaperCutCert.ps1" -Fqdn $Fqdn
