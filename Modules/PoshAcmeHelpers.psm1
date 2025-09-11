function Get-ExistingCertificate {
    param(
        [Parameter(Mandatory)]
        [string]$Fqdn,

        [switch]$UseStaging
    )

    # Decide ACME server
    if ($UseStaging) {
        Set-PAServer LE_STAGE | Out-Null
    } else {
        Set-PAServer LE_PROD  | Out-Null
    }

    # Try to fetch cert from Posh-ACME
    try {
        $cert = Get-PACertificate $Fqdn -ErrorAction Stop
    } catch {
        Write-Warning "No existing certificate found for $Fqdn"
        return $null
    }

    # If PFX exists already, return it
    if ($cert -and (Test-Path $cert.PfxFullChain)) {
        return $cert
    }

    # Otherwise, try to rebuild PFX from PEMs
    $certDir  = Split-Path $cert.CertFile
    $certPem  = Join-Path $certDir "cert.cer"
    $keyPem   = Join-Path $certDir "privkey.key"
    if (-not (Test-Path $keyPem)) {
        $keyPem = Join-Path $certDir "cert.key"
    }
    $chainPem = Join-Path $certDir "chain.cer"

    if (-not (Test-Path $certPem) -or -not (Test-Path $keyPem)) {
        throw "Certificate PEM components missing in $certDir"
    }

    $pfxPath = Join-Path $certDir "fullchain-rebuilt.pfx"
    $plainPw = "Temp123!"   # caller should re-export with PaperCut password later

    try {
        $openssl = "C:\Program Files\OpenSSL-Win64\bin\openssl.exe"
        if (Test-Path $openssl) {
            $bundle = [System.IO.Path]::GetTempFileName()
            Get-Content $keyPem   | Out-File $bundle -Encoding ascii
            Get-Content $certPem  | Out-File $bundle -Append -Encoding ascii
            if (Test-Path $chainPem) {
                Get-Content $chainPem | Out-File $bundle -Append -Encoding ascii
            }

            $pwFile = [System.IO.Path]::GetTempFileName()
            Set-Content -Path $pwFile -Value $plainPw -NoNewline

            & $openssl pkcs12 -export -in $bundle -out $pfxPath -password file:$pwFile -name "PaperCut-$Fqdn"

            Remove-Item $bundle,$pwFile -Force
        } else {
            throw "OpenSSL not installed. Cannot rebuild PFX."
        }
    } catch {
        throw "Failed to rebuild PFX: $_"
    }

    $cert | Add-Member -NotePropertyName PfxFullChain -NotePropertyValue $pfxPath -Force
    return $cert
}
