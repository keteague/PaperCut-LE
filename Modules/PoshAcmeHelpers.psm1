function Get-ExistingCertificate {
    param(
        [string]$Fqdn,
        [switch]$UseStaging,
        [SecureString]$PfxPass,
        [string]$ContactEmail
    )

    # Select ACME server
    if ($UseStaging) {
        Write-Host "Using Let's Encrypt STAGING server"
        Set-PAServer LE_STAGE
    } else {
        Write-Host "Using Let's Encrypt PRODUCTION server"
        Set-PAServer LE_PROD
    }

    # Ensure an account exists
    $acct = Get-PAAccount -ErrorAction SilentlyContinue
    if (-not $acct) {
        Write-Host ("No ACME account found. Creating one with {0}" -f $ContactEmail)
        $acct = New-PAAccount -AcceptTOS -Contact $ContactEmail -Force
        Set-PAAccount $acct.ID | Out-Null
    } else {
        Set-PAAccount $acct.ID | Out-Null
    }

    # Try to fetch cert
    $cert = Get-PACertificate $Fqdn -ErrorAction SilentlyContinue
    if ($cert) {
        Write-Host ("Found existing certificate order for {0}" -f $Fqdn)
        return $cert
    }

    if (-not $PfxPass) {
        throw ("No certificate for {0} and no -PfxPass provided to issue a new one." -f $Fqdn)
    }

    Write-Host ("No existing certificate for {0}. Requesting new one..." -f $Fqdn)
    $newCert = New-PACertificate $Fqdn `
        -Plugin WebSelfHost `
        -PluginArgs @{ } `
        -PfxPass $PfxPass `
        -FriendlyName ("PaperCut-{0}" -f $Fqdn)

    return $newCert
}

function Rebuild-PfxFromPem {
    param(
        [string]$CertFile,
        [string]$KeyFile,
        [string]$ChainFile,
        [string]$OutPfx,
        [SecureString]$Password
    )

    $plain = [System.Net.NetworkCredential]::new('', $Password).Password

    try {
        # --- Load leaf cert ---
        $certPem   = Get-Content $CertFile -Raw
        $certBody  = ($certPem -split "`r?`n") | Where-Object {$_ -and ($_ -notmatch '^---')}
        $certBytes = [Convert]::FromBase64String(($certBody -join ''))
        $leaf      = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList @($certBytes)

        # --- Load key ---
        $keyPem   = Get-Content $KeyFile -Raw
        $keyBody  = ($keyPem -split "`r?`n") | Where-Object {$_ -and ($_ -notmatch '^---')}
        $keyBytes = [Convert]::FromBase64String(($keyBody -join ''))

        $rsa = [System.Security.Cryptography.RSA]::Create()
        $offset = 0
        [void]$rsa.ImportPkcs8PrivateKey($keyBytes, [ref]$offset)

        $leafWithKey = $leaf.CopyWithPrivateKey($rsa)
        if (-not $leafWithKey.HasPrivateKey) {
            throw "Leaf certificate does not have an associated private key"
        }

        $coll = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2Collection
        $coll.Add($leafWithKey) | Out-Null
        Write-Host "✅ Merged with RSA PKCS#8 private key using .NET APIs"

        # --- Add chain certs if present ---
        if ($ChainFile -and (Test-Path $ChainFile)) {
            $chainPem = Get-Content $ChainFile -Raw
            [regex]::Matches($chainPem,"-----BEGIN CERTIFICATE-----.*?-----END CERTIFICATE-----",'Singleline') |
                ForEach-Object {
                    $b64   = ($_.Value -replace '-----BEGIN CERTIFICATE-----','' -replace '-----END CERTIFICATE-----','' -replace '\s','')
                    $bytes = [Convert]::FromBase64String($b64)
                    $c     = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList @($bytes)
                    $coll.Add($c) | Out-Null
                }
        }

        # --- Export PFX ---
        $pfxBytes = $coll.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pkcs12, $plain)
        [IO.File]::WriteAllBytes($OutPfx, $pfxBytes)

        Write-Host ("🎉 Built PFX at {0} with {1} certs (private key included)" -f $OutPfx, $coll.Count)
        return $OutPfx
    }
    catch {
        Write-Warning "⚠️ .NET merge failed: $($_.Exception.Message)"
        Write-Host "➡️ Falling back to OpenSSL for PFX generation..."

        # Temp bundle file
        $tmpPem = [System.IO.Path]::GetTempFileName()
        Get-Content $KeyFile | Out-File -FilePath $tmpPem -Encoding ascii
        Get-Content $CertFile | Out-File -FilePath $tmpPem -Append -Encoding ascii
        if ($ChainFile -and (Test-Path $ChainFile)) {
            Get-Content $ChainFile | Out-File -FilePath $tmpPem -Append -Encoding ascii
        }

        $pwFile = [System.IO.Path]::GetTempFileName()
        Set-Content -Path $pwFile -Value $plain -NoNewline

        $openssl = "C:\Program Files\OpenSSL-Win64\bin\openssl.exe"
        if (-not (Test-Path $openssl)) {
            throw "OpenSSL not found at $openssl. Please install and update the path."
        }

        & $openssl pkcs12 -export `
            -in $tmpPem `
            -out $OutPfx `
            -password file:$pwFile `
            -name "PaperCut-Cert"

        Remove-Item $tmpPem,$pwFile -Force
        if ($LASTEXITCODE -ne 0) {
            throw "OpenSSL PFX generation failed with exit code $LASTEXITCODE"
        }

        Write-Host "✅ Built PFX at $OutPfx using OpenSSL fallback"
        return $OutPfx
    }
}
