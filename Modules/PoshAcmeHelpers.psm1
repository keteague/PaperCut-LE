function Get-ExistingCertificate {
    param(
        [string]$Fqdn,
        [switch]$UseStaging
    )
    if ($UseStaging) { Set-PAServer LE_STAGE } else { Set-PAServer LE_PROD }
    $cert = Get-PACertificate $Fqdn -ErrorAction SilentlyContinue
    return $cert
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

    $leaf = [System.Security.Cryptography.X509Certificates.X509Certificate2]::CreateFromPemFile($CertFile, $KeyFile)
    $coll = [System.Security.Cryptography.X509Certificates.X509Certificate2Collection]::new()
    [void]$coll.Add($leaf)

    if ($ChainFile -and (Test-Path $ChainFile)) {
        $pem = Get-Content $ChainFile -Raw
        [regex]::Matches($pem,"-----BEGIN CERTIFICATE-----.*?-----END CERTIFICATE-----",'Singleline') |
            ForEach-Object {
                $b64 = ($_.Value -replace '-----BEGIN CERTIFICATE-----','' -replace '-----END CERTIFICATE-----','' -replace '\s','')
                $bytes = [Convert]::FromBase64String($b64)
                [void]$coll.Add([System.Security.Cryptography.X509Certificates.X509Certificate2]::new($bytes))
            }
    }

    $pfxBytes = $coll.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pkcs12, $plain)
    [IO.File]::WriteAllBytes($OutPfx, $pfxBytes)
    return $OutPfx
}
