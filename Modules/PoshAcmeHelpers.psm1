function Get-ExistingCertificate {
    param(
        [string]$Fqdn,
        [switch]$UseStaging
    )

    try {
        $cert = Get-PACertificate $Fqdn -ErrorAction Stop
        if ($null -ne $cert) {
            return $cert
        }
    } catch {
        Write-Warning "No existing certificate found for $Fqdn"
    }
    return $null
}
