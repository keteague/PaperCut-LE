function Import-PaperCutPfx {
    param(
        [string]$PfxPath,
        [string]$KeystorePath,
        [string]$KeytoolPath,
        [string]$KeystorePassword,
        [string]$PfxPassword
    )

    if (Test-Path $KeystorePath) {
        Remove-Item $KeystorePath -Force
    }

    & $KeytoolPath -importkeystore `
        -destkeystore $KeystorePath `
        -deststorepass $KeystorePassword `
        -srckeystore $PfxPath `
        -srcstoretype PKCS12 `
        -srcstorepass $PfxPassword `
        -noprompt

    if ($LASTEXITCODE -ne 0) {
        throw "Keytool import failed with exit code $LASTEXITCODE"
    }

    Restart-Service PCAppServer -Force -ErrorAction Stop
}
