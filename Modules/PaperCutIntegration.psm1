function Import-PfxToKeystore {
    param(
        [string]$PfxPath,
        [string]$KeytoolPath,
        [string]$KeystorePath,
        [string]$PfxPassword,
        [string]$KeystorePassword
    )
    if (Test-Path $KeystorePath) { Remove-Item $KeystorePath -Force }

    & $KeytoolPath -importkeystore `
      -srckeystore  $PfxPath `
      -srcstoretype pkcs12 `
      -srcstorepass $PfxPassword `
      -destkeystore $KeystorePath `
      -deststorepass $KeystorePassword `
      -destkeypass  $KeystorePassword `
      -noprompt

    if ($LASTEXITCODE -ne 0) { throw "Keytool import failed with exit code $LASTEXITCODE" }
}

function Update-ServerProperties {
    param(
        [string]$ServerPropsPath,
        [string]$KeystorePath,
        [string]$KeystorePassword
    )
    $ksPathForProps = ($KeystorePath -replace '\\','/')
    $props = if (Test-Path $ServerPropsPath) { Get-Content $ServerPropsPath -Raw } else { "" }
    $props = ($props -split "`r?`n" | Where-Object {
        $_ -notmatch '^\s*server\.ssl\.keystore\s*=' -and
        $_ -notmatch '^\s*server\.ssl\.keystore-password\s*=' -and
        $_ -notmatch '^\s*server\.ssl\.key-password\s*='
    }) -join "`r`n"
    $props += "`r`nserver.ssl.keystore=$ksPathForProps"
    $props += "`r`nserver.ssl.keystore-password=$KeystorePassword"
    $props += "`r`nserver.ssl.key-password=$KeystorePassword`r`n"
    Set-Content -Path $ServerPropsPath -Value $props -Encoding UTF8
}

function Restart-PaperCut {
    Restart-Service -Name 'PCAppServer' -Force
}
