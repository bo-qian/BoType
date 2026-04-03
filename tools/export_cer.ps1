param(
    [string]$PfxPath = "BoType\BoType_TemporaryKey.pfx",
    [string]$PfxPassword = "123456",
    [string]$CerOut = "BoType\BoType_TemporaryKey.cer"
)

if (-not (Test-Path $PfxPath)) { Write-Error "PFX not found: $PfxPath"; exit 1 }

try {
    $pwd = ConvertTo-SecureString -String $PfxPassword -AsPlainText -Force
    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
    $cert.Import($PfxPath, $PfxPassword, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::DefaultKeySet)
    [System.IO.File]::WriteAllBytes($CerOut, $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert))
    Write-Host "Exported CER to: $CerOut"
} catch {
    Write-Error "Export failed: $_"
    exit 1
}