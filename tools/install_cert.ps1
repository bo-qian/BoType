<#
Install the self-signed certificate used to sign the VSTO manifests into LocalMachine stores
Run this script from an elevated PowerShell (Run as Administrator).
Usage:
  pwsh.exe -ExecutionPolicy Bypass -File .\tools\install_cert.ps1 -PfxPath ".\BoType\BoType_TemporaryKey.pfx" -PfxPassword "123456"
#>
param(
    [string]$PfxPath = "BoType\BoType_TemporaryKey.pfx",
    [string]$PfxPassword = "123456"
)

function Write-Info($msg){ Write-Host "[Info] $msg" -ForegroundColor Cyan }
function Write-ErrorAndExit($msg){ Write-Host "[Error] $msg" -ForegroundColor Red; exit 1 }

if (-not (Test-Path $PfxPath)){
    Write-ErrorAndExit "PFX file not found: $PfxPath"
}

try{
    Write-Info "Loading PFX..."
    $pwd = ConvertTo-SecureString -String $PfxPassword -AsPlainText -Force
    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
    $cert.Import($PfxPath, $PfxPassword, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable)
}catch{
    Write-ErrorAndExit "Failed to load PFX: $_"
}

# Try install to LocalMachine stores. Requires elevated privileges.
try{
    Write-Info "Adding certificate to LocalMachine\\My (Personal)..."
    $storeMy = New-Object System.Security.Cryptography.X509Certificates.X509Store("My","LocalMachine")
    $storeMy.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
    $storeMy.Add($cert)
    $storeMy.Close()

    Write-Info "Adding certificate to LocalMachine\\Root (Trusted Root Certification Authorities)..."
    $storeRoot = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root","LocalMachine")
    $storeRoot.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
    $storeRoot.Add($cert)
    $storeRoot.Close()

    Write-Info "Adding certificate to LocalMachine\\TrustedPublisher..."
    $storeTP = New-Object System.Security.Cryptography.X509Certificates.X509Store("TrustedPublisher","LocalMachine")
    $storeTP.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
    $storeTP.Add($cert)
    $storeTP.Close()

    # Export CER next to PFX
    $cerPath = Join-Path (Split-Path -Parent $PfxPath) "BoType_TemporaryKey.cer"
    [System.IO.File]::WriteAllBytes($cerPath, $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert))
    Write-Info "Exported certificate to: $cerPath"

    Write-Info "Certificate thumbprint: $($cert.Thumbprint)"
    Write-Info "Installation complete. You can now run setup.exe without the untrusted-certificate error."
}catch{
    Write-ErrorAndExit "Failed to install certificate. Make sure you ran this script as Administrator. Error: $_"
}
