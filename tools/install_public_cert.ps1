<#
Install the public certificate (.cer) into LocalMachine Trusted Root and TrustedPublisher
Run this script from an elevated PowerShell (Run as Administrator).
Usage:
  pwsh.exe -ExecutionPolicy Bypass -File .\tools\install_public_cert.ps1 -CerPath ".\BoType\BoType_TemporaryKey.cer"
#>
param(
    [string]$CerPath = "BoType\BoType_TemporaryKey.cer"
)
function Write-Info($msg){ Write-Host "[Info] $msg" -ForegroundColor Cyan }
function Write-ErrorAndExit($msg){ Write-Host "[Error] $msg" -ForegroundColor Red; exit 1 }
if (-not (Test-Path $CerPath)){
    Write-ErrorAndExit "CER file not found: $CerPath"
}
try{
    Write-Info "Importing certificate to LocalMachine\\Root (Trusted Root Certification Authorities)..."
    Import-Certificate -FilePath (Resolve-Path $CerPath) -CertStoreLocation Cert:\LocalMachine\Root | Out-Null
    Write-Info "Importing certificate to LocalMachine\\TrustedPublisher..."
    Import-Certificate -FilePath (Resolve-Path $CerPath) -CertStoreLocation Cert:\LocalMachine\TrustedPublisher | Out-Null
    Write-Info "Certificate installed. You can now run setup.exe without the untrusted-certificate error."
}catch{
    Write-ErrorAndExit "Failed to install certificate. Make sure you ran this script as Administrator. Error: $_"
}
