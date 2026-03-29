<#
One-step script to import the project's self-signed PFX into LocalMachine trust stores and run setup.exe.
Run from repository root in an elevated PowerShell (Run as Administrator).
Usage examples:
  pwsh -ExecutionPolicy Bypass -File .\tools\install_and_run_setup.ps1
  pwsh -ExecutionPolicy Bypass -File .\tools\install_and_run_setup.ps1 -PfxPath ".\BoType\BoType_TemporaryKey.pfx" -PfxPassword "123456" -SetupPath ".\BoType_Release\setup.exe"
#>
param(
    [string]$PfxPath = "BoType\BoType_TemporaryKey.pfx",
    [string]$PfxPassword = "123456",
    [string]$SetupPath = "BoType\publish\setup.exe",
    [switch]$NoRunSetup
)

function Write-Info($m){ Write-Host "[Info] $m" -ForegroundColor Cyan }
function Write-Err($m){ Write-Host "[Error] $m" -ForegroundColor Red }

# Relaunch as admin if not elevated
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)){
    Write-Host "This script must be run as Administrator. Relaunching elevated..."
    Start-Process -FilePath pwsh -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" $($MyInvocation.BoundParameters.Keys | ForEach-Object { "-$($_) `"$($MyInvocation.BoundParameters[$_])`"" } )" -Verb RunAs
    exit
}

if (-not (Test-Path $PfxPath)){
    Write-Err "PFX file not found: $PfxPath"
    exit 1
}

try{
    Write-Info "Importing PFX into LocalMachine\\My..."
    $securePwd = ConvertTo-SecureString -String $PfxPassword -AsPlainText -Force
    $imported = Import-PfxCertificate -FilePath (Resolve-Path $PfxPath) -CertStoreLocation Cert:\LocalMachine\My -Password $securePwd -Exportable
    if (-not $imported){ Write-Err "Import-PfxCertificate returned no objects."; exit 1 }
    $cert = $imported[0]

    Write-Info "Exporting public CER to same folder..."
    $cerPath = Join-Path (Split-Path -Parent (Resolve-Path $PfxPath)) "BoType_TemporaryKey.cer"
    Export-Certificate -Cert $cert -FilePath $cerPath -Force | Out-Null
    Write-Info "CER exported to: $cerPath"

    Write-Info "Importing CER into Trusted Root and Trusted Publishers..."
    Import-Certificate -FilePath $cerPath -CertStoreLocation Cert:\LocalMachine\Root | Out-Null
    Import-Certificate -FilePath $cerPath -CertStoreLocation Cert:\LocalMachine\TrustedPublisher | Out-Null

    Write-Info "Certificate thumbprint: $($cert.Thumbprint)"
    Write-Info "Certificate successfully trusted."
}catch{
    Write-Err "Failed to install/trust certificate: $_"
    exit 1
}

if (-not $NoRunSetup){
    if (-not (Test-Path $SetupPath)){
        Write-Err "Setup not found at: $SetupPath. Please provide correct path via -SetupPath or run setup manually."
        exit 1
    }
    Write-Info "Launching installer: $SetupPath"
    Start-Process -FilePath (Resolve-Path $SetupPath) -Wait
    Write-Info "Installer finished."
}

Write-Info "Done."