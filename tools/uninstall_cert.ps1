param(
    [string]$Thumbprint
)

function RemoveFromStore($storeName, $thumb){
    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store($storeName,"LocalMachine")
    $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
    $found = $store.Certificates | Where-Object { $_.Thumbprint -eq $thumb }
    foreach($c in $found){ $store.Remove($c) }
    $store.Close()
}

if (-not $Thumbprint){ Write-Host "Usage: pwsh -File .\\tools\\uninstall_cert.ps1 -Thumbprint <thumbprint>"; exit 1 }
RemoveFromStore "My" $Thumbprint
RemoveFromStore "Root" $Thumbprint
RemoveFromStore "TrustedPublisher" $Thumbprint
Write-Host "Removed certificate $Thumbprint from LocalMachine stores."