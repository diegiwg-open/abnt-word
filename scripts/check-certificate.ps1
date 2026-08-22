# Verifies whether the Office Add-in development CA certificate
# is present in the Windows CurrentUser and LocalMachine Root trust stores.

$certPath = "C:\Users\User\.office-addin-dev-certs\ca.crt"
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($certPath)
Write-Host "Subject:    " $cert.Subject
Write-Host "Thumbprint: " $cert.Thumbprint

$store = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root", "CurrentUser")
$store.Open("ReadOnly")
$found = $store.Certificates.Find(
    [System.Security.Cryptography.X509Certificates.X509FindType]::FindByThumbprint,
    $cert.Thumbprint,
    $false
)
Write-Host "Found in CurrentUser Root:" ($found.Count -gt 0)
$store.Close()

$storeLM = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root", "LocalMachine")
$storeLM.Open("ReadOnly")
$foundLM = $storeLM.Certificates.Find(
    [System.Security.Cryptography.X509Certificates.X509FindType]::FindByThumbprint,
    $cert.Thumbprint,
    $false
)
Write-Host "Found in LocalMachine Root:" ($foundLM.Count -gt 0)
$storeLM.Close()
