# Installs the Office Add-in development CA certificate into the Windows trust store.
# Needed so that Microsoft Word accepts the self-signed HTTPS server without a security warning.

$caPath = "C:\Users\User\.office-addin-dev-certs\ca.crt"

Write-Host "Installing SSL certificate into Windows trust store..." -ForegroundColor Cyan

# Add to Intermediate CA and Personal stores (no UAC prompt required)
Import-Certificate -FilePath $caPath -CertStoreLocation "Cert:\CurrentUser\CA" | Out-Null
Import-Certificate -FilePath $caPath -CertStoreLocation "Cert:\CurrentUser\My" | Out-Null

# Add to Trusted Root store via certutil (may show a system dialog to confirm)
certutil.exe -user -addstore -f Root "$caPath"

Write-Host ""
Write-Host "[OK] Certificate installation completed." -ForegroundColor Green
