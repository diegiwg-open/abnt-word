# Setup Trusted Catalog for Microsoft Word
$catalogPath = "C:\Users\User\Downloads\abnt-word"
$shareName = "ABNTWordAddIn"
$uncPath = "\\$($env:COMPUTERNAME)\$shareName"

Write-Host "Configurando catalogo compartilhado do Office..." -ForegroundColor Cyan

# 1. Cria compartilhamento de rede local Windows
try {
    net share $shareName /DELETE 2>$null
    net share "$shareName=$catalogPath" "/grant:Everyone,FULL"
    Write-Host "[OK] Compartilhamento de rede criado: $uncPath" -ForegroundColor Green
} catch {
    Write-Host "[AVISO] Nao foi possivel criar o share automaticamente. Use o caminho da pasta." -ForegroundColor Yellow
}

# 2. Configura Chave do Registro do Office para Word 2016/2019/2021/2024/365
$regPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs"
$guid1 = "{c4d9a1e0-7612-4cf4-9154-0549c4f828a1}"
$key1 = "$regPath\$guid1"

if (-not (Test-Path $regPath)) {
    New-Item -Path $regPath -Force | Out-Null
}

if (-not (Test-Path $key1)) {
    New-Item -Path $key1 -Force | Out-Null
}

Set-ItemProperty -Path $key1 -Name "Id" -Value $guid1 -Type String
Set-ItemProperty -Path $key1 -Name "Url" -Value $uncPath -Type String
Set-ItemProperty -Path $key1 -Name "Flags" -Value 1 -Type DWord

# Também adiciona o caminho local C: direto como catálogo secundário
$guid2 = "{d5e0b2f1-8723-4df5-0265-1650d5f939b2}"
$key2 = "$regPath\$guid2"
if (-not (Test-Path $key2)) {
    New-Item -Path $key2 -Force | Out-Null
}
Set-ItemProperty -Path $key2 -Name "Id" -Value $guid2 -Type String
Set-ItemProperty -Path $key2 -Name "Url" -Value $catalogPath -Type String
Set-ItemProperty -Path $key2 -Name "Flags" -Value 1 -Type DWord

Write-Host "[OK] Registro configurado com sucesso!" -ForegroundColor Green
Write-Host "Caminhos registrados:"
Write-Host "1. $uncPath" -ForegroundColor Yellow
Write-Host "2. $catalogPath" -ForegroundColor Yellow
