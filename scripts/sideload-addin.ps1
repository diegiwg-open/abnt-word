# Script PowerShell para Registro do Suplemento ABNT no Microsoft Word (2021 / 2024 / 365)
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " Instalador do Suplemento ABNT para Microsoft Word (Windows) " -ForegroundColor Yellow
Write-Host "============================================================" -ForegroundColor Cyan

$CurrentDir = Split-Path -Parent $PSScriptRoot
$ManifestPath = Join-Path $CurrentDir "manifest.xml"
$CatalogDir = "$env:LOCALAPPDATA\Microsoft\Office\AddInSharedCatalog"

if (-not (Test-Path $ManifestPath)) {
    Write-Host "ERRO: manifest.xml não encontrado em $ManifestPath" -ForegroundColor Red
    exit 1
}

# Criar pasta de catálogo compartilhado
if (-not (Test-Path $CatalogDir)) {
    New-Item -ItemType Directory -Path $CatalogDir -Force | Out-Null
    Write-Host "Criada pasta de catálogo: $CatalogDir" -ForegroundColor Green
}

# Copiar manifesto para a pasta de catálogo
Copy-Item -Path $ManifestPath -Destination (Join-Path $CatalogDir "abnt-manifest.xml") -Force
Write-Host "Manifesto copiado para o catálogo do Office." -ForegroundColor Green

# Registrar pasta no Registro do Windows do Office (Word 2016/2019/2021/2024/365 compartilham chave 16.0)
$RegPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs"
$CatalogGuid = "{c4d9a1e0-7612-4cf4-9154-0549c4f828a1}"
$CatalogRegKey = Join-Path $RegPath $CatalogGuid

try {
    if (-not (Test-Path $RegPath)) {
        New-Item -Path $RegPath -Force | Out-Null
    }

    if (-not (Test-Path $CatalogRegKey)) {
        New-Item -Path $CatalogRegKey -Force | Out-Null
    }

    Set-ItemProperty -Path $CatalogRegKey -Name "Id" -Value $CatalogGuid -Type String
    Set-ItemProperty -Path $CatalogRegKey -Name "Url" -Value $CatalogDir -Type String
    Set-ItemProperty -Path $CatalogRegKey -Name "Flags" -Value 1 -Type DWord

    Write-Host "Catálogo registrado com sucesso no Registro do Office!" -ForegroundColor Green
    Write-Host ""
    Write-Host "Como usar no Word:" -ForegroundColor Yellow
    Write-Host "1. Abra o Microsoft Word."
    Write-Host "2. Vá em Inserir -> Obter Suplementos -> Compartilhado / Meus Suplementos."
    Write-Host "3. O 'Formatador ABNT Automático' estará disponível para inserção imediata!"
    Write-Host "4. Uma aba 'Normas ABNT' aparecerá na sua Faixa de Opções (Ribbon)."
}
catch {
    Write-Host "Aviso ao registrar catálogo no registro: $($_.Exception.Message)" -ForegroundColor Yellow
}

Write-Host "============================================================" -ForegroundColor Cyan
