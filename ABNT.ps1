<#
.SYNOPSIS
    Unified idempotent launcher for the ABNT Formatter Word Add-in.
.DESCRIPTION
    Forces port 38472 to be free, starts the local HTTPS server,
    confirms it is responding, and opens Microsoft Word with the 'Normas ABNT' tab loaded.
    Pass a file path as the first argument to open an existing document instead of a blank one.
#>

[CmdletBinding()]
param (
    [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
    [string]$FilePath
)

$ErrorActionPreference = "SilentlyContinue"
$ProjectDir = $PSScriptRoot
if (-not $ProjectDir) { $ProjectDir = (Get-Location).Path }
$Port = 38472

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " FORMATADOR ABNT - INICIALIZADOR AUTOMATICO DO WORD         " -ForegroundColor Yellow
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# ------------------------------------------------------------
# 1. CHECK NODE DEPENDENCIES
# ------------------------------------------------------------
Write-Host "[1/4] Verificando dependencias Node.js..." -ForegroundColor Cyan
$NodeModulesPath = Join-Path $ProjectDir "node_modules"
if (-not (Test-Path $NodeModulesPath)) {
    Write-Host "  -> Instalando dependencias necessarias..." -ForegroundColor Yellow
    npm install --no-audit 2>$null | Out-Null
}
Write-Host "  -> Dependencias OK." -ForegroundColor Green

# ------------------------------------------------------------
# 2. CHECK SSL CERTIFICATES
# ------------------------------------------------------------
Write-Host "[2/4] Verificando certificados de seguranca SSL..." -ForegroundColor Cyan
$CertDir = Join-Path $env:USERPROFILE ".office-addin-dev-certs"
$CaCertPath = Join-Path $CertDir "ca.crt"
$LocalhostCertPath = Join-Path $CertDir "localhost.crt"

if (-not (Test-Path $CaCertPath) -or -not (Test-Path $LocalhostCertPath)) {
    Write-Host "  -> Gerando novos certificados de desenvolvimento..." -ForegroundColor Yellow
    npx office-addin-dev-certs install --days 365 2>$null | Out-Null
}
Write-Host "  -> Certificados SSL configurados." -ForegroundColor Green

# ------------------------------------------------------------
# 3. FORCE-RESTART THE LOCAL SERVER ON PORT 38472
# ------------------------------------------------------------
Write-Host "[3/4] Reiniciando servidor local (Porta $Port)..." -ForegroundColor Cyan

# Kill any process currently listening on the port
try {
    $processes = Get-NetTCPConnection -LocalPort $Port -ErrorAction SilentlyContinue |
                 Select-Object -ExpandProperty OwningProcess -Unique
    foreach ($pidToKill in $processes) {
        if ($pidToKill -gt 0) {
            Stop-Process -Id $pidToKill -Force -ErrorAction SilentlyContinue
            Write-Host "  -> Processo anterior (PID $pidToKill) encerrado." -ForegroundColor Yellow
        }
    }
} catch {}

# Start the Node.js server in background (hidden window)
$serverScript = Join-Path $ProjectDir "server.js"
Start-Process -FilePath "node" -ArgumentList "`"$serverScript`"" -WorkingDirectory $ProjectDir -WindowStyle Hidden

# Wait up to ~6 seconds for the server to start listening
$serverReady = $false
for ($i = 0; $i -lt 10; $i++) {
    Start-Sleep -Milliseconds 600
    try {
        $testConn = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue
        if ($testConn) {
            $serverReady = $true
            break
        }
    } catch {}
}

if ($serverReady) {
    Write-Host "  -> Servidor HTTPS ativo e pronto em https://localhost:$Port" -ForegroundColor Green
} else {
    Write-Host "  -> Aviso: Aguardando inicializacao do servidor na porta $Port..." -ForegroundColor Yellow
}

# ------------------------------------------------------------
# 4. OPEN MICROSOFT WORD WITH THE ADD-IN SIDELOADED
# ------------------------------------------------------------
Write-Host "[4/4] Abrindo Microsoft Word com o Suplemento ABNT..." -ForegroundColor Cyan

$ManifestPath = Join-Path $ProjectDir "manifest.xml"

if ($FilePath -and (Test-Path $FilePath)) {
    $ResolvedPath = (Resolve-Path $FilePath).Path
    Write-Host "  -> Documento existente: $ResolvedPath" -ForegroundColor Yellow
    npx office-addin-debugging start "$ManifestPath" desktop --app word --document "$ResolvedPath" --no-debug
} else {
    Write-Host "  -> Novo Documento..." -ForegroundColor Green
    npx office-addin-debugging start "$ManifestPath" desktop --app word --no-debug
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " SUCESSO: Microsoft Word aberto com o Formatador ABNT!      " -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "Dica: Clique na aba 'Normas ABNT' no topo do Word para usar." -ForegroundColor Yellow
Write-Host ""
