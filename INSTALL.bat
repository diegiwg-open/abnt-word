@echo off
chcp 65001 >nul
title Instalador Formatador ABNT
color 0A

echo.
echo ╔══════════════════════════════════════════════════════════════╗
echo ║     FORMATADOR ABNT - INSTALADOR AUTOMÁTICO                   ║
echo ║     Suplemento para Microsoft Word 2021/2024/365              ║
echo ╚══════════════════════════════════════════════════════════════╝
echo.

:: Verificar se está sendo executado como administrador
net session >nul 2>&1
if %errorLevel% == 0 (
    echo ✓ Executando como administrador
) else (
    echo ⚠️  Aviso: Recomendado executar como administrador
    echo    Alguns recursos podem não funcionar corretamente
    echo.
)

:: Verificar se Node.js está instalado
echo [1/5] Verificando Node.js...
node --version >nul 2>&1
if %errorLevel% == 0 (
    echo ✓ Node.js encontrado
    node --version
) else (
    echo ✗ Node.js não encontrado
    echo   Baixe em: https://nodejs.org/
    echo   O instalador continuará, mas algumas funções podem não funcionar
)
echo.

:: Verificar se Git está instalado (opcional)
echo [2/5] Verificando Git...
git --version >nul 2>&1
if %errorLevel% == 0 (
    echo ✓ Git encontrado
    git --version
) else (
    echo ⚠ Git não encontrado (opcional)
)
echo.

:: Instalar dependências se Node.js estiver disponível
echo [3/5] Instalando dependências...
if exist "package.json" (
    if %errorLevel% == 0 (
        call npm install
        if %errorLevel% == 0 (
            echo ✓ Dependências instaladas com sucesso
        ) else (
            echo ✗ Erro ao instalar dependências
        )
    ) else (
        echo ⚠ Node.js não disponível, pulando instalação de dependências
    )
) else (
    echo ✗ package.json não encontrado
)
echo.

:: Configurar certificados SSL (se disponível)
echo [4/5] Configurando certificados SSL...
if %errorLevel% == 0 (
    call npx office-addin-dev-certs install
    if %errorLevel% == 0 (
        echo ✓ Certificados SSL configurados
    ) else (
        echo ⚠ Erro ao configurar certificados SSL
    )
) else (
    echo ⚠ Node.js não disponível, pulando configuração SSL
)
echo.

:: Iniciar servidor
echo [5/5] Iniciando servidor local...
if %errorLevel% == 0 (
    echo ✓ Iniciando servidor na porta 38472...
    start /B node server.js
    timeout /t 3 /nobreak >nul
    echo ✓ Servidor iniciado com sucesso
) else (
    echo ⚠ Node.js não disponível, servidor não será iniciado
)
echo.

echo ═══════════════════════════════════════════════════════════════
echo.
echo ✓ INSTALAÇÃO CONCLUÍDA!
echo.
echo Para usar o Formatador ABNT:
echo 1. Abra o Microsoft Word
echo 2. Vá em Arquivo ^> Opções ^> Suplementos ^> Trust Center
echo 3. Clique em "Configurações do Trust Center"
echo 4. Em "Local do suplemento", adicione o diretório atual
echo 5. Clique em "Gerenciar Suplementos" e selecione "ABNT"
echo.
echo Ou execute: ABNT.bat (para abrir automaticamente)
echo.
echo ═══════════════════════════════════════════════════════════════
echo.

pause