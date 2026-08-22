@echo off
setlocal
cd /d "%~dp0"
title Formatador ABNT para Word

if "%~1"=="" (
    powershell -ExecutionPolicy Bypass -File "%~dp0ABNT.ps1"
) else (
    powershell -ExecutionPolicy Bypass -File "%~dp0ABNT.ps1" -FilePath "%~1"
)
