@echo off
title Compilador - Ferramenta EACE NOC
echo ================================================================
echo   EXECUTANDO COMPILACAO COMPLETA (PYTHON -> EXE -> INSTALADOR)
echo ================================================================
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0build_installer.ps1"
echo.
echo Processo finalizado!
pause
