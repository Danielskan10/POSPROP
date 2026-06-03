@echo off
chcp 65001 >nul
title Sapienza — Dashboard Posiciones
color 0A
echo.
echo  ╔══════════════════════════════════════════╗
echo  ║  SAPIENZA — Procesador de Posiciones     ║
echo  ╚══════════════════════════════════════════╝
echo.
cd /d "%~dp0"
python procesar_datos.py
if errorlevel 1 (
    echo.
    echo  [ERROR] Revisa que Python este instalado y los datos existan.
    pause
    exit /b 1
)
echo.
pause
