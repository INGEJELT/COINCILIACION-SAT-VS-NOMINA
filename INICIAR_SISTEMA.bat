@echo off
title Servidor El Cardenal - NOMINAS VS SAT
color 0A

echo ===================================================
echo      SISTEMA DE COINCILIACION NOMINA VS SAT
echo ===================================================
echo.
echo Iniciando el servidor y abriendo la red VPN...
echo Por favor, NO cierres esta ventana.
echo Para apagar el sistema, presiona Ctrl + C
echo.

:: Esto fuerza a la consola a ubicarse en la carpeta donde esta el .bat
cd /d "%~dp0"

:: Intenta ejecutar con el comando 'python'
python app.py

:: Si 'python' falla (errorlevel distinto de 0), intenta con 'py'
if %errorlevel% neq 0 (
    echo.
    echo [!] Comando 'python' no encontrado, intentando con 'py'...
    py app.py
)

echo.
echo [!] El servidor se detuvo de manera inesperada o hubo un error de codigo.
pause