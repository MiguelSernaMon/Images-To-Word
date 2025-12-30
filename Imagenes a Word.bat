@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul 2>&1
title Imagenes a Word

echo ========================================
echo    Imagenes a Word - Iniciando...
echo ========================================
echo.

cd /d "%~dp0"

REM Verificar si existe el entorno virtual
if not exist ".venv\Scripts\python.exe" (
    echo [!] No se encontro el entorno virtual.
    echo [*] Ejecuta primero "instalar.bat" para configurar la aplicacion.
    echo.
    pause
    exit /b 1
)

echo [*] Iniciando servidor web...
echo [*] Abriendo navegador en http://localhost:5001
echo.
echo Para cerrar la aplicacion, cierra esta ventana.
echo ========================================

REM Abrir navegador despues de 2 segundos
start "" cmd /c "timeout /t 2 /nobreak >nul && start http://localhost:5001"

REM Iniciar servidor
call ".venv\Scripts\python.exe" app.py

endlocal
