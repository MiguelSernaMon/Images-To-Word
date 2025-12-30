@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul 2>&1
title Imagenes a Word - Instalador

echo ========================================
echo    Imagenes a Word - Instalador
echo ========================================
echo.

cd /d "%~dp0"

REM Verificar si Python esta instalado
where python >nul 2>&1
if !ERRORLEVEL! NEQ 0 (
    echo [ERROR] Python no esta instalado o no esta en PATH.
    echo.
    echo Por favor, instala Python desde:
    echo https://www.python.org/downloads/
    echo.
    echo IMPORTANTE: Durante la instalacion, marca la opcion
    echo "Add Python to PATH"
    echo.
    pause
    exit /b 1
)

echo [OK] Python encontrado
python --version
echo.

REM Eliminar entorno virtual anterior si existe (puede estar corrupto)
if exist ".venv" (
    echo [*] Eliminando entorno virtual anterior...
    rmdir /s /q ".venv"
)

REM Crear entorno virtual
echo [*] Creando entorno virtual...
python -m venv .venv
if !ERRORLEVEL! NEQ 0 (
    echo [ERROR] No se pudo crear el entorno virtual.
    pause
    exit /b 1
)
echo [OK] Entorno virtual creado
echo.

REM Instalar dependencias
echo [*] Instalando dependencias...
echo.
".venv\Scripts\python.exe" -m pip install --upgrade pip
".venv\Scripts\python.exe" -m pip install -r requirements.txt

if !ERRORLEVEL! NEQ 0 (
    echo [ERROR] No se pudieron instalar las dependencias.
    pause
    exit /b 1
)

echo.
echo ========================================
echo [OK] Instalacion completada!
echo ========================================
echo.
echo Ahora puedes ejecutar "Imagenes a Word.bat"
echo para iniciar la aplicacion.
echo.
pause
endlocal
