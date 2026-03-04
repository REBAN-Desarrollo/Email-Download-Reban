@echo off
setlocal enabledelayedexpansion
title Compilando Gmail Downloader a EXE...
color 0B

:: Resolver ruta del script
set "SCRIPT_DIR=%~dp0"
cd /d "!SCRIPT_DIR!"

echo ===================================================
echo   Empaquetando App (Python + Dependencias + Script)
echo ===================================================
echo.
echo Esto convertira tu aplicacion en un unico archivo .exe
echo que puedes llevar a cualquier PC con Windows sin
echo necesidad de instalar Python ni dependencias.
echo.
echo [1/3] Instalando PyInstaller...
python -m pip install pyinstaller xhtml2pdf --quiet
if !errorlevel! neq 0 (
    echo [ERROR] No se pudo instalar PyInstaller. Asegurate de tener internet y Python.
    pause
    exit /b
)

:: Limpiar dist/ viejo antes de compilar
if exist dist rmdir /s /q dist

echo.
echo [2/3] Compilando la aplicacion (esto puede tardar un par de minutos)...
:: --noconsole quita la ventana negra de atras
:: --onefile mete todo (python y dependencias) en un solo archivo
:: --name le da el nombre al archivo final
python -m PyInstaller --noconsole --onefile --name "Gmail_Downloader" app.py

if !errorlevel! neq 0 (
    echo.
    echo [ERROR] Hubo un problema durante la compilacion.
    pause
    exit /b
)

echo.
echo [3/3] Limpiando archivos temporales...
rmdir /s /q build
del /q Gmail_Downloader.spec

echo.
echo ===================================================
echo   COMPILACION EXITOSA!
echo ===================================================
echo.
echo Tu aplicacion portable esta lista.
echo Ve a la carpeta "dist" y ahi encontraras "Gmail_Downloader.exe".
echo Ese archivo contiene TODO lo necesario para funcionar.
echo.
pause
exit
