@echo off
chcp 65001 >nul
title Gmail Downloader - Iniciando...

echo ========================================
echo   Gmail Downloader - Verificando...
echo ========================================
echo.

:: 1. Verificar Python
py --version >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Python no esta instalado.
    echo Descargalo de: https://www.python.org/downloads/
    echo Asegurate de marcar "Add Python to PATH" al instalar.
    pause
    exit /b 1
)
echo [OK] Python encontrado.

:: 2. Verificar playwright pip package
py -c "import playwright" >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo [INSTALANDO] playwright...
    py -m pip install playwright --quiet
    if %ERRORLEVEL% neq 0 (
        echo [ERROR] No se pudo instalar playwright.
        pause
        exit /b 1
    )
)
echo [OK] playwright instalado.

:: 3. Verificar Chromium en carpeta local browsers/
set "PLAYWRIGHT_BROWSERS_PATH=%~dp0browsers"
if not exist "%PLAYWRIGHT_BROWSERS_PATH%\chromium_headless_shell-*" (
    echo [INSTALANDO] Chromium headless en carpeta local...
    py -m playwright install chromium
    :: Eliminar chromium completo si existe (solo necesitamos headless shell)
    for /d %%d in ("%PLAYWRIGHT_BROWSERS_PATH%\chromium-*") do (
        echo [LIMPIEZA] Eliminando browser completo innecesario...
        rmdir /s /q "%%d"
    )
    if %ERRORLEVEL% neq 0 (
        echo [ERROR] No se pudo instalar Chromium.
        pause
        exit /b 1
    )
)
echo [OK] Chromium portable listo.

echo.
echo ========================================
echo   Iniciando Gmail Downloader...
echo ========================================
echo.

:: 4. Ejecutar la app
py "%~dp0app.py"
