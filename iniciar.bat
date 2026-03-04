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
 
:: 4. Verificar Ghostscript (para compresion de PDFs)
set "GS_FOUND=0"
where gswin64c >nul 2>&1 && set "GS_FOUND=1"
if "%GS_FOUND%"=="0" (
    if exist "C:\Program Files\gs\gs*\bin\gswin64c.exe" set "GS_FOUND=1"
)
if "%GS_FOUND%"=="0" (
    echo [INSTALANDO] Ghostscript para compresion de PDFs...
    echo   Descargando instalador...
    powershell -Command "Invoke-WebRequest -Uri 'https://github.com/ArtifexSoftware/ghostpdl-downloads/releases/download/gs10060/gs10060w64.exe' -OutFile '%TEMP%\gs_installer.exe' -UseBasicParsing"
    if %ERRORLEVEL% neq 0 (
        echo [AVISO] No se pudo descargar Ghostscript. La compresion de PDFs no estara disponible.
        goto :skip_gs
    )
    echo   Instalando silenciosamente...
    "%TEMP%\gs_installer.exe" /S
    del "%TEMP%\gs_installer.exe" 2>nul
    echo [OK] Ghostscript instalado.
) else (
    echo [OK] Ghostscript encontrado.
)
:skip_gs

:: 5. Verificar pypdf
py -c "import pypdf" >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo [INSTALANDO] pypdf...
    py -m pip install "pypdf>=4.0,<5.0" --quiet
)
echo [OK] pypdf instalado.

echo.
echo ========================================
echo   Iniciando Gmail Downloader...
echo ========================================
echo.

:: 6. Ejecutar la app
py "%~dp0app.py"
