@echo off
setlocal enabledelayedexpansion
title Iniciando Gmail Downloader...
color 0A

:: Resolver ruta del script para que funcione desde cualquier carpeta
set "SCRIPT_DIR=%~dp0"
cd /d "!SCRIPT_DIR!"

echo ===================================================
echo    Iniciando App de Descarga de Correos
echo ===================================================
echo.

:: Verificar que app.py existe
if not exist "!SCRIPT_DIR!app.py" (
    echo [ERROR] No se encontro app.py en la carpeta del script.
    pause
    exit /b
)

set PYTHON_CMD=python

:: 1. Intentar comando normal (si esta en PATH)
%PYTHON_CMD% --version >nul 2>&1
if !errorlevel! equ 0 goto :PYTHON_FOUND

echo [INFO] Python no esta en el PATH, buscando en carpetas comunes...

:: 2. Buscar en directorios comunes donde suele instalarse
for /d %%D in ("%LOCALAPPDATA%\Programs\Python\Python*") do (
    if exist "%%D\python.exe" (
        set "PYTHON_CMD=%%D\python.exe"
        goto :PYTHON_FOUND
    )
)
for /d %%D in ("C:\Python*") do (
    if exist "%%D\python.exe" (
        set "PYTHON_CMD=%%D\python.exe"
        goto :PYTHON_FOUND
    )
)
for /d %%D in ("C:\Program Files\Python*") do (
    if exist "%%D\python.exe" (
        set "PYTHON_CMD=%%D\python.exe"
        goto :PYTHON_FOUND
    )
)
for /d %%D in ("C:\Program Files (x86)\Python*") do (
    if exist "%%D\python.exe" (
        set "PYTHON_CMD=%%D\python.exe"
        goto :PYTHON_FOUND
    )
)

:: 3. Si no se encontro en ningun lado, preguntar para descargar
echo.
echo [ALERTA] No se pudo encontrar Python en tu sistema.
echo La aplicacion necesita Python para funcionar.
set /p DOWNLOAD_CHOICE="Quieres que lo descargue e instale automaticamente ahora? (S/N): "

if /i "%DOWNLOAD_CHOICE%" neq "S" (
    echo.
    echo Instalacion cancelada. Por favor instala Python manualmente desde python.org
    pause
    exit /b
)

echo.
echo [DESCARGANDO] Descargando instalador de Python 3.11 (Oficial)...
:: Usamos curl que viene incluido en Windows 10/11
curl -# -o python_installer.exe https://www.python.org/ftp/python/3.11.8/python-3.11.8-amd64.exe
if not exist python_installer.exe (
    echo [ERROR] Fallo la descarga. Comprueba tu conexion a internet.
    pause
    exit /b
)

echo [INSTALANDO] Instalando Python silenciosamente (esto tomara unos minutos, por favor espera)...
:: Instalar sin mostrar ventanas, agregando al PATH
start /wait python_installer.exe /quiet InstallAllUsers=0 PrependPath=1 Include_test=0
if !errorlevel! neq 0 (
    echo [ERROR] La instalacion de Python fallo.
    pause
    exit /b
)
echo [INFO] Instalacion terminada. Limpiando el instalador...
del python_installer.exe

:: Volver a buscar despues de instalar
echo [INFO] Buscando la nueva instalacion...
for /d %%D in ("%LOCALAPPDATA%\Programs\Python\Python*") do (
    if exist "%%D\python.exe" (
        set "PYTHON_CMD=%%D\python.exe"
        goto :PYTHON_FOUND
    )
)

:: Si aun asi falla, intentar el comando basico
python --version >nul 2>&1
if !errorlevel! neq 0 (
     echo [ERROR] Instalacion completada, pero es necesario reiniciar esta ventana.
     echo Por favor, cierra esta ventana negra y vuelve a hacer doble click en Arrancar_App.bat
     pause
     exit /b
) else (
     set PYTHON_CMD=python
     goto :PYTHON_FOUND
)


:PYTHON_FOUND
echo [OK] Python detectado en: !PYTHON_CMD!
echo.

echo [1/3] Verificando dependencias...
:: Verificar si ya estan instaladas antes de correr pip install
"!PYTHON_CMD!" -c "import playwright; import pypdf; import xhtml2pdf" >nul 2>&1
if !errorlevel! neq 0 (
    echo [INSTALANDO] Instalando dependencias faltantes...
    "!PYTHON_CMD!" -m pip install -r "!SCRIPT_DIR!requirements.txt" --quiet
    if !errorlevel! neq 0 (
        echo [ERROR] Hubo un problema instalando las dependencias necesarias.
        pause
        exit /b
    )
)
echo [OK] Dependencias listas.
echo.

echo [2/3] Verificando Chromium para PDFs...
set "PLAYWRIGHT_BROWSERS_PATH=!SCRIPT_DIR!browsers"
set "CHROMIUM_FOUND=0"
for /d %%d in ("!PLAYWRIGHT_BROWSERS_PATH!\chromium_headless_shell-*") do set "CHROMIUM_FOUND=1"
if "!CHROMIUM_FOUND!"=="0" (
    echo [INSTALANDO] Chromium headless en carpeta local...
    "!PYTHON_CMD!" -m playwright install chromium
    if !errorlevel! neq 0 (
        echo [ERROR] No se pudo instalar Chromium.
        pause
        exit /b
    )
    :: Eliminar chromium completo si existe (solo necesitamos headless shell)
    for /d %%d in ("!PLAYWRIGHT_BROWSERS_PATH!\chromium-*") do (
        echo [LIMPIEZA] Eliminando browser completo innecesario...
        rmdir /s /q "%%d"
    )
)
echo [OK] Chromium portable listo.
echo.

echo [3/3] Abriendo la aplicacion...
:: Intentar lanzar con pythonw (sin consola) si existe, sino con python normal
:: Obtener directorio real del ejecutable de Python
for %%I in ("!PYTHON_CMD!") do set "PYTHON_DIR=%%~dpI"
if exist "!PYTHON_DIR!pythonw.exe" (
    start "" "!PYTHON_DIR!pythonw.exe" "!SCRIPT_DIR!app.py"
) else (
    start "" "!PYTHON_CMD!" "!SCRIPT_DIR!app.py"
)

exit
