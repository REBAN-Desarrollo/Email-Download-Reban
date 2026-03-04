@echo off
setlocal enabledelayedexpansion
title Iniciando Gmail Downloader...
color 0A

echo ===================================================
echo    Iniciando App de Descarga de Correos
echo ===================================================
echo.

set PYTHON_CMD=python

:: 1. Intentar comando normal (si esta en PATH)
%PYTHON_CMD% --version >nul 2>&1
if %errorlevel% equ 0 goto :PYTHON_FOUND

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
if %errorlevel% neq 0 (
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

echo [1/2] Verificando e instalando dependencias (xhtml2pdf)...
:: Usar pip asociado al python encontrado
"!PYTHON_CMD!" -m pip install xhtml2pdf --quiet
if %errorlevel% neq 0 (
    echo [ERROR] Hubo un problema instalando las dependencias necesarias.
    pause
    exit /b
)
echo Listo.
echo.

echo [2/2] Abriendo la aplicacion...
:: Intentar lanzar con pythonw (sin consola) si existe, sino con python normal
set "PYTHONW_CMD=!PYTHON_CMD:python.exe=pythonw.exe!"
if exist "!PYTHONW_CMD!" (
    start "" "!PYTHONW_CMD!" app.py
) else (
    start "" "!PYTHON_CMD!" app.py
)

exit
