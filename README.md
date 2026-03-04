# Gmail Automator & Downloader

Aplicación de escritorio nativa para Windows (construida con Python y `tkinter`) diseñada para buscar, previsualizar y descargar en lote (batch) correos electrónicos y sus archivos adjuntos desde Gmail mediante IMAP.

Incluye un conversor integrado que transforma automáticamente el cuerpo (texto o HTML) de los correos en archivos `.pdf` legibles.

---

## Características Principales

- **Interfaz Gráfica Sencilla (GUI):** Construida con `tkinter`, sin necesidad de instalaciones pesadas.
- **Probar Conexión:** Botón dedicado para verificar credenciales antes de buscar correos.
- **Búsqueda Avanzada IMAP:**
  - Remitente (`FROM`)
  - Rango de fechas (`SINCE` / `BEFORE`)
  - Asunto (`SUBJECT`)
  - Palabras clave (`TEXT` - Busca en cuerpo y encabezados)
- **Descarga en Lote (Batch):** Permite multiselección (Ctrl/Shift) en la tabla para descargar decenas de correos a la vez.
- **Auto-conversión a PDF:** Genera un archivo `Mensaje_Legible.pdf` por cada correo, incluyendo los metadatos originales (De, Asunto, Fecha).
- **Gestión de Adjuntos:** Descarga y guarda todos los archivos adjuntos (PDFs, Excel, Word, imágenes) en su formato original junto al correo.
- **Formato de Salida Dinámico:** Soporte para variables (wildcards) personalizables en la GUI (`{date}`, `{subject}`, `{sender}`, `{id}`) para nombrar automáticamente las carpetas de descarga.
- **Script de Auto-Instalación:** Incluye un `.bat` para usuarios de Windows que instala Python automáticamente si no lo detecta en el sistema.

---

## Requisitos y Configuración

### 1. Requisitos del Sistema
- **Sistema Operativo:** Windows 10/11 (recomendado)
- **Python:** Python 3.8 o superior (el script `Arrancar_App.bat` intenta instalar la 3.11 automáticamente si no se encuentra).
- **Dependencias Python:** `xhtml2pdf` (para generación de PDFs).

### 2. Configuración en Gmail (Muy Importante)
Debido a las políticas de seguridad de Google, no puedes usar la contraseña normal de tu correo.
1. Activa la **Verificación en 2 pasos** en tu Cuenta de Google.
2. Ve a **Seguridad** > **Contraseñas de aplicaciones**.
3. Genera una nueva contraseña para "Mail" y copia los 16 caracteres. **Esta es la contraseña que se utiliza en la aplicación.**

---

## Cómo Ejecutar

### Opción 1: El modo fácil (Usuarios finales)
Haz doble clic en el archivo **`Arrancar_App.bat`**.
Este script buscará Python en tu sistema, instalará las dependencias necesarias y abrirá la GUI sin mostrar consolas negras. Si no tienes Python, te ofrecerá descargarlo e instalarlo de forma silenciosa.

> **Nota:** El `.bat` funciona correctamente aunque lo ejecutes desde otra carpeta (ej: un acceso directo en el escritorio).

### Opción 2: El modo desarrollador
Si ya tienes el entorno configurado:
```bash
# 1. Instalar requerimientos
pip install -r requirements.txt

# 2. Ejecutar la app
python app.py
```

---

## Compilar a .exe

Para generar un ejecutable portable que no requiere Python instalado:

1. Haz doble clic en **`Compilar_App.bat`**
2. Espera a que termine la compilación (puede tardar unos minutos)
3. El archivo `Gmail_Downloader.exe` estará en la carpeta `dist/`

O manualmente:
```bash
pip install pyinstaller xhtml2pdf
python -m PyInstaller --noconsole --onefile --name "Gmail_Downloader" app.py
```

---

## Solución de Problemas

| Problema | Solución |
|----------|----------|
| **"Autenticación fallida"** | Verifica que usas una App Password de 16 caracteres, no tu contraseña normal de Gmail. Usa el botón "Probar Conexión" para validar. |
| **"Python no encontrado"** | Ejecuta `Arrancar_App.bat` que lo instala automáticamente, o instálalo manualmente desde [python.org](https://python.org). |
| **PDF no se genera** | Verifica que `xhtml2pdf` está instalado: `pip install xhtml2pdf`. La app mostrará un aviso si falta. |
| **Correos marcados como leídos** | Esto ya no ocurre. La app usa `BODY.PEEK[]` tanto en búsqueda como en descarga. |
| **El .bat no abre la app** | Asegúrate de que `app.py` está en la misma carpeta que el `.bat`. |
| **Carpetas duplicadas** | La app ahora añade un sufijo numérico (`_1`, `_2`) automáticamente si la carpeta ya existe. |

---

## Notas para Desarrolladores

### Estructura de Código (`app.py`)
- **UI & Hilos:** La app usa `threading` para separar la interfaz gráfica (Main Thread) de las operaciones de red (IMAP). Todas las actualizaciones de GUI desde hilos usan `root.after()` para thread-safety.
- **Seguridad en IMAP:** Se utiliza `BODY.PEEK` tanto en búsqueda como en descarga. Esto garantiza que los correos **NO se marquen como leídos** en el servidor.
- **Sanitización:** La función `clean_filename()` limpia metadatos decodificados para evitar errores de I/O en Windows. Los metadatos en PDFs se escapan con `html.escape()` para prevenir inyección HTML.
- **Conexión IMAP:** Timeout de 30 segundos. Botón "Probar Conexión" permite validar credenciales sin ejecutar búsquedas.

### Posibles Mejoras (Roadmap)
- Integración con OAuth2 nativo de Google (para evitar la necesidad de App Passwords, aunque requiere validación de proyecto en GCP).
- Soporte extendido para servidores de correo Outlook/Office365 u otros IMAP genéricos (actualmente hardcodeado en `imap.gmail.com`).

---
*Desarrollado para la automatización de la bandeja de entrada y flujos contables / operativos.*
