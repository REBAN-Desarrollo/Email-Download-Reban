# Gmail Automator & Downloader 🚀

Aplicación de escritorio nativa para Windows (construida con Python y `tkinter`) diseñada para buscar, previsualizar y descargar en lote (batch) correos electrónicos y sus archivos adjuntos desde Gmail mediante IMAP. 

Incluye un conversor integrado que transforma automáticamente el cuerpo (texto o HTML) de los correos en archivos `.pdf` legibles.

---

## 🌟 Características Principales

- **Interfaz Gráfica Sencilla (GUI):** Construida con `tkinter`, sin necesidad de instalaciones pesadas.
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

## 🛠 Requisitos y Configuración

### 1. Requisitos del Sistema
- **Sistema Operativo:** Windows 10/11 (recomendado)
- **Python:** Python 3.8 o superior (el script `Arrancar_App.bat` intenta instalar la 3.11 automáticamente si no se encuentra).
- **Dependencias Python:** `xhtml2pdf` (para generación de PDFs).

### 2. Configuración en Gmail (¡Muy Importante!)
Debido a las políticas de seguridad de Google, no puedes usar la contraseña normal de tu correo.
1. Activa la **Verificación en 2 pasos** en tu Cuenta de Google.
2. Ve a **Seguridad** > **Contraseñas de aplicaciones**.
3. Genera una nueva contraseña para "Mail" y copia los 16 caracteres. **Esta es la contraseña que se utiliza en la aplicación.**

---

## 🚀 Cómo Ejecutar

### Opción 1: El modo fácil (Usuarios finales)
Haz doble clic en el archivo **`Arrancar_App.bat`**. 
Este script buscará Python en tu sistema, instalará las dependencias necesarias y abrirá la GUI sin mostrar consolas negras. Si no tienes Python, te ofrecerá descargarlo e instalarlo de forma silenciosa.

### Opción 2: El modo desarrollador
Si ya tienes el entorno configurado:
```bash
# 1. Instalar requerimientos
pip install -r requirements.txt

# 2. Ejecutar la app
python app.py
```

---

## 💻 Notas para Desarrolladores

### Estructura de Código (`app.py`)
- **UI & Hilos:** La app usa `threading` para separar la interfaz gráfica (Main Thread) de las operaciones de red (IMAP). Esto evita que la ventana se congele (estado "No responde") mientras descarga múltiples archivos pesados.
- **Seguridad en IMAP:** En la fase de búsqueda, se utiliza la directiva `BODY.PEEK`. Esto garantiza que los correos que se previsualizan en la tabla **NO se marquen como leídos** en el servidor del usuario. Solo se obtiene el estado raw cuando el usuario confirma la descarga `(RFC822)`.
- **Sanitización de Archivos:** La función `clean_filename()` limpia los metadatos decodificados de los correos para evitar errores de I/O en Windows causados por caracteres reservados (`\ / : * ? " < > |`).

### Posibles Mejoras (Roadmap)
- Integración con Oauth2 nativo de Google (para evitar la necesidad de App Passwords, aunque requiere validación de proyecto en GCP).
- Soporte extendido para servidores de correo Outlook/Office365 u otros IMAP genéricos (actualmente hardcodeado en `imap.gmail.com`).
- Compilación a `.exe` usando `PyInstaller` para portabilidad total en un solo archivo:
  `pyinstaller --noconsole --onefile app.py`

---
*Desarrollado para la automatización de la bandeja de entrada y flujos contables / operativos.*
