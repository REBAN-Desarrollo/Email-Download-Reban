# 🚀 Gmail Automator & Downloader

Una aplicación de escritorio nativa para Windows (construida con Python y `tkinter`) diseñada para buscar, previsualizar y descargar masivamente correos electrónicos y sus archivos adjuntos desde Gmail mediante IMAP. 

Incluye un conversor integrado que transforma automáticamente el cuerpo (texto o HTML) de los correos en archivos `.pdf` legibles, ideal para auditorías, respaldos contables y flujos de automatización administrativa.

---

## 🌟 Características Principales

- **Interfaz Gráfica Intuitiva (GUI):** Fácil de usar, sin necesidad de interactuar con código o consolas.
- **Búsqueda Avanzada IMAP:** Filtra correos sin marcarlos como leídos en tu bandeja por:
  - Remitente (`FROM`)
  - Rango de fechas (`SINCE` / `BEFORE`)
  - Asunto (`SUBJECT`)
  - Palabras clave (`TEXT` - Busca en cuerpo y encabezados)
- **Descarga en Lote (Batch):** Selecciona múltiples correos con `Ctrl` o `Shift` y descárgalos todos a la vez.
- **Auto-conversión a PDF:** Genera un archivo `Mensaje_Legible.pdf` por cada correo, incluyendo los metadatos originales (De, Asunto, Fecha).
- **Extracción de Adjuntos:** Guarda todos los archivos adjuntos (PDFs, Excel, XML, imágenes) en su formato original.
- **Formato de Salida Dinámico:** Usa variables (`{date}`, `{subject}`, `{sender}`, `{id}`) para personalizar cómo se nombran las carpetas de descarga.
- **Portabilidad (Standalone):** Opción para compilar todo en un único archivo `.exe` que no requiere instalar nada en la computadora destino.

---

## 🔑 Configuración Requerida (Gmail App Password)

Por políticas de seguridad, Google **no permite** que inicies sesión en aplicaciones externas usando tu contraseña normal. Debes generar una **Contraseña de Aplicación** de 16 caracteres.

### Paso a paso para obtenerla (Toma 1 minuto):
1. Ve a [Mi Cuenta de Google](https://myaccount.google.com/) e inicia sesión con el correo que vas a usar.
2. En el menú izquierdo, haz clic en **Seguridad**.
3. Busca la sección *"Cómo iniciar sesión en Google"* y asegúrate de que la **Verificación en 2 pasos** esté **Activada** (es un requisito obligatorio de Google).
4. En esa misma página (o usando la barra de búsqueda superior), busca **"Contraseñas de aplicaciones"** (App Passwords).
5. Crea una nueva contraseña. Ponle un nombre para que la recuerdes (ej: *"App Descarga Correos"*).
6. Google te mostrará una ventana emergente con un código de **16 letras**.
7. **¡Copia ese código!** Esa es la contraseña que debes pegar en el campo "App Password" de nuestra aplicación. *(Puedes pegarla con o sin espacios, la app la detectará igual).*

---

## 🚀 Cómo Ejecutar la Aplicación

Tienes tres formas de utilizar este proyecto, desde la más fácil hasta el entorno de desarrollo:

### 1. El modo portable (Recomendado para usuarios finales)
Si alguien ya compiló la aplicación:
1. Ve a la carpeta `dist/`.
2. Haz doble clic en **`Gmail_Downloader.exe`**.
3. ¡Listo! La aplicación se abrirá al instante sin requerir instalaciones previas.

### 2. El modo Script (Windows Automático)
Si tienes el código fuente pero no quieres lidiar con consolas:
1. Haz doble clic en el archivo **`Arrancar_App.bat`**.
2. El script revisará si tienes Python. Si no lo tienes, te ofrecerá descargarlo e instalarlo de forma silenciosa y automática.
3. Luego instalará las librerías necesarias (`xhtml2pdf`) y abrirá la aplicación sin mostrar ventanas negras.

### 3. El modo Desarrollador (Manual)
Si eres desarrollador y quieres modificar el código:
```bash
# 1. Clona el repositorio
git clone https://github.com/REBAN-Desarrollo/Email-Download-Reban.git

# 2. Entra a la carpeta
cd Email-Download-Reban

# 3. Instala los requerimientos
pip install -r requirements.txt

# 4. Ejecuta el script principal
python app.py
```

---

## 🛠️ Cómo Compilar el proyecto a `.exe`

Si modificaste el código y quieres generar un nuevo ejecutable portable para tus compañeros:

1. Simplemente haz doble clic en el archivo **`Compilar_App.bat`**.
2. Este script instalará `PyInstaller`, empaquetará el código y sus dependencias, y limpiará los archivos temporales.
3. Al finalizar, encontrarás tu nuevo archivo en la carpeta `dist/Gmail_Downloader.exe`.

---

## ❓ FAQ & Solución de Problemas (Troubleshooting)

### ❌ Error: "No se pudo conectar" o "Login fallido"
- **Causa más común:** Estás usando tu contraseña normal en lugar de la *Contraseña de Aplicación*, o tu empresa tiene bloqueado el protocolo IMAP.
- **Solución:** Sigue los pasos de la sección de arriba para generar la contraseña de 16 caracteres. Si es un correo corporativo (Google Workspace), asegúrate de que el administrador de IT permita conexiones IMAP.

### ❌ Error: La aplicación no abre o el `.bat` se cierra de inmediato
- **Causa más común:** Python no se instaló correctamente en las variables de entorno (PATH) de Windows.
- **Solución:** Desinstala Python desde el Panel de Control. Vuelve a instalarlo desde la página oficial, y en la primera pantalla del instalador, **asegúrate de marcar la casilla inferior que dice "Add python.exe to PATH"**.

### ❌ Los acentos o la letra "ñ" salen extraños en el PDF
- La librería `xhtml2pdf` intenta usar la codificación `utf-8` por defecto, pero algunos correos viejos vienen en formatos ISO raros. La aplicación hace un esfuerzo por decodificarlos (`errors="ignore"`), pero si el correo original está muy corrupto, pueden perderse tildes.

### ❌ ¿Por qué no aparecen correos cuando busco por fecha?
- El servidor IMAP de Gmail requiere que las fechas estén en estricto formato en inglés: `DD-MMM-YYYY`. 
- **Ejemplo correcto:** `01-Jan-2024` o `15-Mar-2023`.
- **Ejemplo incorrecto:** `01/01/2024` o `15-Marzo-2023` (No usar meses en español).

---

## 💻 Notas Técnicas (Bajo el capó)

- **Prevención de lectura (BODY.PEEK):** Durante la previsualización, la app descarga solo los encabezados usando `BODY.PEEK`. Esto garantiza que los correos que aparecen en la tabla no pierdan su estado de "No Leídos" en el servidor de Gmail.
- **Hilos Paralelos (Threading):** La interfaz utiliza `threading` del sistema operativo. Cuando descargas 100 correos a la vez, el proceso de descarga ocurre en segundo plano (Daemon Thread), permitiendo que la ventana gráfica (Main Thread) siga respondiendo y actualizando el Log.
- **Librerías principales:** `tkinter` (Nativa de Python), `imaplib` (Nativa), `email` (Nativa), `xhtml2pdf` (Externa, para renderizado PDF).
