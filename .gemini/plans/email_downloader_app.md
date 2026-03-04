# Plan: Gmail Downloader App (GUI)

## Objective
Crear una aplicación de escritorio nativa y sencilla para Windows (usando Python y `tkinter`) que permita conectarse a Gmail, previsualizar correos filtrados por fecha/remitente, y descargarlos en lote. Los correos se guardarán en formato PDF y los adjuntos se guardarán organizados en carpetas con un formato de nombre estructurado.

## Scope & Impact
- **Interfaz (GUI):** `tkinter` (incluido en Python, ideal para Windows por su extrema simplicidad y nula necesidad de instalación adicional de librerías visuales).
- **Conversión a PDF:** Uso de `xhtml2pdf` para convertir el cuerpo del correo de HTML a PDF directamente en Python (sin requerir instalar programas de terceros en Windows).
- **Flujo de trabajo:**
  1. Conexión mediante IMAP (requiere Contraseña de Aplicación de Gmail).
  2. Filtro de búsqueda (Remitente, Fecha de inicio, Fecha de fin).
  3. Previsualización en una tabla interactiva (Fecha, Remitente, Asunto).
  4. Descarga en lote de los elementos en pantalla:
     - Crea carpeta principal y subcarpetas: `Descargas_Correos/YYYY-MM-DD - Asunto/`
     - Guarda el cuerpo como `Mensaje.pdf`
     - Guarda los archivos adjuntos en la misma carpeta manteniendo su nombre original.

## Implementation Steps

1. **Setup & Dependencias:**
   - Crear un archivo `requirements.txt` que incluya `xhtml2pdf` (para la conversión a PDF del cuerpo del correo).
2. **Desarrollo de la UI (`app.py`):**
   - **Panel de Configuración:** Campos para Correo, Contraseña de Aplicación, Remitente a buscar y rango de Fechas.
   - **Panel Central (Preview):** Una tabla (`Treeview` de Tkinter) para mostrar los correos encontrados (ID, Fecha, Remitente, Asunto).
   - **Panel Inferior:** Botones para "Buscar Correos", "Descargar Seleccionados" y una caja de texto simple para mostrar el progreso/estado de la herramienta.
3. **Lógica de Conexión y Búsqueda (IMAP):**
   - Uso de `threading` para evitar que la ventana de Windows se congele (el famoso "No responde") mientras descarga datos.
   - Búsqueda mediante criterios IMAP (`FROM`, `SINCE`, `BEFORE`) para traer rápidamente solo la información necesaria para el preview.
4. **Lógica de Descarga y Conversión a PDF:**
   - Descarga el correo completo (`RFC822`) de los IDs seleccionados.
   - Procesa la estructura MIME para extraer el texto/HTML del correo y los datos crudos de los adjuntos.
   - Transforma el texto/HTML a un archivo `.pdf` usando `xhtml2pdf`.
   - Limpia caracteres inválidos para los nombres de las carpetas en Windows.

## Verification & Testing
1. Ejecutar `pip install -r requirements.txt`.
2. Lanzar la aplicación con `python app.py`.
3. Validar la conexión con Gmail introduciendo un App Password.
4. Realizar una búsqueda de un remitente específico.
5. Seleccionar un par de correos y presionar "Descargar".
6. Comprobar en el explorador de Windows que se crearon las carpetas, el archivo PDF del correo se puede abrir y leer, y los adjuntos están intactos.