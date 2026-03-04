import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import imaplib
import email
from email.header import decode_header
from email.utils import parsedate_to_datetime
import os
import re
import html  # para escapar HTML en PDFs

try:
    from xhtml2pdf import pisa
    XHTML2PDF_AVAILABLE = True
except ImportError:
    XHTML2PDF_AVAILABLE = False
    pisa = None

def clean_filename(name):
    """Limpia caracteres inválidos para nombres de archivos en Windows"""
    return re.sub(r'[\\/*?:"<>|]', "", name).strip()[:100]

def decode_str(header_str):
    """Decodifica un header de email (ej: Asunto o Remitente)"""
    if not header_str:
        return "Desconocido"
    decoded_parts = decode_header(header_str)
    result = ""
    for part, enc in decoded_parts:
        if isinstance(part, bytes):
            result += part.decode(enc or "utf-8", errors="ignore")
        else:
            result += part
    return result

def validate_imap_date(date_str):
    """Valida que la fecha tenga formato DD-MMM-YYYY válido para IMAP"""
    valid_months = {"Jan","Feb","Mar","Apr","May","Jun",
                    "Jul","Aug","Sep","Oct","Nov","Dec"}
    parts = date_str.split("-")
    if len(parts) != 3:
        return False
    day, month, year = parts
    if not day.isdigit() or not year.isdigit():
        return False
    if month not in valid_months:
        return False
    if not (1 <= int(day) <= 31) or not (1900 <= int(year) <= 2100):
        return False
    return True

class GmailDownloaderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gmail Downloader Automático")
        self.root.geometry("900x700")
        self.root.resizable(False, False)

        self.imap_server = "imap.gmail.com"
        self.mail_conn = None
        self.output_dir = os.path.join(os.path.expanduser("~"), "Descargas_Correos")

        self.create_widgets()

        if not XHTML2PDF_AVAILABLE:
            self.log("[AVISO] xhtml2pdf no está instalado. No se podrán generar PDFs.")

    def log(self, message):
        """Añade un mensaje al log en la interfaz gráfica (thread-safe)"""
        self.root.after(0, self._log_impl, message)

    def _log_impl(self, message):
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')

    def create_widgets(self):
        # --- Frame de Configuración de Cuenta ---
        frame_auth = ttk.LabelFrame(self.root, text="Credenciales de Gmail")
        frame_auth.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame_auth, text="Correo:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.email_entry = ttk.Entry(frame_auth, width=35)
        self.email_entry.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(frame_auth, text="App Password:").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        self.pass_entry = ttk.Entry(frame_auth, width=30, show="*")
        self.pass_entry.grid(row=0, column=3, padx=5, pady=5)

        self.btn_test = ttk.Button(frame_auth, text="Probar Conexión", command=self.start_test_connection)
        self.btn_test.grid(row=0, column=4, padx=5, pady=5)

        # --- Frame de Búsqueda Avanzada ---
        frame_search = ttk.LabelFrame(self.root, text="Filtros de Búsqueda")
        frame_search.pack(fill="x", padx=10, pady=5)

        # Fila 1
        ttk.Label(frame_search, text="Remitente:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.sender_entry = ttk.Entry(frame_search, width=25)
        self.sender_entry.insert(0, "tecnilaboral")
        self.sender_entry.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(frame_search, text="Desde (DD-MMM-YYYY):").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        self.since_entry = ttk.Entry(frame_search, width=15)
        self.since_entry.insert(0, "01-Jan-2025")
        self.since_entry.grid(row=0, column=3, padx=5, pady=5)

        ttk.Label(frame_search, text="Hasta:").grid(row=0, column=4, padx=5, pady=5, sticky="e")
        self.before_entry = ttk.Entry(frame_search, width=15)
        self.before_entry.insert(0, "31-Dec-2026")
        self.before_entry.grid(row=0, column=5, padx=5, pady=5)

        # Fila 2
        ttk.Label(frame_search, text="Asunto contiene:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.subject_entry = ttk.Entry(frame_search, width=25)
        self.subject_entry.grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(frame_search, text="Palabra Clave (Texto):").grid(row=1, column=2, padx=5, pady=5, sticky="e")
        self.keyword_entry = ttk.Entry(frame_search, width=15)
        self.keyword_entry.grid(row=1, column=3, padx=5, pady=5)

        self.btn_search = ttk.Button(frame_search, text="Buscar Correos", command=self.start_search)
        self.btn_search.grid(row=1, column=4, columnspan=2, padx=10, pady=5, sticky="ew")

        # --- Frame de Resultados (Tabla) ---
        frame_results = ttk.LabelFrame(self.root, text="Correos Encontrados (Usa Ctrl o Shift para multi-selección)")
        frame_results.pack(fill="both", expand=True, padx=10, pady=5)

        columns = ("id", "fecha", "remitente", "asunto")
        self.tree = ttk.Treeview(frame_results, columns=columns, show="headings", selectmode="extended")
        self.tree.heading("id", text="ID")
        self.tree.column("id", width=50, anchor="center")
        self.tree.heading("fecha", text="Fecha")
        self.tree.column("fecha", width=100, anchor="center")
        self.tree.heading("remitente", text="Remitente")
        self.tree.column("remitente", width=200)
        self.tree.heading("asunto", text="Asunto")
        self.tree.column("asunto", width=450)

        scrollbar = ttk.Scrollbar(frame_results, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=(5,0), pady=5)
        scrollbar.pack(side="right", fill="y", padx=(0,5), pady=5)

        # --- Frame de Formato de Salida ---
        frame_format = ttk.LabelFrame(self.root, text="Formato de Salida y Descarga")
        frame_format.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame_format, text="Formato de Carpeta:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.format_entry = ttk.Entry(frame_format, width=40)
        self.format_entry.insert(0, "{date}_{subject}")
        self.format_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(frame_format, text="Variables permitidas: {date}, {sender}, {subject}, {id}", foreground="gray").grid(row=0, column=2, padx=5, pady=5, sticky="w")

        # Acciones de carpeta y descarga
        frame_action = tk.Frame(frame_format)
        frame_action.grid(row=1, column=0, columnspan=3, pady=5, sticky="ew")

        self.btn_folder = ttk.Button(frame_action, text="Cambiar Carpeta Destino", command=self.choose_folder)
        self.btn_folder.pack(side="left", padx=5)

        self.folder_label = ttk.Label(frame_action, text=self.output_dir, foreground="blue")
        self.folder_label.pack(side="left", padx=5)

        self.btn_download = ttk.Button(frame_action, text="Descargar Seleccionados (Batch)", command=self.start_download)
        self.btn_download.pack(side="right", padx=5)

        # --- Log Text ---
        self.log_text = tk.Text(self.root, height=8, state='disabled', bg="#f4f4f4")
        self.log_text.pack(fill="x", padx=10, pady=(0,10))

    def choose_folder(self):
        folder = filedialog.askdirectory(initialdir=self.output_dir)
        if folder:
            self.output_dir = folder
            self.folder_label.config(text=self.output_dir)

    def _close_imap(self):
        """Cierra la conexión IMAP de forma segura"""
        if self.mail_conn:
            try:
                self.mail_conn.logout()
            except Exception:
                pass
            self.mail_conn = None

    def connect_imap(self):
        email_addr = self.email_entry.get().strip()
        pwd = self.pass_entry.get().strip()
        if not email_addr or not pwd:
            raise ValueError("Falta el correo o el App Password.")

        self.log("Conectando a Gmail...")
        self.mail_conn = imaplib.IMAP4_SSL(self.imap_server, timeout=30)
        self.mail_conn.login(email_addr, pwd)
        self.mail_conn.select("INBOX")
        self.log("Conectado exitosamente.")

    def start_test_connection(self):
        """Prueba la conexión IMAP sin buscar correos"""
        self.root.after(0, self.btn_test.config, {"state": "disabled"})
        threading.Thread(target=self._test_connection, daemon=True).start()

    def _test_connection(self):
        try:
            email_addr = self.email_entry.get().strip()
            pwd = self.pass_entry.get().strip()
            if not email_addr or not pwd:
                self.log("[ERROR] Falta el correo o el App Password.")
                return

            self.log("Probando conexión...")
            conn = imaplib.IMAP4_SSL(self.imap_server, timeout=30)
            conn.login(email_addr, pwd)
            conn.logout()
            self.log("Conexión exitosa. Credenciales válidas.")
        except imaplib.IMAP4.error as e:
            self.log(f"[ERROR] Autenticación fallida: {e}")
        except Exception as e:
            self.log(f"[ERROR] No se pudo conectar: {e}")
        finally:
            self.root.after(0, self.btn_test.config, {"state": "normal"})

    def start_search(self):
        self.root.after(0, self.btn_search.config, {"state": "disabled"})
        self.tree.delete(*self.tree.get_children())
        threading.Thread(target=self.search_emails, daemon=True).start()

    def search_emails(self):
        try:
            self.connect_imap()
            sender = self.sender_entry.get().strip()
            since = self.since_entry.get().strip()
            before = self.before_entry.get().strip()
            subject_kw = self.subject_entry.get().strip()
            body_kw = self.keyword_entry.get().strip()

            # Validar formato de fechas
            if since and not validate_imap_date(since):
                self.log(f"[ERROR] Formato de fecha inválido: '{since}'. Use DD-MMM-YYYY (ej: 01-Jan-2025)")
                return
            if before and not validate_imap_date(before):
                self.log(f"[ERROR] Formato de fecha inválido: '{before}'. Use DD-MMM-YYYY (ej: 31-Dec-2026)")
                return

            criteria = []
            if sender: criteria.append(f'FROM "{sender}"')
            if since: criteria.append(f'SINCE "{since}"')
            if before: criteria.append(f'BEFORE "{before}"')
            if subject_kw: criteria.append(f'SUBJECT "{subject_kw}"')
            if body_kw: criteria.append(f'TEXT "{body_kw}"')

            search_query = f"({' '.join(criteria)})" if criteria else "ALL"
            self.log(f"Buscando con criterios IMAP: {search_query}")

            status, data = self.mail_conn.search(None, search_query)
            if status != "OK":
                self.log("Error en la búsqueda IMAP.")
                return

            ids = data[0].split()
            if not ids:
                self.log("No se encontraron correos con esos criterios.")
                return

            self.log(f"Se encontraron {len(ids)} correos. Obteniendo previsualización...")

            for eid in ids:
                # PEEK previene marcar los correos como leídos
                _, msg_data = self.mail_conn.fetch(eid, "(BODY.PEEK[HEADER.FIELDS (DATE FROM SUBJECT)])")
                if not msg_data or not msg_data[0]:
                    continue

                raw_header = msg_data[0][1]
                msg = email.message_from_bytes(raw_header)

                try:
                    fecha_dt = parsedate_to_datetime(msg.get("Date", ""))
                    fecha_str = fecha_dt.strftime("%Y-%m-%d_%H%M")
                except (ValueError, TypeError):
                    fecha_str = "Fecha_Desconocida"

                remitente = decode_str(msg.get("From", ""))
                asunto = decode_str(msg.get("Subject", ""))

                self.root.after(0, self.tree.insert, "", "end", None, {"values": (eid.decode(), fecha_str, remitente, asunto)})

            self.log("Búsqueda completada.")
        except Exception as e:
            self.log(f"Error: {str(e)}")
        finally:
            self.root.after(0, self.btn_search.config, {"state": "normal"})
            self._close_imap()

    def start_download(self):
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("Atención", "Debes seleccionar al menos un correo de la tabla.")
            return

        items_data = [self.tree.item(item, "values") for item in selected_items]
        self.root.after(0, self.btn_download.config, {"state": "disabled"})
        threading.Thread(target=self.download_emails, args=(items_data,), daemon=True).start()

    def convert_html_to_pdf(self, source_html, output_filename):
        if not XHTML2PDF_AVAILABLE:
            self.log("  [AVISO] xhtml2pdf no instalado, PDF omitido")
            return False
        try:
            with open(output_filename, "wb") as result_file:
                pisa_status = pisa.CreatePDF(source_html, dest=result_file)
            return not pisa_status.err
        except Exception as e:
            self.log(f"Error convirtiendo PDF: {e}")
            return False

    def download_emails(self, items_data):
        try:
            os.makedirs(self.output_dir, exist_ok=True)
            self.connect_imap()

            template = self.format_entry.get().strip()
            if not template:
                template = "{date}_{subject}"

            for idx, item in enumerate(items_data, 1):
                eid_str, fecha, remitente, asunto = item
                self.log(f"Descargando [{idx}/{len(items_data)}]: {asunto[:30]}...")

                _, msg_data = self.mail_conn.fetch(eid_str.encode(), "(BODY.PEEK[])")
                raw_email = msg_data[0][1]
                msg = email.message_from_bytes(raw_email)

                # Procesar Wildcards / Template para la carpeta
                folder_name = template
                folder_name = folder_name.replace("{date}", fecha)
                folder_name = folder_name.replace("{sender}", clean_filename(remitente))
                folder_name = folder_name.replace("{subject}", clean_filename(asunto))
                folder_name = folder_name.replace("{id}", eid_str)
                folder_name = clean_filename(folder_name)
                # Prevenir path traversal
                folder_name = folder_name.replace("..", "").replace("/", "").replace("\\", "")

                if not folder_name:
                    folder_name = f"Email_{eid_str}"

                folder_path = os.path.join(self.output_dir, folder_name)
                # Evitar carpetas duplicadas: añadir sufijo incremental
                base_path = folder_path
                counter = 1
                while os.path.exists(folder_path):
                    folder_path = f"{base_path}_{counter}"
                    counter += 1
                os.makedirs(folder_path, exist_ok=True)

                body_html = ""
                body_text = ""

                # Procesar partes del correo (Adjuntos y Cuerpo)
                for part in msg.walk():
                    ctype = part.get_content_type()
                    cdispo = str(part.get("Content-Disposition", ""))

                    # Descarga de Adjuntos
                    if "attachment" in cdispo or part.get_filename():
                        fname = part.get_filename()
                        if fname:
                            fname_dec = decode_str(fname)
                            safe_fname = clean_filename(fname_dec)
                            filepath = os.path.join(folder_path, safe_fname)
                            with open(filepath, "wb") as f:
                                f.write(part.get_payload(decode=True))
                            self.log(f"  Adjunto: {safe_fname}")
                        continue

                    # Extracción de Texto para el PDF
                    if ctype == "text/html" and not body_html:
                        body_html = part.get_payload(decode=True).decode("utf-8", errors="ignore")
                    elif ctype == "text/plain" and not body_text:
                        body_text = part.get_payload(decode=True).decode("utf-8", errors="ignore")

                if not body_html and not body_text:
                    final_html = "<p><em>Este correo no tiene contenido de texto.</em></p>"
                else:
                    final_html = body_html or f"<pre>{html.escape(body_text)}</pre>"

                # Escapar metadatos para prevenir inyección HTML en el PDF
                safe_remitente = html.escape(remitente)
                safe_asunto = html.escape(asunto)
                safe_fecha = html.escape(fecha)

                pdf_html = f"""
                <html>
                <head><meta charset="utf-8">
                <style>
                    body {{ font-family: Arial, sans-serif; padding: 20px; }}
                    .header {{ border-bottom: 2px solid #ccc; padding-bottom: 10px; margin-bottom: 20px; }}
                    .header p {{ margin: 2px 0; }}
                </style>
                </head>
                <body>
                    <div class="header">
                        <p><strong>De:</strong> {safe_remitente}</p>
                        <p><strong>Asunto:</strong> {safe_asunto}</p>
                        <p><strong>Fecha:</strong> {safe_fecha}</p>
                    </div>
                    <div>{final_html}</div>
                </body>
                </html>
                """

                pdf_path = os.path.join(folder_path, "Mensaje_Legible.pdf")
                if self.convert_html_to_pdf(pdf_html, pdf_path):
                    self.log("  PDF guardado en Mensaje_Legible.pdf")
                else:
                    self.log("  [AVISO] No se pudo generar el PDF")

            self.log("Descarga masiva completada exitosamente!")
            self.root.after(0, messagebox.showinfo, "Éxito", f"Se han descargado {len(items_data)} correos exitosamente en:\n{self.output_dir}")

        except Exception as e:
            self.log(f"Error crítico durante descarga: {str(e)}")
            self.root.after(0, messagebox.showerror, "Error", f"Ocurrió un error:\n{str(e)}")
        finally:
            self.root.after(0, self.btn_download.config, {"state": "normal"})
            self._close_imap()

if __name__ == "__main__":
    root = tk.Tk()
    app = GmailDownloaderApp(root)
    root.mainloop()
