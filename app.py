import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import imaplib
import email
from email.header import decode_header
from email.utils import parsedate_to_datetime
import os
import sys
import re
import html  # para escapar HTML en PDFs
import base64
import json
import subprocess
import glob
from datetime import date

def find_ghostscript():
    """Busca gswin64c.exe en rutas estándar de Windows.
    Retorna la ruta completa o None."""
    # 1. Buscar en PATH del sistema
    for cmd in ("gswin64c", "gswin32c", "gs"):
        for d in os.environ.get("PATH", "").split(";"):
            exe = os.path.join(d, cmd + ".exe")
            if os.path.isfile(exe):
                return exe
    # 2. Buscar en rutas de instalación comunes
    search_roots = [
        os.path.join(
            os.environ.get("ProgramFiles", "C:\\Program Files"),
            "gs"
        ),
        os.path.join(
            os.environ.get(
                "ProgramFiles(x86)",
                "C:\\Program Files (x86)"
            ), "gs"
        ),
    ]
    for root in search_roots:
        pattern = os.path.join(root, "gs*", "bin", "gswin*c.exe")
        matches = glob.glob(pattern)
        if matches:
            return matches[0]
    return None

# Playwright portable: buscar Chromium en carpeta local del proyecto
os.environ["PLAYWRIGHT_BROWSERS_PATH"] = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "browsers"
)

try:
    from playwright.sync_api import sync_playwright
    PLAYWRIGHT_AVAILABLE = True
except ImportError:
    PLAYWRIGHT_AVAILABLE = False


def ensure_chromium_installed():
    """Verifica que Chromium exista en browsers/. Si no, lo descarga."""
    if not PLAYWRIGHT_AVAILABLE:
        return
    browsers_dir = os.environ.get("PLAYWRIGHT_BROWSERS_PATH", "")
    if not browsers_dir:
        return
    # Buscar si existe alguna carpeta chromium_headless_shell-*
    if os.path.isdir(browsers_dir):
        for name in os.listdir(browsers_dir):
            if name.startswith("chromium_headless_shell-"):
                # Verificar que el ejecutable exista dentro
                exe = os.path.join(
                    browsers_dir, name,
                    "chrome-win", "headless_shell.exe"
                )
                if os.path.isfile(exe):
                    return  # Todo en orden
    # No existe o está incompleto — descargar
    import subprocess
    subprocess.run(
        [sys.executable, "-m", "playwright",
         "install", "chromium"],
        check=True
    )
    # Limpiar browser completo innecesario (solo headless shell)
    if os.path.isdir(browsers_dir):
        for name in os.listdir(browsers_dir):
            if name.startswith("chromium-"):
                import shutil
                shutil.rmtree(
                    os.path.join(browsers_dir, name),
                    ignore_errors=True
                )

def format_size(size_bytes):
    """Formatea bytes a cadena legible (B, KB, MB)"""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    else:
        return f"{size_bytes / (1024 * 1024):.2f} MB"

def clean_filename(name):
    """Limpia caracteres inválidos para nombres de archivos en Windows"""
    # Eliminar caracteres de control (\r, \n, \t, etc.) y
    # caracteres prohibidos en Windows
    cleaned = re.sub(r'[\x00-\x1f\\/*?:"<>|]', "", name)
    # Colapsar espacios múltiples que queden tras la limpieza
    cleaned = re.sub(r'\s+', ' ', cleaned)
    return cleaned.strip()[:100]

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

GMAIL_DAILY_LIMIT = 2500 * 1024 * 1024  # 2500 MB
APP_DIR = os.path.dirname(os.path.abspath(__file__))
QUOTA_FILE = os.path.join(os.path.expanduser("~"), ".gmail_downloader_quota.json")
SETTINGS_FILE = os.path.join(APP_DIR, "settings.json")


def load_settings():
    """Carga los últimos settings guardados."""
    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}

def save_settings(data):
    """Guarda settings a disco."""
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def load_daily_quota():
    """Carga la cuota diaria. Resetea si es un día nuevo."""
    today = date.today().isoformat()
    try:
        with open(QUOTA_FILE, "r") as f:
            data = json.load(f)
        if data.get("date") == today:
            return data.get("bytes", 0)
    except (FileNotFoundError, json.JSONDecodeError, KeyError):
        pass
    return 0

def save_daily_quota(total_bytes):
    """Guarda la cuota acumulada del día."""
    today = date.today().isoformat()
    with open(QUOTA_FILE, "w") as f:
        json.dump({"date": today, "bytes": total_bytes}, f)

class GmailDownloaderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gmail Downloader Automático")
        self.root.geometry("900x700")
        self.root.resizable(False, False)

        self.imap_server = "imap.gmail.com"
        self.mail_conn = None
        self.output_dir = os.path.join(os.path.expanduser("~"), "Descargas_Correos")

        self.download_body = tk.BooleanVar(value=True)
        self.download_attachments = tk.BooleanVar(value=True)
        self.size_cache = {}  # iid → bytes raw para cálculo de tamaño
        self.daily_bytes = load_daily_quota()

        self.create_widgets()
        self._load_settings()

        if self.daily_bytes > 0:
            pct = self.daily_bytes / GMAIL_DAILY_LIMIT * 100
            self.log(f"[CUOTA] Uso hoy: {format_size(self.daily_bytes)} de 2500 MB ({pct:.1f}%)")

        if not PLAYWRIGHT_AVAILABLE:
            self.log("[AVISO] playwright no está instalado. No se podrán generar PDFs.")
            self.log("  Instalar: pip install playwright && playwright install chromium")
        elif PLAYWRIGHT_AVAILABLE:
            # Auto-verificar/descargar Chromium al iniciar
            try:
                ensure_chromium_installed()
            except Exception as e:
                self.log(f"[AVISO] No se pudo verificar Chromium: {e}")

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

        # Toolbar: Seleccionar/Deseleccionar + Contador
        frame_toolbar = tk.Frame(frame_results)
        frame_toolbar.pack(fill="x", padx=5, pady=(5, 0))

        ttk.Button(frame_toolbar, text="Seleccionar Todos", command=self.select_all).pack(side="left", padx=2)
        ttk.Button(frame_toolbar, text="Deseleccionar", command=self.deselect_all).pack(side="left", padx=2)
        self.selection_label = ttk.Label(frame_toolbar, text="0 de 0 seleccionados")
        self.selection_label.pack(side="right", padx=5)

        columns = ("id", "fecha", "remitente", "asunto", "tamano")
        self.tree = ttk.Treeview(frame_results, columns=columns, show="headings", selectmode="extended")
        self.tree.heading("id", text="ID")
        self.tree.column("id", width=50, anchor="center")
        self.tree.heading("fecha", text="Fecha")
        self.tree.column("fecha", width=100, anchor="center")
        self.tree.heading("remitente", text="Remitente")
        self.tree.column("remitente", width=200)
        self.tree.heading("asunto", text="Asunto")
        self.tree.column("asunto", width=350)
        self.tree.heading("tamano", text="Tamaño")
        self.tree.column("tamano", width=80, anchor="center")

        scrollbar = ttk.Scrollbar(frame_results, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=(5,0), pady=5)
        scrollbar.pack(side="right", fill="y", padx=(0,5), pady=5)

        self.tree.bind("<<TreeviewSelect>>", self.update_selection_count)

        # --- Barra de Progreso ---
        self.progress = ttk.Progressbar(self.root, mode="determinate")

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

        ttk.Checkbutton(frame_action, text="Cuerpo (PDF)", variable=self.download_body).pack(side="left", padx=5)
        ttk.Checkbutton(frame_action, text="Adjuntos", variable=self.download_attachments).pack(side="left", padx=5)

        self.btn_download = ttk.Button(frame_action, text="Descargar Seleccionados (Batch)", command=self.start_download)
        self.btn_download.pack(side="right", padx=5)

        # Botón para comprimir PDFs existentes
        self.btn_compress = ttk.Button(frame_action, text="Comprimir PDFs en Carpeta", command=self.start_compress_pdfs)
        self.btn_compress.pack(side="right", padx=5)

        # --- Log Text ---
        self.log_text = tk.Text(self.root, height=8, state='disabled', bg="#f4f4f4")
        self.log_text.pack(fill="x", padx=10, pady=(0,10))

    def _insert_email_row(self, eid, fecha, remitente, asunto, size_str, size_bytes):
        iid = self.tree.insert("", "end", values=(eid.decode(), fecha, remitente, asunto, size_str))
        self.size_cache[iid] = size_bytes

    def _load_settings(self):
        """Restaura los últimos settings guardados en los campos."""
        s = load_settings()
        if not s:
            return
        # Credenciales
        if s.get("email"):
            self.email_entry.insert(0, s["email"])
        if s.get("password"):
            try:
                self.pass_entry.insert(0, base64.b64decode(s["password"]).decode("utf-8"))
            except Exception:
                pass
        # Filtros — limpiar defaults y poner los guardados
        for field, key in [(self.sender_entry, "sender"), (self.since_entry, "since"),
                           (self.before_entry, "before"), (self.subject_entry, "subject"),
                           (self.keyword_entry, "keyword"), (self.format_entry, "format")]:
            if key in s:
                field.delete(0, tk.END)
                field.insert(0, s[key])
        # Carpeta destino
        if s.get("output_dir"):
            self.output_dir = s["output_dir"]
            self.folder_label.config(text=self.output_dir)
        # Checkboxes
        if "download_body" in s:
            self.download_body.set(s["download_body"])
        if "download_attachments" in s:
            self.download_attachments.set(s["download_attachments"])

    def _save_settings(self):
        """Guarda el estado actual de todos los campos."""
        save_settings({
            "email": self.email_entry.get().strip(),
            "password": base64.b64encode(self.pass_entry.get().strip().encode("utf-8")).decode("ascii"),
            "sender": self.sender_entry.get().strip(),
            "since": self.since_entry.get().strip(),
            "before": self.before_entry.get().strip(),
            "subject": self.subject_entry.get().strip(),
            "keyword": self.keyword_entry.get().strip(),
            "format": self.format_entry.get().strip(),
            "output_dir": self.output_dir,
            "download_body": self.download_body.get(),
            "download_attachments": self.download_attachments.get(),
        })

    def select_all(self):
        children = self.tree.get_children()
        if children:
            self.tree.selection_set(children)

    def deselect_all(self):
        self.tree.selection_remove(self.tree.get_children())

    def update_selection_count(self, event=None):
        total = len(self.tree.get_children())
        selected_items = self.tree.selection()
        selected = len(selected_items)

        total_bytes = sum(self.size_cache.get(item, 0) for item in selected_items)

        size_display = f" | {format_size(total_bytes)}" if selected > 0 and total_bytes > 0 else ""
        self.selection_label.config(text=f"{selected} de {total} seleccionados{size_display}")

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
        self._save_settings()
        self.btn_search.config(state="disabled")
        self.tree.delete(*self.tree.get_children())
        self.size_cache.clear()
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
                # PEEK previene marcar los correos como leídos y pedimos el tamaño total (RFC822.SIZE)
                _, msg_data = self.mail_conn.fetch(eid, "(RFC822.SIZE BODY.PEEK[HEADER.FIELDS (DATE FROM SUBJECT)])")
                if not msg_data or msg_data[0] is None:
                    continue

                # msg_data puede ser [(b'flags', b'header'), b')'] — tomar el primer elemento tupla
                raw_part = msg_data[0]
                size_bytes = 0
                size_str = "Desconocido"
                if isinstance(raw_part, tuple):
                    # Extraer el tamaño del correo
                    info_str = raw_part[0].decode("utf-8", errors="ignore")
                    match = re.search(r"RFC822\.SIZE (\d+)", info_str, re.IGNORECASE)
                    if match:
                        size_bytes = int(match.group(1))
                        size_str = format_size(size_bytes)

                    raw_header = raw_part[1]
                else:
                    continue
                msg = email.message_from_bytes(raw_header)

                try:
                    fecha_dt = parsedate_to_datetime(msg.get("Date", ""))
                    fecha_str = fecha_dt.strftime("%Y-%m-%d_%H%M")
                except (ValueError, TypeError):
                    fecha_str = "Fecha_Desconocida"

                remitente = decode_str(msg.get("From", ""))
                asunto = decode_str(msg.get("Subject", ""))

                self.root.after(0, lambda e=eid, f=fecha_str, r=remitente, a=asunto, s=size_str, sb=size_bytes: self._insert_email_row(e, f, r, a, s, sb))

            self.log("Búsqueda completada.")
        except Exception as e:
            self.log(f"Error: {str(e)}")
        finally:
            self.root.after(0, self.btn_search.config, {"state": "normal"})
            self._close_imap()

    def start_download(self):
        self._save_settings()
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("Atención", "Debes seleccionar al menos un correo de la tabla.")
            return

        if not self.download_body.get() and not self.download_attachments.get():
            messagebox.showwarning("Atención", "Debes marcar al menos una opción: Cuerpo (PDF) o Adjuntos.")
            return

        items_data = [self.tree.item(item, "values") for item in selected_items]
        self.btn_download.config(state="disabled")

        # Mostrar barra de progreso
        self.progress.config(maximum=len(items_data), value=0)
        self.progress.pack(fill="x", padx=10, pady=(0, 5))

        threading.Thread(target=self.download_emails, args=(items_data,), daemon=True).start()

    def convert_html_to_pdf(self, page, source_html, output_filename):
        try:
            page.set_content(source_html, wait_until="networkidle", timeout=15000)
            page.pdf(path=output_filename, format="Letter",
                     margin={"top": "15mm", "bottom": "15mm",
                             "left": "10mm", "right": "10mm"})
            # Auto-comprimir el PDF recién generado
            gs = find_ghostscript()
            if gs:
                self._compress_single_pdf(output_filename, gs)
            return True
        except Exception as e:
            self.log(f"Error convirtiendo PDF: {e}")
            return False

    def download_emails(self, items_data):
        want_body = self.download_body.get()
        want_attachments = self.download_attachments.get()

        pw_context = None
        browser = None
        page = None

        try:
            os.makedirs(self.output_dir, exist_ok=True)
            self.connect_imap()

            # Abrir Playwright una sola vez para todo el batch
            if want_body and PLAYWRIGHT_AVAILABLE:
                self.log("Iniciando motor de renderizado PDF...")
                pw_context = sync_playwright().start()
                browser = pw_context.chromium.launch(headless=True)
                page = browser.new_page()

            template = self.format_entry.get().strip()
            if not template:
                template = "{date}_{subject}"

            batch_bytes = 0

            for idx, item in enumerate(items_data, 1):
                eid_str, fecha, remitente, asunto = item[0], item[1], item[2], item[3]
                self.log(f"Descargando [{idx}/{len(items_data)}]: {asunto[:30]}...")

                _, msg_data = self.mail_conn.fetch(eid_str.encode(), "(BODY.PEEK[])")
                raw_email = msg_data[0][1]
                batch_bytes += len(raw_email)
                msg = email.message_from_bytes(raw_email)

                # Procesar Wildcards / Template para la carpeta
                folder_name = template
                folder_name = folder_name.replace("{date}", fecha)
                folder_name = folder_name.replace("{sender}", clean_filename(remitente))
                folder_name = folder_name.replace("{subject}", clean_filename(asunto))
                folder_name = folder_name.replace("{id}", eid_str)
                folder_name = clean_filename(folder_name)
                folder_name = folder_name.replace("..", "")

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
                cid_map = {}  # Content-ID → data URI (para imágenes inline)

                # Procesar partes del correo (Adjuntos y Cuerpo)
                for part in msg.walk():
                    ctype = part.get_content_type()
                    cdispo = str(part.get("Content-Disposition", ""))

                    # Descarga de Adjuntos
                    if "attachment" in cdispo or part.get_filename():
                        if want_attachments:
                            fname = part.get_filename()
                            if fname:
                                fname_dec = decode_str(fname)
                                safe_fname = clean_filename(fname_dec)
                                filepath = os.path.join(folder_path, safe_fname)
                                with open(filepath, "wb") as f:
                                    f.write(part.get_payload(decode=True))
                                self.log(f"  Adjunto: {safe_fname}")
                        continue

                    # Imágenes inline (Content-ID) → guardar como data URI para PDF
                    if want_body and ctype.startswith("image/"):
                        content_id = part.get("Content-ID", "")
                        if content_id:
                            cid = content_id.strip("<>")
                            payload = part.get_payload(decode=True)
                            if payload:
                                b64 = base64.b64encode(payload).decode("ascii")
                                cid_map[cid] = f"data:{ctype};base64,{b64}"

                    # Extracción de Texto para el PDF
                    if ctype == "text/html" and not body_html:
                        body_html = part.get_payload(decode=True).decode("utf-8", errors="ignore")
                    elif ctype == "text/plain" and not body_text:
                        body_text = part.get_payload(decode=True).decode("utf-8", errors="ignore")

                # Reemplazar referencias cid: en el HTML con data URIs
                if body_html and cid_map:
                    for cid, data_uri in cid_map.items():
                        body_html = body_html.replace(f"cid:{cid}", data_uri)

                if want_body and page:
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
                    if self.convert_html_to_pdf(page, pdf_html, pdf_path):
                        self.log("  PDF guardado en Mensaje_Legible.pdf")
                    else:
                        self.log("  [AVISO] No se pudo generar el PDF")
                elif want_body and not PLAYWRIGHT_AVAILABLE:
                    self.log("  [AVISO] playwright no instalado, PDF omitido")

                # Actualizar barra de progreso
                self.root.after(0, self.progress.config, {"value": idx})

            # Actualizar cuota diaria
            self.daily_bytes += batch_bytes
            save_daily_quota(self.daily_bytes)
            pct = self.daily_bytes / GMAIL_DAILY_LIMIT * 100
            self.log(f"Descarga masiva completada exitosamente!")
            self.log(f"[CUOTA] Este batch: {format_size(batch_bytes)} | Hoy: {format_size(self.daily_bytes)} de 2500 MB ({pct:.1f}%)")
            if pct >= 80:
                self.log("[CUOTA] ADVERTENCIA: Vas por encima del 80% del límite diario de Gmail.")
            self.root.after(0, self._download_finished, len(items_data))

        except Exception as e:
            self.log(f"Error crítico durante descarga: {str(e)}")
            self.root.after(0, messagebox.showerror, "Error", f"Ocurrió un error:\n{str(e)}")
        finally:
            # Cerrar Playwright
            if browser:
                try:
                    browser.close()
                except Exception:
                    pass
            if pw_context:
                try:
                    pw_context.stop()
                except Exception:
                    pass
            self.root.after(0, self.btn_download.config, {"state": "normal"})
            self.root.after(0, self.progress.pack_forget)
            self._close_imap()

    def _compress_single_pdf(self, pdf_path, gs_path=None):
        """Comprime un PDF con Ghostscript (calidad /printer, 300 DPI).
        Retorna (original_bytes, compressed_bytes, error_str)."""
        if not gs_path:
            gs_path = find_ghostscript()
        if not gs_path:
            return 0, 0, "Ghostscript no encontrado"
        try:
            original_size = os.path.getsize(pdf_path)
            tmp_path = pdf_path + ".gs.tmp"

            result = subprocess.run(
                [
                    gs_path,
                    "-sDEVICE=pdfwrite",
                    "-dCompatibilityLevel=1.4",
                    "-dPDFSETTINGS=/printer",
                    "-dNOPAUSE",
                    "-dBATCH",
                    "-dQUIET",
                    f"-sOutputFile={tmp_path}",
                    pdf_path,
                ],
                capture_output=True,
                text=True,
                timeout=60,
            )

            if result.returncode != 0:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)
                err_msg = result.stderr.strip()[:200]
                return 0, 0, err_msg or "Ghostscript falló"

            new_size = os.path.getsize(tmp_path)
            if new_size < original_size:
                os.replace(tmp_path, pdf_path)
                return original_size, new_size, None
            else:
                os.remove(tmp_path)
                return original_size, original_size, None
        except Exception as e:
            tmp_path = pdf_path + ".gs.tmp"
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
            return 0, 0, str(e)

    def start_compress_pdfs(self):
        """Inicia la compresión de PDFs en hilo separado."""
        gs = find_ghostscript()
        if not gs:
            messagebox.showwarning(
                "Ghostscript no encontrado",
                "Instala Ghostscript desde:\n"
                "https://ghostscript.com/releases/gsdnld.html\n\n"
                "Reinicia la app después de instalar."
            )
            return
        if not os.path.isdir(self.output_dir):
            messagebox.showwarning(
                "Carpeta no encontrada",
                f"La carpeta no existe:\n{self.output_dir}"
            )
            return
        self.btn_compress.config(state="disabled")
        threading.Thread(
            target=self._compress_pdfs_in_folder,
            daemon=True
        ).start()

    def _compress_pdfs_in_folder(self):
        """Escanea carpeta de output y comprime todos los PDFs."""
        try:
            # Encontrar todos los PDFs recursivamente
            pdf_files = []
            for dirpath, _, filenames in os.walk(self.output_dir):
                for f in filenames:
                    if f.lower().endswith(".pdf"):
                        pdf_files.append(os.path.join(dirpath, f))

            if not pdf_files:
                self.log("[PDF] No se encontraron archivos PDF.")
                return

            self.log(f"")
            self.log(f"[PDF] Comprimiendo {len(pdf_files)} PDF(s)...")
            total_original = 0
            total_compressed = 0
            compressed_count = 0

            error_count = 0

            for i, pdf_path in enumerate(pdf_files, 1):
                rel_path = os.path.relpath(pdf_path, self.output_dir)
                orig, comp, err = self._compress_single_pdf(pdf_path)
                if err:
                    error_count += 1
                    self.log(
                        f"  [{i}/{len(pdf_files)}] {rel_path}: "
                        f"ERROR - {err}"
                    )
                    continue
                total_original += orig
                total_compressed += comp
                if orig > comp:
                    compressed_count += 1
                    savings = (1 - comp / orig) * 100
                    self.log(
                        f"  [{i}/{len(pdf_files)}] {rel_path}: "
                        f"{format_size(orig)} → "
                        f"{format_size(comp)} "
                        f"({savings:.0f}%)"
                    )
                else:
                    self.log(
                        f"  [{i}/{len(pdf_files)}] {rel_path}: "
                        f"ya optimizado ({format_size(orig)})"
                    )

            # Resumen final
            processed = len(pdf_files) - error_count
            if total_original > 0:
                total_savings = (
                    (1 - total_compressed / total_original) * 100
                )
                self.log(
                    f"[PDF] Listo: {compressed_count}/"
                    f"{processed} comprimidos | "
                    f"{format_size(total_original)} → "
                    f"{format_size(total_compressed)} "
                    f"(ahorro total {total_savings:.0f}%)"
                )
            elif processed > 0:
                self.log(
                    f"[PDF] {processed} PDF(s) ya estaban "
                    f"optimizados."
                )
            if error_count > 0:
                self.log(
                    f"[PDF] {error_count} PDF(s) no se "
                    f"pudieron procesar (ver errores arriba)."
                )

        except Exception as e:
            self.log(f"[PDF] Error: {e}")
        finally:
            self.root.after(
                0, self.btn_compress.config, {"state": "normal"}
            )

    def _download_finished(self, count):
        if messagebox.askyesno("Éxito", f"Se han descargado {count} correos exitosamente.\n¿Abrir carpeta de destino?"):
            os.startfile(self.output_dir)

if __name__ == "__main__":
    root = tk.Tk()
    app = GmailDownloaderApp(root)
    root.mainloop()
