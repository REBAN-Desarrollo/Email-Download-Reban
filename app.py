import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import imaplib
import email
import time
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
import calendar
from datetime import date, datetime

def find_ghostscript():
    """Busca gswin64c.exe en rutas estandar de Windows.
    Retorna la ruta completa o None."""
    for cmd in ("gswin64c", "gswin32c", "gs"):
        for d in os.environ.get("PATH", "").split(";"):
            exe = os.path.join(d, cmd + ".exe")
            if os.path.isfile(exe):
                return exe
    search_roots = [
        os.path.join(
            os.environ.get("ProgramFiles", "C:\\Program Files"), "gs"
        ),
        os.path.join(
            os.environ.get("ProgramFiles(x86)", "C:\\Program Files (x86)"),
            "gs"
        ),
    ]
    for root in search_roots:
        pattern = os.path.join(root, "gs*", "bin", "gswin*c.exe")
        matches = glob.glob(pattern)
        if matches:
            return matches[0]
    return None

# Playwright portable
os.environ["PLAYWRIGHT_BROWSERS_PATH"] = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "browsers"
)

try:
    from playwright.sync_api import sync_playwright
    PLAYWRIGHT_AVAILABLE = True
except ImportError:
    PLAYWRIGHT_AVAILABLE = False


def ensure_chromium_installed():
    if not PLAYWRIGHT_AVAILABLE:
        return
    browsers_dir = os.environ.get("PLAYWRIGHT_BROWSERS_PATH", "")
    if not browsers_dir:
        return
    if os.path.isdir(browsers_dir):
        for name in os.listdir(browsers_dir):
            if name.startswith("chromium_headless_shell-"):
                exe = os.path.join(
                    browsers_dir, name,
                    "chrome-win", "headless_shell.exe"
                )
                if os.path.isfile(exe):
                    return
    import subprocess
    subprocess.run(
        [sys.executable, "-m", "playwright", "install", "chromium"],
        check=True
    )
    if os.path.isdir(browsers_dir):
        for name in os.listdir(browsers_dir):
            if name.startswith("chromium-"):
                import shutil
                shutil.rmtree(
                    os.path.join(browsers_dir, name), ignore_errors=True
                )

def format_size(size_bytes):
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    else:
        return f"{size_bytes / (1024 * 1024):.2f} MB"

def clean_filename(name):
    cleaned = re.sub(r'[\x00-\x1f\\/*?:"<>|]', "", name)
    cleaned = re.sub(r'\s+', ' ', cleaned)
    return cleaned.strip()[:100]

def decode_str(header_str):
    if not header_str:
        return ""
    decoded_parts = decode_header(header_str)
    result = ""
    for part, enc in decoded_parts:
        if isinstance(part, bytes):
            result += part.decode(enc or "utf-8", errors="ignore")
        else:
            result += part
    return result

def validate_imap_date(date_str):
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

GMAIL_DAILY_LIMIT = 2500 * 1024 * 1024
APP_DIR = os.path.dirname(os.path.abspath(__file__))
QUOTA_FILE = os.path.join(os.path.expanduser("~"), ".gmail_downloader_quota.json")
SETTINGS_FILE = os.path.join(APP_DIR, "settings.json")

MONTHS_ES = ["Ene","Feb","Mar","Abr","May","Jun",
             "Jul","Ago","Sep","Oct","Nov","Dic"]
MONTHS_IMAP = ["Jan","Feb","Mar","Apr","May","Jun",
               "Jul","Aug","Sep","Oct","Nov","Dec"]
DAYS_ES = ["Lu","Ma","Mi","Ju","Vi","Sa","Do"]


def load_settings():
    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}

def save_settings(data):
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def load_daily_quota():
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
    today = date.today().isoformat()
    with open(QUOTA_FILE, "w") as f:
        json.dump({"date": today, "bytes": total_bytes}, f)


# ---------------------------------------------------------------------------
# DatePicker — Selector de fecha con popup de calendario
# ---------------------------------------------------------------------------
class DatePicker(tk.Frame):
    """Widget de fecha con boton que despliega un calendario popup."""

    def __init__(self, parent, initial_date=None, **kwargs):
        super().__init__(parent, **kwargs)

        if initial_date:
            self._date = initial_date
        else:
            self._date = date.today()

        self._var = tk.StringVar(value=self._format_imap(self._date))

        self.entry = ttk.Entry(self, textvariable=self._var, width=14)
        self.entry.pack(side="left")
        self.btn = ttk.Button(self, text="...", width=3, command=self._open_popup)
        self.btn.pack(side="left", padx=(2, 0))

        self._popup = None

    def get(self):
        return self._var.get().strip()

    def set_date(self, d):
        self._date = d
        self._var.set(self._format_imap(d))

    def set_text(self, text):
        self._var.set(text)
        try:
            self._date = self._parse_imap(text)
        except Exception:
            pass

    def _format_imap(self, d):
        return f"{d.day:02d}-{MONTHS_IMAP[d.month - 1]}-{d.year}"

    def _parse_imap(self, s):
        parts = s.split("-")
        day = int(parts[0])
        month = MONTHS_IMAP.index(parts[1]) + 1
        year = int(parts[2])
        return date(year, month, day)

    def _open_popup(self):
        if self._popup and self._popup.winfo_exists():
            self._popup.destroy()
            self._popup = None
            return

        try:
            self._date = self._parse_imap(self._var.get().strip())
        except Exception:
            self._date = date.today()

        self._popup = _CalendarPopup(self, self._date, self._on_select)

    def _on_select(self, selected_date):
        self._date = selected_date
        self._var.set(self._format_imap(selected_date))
        if self._popup:
            self._popup.destroy()
            self._popup = None


class _CalendarPopup(tk.Toplevel):
    """Popup de calendario con navegacion mes/ano."""

    def __init__(self, parent, current_date, callback):
        super().__init__(parent)
        self.overrideredirect(True)
        self.callback = callback
        self.current_year = current_date.year
        self.current_month = current_date.month
        self.selected_date = current_date

        # Posicionar debajo del widget padre
        x = parent.winfo_rootx()
        y = parent.winfo_rooty() + parent.winfo_height() + 2
        self.geometry(f"+{x}+{y}")

        self.configure(bg="#ffffff", bd=1, relief="solid")

        # Frame de navegacion
        nav = tk.Frame(self, bg="#2c3e50")
        nav.pack(fill="x")

        btn_style = {"bg": "#2c3e50", "fg": "white", "bd": 0,
                     "font": ("Arial", 10, "bold"), "cursor": "hand2",
                     "activebackground": "#34495e", "activeforeground": "white"}

        tk.Button(nav, text="<<", command=self._prev_year, **btn_style).pack(side="left", padx=2)
        tk.Button(nav, text="<", command=self._prev_month, **btn_style).pack(side="left", padx=2)

        self.lbl_header = tk.Label(nav, text="", bg="#2c3e50", fg="white",
                                   font=("Arial", 11, "bold"), cursor="hand2")
        self.lbl_header.pack(side="left", expand=True, fill="x")
        self.lbl_header.bind("<Button-1>", self._show_month_year_selector)

        tk.Button(nav, text=">", command=self._next_month, **btn_style).pack(side="right", padx=2)
        tk.Button(nav, text=">>", command=self._next_year, **btn_style).pack(side="right", padx=2)

        # Frame para selector rapido mes/ano (oculto por defecto)
        self.quick_frame = tk.Frame(self, bg="#ecf0f1")
        self._quick_visible = False

        # Frame del calendario
        self.cal_frame = tk.Frame(self, bg="white")
        self.cal_frame.pack(padx=2, pady=2)

        # Boton "Hoy"
        tk.Button(self, text="Hoy", command=self._select_today,
                  bg="#3498db", fg="white", bd=0, cursor="hand2",
                  font=("Arial", 9), activebackground="#2980b9").pack(fill="x", padx=2, pady=(0, 2))

        self._build_calendar()

        # Cerrar al perder foco
        self.bind("<FocusOut>", self._on_focus_out)
        self.focus_set()
        self.grab_set()

    def _on_focus_out(self, event):
        # Solo cerrar si el foco sale completamente del popup
        try:
            focused = self.focus_get()
            if focused and (focused == self or str(focused).startswith(str(self))):
                return
        except Exception:
            pass
        self.destroy()

    def _build_calendar(self):
        for w in self.cal_frame.winfo_children():
            w.destroy()

        self.lbl_header.config(
            text=f"{MONTHS_ES[self.current_month - 1]} {self.current_year}"
        )

        # Dias de la semana
        for i, day_name in enumerate(DAYS_ES):
            tk.Label(self.cal_frame, text=day_name, width=4, bg="#ecf0f1",
                     fg="#2c3e50", font=("Arial", 9, "bold")).grid(row=0, column=i, padx=1, pady=1)

        # Dias del mes
        cal = calendar.monthcalendar(self.current_year, self.current_month)
        today = date.today()

        for row_idx, week in enumerate(cal, 1):
            for col_idx, day in enumerate(week):
                if day == 0:
                    tk.Label(self.cal_frame, text="", width=4, bg="white").grid(
                        row=row_idx, column=col_idx, padx=1, pady=1)
                else:
                    d = date(self.current_year, self.current_month, day)
                    bg = "white"
                    fg = "#2c3e50"
                    font = ("Arial", 9)

                    if d == today:
                        bg = "#3498db"
                        fg = "white"
                        font = ("Arial", 9, "bold")
                    if d == self.selected_date:
                        bg = "#e74c3c"
                        fg = "white"
                        font = ("Arial", 9, "bold")

                    btn = tk.Label(self.cal_frame, text=str(day), width=4,
                                   bg=bg, fg=fg, font=font, cursor="hand2",
                                   relief="flat")
                    btn.grid(row=row_idx, column=col_idx, padx=1, pady=1)
                    btn.bind("<Button-1>", lambda e, dd=d: self._on_day_click(dd))
                    btn.bind("<Enter>", lambda e, b=btn: b.config(bg="#bdc3c7") if b.cget("bg") == "white" else None)
                    btn.bind("<Leave>", lambda e, b=btn, bg_orig=bg: b.config(bg=bg_orig))

    def _on_day_click(self, d):
        self.callback(d)

    def _prev_month(self):
        if self.current_month == 1:
            self.current_month = 12
            self.current_year -= 1
        else:
            self.current_month -= 1
        self._build_calendar()

    def _next_month(self):
        if self.current_month == 12:
            self.current_month = 1
            self.current_year += 1
        else:
            self.current_month += 1
        self._build_calendar()

    def _prev_year(self):
        self.current_year -= 1
        self._build_calendar()

    def _next_year(self):
        self.current_year += 1
        self._build_calendar()

    def _select_today(self):
        self.callback(date.today())

    def _show_month_year_selector(self, event=None):
        if self._quick_visible:
            self.quick_frame.pack_forget()
            self._quick_visible = False
            return

        for w in self.quick_frame.winfo_children():
            w.destroy()

        # Selector de ano
        year_frame = tk.Frame(self.quick_frame, bg="#ecf0f1")
        year_frame.pack(fill="x", padx=5, pady=(5, 2))
        tk.Label(year_frame, text="Ano:", bg="#ecf0f1",
                 font=("Arial", 9, "bold")).pack(side="left")

        year_var = tk.IntVar(value=self.current_year)
        tk.Button(year_frame, text="-", width=2, command=lambda: self._quick_change_year(year_var, -1),
                  bg="#bdc3c7", bd=0).pack(side="left", padx=2)
        self.year_lbl = tk.Label(year_frame, textvariable=year_var, bg="#ecf0f1",
                                  font=("Arial", 10, "bold"), width=6)
        self.year_lbl.pack(side="left")
        tk.Button(year_frame, text="+", width=2, command=lambda: self._quick_change_year(year_var, 1),
                  bg="#bdc3c7", bd=0).pack(side="left", padx=2)
        self._year_var = year_var

        # Grid de meses
        month_frame = tk.Frame(self.quick_frame, bg="#ecf0f1")
        month_frame.pack(fill="x", padx=5, pady=(2, 5))

        for i, m_name in enumerate(MONTHS_ES):
            r, c = divmod(i, 4)
            bg = "#3498db" if (i + 1) == self.current_month else "#dfe6e9"
            fg = "white" if (i + 1) == self.current_month else "#2c3e50"
            btn = tk.Label(month_frame, text=m_name, width=5, bg=bg, fg=fg,
                           font=("Arial", 9), cursor="hand2", relief="flat")
            btn.grid(row=r, column=c, padx=2, pady=2)
            btn.bind("<Button-1>", lambda e, mi=i+1, yv=year_var: self._quick_select_month(mi, yv))

        self.quick_frame.pack(fill="x", before=self.cal_frame)
        self._quick_visible = True

    def _quick_change_year(self, year_var, delta):
        year_var.set(year_var.get() + delta)

    def _quick_select_month(self, month, year_var):
        self.current_month = month
        self.current_year = year_var.get()
        self.quick_frame.pack_forget()
        self._quick_visible = False
        self._build_calendar()


# ---------------------------------------------------------------------------
# App principal
# ---------------------------------------------------------------------------
class GmailDownloaderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gmail Downloader")
        self.root.geometry("1250x750")
        self.root.minsize(1050, 600)

        self.imap_server = "imap.gmail.com"
        self.mail_conn = None
        self.output_dir = os.path.join(os.path.expanduser("~"), "Descargas_Correos")

        self.download_body = tk.BooleanVar(value=True)
        self.download_attachments = tk.BooleanVar(value=True)
        self.has_attachments_filter = tk.BooleanVar(value=False)
        self.size_cache = {}
        self.daily_bytes = load_daily_quota()
        self._cancel_event = threading.Event()
        self._log_file = None
        self._preview_conn = None  # conexion IMAP separada para preview
        self._mailboxes = []  # lista de buzones disponibles
        self._email_cache = {}  # eid -> datos cacheados para preview

        self.create_widgets()
        self._load_settings()

        if self.daily_bytes > 0:
            pct = self.daily_bytes / GMAIL_DAILY_LIMIT * 100
            self.log(f"[CUOTA] Uso hoy: {format_size(self.daily_bytes)} de 2500 MB ({pct:.1f}%)")

        if not PLAYWRIGHT_AVAILABLE:
            self.log("[AVISO] playwright no esta instalado. No se podran generar PDFs.")
        elif PLAYWRIGHT_AVAILABLE:
            try:
                ensure_chromium_installed()
            except Exception as e:
                self.log(f"[AVISO] No se pudo verificar Chromium: {e}")

    def log(self, message):
        self.root.after(0, self._log_impl, message)

    def _log_impl(self, message):
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')
        if self._log_file:
            try:
                ts = datetime.now().strftime("%H:%M:%S")
                self._log_file.write(f"[{ts}] {message}\n")
                self._log_file.flush()
            except Exception:
                pass

    def create_widgets(self):
        # === HEADER BAR (credenciales inline) ===
        header = tk.Frame(self.root, bg="#2c3e50", height=45)
        header.pack(fill="x")
        header.pack_propagate(False)

        tk.Label(header, text="Gmail Downloader", bg="#2c3e50", fg="white",
                 font=("Segoe UI", 13, "bold")).pack(side="left", padx=15)

        tk.Label(header, text="Correo:", bg="#2c3e50", fg="#bdc3c7",
                 font=("Segoe UI", 9)).pack(side="left", padx=(20, 3))
        self.email_entry = ttk.Entry(header, width=24)
        self.email_entry.pack(side="left", pady=8)

        tk.Label(header, text="Password:", bg="#2c3e50", fg="#bdc3c7",
                 font=("Segoe UI", 9)).pack(side="left", padx=(12, 3))
        self.pass_entry = ttk.Entry(header, width=20, show="*")
        self.pass_entry.pack(side="left", pady=8)

        self.btn_test = ttk.Button(header, text="Conectar", command=self.start_test_connection)
        self.btn_test.pack(side="left", padx=8)

        # Buzon selector en el header
        tk.Label(header, text="Buzon:", bg="#2c3e50", fg="#bdc3c7",
                 font=("Segoe UI", 9)).pack(side="left", padx=(12, 3))
        self.mailbox_var = tk.StringVar(value="INBOX")
        self.mailbox_combo = ttk.Combobox(header, textvariable=self.mailbox_var,
                                           width=16, state="readonly")
        self.mailbox_combo["values"] = ["INBOX"]
        self.mailbox_combo.pack(side="left", pady=8)

        # === RIBBON DE BUSQUEDA ===
        ribbon = tk.Frame(self.root, bg="#dfe6e9")
        ribbon.pack(fill="x")

        # Fila 1: busqueda global
        r1 = tk.Frame(ribbon, bg="#dfe6e9")
        r1.pack(fill="x", padx=10, pady=(5, 2))

        tk.Label(r1, text="Buscar en todo:", bg="#dfe6e9",
                 font=("Segoe UI", 10, "bold")).pack(side="left")
        self.global_search_entry = ttk.Entry(r1, width=50, font=("Segoe UI", 10))
        self.global_search_entry.pack(side="left", padx=5, fill="x", expand=True)
        tk.Label(r1, text="(OR: De, Para, CC, BCC, Asunto, Texto)", bg="#dfe6e9",
                 fg="#7f8c8d", font=("Segoe UI", 8)).pack(side="left", padx=5)

        self.btn_search = ttk.Button(r1, text="Buscar Correos", command=self.start_search)
        self.btn_search.pack(side="right", padx=5)

        # Fila 2: filtros detallados (AND)
        r2 = tk.Frame(ribbon, bg="#dfe6e9")
        r2.pack(fill="x", padx=10, pady=(0, 5))

        for lbl, attr, w in [("De:", "sender_entry", 14), ("Para:", "to_entry", 14),
                              ("CC:", "cc_entry", 12), ("Asunto:", "subject_entry", 16),
                              ("Texto:", "keyword_entry", 14)]:
            tk.Label(r2, text=lbl, bg="#dfe6e9", font=("Segoe UI", 8)).pack(side="left", padx=(6, 1))
            entry = ttk.Entry(r2, width=w)
            entry.pack(side="left", padx=(0, 3))
            setattr(self, attr, entry)

        tk.Label(r2, text="Desde:", bg="#dfe6e9", font=("Segoe UI", 8)).pack(side="left", padx=(8, 1))
        self.since_picker = DatePicker(r2, initial_date=date(2025, 1, 1))
        self.since_picker.pack(side="left")

        tk.Label(r2, text="Hasta:", bg="#dfe6e9", font=("Segoe UI", 8)).pack(side="left", padx=(5, 1))
        self.before_picker = DatePicker(r2, initial_date=date(2026, 12, 31))
        self.before_picker.pack(side="left")

        ttk.Checkbutton(r2, text="Adj.", variable=self.has_attachments_filter).pack(side="left", padx=6)

        # === CONTENIDO PRINCIPAL: Split horizontal tabla | preview ===
        self.paned = ttk.PanedWindow(self.root, orient="horizontal")
        self.paned.pack(fill="both", expand=True, padx=5, pady=5)

        # --- Panel izquierdo: tabla ---
        left = ttk.Frame(self.paned)
        self.paned.add(left, weight=3)

        # Toolbar tabla
        frame_toolbar = tk.Frame(left)
        frame_toolbar.pack(fill="x", padx=2, pady=(2, 0))
        ttk.Button(frame_toolbar, text="Sel. Todos", command=self.select_all).pack(side="left", padx=2)
        ttk.Button(frame_toolbar, text="Deseleccionar", command=self.deselect_all).pack(side="left", padx=2)
        self.selection_label = ttk.Label(frame_toolbar, text="0 de 0 seleccionados")
        self.selection_label.pack(side="right", padx=5)

        columns = ("id", "fecha", "remitente", "destinatario", "asunto", "adjuntos", "tamano")
        self.tree = ttk.Treeview(left, columns=columns, show="headings", selectmode="extended")
        self.tree.heading("id", text="ID")
        self.tree.column("id", width=40, anchor="center")
        self.tree.heading("fecha", text="Fecha")
        self.tree.column("fecha", width=105, anchor="center")
        self.tree.heading("remitente", text="De")
        self.tree.column("remitente", width=155)
        self.tree.heading("destinatario", text="Para")
        self.tree.column("destinatario", width=130)
        self.tree.heading("asunto", text="Asunto")
        self.tree.column("asunto", width=230)
        self.tree.heading("adjuntos", text="Adj")
        self.tree.column("adjuntos", width=32, anchor="center")
        self.tree.heading("tamano", text="Tamano")
        self.tree.column("tamano", width=65, anchor="center")

        scrollbar = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)

        # --- Panel derecho: preview ---
        right = ttk.LabelFrame(self.paned, text="Vista Previa")
        self.paned.add(right, weight=2)

        # Header del preview
        self.preview_header = tk.Text(right, height=5, wrap="word",
                                       bg="#f7f9fc", fg="#2c3e50", state="disabled",
                                       font=("Segoe UI", 9), bd=0, padx=10, pady=6)
        self.preview_header.pack(fill="x", padx=5, pady=(5, 0))

        # Cuerpo del preview
        preview_body_frame = tk.Frame(right)
        preview_body_frame.pack(fill="both", expand=True, padx=5, pady=(2, 5))

        self.preview_body = tk.Text(preview_body_frame, wrap="word",
                                     bg="white", fg="#333333", state="disabled",
                                     font=("Segoe UI", 9), bd=1, relief="sunken",
                                     padx=10, pady=6)
        preview_scroll = ttk.Scrollbar(preview_body_frame, orient="vertical",
                                        command=self.preview_body.yview)
        self.preview_body.configure(yscrollcommand=preview_scroll.set)
        self.preview_body.pack(side="left", fill="both", expand=True)
        preview_scroll.pack(side="right", fill="y")

        # Tags para formato en preview
        self.preview_header.tag_configure("label", font=("Segoe UI", 9, "bold"))
        self.preview_body.tag_configure("attachment", foreground="#2980b9",
                                         font=("Segoe UI", 9, "bold"))

        # === BARRA INFERIOR: acciones + progreso ===
        bottom = tk.Frame(self.root, bg="#ecf0f1")
        bottom.pack(fill="x")

        actions = tk.Frame(bottom, bg="#ecf0f1")
        actions.pack(fill="x", padx=10, pady=3)

        self.btn_folder = ttk.Button(actions, text="Carpeta Destino", command=self.choose_folder)
        self.btn_folder.pack(side="left", padx=2)
        self.folder_label = ttk.Label(actions, text=self.output_dir, foreground="blue",
                                       background="#ecf0f1")
        self.folder_label.pack(side="left", padx=5)

        ttk.Separator(actions, orient="vertical").pack(side="left", fill="y", padx=8, pady=2)

        ttk.Label(actions, text="Formato:", background="#ecf0f1").pack(side="left", padx=(0, 3))
        self.format_entry = ttk.Entry(actions, width=20)
        self.format_entry.insert(0, "{date}_{subject}")
        self.format_entry.pack(side="left")

        ttk.Checkbutton(actions, text="PDF", variable=self.download_body).pack(side="left", padx=(10, 2))
        ttk.Checkbutton(actions, text="Adjuntos", variable=self.download_attachments).pack(side="left", padx=2)

        self.btn_download = ttk.Button(actions, text="Descargar Seleccionados", command=self.start_download)
        self.btn_download.pack(side="right", padx=2)

        self.btn_cancel = ttk.Button(actions, text="Cancelar", command=self._cancel_download)

        self.btn_compress = ttk.Button(actions, text="Comprimir PDFs", command=self.start_compress_pdfs)
        self.btn_compress.pack(side="right", padx=2)

        # Barra de progreso (oculta por defecto)
        self.progress_frame = tk.Frame(bottom, bg="#ecf0f1")
        self.progress = ttk.Progressbar(self.progress_frame, mode="determinate")
        self.progress.pack(fill="x", side="left", expand=True)
        self.progress_label = ttk.Label(self.progress_frame, text="", width=40)

        # === Log ===
        self.log_text = tk.Text(self.root, height=4, state='disabled', bg="#f4f4f4",
                                font=("Consolas", 8))
        self.log_text.pack(fill="x", padx=5, pady=(0, 3))

    # -----------------------------------------------------------------------
    # Preview
    # -----------------------------------------------------------------------
    def _on_tree_select(self, event=None):
        self.update_selection_count()
        selected = self.tree.selection()
        if len(selected) == 1:
            values = self.tree.item(selected[0], "values")
            eid_str = values[0]
            self._show_preview(eid_str)
        elif len(selected) == 0:
            self._clear_preview()

    def _clear_preview(self):
        self.preview_header.config(state="normal")
        self.preview_header.delete("1.0", tk.END)
        self.preview_header.config(state="disabled")
        self.preview_body.config(state="normal")
        self.preview_body.delete("1.0", tk.END)
        self.preview_body.config(state="disabled")

    def _show_preview(self, eid_str):
        """Muestra preview de un correo. Usa cache o busca via IMAP."""
        if eid_str in self._email_cache:
            self._render_preview(self._email_cache[eid_str])
            return

        # Buscar en hilo para no bloquear la UI
        threading.Thread(target=self._fetch_preview, args=(eid_str,), daemon=True).start()

    def _fetch_preview(self, eid_str):
        """Descarga headers + body preview via IMAP en hilo separado."""
        try:
            if not self._preview_conn:
                email_addr = self.email_entry.get().strip()
                pwd = self.pass_entry.get().strip()
                if not email_addr or not pwd:
                    return
                self._preview_conn = imaplib.IMAP4_SSL(self.imap_server, timeout=30)
                self._preview_conn.login(email_addr, pwd)
                mailbox = self.mailbox_var.get() or "INBOX"
                self._preview_conn.select(mailbox, readonly=True)

            # Fetch headers completos + body (limitado a ~50KB para preview)
            _, msg_data = self._preview_conn.fetch(
                eid_str.encode(),
                "(BODY.PEEK[HEADER] BODY.PEEK[TEXT]<0.51200>)"
            )

            if not msg_data or msg_data[0] is None:
                return

            # Parsear respuesta IMAP (puede venir en multiples partes)
            raw_header = b""
            raw_body = b""
            for part in msg_data:
                if isinstance(part, tuple):
                    info = part[0].decode("utf-8", errors="ignore") if isinstance(part[0], bytes) else str(part[0])
                    if "HEADER" in info.upper():
                        raw_header = part[1]
                    elif "TEXT" in info.upper():
                        raw_body = part[1]

            msg = email.message_from_bytes(raw_header)

            de = decode_str(msg.get("From", ""))
            para = decode_str(msg.get("To", ""))
            cc = decode_str(msg.get("Cc", ""))
            asunto = decode_str(msg.get("Subject", ""))
            fecha = msg.get("Date", "")

            # Decodificar cuerpo
            body_text = ""
            if raw_body:
                # Intentar decodificar como texto
                for enc in ("utf-8", "latin-1", "ascii"):
                    try:
                        body_text = raw_body.decode(enc)
                        break
                    except (UnicodeDecodeError, AttributeError):
                        continue

            # Limpiar HTML si es HTML
            if "<html" in body_text.lower() or "<body" in body_text.lower():
                body_text = re.sub(r'<style[^>]*>.*?</style>', '', body_text, flags=re.DOTALL | re.IGNORECASE)
                body_text = re.sub(r'<script[^>]*>.*?</script>', '', body_text, flags=re.DOTALL | re.IGNORECASE)
                body_text = re.sub(r'<br\s*/?>', '\n', body_text, flags=re.IGNORECASE)
                body_text = re.sub(r'<p[^>]*>', '\n', body_text, flags=re.IGNORECASE)
                body_text = re.sub(r'<div[^>]*>', '\n', body_text, flags=re.IGNORECASE)
                body_text = re.sub(r'<[^>]+>', '', body_text)
                body_text = html.unescape(body_text)
                # Limpiar lineas vacias excesivas
                body_text = re.sub(r'\n{3,}', '\n\n', body_text)

            data = {
                "de": de, "para": para, "cc": cc,
                "asunto": asunto, "fecha": fecha,
                "body": body_text.strip()
            }
            self._email_cache[eid_str] = data
            self.root.after(0, self._render_preview, data)

        except Exception as e:
            self.root.after(0, self._render_preview_error, str(e))
            # Reset conexion si falla
            try:
                if self._preview_conn:
                    self._preview_conn.logout()
            except Exception:
                pass
            self._preview_conn = None

    def _render_preview(self, data):
        # Header
        self.preview_header.config(state="normal")
        self.preview_header.delete("1.0", tk.END)

        self.preview_header.insert(tk.END, "De: ", "label")
        self.preview_header.insert(tk.END, f"{data['de']}\n")
        self.preview_header.insert(tk.END, "Para: ", "label")
        self.preview_header.insert(tk.END, f"{data['para']}\n")
        if data.get("cc"):
            self.preview_header.insert(tk.END, "CC: ", "label")
            self.preview_header.insert(tk.END, f"{data['cc']}\n")
        self.preview_header.insert(tk.END, "Asunto: ", "label")
        self.preview_header.insert(tk.END, f"{data['asunto']}  ")
        self.preview_header.insert(tk.END, f"  [{data['fecha']}]")

        self.preview_header.config(state="disabled")

        # Body
        self.preview_body.config(state="normal")
        self.preview_body.delete("1.0", tk.END)
        body = data.get("body", "")
        if body:
            self.preview_body.insert(tk.END, body)
        else:
            self.preview_body.insert(tk.END, "(Sin contenido de texto)")
        self.preview_body.config(state="disabled")

    def _render_preview_error(self, error_msg):
        self.preview_header.config(state="normal")
        self.preview_header.delete("1.0", tk.END)
        self.preview_header.insert(tk.END, f"Error cargando preview: {error_msg}")
        self.preview_header.config(state="disabled")
        self.preview_body.config(state="normal")
        self.preview_body.delete("1.0", tk.END)
        self.preview_body.config(state="disabled")

    # -----------------------------------------------------------------------
    # Tabla helpers
    # -----------------------------------------------------------------------
    def _insert_email_row(self, eid, fecha, remitente, destinatario, asunto, adj_count, size_str, size_bytes):
        adj_text = str(adj_count) if adj_count > 0 else ""
        iid = self.tree.insert("", "end", values=(
            eid.decode() if isinstance(eid, bytes) else eid,
            fecha, remitente, destinatario, asunto, adj_text, size_str
        ))
        self.size_cache[iid] = size_bytes

    def _load_settings(self):
        s = load_settings()
        if not s:
            return
        if s.get("email"):
            self.email_entry.insert(0, s["email"])
        if s.get("password"):
            try:
                self.pass_entry.insert(0, base64.b64decode(s["password"]).decode("utf-8"))
            except Exception:
                pass
        # Filtros simples
        for field, key in [(self.global_search_entry, "global_search"),
                           (self.sender_entry, "sender"),
                           (self.subject_entry, "subject"),
                           (self.keyword_entry, "keyword"),
                           (self.format_entry, "format")]:
            if key in s:
                field.delete(0, tk.END)
                field.insert(0, s[key])
        # Nuevos campos
        if s.get("to"):
            self.to_entry.insert(0, s["to"])
        if s.get("cc"):
            self.cc_entry.insert(0, s["cc"])
        # DatePickers
        if s.get("since"):
            self.since_picker.set_text(s["since"])
        if s.get("before"):
            self.before_picker.set_text(s["before"])
        # Buzon
        if s.get("mailbox"):
            self.mailbox_var.set(s["mailbox"])
        # Carpeta destino
        if s.get("output_dir"):
            self.output_dir = s["output_dir"]
            self.folder_label.config(text=self.output_dir)
        if "download_body" in s:
            self.download_body.set(s["download_body"])
        if "download_attachments" in s:
            self.download_attachments.set(s["download_attachments"])

    def _save_settings(self):
        save_settings({
            "email": self.email_entry.get().strip(),
            "password": base64.b64encode(self.pass_entry.get().strip().encode("utf-8")).decode("ascii"),
            "global_search": self.global_search_entry.get().strip(),
            "sender": self.sender_entry.get().strip(),
            "to": self.to_entry.get().strip(),
            "cc": self.cc_entry.get().strip(),
            "since": self.since_picker.get(),
            "before": self.before_picker.get(),
            "subject": self.subject_entry.get().strip(),
            "keyword": self.keyword_entry.get().strip(),
            "mailbox": self.mailbox_var.get(),
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
        if self.mail_conn:
            try:
                self.mail_conn.logout()
            except Exception:
                pass
            self.mail_conn = None

    def _close_preview_conn(self):
        if self._preview_conn:
            try:
                self._preview_conn.logout()
            except Exception:
                pass
            self._preview_conn = None

    def connect_imap(self):
        email_addr = self.email_entry.get().strip()
        pwd = self.pass_entry.get().strip()
        if not email_addr or not pwd:
            raise ValueError("Falta el correo o el App Password.")

        self.log("Conectando a Gmail...")
        self.mail_conn = imaplib.IMAP4_SSL(self.imap_server, timeout=120)
        self.mail_conn.login(email_addr, pwd)

        # Listar buzones disponibles
        try:
            status, mailbox_data = self.mail_conn.list()
            if status == "OK":
                boxes = []
                for mb in mailbox_data:
                    if isinstance(mb, bytes):
                        # Formato: (\\flags) "/" "nombre"
                        match = re.search(r'"([^"]*)"$|(\S+)$', mb.decode("utf-8", errors="ignore"))
                        if match:
                            name = match.group(1) or match.group(2)
                            if name:
                                boxes.append(name)
                if boxes:
                    self._mailboxes = sorted(boxes)
                    self.root.after(0, self._update_mailbox_combo)
        except Exception:
            pass

        mailbox = self.mailbox_var.get() or "INBOX"
        self.mail_conn.select(mailbox)
        self.log(f"Conectado a buzon: {mailbox}")

    def _update_mailbox_combo(self):
        self.mailbox_combo["values"] = self._mailboxes

    def start_test_connection(self):
        self.root.after(0, self.btn_test.config, {"state": "disabled"})
        threading.Thread(target=self._test_connection, daemon=True).start()

    def _test_connection(self):
        try:
            email_addr = self.email_entry.get().strip()
            pwd = self.pass_entry.get().strip()
            if not email_addr or not pwd:
                self.log("[ERROR] Falta el correo o el App Password.")
                return

            self.log("Probando conexion...")
            conn = imaplib.IMAP4_SSL(self.imap_server, timeout=30)
            conn.login(email_addr, pwd)

            # Listar buzones
            try:
                status, mailbox_data = conn.list()
                if status == "OK":
                    boxes = []
                    for mb in mailbox_data:
                        if isinstance(mb, bytes):
                            match = re.search(r'"([^"]*)"$|(\S+)$', mb.decode("utf-8", errors="ignore"))
                            if match:
                                name = match.group(1) or match.group(2)
                                if name:
                                    boxes.append(name)
                    if boxes:
                        self._mailboxes = sorted(boxes)
                        self.root.after(0, self._update_mailbox_combo)
                        self.log(f"  {len(boxes)} buzones encontrados.")
            except Exception:
                pass

            conn.logout()
            self.log("Conexion exitosa. Credenciales validas.")
        except imaplib.IMAP4.error as e:
            self.log(f"[ERROR] Autenticacion fallida: {e}")
        except Exception as e:
            self.log(f"[ERROR] No se pudo conectar: {e}")
        finally:
            self.root.after(0, self.btn_test.config, {"state": "normal"})

    def start_search(self):
        self._save_settings()
        self.btn_search.config(state="disabled")
        self.tree.delete(*self.tree.get_children())
        self.size_cache.clear()
        self._email_cache.clear()
        self._close_preview_conn()
        self._clear_preview()
        threading.Thread(target=self.search_emails, daemon=True).start()

    def _build_or_chain(self, parts):
        """Construye una cadena OR anidada para IMAP.
        OR solo acepta 2 operandos, asi que se anida:
        [A, B, C] -> OR A (OR B C)
        """
        if len(parts) == 1:
            return parts[0]
        if len(parts) == 2:
            return f'OR {parts[0]} {parts[1]}'
        return f'OR {parts[0]} ({self._build_or_chain(parts[1:])})'

    def search_emails(self):
        try:
            self.connect_imap()
            global_term = self.global_search_entry.get().strip()
            sender = self.sender_entry.get().strip()
            to_addr = self.to_entry.get().strip()
            cc_addr = self.cc_entry.get().strip()
            since = self.since_picker.get()
            before = self.before_picker.get()
            subject_kw = self.subject_entry.get().strip()
            body_kw = self.keyword_entry.get().strip()
            want_attachments_only = self.has_attachments_filter.get()

            # Validar fechas
            if since and not validate_imap_date(since):
                self.log(f"[ERROR] Fecha invalida: '{since}'. Use DD-MMM-YYYY (ej: 01-Jan-2025)")
                return
            if before and not validate_imap_date(before):
                self.log(f"[ERROR] Fecha invalida: '{before}'. Use DD-MMM-YYYY (ej: 31-Dec-2026)")
                return

            criteria = []

            # "Buscar en todo" -> OR entre FROM, TO, CC, BCC, SUBJECT, TEXT
            if global_term:
                or_parts = [
                    f'FROM "{global_term}"',
                    f'TO "{global_term}"',
                    f'CC "{global_term}"',
                    f'BCC "{global_term}"',
                    f'SUBJECT "{global_term}"',
                    f'TEXT "{global_term}"',
                ]
                criteria.append(self._build_or_chain(or_parts))

            # Filtros especificos (AND con el resultado del OR)
            if sender:
                criteria.append(f'FROM "{sender}"')
            if to_addr:
                criteria.append(f'TO "{to_addr}"')
            if cc_addr:
                criteria.append(f'CC "{cc_addr}"')
            if since:
                criteria.append(f'SINCE "{since}"')
            if before:
                criteria.append(f'BEFORE "{before}"')
            if subject_kw:
                criteria.append(f'SUBJECT "{subject_kw}"')
            if body_kw:
                criteria.append(f'TEXT "{body_kw}"')

            search_query = f"({' '.join(criteria)})" if criteria else "ALL"
            self.log(f"Buscando: {search_query}")

            status, data = self.mail_conn.search(None, search_query)
            if status != "OK":
                self.log("Error en la busqueda IMAP.")
                return

            ids = data[0].split()
            if not ids:
                self.log("No se encontraron correos con esos criterios.")
                return

            self.log(f"Se encontraron {len(ids)} correos. Obteniendo datos...")

            # Fetch headers expandidos: DATE, FROM, TO, SUBJECT, Content-Type
            for eid in ids:
                _, msg_data = self.mail_conn.fetch(
                    eid,
                    "(RFC822.SIZE BODY.PEEK[HEADER.FIELDS (DATE FROM TO CC SUBJECT CONTENT-TYPE)])"
                )
                if not msg_data or msg_data[0] is None:
                    continue

                raw_part = msg_data[0]
                size_bytes = 0
                size_str = "?"
                if isinstance(raw_part, tuple):
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
                    fecha_str = fecha_dt.strftime("%Y-%m-%d %H:%M")
                except (ValueError, TypeError):
                    fecha_str = "Fecha desconocida"

                remitente = decode_str(msg.get("From", ""))
                destinatario = decode_str(msg.get("To", ""))
                asunto = decode_str(msg.get("Subject", ""))

                # Detectar adjuntos via Content-Type header
                content_type = msg.get("Content-Type", "")
                has_att = 1 if "multipart/mixed" in content_type.lower() else 0

                # Filtro de adjuntos
                if want_attachments_only and has_att == 0:
                    continue

                self.root.after(
                    0,
                    lambda e=eid, f=fecha_str, r=remitente, d=destinatario,
                           a=asunto, att=has_att, s=size_str, sb=size_bytes:
                        self._insert_email_row(e, f, r, d, a, att, s, sb)
                )

            total_shown = len(self.tree.get_children())
            self.log(f"Busqueda completada. {total_shown} correos mostrados.")
        except Exception as e:
            self.log(f"Error: {str(e)}")
        finally:
            self.root.after(0, self.btn_search.config, {"state": "normal"})
            self._close_imap()

    def _cancel_download(self):
        self._cancel_event.set()
        self.btn_cancel.config(state="disabled")
        self.log("[CANCELANDO] Deteniendo despues del correo actual...")

    def start_download(self):
        self._save_settings()
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("Atencion", "Debes seleccionar al menos un correo de la tabla.")
            return

        if not self.download_body.get() and not self.download_attachments.get():
            messagebox.showwarning("Atencion", "Debes marcar al menos una opcion: Cuerpo (PDF) o Adjuntos.")
            return

        items_data = [self.tree.item(item, "values") for item in selected_items]
        self._cancel_event.clear()
        self.btn_download.pack_forget()
        self.btn_cancel.pack(side="right", padx=5)

        self.progress.config(maximum=len(items_data), value=0)
        self.progress_label.config(text="Iniciando...")
        self.progress_label.pack(side="right", padx=5)
        self.progress_frame.pack(fill="x", padx=10, pady=(0, 5))

        threading.Thread(target=self.download_emails, args=(items_data,), daemon=True).start()

    def convert_html_to_pdf(self, page, source_html, output_filename):
        try:
            page.set_content(source_html, wait_until="networkidle", timeout=15000)
            page.pdf(path=output_filename, format="Letter",
                     margin={"top": "15mm", "bottom": "15mm",
                             "left": "10mm", "right": "10mm"})
            gs = find_ghostscript()
            if gs:
                self._compress_single_pdf(output_filename, gs)
            return True
        except Exception as e:
            self.log(f"Error convirtiendo PDF: {e}")
            return False

    def _load_manifest(self):
        path = os.path.join(self.output_dir, "manifest.json")
        if os.path.isfile(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return {}

    def _save_manifest(self, manifest):
        path = os.path.join(self.output_dir, "manifest.json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump(manifest, f, ensure_ascii=False, indent=2)

    def _start_playwright(self):
        self.log("Iniciando motor de renderizado PDF...")
        ctx = sync_playwright().start()
        br = ctx.chromium.launch(headless=True)
        pg = br.new_page()
        return ctx, br, pg

    def _restart_playwright(self, pw_ctx, br):
        try:
            if br:
                br.close()
        except Exception:
            pass
        try:
            if pw_ctx:
                pw_ctx.stop()
        except Exception:
            pass
        return self._start_playwright()

    def download_emails(self, items_data):
        want_body = self.download_body.get()
        want_attachments = self.download_attachments.get()

        pw_context = None
        browser = None
        page = None

        try:
            os.makedirs(self.output_dir, exist_ok=True)
            log_path = os.path.join(self.output_dir, "download_log.txt")
            self._log_file = open(log_path, "a", encoding="utf-8")
            self._log_file.write(
                f"\n{'='*50}\n"
                f"Sesion: {datetime.now():%Y-%m-%d %H:%M}\n"
                f"{'='*50}\n"
            )
        except Exception:
            self._log_file = None

        try:
            self.connect_imap()
            manifest = self._load_manifest()

            if want_body and PLAYWRIGHT_AVAILABLE:
                pw_context, browser, page = self._start_playwright()

            template = self.format_entry.get().strip()
            if not template:
                template = "{date}_{subject}"

            batch_bytes = 0
            skipped = 0
            downloaded = 0
            start_time = time.time()
            total = len(items_data)

            for idx, item in enumerate(items_data, 1):
                if self._cancel_event.is_set():
                    self.log(f"[CANCELADO] {downloaded} de {total} descargados.")
                    break

                eid_str = item[0]
                fecha = item[1]
                remitente = item[2]
                # item[3] = destinatario, item[4] = asunto, item[5] = adj, item[6] = tamano
                asunto = item[4]

                if eid_str in manifest:
                    skipped += 1
                    self.log(f"  [{idx}/{total}] [DUPLICADO] {asunto[:30]}...")
                    self.root.after(0, self.progress.config, {"value": idx})
                    continue

                self.log(f"Descargando [{idx}/{total}]: {asunto[:30]}...")

                elapsed = time.time() - start_time
                if downloaded > 0:
                    rate = downloaded / elapsed
                    remaining = (total - idx) / rate
                    mins = int(remaining // 60)
                    secs = int(remaining % 60)
                    eta_text = f"[{idx}/{total}] {idx*100//total}% — ~{mins}m {secs}s restante"
                else:
                    eta_text = f"[{idx}/{total}] Calculando..."
                self.root.after(0, self.progress_label.config, {"text": eta_text})

                raw_email = None
                for attempt in range(3):
                    try:
                        _, msg_data = self.mail_conn.fetch(eid_str.encode(), "(BODY.PEEK[])")
                        raw_email = msg_data[0][1]
                        break
                    except Exception as fetch_err:
                        if attempt < 2:
                            self.log(f"  [RECONECTANDO] Intento {attempt + 2}/3...")
                            try:
                                self._close_imap()
                                self.connect_imap()
                            except Exception:
                                pass
                        else:
                            self.log(f"  [ERROR] No se pudo descargar: {fetch_err}")

                if raw_email is None:
                    continue

                batch_bytes += len(raw_email)
                msg = email.message_from_bytes(raw_email)

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
                os.makedirs(folder_path, exist_ok=True)

                body_html = ""
                body_text = ""
                cid_map = {}

                for part in msg.walk():
                    ctype = part.get_content_type()
                    cdispo = str(part.get("Content-Disposition", ""))

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

                    if want_body and ctype.startswith("image/"):
                        content_id = part.get("Content-ID", "")
                        if content_id:
                            cid = content_id.strip("<>")
                            payload = part.get_payload(decode=True)
                            if payload:
                                b64 = base64.b64encode(payload).decode("ascii")
                                cid_map[cid] = f"data:{ctype};base64,{b64}"

                    if ctype == "text/html" and not body_html:
                        body_html = part.get_payload(decode=True).decode("utf-8", errors="ignore")
                    elif ctype == "text/plain" and not body_text:
                        body_text = part.get_payload(decode=True).decode("utf-8", errors="ignore")

                if body_html and cid_map:
                    for cid, data_uri in cid_map.items():
                        body_html = body_html.replace(f"cid:{cid}", data_uri)

                if want_body and page:
                    if not body_html and not body_text:
                        final_html = "<p><em>Este correo no tiene contenido de texto.</em></p>"
                    else:
                        final_html = body_html or f"<pre>{html.escape(body_text)}</pre>"

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
                    pdf_ok = False
                    for pw_attempt in range(2):
                        try:
                            if self.convert_html_to_pdf(page, pdf_html, pdf_path):
                                pdf_ok = True
                                break
                        except Exception as pw_err:
                            if pw_attempt == 0:
                                self.log("  [PLAYWRIGHT] Reiniciando...")
                                pw_context, browser, page = self._restart_playwright(pw_context, browser)
                            else:
                                self.log(f"  [ERROR PDF] {pw_err}")
                    if pdf_ok:
                        self.log("  PDF guardado en Mensaje_Legible.pdf")
                    else:
                        self.log("  [AVISO] No se pudo generar el PDF")
                elif want_body and not PLAYWRIGHT_AVAILABLE:
                    self.log("  [AVISO] playwright no instalado, PDF omitido")

                manifest[eid_str] = {
                    "fecha": fecha,
                    "asunto": asunto[:80],
                    "descargado": datetime.now().isoformat()
                }
                self._save_manifest(manifest)
                downloaded += 1
                self.root.after(0, self.progress.config, {"value": idx})

            elapsed = time.time() - start_time
            mins = int(elapsed // 60)
            secs = int(elapsed % 60)
            self.daily_bytes += batch_bytes
            save_daily_quota(self.daily_bytes)
            pct = self.daily_bytes / GMAIL_DAILY_LIMIT * 100
            self.log("")
            self.log(f"{'='*40}")
            self.log(f"Descarga finalizada en {mins}m {secs}s")
            self.log(f"  Descargados: {downloaded}")
            if skipped > 0:
                self.log(f"  Duplicados saltados: {skipped}")
            if self._cancel_event.is_set():
                self.log(f"  Cancelados: {total - idx}")
            self.log(f"[CUOTA] Batch: {format_size(batch_bytes)} | Hoy: {format_size(self.daily_bytes)} de 2500 MB ({pct:.1f}%)")
            if pct >= 80:
                self.log("[CUOTA] ADVERTENCIA: >80% del limite diario.")
            self.log(f"{'='*40}")
            self.root.after(0, self._download_finished, downloaded)

        except Exception as e:
            self.log(f"Error critico durante descarga: {str(e)}")
            self.root.after(0, messagebox.showerror, "Error", f"Ocurrio un error:\n{str(e)}")
        finally:
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
            def _restore_ui():
                self.btn_cancel.pack_forget()
                self.btn_download.pack(side="right", padx=5)
                self.progress_frame.pack_forget()
            self.root.after(0, _restore_ui)
            self._close_imap()
            if self._log_file:
                try:
                    self._log_file.close()
                except Exception:
                    pass
                self._log_file = None

    def _compress_single_pdf(self, pdf_path, gs_path=None):
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
                return 0, 0, err_msg or "Ghostscript fallo"

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
        gs = find_ghostscript()
        if not gs:
            messagebox.showwarning(
                "Ghostscript no encontrado",
                "Instala Ghostscript desde:\n"
                "https://ghostscript.com/releases/gsdnld.html\n\n"
                "Reinicia la app despues de instalar."
            )
            return
        if not os.path.isdir(self.output_dir):
            messagebox.showwarning(
                "Carpeta no encontrada",
                f"La carpeta no existe:\n{self.output_dir}"
            )
            return
        self.btn_compress.config(state="disabled")
        threading.Thread(target=self._compress_pdfs_in_folder, daemon=True).start()

    def _compress_pdfs_in_folder(self):
        try:
            pdf_files = []
            for dirpath, _, filenames in os.walk(self.output_dir):
                for f in filenames:
                    if f.lower().endswith(".pdf"):
                        pdf_files.append(os.path.join(dirpath, f))

            if not pdf_files:
                self.log("[PDF] No se encontraron archivos PDF.")
                return

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
                    self.log(f"  [{i}/{len(pdf_files)}] {rel_path}: ERROR - {err}")
                    continue
                total_original += orig
                total_compressed += comp
                if orig > comp:
                    compressed_count += 1
                    savings = (1 - comp / orig) * 100
                    self.log(
                        f"  [{i}/{len(pdf_files)}] {rel_path}: "
                        f"{format_size(orig)} -> {format_size(comp)} ({savings:.0f}%)"
                    )
                else:
                    self.log(f"  [{i}/{len(pdf_files)}] {rel_path}: ya optimizado ({format_size(orig)})")

            processed = len(pdf_files) - error_count
            if total_original > 0:
                total_savings = (1 - total_compressed / total_original) * 100
                self.log(
                    f"[PDF] Listo: {compressed_count}/{processed} comprimidos | "
                    f"{format_size(total_original)} -> {format_size(total_compressed)} "
                    f"(ahorro total {total_savings:.0f}%)"
                )
            elif processed > 0:
                self.log(f"[PDF] {processed} PDF(s) ya estaban optimizados.")
            if error_count > 0:
                self.log(f"[PDF] {error_count} PDF(s) no se pudieron procesar.")

        except Exception as e:
            self.log(f"[PDF] Error: {e}")
        finally:
            self.root.after(0, self.btn_compress.config, {"state": "normal"})

    def _download_finished(self, count):
        if messagebox.askyesno("Exito", f"Se han descargado {count} correos exitosamente.\nAbrir carpeta de destino?"):
            os.startfile(self.output_dir)


if __name__ == "__main__":
    root = tk.Tk()
    app = GmailDownloaderApp(root)
    root.mainloop()
