import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import os
import base64
from datetime import datetime

from utils import (
    format_size, load_settings, save_settings, load_daily_quota,
    ensure_chromium_installed, PLAYWRIGHT_AVAILABLE, GMAIL_DAILY_LIMIT,
    debug_log
)
from datepicker import DatePicker
from imap_handler import IMAPMixin
from download_handler import DownloadMixin


class GmailDownloaderApp(IMAPMixin, DownloadMixin):
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
        self._preview_conn = None
        self._mailboxes = []
        self._email_cache = {}
        self._preview_mailbox = None

        self.create_widgets()
        self._load_settings()

        if self.daily_bytes > 0:
            pct = self.daily_bytes / GMAIL_DAILY_LIMIT * 100
            self.log(f"[CUOTA] Uso hoy: {format_size(self.daily_bytes)} de 2500 MB ({pct:.1f}%)")

        if not PLAYWRIGHT_AVAILABLE:
            self.log("[AVISO] playwright no esta instalado. No se podran generar PDFs.")
        else:
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
        # Log a archivo de descarga
        if self._log_file:
            try:
                ts = datetime.now().strftime("%H:%M:%S")
                self._log_file.write(f"[{ts}] {message}\n")
                self._log_file.flush()
            except Exception:
                pass
        # Log de depuracion
        if message.startswith("[ERROR]") or message.startswith("Error"):
            debug_log.error(message)
        elif message.startswith("[AVISO]") or message.startswith("[CANCELANDO]"):
            debug_log.warning(message)
        else:
            debug_log.info(message)

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

        tk.Label(header, text="App Password:", bg="#2c3e50", fg="#bdc3c7",
                 font=("Segoe UI", 9)).pack(side="left", padx=(12, 3))
        self.pass_entry = ttk.Entry(header, width=20, show="*")
        self.pass_entry.pack(side="left", pady=8)

        self._pass_help = tk.Label(header, text="(?)", bg="#2c3e50", fg="#f39c12",
                                    font=("Segoe UI", 9, "bold"), cursor="hand2")
        self._pass_help.pack(side="left", padx=(2, 0))
        self._pass_help.bind("<Button-1>", lambda e: self._show_app_password_help())

        self.btn_test = ttk.Button(header, text="Conectar", command=self.start_test_connection)
        self.btn_test.pack(side="left", padx=8)

        # Buzon selector en el header
        tk.Label(header, text="Buzon:", bg="#2c3e50", fg="#bdc3c7",
                 font=("Segoe UI", 9)).pack(side="left", padx=(12, 3))
        self.mailbox_var = tk.StringVar(value="[Todos]")
        self.mailbox_combo = ttk.Combobox(header, textvariable=self.mailbox_var,
                                           width=16, state="readonly")
        self.mailbox_combo["values"] = ["[Todos]", "INBOX"]
        self.mailbox_combo.pack(side="left", pady=8)

        # === RIBBON DE BUSQUEDA ===
        ribbon = tk.Frame(self.root, bg="#dfe6e9")
        ribbon.pack(fill="x")

        # Fila 1: busqueda global
        r1 = tk.Frame(ribbon, bg="#dfe6e9")
        r1.pack(fill="x", padx=10, pady=(5, 2))

        ttk.Label(r1, text="Buscar en todo:", background="#dfe6e9").pack(side="left")
        self.global_search_entry = ttk.Entry(r1, width=40)
        self.global_search_entry.pack(side="left", padx=5)

        ttk.Separator(r1, orient="vertical").pack(side="left", fill="y", padx=8, pady=2)

        ttk.Label(r1, text="De:", background="#dfe6e9").pack(side="left")
        self.sender_entry = ttk.Entry(r1, width=20)
        self.sender_entry.pack(side="left", padx=5)

        ttk.Label(r1, text="Para:", background="#dfe6e9").pack(side="left")
        self.to_entry = ttk.Entry(r1, width=20)
        self.to_entry.pack(side="left", padx=5)

        ttk.Label(r1, text="CC:", background="#dfe6e9").pack(side="left")
        self.cc_entry = ttk.Entry(r1, width=15)
        self.cc_entry.pack(side="left", padx=5)

        # Fila 2: asunto, keyword, fechas, buscar
        r2 = tk.Frame(ribbon, bg="#dfe6e9")
        r2.pack(fill="x", padx=10, pady=(2, 2))

        ttk.Label(r2, text="Asunto:", background="#dfe6e9").pack(side="left")
        self.subject_entry = ttk.Entry(r2, width=20)
        self.subject_entry.pack(side="left", padx=5)

        ttk.Label(r2, text="Texto:", background="#dfe6e9").pack(side="left")
        self.keyword_entry = ttk.Entry(r2, width=15)
        self.keyword_entry.pack(side="left", padx=5)

        ttk.Checkbutton(r2, text="Solo con adjuntos",
                         variable=self.has_attachments_filter).pack(side="left", padx=(10, 5))

        # Fila 3: fechas y boton buscar
        r3 = tk.Frame(ribbon, bg="#dfe6e9")
        r3.pack(fill="x", padx=10, pady=(2, 5))

        ttk.Label(r3, text="Desde:", background="#dfe6e9").pack(side="left")
        self.since_picker = DatePicker(r3, bg="#dfe6e9")
        self.since_picker.pack(side="left", padx=5)

        ttk.Label(r3, text="Hasta:", background="#dfe6e9").pack(side="left")
        self.before_picker = DatePicker(r3, bg="#dfe6e9")
        self.before_picker.pack(side="left", padx=5)

        self.btn_search = ttk.Button(r3, text="Buscar", command=self.start_search)
        self.btn_search.pack(side="left", padx=5)

        # === CONTENIDO PRINCIPAL: tabla + preview (PanedWindow horizontal) ===
        self.paned = ttk.PanedWindow(self.root, orient="horizontal")
        self.paned.pack(fill="both", expand=True, padx=5, pady=2)

        # --- Panel izquierdo: tabla ---
        left = ttk.LabelFrame(self.paned, text="Resultados")
        self.paned.add(left, weight=3)

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

        self.preview_header = tk.Text(right, height=5, wrap="word",
                                       bg="#f7f9fc", fg="#2c3e50", state="disabled",
                                       font=("Segoe UI", 9), bd=0, padx=10, pady=6)
        self.preview_header.pack(fill="x", padx=5, pady=(5, 0))

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

        self.preview_header.tag_configure("label", font=("Segoe UI", 9, "bold"))
        self.preview_body.tag_configure("attachment", foreground="#2980b9",
                                         font=("Segoe UI", 9, "bold"))

        # === TOOLBAR: seleccionar (entre tabla y acciones de descarga) ===
        frame_toolbar = tk.Frame(self.root, bg="#ecf0f1")
        frame_toolbar.pack(fill="x", padx=10, pady=2)
        ttk.Button(frame_toolbar, text="Seleccionar todos", command=self.select_all).pack(side="left", padx=2)
        ttk.Button(frame_toolbar, text="Deseleccionar", command=self.deselect_all).pack(side="left", padx=2)
        self.selection_label = ttk.Label(frame_toolbar, text="0 de 0 seleccionados", background="#ecf0f1")
        self.selection_label.pack(side="right", padx=5)

        # === TOOLBAR 2: descargar / comprimir ===
        frame_actions = tk.Frame(self.root, bg="#ecf0f1")
        frame_actions.pack(fill="x", padx=10, pady=2)
        self.btn_download = ttk.Button(frame_actions, text="Descargar Seleccionados",
                                        command=self.start_download, state="disabled")
        self.btn_download.pack(side="left", padx=2)
        self.btn_cancel = ttk.Button(frame_actions, text="Cancelar", command=self._cancel_download)
        self.btn_compress = ttk.Button(frame_actions, text="Comprimir PDFs", command=self.start_compress_pdfs)
        self.btn_compress.pack(side="left", padx=2)

        # === BARRA INFERIOR: acciones + progreso ===
        bottom = tk.Frame(self.root, bg="#ecf0f1")
        bottom.pack(fill="x")

        actions = tk.Frame(bottom, bg="#ecf0f1")
        actions.pack(fill="x", padx=10, pady=3)

        self.btn_folder = ttk.Button(actions, text="Carpeta Destino", command=self.choose_folder)
        self.btn_folder.pack(side="left", padx=2)
        self.folder_label = tk.Label(actions, text=self.output_dir, fg="blue",
                                      bg="#ecf0f1", cursor="hand2",
                                      font=("Segoe UI", 9, "underline"))
        self.folder_label.pack(side="left", padx=5)
        self.folder_label.bind("<Button-1>", lambda e: self._open_output_folder())

        ttk.Separator(actions, orient="vertical").pack(side="left", fill="y", padx=8, pady=2)

        ttk.Label(actions, text="Formato:", background="#ecf0f1").pack(side="left", padx=(0, 3))
        self.format_entry = ttk.Entry(actions, width=40)
        self.format_entry.insert(0, "{date}_{subject}")
        self.format_entry.pack(side="left")

        ttk.Checkbutton(actions, text="PDF", variable=self.download_body).pack(side="left", padx=(10, 2))
        ttk.Checkbutton(actions, text="Adjuntos", variable=self.download_attachments).pack(side="left", padx=2)

        # Barra de progreso (oculta por defecto)
        self.progress_frame = tk.Frame(bottom, bg="#ecf0f1")
        self.progress = ttk.Progressbar(self.progress_frame, mode="determinate")
        self.progress.pack(fill="x", side="left", expand=True)
        self.progress_label = ttk.Label(self.progress_frame, text="", width=40)

        # === Log (redimensionable arrastrando el borde superior) ===
        self.log_text = tk.Text(self.root, height=4, state='disabled', bg="#f4f4f4",
                                font=("Consolas", 8))
        self.log_text.pack(fill="both", padx=5, pady=(0, 3))

        # Grip para redimensionar el log arrastrando
        self._log_grip = tk.Frame(self.root, height=4, cursor="sb_v_double_arrow", bg="#cccccc")
        self._log_grip.pack(fill="x", padx=5, before=self.log_text)
        self._log_grip.bind("<B1-Motion>", self._resize_log)
        self._log_grip.bind("<ButtonRelease-1>", lambda e: self._save_settings())

    def _resize_log(self, event):
        """Redimensiona el area de log arrastrando el grip."""
        # Calcular nueva altura basada en posicion del mouse
        log_bottom = self.log_text.winfo_rooty() + self.log_text.winfo_height()
        new_height = log_bottom - event.y_root
        lines = max(2, min(20, new_height // 14))
        self.log_text.config(height=lines)

    # -------------------------------------------------------------------
    # Tabla helpers
    # -------------------------------------------------------------------
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
        for field, key in [(self.global_search_entry, "global_search"),
                           (self.sender_entry, "sender"),
                           (self.subject_entry, "subject"),
                           (self.keyword_entry, "keyword"),
                           (self.format_entry, "format")]:
            if key in s:
                field.delete(0, tk.END)
                field.insert(0, s[key])
        if s.get("to"):
            self.to_entry.insert(0, s["to"])
        if s.get("cc"):
            self.cc_entry.insert(0, s["cc"])
        if s.get("since"):
            self.since_picker.set_text(s["since"])
        if s.get("before"):
            self.before_picker.set_text(s["before"])
        if s.get("mailbox"):
            self.mailbox_var.set(s["mailbox"])
        if s.get("output_dir"):
            self.output_dir = s["output_dir"]
            self.folder_label.config(text=self.output_dir)
        if "download_body" in s:
            self.download_body.set(s["download_body"])
        if "download_attachments" in s:
            self.download_attachments.set(s["download_attachments"])
        if "log_height" in s:
            self.log_text.config(height=max(2, min(20, s["log_height"])))

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
            "log_height": int(self.log_text.cget("height")),
        })

    def select_all(self):
        children = self.tree.get_children()
        if children:
            self.tree.selection_set(children)

    def deselect_all(self):
        self.tree.selection_remove(self.tree.get_children())

    def update_selection_count(self, event=None):
        total = len(self.tree.get_children())
        selected = len(self.tree.selection())
        self.selection_label.config(text=f"{selected} de {total} seleccionados")
        selected_size = sum(self.size_cache.get(iid, 0) for iid in self.tree.selection())
        if selected_size > 0:
            self.selection_label.config(
                text=f"{selected} de {total} seleccionados ({format_size(selected_size)})"
            )
        self.btn_download.config(state="normal" if selected > 0 else "disabled")

    def choose_folder(self):
        folder = filedialog.askdirectory(initialdir=self.output_dir)
        if folder:
            self.output_dir = folder
            self.folder_label.config(text=folder)

    def _open_output_folder(self):
        os.makedirs(self.output_dir, exist_ok=True)
        os.startfile(self.output_dir)

    def _show_app_password_help(self):
        messagebox.showinfo(
            "Como obtener App Password",
            "Gmail NO acepta tu contrasena normal.\n"
            "Necesitas un App Password (contrasena de aplicacion):\n\n"
            "1. Ve a: myaccount.google.com/apppasswords\n"
            "2. Inicia sesion con tu cuenta Google\n"
            "3. Escribe un nombre (ej: 'Gmail Downloader')\n"
            "4. Haz clic en 'Crear'\n"
            "5. Copia la contrasena de 16 letras generada\n"
            "6. Pegala en el campo 'App Password' de esta app\n\n"
            "REQUISITO: Tener verificacion en 2 pasos activada\n"
            "en tu cuenta de Google."
        )


if __name__ == "__main__":
    root = tk.Tk()
    app = GmailDownloaderApp(root)
    root.mainloop()
