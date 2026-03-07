import imaplib
import email
import html
import re
import threading
import tkinter as tk
from email.utils import parsedate_to_datetime

from utils import decode_str, validate_imap_date, format_size, debug_log


class IMAPMixin:
    """Mixin para operaciones IMAP: conexion, busqueda y preview."""

    def _list_mailboxes(self, conn):
        """Parsea la lista de buzones de una conexion IMAP. Retorna lista ordenada.
        Filtra buzones con flag \\Noselect (carpetas padre no seleccionables)."""
        try:
            status, mailbox_data = conn.list()
            if status == "OK":
                boxes = []
                for mb in mailbox_data:
                    if isinstance(mb, bytes):
                        line = mb.decode("utf-8", errors="ignore")
                        # Saltar buzones con \Noselect
                        if "\\Noselect" in line:
                            continue
                        match = re.search(r'"([^"]*)"$|(\S+)$', line)
                        if match:
                            name = match.group(1) or match.group(2)
                            if name:
                                boxes.append(name)
                if boxes:
                    return sorted(boxes)
        except Exception:
            pass
        return []

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
            self._preview_mailbox = None

    def connect_imap(self):
        email_addr = self.email_entry.get().strip()
        pwd = self.pass_entry.get().strip()
        if not email_addr or not pwd:
            raise ValueError("Falta el correo o el password.")

        self.imap_server = self.server_entry.get().strip() or "imap.gmail.com"
        self.log(f"Conectando a {self.imap_server}...")
        self.mail_conn = imaplib.IMAP4_SSL(self.imap_server, timeout=120)
        self.mail_conn.login(email_addr, pwd)

        # Listar buzones disponibles
        boxes = self._list_mailboxes(self.mail_conn)
        if boxes:
            self._mailboxes = boxes
            self.root.after(0, self._update_mailbox_combo)

        mailbox = self.mailbox_var.get() or "INBOX"
        if mailbox == "[Todos]":
            self.mail_conn.select("INBOX")
            self.log("Conectado (modo Todos los buzones)")
        else:
            self.mail_conn.select(self._quote_mailbox(mailbox))
            self.log(f"Conectado a buzon: {mailbox}")

    def _update_mailbox_combo(self):
        self.mailbox_combo["values"] = ["[Todos]"] + self._mailboxes

    def _quote_mailbox(self, name):
        """Quotea nombre de buzon para IMAP si contiene espacios o caracteres especiales."""
        if " " in name or "[" in name:
            return f'"{name}"'
        return name

    def _parse_eid(self, eid_str):
        """Separa 'mailbox||eid' -> (mailbox, eid). Sin || retorna (None, eid)."""
        if isinstance(eid_str, bytes):
            eid_str = eid_str.decode("utf-8", errors="ignore")
        if "||" in eid_str:
            parts = eid_str.split("||", 1)
            return parts[0], parts[1]
        return None, eid_str

    def start_test_connection(self):
        self.root.after(0, self.btn_test.config, {"state": "disabled"})
        threading.Thread(target=self._test_connection, daemon=True).start()

    def _test_connection(self):
        try:
            email_addr = self.email_entry.get().strip()
            pwd = self.pass_entry.get().strip()
            if not email_addr or not pwd:
                self.log("[ERROR] Falta el correo o el password.")
                return

            self.imap_server = self.server_entry.get().strip() or "imap.gmail.com"
            self.log(f"Probando conexion a {self.imap_server}...")
            conn = imaplib.IMAP4_SSL(self.imap_server, timeout=30)
            conn.login(email_addr, pwd)

            # Listar buzones
            boxes = self._list_mailboxes(conn)
            if boxes:
                self._mailboxes = boxes
                self.root.after(0, self._update_mailbox_combo)
                self.log(f"  {len(boxes)} buzones encontrados.")

            conn.logout()
            self.log("Conexion exitosa. Credenciales validas.")
        except imaplib.IMAP4.error as e:
            err = str(e)
            if "Application-specific password required" in err:
                self.log("[ERROR] Se requiere App Password. Haz clic en (?) junto al campo Password para ver instrucciones.")
            elif "AUTHENTICATIONFAILED" in err:
                self.log("[ERROR] Credenciales invalidas. Asegurate de usar un App Password, no tu contrasena normal.")
            else:
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
        """Construye cadena OR anidada: [A,B,C] -> OR A (OR B C)"""
        if len(parts) == 1:
            return parts[0]
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

            for field, val in [("FROM", sender), ("TO", to_addr), ("CC", cc_addr),
                               ("SINCE", since), ("BEFORE", before),
                               ("SUBJECT", subject_kw), ("TEXT", body_kw)]:
                if val:
                    criteria.append(f'{field} "{val}"')

            search_query = f"({' '.join(criteria)})" if criteria else "ALL"
            self.log(f"Buscando: {search_query}")

            mailbox_selected = self.mailbox_var.get()
            is_todos = mailbox_selected == "[Todos]"

            if is_todos:
                mailboxes_to_search = self._mailboxes if self._mailboxes else ["INBOX"]
            else:
                mailboxes_to_search = [None]  # None = ya seleccionado por connect_imap

            total_shown = 0
            seen_msgids = set()
            for mbox in mailboxes_to_search:
                if mbox:
                    try:
                        status, _ = self.mail_conn.select(self._quote_mailbox(mbox))
                        if status != "OK":
                            continue
                    except Exception:
                        self.log(f"  Buzon no accesible: {mbox}")
                        continue
                    self.log(f"Buscando en {mbox}...")

                try:
                    status, data = self.mail_conn.search(None, search_query)
                except Exception as e:
                    debug_log.warning(f"Search failed in {mbox}: {e}")
                    continue
                if status != "OK":
                    continue

                ids = data[0].split()
                if not ids:
                    continue

                if is_todos:
                    self.log(f"  {len(ids)} correos en {mbox}")

                for eid in ids:
                    try:
                        _, msg_data = self.mail_conn.fetch(
                            eid,
                            "(RFC822.SIZE BODY.PEEK[HEADER.FIELDS (DATE FROM TO CC SUBJECT CONTENT-TYPE MESSAGE-ID)])"
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

                        # Dedup: si ya vimos este Message-ID en otro buzon, saltar
                        if is_todos:
                            msgid = msg.get("Message-ID", "").strip()
                            if msgid and msgid in seen_msgids:
                                continue
                            if msgid:
                                seen_msgids.add(msgid)

                        content_type = msg.get("Content-Type", "")
                        has_att = 1 if "multipart/mixed" in content_type.lower() else 0

                        if want_attachments_only and has_att == 0:
                            continue

                        eid_val = eid.decode() if isinstance(eid, bytes) else eid
                        if is_todos and mbox:
                            eid_display = f"{mbox}||{eid_val}"
                        else:
                            eid_display = eid_val

                        total_shown += 1
                        self.root.after(
                            0,
                            lambda e=eid_display, f=fecha_str, r=remitente, d=destinatario,
                                   a=asunto, att=has_att, s=size_str, sb=size_bytes:
                                self._insert_email_row(e, f, r, d, a, att, s, sb)
                        )
                    except Exception as parse_err:
                        eid_val = eid.decode() if isinstance(eid, bytes) else eid
                        debug_log.warning(f"Error parsing email {eid_val}: {parse_err}")
                        continue

            if total_shown == 0:
                self.log("No se encontraron correos con esos criterios.")
            else:
                self.log(f"Busqueda completada. {total_shown} correos mostrados.")
        except Exception as e:
            self.log(f"Error: {str(e)}")
        finally:
            self.root.after(0, self.btn_search.config, {"state": "normal"})
            self._close_imap()

    # -------------------------------------------------------------------
    # Preview
    # -------------------------------------------------------------------
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
        threading.Thread(target=self._fetch_preview, args=(eid_str,), daemon=True).start()

    def _fetch_preview(self, eid_str):
        """Descarga headers + body preview via IMAP en hilo separado."""
        try:
            mbox, real_eid = self._parse_eid(eid_str)

            if not self._preview_conn:
                email_addr = self.email_entry.get().strip()
                pwd = self.pass_entry.get().strip()
                if not email_addr or not pwd:
                    return
                self.imap_server = self.server_entry.get().strip() or "imap.gmail.com"
                self._preview_conn = imaplib.IMAP4_SSL(self.imap_server, timeout=30)
                self._preview_conn.login(email_addr, pwd)
                self._preview_mailbox = None

            # Seleccionar buzon correcto
            target_mbox = mbox or self.mailbox_var.get() or "INBOX"
            if target_mbox == "[Todos]":
                target_mbox = "INBOX"
            if target_mbox != self._preview_mailbox:
                self._preview_conn.select(self._quote_mailbox(target_mbox), readonly=True)
                self._preview_mailbox = target_mbox

            _, msg_data = self._preview_conn.fetch(
                real_eid.encode(),
                "(BODY.PEEK[HEADER] BODY.PEEK[TEXT]<0.51200>)"
            )

            if not msg_data or msg_data[0] is None:
                return

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

            body_text = ""
            if raw_body:
                for enc in ("utf-8", "latin-1", "ascii"):
                    try:
                        body_text = raw_body.decode(enc)
                        break
                    except (UnicodeDecodeError, AttributeError):
                        continue

            if "<html" in body_text.lower() or "<body" in body_text.lower():
                body_text = re.sub(r'<style[^>]*>.*?</style>', '', body_text, flags=re.DOTALL | re.IGNORECASE)
                body_text = re.sub(r'<script[^>]*>.*?</script>', '', body_text, flags=re.DOTALL | re.IGNORECASE)
                body_text = re.sub(r'<br\s*/?>', '\n', body_text, flags=re.IGNORECASE)
                body_text = re.sub(r'<p[^>]*>', '\n', body_text, flags=re.IGNORECASE)
                body_text = re.sub(r'<div[^>]*>', '\n', body_text, flags=re.IGNORECASE)
                body_text = re.sub(r'<[^>]+>', '', body_text)
                body_text = html.unescape(body_text)
                body_text = re.sub(r'\n{3,}', '\n\n', body_text)

            data = {
                "de": de, "para": para, "cc": cc,
                "asunto": asunto, "fecha": fecha,
                "body": body_text.strip()
            }
            self._email_cache[eid_str] = data
            self.root.after(0, self._render_preview, data)

        except Exception as e:
            debug_log.error(f"Preview fetch failed: {e}", exc_info=True)
            self.root.after(0, self._render_preview_error, str(e))
            self._close_preview_conn()

    def _render_preview(self, data):
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

        self.preview_body.config(state="normal")
        self.preview_body.delete("1.0", tk.END)
        body = data.get("body", "")
        if body:
            self.preview_body.insert(tk.END, body)
        else:
            self.preview_body.insert(tk.END, "(Sin contenido de texto)")
        self.preview_body.config(state="disabled")

    def _render_preview_error(self, error_msg):
        self._clear_preview()
        self.preview_header.config(state="normal")
        self.preview_header.insert(tk.END, f"Error cargando preview: {error_msg}")
        self.preview_header.config(state="disabled")
