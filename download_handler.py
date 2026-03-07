import os
import sys
import json
import time
import threading
import email
import base64
import subprocess
from datetime import datetime
from tkinter import messagebox

from utils import (
    find_ghostscript, format_size, clean_filename, decode_str,
    PLAYWRIGHT_AVAILABLE, save_daily_quota, GMAIL_DAILY_LIMIT,
    APP_DIR, debug_log
)
from email_rendering import (
    BODY_FORMAT_GMAIL_PDF,
    build_gmail_print_document,
    build_original_email_document,
    body_format_needs_original_html,
    body_format_needs_pdf,
    plain_text_to_html_fragment,
    sanitize_email_html,
)


class DownloadMixin:
    """Mixin para descarga de correos, conversion PDF y compresion."""

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
            messagebox.showwarning("Atencion", "Debes marcar al menos una opcion: Cuerpo o Adjuntos.")
            return

        items_data = [self.tree.item(item, "values") for item in selected_items]
        self._cancel_event.clear()
        self.btn_download.pack_forget()
        self.btn_cancel.pack(side="left", padx=2)

        self.progress.config(maximum=len(items_data), value=0)
        self.progress_label.config(text="Iniciando...")
        self.progress_label.pack(side="right", padx=5)
        self.progress_frame.pack(fill="x", padx=10, pady=(0, 5))

        threading.Thread(target=self.download_emails, args=(items_data,), daemon=True).start()

    def convert_html_to_pdf(self, worker, source_html, output_filename):
        tmp_html = output_filename + ".tmp.html"
        try:
            with open(tmp_html, "w", encoding="utf-8") as f:
                f.write(source_html)
            cmd = json.dumps({"html_path": tmp_html, "pdf_path": output_filename})
            worker.stdin.write(cmd + "\n")
            worker.stdin.flush()
            line = worker.stdout.readline()
            if not line:
                raise RuntimeError("PDF worker no responde")
            result = json.loads(line)
            if not result.get("ok"):
                raise RuntimeError(result.get("error", "Error desconocido"))
            gs = find_ghostscript()
            if gs:
                self._compress_single_pdf(output_filename, gs)
            return True
        finally:
            try:
                os.remove(tmp_html)
            except OSError:
                pass

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

    def _start_pdf_worker(self):
        self.log("Iniciando motor de renderizado PDF...")
        py_exe = sys.executable
        if py_exe.lower().endswith("pythonw.exe"):
            py_exe = py_exe[:-len("pythonw.exe")] + "python.exe"
        worker = subprocess.Popen(
            [py_exe, os.path.join(APP_DIR, "pdf_worker.py")],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        line = worker.stdout.readline()
        if not line:
            stderr = worker.stderr.read()
            debug_log.error(f"PDF worker failed to start: {stderr}")
            raise RuntimeError(f"PDF worker no inicio: {stderr[:200]}")
        result = json.loads(line)
        if not result.get("ready"):
            raise RuntimeError("PDF worker no envio señal ready")
        return worker

    def _stop_pdf_worker(self, worker):
        if not worker:
            return
        try:
            worker.stdin.close()
            worker.wait(timeout=10)
        except Exception:
            try:
                worker.kill()
            except Exception:
                pass

    def _restart_pdf_worker(self, worker):
        self._stop_pdf_worker(worker)
        return self._start_pdf_worker()

    def download_emails(self, items_data):
        want_body = self.download_body.get()
        want_attachments = self.download_attachments.get()
        body_format = getattr(self, "get_body_download_mode", lambda: BODY_FORMAT_GMAIL_PDF)()
        want_gmail_pdf = want_body and body_format_needs_pdf(body_format)
        want_original_html = want_body and body_format_needs_original_html(body_format)

        pdf_worker = None

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

            if want_gmail_pdf and PLAYWRIGHT_AVAILABLE:
                try:
                    pdf_worker = self._start_pdf_worker()
                except Exception:
                    self.log("[AVISO] No se pudo iniciar Playwright; PDF estilo Gmail se guardara como HTML.")
            elif want_gmail_pdf:
                self.log("[AVISO] Playwright no disponible; PDF estilo Gmail se guardara como HTML.")

            template = self.format_entry.get().strip()
            if not template:
                template = "{date}_{subject}"

            batch_bytes = 0
            skipped = 0
            downloaded = 0
            start_time = time.time()
            total = len(items_data)
            current_mbox = None

            for idx, item in enumerate(items_data, 1):
                if self._cancel_event.is_set():
                    self.log(f"[CANCELADO] {downloaded} de {total} descargados.")
                    break

                eid_str = item[0]
                mbox, real_eid = self._parse_eid(eid_str)
                fecha = item[1]
                remitente = item[2]
                destinatario = item[3]
                asunto = item[4]

                # Cambiar buzon si es necesario (modo Todos)
                if mbox and mbox != current_mbox:
                    try:
                        self.mail_conn.select(self._quote_mailbox(mbox))
                        current_mbox = mbox
                    except Exception:
                        self.log(f"  [ERROR] No se pudo abrir buzon: {mbox}")
                        continue

                if eid_str in manifest:
                    # Verificar que la carpeta realmente exista
                    prev = manifest[eid_str]
                    prev_folder = prev.get("folder", "")
                    if prev_folder and os.path.isdir(os.path.join(self.output_dir, prev_folder)):
                        skipped += 1
                        self.log(f"  [{idx}/{total}] [DUPLICADO] {asunto[:30]}...")
                        self.root.after(0, self.progress.config, {"value": idx})
                        continue
                    else:
                        del manifest[eid_str]

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
                        _, msg_data = self.mail_conn.fetch(real_eid.encode(), "(BODY.PEEK[])")
                        raw_email = msg_data[0][1]
                        break
                    except Exception as fetch_err:
                        if attempt < 2:
                            self.log(f"  [RECONECTANDO] Intento {attempt + 2}/3...")
                            try:
                                self._close_imap()
                                self.connect_imap()
                                if mbox:
                                    self.mail_conn.select(self._quote_mailbox(mbox))
                                    current_mbox = mbox
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
                folder_name = folder_name.replace("{id}", real_eid)
                folder_name = clean_filename(folder_name)
                folder_name = folder_name.replace("..", "")

                if not folder_name:
                    folder_name = f"Email_{real_eid}"

                folder_path = os.path.join(self.output_dir, folder_name)

                body_html = ""
                body_text = ""
                cid_map = {}

                for part in msg.walk():
                    ctype = part.get_content_type()
                    cdispo = str(part.get("Content-Disposition", ""))

                    if want_body and ctype.startswith("image/"):
                        content_id = part.get("Content-ID", "")
                        if content_id:
                            img_payload = part.get_payload(decode=True)
                            if img_payload:
                                cid = content_id.strip("<>")
                                b64 = base64.b64encode(img_payload).decode("ascii")
                                cid_map[cid] = f"data:{ctype};base64,{b64}"

                    if "attachment" in cdispo or part.get_filename():
                        if want_attachments:
                            fname = part.get_filename()
                            payload = part.get_payload(decode=True)
                            if fname and payload:
                                fname_dec = decode_str(fname)
                                safe_fname = clean_filename(fname_dec)
                                os.makedirs(folder_path, exist_ok=True)
                                filepath = os.path.join(folder_path, safe_fname)
                                with open(filepath, "wb") as f:
                                    f.write(payload)
                                self.log(f"  Adjunto: {safe_fname}")
                        continue

                    payload = part.get_payload(decode=True)
                    if not payload:
                        continue

                    if ctype == "text/html" and not body_html:
                        body_html = payload.decode("utf-8", errors="ignore")
                    elif ctype == "text/plain" and not body_text:
                        body_text = payload.decode("utf-8", errors="ignore")

                if body_html and cid_map:
                    for cid, data_uri in cid_map.items():
                        body_html = body_html.replace(f"cid:{cid}", data_uri)

                saved_something = os.path.isdir(folder_path)

                if want_body and (body_html or body_text):
                    original_source = body_html or plain_text_to_html_fragment(body_text)
                    gmail_fragment = (
                        sanitize_email_html(body_html)
                        if body_html else plain_text_to_html_fragment(body_text)
                    )
                    if not gmail_fragment and body_text:
                        gmail_fragment = plain_text_to_html_fragment(body_text)
                    os.makedirs(folder_path, exist_ok=True)
                    gmail_html = build_gmail_print_document(
                        account_email=self.email_entry.get().strip(),
                        subject=asunto,
                        sender=remitente,
                        recipient=decode_str(msg.get("To", "")) or destinatario,
                        cc=decode_str(msg.get("Cc", "")),
                        sent_at=fecha,
                        body_fragment=gmail_fragment,
                    )
                    original_document = build_original_email_document(original_source)

                    if want_original_html:
                        original_html_path = os.path.join(folder_path, "Mensaje_Original.html")
                        with open(original_html_path, "w", encoding="utf-8") as hf:
                            hf.write(original_document)
                        self.log("  HTML original guardado en Mensaje_Original.html")

                    pdf_ok = False
                    if want_gmail_pdf and pdf_worker:
                        pdf_path = os.path.join(folder_path, "Mensaje_Gmail.pdf")
                        for pw_attempt in range(2):
                            try:
                                if self.convert_html_to_pdf(pdf_worker, gmail_html, pdf_path):
                                    pdf_ok = True
                                    break
                            except Exception as pw_err:
                                if pw_attempt == 0:
                                    self.log("  [PLAYWRIGHT] Reiniciando...")
                                    try:
                                        pdf_worker = self._restart_pdf_worker(pdf_worker)
                                    except Exception:
                                        pdf_worker = None
                                        self.log(f"  [ERROR PDF] {pw_err}")
                                        break
                                else:
                                    self.log(f"  [ERROR PDF] {pw_err}")

                    if pdf_ok:
                        self.log("  PDF Gmail guardado en Mensaje_Gmail.pdf")
                    elif want_gmail_pdf:
                        html_path = os.path.join(folder_path, "Mensaje_Gmail.html")
                        with open(html_path, "w", encoding="utf-8") as hf:
                            hf.write(gmail_html)
                        self.log("  HTML Gmail guardado en Mensaje_Gmail.html")

                    saved_something = True

                if saved_something:
                    manifest[eid_str] = {
                        "fecha": fecha,
                        "asunto": asunto[:80],
                        "folder": folder_name,
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
            debug_log.error(f"Error critico descarga: {e}", exc_info=True)
            self.log(f"Error critico durante descarga: {str(e)}")
            self.root.after(0, messagebox.showerror, "Error", f"Ocurrio un error:\n{str(e)}")
        finally:
            self._stop_pdf_worker(pdf_worker)
            def _restore_ui():
                self.btn_cancel.pack_forget()
                self.btn_download.pack(side="left", padx=2)
                self.update_selection_count()
                self.progress_frame.pack_forget()
            self.root.after(0, _restore_ui)
            self._close_imap()
            if self._log_file:
                try:
                    self._log_file.close()
                except Exception:
                    pass
                self._log_file = None

    # -------------------------------------------------------------------
    # Compresion PDF
    # -------------------------------------------------------------------
    def _compress_single_pdf(self, pdf_path, gs_path=None):
        if not gs_path:
            gs_path = find_ghostscript()
        if not gs_path:
            return 0, 0, "Ghostscript no encontrado"
        tmp_path = pdf_path + ".gs.tmp"
        try:
            original_size = os.path.getsize(pdf_path)
            result = subprocess.run(
                [gs_path, "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.4",
                 "-dPDFSETTINGS=/printer", "-dNOPAUSE", "-dBATCH", "-dQUIET",
                 f"-sOutputFile={tmp_path}", pdf_path],
                capture_output=True, text=True, timeout=60,
            )
            if result.returncode != 0:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)
                return 0, 0, result.stderr.strip()[:200] or "Ghostscript fallo"
            new_size = os.path.getsize(tmp_path)
            if new_size < original_size:
                os.replace(tmp_path, pdf_path)
                return original_size, new_size, None
            os.remove(tmp_path)
            return original_size, original_size, None
        except Exception as e:
            try:
                os.remove(tmp_path)
            except OSError:
                pass
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
        title = "Descarga completada" if not self._cancel_event.is_set() else "Descarga cancelada"
        if messagebox.askyesno(title, f"Se descargaron {count} correos.\nAbrir carpeta de destino?"):
            os.startfile(self.output_dir)
