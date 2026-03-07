"""Subproceso aislado para convertir HTML a PDF con Playwright.

Evita conflictos con tkinter/greenlet ejecutando Playwright en su propio proceso.
Protocolo: lee comandos JSON de stdin, escribe resultados JSON a stdout.
"""
import sys
import os
import json

os.environ["PLAYWRIGHT_BROWSERS_PATH"] = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "browsers"
)


def main():
    from playwright.sync_api import sync_playwright

    pw = sync_playwright().start()
    browser = pw.chromium.launch(headless=True)
    page = browser.new_page()

    sys.stdout.write(json.dumps({"ready": True}) + "\n")
    sys.stdout.flush()

    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue

        try:
            cmd = json.loads(line)
        except json.JSONDecodeError:
            sys.stdout.write(json.dumps({"ok": False, "error": "JSON invalido"}) + "\n")
            sys.stdout.flush()
            continue

        html_path = cmd.get("html_path")
        pdf_path = cmd.get("pdf_path")

        try:
            with open(html_path, "r", encoding="utf-8") as f:
                html_content = f.read()

            page.set_content(html_content, wait_until="networkidle", timeout=15000)
            page.pdf(
                path=pdf_path,
                format="Letter",
                margin={"top": "15mm", "bottom": "15mm",
                         "left": "10mm", "right": "10mm"},
            )
            sys.stdout.write(json.dumps({"ok": True}) + "\n")
            sys.stdout.flush()
        except Exception as e:
            try:
                page.close()
            except Exception:
                pass
            try:
                page = browser.new_page()
            except Exception:
                try:
                    browser.close()
                except Exception:
                    pass
                try:
                    pw.stop()
                except Exception:
                    pass
                pw = sync_playwright().start()
                browser = pw.chromium.launch(headless=True)
                page = browser.new_page()

            sys.stdout.write(json.dumps({"ok": False, "error": str(e)}) + "\n")
            sys.stdout.flush()

    try:
        browser.close()
    except Exception:
        pass
    try:
        pw.stop()
    except Exception:
        pass


if __name__ == "__main__":
    main()
