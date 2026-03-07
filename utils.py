import os
import sys
import re
import json
import glob
import logging
import subprocess
from datetime import date, datetime, timedelta
from logging.handlers import TimedRotatingFileHandler
from email.header import decode_header

# -- Constantes --
GMAIL_DAILY_LIMIT = 2500 * 1024 * 1024
APP_DIR = os.path.dirname(os.path.abspath(__file__))
QUOTA_FILE = os.path.join(os.path.expanduser("~"), ".gmail_downloader_quota.json")
SETTINGS_FILE = os.path.join(APP_DIR, "settings.json")
LOG_DIR = os.path.join(APP_DIR, "logs")


def setup_debug_logger():
    """Configura logger de depuracion con rotacion a 7 dias."""
    os.makedirs(LOG_DIR, exist_ok=True)

    # Limpiar logs antiguos (>7 dias)
    now = datetime.now()
    for f in os.listdir(LOG_DIR):
        fpath = os.path.join(LOG_DIR, f)
        if os.path.isfile(fpath) and f.endswith(".log"):
            age = now - datetime.fromtimestamp(os.path.getmtime(fpath))
            if age > timedelta(days=7):
                try:
                    os.remove(fpath)
                except Exception:
                    pass

    logger = logging.getLogger("gmail_downloader")
    if logger.handlers:
        return logger
    logger.setLevel(logging.DEBUG)

    log_file = os.path.join(LOG_DIR, f"debug_{date.today().isoformat()}.log")
    handler = logging.FileHandler(log_file, encoding="utf-8")
    handler.setFormatter(logging.Formatter(
        "%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    ))
    logger.addHandler(handler)
    logger.info("="*50)
    logger.info("Sesion iniciada")
    logger.info("="*50)
    return logger


debug_log = setup_debug_logger()

MONTHS_ES = ["Ene","Feb","Mar","Abr","May","Jun",
             "Jul","Ago","Sep","Oct","Nov","Dic"]
MONTHS_IMAP = ["Jan","Feb","Mar","Apr","May","Jun",
               "Jul","Aug","Sep","Oct","Nov","Dec"]
DAYS_ES = ["Lu","Ma","Mi","Ju","Vi","Sa","Do"]

# -- Playwright portable --
os.environ["PLAYWRIGHT_BROWSERS_PATH"] = os.path.join(APP_DIR, "browsers")

try:
    import playwright
    PLAYWRIGHT_AVAILABLE = True
except ImportError:
    PLAYWRIGHT_AVAILABLE = False


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


def ensure_chromium_installed():
    if not PLAYWRIGHT_AVAILABLE:
        return
    browsers_dir = os.environ.get("PLAYWRIGHT_BROWSERS_PATH", "")
    if not browsers_dir:
        return
    if os.path.isdir(browsers_dir):
        for name in os.listdir(browsers_dir):
            if name.startswith("chromium") and os.path.isdir(os.path.join(browsers_dir, name)):
                # Buscar cualquier exe de chromium/chrome en subdirectorios
                for sub in os.listdir(os.path.join(browsers_dir, name)):
                    sub_path = os.path.join(browsers_dir, name, sub)
                    if os.path.isdir(sub_path):
                        for f in os.listdir(sub_path):
                            if f.endswith(".exe") and ("chrome" in f.lower() or "headless" in f.lower()):
                                return
    subprocess.run(
        [sys.executable, "-m", "playwright", "install", "chromium"],
        check=True
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
    return cleaned.strip()[:100].rstrip(". ")


def decode_str(header_str):
    if not header_str:
        return ""
    try:
        decoded_parts = decode_header(header_str)
    except Exception:
        return header_str
    result = ""
    for part, enc in decoded_parts:
        if isinstance(part, bytes):
            try:
                result += part.decode(enc or "utf-8", errors="ignore")
            except (LookupError, UnicodeDecodeError):
                result += part.decode("utf-8", errors="ignore")
        else:
            result += part
    return result


def validate_imap_date(date_str):
    parts = date_str.split("-")
    if len(parts) != 3:
        return False
    day, month, year = parts
    if not day.isdigit() or not year.isdigit():
        return False
    if month not in set(MONTHS_IMAP):
        return False
    if not (1 <= int(day) <= 31) or not (1900 <= int(year) <= 2100):
        return False
    return True


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
