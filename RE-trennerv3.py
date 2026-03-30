# RE-TrennerV3 – Kompletter Dienst mit sauberem Stopp-Mechanismus
# Konfiguration wird aus .env-Datei geladen (python-dotenv)

import os
import time
import shutil
import smtplib
import re
import PyPDF2
import pytesseract
import paramiko
import win32api
import win32print
import win32serviceutil
import win32service
import win32event
import servicemanager
from datetime import datetime
from pdf2image import convert_from_path
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from dotenv import load_dotenv

# .env laden – Pfad relativ zum Skript
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(BASE_DIR, ".env"))

# ── Konfiguration aus .env ────────────────────────────────────────────────────
WATCH_FOLDER    = os.getenv("WATCH_FOLDER",   r"C:\Program Files\Python_Rechnungen\ARNEU")
LOG_FOLDER      = os.getenv("LOG_FOLDER",     r"C:\Program Files\Python_Rechnungen\LOG-PY")
ARTR_FOLDER     = os.getenv("ARTR_FOLDER",    r"C:\Program Files\Python_Rechnungen\ARTR")
HOTFOLDER_NET   = os.getenv("HOTFOLDER_NET",  r"C:\Program Files\Python_Rechnungen\Hotfolder_Net")
HOTFOLDER_SFTP  = os.getenv("HOTFOLDER_SFTP", r"C:\Program Files\Python_Rechnungen\Hotfolder_SFTP")
NETWORK_FOLDER  = os.getenv("NETWORK_FOLDER", r"\\192.168.161.11\daten\1-DOPRE")
POPPLER_PATH    = os.getenv("POPPLER_PATH",   r"C:\Program Files\poppler\Library\bin")
PRINTER_NAME    = os.getenv("PRINTER_NAME",   "FFM Drucker_PS")

SFTP_HOST       = os.getenv("SFTP_HOST",   "sftp.hidrive.strato.com")
SFTP_PORT       = int(os.getenv("SFTP_PORT", "22"))
SFTP_USER       = os.getenv("SFTP_USER",   "")
SFTP_PASS       = os.getenv("SFTP_PASS",   "")
SFTP_TARGET     = os.getenv("SFTP_TARGET", "")

SMTP_SERVER     = os.getenv("SMTP_SERVER",    "smtp.strato.de")
SMTP_PORT       = int(os.getenv("SMTP_PORT",  "587"))
SMTP_USER       = os.getenv("SMTP_USER",      "")
SMTP_PASS       = os.getenv("SMTP_PASS",      "")
SMTP_RECIPIENT  = os.getenv("SMTP_RECIPIENT", "")

# ── Tesseract ─────────────────────────────────────────────────────────────────
TESSERACT_EXE = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
if not os.path.exists(TESSERACT_EXE):
    try:
        os.makedirs(r"C:\Temp", exist_ok=True)
        with open(r"C:\Temp\dienststart.log", "a", encoding="utf-8") as f:
            f.write(f"[RE-trennerv3.py] Tesseract NICHT gefunden: {TESSERACT_EXE}\n")
    except Exception:
        pass
else:
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_EXE


# ── Logging ───────────────────────────────────────────────────────────────────
def log_error(context: str, message: str, logfile: str = "fehler.log") -> None:
    os.makedirs(LOG_FOLDER, exist_ok=True)
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(os.path.join(LOG_FOLDER, logfile), "a", encoding="utf-8") as f:
        f.write(f"{timestamp};{context};{message}\n")


# ── E-Mail ────────────────────────────────────────────────────────────────────
def send_failure_email(subject: str, filelist: list) -> None:
    msg = MIMEMultipart()
    msg['From']    = SMTP_USER
    msg['To']      = SMTP_RECIPIENT
    msg['Subject'] = subject
    body = "Folgende Dateien konnten nicht übertragen werden:\n\n" + "\n".join(filelist)
    msg.attach(MIMEText(body, 'plain'))
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USER, SMTP_PASS)
            server.sendmail(SMTP_USER, SMTP_RECIPIENT, msg.as_string())
    except Exception as e:
        log_error("EMAIL", f"E-Mail-Versand fehlgeschlagen: {e}")


# ── Drucken ───────────────────────────────────────────────────────────────────
def print_file(pdf_path: str) -> None:
    try:
        win32api.ShellExecute(0, "printto", pdf_path, f'"{PRINTER_NAME}"', ".", 0)
        log_error("PRINT", f"Datei gedruckt: {pdf_path}", "druck_debug.log")
    except Exception as e:
        log_error("PRINT", f"Fehler beim Drucken {pdf_path}: {e}", "druck_error.log")


# ── Netzwerk-Kopie ────────────────────────────────────────────────────────────
def copy_to_network(local_path: str) -> bool:
    try:
        os.makedirs(NETWORK_FOLDER, exist_ok=True)
        shutil.copy2(local_path, os.path.join(NETWORK_FOLDER, os.path.basename(local_path)))
        log_error("NETZ", f"Erfolgreich übertragen: {local_path}", "netzwerk_info.log")
        return True
    except Exception as e:
        log_error("NETZ", f"Fehler: {e}", "netzwerk_error.log")
        return False


# ── SFTP-Upload ───────────────────────────────────────────────────────────────
def upload_to_sftp(local_path: str) -> bool:
    try:
        transport = paramiko.Transport((SFTP_HOST, SFTP_PORT))
        transport.connect(username=SFTP_USER, password=SFTP_PASS)
        sftp = paramiko.SFTPClient.from_transport(transport)
        sftp.chdir(SFTP_TARGET)
        sftp.put(local_path, os.path.basename(local_path))
        sftp.close()
        transport.close()
        log_error("SFTP", f"Erfolgreich übertragen: {local_path}", "sftp_info.log")
        return True
    except Exception as e:
        log_error("SFTP", f"Fehler: {e}", "sftp_error.log")
        return False


# ── Hotfolder-Retry ───────────────────────────────────────────────────────────
def retry_hotfolder(folder: str, transfer_func, logname: str) -> None:
    """Versucht Dateien im Hotfolder erneut zu übertragen (>4 Minuten alt)."""
    failed = []
    os.makedirs(folder, exist_ok=True)
    for file in os.listdir(folder):
        full_path = os.path.join(folder, file)
        if not os.path.isfile(full_path):
            continue
        age_seconds = time.time() - os.path.getmtime(full_path)
        if age_seconds >= 240:
            if not transfer_func(full_path):
                failed.append(file)
                manuell = os.path.join(folder, "Manuell")
                os.makedirs(manuell, exist_ok=True)
                shutil.move(full_path, os.path.join(manuell, file))
            else:
                # Erfolgreich übertragen → Datei löschen
                try:
                    os.remove(full_path)
                except Exception as e:
                    log_error("HOTFOLDER", f"Löschen nach Übertragung fehlgeschlagen: {e}")
    if failed:
        send_failure_email(f"RE-Trenner Fehler bei {logname}", failed)


# ── OCR & Text-Extraktion ─────────────────────────────────────────────────────
def extract_text_with_ocr(page_image) -> str:
    return pytesseract.image_to_string(page_image, lang='deu')


def extract_invoice_and_customer(text: str):
    invoice_match  = re.search(r"\b(20\d{6})\b", text)
    customer_match = re.search(r"Kunden\s*Nr\.?\s*:?\s*(\d{5})", text)
    return (
        invoice_match.group(1)  if invoice_match  else None,
        customer_match.group(1) if customer_match else None,
    )


# ── PDF speichern + übertragen ────────────────────────────────────────────────
def save_pdf(writer: PyPDF2.PdfWriter, invoice_number: str, customer_number: str) -> None:
    jahr     = invoice_number[:4] if invoice_number else str(datetime.now().year)
    filename = f"RE-{invoice_number or 'UNBEKANNT'}"
    if customer_number:
        filename += f"-{customer_number}"
    filename += ".pdf"

    zielordner = os.path.join(ARTR_FOLDER, f"Rechnungen {jahr}")
    os.makedirs(zielordner, exist_ok=True)
    pdf_path = os.path.join(zielordner, filename)

    # Doppel-Erkennung
    if os.path.exists(pdf_path):
        doppel = os.path.join(zielordner, "Doppel")
        os.makedirs(doppel, exist_ok=True)
        pdf_path = os.path.join(doppel, filename)

    with open(pdf_path, "wb") as out:
        writer.write(out)

    print_file(pdf_path)

    if not copy_to_network(pdf_path):
        shutil.copy2(pdf_path, os.path.join(HOTFOLDER_NET, filename))

    if not upload_to_sftp(pdf_path):
        shutil.copy2(pdf_path, os.path.join(HOTFOLDER_SFTP, filename))


# ── PDF verarbeiten ───────────────────────────────────────────────────────────
def wait_for_file_ready(path: str, timeout: int = 30) -> bool:
    """Wartet bis die Datei vollständig geschrieben ist (Größe stabil)."""
    last_size = -1
    for _ in range(timeout):
        try:
            current_size = os.path.getsize(path)
        except OSError:
            time.sleep(1)
            continue
        if current_size == last_size and current_size > 0:
            return True
        last_size = current_size
        time.sleep(1)
    return False


def process_pdf(pdf_path: str) -> None:
    if not wait_for_file_ready(pdf_path):
        log_error(os.path.basename(pdf_path), "Datei nicht lesbar nach Timeout", "verarbeitung_error.log")
        return
    try:
        pages  = convert_from_path(pdf_path, dpi=300, poppler_path=POPPLER_PATH)
        reader = PyPDF2.PdfReader(pdf_path)
        writer = PyPDF2.PdfWriter()
        invoice_number = customer_number = None

        for i, page in enumerate(reader.pages):
            raw_text = page.extract_text() or ""
            text = raw_text.strip() if raw_text.strip() else extract_text_with_ocr(pages[i])

            if "Bearbeiter:" in text:
                if len(writer.pages) > 0:
                    save_pdf(writer, invoice_number, customer_number)
                    writer = PyPDF2.PdfWriter()
                invoice_number, customer_number = extract_invoice_and_customer(text)

            writer.add_page(page)

        if len(writer.pages) > 0:
            save_pdf(writer, invoice_number, customer_number)

        done_folder = os.path.join(WATCH_FOLDER, "DONE")
        os.makedirs(done_folder, exist_ok=True)
        shutil.move(pdf_path, os.path.join(done_folder, os.path.basename(pdf_path)))

    except Exception as e:
        log_error(os.path.basename(pdf_path), str(e), "verarbeitung_error.log")


# ── Haupt-Loop ────────────────────────────────────────────────────────────────
def run_main_loop(stop_event) -> None:
    log_error("SERVICE", "Dienst gestartet.", "service_debug.log")

    # Hotfolder-Reste aus letztem Lauf abarbeiten
    retry_hotfolder(HOTFOLDER_NET,  copy_to_network,  "Netzwerk")
    retry_hotfolder(HOTFOLDER_SFTP, upload_to_sftp,   "SFTP")

    os.makedirs(WATCH_FOLDER, exist_ok=True)

    # Bereits vorhandene PDFs verarbeiten
    for filename in os.listdir(WATCH_FOLDER):
        if filename.lower().endswith(".pdf"):
            process_pdf(os.path.join(WATCH_FOLDER, filename))

    class Handler(FileSystemEventHandler):
        def on_created(self, event):
            if not event.is_directory and event.src_path.lower().endswith(".pdf"):
                process_pdf(event.src_path)

    observer = Observer()
    observer.schedule(Handler(), WATCH_FOLDER, recursive=False)
    observer.start()

    try:
        while True:
            result = win32event.WaitForSingleObject(stop_event, 1000)
            if result == win32event.WAIT_OBJECT_0:
                break
    finally:
        observer.stop()
        observer.join()
        log_error("SERVICE", "Dienst sauber beendet.", "service_debug.log")


# ── Windows-Dienst ────────────────────────────────────────────────────────────
class ReTrennerService(win32serviceutil.ServiceFramework):
    _svc_name_         = "RechnungTrennerService"
    _svc_display_name_ = "RE-Trenner PDF Dienst"

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        servicemanager.LogMsg(
            servicemanager.EVENTLOG_INFORMATION_TYPE,
            servicemanager.PYS_SERVICE_STARTED,
            (self._svc_name_, ""),
        )
        run_main_loop(self.hWaitStop)


if __name__ == '__main__':
    win32serviceutil.HandleCommandLine(ReTrennerService)
