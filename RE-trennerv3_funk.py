# RE-TrennerV3 – Kompletter Dienst mit Extraktion, Benennung, Druck, Netzwerk- und SFTP-Transfer

import os
import time
import shutil
import smtplib
import PyPDF2
import pytesseract
import paramiko
import win32api
import win32print
import win32serviceutil
import win32service
import win32event
import servicemanager
import subprocess
import sys
from datetime import datetime
from pdf2image import convert_from_path
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

try:
    os.makedirs(r"C:\\Temp", exist_ok=True)
    with open(r"C:\\Temp\\dienststart.log", "a", encoding="utf-8") as f:
        f.write("[RE-trennerv3.py] Skript wurde importiert – Beginn der Ausführung.\n")
except Exception:
    pass

WATCH_FOLDER = r"C:\\Program Files\\Python_Rechnungen\\ARNEU"
LOG_FOLDER = r"C:\\Program Files\\Python_Rechnungen\\LOG-PY"
ARTR_FOLDER = r"C:\\Program Files\\Python_Rechnungen\\ARTR"
NETWORK_FOLDER = r"\\\\192.168.161.111\\scans\\Rechnungseingang"
HOTFOLDER_NET = r"C:\\Program Files\\Python_Rechnungen\\Hotfolder_Net"
HOTFOLDER_SFTP = r"C:\\Program Files\\Python_Rechnungen\\Hotfolder_SFTP"
PRINTER_NAME = "FFM_Kopierer_PCL"
SFTP_HOST = "sftp.hidrive.strato.com"
SFTP_PORT = 22
SFTP_USER = "store-6343"
SFTP_PASS = "Makler99084!"
SFTP_TARGET = "/users/store-6343/Buchhaltung STB/15444-40005/Rechnungsausgang"
SMTP_USER = "Postmaster@demmehvw.de"
SMTP_PASS = "Makler99084"
SMTP_SERVER = "smtp.strato.de"
EMAIL_TO = "p.maurer@demme-immobilien.de"
POPPLER_PATH = r"C:\\Program Files\\poppler\\Library\\bin"


def log_error(filename, message, logfile):
    try:
        os.makedirs(LOG_FOLDER, exist_ok=True)
        with open(os.path.join(LOG_FOLDER, logfile), "a", encoding="utf-8") as f:
            f.write(f"[{datetime.now()}] {filename}: {message}\n")
    except Exception as e:
        with open(r"C:\\Temp\\dienststart.log", "a", encoding="utf-8") as fallback:
            fallback.write(f"[{datetime.now()}] FEHLER beim Schreiben in LOG_FOLDER: {e}\n")
            fallback.write(f"{filename}: {message}\n")

def send_failure_email(subject, filelist):
    body = "Folgende Dateien konnten nicht übertragen werden:\n\n" + "\n".join(filelist)
    msg = MIMEMultipart()
    msg["From"] = SMTP_USER
    msg["To"] = EMAIL_TO
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))
    try:
        with smtplib.SMTP(SMTP_SERVER, 587) as server:
            server.starttls()
            server.login(SMTP_USER, SMTP_PASS)
            server.send_message(msg)
    except Exception as e:
        log_error("MAIL", str(e), "email_error.log")

def print_file(pdf_path):
    try:
        win32api.ShellExecute(0, "printto", pdf_path, f'"{PRINTER_NAME}"', ".", 0)
        log_error("PRINT", f"Datei gedruckt: {pdf_path}", "druck_debug.log")
    except Exception as e:
        log_error("PRINT", f"Fehler beim Drucken {pdf_path}: {str(e)}", "druck_error.log")

def upload_to_sftp(local_path):
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
        log_error("SFTP", f"Fehler: {str(e)}", "sftp_error.log")
        return False

def copy_to_network(local_path):
    try:
        shutil.copy2(local_path, os.path.join(NETWORK_FOLDER, os.path.basename(local_path)))
        log_error("NETZ", f"Erfolgreich übertragen: {local_path}", "netzwerk_info.log")
        return True
    except Exception as e:
        log_error("NETZ", f"Fehler: {str(e)}", "netzwerk_error.log")
        return False

def extract_text_with_ocr(page_image):
    return pytesseract.image_to_string(page_image, lang='deu')

def extract_invoice_and_customer(text):
    import re
    invoice_match = re.search(r"\b(20\d{6})\b", text)
    customer_match = re.search(r"Kunden\s*Nr\.?\s*:?\s*(\d{5})", text)
    return (invoice_match.group(1) if invoice_match else None,
            customer_match.group(1) if customer_match else None)

def save_pdf(writer, invoice_number, customer_number):
    jahr = invoice_number[:4] if invoice_number else str(datetime.now().year)
    filename = f"RE-{invoice_number}"
    if customer_number:
        filename += f"-{customer_number}"
    filename += ".pdf"
    zielordner = os.path.join(ARTR_FOLDER, f"Rechnungen {jahr}")
    os.makedirs(zielordner, exist_ok=True)
    pdf_path = os.path.join(zielordner, filename)

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

def retry_hotfolder(folder, transfer_func, logname):
    failed = []
    os.makedirs(folder, exist_ok=True)
    for file in os.listdir(folder):
        full_path = os.path.join(folder, file)
        if not os.path.isfile(full_path):
            continue
        if (time.time() - os.path.getmtime(full_path)) >= 3600:
            if not transfer_func(full_path):
                failed.append(file)
                manuell = os.path.join(folder, "Manuell")
                os.makedirs(manuell, exist_ok=True)
                shutil.move(full_path, os.path.join(manuell, file))
    if failed:
        send_failure_email(f"RE-Trenner Fehler bei {logname}", failed)

def process_pdf(pdf_path):
    try:
        pages = convert_from_path(pdf_path, dpi=300, poppler_path=POPPLER_PATH)
        reader = PyPDF2.PdfReader(pdf_path)
        writer = PyPDF2.PdfWriter()
        invoice_number = customer_number = None

        for i, page in enumerate(reader.pages):
            text = page.extract_text() or extract_text_with_ocr(pages[i])
            if "Bearbeiter:" in text:
                if len(writer.pages) > 0:
                    save_pdf(writer, invoice_number, customer_number)
                    writer = PyPDF2.PdfWriter()
                invoice_number, customer_number = extract_invoice_and_customer(text)
            writer.add_page(page)

        if len(writer.pages) > 0:
            save_pdf(writer, invoice_number, customer_number)

        os.makedirs(os.path.join(WATCH_FOLDER, "DONE"), exist_ok=True)
        shutil.move(pdf_path, os.path.join(WATCH_FOLDER, "DONE", os.path.basename(pdf_path)))

    except Exception as e:
        log_error(os.path.basename(pdf_path), str(e), "verarbeitung_error.log")

def run_main_loop():
    log_error("SERVICE", "Dienst wurde gestartet und main() wurde erreicht.", "service_debug.log")
    retry_hotfolder(HOTFOLDER_NET, copy_to_network, "Netzwerk")
    retry_hotfolder(HOTFOLDER_SFTP, upload_to_sftp, "SFTP")
    os.makedirs(WATCH_FOLDER, exist_ok=True)

    for filename in os.listdir(WATCH_FOLDER):
        if filename.lower().endswith(".pdf"):
            file_path = os.path.join(WATCH_FOLDER, filename)
            process_pdf(file_path)

    class Handler(FileSystemEventHandler):
        def on_created(self, event):
            if not event.is_directory and event.src_path.endswith(".pdf"):
                time.sleep(10)
                process_pdf(event.src_path)

    observer = Observer()
    observer.schedule(Handler(), WATCH_FOLDER, recursive=False)
    observer.start()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()

class ReTrennerService(win32serviceutil.ServiceFramework):
    _svc_name_ = "RechnungTrennerService"
    _svc_display_name_ = "RE-Trenner PDF Dienst"

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_, ""))
        run_main_loop()
        win32event.WaitForSingleObject(self.hWaitStop, win32event.INFINITE)

if __name__ == '__main__':
    win32serviceutil.HandleCommandLine(ReTrennerService)
