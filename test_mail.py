import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# SMTP-Konfiguration
SMTP_SERVER = "smtp.strato.de"
SMTP_PORT = 587
SMTP_USER = "Postmaster@demmehvw.de"
SMTP_PASS = "Makler99084"

def send_test_email():
    empfänger = "p.maurer@demme-immobilien.de"
    betreff = "SMTP Test erfolgreich"
    nachricht = "✅ Die Verbindung zum SMTP-Server wurde erfolgreich getestet und diese E-Mail wurde automatisch versendet."

    msg = MIMEMultipart()
    msg["From"] = SMTP_USER
    msg["To"] = empfänger
    msg["Subject"] = betreff
    msg.attach(MIMEText(nachricht, "plain"))

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USER, SMTP_PASS)
            server.send_message(msg)
        print("✅ Test-E-Mail erfolgreich an p.maurer@demme-immobilien.de gesendet.")
    except Exception as e:
        print(f"❌ Fehler beim Senden der Test-E-Mail: {e}")

send_test_email()
