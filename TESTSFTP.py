# TEST: SFTP-Upload-Testskript
import paramiko
import os

local_path = r"C:\\Program Files\\Python_Rechnungen\\test.pdf"  # Testdatei (bitte anpassen)
SFTP_TARGET = "/users/store-6343/Buchhaltung STB/15444-40005/Rechnungsausgang"
SFTP_HOST = "sftp.hidrive.strato.com"
SFTP_PORT = 22
SFTP_USER = "store-6343"
SFTP_PASS = "Makler99084!"

try:
    transport = paramiko.Transport((SFTP_HOST, SFTP_PORT))
    transport.connect(username=SFTP_USER, password=SFTP_PASS)
    sftp = paramiko.SFTPClient.from_transport(transport)
    sftp.chdir(SFTP_TARGET)
    sftp.put(local_path, os.path.basename(local_path))
    print("✅ SFTP-Upload erfolgreich.")
    sftp.close()
    transport.close()
except Exception as e:
    print(f"❌ SFTP-Fehler: {e}")
