import shutil
import os

# Testdatei und Zielordner (wie im Hauptskript)
testfile = r"C:\Program Files\Python_Rechnungen\test.pdf"
network_folder = r"\\192.168.161.111\scans\Rechnungseingang"

try:
    os.makedirs(network_folder, exist_ok=True)  # Nur wirksam, wenn Netzwerkordner gemappt ist
    target = os.path.join(network_folder, os.path.basename(testfile))
    shutil.copy2(testfile, target)
    print(f"✅ Datei erfolgreich ins Netzwerk kopiert: {target}")
except Exception as e:
    print(f"❌ Fehler beim Kopieren ins Netzwerk: {e}")
