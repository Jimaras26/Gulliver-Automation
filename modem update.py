import serial
import subprocess
import time
import os
from serial.tools import list_ports

# --- Ρυθμίσεις ---
ARDUINO_VID = 0x2341
QFLASH_PATH = r"C:\QFlash_V7.7"
QFLASH_EXE = os.path.join(QFLASH_PATH, "QFlash_V7.7.exe")
FW_XML = r"C:\bg95update\update\firehose\rawprogram_nand_p2K_b128K.xml"

def run_modem_process():
    # 1. Βρες το Arduino
    arduino_port = next((p.device for p in list_ports.comports() if p.vid == ARDUINO_VID), None)
    if not arduino_port:
        print("❌ Arduino not found!"); return

    try:
        # 2. Στείλε 'B' στο Arduino
        print(f"📡 Connecting to Arduino on {arduino_port}...")
        ser = serial.Serial(arduino_port, 9600, timeout=1)
        time.sleep(2) 
        print("🛠️ Sending 'B' to Arduino...")
        ser.write(b'B\n')
        ser.close()
        
        # 3. Άνοιγμα QFlash και αυτόματη φόρτωση XML
        if os.path.exists(QFLASH_EXE):
            print(f"🚀 Launching QFlash with XML: {os.path.basename(FW_XML)}")
            # Το subprocess.Popen ξεκινά το πρόγραμμα και μας επιτρέπει να περιμένουμε το κλείσιμό του
            qflash_proc = subprocess.Popen([QFLASH_EXE, FW_XML], cwd=QFLASH_PATH, shell=True)
            
            print("\n✅ QFlash is open. Please handle the flash process manually.")
            print("⏳ The script will wait until you close the QFlash window...")
            
            # 4. Αναμονή μέχρι να κλείσει το παράθυρο από τον χρήστη
            qflash_proc.wait()
            
            # 5. Αποστολή CANCEL στο Arduino
            print("\n📂 QFlash closed. Sending 'CANCEL' to Arduino...")
            ser = serial.Serial(arduino_port, 9600, timeout=1)
            time.sleep(2)
            ser.write(b'CANCEL\n')
            ser.close()
            print("🏁 Done! Arduino reset successfully.")
        else:
            print(f"❌ QFlash not found at: {QFLASH_EXE}")

    except Exception as e:
        print(f"💥 Error: {e}")

if __name__ == "__main__":
    run_modem_process()