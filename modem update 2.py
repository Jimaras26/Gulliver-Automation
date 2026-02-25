import serial
import subprocess
import time
from serial.tools import list_ports

# --- CONFIG ---
ARDUINO_VID = 0x2341
QUECTEL_VID = 0x2C7C
QUECTEL_PID = 0x0700 
EDL_VID = 0x05C6
EDL_PID = 0x9008

FH_LOADER = r"C:\QFlash_V7.7\QCMM\CH1\fh_loader.exe"
FW_DIR = r"C:\bg95update\update\firehose"
XML_FILE = "rawprogram_nand_p2K_b128K.xml"

def find_port(vid, pid):
    for p in list_ports.comports():
        if p.vid == vid and p.pid == pid:
            return p.device
    return None
def enter_edl_mode(com_port, baud=460800):
    print(f"🚀 Sending Qualcomm EDL command to {com_port}...")
    edl_cmd = bytes.fromhex(
        "7E 7E 26 00 4B 00 00 00 00 00 00 00 00 00 00 00 7E"
    )
    try:
        with serial.Serial(com_port, baud, timeout=1) as s:
            time.sleep(0.3)
            s.write(edl_cmd)
            s.flush()
            print("✅ EDL command sent")
    except Exception as e:
        print("❌ Failed to send EDL:", e)

def run_process():
    # 1. Arduino ON
    arduino = next((p.device for p in list_ports.comports() if p.vid == ARDUINO_VID), None)
    if not arduino: print("❌ Arduino not found!"); return

    print(f"📡 1. Arduino found. Sending 'B'...")
    with serial.Serial(arduino, 9600, timeout=1) as ser:
        time.sleep(2)
        ser.write(b'B\n')

    # 2. Waiting for Port
    print("⏳ 2. Waiting for Quectel Port...")
    target_com = None
    for _ in range(20):
        # Ψάχνουμε είτε για DP (0700) είτε για EDL (9008)
        target_com = find_port(QUECTEL_VID, QUECTEL_PID) or find_port(EDL_VID, EDL_PID)
        if target_com: break
        time.sleep(1)

    if not target_com: print("❌ Port not found."); return
    print(f"✅ Found Port: {target_com}")

    # 3. SWITCH COMMAND (Μόνο αν είναι σε DP mode)
    time.sleep(2)
    if find_port(QUECTEL_VID, QUECTEL_PID):
        enter_edl_mode(target_com, 460800)
        time.sleep(2)

    # 4. ΠΕΡΙΜΕΝΕ ΤΟ 9008
    print("⏳ 4. Waiting for 9008 mode...")
    edl_com = None
    for _ in range(15):
        edl_com = find_port(EDL_VID, EDL_PID)
        if edl_com: break
        time.sleep(1)
    
    if not edl_com: 
        print("❌ Failed to enter 9008 mode."); return

    # 5. EXECUTE FH_LOADER (Flash) - ΜΕ ΤΙΣ ΔΙΟΡΘΩΜΕΝΕΣ ΠΑΡΑΜΕΤΡΟΥΣ
    print(f"🔥 5. Starting Flash on {edl_com}...")
    
    cmd = [
        FH_LOADER,
        f"--port=\\\\.\\{edl_com}",
        f"--sendxml={XML_FILE}",
        f"--search_path={FW_DIR}",
        "--memoryname=nand",
        "--noprompt",
        "--showpercentagecomplete",
        "--zlpawarehost=1",
        "--maxpayloadsizeinbytes=16384", # ΑΥΤΟ ΕΙΝΑΙ ΤΟ ΚΛΕΙΔΙ (16KB)
        "--reset"
    ]

    process = subprocess.Popen(cmd, cwd=FW_DIR, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
    
    for line in process.stdout:
        print(f" [FH_LOG]: {line.strip()}")
    
    process.wait()

    # 6. CANCEL
    print("\n📂 6. Sending 'CANCEL' to Arduino...")
    with serial.Serial(arduino, 9600, timeout=1) as ser:
        time.sleep(2)
        ser.write(b'CANCEL\n')

if __name__ == "__main__":
    run_process()

    