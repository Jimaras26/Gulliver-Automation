import os
import re
import subprocess
import pandas as pd
from datetime import datetime
from serial.tools import list_ports
from config import EXCEL_PATH, BASE_DIR


def detect_ports():
    """Scan system COM ports to identify Arduino and ESP32 interfaces"""
    a, e = None, None
    for p in list_ports.comports():
        d = (p.description or "").lower()
        if "arduino" in d or (p.vid == 0x2341):
            a = p.device
        elif any(x in d for x in ["cp210", "ch340", "ftdi", "usb serial"]):
            e = p.device
    return a, e


def update_excel(sn, device_data):
    """Append production results to the master Excel log"""
    date_str = datetime.now().strftime("%Y-%m-%d")
    time_str = datetime.now().strftime("%H:%M:%S")
    new_data = {
        "Date": [date_str],
        "Time": [time_str],
        "S/N": [sn],
        "IMEI": [device_data.get("IMEI", "N/A")],
        "IMSI": [device_data.get("IMSI", "N/A")],
        "ICCID": [device_data.get("ICCID", "N/A")],
        "FWVER": [device_data.get("FWVER", "N/A")],
        "MODEMVER": [device_data.get("MODEMVER", "N/A")],
        "Status": [device_data.get("Status", "PASS")],
    }
    df_new = pd.DataFrame(new_data)
    if os.path.exists(EXCEL_PATH):
        df_old = pd.read_excel(EXCEL_PATH)
        df_final = pd.concat([df_old, df_new], ignore_index=True, sort=False)
    else:
        df_final = df_new
    df_final.to_excel(EXCEL_PATH, index=False)


def save_device_report(
    device_data, mcu_fw, full_log, voltages_info, esp_stat, mcu_stat
):
    """Generates a detailed text report for the specific device"""
    date_str = datetime.now().strftime("%Y-%m-%d")
    folder_path = os.path.join(BASE_DIR, f"Gulliver Tested Devices ({date_str})")
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

    ime_raw = device_data.get("IMEI", "N/A")
    pure_imei = (
        "".join(filter(str.isalnum, ime_raw)) if ime_raw != "N/A" else "UNKNOWN_IMEI"
    )
    txt_path = os.path.join(folder_path, f"{pure_imei}.txt")

    with open(txt_path, "a", encoding="utf-8") as f:
        f.write(f"\n--- BIBECOFFEE PRODUCTION REPORT ---\n")
        f.write(f"IMEI: {ime_raw}\nDATE: {date_str}\n")
        f.write(f"ESP32 FW: Validated ({esp_stat})\nMCU FW: {mcu_fw} ({mcu_stat})\n")
        f.write(f"FINAL STATUS: {device_data.get('Status', 'FAIL')}\n")
        f.write(f"VOLTAGES: {voltages_info}\n")
        f.write("FUNCTIONAL TEST LOGS:\n")
        f.write("\n".join([l for l in full_log.split("\n") if "[DUT]" in l]))
