import customtkinter as ctk
import threading
import serial
import time
import subprocess
import os
import re
import pandas as pd
from datetime import datetime
from serial.tools import list_ports

# ================= CONFIGURATION =================
VERSION = "v2.0" # Updated Version
ARDUINO_BAUD = 9600
ESP_BAUD = 921600
BASE_DIR = r"C:\Users\DimitrisOikonomou\Desktop\Gulliver Testing"
EXCEL_PATH = os.path.join(BASE_DIR, "Gulliver_Production_Log.xlsx")
JLINK_EXE = r"C:\Program Files\SEGGER\JLink_V866\JLink.exe"
JLINK_SCRIPT = r"C:\Users\DimitrisOikonomou\Desktop\Gulliver Testing\flash.txt"
JLINK_LOG = r"C:\Users\DimitrisOikonomou\Desktop\Gulliver Testing\logs\jlink_log.txt"

FLASH_ARGS = [
    "write_flash", "-z",
    "0x0000", r"C:\Users\DimitrisOikonomou\Desktop\Gulliver Testing\ESPFW\ESP32-C3-MINI-1-V3.2.0.0\bootloader\bootloader.bin",
    "0x60000", r"C:\Users\DimitrisOikonomou\Desktop\Gulliver Testing\ESPFW\ESP32-C3-MINI-1-V3.2.0.0\esp-at.bin",
    "0x8000", r"C:\Users\DimitrisOikonomou\Desktop\Gulliver Testing\ESPFW\ESP32-C3-MINI-1-V3.2.0.0\partition_table\partition-table.bin",
    "0xD000", r"C:\Users\DimitrisOikonomou\Desktop\Gulliver Testing\ESPFW\ESP32-C3-MINI-1-V3.2.0.0\ota_data_initial.bin",
    "0xF000", r"C:\Users\DimitrisOikonomou\Desktop\Gulliver Testing\ESPFW\ESP32-C3-MINI-1-V3.2.0.0\phy_multiple_init_data.bin",
    "0x1E000", r"C:\Users\DimitrisOikonomou\Desktop\Gulliver Testing\ESPFW\ESP32-C3-MINI-1-V3.2.0.0\at_customize.bin",
    "0x1F000", r"C:\Users\DimitrisOikonomou\Desktop\Gulliver Testing\ESPFW\ESP32-C3-MINI-1-V3.2.0.0\customized_partitions\mfg_nvs.bin",
]

TEST_LIST = [
    "RS232", "Current Transformers", "SIM / IMEI", "Modem GPRS",
    "DC IN (Power)", "Battery", "External Flash", "Accelerometer",
    "WiFi Scan", "Flowmeters (Int)", "Flowmeter (Ext)", "GPS"
]

class GulliverApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"Gulliver Test Jig Tool {VERSION}")
        self.geometry("1200x950")
        ctk.set_appearance_mode("dark")
        
        self.stop_requested = False
        self.ser = None
        self.current_process = None 
        self.test_labels = [] 
        
        # New Data Variables
        self.device_data = {"IMEI": "N/A", "ICCID": "N/A", "IMSI": "N/A", "Status": "FAIL"}
        self.current_full_log = "" 
        
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.setup_ui()

    def setup_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=2)
        self.grid_rowconfigure(0, weight=1)

        # --- LEFT PANEL ---
        self.left_panel = ctk.CTkFrame(self, corner_radius=10)
        self.left_panel.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        ctk.CTkLabel(self.left_panel, text="SETTINGS", font=("Arial", 22, "bold")).pack(pady=20)

        self.check_esp = ctk.CTkCheckBox(self.left_panel, text="1. Flash ESP Firmware"); self.check_esp.pack(pady=10, anchor="w", padx=30); self.check_esp.select()
        self.check_mcu = ctk.CTkCheckBox(self.left_panel, text="2. Flash MCU Firmware"); self.check_mcu.pack(pady=10, anchor="w", padx=30); self.check_mcu.select()
        self.check_modem = ctk.CTkCheckBox(self.left_panel, text="3. Flash Modem Firmware"); self.check_modem.pack(pady=10, anchor="w", padx=30)
        self.check_test_mode = ctk.CTkCheckBox(self.left_panel, text="4. Run Functional Tests"); self.check_test_mode.pack(pady=10, anchor="w", padx=30); self.check_test_mode.select()

        self.start_btn = ctk.CTkButton(self.left_panel, text="START TEST", fg_color="#28a745", font=("Arial", 16, "bold"), command=self.start_test_thread)
        self.start_btn.pack(pady=30, padx=30, fill="x")
        self.cancel_btn = ctk.CTkButton(self.left_panel, text="CANCEL / RESET", fg_color="#dc3545", state="disabled", command=self.request_stop)
        self.cancel_btn.pack(pady=5, padx=30, fill="x")

        # Serial Number entry starts DISABLED
        self.sn_entry = ctk.CTkEntry(self.left_panel, placeholder_text="Enter SN...", height=35, state="disabled")
        self.sn_entry.pack(pady=(40, 10), padx=30, fill="x")
        self.assign_btn = ctk.CTkButton(self.left_panel, text="COMPLETE & SAVE", state="disabled", command=self.save_all_data)
        self.assign_btn.pack(pady=10, padx=30, fill="x")

        # --- RIGHT PANEL ---
        self.right_panel = ctk.CTkFrame(self, corner_radius=10)
        self.right_panel.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")
        self.right_panel.grid_rowconfigure(1, weight=1) 
        self.right_panel.grid_columnconfigure(0, weight=1)

        self.grid_frame = ctk.CTkFrame(self.right_panel, fg_color="transparent")
        self.grid_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        
        ctk.CTkLabel(self.grid_frame, text="FUNCTIONAL TEST STATUS", font=("Arial", 16, "bold")).grid(row=0, column=0, columnspan=2, pady=10)
        
        for i, name in enumerate(TEST_LIST):
            lbl = ctk.CTkLabel(self.grid_frame, text=f" {name} ", fg_color="#333333", corner_radius=6, width=200, height=35)
            row = (i // 2) + 1
            col = i % 2
            lbl.grid(row=row, column=col, padx=5, pady=5, sticky="ew")
            self.test_labels.append(lbl)

        self.log_view = ctk.CTkTextbox(self.right_panel, font=("Consolas", 11), state="disabled")
        self.log_view.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")

    def log(self, msg):
        timestamp = datetime.now().strftime('%H:%M:%S')
        formatted_msg = f"[{timestamp}] {msg}"
        self.current_full_log += formatted_msg + "\n" # Store for TXT
        self.log_view.configure(state="normal")
        self.log_view.insert("end", formatted_msg + "\n")
        self.log_view.see("end")
        self.log_view.configure(state="disabled")

    def save_all_data(self):
        sn = self.sn_entry.get().strip()
        if not sn:
            self.log("⚠️ ERROR: Please enter Serial Number first!")
            return

        # 1. Create Folder
        date_str = datetime.now().strftime("%Y-%m-%d")
        folder_name = f"gulliver tested Devices ({date_str})"
        folder_path = os.path.join(BASE_DIR, folder_name)
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)

        # 2. Save/Append to TXT (Filename = IMEI)
        file_id = self.device_data["IMEI"] if self.device_data["IMEI"] != "N/A" else f"Unknown_IMEI_{sn}"
        txt_path = os.path.join(folder_path, f"{file_id}.txt")
        
        with open(txt_path, "a", encoding="utf-8") as f:
            f.write(f"\n{'='*50}\n")
            f.write(f"DATE: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"TESTER VERSION: {VERSION}\n")
            f.write(f"S/N: {sn} | STATUS: {self.device_data['Status']}\n")
            f.write(f"IMEI: {self.device_data['IMEI']} | ICCID: {self.device_data['ICCID']}\n")
            f.write(f"{'-'*50}\n")
            f.write(self.current_full_log)
            f.write(f"\n{'='*50}\n")

        # 3. Save to Excel
        try:
            excel_row = {
                "Date": [date_str],
                "Time": [datetime.now().strftime("%H:%M:%S")],
                "S/N": [sn],
                "IMEI": [f"'{self.device_data['IMEI']}"], # Force string
                "ICCID": [f"'{self.device_data['ICCID']}"],
                "IMSI": [f"'{self.device_data['IMSI']}"],
                "Status": [self.device_data["Status"]],
                "Tester_Ver": [VERSION]
            }
            df_new = pd.DataFrame(excel_row)
            if os.path.exists(EXCEL_PATH):
                df_old = pd.read_excel(EXCEL_PATH)
                df_final = pd.concat([df_old, df_new], ignore_index=True)
            else:
                df_final = df_new
            df_final.to_excel(EXCEL_PATH, index=False)
            self.log("✅ Data successfully exported to Excel and TXT.")
        except Exception as e:
            self.log(f"❌ Excel Error: {e}")

        self.assign_btn.configure(state="disabled")
        self.sn_entry.delete(0, 'end')
        self.sn_entry.configure(state="disabled")

    def update_test_ui(self, result_value):
        for i in range(len(TEST_LIST)):
            if (result_value & (1 << i)):
                self.test_labels[i].configure(fg_color="#28a745", text=f"✔ {TEST_LIST[i]}")

    def mark_failed_tests(self):
        for lbl in self.test_labels:
            if lbl.cget("fg_color") == "#333333":
                lbl.configure(fg_color="#941c1c")

    def detect_ports(self):
        arduino, esp = None, None
        for p in list_ports.comports():
            desc = (p.description or "").lower()
            if "arduino" in desc or (p.vid == 0x2341): arduino = p.device
            elif any(x in desc for x in ["cp210", "ch340", "ftdi", "usb serial"]): esp = p.device
        return arduino, esp

    def run_subprocess(self, cmd):
        self.current_process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
        for line in self.current_process.stdout:
            if self.stop_requested: 
                try: self.current_process.terminate()
                except: pass
                return False
            self.log(f"[EXEC] {line.strip()}")
        return self.current_process.wait() == 0

    def on_closing(self):
        self.request_stop()
        self.destroy()

    def request_stop(self):
        self.stop_requested = True
        if self.current_process:
            try: self.current_process.terminate()
            except: pass
        if self.ser and self.ser.is_open:
            try:
                self.ser.write(b"CANCEL\n")
                time.sleep(0.1)
                self.ser.close()
            except: pass
        self.start_btn.configure(state="normal")
        self.cancel_btn.configure(state="disabled")

    def start_test_thread(self):
        self.stop_requested = False
        self.current_full_log = ""
        self.device_data = {"IMEI": "N/A", "ICCID": "N/A", "IMSI": "N/A", "Status": "FAIL"}
        self.start_btn.configure(state="disabled")
        self.cancel_btn.configure(state="normal")
        self.assign_btn.configure(state="disabled")
        self.sn_entry.configure(state="disabled")
        self.log_view.configure(state="normal"); self.log_view.delete("1.0", "end"); self.log_view.configure(state="disabled")
        for lbl, name in zip(self.test_labels, TEST_LIST):
            lbl.configure(fg_color="#333333", text=name)
            
        threading.Thread(target=self.main_test_loop, daemon=True).start()

    def main_test_loop(self):
        arduino_port, esp_port = self.detect_ports()
        if not arduino_port: self.log("❌ ERROR: Arduino Port Not Found!"); self.request_stop(); return

        try:
            self.ser = serial.Serial(arduino_port, ARDUINO_BAUD, timeout=1)
            self.log("⏳ Initializing Jig...")
            time.sleep(2)
            self.ser.reset_input_buffer()
            
            self.ser.write(b'?\n')
            response = self.ser.readline().decode(errors="ignore").strip()
            if "ARDUINO_JIG_OK" not in response:
                self.log(f"❌ Handshake Failed! ({response})")
                self.request_stop(); return
            self.log("✅ Connection Established.")

            # 1. ESP Flash
            if self.check_esp.get() and not self.stop_requested:
                if not esp_port: self.log("❌ ESP Port Not Found!"); self.request_stop(); return
                self.log("🚀 Step 1: Flashing ESP..."); self.ser.write(b'P\n'); time.sleep(1)
                cmd = ["python", "-m", "esptool", "--chip", "esp32c3", "--port", esp_port, "--baud", str(ESP_BAUD)] + FLASH_ARGS
                if not self.run_subprocess(cmd): self.log("❌ ESP Flash Fail!"); self.request_stop(); return

            # 2. MCU Flash
            if self.check_mcu.get() and not self.stop_requested:
                self.log("🚀 Step 2: Flashing MCU..."); self.ser.write(b'O\n'); time.sleep(1)
                cmd = [JLINK_EXE, "-CommandFile", JLINK_SCRIPT, "-Log", JLINK_LOG]
                if not self.run_subprocess(cmd): self.request_stop(); return
                time.sleep(1)
                if not self.run_subprocess(cmd): self.request_stop(); return
                self.log("⏳ Waiting for boot..."); time.sleep(3)

            # 3. Functional Tests
            if self.check_test_mode.get() and not self.stop_requested:
                self.log("🚀 Step 4: Starting Tests...")
                self.ser.write(b'T\n')
                start_time = time.time()
                last_res = 0
                
                while (time.time() - start_time) < 180 and not self.stop_requested:
                    line = self.ser.readline().decode(errors="ignore").strip()
                    if line:
                        self.log(f"[DUT] {line}")
                        
                        # Catch Modem Info from Serial
                        if "IMEI:" in line: self.device_data["IMEI"] = line.split("IMEI:")[1].strip()
                        if "ICCID:" in line: self.device_data["ICCID"] = line.split("ICCID:")[1].strip()
                        if "IMSI:" in line: self.device_data["IMSI"] = line.split("IMSI:")[1].strip()

                        match = re.search(r"RESULT:(\d+)", line) or re.search(r"FINALRESULT:(\d+)", line)
                        if match:
                            last_res = int(match.group(1))
                            self.after(0, lambda v=last_res: self.update_test_ui(v))
                            if last_res == 4095:
                                self.device_data["Status"] = "PASS"
                                self.log("🎯 ALL TESTS PASSED!"); break
                
                if last_res != 4095 and not self.stop_requested:
                    self.device_data["Status"] = "FAIL"
                    self.log(f"❌ Incomplete Results: {last_res}")
                    self.after(0, self.mark_failed_tests)
                    # Automatically save FAIL logs even without S/N
                    self.save_all_data()

            # Unlock Serial entry ONLY if all passed
            if not self.stop_requested and last_res == 4095:
                self.ser.write(b"TESTCOMPLETED\n")
                self.after(0, lambda: self.sn_entry.configure(state="normal"))
                self.after(0, lambda: self.assign_btn.configure(state="normal"))

        except Exception as e: self.log(f"⚠️ Error: {str(e)}")
        finally:
            if self.ser and self.ser.is_open: self.ser.close()
            self.after(0, lambda: self.cancel_btn.configure(state="disabled"))
            self.after(0, lambda: self.start_btn.configure(state="normal"))

if __name__ == "__main__":
    app = GulliverApp()
    app.mainloop()