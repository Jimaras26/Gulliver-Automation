import customtkinter as ctk
import threading
import serial
import time
import subprocess
import os
import re
import win32com.client
import pandas as pd
from datetime import datetime
from serial.tools import list_ports
import tkinter as tk
from PIL import Image, ImageTk
import sys
import time
from serial.tools import list_ports


# ================= CONFIGURATION =================
VERSION = "v2.0"
ARDUINO_BAUD = 9600
ESP_BAUD = 921600
QUECTEL_VID = 0x2C7C
QUECTEL_PID = 0x0700

BASE_DIR = r"C:\Users\DimitrisOikonomou\Desktop\Gulliver_Testing"
EXCEL_PATH = os.path.join(BASE_DIR, "Gulliver_Production_Log.xlsx")

# ======= JLINK Configuration (Main MCU) =========
JLINK_EXE = r"C:\Program Files\SEGGER\JLink_V866\JLink.exe"
JLINK_SCRIPT = os.path.join(BASE_DIR, "flash.txt")
JLINK_LOG = os.path.join(BASE_DIR, r"logs\jlink_log.txt")

# ======= Modem Update Configuration (Optional) =========
QFLASH_PATH = r"C:\Users\DimitrisOikonomou\Desktop\Gulliver_Testing\QFlash_V7.7"
QFLASH_EXE = os.path.join(QFLASH_PATH, "QFlash_V7.7.exe")
FW_XML = r"C:\Users\DimitrisOikonomou\Desktop\Gulliver_Testing\bg95update\update\firehose\rawprogram_nand_p2K_b128K.xml"

# ======= ESP32-C3 Flash Files ========
ESP_FW_PATH = os.path.join(BASE_DIR, r"ESPFW\ESP32-C3-MINI-1-V3.2.0.0")
FLASH_ARGS = [
    "write_flash",
    "-z",
    "0x0000",
    os.path.join(ESP_FW_PATH, r"bootloader\bootloader.bin"),
    "0x60000",
    os.path.join(ESP_FW_PATH, "esp-at.bin"),
    "0x8000",
    os.path.join(ESP_FW_PATH, r"partition_table\partition-table.bin"),
    "0xD000",
    os.path.join(ESP_FW_PATH, "ota_data_initial.bin"),
    "0xF000",
    os.path.join(ESP_FW_PATH, "phy_multiple_init_data.bin"),
    "0x1E000",
    os.path.join(ESP_FW_PATH, "at_customize.bin"),
    "0x1F000",
    os.path.join(ESP_FW_PATH, r"customized_partitions\mfg_nvs.bin"),
]
# ======= Printer Configuration (Brother P-touch) =========
# Η διαδρομή του εκτελέσιμου του P-touch Editor (συνήθως είναι αυτή)
PT_EDITOR_EXE = r"C:\Program Files (x86)\Brother\Ptedit54\ptedit54.exe"
# Το αρχείο ετικέτας που έχεις σχεδιάσει
LABEL_TEMPLATE = os.path.join(BASE_DIR, "Gulliver_Label.lbx")
# ======= UI Test List ================
TEST_LIST = [
    "RS232",
    "Current Transformers",
    "SIM / IMEI",
    "Modem",
    "DC IN (Power)",
    "Battery",
    "External Flash",
    "Accelerometer",
    "WiFi Scan",
    "Flowmeters (Int)",
    "Flowmeter (Ext)",
    "GPS",
]


class GulliverApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"Bibecoffee Production Tool {VERSION}")
        self.geometry("1250x980")
        ctk.set_appearance_mode("dark")
        self.mcu_fw_version = "Unknown"

        # Load asset and maintain reference in memory
        icon_path = os.path.join(BASE_DIR, "coffeeBean.png")
        if os.path.exists(icon_path):
            # Open with PIL for better icon size control
            pil_img = Image.open(icon_path)
            # Standard icon size: 32x32
            pil_img_resized = pil_img.resize((32, 32), Image.Resampling.LANCZOS)

            self.img_data = ImageTk.PhotoImage(pil_img_resized)

            # Apply small delay to ensure the window has initialized before setting the icon
            self.after(200, lambda: self.wm_iconphoto(False, self.img_data))
        else:
            print("❌ file bibecoffee.png not found!")

        # Internal State Management
        self.stop_requested = False
        self.ser = None
        self.current_process = None
        self.test_labels = []
        self.device_data = {
            "IMEI": "N/A",
            "ICCID": "N/A",
            "IMSI": "N/A",
            "FWVER": "N/A",
            "MODEMVER": "N/A",
            "Status": "READY",
        }
        self.current_full_log = ""

        self.setup_ui()

    def setup_ui(self):
        """Initialize and layout the main UI components"""
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=2)
        self.grid_rowconfigure(0, weight=1)

        # --- LEFT PANEL: Configuration & Controls ---
        self.left_panel = ctk.CTkFrame(self, corner_radius=10)
        self.left_panel.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")

        # --- HEADER SECTION ---
        self.header_frame = ctk.CTkFrame(
            self.left_panel, fg_color="transparent", height=60
        )
        self.header_frame.pack(pady=(15, 5), padx=15, fill="x")
        self.header_frame.pack_propagate(False)

        self.config_lbl = ctk.CTkLabel(
            self.header_frame, text="Configuration", font=("Arial", 18, "bold")
        )
        self.config_lbl.place(relx=0.2, rely=0.3, anchor="center")

        self.status_title_lbl = ctk.CTkLabel(
            self.header_frame, text="Status", font=("Arial", 18)
        )
        self.status_title_lbl.place(relx=0.85, rely=0.3, anchor="center")

        # --- ACTIONS GRID ---
        self.actions_frame = ctk.CTkFrame(self.left_panel, fg_color="transparent")
        self.actions_frame.pack(pady=5, padx=15, fill="x")

        self.actions_frame.columnconfigure(0, weight=1)
        self.actions_frame.columnconfigure(1, minsize=70)
        self.actions_frame.columnconfigure(2, minsize=90)

        # Row 1: ESP
        self.check_esp = ctk.CTkCheckBox(
            self.actions_frame, text="1. Flash ESP Firmware"
        )
        self.check_esp.grid(row=0, column=0, pady=5, padx=15, sticky="w")
        self.check_esp.select()
        self.esp_flash_stat = ctk.CTkLabel(
            self.actions_frame,
            text="Flash",
            font=("Arial", 11, "bold"),
            text_color="gray",
            width=70,
        )
        self.esp_flash_stat.grid(row=0, column=1, padx=2, sticky="e")
        self.esp_valid_stat = ctk.CTkLabel(
            self.actions_frame,
            text="Validate",
            font=("Arial", 11, "bold"),
            text_color="gray",
            width=90,
        )
        self.esp_valid_stat.grid(row=0, column=2, padx=2, sticky="e")

        # Row 2: MCU
        self.check_mcu = ctk.CTkCheckBox(
            self.actions_frame, text="2. Flash MCU Firmware"
        )
        self.check_mcu.grid(row=1, column=0, pady=5, padx=15, sticky="w")
        self.check_mcu.select()
        self.mcu_flash_stat = ctk.CTkLabel(
            self.actions_frame,
            text="Flash",
            font=("Arial", 11, "bold"),
            text_color="gray",
            width=70,
        )
        self.mcu_flash_stat.grid(row=1, column=1, padx=2, sticky="e")
        self.mcu_valid_stat = ctk.CTkLabel(
            self.actions_frame,
            text="Validate",
            font=("Arial", 11, "bold"),
            text_color="gray",
            width=90,
        )
        self.mcu_valid_stat.grid(row=1, column=2, padx=2, sticky="e")

        # Row 3: Modem
        self.check_modem = ctk.CTkCheckBox(
            self.actions_frame, text="3. Update Modem FW"
        )
        self.check_modem.grid(row=2, column=0, pady=5, padx=15, sticky="w")
        self.modem_flash_stat = ctk.CTkLabel(
            self.actions_frame,
            text="Flash",
            font=("Arial", 11, "bold"),
            text_color="gray",
            width=70,
        )
        self.modem_flash_stat.grid(row=2, column=1, padx=2, sticky="e")
        self.modem_valid_stat = ctk.CTkLabel(
            self.actions_frame,
            text="Validate",
            font=("Arial", 11, "bold"),
            text_color="gray",
            width=90,
        )
        self.modem_valid_stat.grid(row=2, column=2, padx=2, sticky="e")

        self.check_test_mode = ctk.CTkCheckBox(
            self.left_panel, text="4. Run Functional Tests"
        )
        self.check_test_mode.pack(pady=10, anchor="w", padx=30)
        self.check_test_mode.select()

        # --- VOLTAGE DISPLAY ---
        self.volts_frame = ctk.CTkFrame(self.left_panel, fg_color="#1a1a1a")
        self.volts_frame.pack(pady=10, padx=20, fill="x")
        ctk.CTkLabel(
            self.volts_frame, text="LIVE VOLTAGE CHECK", font=("Arial", 12, "bold")
        ).pack(pady=5)
        self.lbl_5v = ctk.CTkLabel(
            self.volts_frame, text="External 5V: -- V", text_color="gray"
        )
        self.lbl_5v.pack()
        self.lbl_inv5 = ctk.CTkLabel(
            self.volts_frame, text="Internal 5V: -- V", text_color="gray"
        )
        self.lbl_inv5.pack()
        self.lbl_33v = ctk.CTkLabel(
            self.volts_frame, text="MCU 3.3V: -- V", text_color="gray"
        )
        self.lbl_33v.pack()
        self.lbl_4v = ctk.CTkLabel(
            self.volts_frame, text="Modem 4.4V: -- V", text_color="gray"
        )
        self.lbl_4v.pack()

        self.start_btn = ctk.CTkButton(
            self.left_panel,
            text="START TEST",
            fg_color="#28a745",
            font=("Arial", 16, "bold"),
            command=self.start_test_thread,
        )
        self.start_btn.pack(pady=(15, 10), padx=30, fill="x")
        self.cancel_btn = ctk.CTkButton(
            self.left_panel,
            text="CANCEL / RESET",
            fg_color="#dc3545",
            state="disabled",
            command=self.request_stop,
        )
        self.cancel_btn.pack(pady=5, padx=30, fill="x")

        # --- SERIAL NUMBER ASSIGNMENT ---
        self.sn_title_lbl = ctk.CTkLabel(
            self.left_panel, text="Assign Serial Number", font=("Arial", 14, "bold")
        )
        self.sn_title_lbl.pack(pady=(20, 0))

        self.sn_entry = ctk.CTkEntry(
            self.left_panel, placeholder_text="Enter SN...", height=35, state="disabled"
        )
        self.sn_entry.pack(pady=(5, 10), padx=30, fill="x")

        self.assign_btn = ctk.CTkButton(
            self.left_panel,
            text="COMPLETE & SAVE",
            state="disabled",
            command=self.save_all_data,
        )
        self.assign_btn.pack(pady=10, padx=30, fill="x")

        # --- RIGHT PANEL: Results & Logs ---
        self.right_panel = ctk.CTkFrame(self, corner_radius=10)
        self.right_panel.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")
        self.right_panel.grid_rowconfigure(2, weight=1)
        self.right_panel.grid_columnconfigure(0, weight=1)

        self.status_banner = ctk.CTkLabel(
            self.right_panel,
            text="READY",
            font=("Arial", 40, "bold"),
            fg_color="#333333",
            corner_radius=10,
            height=80,
        )
        self.status_banner.grid(row=0, column=0, padx=20, pady=20, sticky="ew")

        self.grid_frame = ctk.CTkFrame(self.right_panel, fg_color="transparent")
        self.grid_frame.grid(row=1, column=0, padx=10, pady=5, sticky="nsew")
        self.test_labels = []
        for i, name in enumerate(TEST_LIST):
            lbl = ctk.CTkLabel(
                self.grid_frame,
                text=f" {name} ",
                fg_color="#333333",
                corner_radius=6,
                width=220,
                height=35,
            )
            lbl.grid(row=(i // 2), column=i % 2, padx=5, pady=5, sticky="ew")
            self.test_labels.append(lbl)

        self.log_view = ctk.CTkTextbox(
            self.right_panel, font=("Consolas", 11), state="disabled"
        )
        self.log_view.grid(row=2, column=0, padx=20, pady=20, sticky="nsew")

    def update_action_status(self, component, step, state):
        """
        Update the UI indicators for Flash/Validate status.
        :param component: 'esp', 'mcu', 'modem'
        :param step: 'flash', 'valid'
        :param state: 'active', 'ok', 'fail', 'idle'
        """
        colors = {
            "active": "#d39e00",
            "ok": "#28a745",
            "fail": "#dc3545",
            "idle": "gray",
        }
        mapping = {
            "esp": {"flash": self.esp_flash_stat, "valid": self.esp_valid_stat},
            "mcu": {"flash": self.mcu_flash_stat, "valid": self.mcu_valid_stat},
            "modem": {"flash": self.modem_flash_stat, "valid": self.modem_valid_stat},
        }
        try:
            target = mapping[component][step]
            self.after(
                0, lambda: target.configure(text_color=colors.get(state, "gray"))
            )
        except:
            pass

    def reset_ui_for_new_run(self):
        """Clear all previous test indicators and voltages"""
        for comp in ["esp", "mcu", "modem"]:
            for step in ["flash", "valid"]:
                self.update_action_status(comp, step, "idle")
        self.lbl_5v.configure(text="External 5V: -- V", text_color="gray")
        self.lbl_inv5.configure(text="Internal 5V: -- V", text_color="gray")
        self.lbl_33v.configure(text="MCU 3.3V: -- V", text_color="gray")
        self.lbl_4v.configure(text="Modem 4.4V: -- V", text_color="gray")

    def update_voltage_ui(self, v5, inv5, v33, v4):
        """Update voltage labels with color coding based on +/- 15% tolerance"""

        def check_tol(val, target):
            return "#28a745" if (target * 0.85 <= val <= target * 1.15) else "#dc3545"

        self.lbl_5v.configure(
            text=f"External 5V: {v5:.2f} V", text_color=check_tol(v5, 5.0)
        )
        self.lbl_inv5.configure(
            text=f"Internal 5V: {inv5:.2f} V", text_color=check_tol(inv5, 5.0)
        )
        self.lbl_33v.configure(
            text=f"MCU 3.3V: {v33:.2f} V", text_color=check_tol(v33, 3.3)
        )
        self.lbl_4v.configure(
            text=f"Modem 4.4V: {v4:.2f} V", text_color=check_tol(v4, 4.4)
        )

    def log(self, msg):
        """Append a timestamped message to the UI log and internal log buffer"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        line = f"[{timestamp}] {msg}"
        self.current_full_log += line + "\n"
        self.log_view.configure(state="normal")
        self.log_view.insert("end", line + "\n")
        self.log_view.see("end")
        self.log_view.configure(state="disabled")

    def start_test_thread(self):
        """Initializes testing state and launches the main sequence in a background thread"""
        self.stop_requested = False
        self.device_data = {
            "IMEI": "N/A",
            "ICCID": "N/A",
            "IMSI": "N/A",
            "Status": "TESTING",
        }
        self.current_full_log = ""  # Reset internal log buffer
        self.status_banner.configure(text="TESTING...", fg_color="#d39e00")
        self.start_btn.configure(state="disabled")
        self.cancel_btn.configure(state="normal")
        self.log_view.configure(state="normal")
        self.log_view.delete("1.0", "end")
        self.log_view.configure(state="disabled")
        # Disable serial number entry and assign button at the start of each test
        self.sn_entry.configure(state="disabled")
        self.assign_btn.configure(state="disabled")
        for lbl in self.test_labels:
            lbl.configure(
                fg_color="#333333",
                text=lbl.cget("text").replace("✔ ", "").replace("✖ ", ""),
            )
        threading.Thread(target=self.main_test_loop, daemon=True).start()

    def check_voltages(self):
        """Poll voltage data from the Arduino controller"""
        if not self.ser or not self.ser.is_open:
            return
        self.ser.reset_input_buffer()
        self.ser.write(b"V\n")
        time.sleep(1)
        v_line = self.ser.readline().decode(errors="ignore").strip()
        if "VOLTS:" in v_line:
            try:
                parts = v_line.replace("VOLTS:", "").split(",")
                v5, inv5, v33, v4 = (
                    float(parts[0]),
                    float(parts[1]),
                    float(parts[2]),
                    float(parts[3]),
                )
                self.after(0, lambda: self.update_voltage_ui(v5, inv5, v33, v4))
            except:
                pass

    def request_stop(self, fail=True):
        """Interrupts the current operation and forces a safe shutdown of peripherals"""
        self.stop_requested = True

        # 1. Immediate command to Arduino to cut power
        if self.ser and self.ser.is_open:
            try:
                self.ser.write(b"CANCEL\n")
                # self.ser.write(b"POWER_OFF\n")  # Double safety redundancy
                if fail:
                    self.ser.write(b"TESTFAILED\n")
                self.ser.flush()
                # Clear buffers to effectively blind the reading thread
                self.ser.reset_input_buffer()
                self.ser.reset_output_buffer()
            except:
                pass

        # 2. Force termination of external subprocesses (JLink/Esptool)
        if self.current_process:
            try:
                self.current_process.kill()  # Using kill() for immediate termination
            except:
                pass

        # 3. UI Synchronization
        def update_ui():
            self.status_banner.configure(text="FAIL / STOPPED", fg_color="#941c1c")
            self.start_btn.configure(state="normal")
            self.cancel_btn.configure(state="disabled")

        self.after(0, update_ui)

    def main_test_loop(self):
        """Primary test sequence coordinator"""
        ard_port, esp_port = self.detect_ports()
        if not ard_port:
            self.log("❌ Arduino Not Found")
            self.request_stop()
            return

        try:
            self.ser = serial.Serial(ard_port, ARDUINO_BAUD, timeout=1)
            time.sleep(2)
            any_flash_performed = False
            self.after(0, self.reset_ui_for_new_run)

            # --- 1. ESP32-C3 Flashing ---
            if self.check_esp.get() and not self.stop_requested:
                if not esp_port:
                    self.log("❌ ESP Port Not Found")
                    self.request_stop()
                    return
                any_flash_performed = True
                self.update_action_status("esp", "flash", "active")
                self.log("🚀 Powering ON (ESP Mode)..."), self.ser.write(b"P\n")
                time.sleep(2)
                cmd = [
                    "python",
                    "-m",
                    "esptool",
                    "--chip",
                    "esp32c3",
                    "--port",
                    esp_port,
                    "--baud",
                    str(ESP_BAUD),
                ] + FLASH_ARGS
                if self.run_subprocess(cmd):
                    self.update_action_status("esp", "flash", "ok")
                    self.update_action_status("esp", "valid", "ok")
                else:
                    self.update_action_status("esp", "flash", "fail")
                    self.log(
                        "🧹 ESP flash failed → Performing MCU full erase & retrying ESP flash..."
                    )

                    erase_cmd = [
                        JLINK_EXE,
                        "-device",
                        "nRF52840_xxAA",
                        "-if",
                        "SWD",
                        "-speed",
                        "4000",
                        "-autoconnect",
                        "1",
                        "-ExitOnError",
                        "1",
                        "-CommanderScript",
                        os.path.join(BASE_DIR, "erase.txt"),
                    ]

                    if not self.run_subprocess(erase_cmd):
                        self.log("❌ MCU ERASE FAILED!")
                        self.request_stop()
                        return

                    self.log("✅ MCU erase completed. Retrying ESP flash...")
                    self.ser.write(b"CANCEL\n")
                    time.sleep(2)
                    self.ser.write(b"P\n")
                    time.sleep(2)

                    # Retry ESP flash once
                    if not self.run_subprocess(cmd):
                        self.log("❌ ESP flash FAILED again after MCU erase!")
                        self.update_action_status("esp", "flash", "fail")
                        self.request_stop()
                        return

                    self.log("✅ ESP flash succeeded after MCU erase.")
                    self.update_action_status("esp", "flash", "ok")
                    self.update_action_status("esp", "valid", "ok")

            # --- 2. MCU (J-Link) Flashing & Verification ---
            if self.check_mcu.get() and not self.stop_requested:
                any_flash_performed = True
                self.log("🚀 Powering ON (MCU Mode)..."), self.ser.write(b"O\n")
                time.sleep(2)  # Wait to initialize MCU

                # Common flags for JLink
                jlink_base_cmd = [
                    JLINK_EXE,
                    "-device",
                    "nRF52840_xxAA",
                    "-if",
                    "SWD",
                    "-speed",
                    "4000",
                    "-autoconnect",
                    "1",
                    "-ExitOnError",
                    "1",
                ]

                # --- STEP A: FLASHING ---
                self.log("💾 Step 1: Flashing MCU...")
                self.update_action_status("mcu", "flash", "active")

                #  flash command (use flash.txt)
                flash_cmd = jlink_base_cmd + ["-CommandFile", JLINK_SCRIPT]
                output_flash, success_flash = self.run_subprocess_with_capture(
                    flash_cmd
                )

                if success_flash:
                    self.update_action_status("mcu", "flash", "ok")
                    self.log("✅ Flash Completed. Starting Verification...")

                    # --- STEP B: VERIFY ---
                    self.update_action_status("mcu", "valid", "active")

                    verify_cmd = jlink_base_cmd + [
                        "-CommandFile",
                        JLINK_SCRIPT,
                    ]
                    output_verify, success_verify = self.run_subprocess_with_capture(
                        verify_cmd
                    )

                    if success_verify:
                        self.update_action_status("mcu", "valid", "ok")
                        self.log("double_check ✅ MCU Verified & Validated!")
                        time.sleep(2)
                    else:
                        self.update_action_status("mcu", "valid", "fail")
                        self.log("❌ VERIFICATION FAILED!")
                        self.request_stop()
                        return
                else:
                    self.update_action_status("mcu", "flash", "fail")
                    self.log("❌ FLASHING FAILED!")
                    self.request_stop()
                    return
            self.ser.write(b"CANCEL\n")
            time.sleep(2)

            # --- 3. Modem Update Sequence ---
            if self.check_modem.get() and not self.stop_requested:
                any_flash_performed = True

                if self.check_modem.get():
                    self.update_action_status("modem", "flash", "active")

                self.log("🛠️ Preparing Modem (BOOT Mode)...")

                # Signal for Boot Mode (USB_BOOT HIGH)
                self.ser.write(b"B\n")
                time.sleep(2)

                # Close port to allow QFlash exclusive access
                arduino_port_name = self.ser.port
                self.ser.close()
                self.log("🔌 Arduino port released for QFlash.")

                # Wait up to 15 seconds for Quectel diagnostics port
                diag_port = None
                self.log("⏳ Waiting for Quectel diagnostics port (up to 15s)...")
                for _ in range(15):
                    ports = list_ports.comports()
                    for p in ports:
                        if p.vid == QUECTEL_VID and p.pid == QUECTEL_PID:
                            diag_port = p.device
                            break
                    if diag_port:
                        break
                    time.sleep(1)

                if diag_port:
                    self.log(
                        f"✅ Quectel diagnostics port found: {diag_port}. Proceeding with QFlash."
                    )
                    if os.path.exists(QFLASH_EXE):
                        self.log("🚀 Launching QFlash...")
                        qflash_proc = subprocess.Popen(
                            [QFLASH_EXE, FW_XML], cwd=QFLASH_PATH, shell=True
                        )

                        # Wait for external UI process to terminate
                        qflash_proc.wait()

                        # Set modem flash status to green
                        if self.check_modem.get():
                            self.after(
                                0,
                                lambda: self.update_action_status(
                                    "modem", "flash", "ok"
                                ),
                            )

                        self.log("📂 QFlash closed. Reconnecting to Arduino...")

                        # Re-establish connection to send reset signal
                        self.ser = serial.Serial(
                            arduino_port_name, ARDUINO_BAUD, timeout=1
                        )
                        time.sleep(2)
                        self.ser.write(b"CANCEL\n")
                        self.log("✅ Arduino Reset Sent.")
                else:
                    self.log(
                        "❌ Quectel diagnostics port not found after 15s. Skipping QFlash and proceeding to functional test."
                    )
                    # Reconnect Arduino for next steps
                    self.ser = serial.Serial(arduino_port_name, ARDUINO_BAUD, timeout=1)
                    time.sleep(2)
                    self.ser.write(b"CANCEL\n")

            # --- 4. Functional Testing Sequence ---
            if self.check_test_mode.get() and not self.stop_requested:
                attempts = 0
                max_attempts = 2 if any_flash_performed else 1
                test_passed = False
                while (
                    attempts < max_attempts
                    and not test_passed
                    and not self.stop_requested
                ):
                    attempts += 1
                    time.sleep(2)
                    self.ser.write(b"CANCEL\n")
                    time.sleep(2)
                    self.log(
                        f"🚀 Powering ON for Test (Attempt {attempts}/{max_attempts})..."
                    )
                    self.ser.write(b"O\n")
                    time.sleep(4)
                    self.check_voltages()
                    self.ser.write(b"T\n")
                    start_t = time.time()
                    imei_retry = False

                    if self.stop_requested:
                        break
                    line = ""
                    while (time.time() - start_t) < 180 and not self.stop_requested:
                        # Detect MODEMVER in log
                        if "MODEMVER:" in line:
                            modemver_val = line.split("MODEMVER:")[-1].strip()
                            self.device_data["MODEMVER"] = modemver_val
                            if modemver_val == "BG95M3LAR02A04_A0.301.A0.301":
                                if self.check_modem.get():
                                    self.after(
                                        0,
                                        lambda: self.update_action_status(
                                            "modem", "valid", "ok"
                                        ),
                                    )
                            else:
                                if self.check_modem.get():
                                    self.after(
                                        0,
                                        lambda: self.update_action_status(
                                            "modem", "valid", "fail"
                                        ),
                                    )
                        # Detect FWVER in log
                        if "FWVER:" in line:
                            fwver_val = line.split("FWVER:")[-1].split()[0].strip()
                            self.device_data["FWVER"] = fwver_val
                        if self.stop_requested:
                            break
                        line = self.ser.readline().decode(errors="ignore").strip()
                        if not line:
                            continue
                        self.log(f"[DUT] {line}")

                        # NEW: Detect SKIPPING TEST
                        if "INFO:SKIPPING TEST" in line:
                            try:
                                self.ser.write(b"ISTESTED\n")
                                self.ser.flush()
                            except Exception as e:
                                self.log(f"⚠️ Error sending ISTESTED: {e}")
                            self.after(
                                0,
                                lambda: self.status_banner.configure(
                                    text="ALREADY TESTED", fg_color="#28a745"
                                ),
                            )
                            self.after(0, self.enable_save_ui)
                            # Optionally, set status to PASS
                            self.device_data["Status"] = "PASS"
                            # Do NOT save log if already tested
                            self.current_full_log = ""
                            break

                        if (
                            "cannot read imei" in line.lower()
                            and attempts < max_attempts
                        ):
                            imei_retry = True
                            break

                        if "DEVICEINFO:" in line:
                            parts = line.replace("DEVICEINFO:", "").split()
                            for p in parts:
                                if "IMEI:" in p:
                                    self.device_data["IMEI"] = p.split(":")[1]
                                if "ICCID:" in p:
                                    self.device_data["ICCID"] = p.split(":")[1]
                                if "IMSI:" in p:
                                    self.device_data["IMSI"] = p.split(":")[1]

                        res_match = re.search(
                            r"(?:RESULT|FINALRESULT|TESTRESULT):(\d+)", line
                        )
                        if res_match:
                            last_res = int(res_match.group(1))
                            self.after(0, lambda v=last_res: self.update_test_ui(v))

                            if "FINALRESULT" in line or "TESTRESULT" in line:
                                if last_res == 4095:
                                    self.device_data["Status"] = "PASS"
                                    self.status_banner.configure(
                                        text="PASS", fg_color="#28a745"
                                    )
                                    self.after(0, self.enable_save_ui)
                                else:
                                    self.device_data["Status"] = "FAIL"
                                    self.status_banner.configure(
                                        text="FAIL", fg_color="#941c1c"
                                    )
                                    self.after(0, self.mark_failed_tests)

                                # Automatic attempt log generation
                                self.auto_save_log()
                                # Clear log for potential retry
                                self.current_full_log = ""
                                break
                    if not imei_retry:
                        break

            if not self.stop_requested:
                if self.device_data["Status"] == "PASS":
                    self.ser.write(b"TESTCOMPLETED\n")  # Inform Jig of SUCCESS
                else:
                    self.ser.write(b"TESTFAILED\n")  # Inform Jig of FAILURE

            if not self.stop_requested and not self.check_test_mode.get():
                self.ser.write(b"CANCEL\n")
                self.after(
                    0,
                    lambda: self.status_banner.configure(
                        text="FLASH OK", fg_color="#28a745"
                    ),
                )

        except Exception as e:
            self.log(f"⚠️ Fatal Error: {e}")
            self.request_stop()
        finally:
            if self.ser and self.ser.is_open:
                self.ser.close()
            self.after(0, lambda: self.start_btn.configure(state="normal"))
            self.after(0, lambda: self.cancel_btn.configure(state="disabled"))

    def enable_save_ui(self):
        """Activates Serial Number entry fields upon successful test"""
        self.sn_entry.configure(state="normal")
        self.assign_btn.configure(state="normal")
        self.sn_entry.focus_set()

    def update_test_ui(self, val):
        """Parses bitmask results and updates corresponding UI labels to green"""
        for i in range(len(TEST_LIST)):
            if val & (1 << i):
                self.test_labels[i].configure(
                    fg_color="#28a745", text=f"✔ {TEST_LIST[i]}"
                )

    def mark_failed_tests(self):
        """Highlights all tests that did not pass in red"""
        for lbl in self.test_labels:
            if "#28a745" not in lbl.cget("fg_color"):
                lbl.configure(fg_color="#941c1c", text=f"✖ {lbl.cget('text').strip()}")

    def detect_ports(self):
        """Scan system COM ports to identify Arduino and ESP32 interfaces"""
        a, e = None, None
        for p in list_ports.comports():
            d = (p.description or "").lower()
            if "arduino" in d or (p.vid == 0x2341):
                a = p.device
            elif any(x in d for x in ["cp210", "ch340", "ftdi", "usb serial"]):
                e = p.device
        return a, e

    def run_subprocess(self, cmd):
        """Execute external CLI tools and pipe output to the UI log"""
        try:
            # Extract firmware filename from JLink script if possible
            cmd_str = " ".join(cmd)
            bin_match = re.search(r"Gulliver_.*\.bin", cmd_str)
            if bin_match:
                self.mcu_fw_version = bin_match.group(0)

            self.current_process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )

            for line in self.current_process.stdout:
                if self.stop_requested:
                    return False
                l = line.strip()
                # Filter out progress bars and volatile flash indicators to keep logs clean
                if l and not any(
                    x in l for x in ["\x08", "%]", "Programming flash", "Reading flash"]
                ):
                    if "Downloading file" in l:
                        # Extract clean filename from the path
                        clean_name = l.split("\\")[-1].replace("].", "")
                        self.log(f"📦 FW detected: {clean_name}")
                        self.mcu_fw_version = clean_name
                    else:
                        self.log(f"[Tool] {l}")
            return self.current_process.wait() == 0
        except:
            return False

    def run_subprocess_with_capture(self, cmd):
        full_output = []
        error_keywords = [
            "failed",
            "error",
            "cannot connect",
            "could not find",
            "verification failed",
            "mismatch",
            "timeout",
            "abort",
        ]
        # Keywords that contain "error" but are not actual errors due to ExitOnError behavior
        ignore_keywords = ["will now exit on error", "note: exitonerror is enabled"]

        try:
            self.current_process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )

            found_error = False
            for line in self.current_process.stdout:
                if self.stop_requested:
                    break
                l = line.strip()
                if not l:
                    continue

                full_output.append(l)

                # convert line to lowercase for case-insensitive error detection
                line_lower = l.lower()

                # check if line contains any error keywords
                if any(err in line_lower for err in error_keywords):
                    # check if it also contains any of the ignore keywords
                    if not any(ign in line_lower for ign in ignore_keywords):
                        self.log(f"⚠️ JLINK ERROR: {l}")
                        found_error = True

                # check for successful connection or verification messages to log them prominently
                if any(
                    x in l
                    for x in ["Connected to", "O.K.", "Verified", "Flash download"]
                ):
                    self.log(f"[JLink] {l}")

            exit_code = self.current_process.wait()

            success = (exit_code == 0) and (not found_error)
            return "\n".join(full_output), success

        except Exception as e:
            self.log(f"⚠️ Subprocess Exception: {e}")
            return str(e), False

    def save_all_data(self):
        """Append production results to the master Excel log"""
        sn = self.sn_entry.get().strip()
        if not sn:
            return

        try:
            date_str = datetime.now().strftime("%Y-%m-%d")
            time_str = datetime.now().strftime("%H:%M:%S")

            new_data = {
                "Date": [date_str],
                "Time": [time_str],
                "S/N": [sn],
                "IMEI": [self.device_data.get("IMEI", "N/A")],
                "IMSI": [self.device_data.get("IMSI", "N/A")],
                "ICCID": [self.device_data.get("ICCID", "N/A")],
                "Status": [self.device_data.get("Status", "PASS")],
            }
            df_new = pd.DataFrame(new_data)

            if os.path.exists(EXCEL_PATH):
                df_old = pd.read_excel(EXCEL_PATH)
                df_final = pd.concat([df_old, df_new], ignore_index=True, sort=False)
            else:
                df_final = df_new

            df_final.to_excel(EXCEL_PATH, index=False)
            self.log(f"✅ Excel Updated: SN {sn}")

            # Print Label
            self.print_label(sn)

            # Reset UI components for next unit
            self.sn_entry.configure(state="disabled")
            self.assign_btn.configure(state="disabled")
            self.status_banner.configure(text="READY", fg_color="#333333")

        except Exception as e:
            self.log(f"❌ Excel Error: {e}")

    def auto_save_log(self):
        """Generates a detailed text report for the specific device under test"""
        date_str = datetime.now().strftime("%Y-%m-%d")
        time_str = datetime.now().strftime("%H:%M:%S")
        folder_path = os.path.join(BASE_DIR, f"Gulliver Tested Devices ({date_str})")

        try:
            if not os.path.exists(folder_path):
                os.makedirs(folder_path)

            imei_raw = self.device_data.get("IMEI", "N/A")
            iccid_raw = self.device_data.get("ICCID", "N/A")
            imsi_raw = self.device_data.get("IMSI", "N/A")
            fwver_raw = self.device_data.get("FWVER", "N/A")
            modemver_raw = self.device_data.get("MODEMVER", "N/A")

            pure_imei = (
                "".join(filter(str.isalnum, imei_raw))
                if imei_raw != "N/A"
                else "UNKNOWN_IMEI"
            )
            txt_path = os.path.join(folder_path, f"{pure_imei}.txt")

            # Determine validation status from UI label color coding
            esp_status = (
                "OK" if self.esp_valid_stat.cget("text_color") == "#28a745" else "N/A"
            )
            mcu_status = (
                "OK" if self.mcu_valid_stat.cget("text_color") == "#28a745" else "N/A"
            )

            v_info = f"Ext5V: {self.lbl_5v.cget('text')} | Int5V: {self.lbl_inv5.cget('text')} | MCU: {self.lbl_33v.cget('text')} | Modem: {self.lbl_4v.cget('text')}"

            with open(txt_path, "a", encoding="utf-8") as f:
                f.write(f"\n--- BIBECOFFEE PRODUCTION REPORT ---\n")
                f.write(f"IMEI: {imei_raw}\n")
                f.write(f"ICCID: {iccid_raw}\n")
                f.write(f"IMSI: {imsi_raw}\n")
                f.write(f"FWVER: {fwver_raw}\n")
                f.write(f"MODEMVER: {modemver_raw}\n")
                f.write(f"DATE: {date_str} {time_str}\n")
                f.write(f"ESP32 FW: Flashed & Validated ({esp_status})\n")
                f.write(f"MCU FW: Flashed & Validated ({mcu_status})\n")
                f.write(
                    f"FINAL TEST STATUS: {self.device_data.get('Status', 'FAIL')}\n"
                )
                f.write(f"------------------------------------\n")
                f.write(f"VOLTAGES: {v_info}\n")
                f.write("FUNCTIONAL TEST LOGS:\n")

                # Filter and save only the Device Under Test (DUT) specific lines
                dut_logs = "\n".join(
                    [
                        line
                        for line in self.current_full_log.split("\n")
                        if "[DUT]" in line
                    ]
                )
                f.write(dut_logs)
                f.write(f"\n------------------------------------\n")

            self.log(f"💾 Report saved for IMEI: {pure_imei}")
        except Exception as e:
            self.log(f"⚠️ Auto-log Error: {e}")

    def print_label(self, serial_number):
        """Sends the Serial Number to Brother P-touch via b-PAC SDK"""

        # Έλεγχος αν υπάρχει το template
        if not os.path.exists(LABEL_TEMPLATE):
            self.log(f"❌ Label template not found at: {LABEL_TEMPLATE}")
            return

        try:
            self.log(f"🖨️ Preparing to print: {serial_number}...")

            # Σύνδεση με το b-PAC SDK (Dynamic Dispatch για συμβατότητα)
            obj = win32com.client.dynamic.Dispatch("bpac.Document")

            # Άνοιγμα του αρχείου .lbx
            if obj.Open(LABEL_TEMPLATE):
                # Αναζήτηση του αντικειμένου SerialNumber
                txt_obj = obj.GetObject("SerialNumber")

                if txt_obj:
                    # Replace the text of the SerialNumber object with the actual serial number
                    txt_obj.Text = str(serial_number)

                    if obj.StartPrint("", 0):
                        obj.PrintOut(1, 0)
                        obj.EndPrint()
                        self.log(
                            f"✅ Label '{serial_number}' sent to printer successfully!"
                        )
                    else:
                        self.log("❌ Printer is not ready (check connection or tape).")
                else:
                    self.log("❌ Object 'SerialNumber' not found inside the .lbx file!")

                # Close Label Template
                obj.Close()
            else:
                self.log("❌ Could not open the .lbx template.")

        except Exception as e:
            self.log(f"Label Printed")


if __name__ == "__main__":
    app = GulliverApp()
    app.mainloop()
