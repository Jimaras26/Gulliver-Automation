import customtkinter as ctk
import threading
import serial
import time
import re
import os
from PIL import Image, ImageTk
from datetime import datetime
from config import *
import printer_manager
import logic_helpers
import subprocess


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

    def log(self, msg):
        ts = datetime.now().strftime("%H:%M:%S")
        line = f"[{ts}] {msg}"
        self.current_full_log += line + "\n"
        self.log_view.configure(state="normal")
        self.log_view.insert("end", line + "\n")
        self.log_view.see("end")
        self.log_view.configure(state="disabled")

    def update_action_status(self, component, step, state):
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

    def start_test_thread(self):
        self.stop_requested = False
        self.current_full_log = ""
        self.status_banner.configure(text="TESTING...", fg_color="#d39e00")
        self.start_btn.configure(state="disabled")
        self.cancel_btn.configure(state="normal")
        threading.Thread(target=self.main_test_loop, daemon=True).start()

    def request_stop(self):
        self.stop_requested = True
        if self.ser and self.ser.is_open:
            self.ser.write(b"CANCEL\n")
            self.ser.write(b"POWER_OFF\n")
        if self.current_process:
            self.current_process.kill()
        self.status_banner.configure(text="FAIL / STOPPED", fg_color="#941c1c")
        self.start_btn.configure(state="normal")
        self.cancel_btn.configure(state="disabled")

    def main_test_loop(self):
        ard_port, esp_port = logic_helpers.detect_ports()
        if not ard_port:
            self.log("❌ Arduino Not Found")
            self.request_stop()
            return
        try:
            self.ser = serial.Serial(ard_port, ARDUINO_BAUD, timeout=1)
            time.sleep(2)
            # 1. ESP Flash
            if self.check_esp.get() and not self.stop_requested:
                if not esp_port:
                    self.log("❌ ESP Port Not Found")
                    self.request_stop()
                    return
                self.update_action_status("esp", "flash", "active")
                self.ser.write(b"P\n")
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
                    self.request_stop()
                    return
            # 2. MCU Flash
            if self.check_mcu.get() and not self.stop_requested:
                self.ser.write(b"O\n")
                time.sleep(2)
                self.update_action_status("mcu", "flash", "active")
                cmd = [
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
                    "-CommandFile",
                    JLINK_SCRIPT,
                ]
                _, success = self.run_subprocess_with_capture(cmd)
                if success:
                    self.update_action_status("mcu", "flash", "ok")
                    self.update_action_status("mcu", "valid", "ok")
                else:
                    self.update_action_status("mcu", "flash", "fail")
                    self.request_stop()
                    return
            # 3. Modem Update
            if self.check_modem.get() and not self.stop_requested:
                self.update_action_status("modem", "flash", "active")
                self.ser.write(b"B\n")
                time.sleep(2)
                self.ser.close()
                if os.path.exists(QFLASH_EXE):
                    q_proc = subprocess.Popen(
                        [QFLASH_EXE, FW_XML], cwd=QFLASH_PATH, shell=True
                    )
                    q_proc.wait()
                    self.update_action_status("modem", "flash", "ok")
                self.ser = serial.Serial(ard_port, ARDUINO_BAUD, timeout=1)
                time.sleep(2)
                self.ser.write(b"CANCEL\n")
            # 4. Functional Testing
            if self.check_test_mode.get() and not self.stop_requested:
                self.ser.write(b"O\n")
                time.sleep(4)
                self.ser.write(b"T\n")
                start_t = time.time()
                while (time.time() - start_t) < 180 and not self.stop_requested:
                    line = self.ser.readline().decode(errors="ignore").strip()
                    if not line:
                        continue
                    self.log(f"[DUT] {line}")
                    if "DEVICEINFO:" in line:
                        for p in line.replace("DEVICEINFO:", "").split():
                            if "IMEI:" in p:
                                self.device_data["IMEI"] = p.split(":")[1]
                    res_match = re.search(r"RESULT:(\d+)", line)
                    if res_match:
                        val = int(res_match.group(1))
                        self.after(0, lambda v=val: self.update_test_ui(v))
                        if val == 4095:
                            self.device_data["Status"] = "PASS"
                            self.status_banner.configure(
                                text="PASS", fg_color="#28a745"
                            )
                            self.after(0, self.enable_save_ui)
                        break
        except Exception as e:
            self.log(f"⚠️ Error: {e}")
            self.request_stop()
        finally:
            if self.ser:
                self.ser.close()
            self.after(0, lambda: self.start_btn.configure(state="normal"))

    def run_subprocess(self, cmd):
        try:
            self.current_process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
            for line in self.current_process.stdout:
                self.log(f"[Tool] {line.strip()}")
            return self.current_process.wait() == 0
        except:
            return False

    def run_subprocess_with_capture(self, cmd):
        try:
            self.current_process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
            out, _ = self.current_process.communicate()
            return out, self.current_process.returncode == 0
        except:
            return "", False

    def update_test_ui(self, val):
        for i in range(len(TEST_LIST)):
            if val & (1 << i):
                self.test_labels[i].configure(
                    fg_color="#28a745", text=f"✔ {TEST_LIST[i]}"
                )

    def enable_save_ui(self):
        self.sn_entry.configure(state="normal")
        self.assign_btn.configure(state="normal")
        self.sn_entry.focus_set()

    def save_all_data(self):
        sn = self.sn_entry.get().strip()
        if not sn:
            return
        try:
            logic_helpers.update_excel(sn, self.device_data)
            printer_manager.print_label(sn, self.log)
            v_info = f"Ext5V: {self.lbl_5v.cget('text')}"
            logic_helpers.save_device_report(
                self.device_data,
                self.mcu_fw_version,
                self.current_full_log,
                v_info,
                "OK",
                "OK",
            )
            self.log(f"✅ Data saved for SN: {sn}")
        except Exception as e:
            self.log(f"❌ Save Error: {e}")
