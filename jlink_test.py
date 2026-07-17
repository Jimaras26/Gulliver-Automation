import customtkinter as ctk
import threading
import subprocess
import os
import re
import sys
from datetime import datetime

VERSION = "v2.0"

if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = r"C:\Users\DimitrisOikonomou\Desktop\Gulliver_Testing"

JLINK_EXE    = r"C:\Program Files\SEGGER\JLink_V866\JLink.exe"
JLINK_SCRIPT = os.path.join(BASE_DIR, "flash1.txt")
LOCK_SCRIPT  = os.path.join(BASE_DIR, "lock_bootloader.txt")
LOCK_EXPECTED = ["fc", "e9", "00", "d8", "5d", "fc", "ff", "ff"]

JLINK_BASE = [
    JLINK_EXE,
    "-device", "ATSAMD21J18A",
    "-if",     "SWD",
    "-speed",  "4000",
    "-autoconnect", "1",
    "-ExitOnError",  "1",
]


class JLinkTestApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"JLink Test Tool {VERSION}")
        self.geometry("950x680")
        ctk.set_appearance_mode("dark")
        self.stop_requested  = False
        self.current_process = None
        self.setup_ui()

    def setup_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=2)
        self.grid_rowconfigure(0, weight=1)

        # ── Left panel ──────────────────────────────────────────────
        lp = ctk.CTkFrame(self, corner_radius=10)
        lp.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")

        ctk.CTkLabel(lp, text="JLink Test Tool",
                     font=("Arial", 18, "bold")).pack(pady=(20, 15))

        # Checkboxes
        self.check_flash = ctk.CTkCheckBox(lp, text="1. Flash MCU Firmware")
        self.check_flash.pack(pady=8, anchor="w", padx=30)
        self.check_flash.select()

        self.check_verify = ctk.CTkCheckBox(lp, text="2. Verify MCU Firmware")
        self.check_verify.pack(pady=8, anchor="w", padx=30)
        self.check_verify.select()

        self.check_lock = ctk.CTkCheckBox(lp, text="3. Lock Bootloader (BOD33)")
        self.check_lock.pack(pady=8, anchor="w", padx=30)
        self.check_lock.select()

        # Status indicators
        ind_frame = ctk.CTkFrame(lp, fg_color="#1a1a1a")
        ind_frame.pack(pady=20, padx=20, fill="x")
        ctk.CTkLabel(ind_frame, text="Step Status",
                     font=("Arial", 12, "bold")).pack(pady=5)

        self.stat_flash  = ctk.CTkLabel(ind_frame, text="Flash   : —",
                                        font=("Consolas", 12), text_color="gray")
        self.stat_flash.pack(anchor="w", padx=15, pady=2)

        self.stat_verify = ctk.CTkLabel(ind_frame, text="Verify  : —",
                                        font=("Consolas", 12), text_color="gray")
        self.stat_verify.pack(anchor="w", padx=15, pady=2)

        self.stat_lock   = ctk.CTkLabel(ind_frame, text="Lock    : —",
                                        font=("Consolas", 12), text_color="gray")
        self.stat_lock.pack(anchor="w", padx=15, pady=(2, 10))

        self.start_btn = ctk.CTkButton(
            lp, text="START", fg_color="#28a745",
            font=("Arial", 16, "bold"), command=self.start_thread)
        self.start_btn.pack(pady=(10, 5), padx=30, fill="x")

        self.cancel_btn = ctk.CTkButton(
            lp, text="CANCEL", fg_color="#dc3545",
            state="disabled", command=self.request_stop)
        self.cancel_btn.pack(pady=5, padx=30, fill="x")

        # ── Right panel ─────────────────────────────────────────────
        rp = ctk.CTkFrame(self, corner_radius=10)
        rp.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")
        rp.grid_rowconfigure(1, weight=1)
        rp.grid_columnconfigure(0, weight=1)

        self.status_banner = ctk.CTkLabel(
            rp, text="READY", font=("Arial", 40, "bold"),
            fg_color="#333333", corner_radius=10, height=80)
        self.status_banner.grid(row=0, column=0, padx=20, pady=20, sticky="ew")

        self.log_view = ctk.CTkTextbox(rp, font=("Consolas", 11), state="disabled")
        self.log_view.grid(row=1, column=0, padx=20, pady=(0, 20), sticky="nsew")

    # ── helpers ──────────────────────────────────────────────────────

    def log(self, msg):
        ts   = datetime.now().strftime("%H:%M:%S")
        line = f"[{ts}] {msg}"
        self.log_view.configure(state="normal")
        self.log_view.insert("end", line + "\n")
        self.log_view.see("end")
        self.log_view.configure(state="disabled")

    def set_stat(self, label, state):
        colors = {"ok": "#28a745", "fail": "#dc3545",
                  "active": "#d39e00", "idle": "gray"}
        icons  = {"ok": "✔", "fail": "✖", "active": "▶", "idle": "—"}
        names  = {self.stat_flash: "Flash  ", self.stat_verify: "Verify ",
                  self.stat_lock:  "Lock   "}
        label.configure(
            text=f"{names[label]}: {icons[state]}",
            text_color=colors[state])

    def start_thread(self):
        self.stop_requested = False
        self.status_banner.configure(text="RUNNING...", fg_color="#d39e00")
        self.start_btn.configure(state="disabled")
        self.cancel_btn.configure(state="normal")
        self.log_view.configure(state="normal")
        self.log_view.delete("1.0", "end")
        self.log_view.configure(state="disabled")
        for s in (self.stat_flash, self.stat_verify, self.stat_lock):
            self.set_stat(s, "idle")
        threading.Thread(target=self.run_sequence, daemon=True).start()

    def request_stop(self):
        self.stop_requested = True
        if self.current_process:
            try:
                self.current_process.kill()
            except Exception:
                pass
        self.after(0, lambda: self.status_banner.configure(
            text="STOPPED", fg_color="#941c1c"))
        self.after(0, lambda: self.start_btn.configure(state="normal"))
        self.after(0, lambda: self.cancel_btn.configure(state="disabled"))

    def finish(self, success):
        def _ui():
            if success:
                self.status_banner.configure(text="PASS", fg_color="#28a745")
            else:
                self.status_banner.configure(text="FAIL", fg_color="#941c1c")
            self.start_btn.configure(state="normal")
            self.cancel_btn.configure(state="disabled")
        self.after(0, _ui)

    # ── main sequence ────────────────────────────────────────────────

    def run_sequence(self):
        try:
            # 1. Flash
            if self.check_flash.get() and not self.stop_requested:
                self.after(0, lambda: self.set_stat(self.stat_flash, "active"))
                self.log("💾 Flashing MCU firmware...")
                _, ok = self.run_jlink(JLINK_SCRIPT)
                if not ok:
                    self.after(0, lambda: self.set_stat(self.stat_flash, "fail"))
                    self.log("❌ Flash FAILED!")
                    self.finish(False)
                    return
                self.after(0, lambda: self.set_stat(self.stat_flash, "ok"))
                self.log("✅ Flash complete.")

            # 2. Verify
            if self.check_verify.get() and not self.stop_requested:
                self.after(0, lambda: self.set_stat(self.stat_verify, "active"))
                self.log("🔍 Verifying MCU firmware...")
                _, ok = self.run_jlink(JLINK_SCRIPT)
                if not ok:
                    self.after(0, lambda: self.set_stat(self.stat_verify, "fail"))
                    self.log("❌ Verification FAILED!")
                    self.finish(False)
                    return
                self.after(0, lambda: self.set_stat(self.stat_verify, "ok"))
                self.log("✅ MCU Verified & Validated!")

            # 3. Lock bootloader
            if self.check_lock.get() and not self.stop_requested:
                self.after(0, lambda: self.set_stat(self.stat_lock, "active"))
                if not self.lock_and_verify_bootloader():
                    self.after(0, lambda: self.set_stat(self.stat_lock, "fail"))
                    self.finish(False)
                    return
                self.after(0, lambda: self.set_stat(self.stat_lock, "ok"))

            self.finish(True)

        except Exception as e:
            self.log(f"⚠️ Fatal Error: {e}")
            self.finish(False)

    def lock_and_verify_bootloader(self):
        MAX_ATTEMPTS = 3
        for attempt in range(1, MAX_ATTEMPTS + 1):
            self.log(f"🔒 Locking bootloader (attempt {attempt}/{MAX_ATTEMPTS})...")
            output, success = self.run_jlink(LOCK_SCRIPT)

            if not success:
                if attempt < MAX_ATTEMPTS:
                    self.log("⚠️ Write failed — retrying all commands from the beginning...")
                    continue
                self.log("❌ Bootloader lock FAILED after all attempts!")
                return False

            # Parse mem8 result
            for line in output.splitlines():
                if "00804000" in line.lower() and "=" in line:
                    rhs   = line.split("=", 1)[1]
                    found = [b.lower() for b in re.findall(r"[0-9a-fA-F]{2}", rhs)]
                    if found[:8] == LOCK_EXPECTED:
                        self.log("✅ Bootloader locked & verified! (FC E9 00 D8 5D FC FF FF)")
                        return True

            self.log("❌ Bootloader verification FAILED (unexpected mem8 values)!")
            return False

        self.log("❌ Bootloader lock FAILED after all attempts!")
        return False

    # ── subprocess wrapper ───────────────────────────────────────────

    def run_jlink(self, script_path):
        cmd = JLINK_BASE + ["-CommandFile", script_path]
        full_output = []
        error_kw  = ["failed", "error", "cannot connect", "could not find",
                     "verification failed", "mismatch", "timeout", "abort"]
        ignore_kw = ["will now exit on error", "note: exitonerror is enabled"]
        try:
            self.current_process = subprocess.Popen(
                cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                text=True, creationflags=subprocess.CREATE_NO_WINDOW)

            found_error = False
            for line in self.current_process.stdout:
                if self.stop_requested:
                    break
                l = line.strip()
                if not l:
                    continue
                full_output.append(l)
                ll = l.lower()
                if any(e in ll for e in error_kw) and not any(i in ll for i in ignore_kw):
                    self.log(f"⚠️ JLink: {l}")
                    found_error = True
                if any(x in l for x in ["Connected to", "O.K.", "Verified",
                                         "Flash download", "Downloading", "Writing", "="]):
                    self.log(f"[JLink] {l}")

            exit_code = self.current_process.wait()
            return "\n".join(full_output), (exit_code == 0) and not found_error

        except Exception as e:
            self.log(f"⚠️ Subprocess Error: {e}")
            return str(e), False


if __name__ == "__main__":
    app = JLinkTestApp()
    app.mainloop()
