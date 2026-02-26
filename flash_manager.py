import re
import subprocess
import os
import pandas as pd
from datetime import datetime
from config import JLINK_EXE, JLINK_SCRIPT, JLINK_LOG, EXCEL_PATH


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
                x in l for x in ["Connected to", "O.K.", "Verified", "Flash download"]
            ):
                self.log(f"[JLink] {l}")

        exit_code = self.current_process.wait()

        success = (exit_code == 0) and (not found_error)
        return "\n".join(full_output), success

    except Exception as e:
        self.log(f"⚠️ Subprocess Exception: {e}")
        return str(e), False


def update_excel(sn, imei, results, log_func):
    log_func("📊 Updating Excel log...")
    new_data = {
        "Timestamp": [datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
        "Serial Number": [sn],
        "IMEI": [imei],
    }
    for test, status in results.items():
        new_data[test] = [status]

    df_new = pd.DataFrame(new_data)
    if os.path.exists(EXCEL_PATH):
        df_old = pd.read_excel(EXCEL_PATH)
        df_final = pd.concat([df_old, df_new], ignore_index=True)
    else:
        df_final = df_new
    df_final.to_excel(EXCEL_PATH, index=False)
    log_func("✅ Excel updated.")


def flash_mcu(log_func):
    log_func("🚀 Flashing Main MCU...")
    cmd = [
        JLINK_EXE,
        "-device",
        "STM32L476RG",
        "-If",
        "SWD",
        "-Speed",
        "4000",
        "-CommanderScript",
        JLINK_SCRIPT,
    ]
    rc, out, err = run_subprocess_with_capture(cmd, log_func)
    return rc == 0, out
