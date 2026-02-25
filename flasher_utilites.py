import subprocess
import re
import os

def run_jlink(jlink_exe, script, log_path):
    cmd = [jlink_exe, "-CommandFile", script, "-Log", log_path]
    try:
        process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, 
                                 text=True, creationflags=subprocess.CREATE_NO_WINDOW)
        output, _ = process.communicate()
        success = (process.returncode == 0) and ("O.K." in output)
        return output, success
    except Exception as e:
        return str(e), False

def run_esp_flash(port, baud, args):
    cmd = ["python", "-m", "esptool", "--chip", "esp32c3", "--port", port, "--baud", str(baud)] + args
    try:
        process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, 
                                 text=True, creationflags=subprocess.CREATE_NO_WINDOW)
        output, _ = process.communicate()
        # Το esptool επιστρέφει "Hash of data verified" στο τέλος
        success = (process.returncode == 0) and ("Hash of data verified" in output)
        return output, success
    except Exception as e:
        return str(e), False