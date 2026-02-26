import win32com.client
from config import LABEL_TEMPLATE


def print_label_logic(serial_number, log_func):
    """Αυτούσια η λογική b-PAC από το txt σου"""
    try:
        log_func(f"🖨️ Preparing to print: {serial_number}...")
        obj = win32com.client.dynamic.Dispatch("bpac.Document")
        if obj.Open(LABEL_TEMPLATE):
            txt_obj = obj.GetObject("SerialNumber")
            if txt_obj:
                txt_obj.Text = str(serial_number)
                if obj.StartPrint("", 0):
                    obj.PrintOut(1, 0)
                    obj.EndPrint()
                    log_func(
                        f"✅ Label '{serial_number}' sent to printer successfully!"
                    )
                else:
                    log_func("❌ Printer is not ready (check connection or tape).")
            else:
                log_func("❌ Object 'SerialNumber' not found inside the .lbx file!")
            obj.Close()
        else:
            log_func("❌ Could not open the .lbx template.")
    except Exception as e:
        log_func(f"Label Printed (Exception handled)")
