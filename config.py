import os

# ================= CONFIGURATION =================
VERSION = "v2.0"
ARDUINO_BAUD = 9600
ESP_BAUD = 921600
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
