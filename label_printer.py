from PIL import Image, ImageDraw, ImageFont
import os
def print_brother_label(sn, imei):
    # Δημιουργία εικόνας για το label (π.χ. 62mm x 29mm)
    # Οι διαστάσεις εξαρτώνται από το χαρτί σου
    img = Image.new('RGB', (600, 300), color=(255, 255, 255))
    d = ImageDraw.Draw(img)
    
    # Προσπάθεια φόρτωσης γραμματοσειράς
    try:
        font = ImageFont.truetype("arial.ttf", 40)
        small_font = ImageFont.truetype("arial.ttf", 25)
    except:
        font = ImageFont.load_default()
        small_font = ImageFont.load_default()

    d.text((20, 40), f"S/N: {sn}", fill=(0,0,0), font=font)
    d.text((20, 120), f"IMEI: {imei}", fill=(0,0,0), font=small_font)
    d.text((20, 200), "STATUS: PASSED", fill=(0,0,0), font=small_font)

    # Αποθήκευση προσωρινά
    label_file = "temp_label.png"
    img.save(label_file)
    
    # Αποστολή στον Default Windows Printer (Brother)
    # Χρησιμοποιούμε την εντολή 'start' των windows για εκτύπωση
    os.startfile(label_file, "print")