import customtkinter as ctk
from ui_main import GulliverApp

if __name__ == "__main__":
    # Προετοιμασία εμφάνισης
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")

    # Εκκίνηση της εφαρμογής
    app = GulliverApp()
    app.mainloop()
