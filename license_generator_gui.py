import os
import uuid
import hashlib
import subprocess
import tkinter as tk
from tkinter import messagebox
import sys
import ctypes

# Set taskbar icon for Windows
try:
    if sys.platform == "win32":
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("FutureWave.LicenseGenerator.1")
except Exception:
    pass

# --- Constants ---
SALT = "FUTUREWAVE_SALT"
SECRET = "FUTUREWAVE_SECRET"
TARGET_SCRIPT = "PDF_PROCESSING_TOOL.py"
LICENSE_FILE = "license.key"

# --- License & Security ---
def get_mac_address():
    mac = uuid.getnode()
    return ':'.join(("%012X" % mac)[i:i+2] for i in range(0, 12, 2))

def generate_license_key(mac_address):
    raw = mac_address + SECRET
    return hashlib.sha256(raw.encode()).hexdigest().upper()

def save_license_key(key):
    with open(LICENSE_FILE, "w") as f:
        f.write(key)

def verify_password(input_password):
    hashed_input = hashlib.sha256((input_password + SALT).encode()).hexdigest()
    correct_hash = hashlib.sha256((TEMP2 + SALT).encode()).hexdigest()
    return hashed_input == correct_hash

# --- GUI Logic ---
class FutureWaveApp:
    def __init__(self, master):
        self.master = master
        master.title("Rishabh private limited Login")
        master.geometry("400x220")
        master.configure(bg="#2C3E50")
        master.resizable(False, False)

        self.label_title = tk.Label(master, text="Rishabh private limited Login", font=("Arial", 18, "bold"), fg="white", bg="#2C3E50")
        self.label_title.pack(pady=10)

        self.label_instruction = tk.Label(master, text="Enter your password to continue:", fg="white", bg="#2C3E50")
        self.label_instruction.pack()

        self.password_entry = tk.Entry(master, show="*", width=30, font=("Arial", 12))
        self.password_entry.pack(pady=10)

        # Bind Enter key to login function
        self.password_entry.bind("<Return>", lambda event: self.check_password())

        self.login_button = tk.Button(master, text="Login", command=self.check_password, bg="#2980B9", fg="white", width=20)
        self.login_button.pack(pady=10)

        self.status_label = tk.Label(master, text="", fg="yellow", bg="#2C3E50")
        self.status_label.pack()

    def check_password(self):
        password = self.password_entry.get()

        if verify_password(password):
            mac = get_mac_address()
            license_key = generate_license_key(mac)
            save_license_key(license_key)
            self.status_label.config(text="Access granted.", fg="lightgreen")
            self.master.after(1500, self.master.destroy)  # Close after 1.5 seconds
        else:
            self.status_label.config(text="Incorrect password. Access denied.", fg="red")

    def launch_script(self):
        # This method is no longer needed as the launch script will handle launching
        pass

# --- Entry Point ---
if __name__ == "__main__":
    root = tk.Tk()
    icon_path = os.path.join(os.path.dirname(sys.argv[0]), "icon.ico")
    root.iconbitmap(icon_path)
    app = FutureWaveApp(root)
    TEMP2 = "FUTUREWAVE2025"
    root.mainloop()
