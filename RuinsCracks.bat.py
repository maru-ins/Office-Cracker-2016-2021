import os
import sys
import subprocess
import ctypes
from colorama import Fore, init

init(autoreset=True)

# === Fungsi cek & minta admin ===
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

# Kalau tidak admin → minta UAC prompt
if not is_admin():
    # Jalankan ulang script ini sebagai admin di CMD
    script = os.path.abspath(__file__)
    params = " ".join(sys.argv[1:])
    subprocess.run(["powershell", "-Command",
                    f'Start-Process cmd -ArgumentList " python \"{script}\" {params}" -Verb RunAs'])
    sys.exit()

# === Daftar versi Office ===
office_versions = [
    "Office 2016",
    "Office 2019",
    "Office 2021",
    "Office 365"
]

# === Logo ASCII ===
logo = """
················································
: ______   __    __ _____     __      _  _____ :
:(   __ \  ) )  ( ((_   _)   /  \    / )/ ____\:
: ) (__) )( (    ) ) | |    / /\ \  / /( (___  :
:(    __/  ) )  ( (  | |    ) ) ) ) ) ) \___ \ :
: ) \ \  _( (    ) ) | |   ( ( ( ( ( (      ) ):
:( ( \ \_))) \__/ ( _| |__ / /  \ \/ /  ___/ / :
: )_) \__/ \______//_____((_/    \__/  /____/  :
················································
"""
print(Fore.CYAN + logo)
print(Fore.YELLOW + "Welcome to Ruins Office Activator\n")

# === Menu pilihan ===
for i, ver in enumerate(office_versions, 1):
    print(f"{Fore.GREEN}{i}. {ver}")

try:
    choice = int(input(Fore.CYAN + "\nPilih versi Office: "))
    if choice < 1 or choice > len(office_versions):
        print(Fore.RED + "Pilihan tidak valid!")
        input("\nTekan ENTER untuk keluar...")
        sys.exit()
except ValueError:
    print(Fore.RED + "Input harus angka!")
    input("\nTekan ENTER untuk keluar...")
    sys.exit()

selected_version = office_versions[choice - 1]
print(Fore.YELLOW + f"\nKamu memilih: {selected_version}")

# === Konfirmasi ===
confirm = input(Fore.CYAN + "Lanjutkan aktivasi? (y/n): ").strip().lower()
if confirm != "y":
    print(Fore.RED + "\nDibatalkan oleh user.")
    input("\nTekan ENTER untuk keluar...")
    sys.exit()

# === Cari file batch ===
batch_file = None
for file in os.listdir():
    if file.lower().endswith(".bat") and selected_version.lower().replace(" ", "_") in file.lower():
        batch_file = file
        break

if not batch_file:
    print(Fore.RED + f"\nFile batch untuk {selected_version} tidak ditemukan!")
else:
    print(Fore.YELLOW + f"\nMenjalankan: {batch_file}")
    subprocess.run(["cmd", "/k", batch_file], shell=True)

input(Fore.MAGENTA + "\nTekan ENTER untuk keluar...")
