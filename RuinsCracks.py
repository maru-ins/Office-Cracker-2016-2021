import init
import os
import subprocess
import ctypes
import sys
# import colorama from colorama
import init, Fore, Style

# Inisialisasi Colorama
init(autoreset=True)

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if not is_admin():
    # Minta ulang script dijalankan sebagai admin
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, " ".join(sys.argv), None, 1
    )
    sys.exit()

# ASCII art RUINS
RUINS_LOGO = f"""{Fore.CYAN}
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

def find_batch_file(keyword):
    """Cari file batch yang mengandung keyword di folder saat ini"""
    for file in os.listdir():
        if file.lower().endswith(".bat") and keyword.lower() in file.lower():
            return file
    return None 

def main():
    print(Style.BRIGHT + RUINS_LOGO)
    print(Fore.YELLOW + "=" * 50)
    print(Fore.GREEN + "       TOOL PILIHAN FILE BATCH (Versi Interaktif)")
    print(Fore.YELLOW + "=" * 50)
    print(Fore.WHITE + "Selamat datang! Pilih versi Microsoft Office\n")

    # Daftar pilihan
    versions = {
        "1": "2016",
        "2": "2019",
        "3": "2021"
    }

    # Tampilkan menu
    for key, value in versions.items():
        print(f"{Fore.CYAN}[{key}] {Fore.WHITE}Office {value}")

    # Input pilihan
    choice = input(Fore.YELLOW + "\nMasukkan pilihan (1-3): " + Fore.WHITE).strip()

    if choice in versions:
        version_keyword = versions[choice]
        batch_file = find_batch_file(version_keyword)

        if batch_file:
            # Konfirmasi
            confirm = input(Fore.YELLOW + f"\nJalankan '{batch_file}'? (y/n): " + Fore.WHITE).strip().lower()
            if confirm != "y":
                print(Fore.RED + "Dibatalkan oleh user.")
                return

            print(Fore.GREEN + f"\nMenjalankan {batch_file}...\n" + Fore.YELLOW + "-"*50)

            # Jalankan batch & tampilkan output real-time
            process = subprocess.Popen(
                batch_file,
                shell=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True
            )

            for line in process.stdout:
                print(Fore.WHITE + line, end="")

            process.wait()
            print(Fore.YELLOW + "-"*50)
            print(Fore.GREEN + f"Proses Office {version_keyword} selesai.")
        else:
            print(Fore.RED + f"\nFile batch dengan kata '{version_keyword}' tidak ditemukan!")
    else:
        print(Fore.RED + "\nPilihan tidak valid!")

if __name__ == "__main__":
    main()
