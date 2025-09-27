import requests
import os
import sys
import importlib
import subprocess

# URL GitHub raw file
GITHUB_RAW_URL = "https://raw.githubusercontent.com/clifthanger/python_project/refs/heads/main/main.py"
MARKER = "# === IMPORT SETELAH DIPASTIKAN ADA ==="

# Mapping pip -> module (hanya yang beda nama)
REQUIRED_PACKAGES = {
    "python-dotenv": "dotenv",
}

def check_and_install_dependencies():
    for pip_name, import_name in REQUIRED_PACKAGES.items():
        try:
            importlib.import_module(import_name)
        except ImportError:
            print(f"[INFO] Installing missing dependency: {pip_name}")
            subprocess.check_call([sys.executable, "-m", "pip", "install", pip_name])
            print(f"[INFO] ✅ {pip_name} berhasil diinstall")

def self_update():
    try:
        resp = requests.get(GITHUB_RAW_URL, timeout=10)
        resp.raise_for_status()
        remote_code = resp.text

        if MARKER not in remote_code:
            print("[WARNING] Marker tidak ditemukan di remote code, skip update.")
            return

        local_file = os.path.abspath(__file__)
        with open(local_file, "r", encoding="utf-8") as f:
            local_code = f.read()

        if MARKER not in local_code:
            print("[WARNING] Marker tidak ditemukan di local code, skip update.")
            return

        # Pisahkan berdasarkan marker
        local_before, local_after = local_code.split(MARKER, 1)
        remote_before, remote_after = remote_code.split(MARKER, 1)

        if local_after.strip() != remote_after.strip():
            print("[INFO] Update tersedia, menimpa bagian setelah marker...")
            with open(local_file, "w", encoding="utf-8") as f:
                f.write(local_before + MARKER + remote_after)
            print("[INFO] Update selesai. Silakan jalankan ulang script.")
            sys.exit(0)
        else:
            print("[INFO] Tidak ada update, lanjut eksekusi...")
    except Exception as e:
        print(f"[ERROR] Gagal cek update: {e}")

# === PANGGIL UPDATE SEBELUM IMPORT DLL ===
self_update()
check_and_install_dependencies()

# === IMPORT SETELAH DIPASTIKAN ADA ===
import time
import logging
import jaydebeapi
import gspread
from dotenv import load_dotenv
from oauth2client.service_account import ServiceAccountCredentials
from getpass import getpass
import requests
from tqdm import tqdm
# === CONFIG PATH ===
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
ENV_PATH = os.path.join(CURRENT_DIR, ".env")

# === LOAD ENV ===
load_dotenv(ENV_PATH)

# === CLEAR SCREEN ===
def clear_screen():
    os.system("cls" if os.name == "nt" else "clear")

# === LOGGING CONFIG ===
logging.basicConfig(
    level=logging.INFO,
    format="%(levelname)s: %(message)s"
)

# === INTRO LOGO ===
def display_intro():
    logo = r"""
  █████▒▄▄▄       ██▀███   ██░ ██  ▒█████   ██▀███   ██▓▒███████▒ ▒█████   ███▄    █ 
▓██   ▒▒████▄    ▓██ ▒ ██▒▓██░ ██▒▒██▒  ██▒▓██ ▒ ██▒▓██▒▒ ▒ ▒ ▄▀░▒██▒  ██▒ ██ ▀█   █ 
▒████ ░▒██  ▀█▄  ▓██ ░▄█ ▒▒██▀▀██░▒██░  ██▒▓██ ░▄█ ▒▒██▒░ ▒ ▄▀▒░ ▒██░  ██▒▓██  ▀█ ██▒
░▓█▒  ░░██▄▄▄▄██ ▒██▀▀█▄  ░▓█ ░██ ▒██   ██░▒██▀▀█▄  ░██░  ▄▀▒   ░▒██   ██░▓██▒  ▐▌██▒
░▒█░    ▓█   ▓██▒░██▓ ▒██▒░▓█▒░██▓░ ████▓▒░░██▓ ▒██▒░██░▒███████▒░ ████▓▒░▒██░   ▓██░
 ▒ ░    ▒▒   ▓▒█░░ ▒▓ ░▒▓░ ▒ ░░▒░▒░ ▒░▒░▒░ ░ ▒▓ ░▒▓░░▓  ░▒▒ ▓░▒░▒░ ▒░▒░▒░ ░ ▒░   ▒ ▒ 
 ░       ▒   ▒▒ ░  ░▒ ░ ▒░ ▒ ░▒░ ░  ░ ▒ ▒░   ░▒ ░ ▒░ ▒ ░░░▒ ▒ ░ ▒  ░ ▒ ▒░ ░ ░░   ░ ▒░
 ░ ░     ░   ▒     ░░   ░  ░  ░░ ░░ ░ ░ ▒    ░░   ░  ▒ ░░ ░ ░ ░ ░░ ░ ░ ▒     ░   ░ ░ 
             ░  ░   ░      ░  ░  ░    ░ ░     ░      ░    ░ ░        ░ ░           ░ 
                                                        ░                      V.1.2.1a
"""
    print(logo)
    time.sleep(3)
    clear_screen()

# === FILE DETECTION ===
def pick_file(extension, download_url=None, save_as=None):
    files = [f for f in os.listdir(CURRENT_DIR) if f.endswith(extension)]
    if not files:
        if download_url:
            logging.warning(f"File {extension} tidak ditemukan. Mengunduh otomatis...")

            resp = requests.get(download_url, stream=True)
            resp.raise_for_status()

            cd = resp.headers.get("content-disposition")
            if cd and "filename=" in cd:
                filename = cd.split("filename=")[1].strip('"')
            else:
                filename = save_as or f"downloaded{extension}"

            total_size = int(resp.headers.get("content-length", 0))
            block_size = 1024
            with open(filename, "wb") as f, tqdm(
                desc=filename,
                total=total_size,
                unit="B",
                unit_scale=True,
                unit_divisor=1024,
            ) as bar:
                for chunk in resp.iter_content(block_size):
                    f.write(chunk)
                    bar.update(len(chunk))

            logging.info(f"File {filename} berhasil diunduh.")
            return os.path.join(CURRENT_DIR, filename)

        else:
            logging.error(f"File {extension} tidak ditemukan di root dan tidak ada link download.")
            sys.exit(1)

    if len(files) == 1:
        logging.info(f"Otomatis pakai {files[0]}")
        time.sleep(1)
        clear_screen()
        return os.path.join(CURRENT_DIR, files[0])

    print(f"\nBeberapa file {extension} ditemukan:")
    for i, f in enumerate(files, 1):
        print(f"{i}. {f}")
    choice = int(input("Pilih nomor file: ")) - 1
    clear_screen()
    return os.path.join(CURRENT_DIR, files[choice])

def pick_sql_file():
    files = [f for f in os.listdir(CURRENT_DIR) if f.endswith(".sql")]
    if not files:
        logging.error("Tidak ada file .sql ditemukan di root.")
        sys.exit(1)
    if len(files) == 1:
        logging.info(f"Otomatis pakai {files[0]}")
        time.sleep(1)
        clear_screen()
        return os.path.join(CURRENT_DIR, files[0])

    print("\nBeberapa file SQL ditemukan:")
    for i, f in enumerate(files, 1):
        print(f"{i}. {f}")
    choice = int(input("Pilih nomor file SQL: ")) - 1
    clear_screen()
    return os.path.join(CURRENT_DIR, files[choice])

def load_sql(file_path):
    with open(file_path, "r", encoding="utf-8") as f:
        return f.read()

# === LOADING EFFECT ===
def simple_loader(message, duration=3):
    for i in range(duration * 4):
        dots = "." * (i % 4)
        sys.stdout.write(f"\r{message}{dots}   ")
        sys.stdout.flush()
        time.sleep(0.25)
    print("\r" + " " * (len(message) + 5), end="\r")

# === SETUP ===
JDBC_JAR = pick_file(
    ".jar",
    download_url="https://drive.google.com/uc?export=download&id=1ZkPcQB81FocitiQGiYFA21yA9bfGk6Ab",
    save_as="jt400.jar"
)
GOOGLE_CREDS_PATH = pick_file(".json")
SQL_FILE = pick_sql_file()
DB_URL = "jdbc:as400://10.10.1.1"
DRIVER_CLASS = "com.ibm.as400.access.AS400JDBCDriver"

# === GOOGLE SHEETS UTILITY ===
def get_google_sheets_client(json_cred_file):
    simple_loader("Membaca kredensial Google Sheets", 2)
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(json_cred_file, scope)
    return gspread.authorize(creds)

def choose_spreadsheet_and_sheet(client):
    simple_loader("Mengambil daftar spreadsheet", 3)
    spreadsheets = client.openall()
    print("\nDaftar Spreadsheet di Drive:")
    for i, ss in enumerate(spreadsheets, start=1):
        print(f"{i}. {ss.title}")
    ss_choice = int(input("Pilih nomor spreadsheet: ")) - 1
    spreadsheet = spreadsheets[ss_choice]
    clear_screen()

    worksheets = spreadsheet.worksheets()
    print(f"\nSpreadsheet '{spreadsheet.title}' punya {len(worksheets)} sheet:")
    for i, ws in enumerate(worksheets, start=1):
        print(f"{i}. {ws.title}")
    print(f"{len(worksheets)+1}. + Buat sheet baru")

    ws_choice = int(input("Pilih nomor sheet: ")) - 1
    clear_screen()

    if ws_choice == len(worksheets):
        new_name = input("Masukkan nama sheet baru: ").strip()
        rows = int(input("Jumlah baris awal (default 100): ") or 100)
        cols = int(input("Jumlah kolom awal (default 20): ") or 20)
        worksheet = spreadsheet.add_worksheet(title=new_name, rows=str(rows), cols=str(cols))
        logging.info(f"Sheet baru '{new_name}' berhasil dibuat.")
    else:
        worksheet = worksheets[ws_choice]

    start_cell = input("\nMasukkan cell awal (contoh X2, default A1): ").strip().upper()
    if not start_cell:
        start_cell = "A1"
    clear_screen()

    return spreadsheet.title, worksheet.title, start_cell, worksheet

def get_or_configure_sheet(client):
    prev_spreadsheet = os.getenv("SPREADSHEET_NAME")
    prev_sheet = os.getenv("SHEET_NAME")
    prev_cell = os.getenv("START_CELL")

    if prev_spreadsheet and prev_sheet and prev_cell:
        use_prev = input(f"\nPakai konfigurasi sebelumnya? {prev_spreadsheet} -> {prev_sheet} mulai {prev_cell} (Y/n): ").strip().lower()
        if use_prev in ("", "y", "yes"):
            clear_screen()
            spreadsheet = client.open(prev_spreadsheet)
            worksheet = spreadsheet.worksheet(prev_sheet)
            return prev_spreadsheet, prev_sheet, prev_cell, worksheet
        clear_screen()

    spreadsheet_name, sheet_name, start_cell, worksheet = choose_spreadsheet_and_sheet(client)
    with open(ENV_PATH, "w") as f:
        f.write(f"SPREADSHEET_NAME={spreadsheet_name}\n")
        f.write(f"SHEET_NAME={sheet_name}\n")
        f.write(f"START_CELL={start_cell}\n")

    return spreadsheet_name, sheet_name, start_cell, worksheet

# === MAIN EXECUTION ===
def main():
    display_intro()
    client = get_google_sheets_client(GOOGLE_CREDS_PATH)
    spreadsheet_name, sheet_name, start_cell, sheet = get_or_configure_sheet(client)

    values = sheet.get_all_values()
    is_empty = all(all(cell == "" for cell in row) for row in values)

    if not is_empty:
        confirm_clear = input(f"\nSheet '{sheet_name}' sudah ada data. Clear semua dulu? (Y/n): ").strip().lower()
        if confirm_clear in ("", "y", "yes"):
            sheet.clear()
            logging.info("Sheet dikosongkan.")
            time.sleep(1)
            clear_screen()

    db_user = input("\nUser ID AS400: ")
    db_password = getpass("Password: ")
    clear_screen()

    sql = load_sql(SQL_FILE)

    simple_loader("Menghubungkan ke AS400", 4)
    conn = jaydebeapi.connect(DRIVER_CLASS, DB_URL, [db_user, db_password], JDBC_JAR)
    curs = conn.cursor()

    logging.info("Menjalankan query...")
    with tqdm(total=1, desc="Fetching data", bar_format="{l_bar}{bar}| {n_fmt}/{total_fmt}") as pbar:
        curs.execute(sql)
        rows = curs.fetchall()
        pbar.update(1)

    columns = [desc[0] for desc in curs.description]
    data = [columns] + [list(r) for r in rows]

    logging.info(f"Uploading {len(rows)} baris ke Google Sheets...")
    with tqdm(total=1, desc="Uploading", bar_format="{l_bar}{bar}| {n_fmt}/{total_fmt}") as pbar:
        sheet.update(start_cell, data)
        pbar.update(1)

    curs.close()
    conn.close()
    clear_screen()
    logging.info(f"✅ Selesai menambahkan {len(rows)} baris ke {spreadsheet_name} pada sheet {sheet_name}, dimulai dari cell {start_cell}")

if __name__ == "__main__":
    main()
