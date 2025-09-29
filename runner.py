# === IMPORT SETELAH DIPASTIKAN ADA ===
import os
import sys
import time
import logging
import jaydebeapi
import gspread
from dotenv import load_dotenv
from oauth2client.service_account import ServiceAccountCredentials
from getpass import getpass
import requests
from tqdm import tqdm
import subprocess
import openpyxl
import json
import re
import threading

# === CONFIG PATH ===
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
ENV_PATH = os.path.join(CURRENT_DIR, ".env")

# === LOAD ENV ===
load_dotenv(ENV_PATH)

# === SPINNER ===
def spinner(msg, stop_event):
    spinner_chars = ["X", "■", "▲", "○"]
    idx = 0
    while not stop_event.is_set():
        sys.stdout.write(f"\r{msg} {spinner_chars[idx % len(spinner_chars)]}")
        sys.stdout.flush()
        idx += 1
        time.sleep(0.1)
    # bersihin baris terakhir setelah selesai
    sys.stdout.write("\r" + " " * (len(msg) + 3) + "\r")
    sys.stdout.flush()

# === SIMPLE LOADER YANG PAKAI SPINNER ===
def simple_loader(msg, duration):
    stop_event = threading.Event()
    t = threading.Thread(target=spinner, args=(msg, stop_event))
    t.start()
    time.sleep(duration)   # durasi spinner jalan
    stop_event.set()
    t.join()
    print(f"{msg} [OK]")   # optional: tanda selesai

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
                                                        ░                    V.1.3.2
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

def output_menu():
    while True:
        print("\nPilih output data:")
        print("1. Upload ke Google Sheets")
        print("2. Simpan ke Excel (.xlsx)")
        choice = input("Pilihan (1/2): ").strip()
        if choice in ("1", "2"):
            return choice
        print("[WARN] Pilihan tidak valid, coba lagi.\n")

def save_to_excel(columns, rows, filename):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(columns)
    for r in rows:
        ws.append(r)
    wb.save(filename)
    logging.info(f"✅ Data berhasil disimpan ke file Excel: {filename}")

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

def sql_menu():
    while True:
        print("\nPilih opsi eksekusi SQL:")
        print("1. Jalankan query langsung")
        print("2. Preview query dan jalankan")
        print("3. Edit file query kemudian jalankan")
        choice = input("Pilihan (1/2/3): ").strip()

        if choice == "1":
            return "run"
        elif choice == "2":
            sql_preview = load_sql(SQL_FILE)
            clear_screen()
            print("=== ISI FILE QUERY ===\n")
            print(sql_preview)
            print("\nPilihan selanjutnya:")
            print("1. Jalankan sekarang")
            print("2. Kembali ke menu awal")
            sub_choice = input("Pilihan (1/2): ").strip()
            clear_screen()
            if sub_choice == "1":
                return "run"
            else:
                continue
        elif choice == "3":
            edit_sql_file(SQL_FILE)
            clear_screen()
            return "run"
        else:
            print("[WARN] Pilihan tidak valid, coba lagi.\n")
            continue

def edit_sql_file(sql_file):
    editor = os.environ.get("EDITOR", "notepad" if os.name == "nt" else "nano")
    print(f"[INFO] Membuka {sql_file} di editor ({editor})... Tutup Editor dan sql akan berjalan...")
    try:
        subprocess.call([editor, sql_file])
    except Exception as e:
        print(f"[ERROR] Gagal membuka editor: {e}")


# === SETUP ===
JDBC_JAR = pick_file(
    ".jar",
    download_url="https://drive.google.com/uc?export=download&id=1ZkPcQB81FocitiQGiYFA21yA9bfGk6Ab",
    save_as="jt400.jar"
)
GOOGLE_CREDS_PATH = pick_file(".json")
DB_URL = "jdbc:as400://10.10.1.1"
DRIVER_CLASS = "com.ibm.as400.access.AS400JDBCDriver"

# === GOOGLE SHEETS UTILITY ===
def read_env_dict():
    env_dict = {}
    if os.path.exists(ENV_PATH):
        with open(ENV_PATH, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                if "=" in line:
                    k, v = line.split("=", 1)
                    env_dict[k] = v
    return env_dict

def write_env_dict(env_dict):
    """Tulis seluruh dict ke .env (overwrite)."""
    lines = []
    for k, v in env_dict.items():
        lines.append(f"{k}={v}")
    with open(ENV_PATH, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + ("\n" if lines else ""))

def append_env_var(key, value):
    """Set/append a single env var (will update existing or add new)."""
    env = read_env_dict()
    env[key] = value
    write_env_dict(env)
    
def get_google_sheets_client(json_cred_file):
    simple_loader("Membaca kredensial Google Sheets", 2)
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(json_cred_file, scope)
    return gspread.authorize(creds)

def choose_spreadsheet_and_sheet(client):
    env_dict = read_env_dict()
    owner_email = env_dict.get("OWNER_EMAIL")

    simple_loader("Mengambil daftar Spreadsheet", 4)
    clear_screen()
    spreadsheets = client.openall()

    
    print("\nDaftar Spreadsheet di Drive:")
    for i, ss in enumerate(spreadsheets, start=1):
        print(f"{i}. {ss.title}")

    create_option_index = len(spreadsheets) + 1
    delete_spreadsheet_index = len(spreadsheets) + 2

    print(f"{create_option_index}. + Buat spreadsheet baru")
    print(f"{delete_spreadsheet_index}. - Hapus spreadsheet")

    ss_choice = input("Pilih nomor spreadsheet: ").strip()

    # --- Buat spreadsheet baru ---
    if ss_choice == str(create_option_index):
        base_name = input("Masukkan nama spreadsheet baru (default Export): ").strip()
        if not base_name:
            base_name = "Export"

        pattern = re.compile(rf"^{re.escape(base_name)}\s*(\d+)$", re.IGNORECASE)
        max_num = 0
        for ss in spreadsheets:
            m = pattern.match(ss.title)
            if m:
                try:
                    num = int(m.group(1))
                    max_num = max(max_num, num)
                except:
                    pass
        next_num = max_num + 1 if max_num > 0 else 1
        new_title = f"{base_name} {next_num}" if base_name.lower() == "export" else base_name

        logging.info(f"[INFO] Membuat spreadsheet baru: {new_title}")
        new_ss = client.create(new_title)
        time.sleep(1)

        if not owner_email:
            owner_email = input("Masukkan email kamu untuk akses sheet: ").strip()
            if owner_email:
                append_env_var("OWNER_EMAIL", owner_email)

        try:
            if owner_email:
                logging.info(f"[INFO] Sharing spreadsheet ke {owner_email} ...")
                new_ss.share(owner_email, perm_type='user', role='writer', notify=True)
                logging.info(f"[OK] Berhasil share ke {owner_email}")
            else:
                logging.info("[INFO] Tidak ada OWNER_EMAIL, skip sharing.")
        except Exception as e:
            logging.warning(f"[WARN] Gagal share spreadsheet ke {owner_email}: {e}")

        spreadsheet = client.open(new_title)
        clear_screen()

    # --- Hapus spreadsheet ---
    elif ss_choice == str(delete_spreadsheet_index):
        print("\n=== Daftar Spreadsheet untuk Dihapus ===")
        for i, ss in enumerate(spreadsheets, start=1):
            print(f"{i}. {ss.title}")
        del_choice = input("Masukkan nomor spreadsheet yang ingin dihapus: ").strip()
        try:
            del_index = int(del_choice) - 1
            if 0 <= del_index < len(spreadsheets):
                ss_to_delete = spreadsheets[del_index]
                confirm = input(f"Yakin mau hapus spreadsheet '{ss_to_delete.title}'? (y/n): ").strip().lower()
                if confirm == "y":
                    try:
                        client.del_spreadsheet(ss_to_delete.id)   # ✅ FIX
                        logging.info(f"Spreadsheet '{ss_to_delete.title}' berhasil dihapus.")
                    except Exception as e:
                        logging.warning(f"Gagal menghapus spreadsheet: {e}")
                        logging.warning("Pastikan service account adalah OWNER file ini agar bisa menghapus.")
                else:
                    logging.info("Pembatalan hapus spreadsheet.")
            else:
                logging.warning("Nomor spreadsheet tidak valid.")
        except ValueError:
            logging.warning("Input tidak valid.")
        time.sleep(1)
        clear_screen()
        return choose_spreadsheet_and_sheet(client)

    # --- Pilih spreadsheet yang ada ---
    else:
        try:
            idx = int(ss_choice) - 1
            if 0 <= idx < len(spreadsheets):
                spreadsheet = spreadsheets[idx]
                clear_screen()
            else:
                logging.warning("Nomor spreadsheet tidak valid.")
                return choose_spreadsheet_and_sheet(client)
        except ValueError:
            logging.warning("Input tidak valid.")
            return choose_spreadsheet_and_sheet(client)

    # --- Loop pilihan sheet ---
    while True:
        worksheets = spreadsheet.worksheets()
        print(f"\nSpreadsheet '{spreadsheet.title}' punya {len(worksheets)} sheet:")
        for i, ws in enumerate(worksheets, start=1):
            print(f"{i}. {ws.title}")
        print(f"{len(worksheets)+1}. + Buat sheet baru")
        print(f"{len(worksheets)+2}. - Hapus sheet")

        ws_choice = input("Pilih nomor sheet: ").strip()

        # --- Buat sheet baru ---
        if ws_choice == str(len(worksheets)+1):
            new_name = input("Masukkan nama sheet baru: ").strip()
            rows = int(input("Jumlah baris awal (default 100): ") or 100)
            cols = int(input("Jumlah kolom awal (default 20): ") or 20)
            worksheet = spreadsheet.add_worksheet(title=new_name, rows=str(rows), cols=str(cols))
            logging.info(f"Sheet baru '{new_name}' berhasil dibuat.")

            start_cell = input("\nMasukkan cell awal (contoh X2, default A1): ").strip().upper()
            if not start_cell:
                start_cell = "A1"
            clear_screen()
            return spreadsheet.title, worksheet.title, start_cell, worksheet

        # --- Hapus sheet ---
        elif ws_choice == str(len(worksheets)+2):
            if len(worksheets) == 1:
                logging.warning("Tidak bisa menghapus semua sheet. Spreadsheet harus minimal punya 1 sheet.")
                time.sleep(1)
                clear_screen()
                continue

            print("\n=== Daftar Sheet untuk Dihapus ===")
            for i, ws in enumerate(worksheets, start=1):
                print(f"{i}. {ws.title}")
            del_choice = input("Masukkan nomor sheet yang ingin dihapus: ").strip()
            try:
                del_index = int(del_choice) - 1
                if 0 <= del_index < len(worksheets):
                    sheet_to_delete = worksheets[del_index]
                    confirm = input(f"Yakin mau hapus sheet '{sheet_to_delete.title}'? (y/n): ").strip().lower()
                    if confirm == "y":
                        try:
                            spreadsheet.del_worksheet(sheet_to_delete)
                            logging.info(f"Sheet '{sheet_to_delete.title}' berhasil dihapus.")
                        except Exception as e:
                            logging.warning(f"Gagal menghapus sheet: {e}")
                    else:
                        logging.info("Pembatalan hapus sheet.")
                else:
                    logging.warning("Nomor sheet tidak valid.")
            except ValueError:
                logging.warning("Input tidak valid.")
            time.sleep(1)
            clear_screen()
            continue

        # --- Pilih sheet ---
        else:
            try:
                idx = int(ws_choice) - 1
                if 0 <= idx < len(worksheets):
                    worksheet = worksheets[idx]
                    clear_screen()
                    break
                else:
                    logging.warning("Nomor sheet tidak valid.")
            except ValueError:
                logging.warning("Input tidak valid.")
            time.sleep(1)
            clear_screen()
            continue

    start_cell = input("\nMasukkan cell awal (contoh X2, default A1): ").strip().upper()
    if not start_cell:
        start_cell = "A1"
    clear_screen()
    return spreadsheet.title, worksheet.title, start_cell, worksheet

def get_or_configure_sheet(client):
    env_dict = read_env_dict()
    prev_spreadsheet = env_dict.get("SPREADSHEET_NAME")
    prev_sheet = env_dict.get("SHEET_NAME")
    prev_cell = env_dict.get("START_CELL")

    if prev_spreadsheet and prev_sheet and prev_cell:
        use_prev = input(f"\nPakai konfigurasi sebelumnya? {prev_spreadsheet} -> {prev_sheet} mulai {prev_cell} (Y/n): ").strip().lower()
        if use_prev in ("", "y", "yes"):
            clear_screen()
            try:
                spreadsheet = client.open(prev_spreadsheet)
                worksheet = spreadsheet.worksheet(prev_sheet)
                return prev_spreadsheet, prev_sheet, prev_cell, worksheet
            except Exception as e:
                logging.warning(f"Gagal membuka konfigurasi sebelumnya: {e}")
                # fallback to interactive selection
        clear_screen()

    # Interaktif pilih/buat spreadsheet & sheet
    spreadsheet_name, sheet_name, start_cell, worksheet = choose_spreadsheet_and_sheet(client)

    # baca .env (lagi) untuk menjaga nilai lain, lalu update keys yang berubah
    env = read_env_dict()
    env["SPREADSHEET_NAME"] = spreadsheet_name
    env["SHEET_NAME"] = sheet_name
    env["START_CELL"] = start_cell
    write_env_dict(env)

    return spreadsheet_name, sheet_name, start_cell, worksheet
# === MAIN EXECUTION ===
def main():
    display_intro()

    print("Pilih mode output data:")
    print("1. Upload ke Google Sheets")
    print("2. Simpan ke Excel (xlsx)")
    mode = input("Pilihan (1/2): ").strip()
    use_gsheet = (mode == "1")
    clear_screen()

    if use_gsheet:
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
    else:
        spreadsheet_name = None
        sheet_name = None
        start_cell = "A1"
        sheet = None
        excel_filename = input("\nMasukkan nama file Excel output (tanpa .xlsx): ").strip()
        if not excel_filename:
            excel_filename = "output"
        excel_filename = excel_filename + ".xlsx"
        clear_screen()

    global SQL_FILE
    SQL_FILE = pick_sql_file()

    action = sql_menu()
    if action != "run":
        sys.exit(0)

    clear_screen()

    db_user = input("\nUser ID AS400: ")
    db_password = getpass("Password: ")
    clear_screen()

    sql = load_sql(SQL_FILE)

    simple_loader("Menghubungkan ke AS400", 4)
    conn = jaydebeapi.connect(DRIVER_CLASS, DB_URL, [db_user, db_password], JDBC_JAR)
    curs = conn.cursor()
    clear_screen()

    logging.info("Menjalankan query...")
    with tqdm(total=1, desc="Fetching data", bar_format="{l_bar}{bar}| {n_fmt}/{total_fmt}") as pbar:
        curs.execute(sql)
        rows = curs.fetchall()
        pbar.update(1)
    clear_screen()

    columns = [desc[0] for desc in curs.description]
    data = [columns] + [list(r) for r in rows]

    if use_gsheet:
        logging.info(f"Uploading {len(rows)} baris ke Google Sheets...")
        with tqdm(total=1, desc="Uploading", bar_format="{l_bar}{bar}| {n_fmt}/{total_fmt}") as pbar:
            sheet.update(start_cell, data)
            pbar.update(1)
        logging.info(f"[OK] Selesai menambahkan {len(rows)} baris ke {spreadsheet_name} pada sheet {sheet_name}, dimulai dari cell {start_cell}")
    else:
        logging.info(f"Menyimpan {len(rows)} baris ke file Excel: {excel_filename}")
        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Data"
            for row in data:
                ws.append(row)
            wb.save(excel_filename)
            logging.info(f"[OK] Data berhasil disimpan ke {excel_filename}")
        except Exception as e:
            logging.error(f"Gagal menyimpan ke Excel: {e}")

    curs.close()
    conn.close()
    #clear_screen()

if __name__ == "__main__":
    main()
