import os
import sys
import subprocess
import requests
import json
import importlib.util

# === CONFIG ===
MAIN_RAW_URL = "https://raw.githubusercontent.com/clifthanger/python_project/main/main.py"
RUNNER_RAW_URL = "https://raw.githubusercontent.com/clifthanger/python_project/main/runner.py"

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
RUNNER_FILE = os.path.join(CURRENT_DIR, "runner.py")
REQ_FILE = os.path.join(CURRENT_DIR, "requirements.txt")


# === AUTO INSTALL DEPENDENCIES ===
def ensure_requirements():
    """Baca requirements.txt lalu install modul yang belum ada (auto pip + clean output)"""
    if not os.path.exists(REQ_FILE):
        print("[WARN] requirements.txt tidak ditemukan — lewati pengecekan paket.")
        return

    print("[INFO] Mengecek dependensi Python...")

    # Baca daftar paket dari file
    with open(REQ_FILE, "r", encoding="utf-8") as f:
        required_packages = [line.strip() for line in f if line.strip() and not line.startswith("#")]

    # mapping pip_name -> module_name (untuk import check)
    name_map = {
        "JayDeBeApi": "jaydebeapi",
        "python-dotenv": "dotenv",
        "oauth2client": "oauth2client",
        "gspread": "gspread",
        "tqdm": "tqdm",
        "openpyxl": "openpyxl",
    }

    missing_packages = []
    for pkg in required_packages:
        pkg_clean = pkg.split("==")[0].strip()
        module_name = name_map.get(pkg_clean, pkg_clean)
        if not importlib.util.find_spec(module_name):
            missing_packages.append(pkg_clean)

    if not missing_packages:
        print("[INFO] Semua dependensi sudah terinstall.")
        return

    print("[INFO] Paket yang belum ada:")
    for p in missing_packages:
        print(f"   - {p}")

    # --- pastikan pip tersedia ---
    try:
        import pip  # noqa
    except ImportError:
        print("[INFO] Pip belum ada, menginstall dengan ensurepip...")
        subprocess.check_call([sys.executable, "-m", "ensurepip", "--upgrade"])

    print("\n[INFO] Menginstall paket yang hilang...\n")

    try:
        for pkg in missing_packages:
            subprocess.check_call(
                [sys.executable, "-m", "pip", "install", "--quiet", pkg],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.STDOUT,
            )
            print(f"   [+] {pkg} terinstall.")
    except subprocess.CalledProcessError as e:
        print(f"[ERROR] Gagal menginstall {pkg}: {e}")
        sys.exit(1)

    print("\n[INFO] Semua paket sudah terinstall. Restarting program...\n")
    os.execv(sys.executable, [sys.executable] + sys.argv)

def ensure_json_exists():
    """Pastikan ada minimal satu file .json di root."""
    json_files = [f for f in os.listdir(CURRENT_DIR) if f.endswith(".json")]

    if not json_files:
        print("[WARN] Tidak ada file .json credentials di root!")
        dummy_file = os.path.join(CURRENT_DIR, "dummy.json")
        dummy_content = {
            "INFO": "INI FILE DUMMY. Program tidak akan jalan normal. "
                    "Segera copy-paste file credentials .json asli ke root program."
        }
        try:
            with open(dummy_file, "w", encoding="utf-8") as f:
                json.dump(dummy_content, f, indent=4, ensure_ascii=False)
            print(f"[INFO] Dummy file dibuat: {dummy_file}")
        finally:
            sys.exit(1)

    if len(json_files) == 1 and json_files[0] == "dummy.json":
        print("[WARN] Hanya ditemukan dummy.json. Harap ganti dengan file credentials asli!")
        sys.exit(1)

    print(f"[INFO] Ditemukan file JSON: {json_files}")


def download_runner():
    """Download runner.py dari GitHub RAW"""
    try:
        print("[INFO] Runner tidak ditemukan, mendownload dari Server...")
        response = requests.get(RUNNER_RAW_URL, timeout=10)
        response.raise_for_status()
        with open(RUNNER_FILE, "w", encoding="utf-8") as f:
            f.write(response.text)
        print("[INFO] Runner berhasil didownload.")
    except Exception as e:
        print(f"[ERROR] Gagal download Runner: {e}")
        sys.exit(1)


def check_update_runner():
    """Cek apakah ada update di GitHub untuk runner.py"""
    try:
        response = requests.get(RUNNER_RAW_URL, timeout=10)
        response.raise_for_status()
        remote_code = response.text

        if not os.path.exists(RUNNER_FILE):
            download_runner()
            return

        with open(RUNNER_FILE, "r", encoding="utf-8") as f:
            local_code = f.read()

        if remote_code.strip() != local_code.strip():
            print("[INFO] Update tersedia!")
            with open(RUNNER_FILE, "w", encoding="utf-8") as f:
                f.write(remote_code)
            print("[INFO] Runner berhasil diperbarui.")
        else:
            print("[INFO] Sudah versi terbaru.")
    except Exception as e:
        print(f"[WARN] Tidak bisa cek update: {e}")


def run_runner():
    """Jalankan runner.py"""
    try:
        subprocess.run([sys.executable, RUNNER_FILE], check=True)
    except subprocess.CalledProcessError as e:
        print(f"[ERROR] runner.py error: {e}")


if __name__ == "__main__":
    ensure_requirements()   # <--- Tambahan di sini
    ensure_json_exists()

    if not os.path.exists(RUNNER_FILE):
        download_runner()

    check_update_runner()
    run_runner()
