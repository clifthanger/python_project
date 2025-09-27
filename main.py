import os
import sys
import subprocess
import requests
import json

# === CONFIG ===
MAIN_RAW_URL = "https://raw.githubusercontent.com/clifthanger/python_project/main/main.py"
RUNNER_RAW_URL = "https://raw.githubusercontent.com/clifthanger/python_project/main/runner.py"

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
RUNNER_FILE = os.path.join(CURRENT_DIR, "runner.py")


def ensure_json_exists():
    """Pastikan ada minimal satu file .json di root. 
       Kalau tidak ada → buat dummy.json lalu exit.
       Kalau hanya dummy.json → exit juga.
    """
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
            sys.exit(1)  # langsung stop program

    # Kalau cuma ada dummy.json doang
    if len(json_files) == 1 and json_files[0] == "dummy.json":
        print("[WARN] Hanya ditemukan dummy.json. Harap ganti dengan file credentials asli!")
        sys.exit(1)

    # Kalau ada file json asli selain dummy.json → lanjut jalan
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
    ensure_json_exists()  # <- cek ada json atau bikin dummy

    if not os.path.exists(RUNNER_FILE):
        download_runner()

    check_update_runner()
    run_runner()
