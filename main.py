import os
import sys
import subprocess
import requests

# === CONFIG ===
MAIN_RAW_URL = "https://raw.githubusercontent.com/clifthanger/python_project/main/main.py"
RUNNER_RAW_URL = "https://raw.githubusercontent.com/clifthanger/python_project/main/runner.py"

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
RUNNER_FILE = os.path.join(CURRENT_DIR, "runner.py")


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
            print("[INFO] Berhasil diperbarui.")
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
    if not os.path.exists(RUNNER_FILE):
        download_runner()

    check_update_runner()
    run_runner()
