"""
GitHub Repository Uploader for vsinghidelhi/gst-reconciliation
Pushes all project code, documentation, and configuration files to GitHub repository.
"""

import os
import sys
import base64
import requests
from pathlib import Path
from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent

# Check for .env in current folder or cch_ifirm_sync
if (BASE_DIR / ".env").exists():
    load_dotenv(BASE_DIR / ".env")
elif Path(r"C:\Users\lenovo\cch_ifirm_sync\.env").exists():
    load_dotenv(r"C:\Users\lenovo\cch_ifirm_sync\.env")

# Token is loaded strictly from environment or .env file (never committed to GitHub)
GITHUB_TOKEN = os.getenv("GITHUB_TOKEN")
if not GITHUB_TOKEN:
    # Try reading from cch_ifirm_sync .env directly if not loaded
    cch_env = Path(r"C:\Users\lenovo\cch_ifirm_sync\.env")
    if cch_env.exists():
        load_dotenv(cch_env)
        GITHUB_TOKEN = os.getenv("GITHUB_TOKEN")

REPO_OWNER = "vsinghidelhi"
REPO_NAME = "gst-reconciliation"
BASE_URL = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/contents"

HEADERS = {
    "Authorization": f"token {GITHUB_TOKEN}",
    "Accept": "application/vnd.github.v3+json"
}

FILES_TO_UPLOAD = [
    "reconciler.py",
    "run_reconciliation.py",
    "Code.gs",
    "Index.html",
    "deploy_to_gas.py",
    "requirements.txt",
    ".gitignore",
    "README.md",
    "github_uploader.py"
]

def upload_file(filename):
    file_path = BASE_DIR / filename
    if not file_path.exists():
        print(f"[!] File not found: {filename}")
        return False

    with open(file_path, "rb") as f:
        content_bytes = f.read()

    content_b64 = base64.b64encode(content_bytes).decode("utf-8")
    url = f"{BASE_URL}/{filename}"

    # Check if exists to retrieve SHA
    r = requests.get(url, headers=HEADERS)
    payload = {
        "message": f"Add / Update {filename}",
        "content": content_b64
    }
    if r.status_code == 200:
        payload["sha"] = r.json().get("sha")

    put_res = requests.put(url, headers=HEADERS, json=payload)
    if put_res.status_code in [200, 201]:
        print(f"[+] Successfully uploaded: {filename}")
        return True
    else:
        print(f"[!] Failed to upload {filename}: {put_res.status_code} - {put_res.text}")
        return False

def main():
    print(f"[*] Connecting to GitHub repository: {REPO_OWNER}/{REPO_NAME} ...")
    chk = requests.get(f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}", headers=HEADERS)
    if chk.status_code != 200:
        print(f"[!] Repository check failed: {chk.status_code} - {chk.text}")
        sys.exit(1)

    print(f"[+] Repository found: {chk.json().get('html_url')}")
    print(f"[*] Uploading project files...")

    success_count = 0
    for filename in FILES_TO_UPLOAD:
        if upload_file(filename):
            success_count += 1

    print(f"\n[OK] Upload complete: {success_count}/{len(FILES_TO_UPLOAD)} files successfully committed to GitHub.")
    print(f"     Repository URL: https://github.com/{REPO_OWNER}/{REPO_NAME}")

if __name__ == "__main__":
    main()
