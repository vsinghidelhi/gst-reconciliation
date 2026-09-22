"""
Push Code.gs, sample_data.gs, Index.html, and appsscript.json to Google Apps Script
and create a live Web App deployment.
"""

import os
import sys
import json
import urllib.request
import urllib.parse
from pathlib import Path

sys.path.append(r'C:\Users\lenovo\VSA_Compliance_Dashboard')
from push_to_gas import get_valid_access_token

SCRIPT_ID = "1hZGoAw_AX-mQtGgk3baPHUlU1X3o_AHqOEqUNBOEIujk0BX662fVWjtu"
BASE_DIR = Path(__file__).resolve().parent

def main():
    print(f"[*] Getting valid Google OAuth access token...")
    token = get_valid_access_token()
    print(f"[+] Access token active!")

    # 1. Read files
    code_path = BASE_DIR / "Code.gs"
    sample_path = BASE_DIR / "sample_data.gs"
    index_path = BASE_DIR / "Index.html"

    with open(code_path, "r", encoding="utf-8") as f:
        code_content = f.read()

    with open(sample_path, "r", encoding="utf-8") as f:
        sample_content = f.read()

    with open(index_path, "r", encoding="utf-8") as f:
        index_content = f.read()

    manifest = {
        "timeZone": "Asia/Kolkata",
        "dependencies": {},
        "exceptionLogging": "STACKDRIVER",
        "runtimeVersion": "V8",
        "webapp": {
            "executeAs": "USER_DEPLOYING",
            "access": "ANYONE_ANONYMOUS"
        }
    }

    files = [
        {
            "name": "appsscript",
            "type": "JSON",
            "source": json.dumps(manifest, indent=2)
        },
        {
            "name": "Code",
            "type": "SERVER_JS",
            "source": code_content
        },
        {
            "name": "sample_data",
            "type": "SERVER_JS",
            "source": sample_content
        },
        {
            "name": "Index",
            "type": "HTML",
            "source": index_content
        }
    ]

    # 2. Upload to Google Apps Script
    print(f"[*] Uploading files to Apps Script Project: {SCRIPT_ID} ...")
    content_url = f"https://script.googleapis.com/v1/projects/{SCRIPT_ID}/content"
    put_req = urllib.request.Request(
        content_url,
        data=json.dumps({"files": files}).encode("utf-8"),
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        method="PUT"
    )

    try:
        with urllib.request.urlopen(put_req) as resp:
            print("[+] Successfully uploaded all files to Google Apps Script!")
    except urllib.error.HTTPError as e:
        print(f"[!] Upload failed ({e.code}): {e.read().decode('utf-8')}")
        sys.exit(1)

    # 3. Create Version
    print(f"[*] Creating project version...")
    ver_url = f"https://script.googleapis.com/v1/projects/{SCRIPT_ID}/versions"
    ver_req = urllib.request.Request(
        ver_url,
        data=json.dumps({"description": "v1.0 - Initial Production Release of VSA GST Reconciliation Portal"}).encode("utf-8"),
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        method="POST"
    )

    try:
        with urllib.request.urlopen(ver_req) as resp:
            ver_res = json.loads(resp.read().decode("utf-8"))
            version_number = ver_res.get("versionNumber", 1)
            print(f"[+] Version created: v{version_number}")
    except urllib.error.HTTPError as e:
        print(f"[!] Version creation failed: {e.read().decode('utf-8')}")
        version_number = 1

    # 4. Create Web App Deployment
    print(f"[*] Creating Web App deployment...")
    dep_url = f"https://script.googleapis.com/v1/projects/{SCRIPT_ID}/deployments"
    dep_payload = {
        "versionNumber": version_number,
        "manifestFileName": "appsscript",
        "description": "v1.0 Production Web App"
    }

    dep_req = urllib.request.Request(
        dep_url,
        data=json.dumps(dep_payload).encode("utf-8"),
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        method="POST"
    )

    try:
        with urllib.request.urlopen(dep_req) as resp:
            dep_res = json.loads(resp.read().decode("utf-8"))
            deployment_id = dep_res.get("deploymentId")
            print(f"[+] Deployment successful!")
            print(f"    Deployment ID: {deployment_id}")
            for entry in dep_res.get("entryPoints", []):
                if entry.get("entryPointType") == "WEB_APP":
                    web_url = entry.get("webApp", {}).get("url")
                    print(f"\n[OK] LIVE WEB APP URL:\n     {web_url}\n")
    except urllib.error.HTTPError as e:
        print(f"[!] Deployment failed ({e.code}): {e.read().decode('utf-8')}")

if __name__ == "__main__":
    main()
