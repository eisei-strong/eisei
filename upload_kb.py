#!/usr/bin/env python3
"""KBデータをGoogle Driveにアップロード"""

import json
import urllib.request
import urllib.parse

CSV_FILE = "/Users/kodai/sales-dashboard/hoshi_feedback_kb.csv"
CLASP_CREDS = "/Users/kodai/.clasprc.json"

def get_access_token():
    with open(CLASP_CREDS) as f:
        creds = json.load(f)
    token_data = creds["tokens"]["default"]
    data = urllib.parse.urlencode({
        "client_id": token_data["client_id"],
        "client_secret": token_data["client_secret"],
        "refresh_token": token_data["refresh_token"],
        "grant_type": "refresh_token"
    }).encode()
    req = urllib.request.Request("https://oauth2.googleapis.com/token", data=data)
    resp = urllib.request.urlopen(req)
    return json.loads(resp.read())["access_token"]

def upload_csv_to_drive(token):
    """CSVをGoogle Driveにアップロード（Spreadsheet形式で変換）"""
    # Multipart upload
    with open(CSV_FILE, "rb") as f:
        csv_data = f.read()

    metadata = json.dumps({
        "name": "hoshi_feedback_kb",
        "mimeType": "application/vnd.google-apps.spreadsheet"
    }).encode()

    boundary = b"----boundary123456"
    body = b""
    body += b"--" + boundary + b"\r\n"
    body += b"Content-Type: application/json; charset=UTF-8\r\n\r\n"
    body += metadata + b"\r\n"
    body += b"--" + boundary + b"\r\n"
    body += b"Content-Type: text/csv\r\n\r\n"
    body += csv_data + b"\r\n"
    body += b"--" + boundary + b"--\r\n"

    url = "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&convert=true"
    req = urllib.request.Request(url, data=body, headers={
        "Authorization": f"Bearer {token}",
        "Content-Type": f"multipart/related; boundary={boundary.decode()}"
    })

    resp = urllib.request.urlopen(req, timeout=120)
    result = json.loads(resp.read())
    return result

def main():
    print("=== KB CSV → Google Drive アップロード ===")
    print("Authenticating...")
    token = get_access_token()

    print(f"Uploading {CSV_FILE}...")
    result = upload_csv_to_drive(token)
    file_id = result["id"]
    print(f"OK! File ID: {file_id}")
    print(f"URL: https://docs.google.com/spreadsheets/d/{file_id}")
    print(f"\nこのFile IDをGASで使います: {file_id}")

if __name__ == "__main__":
    main()
