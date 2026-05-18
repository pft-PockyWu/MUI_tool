#!/usr/bin/env python3
"""
MUI Tool — Upload to Google Drive
Build 完後自動上傳 exe 並更新 latest_version.txt

首次執行時會開啟瀏覽器做 Google 授權（只需做一次），
之後 token 會存在 token.json，不需再手動授權。

使用方式：
    python upload_to_drive.py
"""

import io
import re
import sys
from pathlib import Path

# ── Google Drive 設定 ─────────────────────────────────────────────────────────
FOLDER_ID        = "1-DumwCUoqOV-zgMRhUtI4QRMtqGn3qDq"   # 公開下載資料夾
VERSION_FILE_ID  = "11dRxedUqKwtD26P4EcNK17qmICefJ4-w"    # latest_version.txt
SCOPES           = ["https://www.googleapis.com/auth/drive"]
HERE             = Path(__file__).parent
CREDENTIALS_PATH = HERE / "credentials.json"
TOKEN_PATH       = HERE / "token.json"
DIST_DIR         = HERE / "dist"


# ── 套件自動安裝 ──────────────────────────────────────────────────────────────
def _ensure_packages():
    try:
        from googleapiclient.discovery import build   # noqa
        from google_auth_oauthlib.flow import InstalledAppFlow  # noqa
    except ImportError:
        print("⚠️  正在安裝 Google API 套件，請稍候...")
        import subprocess
        subprocess.check_call([sys.executable, "-m", "pip", "install",
                               "google-api-python-client",
                               "google-auth-oauthlib",
                               "google-auth-httplib2"])
        print("✅ 安裝完成，重新啟動腳本...\n")
        import os
        os.execv(sys.executable, [sys.executable] + sys.argv)


# ── Google Drive 認證 ─────────────────────────────────────────────────────────
def _get_service():
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build

    creds = None
    if TOKEN_PATH.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not CREDENTIALS_PATH.exists():
                print("❌ 找不到 credentials.json！")
                print()
                print("   請依照以下步驟取得：")
                print("   1. 前往 https://console.cloud.google.com/")
                print("   2. 建立新專案（或選既有專案）")
                print("   3. 啟用 Google Drive API")
                print("   4. 建立 OAuth 憑證 → 桌面應用程式")
                print("   5. 下載 JSON，重新命名為 credentials.json")
                print(f"   6. 放到 {HERE}")
                sys.exit(1)
            flow = InstalledAppFlow.from_client_secrets_file(
                str(CREDENTIALS_PATH), SCOPES)
            creds = flow.run_local_server(port=0)
        TOKEN_PATH.write_text(creds.to_json(), encoding="utf-8")

    return build("drive", "v3", credentials=creds)


# ── 找最新 exe ────────────────────────────────────────────────────────────────
def _find_latest_exe() -> Path:
    exes = sorted(
        DIST_DIR.glob("MUI_tool_*.exe"),
        key=lambda p: p.stat().st_mtime,
        reverse=True
    )
    if not exes:
        print(f"❌ 在 {DIST_DIR} 找不到 MUI_tool_*.exe")
        print("   請先執行 build_windows.bat 完成打包")
        sys.exit(1)
    return exes[0]


# ── 解析版號 ──────────────────────────────────────────────────────────────────
def _parse_version(exe_path: Path) -> str:
    """MUI_tool_1_3_260508_1030.exe  →  '1.3.260508_1030'"""
    m = re.search(r'MUI_tool_(\d+)_(\d+)_(\d{6}_\d{4})\.exe', exe_path.name)
    if not m:
        print(f"❌ 無法從檔名解析版號：{exe_path.name}")
        sys.exit(1)
    return f"{m.group(1)}.{m.group(2)}.{m.group(3)}"


# ── 刪除舊版 exe ──────────────────────────────────────────────────────────────
def _delete_old_exes(service, new_name: str):
    results = service.files().list(
        q=(f"'{FOLDER_ID}' in parents"
           " and name contains 'MUI_tool_'"
           " and name contains '.exe'"
           " and trashed = false"),
        fields="files(id, name)"
    ).execute()
    for f in results.get("files", []):
        if f["name"] == new_name:
            continue   # 同名就直接覆蓋，不必先刪
        service.files().delete(fileId=f["id"]).execute()
        print(f"   🗑️  已刪除舊版：{f['name']}")


# ── 上傳 exe ──────────────────────────────────────────────────────────────────
def _upload_exe(service, exe_path: Path):
    from googleapiclient.http import MediaFileUpload

    size_mb = exe_path.stat().st_size / 1024 / 1024
    print(f"   ⬆️  上傳中：{exe_path.name}  ({size_mb:.1f} MB)")

    media = MediaFileUpload(
        str(exe_path),
        mimetype="application/octet-stream",
        resumable=True,
        chunksize=4 * 1024 * 1024   # 4 MB chunks
    )
    file_meta = {"name": exe_path.name, "parents": [FOLDER_ID]}
    request = service.files().create(
        body=file_meta, media_body=media, fields="id,name")

    response = None
    while response is None:
        status, response = request.next_chunk()
        if status:
            print(f"      {int(status.progress() * 100):3d}%", end="\r")
    print(f"   ✅ 上傳完成：{response['name']}      ")


# ── 更新 latest_version.txt ───────────────────────────────────────────────────
def _update_version_file(service, version: str):
    media = io.BytesIO(version.encode("utf-8"))
    from googleapiclient.http import MediaIoBaseUpload
    service.files().update(
        fileId=VERSION_FILE_ID,
        media_body=MediaIoBaseUpload(media, mimetype="text/plain")
    ).execute()
    print(f"   ✅ latest_version.txt 已更新為：{version}")


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    _ensure_packages()

    exe_path = _find_latest_exe()
    version  = _parse_version(exe_path)

    print(f"\n📦 準備發佈：{exe_path.name}")
    print(f"   版號：v{version}\n")

    service = _get_service()
    _delete_old_exes(service, exe_path.name)
    _upload_exe(service, exe_path)
    _update_version_file(service, version)

    print(f"\n🎉 完成！v{version} 已發佈至 Google Drive。")
    print(f"   下載連結：https://drive.google.com/drive/folders/{FOLDER_ID}\n")


if __name__ == "__main__":
    main()
