#!/usr/bin/env python3
"""
Script สำหรับ Login และ Download ไฟล์ .xlsx จาก cenproject.rid.go.th
ใช้งานผ่าน GitHub Actions
"""
import os
import sys
import requests
from datetime import datetime

# ============ CONFIG จาก GitHub Secrets ============
USERNAME = os.environ.get("CENPROJECT_USER", "")
PASSWORD = os.environ.get("CENPROJECT_PASS", "")
BUDGET_YEAR = os.environ.get("BUDGET_YEAR", "2026")
LOGIN_URL = "https://cenproject.rid.go.th/login"
TRACK_URL = f"https://cenproject.rid.go.th/track/project?BudgetYear={BUDGET_YEAR}"
EXPORT_URL = f"https://cenproject.rid.go.th/track/export?BudgetYear={BUDGET_YEAR}"
OUTPUT_PATH = "data/cenproject_data.xlsx"
META_PATH = "data/meta.json"

def main():
    if not USERNAME or not PASSWORD:
        print("❌ ไม่พบ CENPROJECT_USER หรือ CENPROJECT_PASS ใน environment")
        sys.exit(1)

    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "th,en;q=0.9",
        "Referer": "https://cenproject.rid.go.th/",
    })

    # ---- Step 1: GET login page (เพื่อรับ CSRF token ถ้ามี) ----
    print("📡 กำลังเชื่อมต่อ cenproject.rid.go.th...")
    try:
        resp = session.get("https://cenproject.rid.go.th/", timeout=30, verify=False)
        resp.raise_for_status()
        print(f"  → GET / : {resp.status_code}")
    except Exception as e:
        print(f"❌ เชื่อมต่อไม่ได้: {e}")
        sys.exit(1)

    # ---- Step 2: POST login ----
    print("🔐 กำลัง Login...")
    login_data = {
        "username": USERNAME,
        "password": PASSWORD,
        "_token": extract_csrf(resp.text),  # Laravel CSRF
    }
    try:
        resp2 = session.post(LOGIN_URL, data=login_data, timeout=30, verify=False,
                             allow_redirects=True)
        print(f"  → POST login : {resp2.status_code} | URL: {resp2.url}")
        if "login" in resp2.url.lower() and resp2.status_code == 200:
            print("❌ Login ล้มเหลว — ตรวจสอบ username/password")
            sys.exit(1)
        print("✅ Login สำเร็จ")
    except Exception as e:
        print(f"❌ Login error: {e}")
        sys.exit(1)

    # ---- Step 3: เลือก ระบบติดตาม ----
    print("🗂️ กำลังเข้าระบบติดตาม...")
    try:
        resp3 = session.get(TRACK_URL, timeout=30, verify=False)
        print(f"  → GET track : {resp3.status_code}")
    except Exception as e:
        print(f"⚠️ เข้าระบบติดตามไม่ได้: {e}")

    # ---- Step 4: Export Excel ----
    print("📥 กำลัง Export Excel...")
    try:
        resp4 = session.get(EXPORT_URL, timeout=120, verify=False,
                            headers={"Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,*/*"})
        print(f"  → GET export : {resp4.status_code}")
        content_type = resp4.headers.get("Content-Type", "")
        print(f"  → Content-Type: {content_type}")

        # ตรวจสอบว่าได้ไฟล์ Excel จริง
        is_excel = (
            "spreadsheet" in content_type or
            "excel" in content_type or
            "octet-stream" in content_type or
            resp4.content[:4] == b'PK\x03\x04'  # ZIP header = xlsx
        )
        if not is_excel:
            print(f"❌ Response ไม่ใช่ Excel file")
            print(f"  Content preview: {resp4.content[:200]}")
            sys.exit(1)

        # ---- Step 5: บันทึกไฟล์ ----
        os.makedirs("data", exist_ok=True)
        with open(OUTPUT_PATH, "wb") as f:
            f.write(resp4.content)
        file_size = os.path.getsize(OUTPUT_PATH)
        print(f"✅ บันทึกไฟล์แล้ว: {OUTPUT_PATH} ({file_size:,} bytes)")

        # ---- Step 6: เขียน meta.json ----
        import json
        now = datetime.utcnow()
        meta = {
            "updated_at": now.isoformat() + "Z",
            "updated_th": now.strftime("%d/%m/") + str(int(now.strftime("%Y")) + 543) + " " + now.strftime("%H:%M UTC"),
            "file_size": file_size,
            "budget_year": BUDGET_YEAR,
            "filename": os.path.basename(OUTPUT_PATH),
        }
        with open(META_PATH, "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)
        print(f"✅ เขียน meta.json แล้ว")

    except Exception as e:
        print(f"❌ Export error: {e}")
        sys.exit(1)

def extract_csrf(html):
    """ดึง CSRF token จาก HTML (Laravel)"""
    import re
    m = re.search(r'<meta name="csrf-token" content="([^"]+)"', html)
    if m:
        return m.group(1)
    m = re.search(r'name="_token"[^>]*value="([^"]+)"', html)
    if m:
        return m.group(1)
    return ""

if __name__ == "__main__":
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    main()
