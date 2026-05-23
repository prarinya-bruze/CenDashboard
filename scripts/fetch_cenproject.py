#!/usr/bin/env python3
"""
Fetch CenProject Data — GitHub Actions Script
Login → เลือกระบบติดตาม → Export Excel
"""
import os, sys, re, json
from datetime import datetime
import requests
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ============ CONFIG ============
USERNAME    = os.environ.get("CENPROJECT_USER", "")
PASSWORD    = os.environ.get("CENPROJECT_PASS", "")
BUDGET_YEAR = os.environ.get("BUDGET_YEAR", "2026")
BASE        = "https://cenproject.rid.go.th"
OUTPUT_DIR  = "data"
OUTPUT_XLSX = os.path.join(OUTPUT_DIR, "cenproject_data.xlsx")
OUTPUT_META = os.path.join(OUTPUT_DIR, "meta.json")

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "th-TH,th;q=0.9,en;q=0.8",
}

def log(msg): print(msg, flush=True)

def extract_token(html):
    """ดึง CSRF / _token จาก HTML หลายรูปแบบ"""
    patterns = [
        r'<meta[^>]+name=["\']csrf-token["\'][^>]+content=["\']([^"\']+)["\']',
        r'<meta[^>]+content=["\']([^"\']+)["\'][^>]+name=["\']csrf-token["\']',
        r'name=["\']_token["\'][^>]*value=["\']([^"\']+)["\']',
        r'value=["\']([^"\']+)["\'][^>]*name=["\']_token["\']',
        r'"_token"\s*:\s*"([^"]+)"',
        r"'_token'\s*:\s*'([^']+)'",
    ]
    for p in patterns:
        m = re.search(p, html, re.IGNORECASE)
        if m:
            log(f"  → CSRF token พบด้วย pattern: {p[:40]}...")
            return m.group(1)
    log("  ⚠️ ไม่พบ CSRF token")
    return ""

def main():
    if not USERNAME or not PASSWORD:
        log("❌ กรุณาตั้งค่า CENPROJECT_USER และ CENPROJECT_PASS ใน GitHub Secrets")
        sys.exit(1)

    log(f"👤 Username: {USERNAME}")
    log(f"🌐 Base URL: {BASE}")
    log(f"📅 Budget Year: {BUDGET_YEAR}")
    log("─" * 50)

    s = requests.Session()
    s.headers.update(HEADERS)
    s.verify = False

    # ──────────────────────────────────────────
    # STEP 1: GET หน้า Login — รับ CSRF token
    # ──────────────────────────────────────────
    log("📡 STEP 1: เข้าหน้า Login...")
    try:
        r1 = s.get(f"{BASE}/", timeout=30)
        log(f"  → Status: {r1.status_code} | URL: {r1.url}")
        log(f"  → Cookies: {dict(s.cookies)}")
        token = extract_token(r1.text)
        log(f"  → Token: {token[:20]}..." if token else "  → Token: (ไม่พบ)")

        # ถ้าหน้าแรก redirect ไปที่ login ให้ดึง token จาก login page แทน
        if "login" in r1.url.lower() or not token:
            r1b = s.get(f"{BASE}/login", timeout=30)
            log(f"  → GET /login: {r1b.status_code}")
            token = extract_token(r1b.text) or token

    except Exception as e:
        log(f"❌ เชื่อมต่อ {BASE} ไม่ได้: {e}")
        sys.exit(1)

    # ──────────────────────────────────────────
    # STEP 2: POST Login
    # ──────────────────────────────────────────
    log("\n🔐 STEP 2: Login...")
    login_payload = {
        "username": USERNAME,
        "password": PASSWORD,
    }
    if token:
        login_payload["_token"] = token

    try:
        r2 = s.post(
            f"{BASE}/login",
            data=login_payload,
            timeout=30,
            allow_redirects=True,
        )
        log(f"  → Status: {r2.status_code} | URL: {r2.url}")
        log(f"  → Cookies after login: {list(s.cookies.keys())}")

        # ตรวจสอบว่า Login สำเร็จ
        if "/login" in r2.url and r2.status_code == 200:
            # ลองหา error message ใน response
            err_patterns = [
                r'class=["\'][^"\']*alert[^"\']*["\'][^>]*>(.*?)</div>',
                r'class=["\'][^"\']*error[^"\']*["\'][^>]*>(.*?)</\w+>',
                r'<p[^>]*style=["\'][^"\']*color\s*:\s*red[^"\']*["\'][^>]*>(.*?)</p>',
            ]
            err_msg = ""
            for ep in err_patterns:
                m = re.search(ep, r2.text, re.IGNORECASE | re.DOTALL)
                if m:
                    err_msg = re.sub(r'<[^>]+>', '', m.group(1)).strip()
                    break
            log(f"❌ Login ล้มเหลว | Error: {err_msg or '(ไม่มีข้อความ error)'}")
            log(f"  → Response preview: {r2.text[:500]}")
            sys.exit(1)

        log("✅ Login สำเร็จ")

    except Exception as e:
        log(f"❌ Login error: {e}")
        sys.exit(1)

    # ──────────────────────────────────────────
    # STEP 3: เข้าระบบติดตาม
    # ──────────────────────────────────────────
    log("\n🗂️ STEP 3: เข้าระบบติดตาม...")
    track_url = f"{BASE}/track/project?BudgetYear={BUDGET_YEAR}"
    try:
        r3 = s.get(track_url, timeout=30)
        log(f"  → Status: {r3.status_code} | URL: {r3.url}")
        if "/login" in r3.url:
            log("❌ Session หมดอายุ — ต้อง login ใหม่")
            sys.exit(1)
    except Exception as e:
        log(f"⚠️ เข้าระบบติดตามไม่ได้: {e} (ข้ามไป)")

    # ──────────────────────────────────────────
    # STEP 4: Export Excel
    # ──────────────────────────────────────────
    log("\n📥 STEP 4: Export Excel...")
    export_url = f"{BASE}/track/export?BudgetYear={BUDGET_YEAR}"
    try:
        r4 = s.get(
            export_url,
            timeout=120,
            headers={
                **HEADERS,
                "Accept": (
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,"
                    "application/vnd.ms-excel,application/octet-stream,*/*"
                ),
                "Referer": track_url,
            },
            stream=True,
        )
        log(f"  → Status: {r4.status_code}")
        log(f"  → Content-Type: {r4.headers.get('Content-Type','?')}")
        log(f"  → Content-Disposition: {r4.headers.get('Content-Disposition','?')}")

        content = r4.content
        log(f"  → Content length: {len(content):,} bytes")
        log(f"  → First 8 bytes (hex): {content[:8].hex()}")

        # ตรวจสอบ magic bytes ของ xlsx (ZIP = PK\x03\x04)
        is_xlsx = content[:4] == b'PK\x03\x04'
        is_xls  = content[:8] == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'  # OLE2
        ct = r4.headers.get("Content-Type", "")
        is_ct_ok = any(x in ct for x in ["spreadsheet", "excel", "octet-stream"])

        if not (is_xlsx or is_xls or is_ct_ok):
            log("❌ Response ไม่ใช่ Excel file")
            log(f"  → HTML preview: {content[:300].decode('utf-8','ignore')}")
            sys.exit(1)

        # ──────────────────────────────────────────
        # STEP 5: บันทึกไฟล์
        # ──────────────────────────────────────────
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        with open(OUTPUT_XLSX, "wb") as f:
            f.write(content)
        file_size = os.path.getsize(OUTPUT_XLSX)
        log(f"\n✅ บันทึก: {OUTPUT_XLSX} ({file_size:,} bytes)")

        # เขียน meta.json
        now = datetime.utcnow()
        th_year = now.year + 543
        meta = {
            "updated_at": now.isoformat() + "Z",
            "updated_th": f"{now.day:02d}/{now.month:02d}/{th_year} {now.hour+7:02d}:{now.minute:02d} ICT",
            "file_size_bytes": file_size,
            "budget_year": BUDGET_YEAR,
            "filename": os.path.basename(OUTPUT_XLSX),
            "source": BASE,
        }
        with open(OUTPUT_META, "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)
        log(f"✅ บันทึก: {OUTPUT_META}")
        log(f"\n🎉 เสร็จสิ้น! ข้อมูล ณ {meta['updated_th']}")

    except Exception as e:
        log(f"❌ Export error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()
