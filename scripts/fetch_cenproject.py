"""
Fetch CenProject Data — Playwright version
Login → Export Excel → บันทึกไฟล์ + meta.json
รันได้ทั้ง local และ GitHub Actions (headless)
"""
from playwright.sync_api import sync_playwright
import os, time, json
from datetime import datetime

USERNAME     = os.environ.get("CENPROJECT_USER", "onfarm23")
PASSWORD     = os.environ.get("CENPROJECT_PASS", "")
BUDGET_YEAR  = os.environ.get("BUDGET_YEAR", "2026")
BASE_URL     = "https://cenproject.rid.go.th"
EXPORT_URL   = f"{BASE_URL}/track/export?BudgetYear={BUDGET_YEAR}"
OUTPUT_DIR   = "data"
OUTPUT_META  = os.path.join(OUTPUT_DIR, "meta.json")
IS_CI        = os.environ.get("CI", "") == "true"   # GitHub Actions = headless

def log(msg): print(msg, flush=True)

def run():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    with sync_playwright() as p:
        log(f"🌐 Launch browser (headless={IS_CI})...")
        browser = p.chromium.launch(
            headless=IS_CI,
            slow_mo=200 if not IS_CI else 0,
            args=["--no-sandbox", "--disable-setuid-sandbox",
                  "--disable-dev-shm-usage"] if IS_CI else [],
        )
        ctx = browser.new_context(
            accept_downloads=True,
            viewport={"width": 1280, "height": 900},
        )
        page = ctx.new_page()

        # ── STEP 1: เปิดเว็บ ──
        log(f"\n📡 STEP 1: เปิด {BASE_URL}")
        page.goto(BASE_URL, wait_until="domcontentloaded", timeout=30000)
        page.wait_for_timeout(2000)
        log(f"  → URL: {page.url}")

        # ── STEP 2: Login ──
        log("\n🔐 STEP 2: Login...")
        page.get_by_placeholder("Username").fill(USERNAME)
        page.get_by_placeholder("Password").fill(PASSWORD)
        page.get_by_role("button", name="เข้าสู่ระบบ").click()
        page.wait_for_load_state("domcontentloaded", timeout=20000)
        page.wait_for_timeout(2000)
        log(f"  → URL หลัง login: {page.url}")

        if "/login" in page.url:
            log("❌ Login ล้มเหลว — ตรวจสอบ username/password")
            browser.close()
            raise SystemExit(1)
        log("  ✅ Login สำเร็จ")

        # ── STEP 3: เลือกระบบติดตาม ──
        log("\n🗂️ STEP 3: เลือกระบบติดตาม...")
        try:
            # คลิก "ระบบติดตาม" ถ้ามี modal หลัง login
            track_btn = page.locator("text=ระบบติดตาม").first
            if track_btn.is_visible(timeout=3000):
                track_btn.click()
                page.wait_for_load_state("domcontentloaded", timeout=15000)
                page.wait_for_timeout(2000)
                log(f"  → คลิกระบบติดตาม | URL: {page.url}")
        except Exception:
            log("  → ไม่มี modal ระบบติดตาม (ข้ามไป)")

        # ── STEP 4: เข้าหน้า Export ──
        log(f"\n📋 STEP 4: ไปหน้า Export ({EXPORT_URL})")
        page.goto(EXPORT_URL, wait_until="domcontentloaded", timeout=30000)
        page.wait_for_timeout(3000)
        log(f"  → URL: {page.url}")

        if IS_CI:
            page.screenshot(path=os.path.join(OUTPUT_DIR, "debug_export_page.png"))
            log("  → screenshot บันทึกแล้ว")

        # ── STEP 5: กด Export Excel ──
        log("\n📥 STEP 5: กดปุ่ม Export Excel...")
        log("  → รอไฟล์ download (อาจนาน 3-10 นาที สำหรับข้อมูลเยอะ)...")

        try:
            with page.expect_download(timeout=600000) as dl_info:  # รอ 10 นาที
                # ลองหาปุ่มหลายวิธี
                btn = None
                selectors = [
                    "button:has-text('Export Excel')",
                    "button:has-text('export excel')",
                    "[class*='export']:has-text('Excel')",
                    "button.btn-success:has-text('Excel')",
                ]
                for sel in selectors:
                    try:
                        candidate = page.locator(sel).first
                        if candidate.is_visible(timeout=2000):
                            btn = candidate
                            log(f"  → พบปุ่มด้วย: {sel}")
                            break
                    except Exception:
                        continue

                if btn is None:
                    log("  ⚠️ ไม่พบปุ่ม Export Excel ด้วย selector — ลอง screenshot")
                    if IS_CI:
                        page.screenshot(path=os.path.join(OUTPUT_DIR, "debug_no_button.png"))
                    raise Exception("ไม่พบปุ่ม Export Excel")

                btn.click()
                log("  → คลิกปุ่มแล้ว กำลังรอ download...")

            download = dl_info.value
            fname = download.suggested_filename or f"cenproject_{BUDGET_YEAR}_{int(time.time())}.xlsx"
            save_path = os.path.join(OUTPUT_DIR, "cenproject_data.xlsx")

            download.save_as(save_path)
            sz = os.path.getsize(save_path)
            log(f"\n  ✅ บันทึก: {save_path} ({sz:,} bytes)")

        except Exception as e:
            log(f"❌ Export ล้มเหลว: {e}")
            if IS_CI:
                page.screenshot(path=os.path.join(OUTPUT_DIR, "debug_error.png"))
            browser.close()
            raise SystemExit(1)

        browser.close()

    # ── เขียน meta.json ──
    now = datetime.utcnow()
    meta = {
        "updated_at": now.isoformat() + "Z",
        "updated_th": (
            f"{now.day:02d}/{now.month:02d}/{now.year+543} "
            f"{(now.hour+7)%24:02d}:{now.minute:02d} ICT"
        ),
        "file_size_bytes": os.path.getsize(os.path.join(OUTPUT_DIR, "cenproject_data.xlsx")),
        "budget_year": BUDGET_YEAR,
        "filename": "cenproject_data.xlsx",
    }
    with open(OUTPUT_META, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)

    log(f"✅ meta.json บันทึกแล้ว")
    log(f"\n🎉 สำเร็จ! อัปเดต ณ {meta['updated_th']}")

if __name__ == "__main__":
    run()
