#!/usr/bin/env python3
"""
Fetch CenProject Data v8 — FINAL
Variables ครบถ้วนจาก Network tab จริง:
  BudgetYear, ProgressResultID, Order, VendorPurchaseFilter, SelfPurchaseFilter,
  TroubleTypeFilter, UsedBudgetFilter, ProgressPlanFilter, ProjectTypeFilter,
  KPIFilter, KPISupFilter, KPISupNumber, IsExportImage
  + fields อื่นๆ ทั้งหมด (null)
"""
import os, sys, re, json
from datetime import datetime
import requests, urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

USERNAME    = os.environ.get("CENPROJECT_USER", "")
PASSWORD    = os.environ.get("CENPROJECT_PASS", "")
BUDGET_YEAR = int(os.environ.get("BUDGET_YEAR", "2026"))
BASE        = "https://cenproject.rid.go.th"
GQL_URL     = f"{BASE}/track-service"
OUTPUT_DIR  = "data"
OUTPUT_XLSX = os.path.join(OUTPUT_DIR, "cenproject_data.xlsx")
OUTPUT_META = os.path.join(OUTPUT_DIR, "meta.json")
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
      "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")

# ── GraphQL query จาก View source จริง ──
GQL_QUERY = (
    "query ExportExcelProjectExport("
    "$BudgetYear: [Int!], $ProjectName: String, $ProvinceID: [Int], "
    "$DistrictID: [Int], $SubdistrictID: [Int], $BasinID: [Int], "
    "$SubbasinID: [Int], $OperationStartYear: Int, $ProductID: [Int], "
    "$ActivityID: [Int], $ProjectTypeID: [Int], "
    "$BudgetRangeStart: Float, $BudgetRangeEnd: Float, "
    "$PurchaseTypeID: [Int], $PurchaseStepID: [Int], "
    "$ProgressResultRangeStart: Float, $ProgressResultRangeEnd: Float, "
    "$UsedBudgetResultRangeStart: Float, $UsedBudgetResultRangeEnd: Float, "
    "$OfficeOrganizationID: [Int], $OrganizationID: [Int], "
    "$BudgetSourceID: [Int], $BudgetTypeID: [Int], "
    "$ProjectCode: [String], $ProjectSizeID: [Int], "
    "$TagID: [Int], $RoyalID: [Int], "
    "$PercentRangeStart: Float, $PercentRangeEnd: Float, "
    "$IsCanceled: Int, $IsMergeDraft: Int, "
    "$VendorPurchaseFilter: Boolean!, $SelfPurchaseFilter: Boolean!, "
    "$TroubleTypeFilter: Boolean!, $UsedBudgetFilter: Boolean!, "
    "$ProgressPlanFilter: Boolean!, $ProjectTypeFilter: Boolean!, "
    "$KPIFilter: Boolean!, $KPISupFilter: Boolean!, "
    "$KPISupNumber: Int, $PictureFilter: Boolean, "
    "$ProgressResultID: [Int], $ProgressResultInProgressID: [Int], "
    "$BudgetDimensionID: [Int], $SupervisorName: String, "
    "$IsExportImage: Boolean, $Order: Int, $Offset: Int, $Limit: Int"
    ") {\n"
    "  ExportExcelProjectExport(\n"
    "    BudgetYear: $BudgetYear\n"
    "    ProjectName: $ProjectName\n"
    "    ProvinceID: $ProvinceID\n"
    "    DistrictID: $DistrictID\n"
    "    SubdistrictID: $SubdistrictID\n"
    "    BasinID: $BasinID\n"
    "    SubbasinID: $SubbasinID\n"
    "    OperationStartYear: $OperationStartYear\n"
    "    ProductID: $ProductID\n"
    "    ActivityID: $ActivityID\n"
    "    ProjectTypeID: $ProjectTypeID\n"
    "    BudgetRangeStart: $BudgetRangeStart\n"
    "    BudgetRangeEnd: $BudgetRangeEnd\n"
    "    PurchaseTypeID: $PurchaseTypeID\n"
    "    PurchaseStepID: $PurchaseStepID\n"
    "    ProgressResultRangeStart: $ProgressResultRangeStart\n"
    "    ProgressResultRangeEnd: $ProgressResultRangeEnd\n"
    "    UsedBudgetResultRangeStart: $UsedBudgetResultRangeStart\n"
    "    UsedBudgetResultRangeEnd: $UsedBudgetResultRangeEnd\n"
    "    OfficeOrganizationID: $OfficeOrganizationID\n"
    "    OrganizationID: $OrganizationID\n"
    "    BudgetSourceID: $BudgetSourceID\n"
    "    BudgetTypeID: $BudgetTypeID\n"
    "    ProjectCode: $ProjectCode\n"
    "    ProjectSizeID: $ProjectSizeID\n"
    "    TagID: $TagID\n"
    "    RoyalID: $RoyalID\n"
    "    PercentRangeStart: $PercentRangeStart\n"
    "    PercentRangeEnd: $PercentRangeEnd\n"
    "    IsCanceled: $IsCanceled\n"
    "    IsMergeDraft: $IsMergeDraft\n"
    "    VendorPurchaseFilter: $VendorPurchaseFilter\n"
    "    SelfPurchaseFilter: $SelfPurchaseFilter\n"
    "    TroubleTypeFilter: $TroubleTypeFilter\n"
    "    UsedBudgetFilter: $UsedBudgetFilter\n"
    "    ProgressPlanFilter: $ProgressPlanFilter\n"
    "    ProjectTypeFilter: $ProjectTypeFilter\n"
    "    KPIFilter: $KPIFilter\n"
    "    KPISupFilter: $KPISupFilter\n"
    "    KPISupNumber: $KPISupNumber\n"
    "    PictureFilter: $PictureFilter\n"
    "    ProgressResultID: $ProgressResultID\n"
    "    ProgressResultInProgressID: $ProgressResultInProgressID\n"
    "    BudgetDimensionID: $BudgetDimensionID\n"
    "    SupervisorName: $SupervisorName\n"
    "    IsExportImage: $IsExportImage\n"
    "    Order: $Order\n"
    "    Offset: $Offset\n"
    "    Limit: $Limit\n"
    "  ) {\n"
    "    FileName\n"
    "    Link\n"
    "    __typename\n"
    "  }\n"
    "}"
)

# Variables จาก Payload จริง (View source)
GQL_VARIABLES = {
    "BudgetYear": None,          # จะใส่ค่าจริงตอน runtime
    "ProgressResultID": None,
    "Order": None,
    "VendorPurchaseFilter": True,
    "SelfPurchaseFilter": True,
    "TroubleTypeFilter": True,
    "UsedBudgetFilter": True,
    "ProgressPlanFilter": True,
    "ProjectTypeFilter": True,
    "KPIFilter": True,
    "KPISupFilter": True,
    "KPISupNumber": 5,
    "IsExportImage": False,
    # fields ที่เหลือ null ทั้งหมด
    "ProjectName": None,
    "ProvinceID": None,
    "DistrictID": None,
    "SubdistrictID": None,
    "BasinID": None,
    "SubbasinID": None,
    "OperationStartYear": None,
    "ProductID": None,
    "ActivityID": None,
    "ProjectTypeID": None,
    "BudgetRangeStart": None,
    "BudgetRangeEnd": None,
    "PurchaseTypeID": None,
    "PurchaseStepID": None,
    "ProgressResultRangeStart": None,
    "ProgressResultRangeEnd": None,
    "UsedBudgetResultRangeStart": None,
    "UsedBudgetResultRangeEnd": None,
    "OfficeOrganizationID": None,
    "OrganizationID": None,
    "BudgetSourceID": None,
    "BudgetTypeID": None,
    "ProjectCode": None,
    "ProjectSizeID": None,
    "TagID": None,
    "RoyalID": None,
    "PercentRangeStart": None,
    "PercentRangeEnd": None,
    "IsCanceled": None,
    "IsMergeDraft": None,
    "PictureFilter": None,
    "ProgressResultInProgressID": None,
    "BudgetDimensionID": None,
    "SupervisorName": None,
    "Offset": None,
    "Limit": None,
}

def log(msg): print(msg, flush=True)

def extract_token(html):
    for p in [
        r'<meta[^>]+name=["\']csrf-token["\'][^>]+content=["\']([^"\']+)["\']',
        r'name=["\']_token["\'][^>]*value=["\']([^"\']+)["\']',
        r'"_token"\s*:\s*"([^"]+)"',
    ]:
        m = re.search(p, html, re.IGNORECASE)
        if m: return m.group(1)
    return ""

def main():
    if not USERNAME or not PASSWORD:
        log("❌ ไม่พบ CENPROJECT_USER / CENPROJECT_PASS")
        sys.exit(1)

    log(f"👤 User: {USERNAME} | BudgetYear: {BUDGET_YEAR}")
    log("─" * 60)

    s = requests.Session()
    s.verify = False
    s.headers.update({"User-Agent": UA, "Accept-Language": "th-TH,th;q=0.9"})

    # ── STEP 1: Login ──
    log("📡 STEP 1: GET หน้าหลัก...")
    r1 = s.get(f"{BASE}/", timeout=30)
    log(f"  → {r1.status_code} | {r1.url}")
    token = extract_token(r1.text)
    if not token:
        r1b = s.get(f"{BASE}/login", timeout=30)
        token = extract_token(r1b.text)
    log(f"  → CSRF: {'พบ' if token else 'ไม่พบ'}")

    log("\n🔐 STEP 2: POST Login...")
    r2 = s.post(f"{BASE}/login",
                data={"username": USERNAME, "password": PASSWORD, "_token": token},
                timeout=30, allow_redirects=True)
    log(f"  → {r2.status_code} | {r2.url}")
    if "/login" in r2.url and r2.status_code == 200:
        log("❌ Login ล้มเหลว")
        sys.exit(1)
    log("✅ Login สำเร็จ")

    # ── STEP 2: เข้าหน้า export ──
    log("\n🗂️ STEP 3: เข้าหน้า export...")
    r3 = s.get(f"{BASE}/track/export?BudgetYear={BUDGET_YEAR}", timeout=30)
    log(f"  → {r3.status_code} | {r3.url}")

    # ── STEP 3: GraphQL ──
    log(f"\n🔗 STEP 4: POST GraphQL (v8 FINAL)...")
    variables = dict(GQL_VARIABLES)
    variables["BudgetYear"] = [BUDGET_YEAR]

    gql_payload = {
        "operationName": "ExportExcelProjectExport",
        "variables": variables,
        "query": GQL_QUERY,
    }

    gql_headers = {
        "Content-Type": "application/json",
        "Accept": "application/json, */*",
        "Origin": BASE,
        "Referer": f"{BASE}/track/export?BudgetYear={BUDGET_YEAR}",
        "User-Agent": UA,
    }

    r4 = s.post(GQL_URL, json=gql_payload, headers=gql_headers, timeout=180)
    log(f"  → Status: {r4.status_code}")
    log(f"  → Response (500 chars): {r4.text[:500]}")

    if r4.status_code != 200:
        log(f"❌ GraphQL ตอบ {r4.status_code}")
        sys.exit(1)

    try:
        gql_resp = r4.json()
    except Exception:
        log("❌ Response ไม่ใช่ JSON")
        sys.exit(1)

    errors = gql_resp.get("errors", [])
    if errors:
        log("❌ GraphQL errors:")
        for e in errors:
            log(f"  → {e.get('message','')}")
        sys.exit(1)

    export_data = gql_resp.get("data", {}).get("ExportExcelProjectExport", {})
    if not export_data:
        log(f"❌ ไม่พบ data.ExportExcelProjectExport")
        log(f"  → Full: {json.dumps(gql_resp, ensure_ascii=False)[:800]}")
        sys.exit(1)

    file_link = export_data.get("Link", "")
    file_name = export_data.get("FileName", "cenproject_data.xlsx")
    log(f"\n  ✅ FileName: {file_name}")
    log(f"  ✅ Link: {file_link}")

    if not file_link:
        log("❌ Link ว่างเปล่า")
        sys.exit(1)

    # ── STEP 4: Download ──
    download_url = file_link if file_link.startswith("http") else BASE + file_link
    log(f"\n📥 STEP 5: Download {download_url}")

    r5 = s.get(download_url, timeout=180,
               headers={"Accept": "*/*", "Referer": f"{BASE}/track/export",
                        "User-Agent": UA})
    log(f"  → Status: {r5.status_code}")
    log(f"  → Content-Type: {r5.headers.get('Content-Type','?')}")
    log(f"  → Size: {len(r5.content):,} bytes | Magic: {r5.content[:4].hex()}")

    is_xlsx = r5.content[:4] == b'PK\x03\x04'
    is_xls  = r5.content[:8] == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'
    if not (is_xlsx or is_xls):
        log("❌ ไม่ใช่ไฟล์ Excel")
        log(f"  → Preview: {r5.content[:300]}")
        sys.exit(1)

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    with open(OUTPUT_XLSX, "wb") as f: f.write(r5.content)
    sz = os.path.getsize(OUTPUT_XLSX)
    log(f"\n✅ บันทึก: {OUTPUT_XLSX} ({sz:,} bytes)")

    now = datetime.utcnow()
    meta = {
        "updated_at": now.isoformat() + "Z",
        "updated_th": (f"{now.day:02d}/{now.month:02d}/{now.year+543} "
                       f"{(now.hour+7)%24:02d}:{now.minute:02d} ICT"),
        "file_size_bytes": sz,
        "budget_year": BUDGET_YEAR,
        "filename": file_name,
    }
    with open(OUTPUT_META, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)
    log(f"✅ meta.json บันทึกแล้ว")
    log(f"\n🎉 สำเร็จ! ข้อมูล ณ {meta['updated_th']}")

if __name__ == "__main__":
    main()
