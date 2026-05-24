#!/usr/bin/env python3
"""
Fetch CenProject Data v9 — แก้ type [Int] → [Int!] ทุกตัว
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

# query จาก View source จริง — ใช้ type ตรงตาม server ([Int!] ทุกตัว)
GQL_QUERY = """query ExportExcelProjectExport(
  $BudgetYear: [Int!],
  $ProjectName: String,
  $ProvinceID: [Int!],
  $DistrictID: [Int!],
  $SubdistrictID: [Int!],
  $BasinID: [Int!],
  $SubbasinID: [Int!],
  $OperationStartYear: Int,
  $ProductID: [Int!],
  $ActivityID: [Int!],
  $ProjectTypeID: [Int!],
  $BudgetRangeStart: Float,
  $BudgetRangeEnd: Float,
  $PurchaseTypeID: [Int!],
  $PurchaseStepID: [Int!],
  $ProgressResultRangeStart: Float,
  $ProgressResultRangeEnd: Float,
  $UsedBudgetResultRangeStart: Float,
  $UsedBudgetResultRangeEnd: Float,
  $OfficeOrganizationID: [Int!],
  $OrganizationID: [Int!],
  $BudgetSourceID: [Int!],
  $BudgetTypeID: [Int!],
  $ProjectCode: [String!],
  $ProjectSizeID: [Int!],
  $TagID: [Int!],
  $RoyalID: [Int!],
  $PercentRangeStart: Float,
  $PercentRangeEnd: Float,
  $IsCanceled: Int,
  $IsMergeDraft: Int,
  $VendorPurchaseFilter: Boolean!,
  $SelfPurchaseFilter: Boolean!,
  $TroubleTypeFilter: Boolean!,
  $UsedBudgetFilter: Boolean!,
  $ProgressPlanFilter: Boolean!,
  $ProjectTypeFilter: Boolean!,
  $KPIFilter: Boolean!,
  $KPISupFilter: Boolean!,
  $KPISupNumber: Int,
  $PictureFilter: Boolean,
  $ProgressResultID: [Int!],
  $ProgressResultInProgressID: [Int!],
  $BudgetDimensionID: [Int!],
  $SupervisorName: String,
  $IsExportImage: Boolean,
  $Order: Int,
  $Offset: Int,
  $Limit: Int
) {
  ExportExcelProjectExport(
    BudgetYear: $BudgetYear
    ProjectName: $ProjectName
    ProvinceID: $ProvinceID
    DistrictID: $DistrictID
    SubdistrictID: $SubdistrictID
    BasinID: $BasinID
    SubbasinID: $SubbasinID
    OperationStartYear: $OperationStartYear
    ProductID: $ProductID
    ActivityID: $ActivityID
    ProjectTypeID: $ProjectTypeID
    BudgetRangeStart: $BudgetRangeStart
    BudgetRangeEnd: $BudgetRangeEnd
    PurchaseTypeID: $PurchaseTypeID
    PurchaseStepID: $PurchaseStepID
    ProgressResultRangeStart: $ProgressResultRangeStart
    ProgressResultRangeEnd: $ProgressResultRangeEnd
    UsedBudgetResultRangeStart: $UsedBudgetResultRangeStart
    UsedBudgetResultRangeEnd: $UsedBudgetResultRangeEnd
    OfficeOrganizationID: $OfficeOrganizationID
    OrganizationID: $OrganizationID
    BudgetSourceID: $BudgetSourceID
    BudgetTypeID: $BudgetTypeID
    ProjectCode: $ProjectCode
    ProjectSizeID: $ProjectSizeID
    TagID: $TagID
    RoyalID: $RoyalID
    PercentRangeStart: $PercentRangeStart
    PercentRangeEnd: $PercentRangeEnd
    IsCanceled: $IsCanceled
    IsMergeDraft: $IsMergeDraft
    VendorPurchaseFilter: $VendorPurchaseFilter
    SelfPurchaseFilter: $SelfPurchaseFilter
    TroubleTypeFilter: $TroubleTypeFilter
    UsedBudgetFilter: $UsedBudgetFilter
    ProgressPlanFilter: $ProgressPlanFilter
    ProjectTypeFilter: $ProjectTypeFilter
    KPIFilter: $KPIFilter
    KPISupFilter: $KPISupFilter
    KPISupNumber: $KPISupNumber
    PictureFilter: $PictureFilter
    ProgressResultID: $ProgressResultID
    ProgressResultInProgressID: $ProgressResultInProgressID
    BudgetDimensionID: $BudgetDimensionID
    SupervisorName: $SupervisorName
    IsExportImage: $IsExportImage
    Order: $Order
    Offset: $Offset
    Limit: $Limit
  ) {
    FileName
    Link
    __typename
  }
}"""

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

    # ── Login ──
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

    log("\n🗂️ STEP 3: เข้าหน้า export...")
    r3 = s.get(f"{BASE}/track/export?BudgetYear={BUDGET_YEAR}", timeout=30)
    log(f"  → {r3.status_code} | {r3.url}")

    # ── GraphQL ──
    log(f"\n🔗 STEP 4: POST GraphQL (v9)...")
    variables = {
        "BudgetYear": [BUDGET_YEAR],
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

    gql_payload = {
        "operationName": "ExportExcelProjectExport",
        "variables": variables,
        "query": GQL_QUERY,
    }

    r4 = s.post(GQL_URL, json=gql_payload, timeout=600, headers={
        "Content-Type": "application/json",
        "Accept": "application/json, */*",
        "Origin": BASE,
        "Referer": f"{BASE}/track/export?BudgetYear={BUDGET_YEAR}",
        "User-Agent": UA,
    })
    log(f"  → Status: {r4.status_code}")
    log(f"  → Response (600 chars): {r4.text[:600]}")

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
        log(f"❌ ไม่พบ data: {json.dumps(gql_resp, ensure_ascii=False)[:500]}")
        sys.exit(1)

    file_link = export_data.get("Link", "")
    file_name = export_data.get("FileName", "cenproject_data.xlsx")
    log(f"  ✅ FileName: {file_name}")
    log(f"  ✅ Link: {file_link}")

    if not file_link:
        log("❌ Link ว่างเปล่า")
        sys.exit(1)

    # ── Download ──
    download_url = file_link if file_link.startswith("http") else BASE + file_link
    log(f"\n📥 STEP 5: Download {download_url}")

    r5 = s.get(download_url, timeout=600,
               headers={"Accept": "*/*", "Referer": f"{BASE}/track/export",
                        "User-Agent": UA})
    log(f"  → Status: {r5.status_code} | Size: {len(r5.content):,} | Magic: {r5.content[:4].hex()}")

    is_xlsx = r5.content[:4] == b'PK\x03\x04'
    is_xls  = r5.content[:8] == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'
    if not (is_xlsx or is_xls):
        log(f"❌ ไม่ใช่ Excel: {r5.content[:200]}")
        sys.exit(1)

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    with open(OUTPUT_XLSX, "wb") as f: f.write(r5.content)
    sz = os.path.getsize(OUTPUT_XLSX)
    log(f"\n✅ บันทึก: {OUTPUT_XLSX} ({sz:,} bytes)")

    now = datetime.utcnow()
    meta = {
        "updated_at": now.isoformat() + "Z",
        "updated_th": f"{now.day:02d}/{now.month:02d}/{now.year+543} {(now.hour+7)%24:02d}:{now.minute:02d} ICT",
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
