from __future__ import annotations

import hashlib
import html
import json
import os
import re
import zipfile
from http.cookiejar import CookieJar
from datetime import datetime, timezone
from io import BytesIO
from urllib.parse import parse_qsl, urlencode, urlparse, urlunparse
from urllib.request import HTTPCookieProcessor, Request, build_opener
from xml.etree import ElementTree as ET

ONEDRIVE_SHARE_URL = os.environ.get("ONEDRIVE_SHARE_URL", "").strip()
NS = {
    "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def append_download_parameter(url: str) -> str:
    parsed = urlparse(url)
    query = dict(parse_qsl(parsed.query, keep_blank_values=True))
    query["download"] = "1"
    return urlunparse(parsed._replace(query=urlencode(query)))


def valid_xlsx_bytes(content: bytes) -> bool:
    if not content.startswith(b"PK"):
        return False
    try:
        with zipfile.ZipFile(BytesIO(content)) as archive:
            return "xl/workbook.xml" in archive.namelist()
    except zipfile.BadZipFile:
        return False


def make_opener():
    # OneDrive first opens/redeems the public share page, then uses session
    # cookies for the file download request.  Keeping one cookie jar is
    # essential for the current 1drv.ms / migratedtospo share-link format.
    return build_opener(HTTPCookieProcessor(CookieJar()))


def fetch_url(opener, url: str) -> tuple[bytes, str, str]:
    request = Request(
        url,
        headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124 Safari/537.36",
            "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream,text/html;q=0.8,*/*;q=0.5",
            "Cache-Control": "no-cache",
        },
    )
    with opener.open(request, timeout=25) as response:
        return response.read(), response.geturl(), response.headers.get("Content-Type", "")


def decode_embedded_url(value: str) -> str:
    value = html.unescape(value)
    value = value.replace("\\u0026", "&").replace("\\/", "/")
    value = value.replace("\u0026", "&")
    return value


def extract_download_candidates(page: bytes) -> list[str]:
    # Excel/OneDrive viewer HTML sometimes embeds a temporary download URL.
    text = page.decode("utf-8", errors="ignore")
    urls = []
    patterns = [
        r'"(?:downloadUrl|downloadURL|@microsoft\.graph\.downloadUrl)"\s*:\s*"([^"]+)"',
        r'(https?://[^"\s<>]+(?:download|download.aspx|\.xlsx)[^"\s<>]*)',
    ]
    for pattern in patterns:
        for match in re.findall(pattern, text, flags=re.IGNORECASE):
            url = decode_embedded_url(match)
            if url.startswith("http"):
                urls.append(url)
    return list(dict.fromkeys(urls))


def download_workbook() -> bytes:
    share_url = ONEDRIVE_SHARE_URL
    override_url = os.environ.get("ONEDRIVE_DOWNLOAD_URL", "").strip()
    if not share_url and not override_url:
        raise RuntimeError("ONEDRIVE_SHARE_URL is not configured in Vercel Environment Variables.")

    opener = make_opener()
    candidates = []
    diagnostics = []

    if override_url:
        candidates.append(override_url)

    resolved = None
    if share_url:
        try:
            landing, resolved, ctype = fetch_url(opener, share_url)
            if valid_xlsx_bytes(landing):
                return landing
            candidates.extend(extract_download_candidates(landing))
            candidates.append(append_download_parameter(share_url))
            if resolved:
                candidates.append(append_download_parameter(resolved))
                parsed = urlparse(resolved)
                query = dict(parse_qsl(parsed.query, keep_blank_values=True))
                # Consumer OneDrive's viewer normally exposes resid; after the
                # initial request cookies carry the public-share redemption.
                if query.get("resid"):
                    download_query = {k: v for k, v in query.items() if k in {"resid", "authkey", "redeem", "cid"}}
                    candidates.append("https://onedrive.live.com/download?" + urlencode(download_query))
                    candidates.append("https://onedrive.live.com/download?" + parsed.query)
            diagnostics.append(f"landing={ctype or 'unknown'}")
        except Exception as exc:
            diagnostics.append("landing request failed: " + str(exc))

    for candidate in list(dict.fromkeys(candidates)):
        try:
            content, final_url, ctype = fetch_url(opener, candidate)
            if valid_xlsx_bytes(content):
                return content
            for embedded in extract_download_candidates(content):
                embedded_content, _, _ = fetch_url(opener, embedded)
                if valid_xlsx_bytes(embedded_content):
                    return embedded_content
            diagnostics.append(f"not-xlsx ({ctype or 'unknown'}) from {urlparse(final_url).netloc}")
        except Exception as exc:
            diagnostics.append("candidate failed: " + str(exc))

    detail = "; ".join(diagnostics[-4:])
    raise RuntimeError(
        "OneDrive opened the public Excel viewer but did not return the workbook file. "
        "This build now keeps OneDrive cookies and retries direct-download paths. "
        "If this message remains, the link type cannot be downloaded anonymously and the connection must use Microsoft Graph authentication. "
        + detail
    )


def col_index(cell_ref: str) -> int:
    letters = re.match(r"[A-Z]+", cell_ref or "A").group(0)
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n - 1


def as_number(value):
    if value in (None, ""):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).replace(",", "").replace("٬", "").strip())
    except ValueError:
        return 0.0


def clean_text(value) -> str:
    return "" if value is None else str(value).strip()


def parse_xlsx(content: bytes) -> dict[str, list[list]]:
    with zipfile.ZipFile(BytesIO(content)) as z:
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for item in root.findall("a:si", NS):
                shared.append("".join(t.text or "" for t in item.findall(".//a:t", NS)))
        workbook = ET.fromstring(z.read("xl/workbook.xml"))
        rels_root = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        rels = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels_root.findall("rel:Relationship", NS)}
        sheet_paths = {}
        for sh in workbook.find("a:sheets", NS):
            name = sh.attrib["name"].strip()
            target = rels[sh.attrib[f"{{{NS['r']}}}id"]]
            if not target.startswith("xl/"):
                target = "xl/" + target.lstrip("/")
            sheet_paths[name] = target
        result = {}
        for name, path in sheet_paths.items():
            root = ET.fromstring(z.read(path))
            matrix = []
            for row in root.findall(".//a:sheetData/a:row", NS):
                row_values = []
                for cell in row.findall("a:c", NS):
                    idx = col_index(cell.attrib.get("r", "A1"))
                    while len(row_values) <= idx:
                        row_values.append(None)
                    cell_type = cell.attrib.get("t")
                    v_node = cell.find("a:v", NS)
                    inline = cell.find("a:is", NS)
                    raw = None
                    if cell_type == "inlineStr" and inline is not None:
                        raw = "".join(t.text or "" for t in inline.findall(".//a:t", NS))
                    elif v_node is not None:
                        raw = v_node.text
                        if cell_type == "s":
                            raw = shared[int(raw)]
                        elif cell_type in (None, "n"):
                            try:
                                raw = float(raw)
                                if raw.is_integer():
                                    raw = int(raw)
                            except (ValueError, AttributeError):
                                pass
                    row_values[idx] = raw
                matrix.append(row_values)
            result[name] = matrix
        return result


def get(matrix: list[list], row: int, col: int):
    if col < 0 or row >= len(matrix) or col >= len(matrix[row]):
        return None
    return matrix[row][col]


def normalize_month(label: str) -> str:
    return " ".join(clean_text(label).replace("\n", " ").split())


def risk_for(task: dict) -> str:
    budget, paid, progress = task["budget"], task["paid"], task["progress"]
    completed = task["completed"] in {"نعم", "Yes", "yes"} or progress >= 0.999
    if budget <= 0 and paid > 0:
        return "غير مخطط"
    if paid > budget and budget > 0:
        return "تجاوز الميزانية"
    if completed:
        return "مكتملة"
    spend_ratio = paid / budget if budget else 0
    if spend_ratio >= 0.80 and progress < 0.80:
        return "عالية المخاطر"
    if paid > 0 and progress < 0.25:
        return "تحتاج متابعة"
    return "ضمن الخطة"


def category_for(name: str) -> str:
    if any(k in name for k in ["الأرض", "شراء الأرض", "غانم"]): return "شراء الأرض"
    if any(k in name for k in ["فسائل", "مصدات رياح", "المشتل", "تحضين"]): return "الزراعة والفسائل"
    if any(k in name for k in ["بئر", "ري", "طلمبة", "بركة"]): return "المياه والري"
    if any(k in name for k in ["طاقة", "كابلات", "شبكات"]): return "الطاقة والبنية التحتية"
    if any(k in name for k in ["طرق", "تسوية", "حفر الجور", "أعمال مدنية"]): return "تجهيز الموقع"
    if any(k in name for k in ["عمالة", "إستشارات"]): return "التشغيل والخدمات"
    if any(k in name for k in ["سيارة", "انتقالات"]): return "النقل"
    if any(k in name for k in ["محامي", "متفرقات"]): return "إدارية ومتفرقات"
    return "أخرى"


def extract_dashboard_data(content: bytes) -> dict:
    sheets = parse_xlsx(content)
    expenses = sheets.get("المصروفات", [])
    capital = sheets.get("رأس المال", [])
    misc = sheets.get("المتفرقات", [])
    if len(expenses) < 3:
        raise RuntimeError("The worksheet المصروفات is empty or not found.")
    top, sub = expenses[0], expenses[1]
    top_filled, current = [], None
    for i in range(max(len(top), len(sub))):
        value = get(expenses, 0, i)
        if value not in (None, ""):
            current = value
        top_filled.append(current)
    months, month_cols = [], {}
    for i in range(8, max(len(top_filled), len(sub))):
        month = normalize_month(top_filled[i])
        metric = clean_text(get(expenses, 1, i))
        if not month:
            continue
        if month not in months:
            months.append(month)
        month_cols.setdefault(month, {})["actual" if "الفعلي" in metric else "planned"] = i
    tasks = []
    for row_idx in range(2, len(expenses)):
        name = clean_text(get(expenses, row_idx, 1))
        row_id = get(expenses, row_idx, 0)
        if not name or name == "الإجمالي" or row_id in (None, ""):
            continue
        task = {
            "id": row_id,
            "name": name,
            "budget": as_number(get(expenses, row_idx, 2)),
            "paid": as_number(get(expenses, row_idx, 3)),
            "remaining": as_number(get(expenses, row_idx, 4)),
            "completed": clean_text(get(expenses, row_idx, 5)),
            "progress": as_number(get(expenses, row_idx, 6)),
            "savings": as_number(get(expenses, row_idx, 7)),
            "monthly": {},
            "category": category_for(name),
        }
        if task["progress"] > 1:
            task["progress"] /= 100
        for month in months:
            cols = month_cols.get(month, {})
            task["monthly"][month] = {
                "actual": as_number(get(expenses, row_idx, cols.get("actual", -1))),
                "planned": as_number(get(expenses, row_idx, cols.get("planned", -1))),
            }
        task["spend_ratio"] = task["paid"] / task["budget"] if task["budget"] else (1 if task["paid"] else 0)
        task["risk"] = risk_for(task)
        tasks.append(task)
    partners = []
    for row_idx in range(1, min(len(capital), 9)):
        partner = clean_text(get(capital, row_idx, 1))
        if partner and "الإجمالي" not in partner:
            partners.append({"name": partner, "share": as_number(get(capital, row_idx, 2)), "paid": as_number(get(capital, row_idx, 3)), "remaining": as_number(get(capital, row_idx, 4))})
    misc_items = []
    for row_idx in range(2, len(misc)):
        name = clean_text(get(misc, row_idx, 1))
        paid = as_number(get(misc, row_idx, 3))
        if name and "الإجمالي" not in name and paid > 0:
            misc_items.append({"name": name, "paid": paid})
    total_budget = sum(t["budget"] for t in tasks)
    total_paid = sum(t["paid"] for t in tasks)
    total_remaining = sum(t["remaining"] for t in tasks)
    capital_paid = sum(p["paid"] for p in partners)
    return {
        "version": hashlib.sha256(content).hexdigest(),
        "loaded_at": datetime.now(timezone.utc).isoformat(),
        "source_mode": "onedrive",
        "overview": {
            "total_budget": total_budget,
            "total_paid": total_paid,
            "total_remaining": total_remaining,
            "total_savings": sum(t["savings"] for t in tasks),
            "spend_ratio": total_paid / total_budget if total_budget else 0,
            "capital_paid": capital_paid,
            "cash_remaining": capital_paid - total_paid,
        },
        "months": months,
        "monthly": [{"month": m, "actual": sum(t["monthly"][m]["actual"] for t in tasks), "planned": sum(t["monthly"][m]["planned"] for t in tasks)} for m in months],
        "tasks": tasks,
        "partners": partners,
        "misc": misc_items,
    }


def app(environ, start_response):
    try:
        body = json.dumps(extract_dashboard_data(download_workbook()), ensure_ascii=False).encode("utf-8")
        status = "200 OK"
    except Exception as exc:
        body = json.dumps({"error": str(exc)}, ensure_ascii=False).encode("utf-8")
        status = "500 Internal Server Error"
    start_response(status, [
        ("Content-Type", "application/json; charset=utf-8"),
        ("Cache-Control", "no-store, max-age=0"),
        ("Content-Length", str(len(body))),
    ])
    return [body]
