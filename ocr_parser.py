import re
import os
import base64
import json
import google.auth
from google.cloud import vision

# 22 個已知品項代碼
KNOWN_CODES = {
    "RGT5F", "RGM5F", "RGG9B", "RGG3B", "RTT9B", "RTT3B",
    "RBZ9B", "RBZ3B", "RBZ5F", "RCB9B", "RCB3B", "GAB5F",
    "GBW5F", "GB2M",  "GLW5F", "GL2M",  "GCL5F", "GCL2M",
    "GHW5F", "GHW2M", "RYB5F", "KMT5F",
}

_client = None

def get_client():
    global _client
    if _client is None:
        credentials, _ = google.auth.default(
            scopes=["https://www.googleapis.com/auth/cloud-vision"]
        )
        _client = vision.ImageAnnotatorClient(credentials=credentials)
    return _client


def parse_transfer_order(image_path: str) -> dict:
    with open(image_path, "rb") as f:
        content = f.read()

    image    = vision.Image(content=content)
    response = get_client().document_text_detection(image=image)

    if response.error.message:
        raise RuntimeError(f"Vision API error: {response.error.message}")

    full_text = response.full_text_annotation.text
    items     = _extract_items(response)
    rows      = _group_by_y(items, tol=15)

    receiving_loc  = _extract_receiving_location(full_text)
    shipment_no    = _extract_shipment_no(full_text)
    warehouse_label = f"{receiving_loc} - {shipment_no}".strip() if receiving_loc else shipment_no

    result = {
        "date":         _extract_date(full_text),
        "warehouse":    warehouse_label,
        "shipment_no":  shipment_no,
        "matched_items":   [],
        "unmatched_items": [],
    }

    table_items = _parse_table(rows, full_text)
    for item in table_items:
        key = "matched_items" if item["code"] in KNOWN_CODES else "unmatched_items"
        result[key].append(item)

    return result


# ── 表頭欄位擷取 ──────────────────────────────────────────

def _extract_date(text: str) -> str:
    for pattern in [
        r'交貨日期[：:\s]*(\d{4}/\d{2}/\d{2})',
        r'裝載日期[：:\s]*(\d{4}/\d{2}/\d{2})',
    ]:
        m = re.search(pattern, text)
        if m:
            return m.group(1)
    return ""

def _extract_receiving_location(text: str) -> str:
    # 收貨儲位：1108 北區供貨倉庫 → 取名稱部分
    # 用多種停止條件相容 OCR（雙空格）與 PDF（關鍵字緊跟）兩種格式
    m = re.search(r'收貨儲位[：:\s]*\d+\s+(.+?)(?:\s{2,}|\t|\n|採購單號|裝運單號|發貨|$)', text)
    return m.group(1).strip() if m else ""

def _extract_shipping_point(text: str) -> str:
    # 出貨點：1403 宏全無菌三倉 → 取名稱部分（裝載明細表格式）
    m = re.search(r'出貨點[：:\s]*\d+\s+(.+?)(?:\s{2,}|\t|\n|裝載日期|裝運單號|$)', text)
    return m.group(1).strip() if m else ""

def _extract_shipment_no(text: str) -> str:
    m = re.search(r'裝運單號[：:\s]*(\d+)', text)
    return m.group(1) if m else ""


# ── 產品表格解析 ──────────────────────────────────────────

def _parse_table(rows: list, full_text: str) -> list:
    # 找表格標題列（含「序號」「品號」「數量」）
    header_idx, qty_x, pallets_x = _find_header(rows)
    if header_idx == -1:
        return _fallback_parse(full_text)

    result = []
    for row in rows[header_idx + 1:]:
        row_text = " ".join(i["text"] for i in row)

        # 跳過合計列與棧板列
        if any(kw in row_text for kw in ["合計", "全家棧板", "FPAL"]):
            continue

        code = _extract_code(row)
        if not code:
            continue

        qty     = _extract_qty(row, qty_x)
        pallets = _extract_pallets(row, pallets_x, qty_x, exclude_qty=qty)
        result.append({"code": code, "qty": qty, "pallets": pallets})

    return result


def _find_header(rows: list):
    """回傳 (header列index, 數量欄X座標, 棧板欄X座標)"""
    for i, row in enumerate(rows):
        text = "".join(item["text"] for item in row)
        if "序號" in text and ("品號" in text or "數量" in text):
            qty_x = next(
                (item["x"] for item in row if "數量" in item["text"]), None
            )
            pallets_x = next(
                (item["x"] for item in row if "棧板" in item["text"]), None
            )
            return i, qty_x, pallets_x
    return -1, None, None


def _extract_code(row_items: list) -> str:
    """從一列中擷取產品代碼（如 GLW5F）"""
    for item in sorted(row_items, key=lambda i: i["x"]):
        text = item["text"].strip()
        # 直接比對已知代碼
        if text in KNOWN_CODES:
            return text
        # 格式：「數字 代碼」（如 3383 GLW5F）
        m = re.match(r'^\d+\s+([A-Z]{2,5}\d[A-Z0-9]{0,3})$', text)
        if m and m.group(1) in KNOWN_CODES:
            return m.group(1)
        # 代碼混在長字串中
        for code in KNOWN_CODES:
            if code in text:
                return code
    return ""


def _extract_qty(row_items: list, qty_x) -> int:
    """從一列中擷取數量（優先取 X 座標最接近 qty_x 的純數字）"""
    candidates = []
    for item in row_items:
        clean = item["text"].replace(",", "").strip()
        if re.match(r'^\d+$', clean):
            n = int(clean)
            # 排除序號（1~9）和棧板數（通常 < 30）
            if n >= 10:
                candidates.append((item["x"], n))

    if not candidates:
        return 0

    if qty_x:
        return min(candidates, key=lambda c: abs(c[0] - qty_x))[1]

    # 無座標基準時取最大值（數量通常最大）
    return max(c[1] for c in candidates)


def _extract_pallets(row_items: list, pallets_x, qty_x=None, exclude_qty: int = 0) -> int:
    """從一列中擷取棧板數。
    雙重排除策略：
      1. 跳過 X ≤ qty_x 的候選（棧板欄在數量欄右側）
      2. 跳過值等於 qty 的候選（防止列對齊偏差時誤取數量值）
    """
    if pallets_x is None:
        return 0
    candidates = []
    for item in row_items:
        clean = item["text"].replace(",", "").strip()
        if re.match(r'^\d+$', clean):
            n = int(clean)
            if n < 1:
                continue
            if n == exclude_qty:          # 直接排除數量那個數字
                continue
            if qty_x is not None and item["x"] <= qty_x:
                continue
            candidates.append((item["x"], n))
    if not candidates:
        return 0
    return min(candidates, key=lambda c: abs(c[0] - pallets_x))[1]


# ── 降級解析（無法定位表格標題時）────────────────────────

def _fallback_parse(full_text: str) -> list:
    result = []
    for line in full_text.split("\n"):
        for code in KNOWN_CODES:
            if code in line:
                nums = re.findall(r'\b(\d{2,5})\b', line.replace(",", ""))
                candidates = [int(n) for n in nums if int(n) >= 10]
                qty = max(candidates) if candidates else 0
                result.append({"code": code, "qty": qty})
                break
    return result


# ── 工具函式 ─────────────────────────────────────────────

def _extract_items(response) -> list:
    items = []
    for page in response.full_text_annotation.pages:
        for block in page.blocks:
            for para in block.paragraphs:
                for word in para.words:
                    text = "".join(s.text for s in word.symbols).strip()
                    if not text:
                        continue
                    verts = word.bounding_box.vertices
                    xs = [v.x for v in verts]
                    ys = [v.y for v in verts]
                    items.append({
                        "text": text,
                        "x": (min(xs) + max(xs)) / 2,
                        "y": (min(ys) + max(ys)) / 2,
                    })
    return items


def _group_by_y(items: list, tol: int = 15) -> list:
    if not items:
        return []
    sorted_items = sorted(items, key=lambda i: i["y"])
    groups, cur  = [], [sorted_items[0]]
    ref_y        = sorted_items[0]["y"]
    for item in sorted_items[1:]:
        if abs(item["y"] - ref_y) <= tol:
            cur.append(item)
        else:
            groups.append(sorted(cur, key=lambda i: i["x"]))
            cur, ref_y = [item], item["y"]
    groups.append(sorted(cur, key=lambda i: i["x"]))
    return groups


# ── PDF 批次解析 ──────────────────────────────────────────

def parse_pdf_transfer_orders(pdf_path: str, filename: str = "") -> list:
    """Parse every page of a PDF transfer-order file.
    Each page is one order; returns a list in the same dict shape as parse_transfer_order."""
    import pdfplumber  # lazy import — keeps other endpoints alive even if pdfplumber is absent

    # 從檔名提取識別字（如「良鎮裝載明細表.pdf」→「良鎮」）
    filename_label = ""
    if filename:
        stem = re.sub(r'\.pdf$', '', filename, flags=re.IGNORECASE)  # 去副檔名
        stem = re.sub(r'^\d+\s*', '', stem)                          # 去開頭日期
        stem = re.sub(r'裝載明細表$', '', stem).strip()              # 去結尾
        filename_label = stem

    results = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text   = page.extract_text() or ""
            tables = page.extract_tables()

            if "出貨點" in text:
                # 裝載明細表：所別/備註 = 檔名識別字 + 裝運單號
                shipment_no     = _extract_shipment_no(text)
                warehouse_label = f"{filename_label} - {shipment_no}".strip() if filename_label else shipment_no
            else:
                receiving_loc   = _extract_receiving_location(text)
                shipment_no     = _extract_shipment_no(text)
                warehouse_label = f"{receiving_loc} - {shipment_no}".strip() if receiving_loc else shipment_no

            result = {
                "date":            _extract_date(text),
                "warehouse":       warehouse_label,
                "shipment_no":     shipment_no,
                "matched_items":   [],
                "unmatched_items": [],
            }

            for item in _parse_pdf_table(tables):
                key = "matched_items" if item["code"] in KNOWN_CODES else "unmatched_items"
                result[key].append(item)

            if result["matched_items"] or result["unmatched_items"]:
                results.append(result)

    return results


def _parse_pdf_table(tables: list) -> list:
    """Extract items from a pdfplumber table list for one page."""
    for table in tables:
        if not table:
            continue

        # Find header row: 調撥單格式（序號+品號）或裝載表格式（商品編號+數量）
        header_idx = -1
        for i, row in enumerate(table):
            row_text = " ".join(str(c) for c in row if c is not None)
            if ("序號" in row_text and ("品號" in row_text or "數量" in row_text)) or \
               ("商品編號" in row_text and "數量" in row_text):
                header_idx = i
                break

        if header_idx == -1:
            continue

        header_row = table[header_idx]

        # Locate column indices from header
        code_col = qty_col = pallets_col = None
        for j, cell in enumerate(header_row):
            if cell is None:
                continue
            s = str(cell)
            if ("品號" in s or "產品小名" in s) and code_col is None:
                code_col = j
            if "數量" in s and qty_col is None:
                qty_col = j
            if "棧板" in s and pallets_col is None:
                pallets_col = j

        if code_col is None:
            code_col = 1  # fallback

        result = []
        for row in table[header_idx + 1:]:
            if not row:
                continue
            row_text = " ".join(str(c) for c in row if c is not None)
            if any(kw in row_text for kw in ["合計", "全家棧板", "FPAL", "第一聯", "第二聯", "成品倉庫"]):
                continue

            code_cell = str(row[code_col] or "").strip() if code_col < len(row) else ""
            code = _extract_code_from_cell(code_cell)
            if not code:
                continue

            qty = 0
            if qty_col is not None and qty_col < len(row):
                raw = str(row[qty_col] or "").replace(",", "").strip()
                if re.match(r'^\d+$', raw):
                    qty = int(raw)

            pallets = 0
            if pallets_col is not None and pallets_col < len(row):
                raw = str(row[pallets_col] or "").replace(",", "").strip()
                if re.match(r'^\d+$', raw):
                    pallets = int(raw)

            result.append({"code": code, "qty": qty, "pallets": pallets})

        return result

    return []


def _extract_code_from_cell(text: str) -> str:
    """Extract product code from a PDF table cell like '3441 RGT5F' or 'RGT5F'."""
    text = text.strip()
    if not text:
        return ""
    if text in KNOWN_CODES:
        return text
    # "number code" format
    m = re.match(r'^\d+\s+([A-Z][A-Z0-9]{2,7})$', text)
    if m:
        return m.group(1)
    # Known code embedded in text
    for code in KNOWN_CODES:
        if code in text:
            return code
    # Generic code pattern
    m = re.search(r'\b([A-Z]{2,5}\d[A-Z0-9]{0,3})\b', text)
    if m:
        return m.group(1)
    return ""


# ── 味全出貨單解析 ────────────────────────────────────────────

def parse_weichuan_transfer_order(image_path: str) -> dict:
    """Parse 味全 出貨單 image (JPG/PNG).
    Returns dict: {date, warehouse, matched_items, unmatched_items}
    matched_items: [{code, qty}] where code ∈ {"黑咖啡_全台","黑咖啡_4入","拿鐵_全台","拿鐵_4入"}
    """
    with open(image_path, "rb") as f:
        content = f.read()

    image    = vision.Image(content=content)
    response = get_client().document_text_detection(image=image)

    if response.error.message:
        raise RuntimeError(f"Vision API error: {response.error.message}")

    return _parse_weichuan_ocr_response(response)


def parse_weichuan_pdf(pdf_path: str) -> dict:
    """Parse 味全 出貨單 as scanned PDF (converts to image, then OCR)."""
    from pdf2image import convert_from_path
    import tempfile, os

    images = convert_from_path(pdf_path, dpi=200)
    if not images:
        return {"date": "", "warehouse": "", "matched_items": [], "unmatched_items": []}

    # Use the first page (one order per file)
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
        images[0].save(tmp.name, format="PNG")
        tmp_path = tmp.name

    try:
        return parse_weichuan_transfer_order(tmp_path)
    finally:
        os.unlink(tmp_path)


def _parse_weichuan_ocr_response(response) -> dict:
    full_text = response.full_text_annotation.text
    items     = _extract_items(response)
    rows      = _group_by_y(items, tol=15)

    date           = _extract_weichuan_date(full_text)
    receiving_unit = _extract_weichuan_receiving_unit(full_text)
    doc_no         = _extract_weichuan_doc_no(full_text)
    warehouse      = f"{receiving_unit}{doc_no}".strip()

    products = _parse_weichuan_rows(rows, full_text)

    return {
        "date":            date,
        "warehouse":       warehouse,
        "matched_items":   products,
        "unmatched_items": [],
    }


def _roc_to_gregorian(roc_date: str) -> str:
    """Convert ROC date string (e.g. '115.04.23') to '2026/04/23'."""
    parts = roc_date.split(".")
    if len(parts) != 3:
        return roc_date
    try:
        year = int(parts[0]) + 1911
        return f"{year}/{parts[1]}/{parts[2]}"
    except ValueError:
        return roc_date


def _extract_weichuan_date(text: str) -> str:
    """Extract 預計到貨日 and convert from ROC to Gregorian."""
    # Try inline: 預計到貨日：115.04.23
    m = re.search(r'預計到貨日[：:\s]+(\d{2,3}\.\d{2}\.\d{2})(?!-)', text)
    if m:
        return _roc_to_gregorian(m.group(1))
    # Try multi-line: 預計到貨日 on one line, date on next
    m = re.search(r'預計到貨日[：:\s]*\n+\s*(\d{2,3}\.\d{2}\.\d{2})(?!-)', text)
    if m:
        return _roc_to_gregorian(m.group(1))
    # Fallback: first ROC date without dash (avoid matching the 單據編號)
    # Only match if NOT preceded by another number-dot pattern
    for seg in re.finditer(r'(\d{2,3}\.\d{2}\.\d{2})(?!-)', text):
        return _roc_to_gregorian(seg.group(1))
    return ""


def _extract_weichuan_receiving_unit(text: str) -> str:
    """Extract 收貨單位 name (e.g. '味全斗六廠').

    When 收貨單位：and 出貨日期：are on the same OCR line the naive regex
    captures '出貨日期' instead of the factory name.  We instead search
    directly for the '味全XX廠' pattern, excluding the company header
    ('味全食品工業…') and footer ('味全公司XX廠').
    """
    # Direct search: 味全 not followed by 食品 or 公司, ending in 廠
    m = re.search(r'味全(?!食品|公司)(\S+廠)', text)
    if m:
        return "味全" + m.group(1)
    # Fallback: look on the line *after* 收貨單位：
    m = re.search(r'收貨單位[：:][^\n]*\n\s*([^\n\t出貨]+?)(?:\s{2,}|\t|\n|出貨日期|$)', text)
    if m:
        return m.group(1).strip()
    return ""


def _extract_weichuan_doc_no(text: str) -> str:
    """Extract 單據編號 (e.g. '115.04.23-02')."""
    m = re.search(r'單據編號[：:\s]*\n*\s*(\d{2,3}\.\d{2}\.\d{2}-\d+)', text)
    if m:
        return m.group(1)
    # Fallback: any XXX.XX.XX-XX pattern
    m = re.search(r'\b(\d{3}\.\d{2}\.\d{2}-\d+)\b', text)
    return m.group(1) if m else ""


def _parse_weichuan_rows(rows: list, full_text: str) -> list:
    """Extract products from coordinate-grouped OCR rows.

    Two-pass strategy:
      1. Coordinate-based: group items by Y, match product+qty in same row.
         When qty is missing from the product row, look ahead ±2 rows.
      2. Text-based fallback: scan full_text lines for product keywords
         and the "N 箱" pattern on the same or adjacent line.
    """
    # ── Pass 1: coordinate-based ──────────────────────────────────────────
    header_idx = qty_x = None
    for i, row in enumerate(rows):
        text = "".join(item["text"] for item in row)
        if "品名" in text and ("數量" in text or "序號" in text):
            header_idx = i
            qty_x = next((item["x"] for item in row if "數量" in item["text"]), None)
            break

    SKIP_KW = {"合計", "本車", "承運", "收貨單位", "警檢", "備註", "貨運行"}
    result   = []
    found    = set()
    scan_rows = rows[header_idx + 1:] if header_idx is not None else rows

    for i, row in enumerate(scan_rows):
        row_text = " ".join(item["text"] for item in row)
        if any(kw in row_text for kw in SKIP_KW):
            continue
        code = _detect_weichuan_product(row_text)
        if not code or code in found:
            continue

        qty = _extract_weichuan_qty(row, row_text, qty_x)
        # If qty not on same row, look in adjacent rows (±2)
        if qty == 0:
            for offset in range(1, 3):
                for delta in [offset, -offset]:
                    adj = i + delta
                    if 0 <= adj < len(scan_rows):
                        adj_text = " ".join(item["text"] for item in scan_rows[adj])
                        qty = _extract_weichuan_qty(scan_rows[adj], adj_text, qty_x)
                        if qty:
                            break
                if qty:
                    break

        if qty > 0:
            result.append({"code": code, "qty": qty})
            found.add(code)

    if result:
        return result

    # ── Pass 2: text-based fallback ───────────────────────────────────────
    # Only trust "N 箱" as qty source — raw numbers include "(340ml)" noise.
    lines = full_text.split("\n")
    for i, line in enumerate(lines):
        code = _detect_weichuan_product(line)
        if not code or code in found:
            continue
        # Search current line + next 2 lines for "N 箱"
        qty = 0
        for j in range(i, min(i + 3, len(lines))):
            qty = _extract_weichuan_qty_from_text(lines[j])
            if qty:
                break
        if qty > 0:
            result.append({"code": code, "qty": qty})
            found.add(code)

    return result


def _extract_weichuan_qty_from_text(text: str) -> int:
    """Extract quantity using only the '箱' unit marker.
    Never falls back to raw number guessing — product names contain
    '(340ml)' which would produce a wrong qty of 340.
    """
    m = re.search(r'([\d,]+)\s*箱', text)
    return int(m.group(1).replace(",", "")) if m else 0


def _detect_weichuan_product(text: str) -> str:
    """Map product name text to internal code."""
    has_4ru = "4入" in text
    if "拿鐵" in text:
        return "拿鐵_4入" if has_4ru else "拿鐵_全台"
    if "黑咖啡" in text:
        return "黑咖啡_4入" if has_4ru else "黑咖啡_全台"
    return ""


def _extract_weichuan_qty(row_items: list, row_text: str, qty_x=None) -> int:
    """Extract quantity (箱 count) from a product row."""
    # Prefer explicit "N 箱" pattern
    m = re.search(r'([\d,]+)\s*箱', row_text)
    if m:
        return int(m.group(1).replace(",", ""))
    # Coordinate-based fallback
    candidates = []
    for item in row_items:
        clean = item["text"].replace(",", "").strip()
        if re.match(r'^\d+$', clean):
            n = int(clean)
            if n >= 1:
                candidates.append((item["x"], n))
    if not candidates:
        return 0
    if qty_x:
        return min(candidates, key=lambda c: abs(c[0] - qty_x))[1]
    return max(c[1] for c in candidates)


# ── 維他露 section ──────────────────────────────────────────────────────────

_VTL_CUSTOMER_FIXES = {
    "員林中食": "員林中倉",
    "貝林中食": "員林中倉",
    "貝林中倉": "員林中倉",
}


def _normalize_vtl_code(raw: str) -> str:
    """
    Convert OCR product code to template format.
    e.g. "JDP05901A61" → "JDP0590 1A61"
    Also corrects common OCR char confusions at digit positions (3-6).
    """
    code = re.sub(r"^\d+-", "", raw.strip())
    chars = list(code)
    for i in range(3, min(7, len(chars))):
        if chars[i] == "O": chars[i] = "0"
        if chars[i] == "S": chars[i] = "5"
        if chars[i] == "I": chars[i] = "1"
        if chars[i] == "l": chars[i] = "1"
    code = "".join(chars)
    if len(code) >= 11:
        return code[:7] + " " + code[7:]
    return code


def _parse_vtl_qty(text: str) -> int:
    """
    Parse printed quantity from an OCR line, ignoring appended handwriting.

    Handles:
      "2702/27"  → 270   (handwriting "2/27" glued to "270")
      "1.4306"   → 1430  (comma→period, "1,430" misread)
      "520 25"   → 520   (trailing handwriting after space)
      "540"      → 540
    """
    text = text.strip()
    # Full date strings (DD/MM/YY or D/M/YY) are not quantities
    if re.match(r"^\d{1,2}/\d{2}/\d{2}", text):
        return 0
    # Standalone handwriting date annotations like "2/3", "3/6", "1/8" are not quantities
    if re.match(r"^\d{1,2}/\d{1,2}$", text):
        return 0
    # Handwriting appended as "N/D" or "N/DD" — stop before digit-slash-digit(s)
    m = re.match(r"^(\d+?)(?=[1-9]/\d{1,2})", text)
    if m and len(m.group(1)) >= 2:
        return int(m.group(1))
    # Thousands separator misread as decimal "1.430"
    m = re.match(r"^(\d+)\.(\d{3})", text)
    if m:
        return int(m.group(1) + m.group(2))
    # Comma thousands separator "1,430"
    m = re.match(r"^(\d+),(\d{3})", text)
    if m:
        return int(m.group(1) + m.group(2))
    # Plain number (take leading digits, ignore trailing noise)
    m = re.match(r"^(\d+)", text)
    if m:
        return int(m.group(1))
    return 0


def _parse_vtl_blocks(full_text: str) -> list:
    """
    Split OCR text of a 維他露 調撥單 page into order blocks.
    Each block → one row in the 代工明細表.
    Returns list of {date, warehouse, matched_items, unmatched_items}.
    """
    # 营/営/營 are CJK variants of the same character — match all three
    raw_blocks = re.split(r"(?=\d+\s+[营営營]業部)", full_text)
    results = []

    for block in raw_blocks:
        if not re.search(r"[营営營]業部", block):
            continue
        lines = [l.strip() for l in block.split("\n") if l.strip()]

        # Trim at summary/footer section to prevent totals from being parsed as quantities
        trimmed = []
        for ln in lines:
            if re.match(r"^-{3,}", ln) or re.match(r"^總計", ln) or re.match(r"^單位\b", ln):
                break
            trimmed.append(ln)
        if trimmed:
            lines = trimmed

        date = ""
        order_no = ""
        destination = ""
        items = []

        # Extract date (排出日期) and order number from line containing VR1-
        for line in lines:
            m = re.search(r"(\d{2}/\d{2}/\d{2})\s+VR1?-(\d+)", line)
            if m:
                parts = m.group(1).split("/")
                date = f"20{parts[0]}/{parts[1]}/{parts[2]}"
                order_no = f"VSR-{m.group(2)}"
                break

        # Extract destination: first meaningful line after VR1- line
        found_vr = False
        for line in lines:
            if re.search(r"VR1?-\d+", line):
                found_vr = True
                continue
            if not found_vr:
                continue
            if re.match(r"^DP\d+", line):     continue
            if re.match(r"^\d{2}/\d{2}/\d{2}", line): continue
            if re.match(r"^SM[A-Z0-9]+", line): continue
            if re.match(r"^\d+\s*[营営營]業部", line): continue
            if re.match(r"^備註", line):        continue
            if "(" in line and "宏" in line:    continue
            # Must look like a Chinese location name
            if re.search(r"[一-鿿]", line) and len(line) >= 2:
                destination = _VTL_CUSTOMER_FIXES.get(line, line)
                break

        # Extract items: two-pass approach.
        # OCR sometimes outputs all product-code lines first, then CTN/qty lines,
        # so a simple scan-ahead would steal a later item's CTN for an earlier one.
        # Pass 1: collect item code positions in order.
        item_starts = []
        for ii, ln in enumerate(lines):
            mm = re.match(r"\d+-([A-Z]{3}[\dOSIl]{4}\w+)", ln)
            if mm:
                item_starts.append((ii, mm.group(1)))

        _IP = re.compile(r"\d+-[A-Z]{3}[\dOSIl]{4}\w+")
        claimed_j = set()

        def _find_qty(start, end, skip_after=None):
            """Return first unclaimed CTN qty in lines[start:end].
            skip_after: skip lines that are item-code lines at j > skip_after."""
            for j in range(start, end):
                if j in claimed_j:
                    continue
                ln2 = lines[j]
                if skip_after is not None and j > skip_after and _IP.match(ln2):
                    continue
                im = re.search(r"(?:CIN|CTN)\s+(.*)", ln2)
                if im:
                    q = 0
                    for tok in im.group(1).split():
                        q = _parse_vtl_qty(tok)
                        if q > 0:
                            break
                    if q == 0:
                        for k in range(j + 1, min(j + 3, len(lines))):
                            if k not in claimed_j:
                                q = _parse_vtl_qty(lines[k])
                                if q > 0:
                                    break
                    if q > 0:
                        claimed_j.add(j)
                        return q
                elif re.match(r"^(?:CIN|CTN)$", ln2):
                    q = 0
                    for k in range(j + 1, min(j + 3, len(lines))):
                        if k not in claimed_j:
                            q = _parse_vtl_qty(lines[k])
                            if q > 0:
                                break
                    if q > 0:
                        claimed_j.add(j)
                        return q
            return 0

        # Pass 2: assign each item its CTN quantity.
        # First try within territory (before next item's line);
        # fall back to full-block scan skipping other item-code lines.
        for n, (ii, raw_code) in enumerate(item_starts):
            normalized = _normalize_vtl_code(raw_code)
            nxt = item_starts[n + 1][0] if n + 1 < len(item_starts) else min(ii + 7, len(lines))
            qty = _find_qty(ii, nxt)
            if qty == 0:
                qty = _find_qty(ii, len(lines), skip_after=ii)
            if qty > 0:
                items.append({"code": normalized, "qty": qty})

        if items:
            wh = (destination + order_no) if destination else order_no
            matched = [{"code": it["code"], "qty": it["qty"]} for it in items]
            results.append({
                "date":            date,
                "warehouse":       wh,
                "matched_items":   matched,
                "unmatched_items": [],
            })

    return results


_VTL_CLAUDE_PROMPT = """你正在解析維他露食品股份有限公司的調撥單（訂明細表）。

請提取圖片中每個「N 營業部」區塊的資訊，回傳 JSON。

格式：
{
  "orders": [
    {
      "date": "YYYY/MM/DD",
      "warehouse": "倉庫名稱VSR-XXXXXX",
      "items": [
        {"code": "XXXXXXXXXXX", "qty": 數量},
        ...
      ]
    }
  ]
}

規則：
1. date：排出日期，YY/MM/DD 轉為 YYYY/MM/DD（如 26/05/07 → 2026/05/07）
2. warehouse：送貨客戶倉庫名稱（如「巨航倉」「員林中倉」）加上 VSR-XXXXXX 訂單編號，合併為一個字串
3. code：品號原文，不含前面的序號（如 1-），不加空格（如 JAP05801B71、JDP05901BT1）
4. qty：該品項印刷的阿拉伯數字數量，忽略旁邊所有手寫標記（如 2/3、3/6、1/8 等）
5. 每個品項前面有 CTN 或 CIN 標記，數量在 CTN/CIN 之後
6. 只回傳 JSON，不要任何說明文字"""


def parse_vtl_with_claude(image_path: str) -> list:
    """Parse 維他露 調撥單 using Claude Vision API. Returns same format as parse_vtl_transfer_order."""
    import anthropic

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise RuntimeError("ANTHROPIC_API_KEY not set")

    with open(image_path, "rb") as f:
        img_data = base64.standard_b64encode(f.read()).decode("utf-8")

    ext = os.path.splitext(image_path)[1].lower()
    media_type_map = {".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png", ".webp": "image/webp"}
    media_type = media_type_map.get(ext, "image/jpeg")

    client = anthropic.Anthropic(api_key=api_key)
    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=4096,
        messages=[{
            "role": "user",
            "content": [
                {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": img_data}},
                {"type": "text", "text": _VTL_CLAUDE_PROMPT},
            ],
        }],
    )

    raw = message.content[0].text.strip()
    # Strip markdown code fences if present
    if raw.startswith("```"):
        raw = re.sub(r"^```[a-z]*\n?", "", raw)
        raw = re.sub(r"\n?```$", "", raw)

    data = json.loads(raw)
    results = []
    for order in data.get("orders", []):
        items = []
        for it in order.get("items", []):
            code = _normalize_vtl_code(str(it.get("code", "")))
            qty = int(it.get("qty", 0))
            if code and qty > 0:
                items.append({"code": code, "qty": qty})
        if items:
            results.append({
                "date":            order.get("date", ""),
                "warehouse":       order.get("warehouse", ""),
                "matched_items":   items,
                "unmatched_items": [],
            })
    return results


def parse_vtl_transfer_order(image_path: str) -> list:
    """Parse a 維他露 調撥單 image. Returns list of order blocks."""
    with open(image_path, "rb") as f:
        content = f.read()
    client = get_client()
    image = vision.Image(content=content)
    response = client.document_text_detection(image=image)
    if response.error.message:
        raise RuntimeError(response.error.message)
    return _parse_vtl_blocks(response.full_text_annotation.text)


def parse_vtl_pdf(pdf_path: str) -> list:
    """Parse a 維他露 PDF (multiple pages). Returns all order blocks."""
    import tempfile
    from pdf2image import convert_from_path

    images = convert_from_path(pdf_path)
    all_blocks = []
    for img in images:
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
            img.save(tmp.name, "PNG")
            tmp_path = tmp.name
        try:
            blocks = parse_vtl_transfer_order(tmp_path)
            all_blocks.extend(blocks)
        finally:
            import os as _os
            _os.unlink(tmp_path)
    return all_blocks
