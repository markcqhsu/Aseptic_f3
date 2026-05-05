import re
import google.auth
from google.cloud import vision
import pdfplumber

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
    warehouse_label = f"{receiving_loc} {shipment_no}".strip() if receiving_loc else shipment_no

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

def parse_pdf_transfer_orders(pdf_path: str) -> list:
    """Parse every page of a PDF transfer-order file.
    Each page is one order; returns a list in the same dict shape as parse_transfer_order."""
    results = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text   = page.extract_text() or ""
            tables = page.extract_tables()

            receiving_loc   = _extract_receiving_location(text)
            shipment_no     = _extract_shipment_no(text)
            warehouse_label = f"{receiving_loc} {shipment_no}".strip() if receiving_loc else shipment_no

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

        # Find header row containing 序號 and 品號/數量
        header_idx = -1
        for i, row in enumerate(table):
            row_text = " ".join(str(c) for c in row if c is not None)
            if "序號" in row_text and ("品號" in row_text or "數量" in row_text):
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
            if "品號" in s and code_col is None:
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
            if any(kw in row_text for kw in ["合計", "全家棧板", "FPAL", "第一聯", "第二聯"]):
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
