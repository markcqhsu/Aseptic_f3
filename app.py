import os
import tempfile
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS

from ocr_parser import (
    parse_transfer_order, parse_pdf_transfer_orders,
    parse_weichuan_transfer_order, parse_weichuan_pdf,
    parse_vtl_transfer_order, parse_vtl_pdf,
    parse_vtl_with_claude,
)
from form_template import build_workbook
from form_template_weichuan import build_weichuan_workbook
from form_template_vtl import build_vtl_workbook

app = Flask(__name__)
CORS(app, origins=["https://markcqhsu.github.io"])


@app.get("/health")
def health():
    return jsonify({"status": "ok"})


@app.post("/ocr")
def ocr():
    """
    Accepts multipart/form-data with field 'image'.
    Returns parsed order data only — session state is managed client-side.
    """
    if "image" not in request.files:
        return jsonify({"error": "no image field"}), 400

    file = request.files["image"]
    if file.filename == "":
        return jsonify({"error": "empty filename"}), 400

    suffix = os.path.splitext(file.filename)[1] or ".jpg"
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        file.save(tmp.name)
        tmp_path = tmp.name

    try:
        parsed = parse_transfer_order(tmp_path)
    except Exception:
        return jsonify({"error": "OCR 處理失敗，請確認圖片格式"}), 500
    finally:
        os.unlink(tmp_path)

    return jsonify({
        "date":            parsed["date"],
        "warehouse":       parsed["warehouse"],
        "shipment_no":     parsed["shipment_no"],
        "matched_items":   parsed["matched_items"],
        "unmatched_items": parsed["unmatched_items"],
    })


@app.post("/ocr-pdf")
def ocr_pdf():
    """
    Accepts multipart/form-data with field 'file' (PDF).
    Returns parsed order data for every page of the PDF.
    """
    if "file" not in request.files:
        return jsonify({"error": "no file field"}), 400

    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "empty filename"}), 400

    if not file.filename.lower().endswith(".pdf"):
        return jsonify({"error": "file must be a PDF"}), 400

    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
        file.save(tmp.name)
        tmp_path = tmp.name

    try:
        orders = parse_pdf_transfer_orders(tmp_path, filename=file.filename)
    except Exception:
        return jsonify({"error": "PDF 處理失敗，請確認檔案格式"}), 500
    finally:
        os.unlink(tmp_path)

    return jsonify({"orders": orders})


@app.post("/export")
def export():
    """
    Builds an Excel file from the session rows sent by the client.
    Body JSON: {"rows": [{date, warehouse, items:[{code,qty}]}, ...]}
    """
    data = request.get_json(force=True, silent=True) or {}
    rows = data.get("rows", [])

    if not rows:
        return jsonify({"error": "no data"}), 400

    try:
        xlsx_bytes = build_workbook(rows)
    except Exception:
        return jsonify({"error": "Excel 產生失敗"}), 500

    return send_file(
        __import__("io").BytesIO(xlsx_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="宏全三廠_成品庫存.xlsx",
    )


@app.post("/ocr-weichuan")
def ocr_weichuan():
    """
    Parse 味全 出貨單 — accepts image (field 'image') or PDF (field 'file').
    Returns: {date, warehouse, matched_items, unmatched_items}
    """
    is_pdf = "file" in request.files
    file   = request.files.get("file") or request.files.get("image")
    if file is None:
        return jsonify({"error": "no file field"}), 400
    if file.filename == "":
        return jsonify({"error": "empty filename"}), 400

    suffix = ".pdf" if is_pdf else (os.path.splitext(file.filename)[1] or ".jpg")
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        file.save(tmp.name)
        tmp_path = tmp.name

    try:
        if is_pdf:
            parsed = parse_weichuan_pdf(tmp_path)
        else:
            parsed = parse_weichuan_transfer_order(tmp_path)
    except Exception:
        return jsonify({"error": "OCR 處理失敗，請確認圖片格式"}), 500
    finally:
        os.unlink(tmp_path)

    return jsonify(parsed)


@app.post("/export-weichuan")
def export_weichuan():
    """
    Build 味全代工明細表 Excel.
    Body JSON: {"rows": [{date, warehouse, items:[{code,qty}]}, ...]}
    """
    data = request.get_json(force=True, silent=True) or {}
    rows = data.get("rows", [])

    if not rows:
        return jsonify({"error": "no data"}), 400

    try:
        xlsx_bytes = build_weichuan_workbook(rows)
    except Exception:
        return jsonify({"error": "Excel 產生失敗"}), 500

    return send_file(
        __import__("io").BytesIO(xlsx_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="味全代工明細表.xlsx",
    )


@app.post("/ocr-vtl")
def ocr_vtl():
    """
    Parse 維他露 調撥單 — accepts image (field 'image') or PDF (field 'file').
    Returns: {orders: [{date, warehouse, matched_items, unmatched_items}, ...]}
    Each order = one 營業部 block = one Excel row.
    """
    is_pdf = "file" in request.files
    file   = request.files.get("file") or request.files.get("image")
    if file is None:
        return jsonify({"error": "no file field"}), 400
    if file.filename == "":
        return jsonify({"error": "empty filename"}), 400

    suffix = ".pdf" if is_pdf else (os.path.splitext(file.filename)[1] or ".jpg")
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        file.save(tmp.name)
        tmp_path = tmp.name

    try:
        if is_pdf:
            orders = parse_vtl_pdf(tmp_path)
        else:
            try:
                orders = parse_vtl_with_claude(tmp_path)
            except Exception:
                orders = parse_vtl_transfer_order(tmp_path)
    except Exception:
        return jsonify({"error": "OCR 處理失敗，請確認圖片格式"}), 500
    finally:
        os.unlink(tmp_path)

    return jsonify({"orders": orders})


@app.post("/export-vtl")
def export_vtl():
    """
    Build 維他露代工明細表 Excel.
    Body JSON: {"rows": [{date, warehouse, items:[{code,qty}]}, ...]}
    """
    data = request.get_json(force=True, silent=True) or {}
    rows = data.get("rows", [])

    if not rows:
        return jsonify({"error": "no data"}), 400

    try:
        xlsx_bytes = build_vtl_workbook(rows)
    except Exception:
        return jsonify({"error": "Excel 產生失敗"}), 500

    return send_file(
        __import__("io").BytesIO(xlsx_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="維他露代工明細表.xlsx",
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080, debug=False)
