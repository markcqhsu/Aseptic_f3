import os
import tempfile
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS

from ocr_parser import parse_transfer_order, parse_pdf_transfer_orders
from form_template import build_workbook

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


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080, debug=False)
