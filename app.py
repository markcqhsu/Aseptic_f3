import os
import tempfile
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS

from ocr_parser import parse_transfer_order
from form_template import build_workbook

app = Flask(__name__)
CORS(app)

# In-memory session: list of parsed order dicts
# Each entry: {"date": str, "warehouse": str, "items": [...]}
_session: list = []


@app.get("/health")
def health():
    return jsonify({"status": "ok"})


@app.post("/ocr")
def ocr():
    """
    Accepts a multipart/form-data upload with field 'image'.
    Returns parsed order data and appends it to the session.
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
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        os.unlink(tmp_path)

    entry = {
        "date":      parsed["date"],
        "warehouse": parsed["warehouse"],
        "items":     parsed["matched_items"],
    }
    _session.append(entry)

    return jsonify({
        "index":           len(_session) - 1,
        "date":            parsed["date"],
        "warehouse":       parsed["warehouse"],
        "shipment_no":     parsed["shipment_no"],
        "matched_items":   parsed["matched_items"],
        "unmatched_items": parsed["unmatched_items"],
        "session_count":   len(_session),
    })


@app.put("/session/<int:index>")
def update_session(index: int):
    """
    Allows the frontend to push corrected data back for a given order index.
    Body JSON: {"date": "...", "warehouse": "...", "items": [...]}
    """
    if index < 0 or index >= len(_session):
        return jsonify({"error": "index out of range"}), 404

    data = request.get_json(force=True, silent=True)
    if not data:
        return jsonify({"error": "invalid JSON"}), 400

    _session[index] = {
        "date":      data.get("date",      _session[index]["date"]),
        "warehouse": data.get("warehouse", _session[index]["warehouse"]),
        "items":     data.get("items",     _session[index]["items"]),
    }
    return jsonify({"ok": True})


@app.delete("/session/<int:index>")
def delete_session(index: int):
    if index < 0 or index >= len(_session):
        return jsonify({"error": "index out of range"}), 404
    _session.pop(index)
    return jsonify({"ok": True, "session_count": len(_session)})


@app.delete("/session")
def clear_session():
    _session.clear()
    return jsonify({"ok": True})


@app.get("/session")
def get_session():
    return jsonify({"rows": _session, "session_count": len(_session)})


@app.post("/export")
def export():
    """
    Builds an Excel file from the current session and returns it for download.
    Optionally accepts JSON body with overridden rows to allow one-shot export
    without a prior /ocr call.
    """
    data = request.get_json(force=True, silent=True) or {}
    rows = data.get("rows", _session)

    if not rows:
        return jsonify({"error": "no data in session"}), 400

    try:
        xlsx_bytes = build_workbook(rows)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    return send_file(
        __import__("io").BytesIO(xlsx_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="宏全三廠_成品庫存.xlsx",
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080, debug=False)
