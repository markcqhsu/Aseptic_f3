import hmac
import io
import logging
import os
import tempfile
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address

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

logger = logging.getLogger("aseptic-f3-api")

limiter = Limiter(get_remote_address, app=app, default_limits=[])

# Shared secret required on every request (except /health). Set via env/Secret
# Manager in deploy.sh. If unset, the check is skipped (local dev convenience).
APP_SHARED_KEY = os.environ.get("APP_SHARED_KEY")


@app.before_request
def _require_app_key():
    if not APP_SHARED_KEY or request.method == "OPTIONS" or request.path == "/health":
        return None
    supplied = request.headers.get("X-App-Key", "")
    if not hmac.compare_digest(supplied, APP_SHARED_KEY):
        return jsonify({"error": "unauthorized"}), 401
    return None


# Limit upload size to 20 MB
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024

_ALLOWED_IMAGE_EXT = {".jpg", ".jpeg", ".png", ".webp", ".gif", ".bmp", ".tiff", ".tif"}


def _image_suffix(filename: str) -> str:
    ext = os.path.splitext(filename)[1].lower()
    return ext if ext in _ALLOWED_IMAGE_EXT else ".jpg"


@app.get("/health")
def health():
    return jsonify({"status": "ok"})


@app.post("/ocr")
@limiter.limit("30/hour")
def ocr():
    if "image" not in request.files:
        return jsonify({"error": "no image field"}), 400

    file = request.files["image"]
    if file.filename == "":
        return jsonify({"error": "empty filename"}), 400

    suffix = _image_suffix(file.filename)
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
@limiter.limit("30/hour")
def ocr_pdf():
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
@limiter.limit("60/hour")
def export():
    data = request.get_json(force=True, silent=True) or {}
    rows = data.get("rows", [])

    if not rows:
        return jsonify({"error": "no data"}), 400

    try:
        xlsx_bytes = build_workbook(rows)
    except Exception:
        return jsonify({"error": "Excel 產生失敗"}), 500

    return send_file(
        io.BytesIO(xlsx_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="宏全三廠_成品庫存.xlsx",
    )


@app.post("/ocr-weichuan")
@limiter.limit("30/hour")
def ocr_weichuan():
    is_pdf = "file" in request.files
    file   = request.files.get("file") or request.files.get("image")
    if file is None:
        return jsonify({"error": "no file field"}), 400
    if file.filename == "":
        return jsonify({"error": "empty filename"}), 400

    if is_pdf:
        if not file.filename.lower().endswith(".pdf"):
            return jsonify({"error": "file must be a PDF"}), 400
        suffix = ".pdf"
    else:
        suffix = _image_suffix(file.filename)

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
@limiter.limit("60/hour")
def export_weichuan():
    data = request.get_json(force=True, silent=True) or {}
    rows = data.get("rows", [])

    if not rows:
        return jsonify({"error": "no data"}), 400

    try:
        xlsx_bytes = build_weichuan_workbook(rows)
    except Exception:
        return jsonify({"error": "Excel 產生失敗"}), 500

    return send_file(
        io.BytesIO(xlsx_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="味全代工明細表.xlsx",
    )


@app.post("/ocr-vtl")
@limiter.limit("30/hour")
def ocr_vtl():
    is_pdf = "file" in request.files
    file   = request.files.get("file") or request.files.get("image")
    if file is None:
        return jsonify({"error": "no file field"}), 400
    if file.filename == "":
        return jsonify({"error": "empty filename"}), 400

    if is_pdf:
        if not file.filename.lower().endswith(".pdf"):
            return jsonify({"error": "file must be a PDF"}), 400
        suffix = ".pdf"
    else:
        suffix = _image_suffix(file.filename)

    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        file.save(tmp.name)
        tmp_path = tmp.name

    method = "unknown"
    debug_info = []
    total_items = None
    try:
        if is_pdf:
            method = "claude_vision_pdf"
            result = parse_vtl_pdf(tmp_path)
            orders = result["orders"]
            total_items = result.get("total_items")
        else:
            try:
                method = "claude_vision"
                orders = parse_vtl_with_claude(tmp_path)
                debug_info.append("Claude Vision 成功")
            except Exception as e:
                logger.warning("Claude Vision failed, falling back: %s: %s", type(e).__name__, e)
                debug_info.append(f"Claude Vision 失敗：{type(e).__name__}")
                method = "google_vision_fallback"
                try:
                    orders = parse_vtl_transfer_order(tmp_path)
                    debug_info.append("Google Vision fallback 成功")
                except Exception as e2:
                    logger.warning("Google Vision fallback also failed: %s: %s", type(e2).__name__, e2)
                    debug_info.append(f"Google Vision 也失敗：{type(e2).__name__}")
                    raise
    except Exception as e:
        logger.exception("VTL OCR failed (method=%s, debug=%s)", method, debug_info)
        return jsonify({"error": "OCR 處理失敗，請確認圖片或檔案格式"}), 500
    finally:
        os.unlink(tmp_path)

    return jsonify({"orders": orders, "total_items": total_items, "method": method, "debug": debug_info})


@app.post("/export-vtl")
@limiter.limit("60/hour")
def export_vtl():
    data = request.get_json(force=True, silent=True) or {}
    rows = data.get("rows", [])

    if not rows:
        return jsonify({"error": "no data"}), 400

    try:
        xlsx_bytes = build_vtl_workbook(rows)
    except Exception:
        return jsonify({"error": "Excel 產生失敗"}), 500

    return send_file(
        io.BytesIO(xlsx_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="維他露代工明細表.xlsx",
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080, debug=False)
