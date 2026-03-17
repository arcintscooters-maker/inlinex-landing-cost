import os
import traceback
from flask import Flask, request, jsonify, render_template
from invoice_parser import parse_invoice

app = Flask(__name__)

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/parse", methods=["POST"])
def parse():
    try:
        invoice_sgd = float(request.form.get("invoice_sgd", 0))
        shipping_sgd = float(request.form.get("shipping_sgd", 0))
        file = request.files.get("invoice")

        if not file:
            return jsonify({"error": "No file uploaded"}), 400
        if invoice_sgd <= 0:
            return jsonify({"error": "Invoice SGD amount must be > 0"}), 400

        filename = file.filename
        file_bytes = file.read()

        print(f"[DEBUG] filename={filename} size={len(file_bytes)} bytes")

        result = parse_invoice(filename, file_bytes)

        print(f"[DEBUG] parsed ok: {len(result['items'])} items, total={result['invoice_total_usd']}")

        inv_usd = result["invoice_total_usd"]
        for item in result["items"]:
            pct = item["total_usd"] / inv_usd
            line_sgd = pct * invoice_sgd
            ship_alloc = pct * shipping_sgd
            item["pct"] = round(pct * 100, 4)
            item["line_sgd"] = round(line_sgd, 2)
            item["ship_alloc"] = round(ship_alloc, 2)
            item["landed_per_unit"] = round((line_sgd + ship_alloc) / item["qty"], 2)

        result["invoice_sgd"] = invoice_sgd
        result["shipping_sgd"] = shipping_sgd
        result["total_landed"] = round(invoice_sgd + shipping_sgd, 2)

        return jsonify(result)

    except Exception as e:
        tb = traceback.format_exc()
        print(f"[ERROR] {tb}")
        return jsonify({"error": str(e), "detail": tb}), 500


@app.route("/debug", methods=["GET"])
def debug():
    """Simple endpoint to verify the app and imports are working."""
    try:
        import pdfplumber
        import openpyxl
        return jsonify({
            "status": "ok",
            "pdfplumber": pdfplumber.__version__,
            "openpyxl": openpyxl.__version__,
        })
    except Exception as e:
        return jsonify({"status": "error", "error": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"Starting on port {port}...")
    app.run(debug=False, host="0.0.0.0", port=port)
