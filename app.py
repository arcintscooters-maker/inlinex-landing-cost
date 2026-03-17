import os
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

        file_bytes = file.read()
        result = parse_invoice(file.filename, file_bytes)

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

    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        return jsonify({"error": f"Failed to parse invoice: {str(e)}"}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"Starting on port {port}...")
    app.run(debug=False, host="0.0.0.0", port=port)
