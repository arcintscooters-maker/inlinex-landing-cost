import os
import base64
import json
from flask import Flask, request, jsonify, render_template
import anthropic

app = Flask(__name__)

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/parse", methods=["POST"])
def parse():
    try:
        api_key = os.environ.get("ANTHROPIC_API_KEY", "").strip()
        invoice_sgd = float(request.form.get("invoice_sgd", 0))
        shipping_sgd = float(request.form.get("shipping_sgd", 0))
        file = request.files.get("invoice")

        if not api_key:
            return jsonify({"error": "API key not configured on server. Add ANTHROPIC_API_KEY in Railway Variables."}), 500
        if not file:
            return jsonify({"error": "No file uploaded"}), 400
        if invoice_sgd <= 0:
            return jsonify({"error": "Invoice SGD amount must be > 0"}), 400

        filename = file.filename.lower()
        file_bytes = file.read()
        b64 = base64.standard_b64encode(file_bytes).decode("utf-8")

        if filename.endswith(".pdf"):
            media_type = "application/pdf"
        elif filename.endswith((".xlsx", ".xls")):
            media_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        else:
            return jsonify({"error": "Unsupported file type. Upload PDF or Excel."}), 400

        prompt = """This is a supplier invoice. Extract ALL line items and return ONLY a valid JSON object — no markdown, no explanation, no extra text.

JSON structure:
{
  "invoice_no": "string",
  "supplier": "string",
  "invoice_total_usd": number,
  "items": [
    {
      "pos": number,
      "sku": "string",
      "ean": "string",
      "description": "string",
      "brand": "string",
      "qty": number,
      "unit_usd": number,
      "total_usd": number
    }
  ],
  "notes": "string - warnings like out of stock, substitutions. Empty string if none."
}

Rules:
- sku = article/item number from invoice
- ean = barcode if present, else empty string
- brand = sub-brand or product line (e.g. Ennui, IQON, Powerslide, Kizer, Flying Eagle)
- total_usd = actual line total after any discount
- invoice_total_usd = grand total of all lines
- Return ONLY the JSON object, nothing else"""

        client = anthropic.Anthropic(api_key=api_key)
        message = client.messages.create(
            model="claude-opus-4-5",
            max_tokens=4000,
            messages=[{
                "role": "user",
                "content": [
                    {
                        "type": "document",
                        "source": {
                            "type": "base64",
                            "media_type": media_type,
                            "data": b64
                        }
                    },
                    {"type": "text", "text": prompt}
                ]
            }]
        )

        raw = message.content[0].text.strip()
        clean = raw.replace("```json", "").replace("```", "").strip()
        parsed = json.loads(clean)

        inv_usd = parsed["invoice_total_usd"]

        for item in parsed["items"]:
            pct = item["total_usd"] / inv_usd
            line_sgd = pct * invoice_sgd
            ship_alloc = pct * shipping_sgd
            item["pct"] = round(pct * 100, 4)
            item["line_sgd"] = round(line_sgd, 2)
            item["ship_alloc"] = round(ship_alloc, 2)
            item["landed_per_unit"] = round((line_sgd + ship_alloc) / item["qty"], 2)

        parsed["invoice_sgd"] = invoice_sgd
        parsed["shipping_sgd"] = shipping_sgd
        parsed["total_landed"] = round(invoice_sgd + shipping_sgd, 2)

        return jsonify(parsed)

    except json.JSONDecodeError:
        return jsonify({"error": "Could not parse AI response as JSON. Try again."}), 500
    except anthropic.AuthenticationError:
        return jsonify({"error": "Invalid API key on server. Check Railway Variables."}), 401
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"Starting InlineX Landing Cost Calculator on port {port}...")
    app.run(debug=False, host="0.0.0.0", port=port)
