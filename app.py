import os
import tempfile
from flask import Flask, request, send_file, render_template_string

from dd1750_core import generate_dd1750_from_pdf

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024  # 200MB

INDEX_HTML = """
<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <title>BOM → DD1750 (Flask v4 stable)</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 24px; }
    .row { display: flex; gap: 48px; }
    .col { flex: 1; }
    .card { border: 1px solid #ddd; padding: 16px; border-radius: 8px; }
    label { font-weight: bold; display:block; margin-top: 10px; }
    input[type=file], input[type=number] { width: 100%; padding: 6px; }
    button { padding: 10px 14px; margin-top: 16px; }
    .err { color: #b00020; font-weight: bold; margin-top: 12px; }
    .hint { color: #555; font-size: 0.95em; }
  </style>
</head>
<body>
  <h1>BOM → DD1750 (Flask v4 stable)</h1>
  <p class="hint">Upload a BOM PDF and a flat DD1750 template PDF (your blank form). Then click Generate.</p>

  <form method="post" action="/generate" enctype="multipart/form-data">
    <div class="row">
      <div class="col card">
        <label>BOM (PDF)</label>
        <input type="file" name="bom" accept="application/pdf" required />

        <label>Start parsing at page (0-based)</label>
        <input type="number" name="start_page" value="0" min="0" step="1" />
        <div class="hint">Use 0 for normal BOMs. If your BOM has a cover sheet, try 1.</div>
      </div>

      <div class="col card">
        <label>DD1750 template (flat PDF)</label>
        <input type="file" name="template" accept="application/pdf" required />
        <div class="hint">Use your blank flat form (no fillable fields). If you only have a fillable PDF, print-to-PDF first.</div>
      </div>
    </div>

    <button type="submit">Generate DD1750</button>

    {% if error %}
      <div class="err">{{ error }}</div>
    {% endif %}
  </form>
</body>
</html>
"""


@app.get("/")
def index():
    return render_template_string(INDEX_HTML, error=None)


@app.post("/generate")
def generate():
    bom = request.files.get("bom")
    template = request.files.get("template")

    if not bom or not template:
        return render_template_string(INDEX_HTML, error="Missing BOM or template file.")

    try:
        start_page = int(request.form.get("start_page", "0"))
    except ValueError:
        start_page = 0

    with tempfile.TemporaryDirectory() as td:
        bom_path = os.path.join(td, "bom.pdf")
        tpl_path = os.path.join(td, "template.pdf")
        out_path = os.path.join(td, "DD1750_OUTPUT.pdf")

        bom.save(bom_path)
        template.save(tpl_path)

        try:
            out_pdf, item_count = generate_dd1750_from_pdf(
                bom_pdf_path=bom_path,
                template_pdf_path=tpl_path,
                out_pdf_path=out_path,
                start_page=start_page,
            )
        except Exception as e:
            # Show a friendly error instead of crashing the container
            return render_template_string(INDEX_HTML, error=f"Error generating DD1750: {e}")

        if item_count == 0:
            return render_template_string(
                INDEX_HTML,
                error=(
                    "No readable items parsed from BOM. If this BOM is scanned, convert it to a text PDF "
                    "(OCR/searchable) or export to Excel first."
                ),
            )

        return send_file(out_pdf, as_attachment=True, download_name="DD1750_OUTPUT.pdf")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8080"))
    app.run(host="0.0.0.0", port=port)
