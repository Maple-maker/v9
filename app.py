import os, tempfile, re
from flask import Flask, request, send_file, render_template_string
from pypdf import PdfReader, PdfWriter
from pypdf._page import PageObject
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Paragraph

VERSION = "2025-01-02.flask.v5"

HTML = """
<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>BOM → DD1750</title>
  <style>
    body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif;max-width:980px;margin:40px auto;padding:0 16px}
    .box{border:1px solid #ddd;border-radius:10px;padding:18px;margin:16px 0}
    label{display:block;margin:10px 0 4px}
    input,select{padding:8px;width:100%}
    button{padding:10px 14px;border:0;border-radius:8px;background:#111;color:#fff;cursor:pointer}
    small{color:#555}
    .row{display:flex;gap:18px}
    .col{flex:1}
  </style>
</head>
<body>
  <h2>BOM → DD1750 Automator</h2>
  <p><small>Version: {{version}}</small></p>

  <div class="box">
    <form method="post" action="/generate" enctype="multipart/form-data">
      <div class="row">
        <div class="col">
          <label>BOM (PDF)</label>
          <input type="file" name="bom" accept=".pdf" required>
        </div>
        <div class="col">
          <label>Blank DD1750 template (PDF)</label>
          <input type="file" name="template" accept=".pdf" required>
        </div>
      </div>
      <div class="row">
        <div class="col">
          <label>Label under description</label>
          <select name="label">
            <option value="NSN" selected>NSN</option>
            <option value="SN">SN</option>
          </select>
        </div>
        <div class="col">
          <label>Start Page (0-based)</label>
          <input type="number" name="start_page" value="0" min="0">
        </div>
      </div>
      <p><small>Ensure your BOM is OCR-enabled before uploading.</small></p>
      <button type="submit">Generate DD1750</button>
    </form>
  </div>
</body>
</html>
"""

app = Flask(__name__)

# --- BOM Parsing Logic ---
MAT_RE = re.compile(r"^\s*(\d{7,13})") # Matches NSN style strings [cite: 7, 20]
QTY_RE = re.compile(r"(\d+)\s*$")

def is_header_noise(s: str) -> bool:
    u = s.upper()
    return any(h in u for h in ["LV", "DESCRIPTION", "WTY", "ARC", "CIIC", "UI", "SCMC", "AUTH", "OH QTY", "COMPONENT", "PAGE"])

def extract_items_bom_style(pdf_path: str, start_page: int = 0):
    reader = PdfReader(pdf_path)
    items = []
    cur = {"desc": None, "mat": None, "qty": None}

    def flush():
        nonlocal cur
        if cur["desc"] and cur["mat"] and cur["qty"] is not None and cur["qty"] > 0:
            items.append({"desc": cur["desc"], "mat": cur["mat"], "qty": cur["qty"]})
        cur = {"desc": None, "mat": None, "qty": None}

    for pi in range(start_page, len(reader.pages)):
        txt = reader.pages[pi].extract_text() or ""
        lines = [l.strip() for l in txt.splitlines() if l.strip()]
        for s in lines:
            if is_header_noise(s): continue
            
            # Match Material Numbers (NSNs) [cite: 7, 20]
            mm = MAT_RE.match(s)
            if mm:
                if cur["mat"]: flush()
                cur["mat"] = mm.group(1)
                continue

            # Identify quantities (last digit on lines with codes like EA or AY) [cite: 7, 20]
            if any(tok in s.upper() for tok in ["EA", "AY", "9G", "9K", "SCMC"]):
                qm = QTY_RE.search(s)
                if qm: cur["qty"] = int(qm.group(1))
                continue

            # Capture Description [cite: 7, 20]
            if not cur["mat"] and len(s) > 5:
                cur["desc"] = s if not cur["desc"] else cur["desc"] + " " + s

    flush()
    return items

# --- DD1750 Overlay with Text Wrapping ---
def make_overlay(pages, label="NSN"):
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    c = canvas.Canvas(tmp.name, pagesize=letter)
    styles = getSampleStyleSheet()
    
    # Custom style for wrapped description
    desc_style = styles["Normal"]
    desc_style.fontSize = 8
    desc_style.leading = 9 

    # Field coordinates (Inches) [cite: 119, 120, 121, 122, 131, 132, 136]
    col_box = 0.85 * inch
    col_contents = 1.62 * inch
    col_uoi = 6.18 * inch
    col_init = 7.05 * inch
    col_run = 7.83 * inch
    col_total = 8.58 * inch

    y_start = 6.55 * inch
    row_height = 0.305 * inch 

    for rows in pages:
        y = y_start
        for idx, it in enumerate(rows):
            # Box No [cite: 119]
            c.setFont("Helvetica", 9)
            c.drawCentredString(col_box, y - 0.15*inch, str(idx + 1))

            # Contents (Wrapped) [cite: 120, 121]
            content_html = f"<b>{it['desc']}</b><br/>{label}: {it['mat']}"
            p = Paragraph(content_html, desc_style)
            w, h = p.wrap(4.2 * inch, row_height)
            p.drawOn(c, col_contents, y - h - 0.05*inch)

            # UOI [cite: 122]
            c.drawCentredString(col_uoi, y - 0.15*inch, "EA")

            # Quantities [cite: 131, 132, 133, 136]
            q_str = str(it["qty"])
            c.drawCentredString(col_init, y - 0.15*inch, q_str)
            c.drawCentredString(col_run, y - 0.15*inch, "0")
            c.drawCentredString(col_total, y - 0.15*inch, q_str)

            y -= row_height
        c.showPage()
    c.save()
    return tmp.name

def merge_with_template(template_pdf, overlay_pdf, out_pdf):
    tpl = PdfReader(template_pdf)
    ov = PdfReader(overlay_pdf)
    writer = PdfWriter()
    for ovp in ov.pages:
        merged = PageObject.create_blank_page(width=tpl.pages[0].mediabox.width, height=tpl.pages[0].mediabox.height)
        merged.merge_page(tpl.pages[0])
        merged.merge_page(ovp)
        writer.add_page(merged)
    with open(out_pdf, "wb") as f: writer.write(f)

@app.route("/")
def index(): return render_template_string(HTML, version=VERSION)

@app.post("/generate")
def generate():
    bom = request.files["bom"]
    template = request.files["template"]
    label = request.form.get("label", "NSN")
    start_page = int(request.form.get("start_page", 0))

    with tempfile.TemporaryDirectory() as td:
        bp, tp = os.path.join(td, "b.pdf"), os.path.join(td, "t.pdf")
        bom.save(bp); template.save(tp)
        
        items = extract_items_bom_style(bp, start_page)
        chunked = [items[i:i+18] for i in range(0, len(items), 18)] or [[]]
        
        overlay = make_overlay(chunked, label)
        out = os.path.join(td, "DD1750.pdf")
        merge_with_template(tp, overlay, out)
        return send_file(out, as_attachment=True, download_name="DD1750_Automated.pdf")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8080)))
