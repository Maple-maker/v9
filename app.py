import os
import re
import io
import tempfile
from typing import List, Dict, Tuple

from flask import Flask, request, send_file, render_template_string
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.pdfbase.pdfmetrics import stringWidth

app = Flask(__name__)

HTML = """
<!doctype html>
<title>BOM → DD1750 (v4 items patch)</title>
<h1>BOM → DD1750 (v4 items patch)</h1>
<form method=post enctype=multipart/form-data>
  <p><b>BOM (PDF)</b><br><input type=file name=bom_pdf required></p>
  <p><b>DD1750 template (flat PDF)</b><br><input type=file name=template_pdf required>
     <br><small>Upload your blank_flat.pdf template.</small>
  </p>
  <p><b>Start parsing at page (0-based)</b><br>
     <input type=number name=start_page value="0" min="0" style="width:80px;">
     <small>Use 0 for normal BOMs.</small>
  </p>
  <p><button type=submit>Generate DD1750</button></p>
</form>
{% if error %}<p style="color:#b00;"><b>{{error}}</b></p>{% endif %}
"""

# ----- Parsing -----

NSN_LINE_RE = re.compile(r"^\d{9}$")

def _clean_line(s: str) -> str:
    s = s.replace("\u00a0", " ")
    return re.sub(r"\s+", " ", s).strip()

def _extract_qty_from_line(line: str) -> int:
    # quantity is usually the last integer token
    m = re.search(r"(\d+)\s*$", line)
    if not m:
        return 0
    try:
        q = int(m.group(1))
    except ValueError:
        return 0
    # guard against garbage huge numbers
    if q < 0 or q > 99999:
        return 0
    return q

def _pick_description(block_lines: List[str]) -> str:
    # Prefer the line right after a single-letter line (A/B/C...), typical in these BOMs.
    for i, ln in enumerate(block_lines[:-1]):
        if re.fullmatch(r"[A-Z]", ln):
            cand = block_lines[i+1]
            if cand and not NSN_LINE_RE.match(cand) and "EA" not in cand and not cand.startswith("C_"):
                return cand
    # Otherwise prefer first reasonable line that isn't codes or generic labels
    for ln in block_lines:
        if not ln or NSN_LINE_RE.match(ln):
            continue
        if ln.startswith("C_") or "~" in ln:
            continue
        # skip generic placeholders
        if ln.upper() in {"ITEM", "NSN", "NSN:", "NOMENCLATURE", "DESCRIPTION"}:
            continue
        if "EA" in ln:
            continue
        if len(ln) < 3:
            continue
        return ln
    return ""

def parse_bom_pdf(bom_path: str, start_page: int = 0) -> List[Dict[str, str]]:
    reader = PdfReader(bom_path)
    items: List[Dict[str, str]] = []

    current_nsn = None
    block: List[str] = []

    def flush_block():
        nonlocal current_nsn, block, items
        if not current_nsn:
            block = []
            return
        qty = 0
        # find the last line in the block containing EA and a trailing qty
        for ln in reversed(block):
            if "EA" in ln:
                q = _extract_qty_from_line(ln)
                if q:
                    qty = q
                    break
        if qty <= 0:
            block = []
            current_nsn = None
            return
        desc = _pick_description(block)
        # If we only found a generic placeholder (common in these BOM exports),
        # fall back to the line that contains "EA ... <qty>" and extract the text before EA.
        if not desc or desc.upper() in {"ITEM", "NSN", "DESCRIPTION"}:
            for ln in block:
                if EA_QTY_RE.search(ln):
                    left = re.split(r"\bEA\b", ln, maxsplit=1)[0]
                    # Trim common trailing code patterns (e.g., "X U", "X J", etc.)
                    left = re.sub(r"\bX\s+[A-Z]{1,2}\s*$", "", left).strip()
                    # Remove duplicated commas/spaces
                    left = re.sub(r"\s{2,}", " ", left).strip(" ,-")
                    if left:
                        desc = left
                        break
        desc = desc or "ITEM"
        items.append({
            "nsn": current_nsn,
            "desc": desc,
            "qty": str(qty),
        })
        block = []
        current_nsn = None

    for p in range(start_page, len(reader.pages)):
        text = reader.pages[p].extract_text() or ""
        for raw in text.splitlines():
            ln = _clean_line(raw)
            if not ln:
                continue
            if NSN_LINE_RE.match(ln):
                # new item starts; flush previous
                flush_block()
                current_nsn = ln
                block = [ln]
                continue
            if current_nsn:
                block.append(ln)

    flush_block()
    return items

# ----- Rendering -----

def wrap_to_width(text: str, max_width: float, font_name: str, font_size: float, max_lines: int) -> List[str]:
    # help wrapping by adding spaces after commas/slashes
    t = re.sub(r"([,/])", r"\1 ", text)
    words = t.split()
    if not words:
        return [""]

    lines: List[str] = []
    cur = ""
    for w in words:
        cand = (cur + " " + w).strip() if cur else w
        if stringWidth(cand, font_name, font_size) <= max_width:
            cur = cand
        else:
            if cur:
                lines.append(cur)
            # if single word too long, hard-break
            if stringWidth(w, font_name, font_size) > max_width:
                tmp = ""
                for ch in w:
                    cand2 = tmp + ch
                    if stringWidth(cand2, font_name, font_size) <= max_width:
                        tmp = cand2
                    else:
                        if tmp:
                            lines.append(tmp)
                        tmp = ch
                cur = tmp
            else:
                cur = w
        if len(lines) >= max_lines:
            break

    if len(lines) < max_lines and cur:
        lines.append(cur)

    # truncate if too many
    if len(lines) > max_lines:
        lines = lines[:max_lines]

    # add ellipsis if we truncated mid-text
    joined = " ".join(lines)
    if len(joined) < len(t):
        # ensure last line ends with … if possible
        last = lines[-1]
        ell = "…"
        while last and stringWidth(last + ell, font_name, font_size) > max_width:
            last = last[:-1]
        lines[-1] = (last + ell) if last else ell

    return lines


def build_overlay_page(c: canvas.Canvas, rows: List[Dict[str, str]], page_num: int, total_pages: int):
    # Page is letter-size template; coordinates tuned to the provided blank_flat.pdf
    # Row geometry
    rows_per_page = 18
    row_h = 22.0
    top_y = 600.0  # baseline for row 1 (numbers look good here)

    # Column x positions
    x_box = 60.0
    x_desc = 110.0
    desc_w = 270.0
    x_uoi = 410.0
    x_init = 455.0
    x_spares = 500.0
    x_total = 548.0

    # Fonts
    font = "Helvetica"

    for i, it in enumerate(rows):
        row_idx = i + 1
        y = top_y - (row_idx - 1) * row_h

        # Box no.
        c.setFont(font, 8)
        c.drawRightString(x_box, y, str((page_num * rows_per_page) + row_idx))

        # Description + NSN inside contents column, top-aligned within the row
        desc = it["desc"]
        nsn = it["nsn"]

        # Fit desc to 2 lines, then a NSN line.
        # Try slightly larger then shrink if needed.
        for fs in (7.0, 6.5, 6.0, 5.5):
            lines = wrap_to_width(desc, desc_w, font, fs, max_lines=2)
            if len(lines) <= 2:
                desc_fs = fs
                break
        else:
            desc_fs = 5.5
            lines = wrap_to_width(desc, desc_w, font, desc_fs, max_lines=2)

        # top inset so it sits higher in the cell
        y_top = y + 7.0
        c.setFont(font, desc_fs)
        for li, ln in enumerate(lines):
            c.drawString(x_desc, y_top - li * (desc_fs + 1.0), ln)

        c.setFont(font, 6.0)
        c.drawString(x_desc, y_top - 2 * (desc_fs + 1.0), f"NSN: {nsn}")

        # UOI and quantities (center-ish)
        c.setFont(font, 8)
        c.drawCentredString(x_uoi, y, "EA")
        qty = str(int(it["qty"]))
        c.drawCentredString(x_init, y, qty)
        c.drawCentredString(x_spares, y, "0")
        c.drawCentredString(x_total, y, qty)


def render_dd1750(items: List[Dict[str, str]], template_path: str) -> bytes:
    if not items:
        raise ValueError("No items parsed from BOM. If this BOM is scanned, convert to a text PDF or provide an Excel export.")

    rows_per_page = 18
    pages = (len(items) + rows_per_page - 1) // rows_per_page

    # Build overlay PDF in memory
    overlay_buf = io.BytesIO()
    c = canvas.Canvas(overlay_buf, pagesize=template_page_size(template_path))

    for p in range(pages):
        start = p * rows_per_page
        end = min(len(items), start + rows_per_page)
        build_overlay_page(c, items[start:end], p, pages)
        c.showPage()
    c.save()
    overlay_buf.seek(0)

    # Merge overlays with template
    template_reader = PdfReader(template_path)
    overlay_reader = PdfReader(overlay_buf)

    out = PdfWriter()
    base_page = template_reader.pages[0]

    for i in range(pages):
        page = base_page
        # copy to avoid mutating base page for subsequent merges
        page = page.copy()
        page.merge_page(overlay_reader.pages[i])
        out.add_page(page)

    out_buf = io.BytesIO()
    out.write(out_buf)
    out_buf.seek(0)
    return out_buf.getvalue()


def template_page_size(template_path: str) -> Tuple[float, float]:
    r = PdfReader(template_path)
    mb = r.pages[0].mediabox
    return float(mb.width), float(mb.height)


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'GET':
        return render_template_string(HTML, error=None)

    bom_file = request.files.get('bom_pdf')
    tpl_file = request.files.get('template_pdf')
    start_page = int(request.form.get('start_page', '0') or 0)

    if not bom_file or not tpl_file:
        return render_template_string(HTML, error="Please upload both the BOM PDF and the template PDF.")

    with tempfile.TemporaryDirectory() as td:
        bom_path = os.path.join(td, 'bom.pdf')
        tpl_path = os.path.join(td, 'template.pdf')
        bom_file.save(bom_path)
        tpl_file.save(tpl_path)

        try:
            items = parse_bom_pdf(bom_path, start_page=start_page)
            pdf_bytes = render_dd1750(items, tpl_path)
        except Exception as e:
            return render_template_string(HTML, error=str(e))

    return send_file(
        io.BytesIO(pdf_bytes),
        mimetype='application/pdf',
        as_attachment=True,
        download_name='DD1750_OUTPUT.pdf'
    )


if __name__ == '__main__':
    port = int(os.environ.get('PORT', '5000'))
    app.run(host='0.0.0.0', port=port)
