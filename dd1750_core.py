"""DD1750 core: parse BOM PDFs and render DD Form 1750 overlays.

Designed for consistent 'Component Listing / Hand Receipt' style BOM PDFs.

- Parses line items that begin with 'B ' and end with a quantity.
- Captures the nearest subsequent 9-digit NSN line.
- Generates a multi-page DD1750 PDF by overlaying text onto a flat template.

This file is intentionally self-contained and safe to import on Railway.
"""

from __future__ import annotations

import io
import math
import os
import re
from dataclasses import dataclass
from copy import deepcopy
from typing import List, Optional, Tuple

import pdfplumber
from pypdf import PdfReader, PdfWriter

from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics


# --- Constants derived from the supplied blank template (letter: 612x792)
# These coordinates match the grid lines in the blank template render.
PAGE_W, PAGE_H = 612.0, 792.0

# Column x-bounds (points)
X_BOX_L, X_BOX_R = 44.0, 88.0
X_CONTENT_L, X_CONTENT_R = 88.0, 365.0
X_UOI_L, X_UOI_R = 365.0, 408.5
X_INIT_L, X_INIT_R = 408.5, 453.5
X_SPARES_L, X_SPARES_R = 453.5, 514.5
X_TOTAL_L, X_TOTAL_R = 514.5, 566.0

# Table y-bounds (points)
# Header divider line (under column titles) is ~616pt; bottom divider above certification ~89.5pt.
Y_TABLE_TOP_LINE = 616.0
Y_TABLE_BOTTOM_LINE = 89.5

# Standard DD1750 has 40 lines.
ROWS_PER_PAGE = 40

# Compute row height from table bounds
ROW_H = (Y_TABLE_TOP_LINE - Y_TABLE_BOTTOM_LINE) / ROWS_PER_PAGE

# Text padding inside cells
PAD_X = 3.0


@dataclass
class BomItem:
    line_no: int
    description: str
    nsn: str
    qty: int


_B_PREFIX_RE = re.compile(r"^B\s+(.+?)\s+(\d+)\s*$")
_NSN_RE = re.compile(r"^(\d{9})$")


def _clean_desc(desc: str) -> str:
    desc = desc.strip()
    desc = re.sub(r"\s+", " ", desc)
    # remove obvious control strings that show up in some exports
    desc = desc.replace("COMPONENT LISTING / HAND RECEIPT", "").strip()
    return desc


def extract_items_from_pdf(pdf_path: str, start_page: int = 0) -> List[BomItem]:
    """Extract line items from a text-based BOM PDF.

    Heuristic:
      - Item start line: begins with 'B ' and ends with an integer qty.
      - NSN: the next line within a short window that is exactly 9 digits.

    Returns a sequential list (line_no starts at 1).
    """

    items: List[Tuple[str, str, int]] = []

    with pdfplumber.open(pdf_path) as pdf:
        pages = pdf.pages[start_page:]
        for p in pages:
            text = p.extract_text() or ""
            lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
            i = 0
            while i < len(lines):
                ln = lines[i]
                m = _B_PREFIX_RE.match(ln)
                if m:
                    desc_raw, qty_s = m.group(1), m.group(2)
                    qty = int(qty_s)
                    desc = _clean_desc(desc_raw)

                    nsn = ""
                    # look ahead for nsn line
                    for j in range(i + 1, min(i + 10, len(lines))):
                        if _NSN_RE.match(lines[j]):
                            nsn = lines[j]
                            break

                    # Some exports break descriptions across multiple lines.
                    # If the next line isn't an NSN and doesn't start a new item, append it.
                    # Keep it conservative so we don't swallow the next item.
                    if i + 1 < len(lines):
                        nxt = lines[i + 1]
                        if (not _NSN_RE.match(nxt)) and (not _B_PREFIX_RE.match(nxt)):
                            # Don't append material/part lines (often start with C_ or digits/underscore combos)
                            if not re.match(r"^[A-Z]_", nxt):
                                # append only if it doesn't end with a qty that looks like a new item
                                if not re.search(r"\s\d+\s*$", nxt):
                                    desc = _clean_desc(desc + " " + nxt)

                    items.append((desc, nsn, qty))
                i += 1

    out: List[BomItem] = []
    for idx, (desc, nsn, qty) in enumerate(items, start=1):
        out.append(BomItem(line_no=idx, description=desc, nsn=nsn, qty=qty))
    return out


def _wrap_to_width(text: str, font: str, size: float, max_w: float, max_lines: int) -> List[str]:
    """Greedy word-wrap using actual font metrics."""
    text = re.sub(r"\s+", " ", text).strip()
    if not text:
        return [""]

    words = text.split(" ")
    lines: List[str] = []
    cur = ""

    def fits(s: str) -> bool:
        return pdfmetrics.stringWidth(s, font, size) <= max_w

    for w in words:
        if not cur:
            trial = w
        else:
            trial = cur + " " + w

        if fits(trial):
            cur = trial
            continue

        # if single word doesn't fit, hard-break it
        if not cur:
            chunk = w
            while chunk:
                # find longest prefix that fits
                lo, hi = 1, len(chunk)
                best = 1
                while lo <= hi:
                    mid = (lo + hi) // 2
                    cand = chunk[:mid]
                    if fits(cand):
                        best = mid
                        lo = mid + 1
                    else:
                        hi = mid - 1
                lines.append(chunk[:best])
                chunk = chunk[best:]
                if len(lines) >= max_lines:
                    return lines[:max_lines]
            cur = ""
        else:
            lines.append(cur)
            cur = w if fits(w) else w
            if len(lines) >= max_lines:
                return lines[:max_lines]

    if cur:
        lines.append(cur)

    if len(lines) > max_lines:
        lines = lines[:max_lines]

    return lines


def _draw_center(c: canvas.Canvas, txt: str, x_l: float, x_r: float, y: float, font: str, size: float):
    c.setFont(font, size)
    x = (x_l + x_r) / 2.0
    c.drawCentredString(x, y, txt)


def _build_overlay_page(items: List[BomItem], page_num: int, total_pages: int) -> bytes:
    """Return a PDF bytes for a single overlay page."""
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(PAGE_W, PAGE_H))

    # fonts
    FONT_MAIN = "Helvetica"
    FONT_SMALL = "Helvetica"

    # Start baseline for first row.
    # Place near top of first row area (slightly below the top divider line).
    first_row_top = Y_TABLE_TOP_LINE - 2.0

    max_content_w = (X_CONTENT_R - X_CONTENT_L) - 2 * PAD_X

    for row_idx in range(ROWS_PER_PAGE):
        item_idx = row_idx
        y_row_top = first_row_top - row_idx * ROW_H
        # inside the row box
        y_desc = y_row_top - 7.0
        y_nsn = y_row_top - 12.2

        if item_idx >= len(items):
            continue

        it = items[item_idx]

        # Box number
        _draw_center(c, str(it.line_no), X_BOX_L, X_BOX_R, y_desc, FONT_MAIN, 8)

        # Contents: description + NSN
        desc_lines = _wrap_to_width(it.description, FONT_MAIN, 6.8, max_content_w, max_lines=2)

        # If we used 2 lines for description, we may not have room for a separate NSN line.
        # In that case, append NSN to the last line (trim if needed).
        if it.nsn:
            nsn_label = f"NSN: {it.nsn}"
        else:
            nsn_label = ""

        if len(desc_lines) == 1:
            c.setFont(FONT_MAIN, 6.8)
            c.drawString(X_CONTENT_L + PAD_X, y_desc, desc_lines[0])
            if nsn_label:
                c.setFont(FONT_SMALL, 5.8)
                c.drawString(X_CONTENT_L + PAD_X, y_nsn, nsn_label)
        else:
            c.setFont(FONT_MAIN, 6.5)
            c.drawString(X_CONTENT_L + PAD_X, y_desc, desc_lines[0])
            line2 = desc_lines[1]
            if nsn_label:
                # try to append
                appended = (line2 + "  " + nsn_label).strip()
                if pdfmetrics.stringWidth(appended, FONT_MAIN, 6.0) <= max_content_w:
                    c.setFont(FONT_MAIN, 6.0)
                    c.drawString(X_CONTENT_L + PAD_X, y_nsn, appended)
                else:
                    # trim line2 to make space
                    c.setFont(FONT_MAIN, 6.0)
                    # reserve width for " … " + nsn_label
                    reserve = pdfmetrics.stringWidth(" … " + nsn_label, FONT_MAIN, 6.0)
                    avail = max(10.0, max_content_w - reserve)
                    trimmed = line2
                    while trimmed and pdfmetrics.stringWidth(trimmed, FONT_MAIN, 6.0) > avail:
                        trimmed = trimmed[:-1]
                    out_line = (trimmed.rstrip() + " … " + nsn_label).strip()
                    c.drawString(X_CONTENT_L + PAD_X, y_nsn, out_line)
            else:
                c.setFont(FONT_MAIN, 6.0)
                c.drawString(X_CONTENT_L + PAD_X, y_nsn, line2)

        # UOI + quantities
        _draw_center(c, "EA", X_UOI_L, X_UOI_R, y_desc, FONT_MAIN, 8)
        _draw_center(c, str(it.qty), X_INIT_L, X_INIT_R, y_desc, FONT_MAIN, 8)
        _draw_center(c, "0", X_SPARES_L, X_SPARES_R, y_desc, FONT_MAIN, 8)
        _draw_center(c, str(it.qty), X_TOTAL_L, X_TOTAL_R, y_desc, FONT_MAIN, 8)

    # Optional: page numbering fields could be filled here if desired.
    # Many units leave them blank; template might have them in header.

    c.showPage()
    c.save()
    return buf.getvalue()


def generate_dd1750_from_pdf(
    bom_pdf_path: str,
    template_pdf_path: str,
    out_pdf_path: str,
    start_page: int = 0,
) -> Tuple[str, int]:
    """Generate DD1750 PDF.

    Returns (out_pdf_path, item_count)
    """

    items = extract_items_from_pdf(bom_pdf_path, start_page=start_page)
    item_count = len(items)

    if item_count == 0:
        # still create a single-page copy of the template
        reader = PdfReader(template_pdf_path)
        writer = PdfWriter()
        writer.add_page(reader.pages[0])
        with open(out_pdf_path, "wb") as f:
            writer.write(f)
        return out_pdf_path, 0

    total_pages = math.ceil(item_count / ROWS_PER_PAGE)

    writer = PdfWriter()

    # NOTE: In some hosting environments, reusing (or even deep-copying)
    # a PageObject can still lead to aliasing where the last overlay is
    # replicated across all pages. Re-reading the template for each
    # output page avoids that class of bugs.
    for p in range(total_pages):
        chunk = items[p * ROWS_PER_PAGE : (p + 1) * ROWS_PER_PAGE]
        overlay_pdf = _build_overlay_page(chunk, page_num=p + 1, total_pages=total_pages)
        overlay_reader = PdfReader(io.BytesIO(overlay_pdf))

        fresh_template = PdfReader(template_pdf_path).pages[0]
        fresh_template.merge_page(overlay_reader.pages[0])
        writer.add_page(fresh_template)

    with open(out_pdf_path, "wb") as f:
        writer.write(f)

    return out_pdf_path, item_count
