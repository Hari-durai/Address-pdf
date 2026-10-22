import os
import io
import csv
import logging
import pathlib
import re
from werkzeug.exceptions import RequestEntityTooLarge
from werkzeug.middleware.proxy_fix import ProxyFix
from zipfile import ZipFile
from flask import Flask, render_template, request, send_file, abort
from openpyxl import Workbook
from openpyxl.styles import Alignment
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

app = Flask(__name__)
# Trust Render's proxy to detect HTTPS correctly
app.wsgi_app = ProxyFix(app.wsgi_app, x_proto=1, x_host=1)
app.config.update(
    MAX_CONTENT_LENGTH=int(os.environ.get("MAX_UPLOAD_MB", "5")) * 1024 * 1024,
    SESSION_COOKIE_SECURE=True,
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    PREFERRED_URL_SCHEME="https",
)
logging.getLogger("werkzeug").setLevel(logging.WARNING)
app.logger.setLevel(logging.WARNING)


def read_csv_from_upload(file_storage):
    data = file_storage.read()
    try:
        text = data.decode("utf-8-sig")
    except Exception:
        text = data.decode("utf-8", errors="ignore")
    rows = list(csv.reader(io.StringIO(text)))
    return rows


def build_xlsx(rows):
    wb = Workbook()
    ws = wb.active
    for r in rows:
        ws.append(r)
    wrap_cols = set()
    if rows:
        header = rows[0]
        for idx, name in enumerate(header):
            n = (name or "").strip().lower()
            if n in {"billing address", "items ordered", "address", "notes"}:
                wrap_cols.add(idx + 1)
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            if cell.column in wrap_cols:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
    default_widths = {
        1: 12,
        2: 14,
        3: 18,
        4: 18,
        5: 28,
        6: 16,
        7: 50,
        8: 50,
        9: 12,
        10: 20,
    }
    for col in range(1, ws.max_column + 1):
        width = default_widths.get(col, 18)
        ws.column_dimensions[chr(64 + col)].width = width
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out


def try_register_unicode_font():
    font_path = os.path.join(app.root_path, "static", "fonts", "NotoSansTamil-Regular.ttf")
    if os.path.exists(font_path):
        try:
            pdfmetrics.registerFont(TTFont("NotoSansTamil", font_path))
            return "NotoSansTamil"
        except Exception:
            pass
    return "Helvetica"


def normalize_text(s: str) -> str:
    # Collapse whitespace, remove duplicate spaces/commas, strip control chars
    s = s.replace("\r", "\n")
    s = re.sub(r"[\t\u200b\ufeff]", "", s)
    s = re.sub(r"[ ]{2,}", " ", s)
    s = re.sub(r"\n{2,}", "\n", s)
    s = re.sub(r"\s+,", ",", s)
    s = re.sub(r",\s+", ", ", s)
    return s.strip()


def sanitize_filename(name: str) -> str:
    # Keep only safe characters and ensure non-empty
    base = re.sub(r"[^A-Za-z0-9._-]+", "_", name).strip("._-")
    return base or "file"


def get_from_text() -> str:
    # Static FROM address as requested by user
    return "\n \n MANNAI FOODS,\n Asesham, mannargudi,\npin-614016,\nph:9344769473"


def wrap_text(c, text, x, y, max_width, line_height, font_name, font_size):
    if not text:
        return y
    paragraphs = text.replace("\r", "").split("\n")
    for para in paragraphs:
        words = para.split()
        if not words:
            # blank line: keep vertical space
            y -= line_height
            continue
        line = ""
        for w in words:
            test = (line + (" " if line else "") + w)
            width = c.stringWidth(test, font_name, font_size)
            if width <= max_width:
                line = test
            else:
                c.drawString(x, y, line)
                y -= line_height
                line = w
        if line:
            c.drawString(x, y, line)
            y -= line_height
    return y


def build_pdf(rows):
    font_name = try_register_unicode_font()
    buff = io.BytesIO()
    c = canvas.Canvas(buff, pagesize=A4)
    page_w, page_h = A4
    margin = 12 * mm
    row_gap = 6 * mm
    rows_per_page = 3
    row_h = (page_h - 2 * margin - (rows_per_page - 1) * row_gap) / rows_per_page
    table_w = page_w - 2 * margin
    col_w = table_w / 2.0
    from_text = get_from_text() or ""
    font_size_title = 14
    font_size_text = 12
    line_height = font_size_text + 3
    pad_x = 6 * mm
    pad_y = 6 * mm

    def draw_row(row_index, to_lines):
        y_top = page_h - (margin + row_index * (row_h + row_gap))
        # Outer row border
        c.rect(margin, y_top - row_h, table_w, row_h, stroke=1, fill=0)
        # Middle vertical line splitting TO | FROM
        c.line(margin + col_w, y_top, margin + col_w, y_top - row_h)

        # Left cell: FROM
        from_x = margin + pad_x
        from_y = y_top - pad_y
        c.setFont("Helvetica", font_size_title)
        c.drawString(from_x, from_y, "FROM:")
        c.setFont(font_name, font_size_text)
        yy2 = from_y - 16
        # Center-align FROM text
        from_lines = from_text.replace("\r", "").split("\n")
        for line in from_lines:
            line_width = c.stringWidth(line, font_name, font_size_text)
            x_centered = from_x + (col_w - 2 * pad_x - line_width) / 2
            if x_centered < from_x:
                x_centered = from_x  # avoid overflow
            c.drawString(x_centered, yy2, line)
            yy2 -= line_height

        # Right cell: TO
        to_x = margin + col_w + pad_x
        to_y = y_top - pad_y
        c.setFont("Helvetica", font_size_title)
        c.drawString(to_x, to_y, "TO: ")
        c.setFont(font_name, font_size_text)
        yy = to_y - 16
        # Add double blank line after "TO:" label
        yy -= 2 * line_height
        for i, line in enumerate(to_lines):
            yy = wrap_text(c, line, to_x, yy, col_w - 2 * pad_x, line_height, font_name, font_size_text)
            # Add a blank line between TO address lines (but not after the last line)
            if i < len(to_lines) - 1:
                yy -= line_height

    to_blocks = []
    for i, row in enumerate(rows[1:] if rows else []):
        order_id = row[0] if len(row) > 0 else ""
        status = row[1] if len(row) > 1 else ""
        created = row[2] if len(row) > 2 else ""
        customer = row[3] if len(row) > 3 else ""
        email = row[4] if len(row) > 4 else ""
        phone = row[5] if len(row) > 5 else ""
        billing_address = row[6] if len(row) > 6 else ""
        block = []
        if customer:
            block.append(str(customer))
        if order_id:
            block.append("Order ID: " + str(order_id))
        if billing_address:
            block.append(str(billing_address))
        if phone:
            block.append("Phone: " + str(phone))
        if block:
            to_blocks.append(block)

    idx = 0
    for block in to_blocks:
        draw_row(idx % rows_per_page, block)
        idx += 1
        if idx % rows_per_page == 0:
            c.showPage()
    if idx % rows_per_page != 0:
        c.showPage()
    c.save()
    buff.seek(0)
    return buff


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/process", methods=["POST"])
def process():
    if "file" not in request.files:
        abort(400, "CSV file missing")
    f = request.files["file"]
    if not f.filename.lower().endswith(".csv"):
        abort(400, "Please upload a .csv file")
    rows = read_csv_from_upload(f)
    if not rows:
        abort(400, "CSV is empty")
    xlsx_buf = build_xlsx(rows)
    pdf_buf = build_pdf(rows)
    zip_buf = io.BytesIO()
    # Use CSV filename (without extension) for output files, sanitized
    base_name = os.path.splitext(f.filename)[0]
    base_name = sanitize_filename(base_name)
    with ZipFile(zip_buf, mode="w") as z:
        z.writestr(f"{base_name}.xlsx", xlsx_buf.getvalue())
        z.writestr(f"{base_name}.pdf", pdf_buf.getvalue())
    zip_buf.seek(0)
    return send_file(
        zip_buf,
        mimetype="application/zip",
        as_attachment=True,
        download_name=f"{base_name}_outputss.zip",
    )


@app.errorhandler(RequestEntityTooLarge)
def handle_file_too_large(e):
    mb = int(os.environ.get("MAX_UPLOAD_MB", "5"))
    return (f"File too large. Max {mb} MB.", 413)


@app.after_request
def add_security_headers(resp):
    resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    resp.headers["Pragma"] = "no-cache"
    resp.headers["Expires"] = "0"
    resp.headers["X-Content-Type-Options"] = "nosniff"
    resp.headers["X-Frame-Options"] = "DENY"
    if request.is_secure:
        resp.headers["Strict-Transport-Security"] = "max-age=31536000; includeSubDomains"
    return resp


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    app.run(host="0.0.0.0", port=port)
