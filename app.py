from flask import (
    Flask,
    render_template,
    request,
    send_file,
    redirect,
    url_for
)

import fitz
import pandas as pd
import os
import re
import shutil
import zipfile

from pathlib import Path
from openpyxl import Workbook

# ==========================================
# Flask
# ==========================================

app = Flask(__name__)

# ==========================================
# Disable MuPDF warnings
# ==========================================

fitz.TOOLS.mupdf_display_errors(False)

# ==========================================
# Base folders
# ==========================================

BASE_DIR = Path(__file__).resolve().parent

FONT_PATH = BASE_DIR / "fonts" / "NotoSans-Regular.ttf"

STAMP_UPLOAD = BASE_DIR / "uploads" / "stamp"
STAMP_OUTPUT = BASE_DIR / "output" / "stamp"

QTY_UPLOAD = BASE_DIR / "uploads" / "quantity"
QTY_OUTPUT = BASE_DIR / "output" / "quantity"

STAMP_UPLOAD.mkdir(parents=True, exist_ok=True)
STAMP_OUTPUT.mkdir(parents=True, exist_ok=True)

QTY_UPLOAD.mkdir(parents=True, exist_ok=True)
QTY_OUTPUT.mkdir(parents=True, exist_ok=True)

FONT = fitz.Font(fontfile=str(FONT_PATH))

PAGE_NUMBER = 0
TEXT_X = 850
TEXT_Y = 812
FONT_SIZE = 14

# ==========================================
# Home
# ==========================================

@app.route("/")
def home():

    return render_template(
        "index.html"
    )

# ==========================================
# Stamp Page
# ==========================================

@app.route("/stamp")
def stamp_page():

    return render_template(
        "stamp.html"
    )

# ==========================================
# Quantity Page
# ==========================================

@app.route("/quantity")
def quantity_page():

    return render_template(
        "quantity.html"
    )

# ==========================================
# Stamp PDFs
# ==========================================

@app.route(
    "/process_stamp",
    methods=["POST"]
)
def process_stamp():

    shutil.rmtree(
        STAMP_UPLOAD,
        ignore_errors=True
    )

    shutil.rmtree(
        STAMP_OUTPUT,
        ignore_errors=True
    )

    STAMP_UPLOAD.mkdir(
        parents=True,
        exist_ok=True
    )

    STAMP_OUTPUT.mkdir(
        parents=True,
        exist_ok=True
    )

    order = request.form.get(
        "order",
        ""
    ).strip()

    raw = request.form.get(
        "positions",
        ""
    )

    rows = []

    for line in raw.splitlines():

        line = line.strip()

        if not line:
            continue

        parts = re.split(
            r"[\s,;\t]+",
            line
        )

        if len(parts) < 2:
            continue

        try:

            pos = int(parts[0])

            qty = int(parts[1])

        except:

            continue

        rows.append({

            "Поз.": pos,

            "Кол-во": qty

        })

    if len(rows) == 0:

        return "No positions."

    df = pd.DataFrame(rows)

    pdfs = request.files.getlist(
        "pdfs"
    )

    if len(pdfs) == 0:

        return "No PDFs."

    for pdf in pdfs:

        if pdf.filename:

            pdf.save(

                STAMP_UPLOAD / pdf.filename

            )


                # ==========================================
    # Process Stamp PDFs
    # ==========================================

    pdf_files = sorted(STAMP_UPLOAD.glob("*.pdf"))

    for pdf_path in pdf_files:

        filename = pdf_path.name

        clean_name = re.sub(
            r"\s*\(\d+\)(?=\.pdf$)",
            "",
            filename
        )

        m = re.match(
            r"^\s*(\d+)",
            clean_name
        )

        if not m:
            print("Skipped:", filename)
            continue

        position = int(m.group(1))

        row = df[
            df["Поз."] == position
        ]

        if row.empty:
            print("Position not found:", position)
            continue

        qty = int(
            row.iloc[0]["Кол-во"]
        )

        stamp_text = (
            f"{order} заказ - "
            f"{position} поз. - "
            f"{qty} шт"
        )

        try:

            doc = fitz.open(str(pdf_path))

        except Exception as e:

            print(e)

            continue

        if PAGE_NUMBER >= len(doc):

            doc.close()

            continue

        page = doc[PAGE_NUMBER]

        writer = fitz.TextWriter(
            page.rect
        )

        writer.append(

            fitz.Point(
                TEXT_X,
                TEXT_Y
            ),

            stamp_text,

            font=FONT,

            fontsize=FONT_SIZE

        )

        writer.write_text(

            page,

            color=(0, 0, 0)

        )

        output_pdf = (
            STAMP_OUTPUT /
            clean_name
        )

        try:

            doc.save(

                str(output_pdf),

                garbage=4,

                clean=True,

                deflate=True

            )

        except Exception as e:

            print(e)

        doc.close()

    # ==========================================
    # Create ZIP
    # ==========================================

    zip_path = BASE_DIR / "Stamped_PDFs.zip"

    if zip_path.exists():

        zip_path.unlink()

    with zipfile.ZipFile(

        zip_path,

        "w",

        compression=zipfile.ZIP_DEFLATED

    ) as zipf:

        for pdf in STAMP_OUTPUT.glob("*.pdf"):

            zipf.write(

                pdf,

                arcname=pdf.name

            )

    return send_file(

        zip_path,

        as_attachment=True,

        download_name="Stamped_PDFs.zip"

    )

# ==========================================
# Quantity Extractor
# ==========================================

@app.route(
    "/process_quantity",
    methods=["POST"]
)
def process_quantity():

    shutil.rmtree(
        QTY_UPLOAD,
        ignore_errors=True
    )

    shutil.rmtree(
        QTY_OUTPUT,
        ignore_errors=True
    )

    QTY_UPLOAD.mkdir(
        parents=True,
        exist_ok=True
    )

    QTY_OUTPUT.mkdir(
        parents=True,
        exist_ok=True
    )

    pdfs = request.files.getlist(
        "pdfs"
    )

    if len(pdfs) == 0:

        return "No PDFs."

    for pdf in pdfs:

        if pdf.filename:

            pdf.save(

                QTY_UPLOAD /
                pdf.filename

            )

    wb = Workbook()

    ws = wb.active

    ws.title = "Results"

    ws.append([

    "File",

    "Quantity"

    ])

    grand_total = 0

    pattern = re.compile(
    r"(\d+)\s*заказ\s*-\s*(\d+)\s*поз\.\s*-\s*(\d+)\s*шт",
    re.IGNORECASE
    )

    pdf_files = sorted(QTY_UPLOAD.glob("*.pdf"))

    for pdf_path in pdf_files:

        filename = pdf_path.name

        order = ""
        position = ""
        qty = 0

        try:

            doc = fitz.open(str(pdf_path))

            text = ""

            for page in doc:

                text += page.get_text("text")

            doc.close()

            m = pattern.search(text)

            if m:

                order = m.group(1)

                position = m.group(2)

                qty = int(m.group(3))

            else:

                matches = re.findall(
                    r"(\d+)\s*шт",
                    text,
                    re.IGNORECASE
                )

                if matches:

                    qty = int(matches[-1])

            ws.append([
                filename,
                qty
            ])

            grand_total += qty

        except Exception as e:

            ws.append([
                filename,
                f"ERROR: {e}"
            ])

    ws.append([])

    ws.append([
        "TOTAL",
        grand_total
    ])

    excel_path = BASE_DIR / "PDF_QTY_Report.xlsx"

    wb.save(excel_path)

    return send_file(
        excel_path,
        as_attachment=True,
        download_name="PDF_QTY_Report.xlsx"
    )


if __name__ == "__main__":

    app.run(
        host="0.0.0.0",
        port=5000,
        debug=True
    )