
import os
import io
import html
import uuid

from flask import (
    Flask, render_template, request, session,
    send_file, redirect, url_for, flash
)
from werkzeug.utils import secure_filename

from pypdf import PdfReader
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from PIL import Image
import pytesseract

from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
)
from reportlab.lib.units import inch


# -------------------------------------------------
# APPLICATION SETTINGS
# -------------------------------------------------

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

app = Flask(__name__)
app.secret_key = os.environ.get(
    "SECRET_KEY", "question-paper-generator-new-2026"
)

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024

ALLOWED_EXTENSIONS = {"pdf", "docx", "jpg", "jpeg", "png"}

# Paper size: 6 x 9 inches
PAGE_WIDTH = 6 * inch
PAGE_HEIGHT = 9 * inch

# Windows Tesseract location
TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

if os.path.exists(TESSERACT_PATH):
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH


# -------------------------------------------------
# DEFAULT QUESTION PAPER DETAILS
# -------------------------------------------------

DEFAULT_DATA = {
    "course_code": "BSC-S551",
    "college": "ST. ANN'S COLLEGE FOR WOMEN",
    "autonomous": "(AUTONOMOUS)",
    "address": "Malkapuram, Visakhapatnam-530011",
    "exam": "B.Sc. Honors – Degree Examinations – V Semester",
    "batch": "Admitted Batch: 2024-2025",
    "subject": "CHEMISTRY (Major)",
    "paper_title": (
        "ANALYTICAL METHODS IN CHEMISTRY – "
        "QUANTITATIVE ANALYSIS"
    ),
    "time": "3 Hrs",
    "max_marks": "60",
    "section_a_instruction": "Answer any FIVE Questions",
    "section_a_marks": "5 × 4 = 20M",
    "section_b_instruction": "Answer ALL the following questions.",
    "section_b_marks": "5 × 8 = 40M",

    "section_a": [
        "Define molarity and molality with suitable formulas.",
        "Define standard solution and explain its preparation.",
        "Explain the principle of complexometric titrations.",
        "Differentiate between coagulation and peptization.",
        "Write a short note on significant figures.",
        "What is standard deviation? Explain with an example.",
        "Define synergism in solvent extraction.",
        "Explain determination of hardness of water using EDTA."
    ],

    "section_b_a": [
        "Differentiate between primary and secondary standards.",
        "Discuss the theories of acid-base titrations.",
        "What is precision? Explain methods of expressing precision.",
        "Explain the principle of solvent extraction.",
        "Explain determination of dissolved oxygen using Winkler's method."
    ],

    "section_b_b": [
        "Describe common laboratory apparatus used in volumetric analysis.",
        "Explain co-precipitation and post-precipitation.",
        "Explain determinate and indeterminate errors.",
        "Describe the determination of Iron (III) by solvent extraction.",
        "Explain determination of alkalinity in water."
    ]
}


def get_current_data():
    data = DEFAULT_DATA.copy()

    saved = session.get("paper_data")
    if saved:
        data.update(saved)

    for key, count in [
        ("section_a", 8),
        ("section_b_a", 5),
        ("section_b_b", 5)
    ]:
        values = list(data.get(key, []))
        data[key] = (values + [""] * count)[:count]

    return data


def allowed_file(filename):
    return (
        "." in filename
        and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS
    )


def clean_text(value):
    return str(value or "").strip()


def safe_paragraph(value):
    return html.escape(clean_text(value)).replace("\n", "<br/>")


# -------------------------------------------------
# FILE TEXT EXTRACTION / OCR
# -------------------------------------------------

def extract_file_text(filepath, extension):
    extracted = ""

    if extension == "pdf":
        reader = PdfReader(filepath)

        for page in reader.pages:
            extracted += (page.extract_text() or "") + "\n"

    elif extension == "docx":
        document = Document(filepath)

        for paragraph in document.paragraphs:
            extracted += paragraph.text + "\n"

    elif extension in ("jpg", "jpeg", "png"):
        image = Image.open(filepath)
        extracted = pytesseract.image_to_string(image, lang="eng")

    return extracted.strip()


@app.route("/upload", methods=["POST"])
def upload_file():
    file = request.files.get("source_file")

    if not file or not file.filename:
        flash("Please select a file.")
        return redirect(url_for("index"))

    if not allowed_file(file.filename):
        flash("Only PDF, DOCX, JPG and PNG files are allowed.")
        return redirect(url_for("index"))

    filename = secure_filename(file.filename)
    unique_name = str(uuid.uuid4()) + "_" + filename
    filepath = os.path.join(UPLOAD_FOLDER, unique_name)

    file.save(filepath)

    try:
        extension = filename.rsplit(".", 1)[1].lower()
        extracted_text = extract_file_text(filepath, extension)

        text_path = os.path.join(UPLOAD_FOLDER, "extracted_text.txt")

        with open(text_path, "w", encoding="utf-8") as f:
            f.write(extracted_text)

        return render_template(
            "index.html",
            data=get_current_data(),
            extracted_text=extracted_text,
            upload_message="Text extraction completed."
        )

    except Exception as error:
        return render_template(
            "index.html",
            data=get_current_data(),
            extracted_text="",
            upload_message="Extraction error: " + str(error)
        )


# -------------------------------------------------
# HOME PAGE
# -------------------------------------------------

@app.route("/")
def index():
    return render_template(
        "index.html",
        data=get_current_data(),
        extracted_text=""
    )


# -------------------------------------------------
# SAVE / PREVIEW QUESTION PAPER
# -------------------------------------------------

@app.route("/preview", methods=["POST"])
def preview():
    data = {}

    simple_fields = [
        "course_code",
        "college",
        "autonomous",
        "address",
        "exam",
        "batch",
        "subject",
        "paper_title",
        "time",
        "max_marks",
        "section_a_instruction",
        "section_a_marks",
        "section_b_instruction",
        "section_b_marks"
    ]

    for field in simple_fields:
        data[field] = clean_text(request.form.get(field))

    data["section_a"] = [
        clean_text(x)
        for x in request.form.getlist("section_a[]")
    ]

    data["section_b_a"] = [
        clean_text(x)
        for x in request.form.getlist("section_b_a[]")
    ]

    data["section_b_b"] = [
        clean_text(x)
        for x in request.form.getlist("section_b_b[]")
    ]

    data["section_a"] = (data["section_a"] + [""] * 8)[:8]
    data["section_b_a"] = (data["section_b_a"] + [""] * 5)[:5]
    data["section_b_b"] = (data["section_b_b"] + [""] * 5)[:5]

    session["paper_data"] = data

    return render_template("preview.html", data=data)


# -------------------------------------------------
# PDF GENERATION - 6 x 9 INCHES
# -------------------------------------------------

def make_pdf(data):
    buffer = io.BytesIO()

    document = SimpleDocTemplate(
        buffer,
        pagesize=(PAGE_WIDTH, PAGE_HEIGHT),
        rightMargin=0.32 * inch,
        leftMargin=0.32 * inch,
        topMargin=0.28 * inch,
        bottomMargin=0.30 * inch
    )

    content_width = PAGE_WIDTH - (0.64 * inch)

    styles = getSampleStyleSheet()

    college_style = ParagraphStyle(
        "College",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=12,
        leading=13,
        alignment=TA_CENTER,
        spaceAfter=0
    )

    heading_style = ParagraphStyle(
        "Heading",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=9,
        leading=10,
        alignment=TA_CENTER,
        spaceAfter=0
    )

    normal_style = ParagraphStyle(
        "NormalCustom",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8.5,
        leading=10,
        alignment=TA_LEFT,
        spaceAfter=2
    )

    center_style = ParagraphStyle(
        "CenterCustom",
        parent=normal_style,
        alignment=TA_CENTER
    )

    question_style = ParagraphStyle(
        "Question",
        parent=normal_style,
        fontSize=8.5,
        leading=10.5,
        spaceAfter=4
    )

    right_style = ParagraphStyle(
        "RightCustom",
        parent=normal_style,
        alignment=TA_RIGHT
    )

    story = []

    # PAPER CODE - CENTERED AT TOP
    story.append(
        Paragraph(
            "<b>PAPER CODE: "
            + safe_paragraph(data["course_code"])
            + "</b>",
            center_style
        )
    )
    story.append(Spacer(1, 4))

    # COLLEGE HEADING
    story.append(
        Paragraph(safe_paragraph(data["college"]), college_style)
    )
    story.append(
        Paragraph(safe_paragraph(data["autonomous"]), heading_style)
    )
    story.append(
        Paragraph(safe_paragraph(data["address"]), center_style)
    )
    story.append(
        Paragraph(safe_paragraph(data["exam"]), center_style)
    )
    story.append(
        Paragraph(safe_paragraph(data["batch"]), center_style)
    )
    story.append(
        Paragraph(safe_paragraph(data["subject"]), heading_style)
    )
    story.append(
        Paragraph(safe_paragraph(data["paper_title"]), heading_style)
    )

    story.append(Spacer(1, 6))

    # HALL TICKET LABEL OUTSIDE THE 11 BOXES
    ticket_cells = [
        Paragraph("<b>Hall Ticket No.:</b>", normal_style)
    ] + [""] * 11

    ticket_widths = [1.05 * inch] + [0.22 * inch] * 11

    ticket_table = Table(
        [ticket_cells],
        colWidths=ticket_widths,
        rowHeights=[0.23 * inch]
    )

    ticket_table.setStyle(TableStyle([
        # No border around the Hall Ticket label
        ("BOX", (1, 0), (-1, -1), 0.65, colors.black),
        ("INNERGRID", (1, 0), (-1, -1), 0.65, colors.black),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 1),
        ("RIGHTPADDING", (0, 0), (-1, -1), 1),
        ("TOPPADDING", (0, 0), (-1, -1), 1),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
    ]))

    story.append(ticket_table)
    story.append(Spacer(1, 5))

    # TIME LEFT AND MAXIMUM MARKS RIGHT
    info_table = Table(
        [[
            Paragraph(
                "<b>Time: " + safe_paragraph(data["time"]) + "</b>",
                normal_style
            ),
            Paragraph(
                "<b>Max. Marks: "
                + safe_paragraph(data["max_marks"]) + "</b>",
                right_style
            )
        ]],
        colWidths=[content_width / 2, content_width / 2]
    )

    info_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))

    story.append(info_table)
    story.append(Spacer(1, 5))

    # SECTION HEADING CENTER AND MARKS RIGHT
    def add_section_row(title, marks):
        row = Table(
            [[
                Paragraph("<b>" + title + "</b>", center_style),
                Paragraph("<b>" + safe_paragraph(marks) + "</b>", right_style)
            ]],
            colWidths=[
                content_width * 0.70,
                content_width * 0.30
            ]
        )

        row.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ]))

        story.append(row)

    # SECTION A
    add_section_row("SECTION – A", data["section_a_marks"])

    story.append(
        Paragraph(
            safe_paragraph(data["section_a_instruction"]),
            normal_style
        )
    )

    for number, question in enumerate(data["section_a"], start=1):
        if question:
            story.append(
                Paragraph(
                    f"<b>{number}.</b> {safe_paragraph(question)}",
                    question_style
                )
            )

    story.append(Spacer(1, 4))

    # SECTION B
    add_section_row("SECTION – B", data["section_b_marks"])

    story.append(
        Paragraph(
            safe_paragraph(data["section_b_instruction"]),
            normal_style
        )
    )

    for i in range(5):
        part_a = data["section_b_a"][i]
        part_b = data["section_b_b"][i]
        number = 9 + i

        if part_a:
            story.append(
                Paragraph(
                    f"<b>{number}. (a)</b> "
                    + safe_paragraph(part_a),
                    question_style
                )
            )

        # OR - CENTERED, NOT BOLD
        story.append(
            Paragraph("OR", center_style)
        )

        if part_b:
            story.append(
                Paragraph(
                    "<b>(b)</b> " + safe_paragraph(part_b),
                    question_style
                )
            )

    # STARS ONLY AT THE END
    story.append(Spacer(1, 7))
    story.append(
        Paragraph(
            "* * * * *",
            ParagraphStyle(
                "Stars",
                parent=center_style,
                fontSize=10,
                leading=12
            )
        )
    )

    document.build(story)
    buffer.seek(0)

    return buffer


@app.route("/download/pdf")
def download_pdf():
    pdf_buffer = make_pdf(get_current_data())

    return send_file(
        pdf_buffer,
        as_attachment=True,
        download_name="Question_Paper.pdf",
        mimetype="application/pdf"
    )


# -------------------------------------------------
# WORD DOCUMENT - 6 x 9 INCHES
# -------------------------------------------------

def make_word(data):
    document = Document()

    section = document.sections[0]
    section.page_width = Inches(6)
    section.page_height = Inches(9)
    section.top_margin = Inches(0.28)
    section.bottom_margin = Inches(0.30)
    section.left_margin = Inches(0.32)
    section.right_margin = Inches(0.32)

    def format_paragraph(paragraph, alignment=None, after=0):
        if alignment is not None:
            paragraph.alignment = alignment

        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(after)
        paragraph.paragraph_format.line_spacing = 1.0

    def format_run(run, bold=False, size=9):
        run.bold = bold
        run.font.name = "Times New Roman"
        run.font.size = Pt(size)

    def add_center(text, bold=False, size=9):
        paragraph = document.add_paragraph()
        format_paragraph(paragraph, WD_ALIGN_PARAGRAPH.CENTER)

        run = paragraph.add_run(str(text))
        format_run(run, bold, size)

        return paragraph

    def add_question(label, text):
        paragraph = document.add_paragraph()
        format_paragraph(paragraph, after=3)

        run = paragraph.add_run(label + " ")
        format_run(run, True, 9)

        run = paragraph.add_run(text)
        format_run(run, False, 9)

        return paragraph

    def set_cell_text(cell, text, bold=False, size=9, alignment=None):
        cell.text = ""

        paragraph = cell.paragraphs[0]
        format_paragraph(paragraph, alignment)

        run = paragraph.add_run(str(text))
        format_run(run, bold, size)

        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    def remove_cell_borders(cell):
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()

        borders = tcPr.first_child_found_in("w:tcBorders")

        if borders is None:
            borders = OxmlElement("w:tcBorders")
            tcPr.append(borders)

        for edge in (
            "top", "left", "bottom", "right",
            "insideH", "insideV"
        ):
            tag = "w:" + edge
            element = borders.find(qn(tag))

            if element is None:
                element = OxmlElement(tag)
                borders.append(element)

            element.set(qn("w:val"), "nil")

    # PAPER CODE - CENTERED AT TOP
    add_center(
        "PAPER CODE: " + data["course_code"],
        True,
        10
    )

    # COLLEGE HEADING
    add_center(data["college"], True, 12)
    add_center(data["autonomous"], True, 9)
    add_center(data["address"], False, 8)
    add_center(data["exam"], False, 8)
    add_center(data["batch"], False, 8)
    add_center(data["subject"], True, 10)
    add_center(data["paper_title"], True, 9)

    # HALL TICKET LABEL OUTSIDE THE 11 BOXES
    ticket = document.add_table(rows=1, cols=12)
    ticket.alignment = WD_TABLE_ALIGNMENT.CENTER
    ticket.autofit = False
    ticket.style = "Table Grid"

    ticket.columns[0].width = Inches(1.05)

    set_cell_text(
        ticket.cell(0, 0),
        "Hall Ticket No.:",
        True,
        8
    )

    # Remove border around the label cell
    remove_cell_borders(ticket.cell(0, 0))

    # 11 empty boxes
    for i in range(1, 12):
        ticket.columns[i].width = Inches(0.22)
        set_cell_text(
            ticket.cell(0, i),
            "",
            False,
            8
        )

    document.add_paragraph().paragraph_format.space_after = Pt(1)

    # TIME AND MAXIMUM MARKS
    info = document.add_table(rows=1, cols=2)
    info.alignment = WD_TABLE_ALIGNMENT.CENTER
    info.autofit = False

    info.columns[0].width = Inches(2.68)
    info.columns[1].width = Inches(2.68)

    set_cell_text(
        info.cell(0, 0),
        "Time: " + data["time"],
        True,
        9
    )

    set_cell_text(
        info.cell(0, 1),
        "Max. Marks: " + data["max_marks"],
        True,
        9,
        WD_ALIGN_PARAGRAPH.RIGHT
    )

    # SECTION HEADING CENTER AND MARKS RIGHT
    def add_section_row(title, marks):
        table = document.add_table(rows=1, cols=3)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False

        table.columns[0].width = Inches(1.60)
        table.columns[1].width = Inches(2.16)
        table.columns[2].width = Inches(1.60)

        set_cell_text(
            table.cell(0, 0),
            "",
            False,
            9
        )

        set_cell_text(
            table.cell(0, 1),
            title,
            True,
            10,
            WD_ALIGN_PARAGRAPH.CENTER
        )

        set_cell_text(
            table.cell(0, 2),
            marks,
            True,
            9,
            WD_ALIGN_PARAGRAPH.RIGHT
        )

        # Hide table borders
        for row in table.rows:
            for cell in row.cells:
                remove_cell_borders(cell)

    # SECTION A
    add_section_row("SECTION – A", data["section_a_marks"])

    paragraph = document.add_paragraph()
    format_paragraph(paragraph, after=3)

    run = paragraph.add_run(data["section_a_instruction"])
    format_run(run, True, 9)

    for i, question in enumerate(data["section_a"], start=1):
        if question:
            add_question(str(i) + ".", question)

    # SECTION B
    add_section_row("SECTION – B", data["section_b_marks"])

    paragraph = document.add_paragraph()
    format_paragraph(paragraph, after=3)

    run = paragraph.add_run(data["section_b_instruction"])
    format_run(run, True, 9)

    for i in range(5):
        part_a = data["section_b_a"][i]
        part_b = data["section_b_b"][i]
        number = 9 + i

        if part_a:
            add_question(f"{number}. (a)", part_a)

        # OR - CENTERED, NOT BOLD
        add_center("OR", False, 9)

        if part_b:
            add_question("(b)", part_b)

    # STARS ONLY AT THE END
    add_center("* * * * *", True, 11)

    buffer = io.BytesIO()
    document.save(buffer)
    buffer.seek(0)

    return buffer


@app.route("/download/word")
def download_word():
    word_buffer = make_word(get_current_data())

    return send_file(
        word_buffer,
        as_attachment=True,
        download_name="Question_Paper.docx",
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "wordprocessingml.document"
        )
    )


# -------------------------------------------------
# PRINT ROUTE
# -------------------------------------------------

@app.route("/print")
def print_paper():
    return render_template(
        "preview.html",
        data=get_current_data(),
        print_mode=True
    )


# -------------------------------------------------
# ERROR HANDLING
# -------------------------------------------------

@app.errorhandler(413)
def file_too_large(error):
    return "File is too large. Maximum upload size is 25 MB.", 413


@app.errorhandler(404)
def page_not_found(error):
    return "Page not found.", 404


# -------------------------------------------------
# RUN APPLICATION
# -------------------------------------------------

if __name__ == "__main__":
    print("-----------------------------------------")
    print("ST. ANN'S QUESTION PAPER GENERATOR")
    print("NEW APPLICATION")
    print("PAPER SIZE: 6 x 9 INCHES")
    print("PAPER CODE: CENTERED AT TOP")
    print("SECTION A / B: CENTERED")
    print("OR: CENTERED, NOT BOLD")
    print("HALL TICKET: LABEL OUTSIDE 11 BOXES")
    print("http://127.0.0.1:5000")
    print("-----------------------------------------")

    app.run(
        host="127.0.0.1",
        port=5000,
        debug=True
    )
