from __future__ import annotations

import json
import csv
import io
import zipfile
import uuid
from datetime import datetime
from pathlib import Path

from flask import Flask, abort, render_template, request, send_from_directory, url_for, Response
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from werkzeug.utils import secure_filename

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
PDF_DIR = BASE_DIR / "generated_pdfs"
RECORDS_FILE = BASE_DIR / "submission_records.json"
ALLOWED_EXTENSIONS = {"pdf", "jpg", "jpeg", "png", "doc", "docx"}
ALLOWED_IMAGE_EXTENSIONS = {"jpg", "jpeg", "png"}
MAX_STUDENT_IMAGE_BYTES = 500 * 1024
MIN_STUDENT_IMAGE_BYTES = 200 * 1024

UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
PDF_DIR.mkdir(parents=True, exist_ok=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def allowed_image_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_IMAGE_EXTENSIONS


def load_submission_records() -> list[dict]:
    if not RECORDS_FILE.exists():
        return []
    try:
        data = json.loads(RECORDS_FILE.read_text(encoding="utf-8"))
        return data if isinstance(data, list) else []
    except (json.JSONDecodeError, OSError):
        return []


def ensure_record_ids(records: list[dict]) -> list[dict]:
    updated = False
    for row in records:
        if not str(row.get("record_id", "")).strip():
            row["record_id"] = uuid.uuid4().hex
            updated = True
    if updated:
        RECORDS_FILE.write_text(json.dumps(records, ensure_ascii=True, indent=2), encoding="utf-8")
    return records


def append_submission_record(record: dict) -> None:
    records = load_submission_records()
    record["record_id"] = uuid.uuid4().hex
    records.insert(0, record)
    RECORDS_FILE.write_text(json.dumps(records, ensure_ascii=True, indent=2), encoding="utf-8")


def get_record_month_key(record: dict) -> str:
    ts = record.get("created_ts")
    if isinstance(ts, (int, float)):
        return datetime.fromtimestamp(ts).strftime("%Y-%m")
    raw_date = str(record.get("date", "")).strip()
    try:
        return datetime.strptime(raw_date, "%d-%m-%Y %I:%M %p").strftime("%Y-%m")
    except ValueError:
        return ""


def filter_records_by_month(records: list[dict], month_key: str) -> list[dict]:
    if not month_key:
        return records
    return [row for row in records if get_record_month_key(row) == month_key]


def format_month_label(month_key: str) -> str:
    try:
        return datetime.strptime(month_key, "%Y-%m").strftime("%b %Y")
    except ValueError:
        return month_key


def build_csv_text(records: list[dict]) -> str:
    buffer = io.StringIO()
    writer = csv.writer(buffer)
    writer.writerow(
        [
            "LetterHead No",
            "Student Name",
            "Recommendation Letter",
            "Date",
            "Internship PDF",
            "NOC PDF",
        ]
    )
    for row in records:
        writer.writerow(
            [
                row.get("letterhead_no", ""),
                row.get("student_name", ""),
                row.get("recommendation_letter", ""),
                row.get("date", ""),
                row.get("internship_pdf", ""),
                row.get("noc_pdf", ""),
            ]
        )
    return buffer.getvalue()


def write_paragraph(
    c: canvas.Canvas,
    text: str,
    x: int,
    y: int,
    max_width: int,
    font_name: str,
    font_size: int,
    line_height: int,
) -> int:
    words = text.split()
    line = ""
    for word in words:
        test_line = f"{line} {word}".strip()
        if c.stringWidth(test_line, font_name, font_size) <= max_width:
            line = test_line
        else:
            c.drawString(x, y, line)
            y -= line_height
            line = word
    if line:
        c.drawString(x, y, line)
        y -= line_height
    return y


def draw_underlined_value(
    c: canvas.Canvas,
    label: str,
    value: str,
    x: int,
    y: int,
    width: int,
    font_name: str,
    font_size: int,
    line_height: int,
) -> int:
    c.setFont(font_name, font_size)
    c.drawString(x, y, label)
    value_x = x + width
    c.drawString(value_x, y, value)
    return y - max(line_height - 1, 8)


def write_paragraph_with_bold_phrase(
    c: canvas.Canvas,
    text: str,
    bold_phrase: str,
    x: int,
    y: int,
    max_width: int,
    normal_font: str,
    bold_font: str,
    font_size: int,
    line_height: int,
) -> int:
    if bold_phrase not in text:
        return write_paragraph(c, text, x, y, max_width, normal_font, font_size, line_height)

    words = text.split()
    line = ""
    for word in words:
        test_line = f"{line} {word}".strip()
        if c.stringWidth(test_line, normal_font, font_size) <= max_width:
            line = test_line
        else:
            if bold_phrase in line:
                pre, post = line.split(bold_phrase, 1)
                c.setFont(normal_font, font_size)
                c.drawString(x, y, pre)
                pre_w = c.stringWidth(pre, normal_font, font_size)
                c.setFont(bold_font, font_size)
                c.drawString(x + pre_w, y, bold_phrase)
                bold_w = c.stringWidth(bold_phrase, bold_font, font_size)
                c.setFont(normal_font, font_size)
                c.drawString(x + pre_w + bold_w, y, post)
            else:
                c.setFont(normal_font, font_size)
                c.drawString(x, y, line)
            y -= line_height
            line = word
    if line:
        if bold_phrase in line:
            pre, post = line.split(bold_phrase, 1)
            c.setFont(normal_font, font_size)
            c.drawString(x, y, pre)
            pre_w = c.stringWidth(pre, normal_font, font_size)
            c.setFont(bold_font, font_size)
            c.drawString(x + pre_w, y, bold_phrase)
            bold_w = c.stringWidth(bold_phrase, bold_font, font_size)
            c.setFont(normal_font, font_size)
            c.drawString(x + pre_w + bold_w, y, post)
        else:
            c.setFont(normal_font, font_size)
            c.drawString(x, y, line)
        y -= line_height
    return y


def draw_from_to_line(
    c: canvas.Canvas,
    x: int,
    y: int,
    from_date: str,
    to_date: str,
    normal_font: str,
    bold_font: str,
    font_size: int,
    line_height: int,
) -> int:
    c.setFont(bold_font, font_size)
    c.drawString(x, y, "from")
    from_w = c.stringWidth("from", bold_font, font_size)
    c.setFont(normal_font, font_size)
    c.drawString(x + from_w + 6, y, f"{from_date} ")
    date_from_w = c.stringWidth(f"{from_date} ", normal_font, font_size)
    c.setFont(bold_font, font_size)
    c.drawString(x + from_w + 6 + date_from_w, y, "to")
    to_w = c.stringWidth("to", bold_font, font_size)
    c.setFont(normal_font, font_size)
    c.drawString(x + from_w + 6 + date_from_w + to_w + 6, y, f"{to_date}")
    return y - line_height


def format_ddmmyyyy(value: str) -> str:
    try:
        return datetime.strptime(value, "%Y-%m-%d").strftime("%d-%m-%Y")
    except ValueError:
        return value


def _render_pdf(c: canvas.Canvas, form_data: dict, font_size: int, line_height: int) -> int:
    _, height = A4
    left = 70
    width = 450
    bold = "Helvetica-Bold"
    normal = "Helvetica"

    c.setFont(bold, font_size)
    c.drawString(left, height - 75, f"Letterhead No. {form_data['letterhead_no']}")
    c.drawString(360, height - 75, f"Issuing Date: {datetime.now().strftime('%d-%m-%Y')}")

    y = height - 110
    c.setFont(normal, font_size)
    c.drawString(left, y, "To,")
    y -= line_height + 8
    c.drawString(left, y, form_data["reporting_manager"] or "Reporting Head")
    y -= line_height
    c.drawString(left, y, form_data["company_name"])

    y -= line_height + 8
    c.setFont(bold, font_size)
    c.drawString(left, y, "Sub: Internship in your esteemed organization")
    y -= line_height + 2
    c.setFont(bold, font_size)
    c.drawString(left, y, "Sir,")

    y -= line_height + 8
    para_1_prefix = "Kindly allow me to introduce "
    para_1_bold = "Rama University"
    para_1_suffix = ", Uttar Pradesh."
    c.setFont(normal, font_size)
    c.drawString(left, y, para_1_prefix)
    prefix_w = c.stringWidth(para_1_prefix, normal, font_size)
    c.setFont(bold, font_size)
    c.drawString(left + prefix_w, y, para_1_bold)
    bold_w = c.stringWidth(para_1_bold, bold, font_size)
    c.setFont(normal, font_size)
    c.drawString(left + prefix_w + bold_w, y, para_1_suffix)
    y -= line_height

    y -= 4
    para_2 = (
        "It gives me immense happiness to share with you that Rama University is renowned name in the world "
        "of education. Recognized by UGC, Government of India, the university is rising as one of the largest "
        "educational establishments in the country. Involved in imparting world class education, Rama University "
        "has shaped more than 10,000 professionals so far and offers more than 100 courses across 11 specialized "
        "fields. Based out of Delhi-NCR and Kanpur, Rama University has two hi-tech campuses spread across more "
        "than 150 acres of collective area. With lush green environment, serene surroundings and all the necessary "
        "facilities available within campus, the University promises to be a perfect place for learning. With "
        "highly educated faculty members along with the international experience, progressive learning approach and "
        "modern teaching methodology; Rama University has 09 constituent faculties, 5 teaching hospitals and "
        "state-of-the-art research centers to efficiently cater to the students from all over the country. Rama "
        "University offers education at par with global paradigms. The dynamic environment in the faculties not only "
        "ensures enormous growth potential but also promotes intellectual as well as personal growth."
    )
    y = write_paragraph(c, para_2, left, y, width, normal, font_size, line_height)

    y -= 6
    para_3 = (
        "Our University conducts Two year, three years, four years and Five years degree courses. The intake of "
        "students is through MAT/All India Engineering Entrance Examination (AIEEE)/RUET. We are committed to "
        "provide excellent learning environment and facilities for professional orientation."
    )
    y = write_paragraph(c, para_3, left, y, width, normal, font_size, line_height)

    y -= 6
    para_4 = (
        "The curriculum is transacted by highly qualified faculty using world-class facilities in the following "
        "disciplines-"
    )
    y = write_paragraph(c, para_4, left, y, width, normal, font_size, line_height)
    y -= 2
    c.setFont(bold, font_size)
    y = write_paragraph(
        c,
        "BBA, MBA (Digital Marketing, Human Resource, Marketing, Finance & International Business),",
        left,
        y,
        width,
        bold,
        font_size,
        line_height,
    )
    y = write_paragraph(
        c,
        "B.Tech. (CSE, ME, CE, Biotechnology), B.Sc Agriculture, Pharmacy (D.Pharm & B.Pharm)",
        left,
        y,
        width,
        bold,
        font_size,
        line_height,
    )
    y = write_paragraph(
        c,
        "Diploma (CSE, EE, ME, CE), Mass Communication (PGDJMC, BJMC, MJMC),",
        left,
        y,
        width,
        bold,
        font_size,
        line_height,
    )
    y = write_paragraph(
        c,
        "Juridical Sciences (BALLB, BBALLB, LLB, LLM), Nursing, Paramedical, Dental & Medical.",
        left,
        y,
        width,
        bold,
        font_size,
        line_height,
    )

    y -= 6
    c.setFont(normal, font_size)
    para_5 = (
        "I would like to share with you that RAMA University has arranged for round-the-year technical trainings "
        "for its students on campus which are a regular part of their academic curriculum."
    )
    y = write_paragraph_with_bold_phrase(
        c,
        para_5,
        "RAMA University",
        left,
        y,
        width,
        normal,
        bold,
        font_size,
        line_height,
    )

    y -= 6
    para_6_prefix = (
        "Therefore, we would like you to help our students to get the right exposure and kindly allow them to "
        "undergo Online/ Offline training through your esteemed organization "
    )
    y = write_paragraph(c, para_6_prefix, left, y, width, normal, font_size, line_height)
    y = draw_from_to_line(
        c,
        left,
        y,
        format_ddmmyyyy(form_data["date_from"]),
        format_ddmmyyyy(form_data["date_to"]),
        normal,
        bold,
        font_size,
        line_height,
    )
    y -= 4

    y -= 4
    y = draw_underlined_value(
        c,
        "Student Name:",
        form_data["student_name"],
        left,
        y,
        80,
        normal,
        font_size,
        line_height,
    )
    y = draw_underlined_value(
        c,
        "Roll Number:",
        form_data["roll_no"],
        left,
        y,
        80,
        normal,
        font_size,
        line_height,
    )
    y = draw_underlined_value(c, "Course:", form_data["course"], left, y, 80, normal, font_size, line_height)
    y = draw_underlined_value(c, "Contact No:", form_data["phone"], left, y, 80, normal, font_size, line_height)
    y = draw_underlined_value(c, "Mail Id:", form_data["email"], left, y, 80, normal, font_size, line_height)
    y -= 8

    y -= 4
    c.setFont(normal, font_size)
    c.drawString(left, y, "Looking forward to your co-operation.")
    y -= line_height + 14
    c.drawString(left, y, "Thanks and Regards")

    y -= line_height + 10
    c.setFont(bold, font_size)
    c.drawString(left, y, "Mr. Saurabh Samaddar")
    y -= line_height
    c.drawString(left, y, "H.O.D. | Training & Development")
    y -= line_height
    c.setFont(normal, font_size)
    c.drawString(left, y, "Rama University, Uttar Pradesh")
    y -= line_height
    c.drawString(left, y, "Contact: +91- 8588816343")
    y -= line_height
    c.drawString(left, y, "Email: head.tnp@ramauniversity.ac.in")
    y -= line_height
    c.drawString(left, y, "For more details, please visit www.ramauniversity.ac.in")

    return y


def generate_pdf(form_data: dict) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    pdf_path = PDF_DIR / f"internship_form_{form_data['roll_no']}_{timestamp}.pdf"

    for font_size, line_height in [(11, 12), (10, 11), (9, 10)]:
        c = canvas.Canvas(str(pdf_path), pagesize=A4)
        y = _render_pdf(c, form_data, font_size, line_height)
        if y >= 60:
            c.save()
            return pdf_path
        c.save()

    return pdf_path


def write_centered(c: canvas.Canvas, text: str, y: int, font: str, size: int) -> None:
    c.setFont(font, size)
    width, _ = A4
    text_width = c.stringWidth(text, font, size)
    x = (width - text_width) / 2
    c.drawString(x, y, text)


def generate_noc_pdf(form_data: dict) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    pdf_path = PDF_DIR / f"noc_form_{form_data['roll_no']}_{timestamp}.pdf"

    c = canvas.Canvas(str(pdf_path), pagesize=A4)
    page_w, page_h = A4

    c.setFillColorRGB(1, 1, 1)
    c.rect(0, 0, page_w, page_h, stroke=0, fill=1)
    c.setFillColorRGB(0, 0, 0)

    c.setFont("Helvetica", 15)
    c.drawString(55, page_h - 110, f"Letter No. {form_data['letterhead_no']}")
    c.drawRightString(page_w - 55, page_h - 110, "Issued on Date:")
    c.drawRightString(page_w - 55, page_h - 132, datetime.now().strftime("%d/%b/%y"))

    title_y = page_h - 200
    write_centered(c, "NO OBJECTION CERTIFICATE", int(title_y), "Helvetica-Bold", 24)
    c.line(160, title_y - 4, page_w - 160, title_y - 4)

    left = 55
    text_w = page_w - 110
    y = page_h - 260
    c.setFont("Helvetica", 15)

    p1 = (
        f"This is to certify that {form_data['student_name']}, STUDENT OF {form_data['course']} "
        f"{form_data['year']} Year is hereby permitted to take up Internship as per their discretion. "
        f"Furthermore, this update has been duly recorded in the official college records."
    )
    y = write_paragraph(c, p1, left, y, int(text_w), "Helvetica", 15, 28)

    y -= 16
    p2 = (
        f"This certificate is issued upon the student's request for the purpose of Internship only with "
        f"{form_data['company_name']} which is Valid from {format_ddmmyyyy(form_data['date_from'])} "
        f"to {format_ddmmyyyy(form_data['date_to'])}."
    )
    y = write_paragraph(c, p2, left, y, int(text_w), "Helvetica", 15, 28)

    y -= 16
    p3 = (
        f"We extend our best wishes to {form_data['student_name']} for continued success in their future endeavors."
    )
    y = write_paragraph(c, p3, left, y, int(text_w), "Helvetica", 15, 28)

    y -= 70
    c.setFont("Helvetica", 15)
    c.drawString(left, y, "Signed By")
    y -= 34
    c.setFont("Helvetica-Bold", 15)
    c.drawString(left, y, "Saurabh Samaddar")
    y -= 34
    c.setFont("Helvetica", 15)
    c.drawString(left, y, "H.O.D. (Training & Development)")

    c.save()
    return pdf_path


def list_generated_pdfs() -> list[dict]:
    pdfs = []
    for file_path in PDF_DIR.glob("*.pdf"):
        stat = file_path.stat()
        modified_dt = datetime.fromtimestamp(stat.st_mtime)
        pdfs.append(
            {
                "name": file_path.name,
                "size_kb": round(stat.st_size / 1024, 1),
                "modified": modified_dt.strftime("%d-%m-%Y %I:%M %p"),
                "modified_ts": stat.st_mtime,
            }
        )
    pdfs.sort(key=lambda item: item["modified_ts"], reverse=True)
    return pdfs


@app.route("/", methods=["GET"])
def index():
    return render_template("form.html")


@app.route("/teacher", methods=["GET"])
def teacher_dashboard():
    records = ensure_record_ids(load_submission_records())
    all_records = records
    selected_month = request.args.get("month", "").strip()
    valid_month = ""
    if selected_month:
        try:
            valid_month = datetime.strptime(selected_month, "%Y-%m").strftime("%Y-%m")
        except ValueError:
            valid_month = ""

    month_count: dict[str, int] = {}
    for row in all_records:
        key = get_record_month_key(row)
        if not key:
            continue
        month_count[key] = month_count.get(key, 0) + 1

    month_options = sorted(month_count.keys(), reverse=True)
    month_summaries = [
        {"value": key, "label": format_month_label(key), "count": month_count[key]}
        for key in month_options
    ]

    records = filter_records_by_month(all_records, valid_month)

    return render_template(
        "teacher.html",
        records=records,
        selected_month=valid_month,
        selected_month_label=(format_month_label(valid_month) if valid_month else "All Months"),
        month_options=month_options,
        month_summaries=month_summaries,
        filtered_count=len(records),
        total_count=len(all_records),
    )


@app.route("/teacher/export", methods=["GET"])
def export_teacher_records():
    records = ensure_record_ids(load_submission_records())
    selected_month = request.args.get("month", "").strip()
    valid_month = ""
    if selected_month:
        try:
            valid_month = datetime.strptime(selected_month, "%Y-%m").strftime("%Y-%m")
        except ValueError:
            valid_month = ""
    records = filter_records_by_month(records, valid_month)

    csv_data = build_csv_text(records)
    filename_suffix = valid_month if valid_month else datetime.now().strftime("%Y%m%d_%H%M%S")
    response = Response(csv_data, mimetype="text/csv")
    response.headers["Content-Disposition"] = f"attachment; filename=teacher_records_{filename_suffix}.csv"
    return response


@app.route("/teacher/download/<path:filename>", methods=["GET"])
def download_pdf(filename: str):
    safe_name = secure_filename(filename)
    file_path = PDF_DIR / safe_name
    if not file_path.exists():
        abort(404)
    return send_from_directory(PDF_DIR, safe_name, as_attachment=True)


@app.route("/teacher/download-upload/<path:filename>", methods=["GET"])
def download_uploaded_file(filename: str):
    safe_name = secure_filename(filename)
    file_path = UPLOAD_DIR / safe_name
    if not file_path.exists():
        abort(404)
    return send_from_directory(UPLOAD_DIR, safe_name, as_attachment=True)


@app.route("/teacher/share-package", methods=["POST"])
def build_share_package():
    payload = request.get_json(silent=True) or {}
    raw_files = payload.get("files", [])
    if not isinstance(raw_files, list):
        return Response("Invalid payload", status=400)

    safe_files: list[str] = []
    seen = set()
    for name in raw_files:
        if not isinstance(name, str):
            continue
        cleaned = secure_filename(name.strip())
        if not cleaned or not cleaned.lower().endswith(".pdf"):
            continue
        if cleaned in seen:
            continue
        path = PDF_DIR / cleaned
        if path.exists():
            seen.add(cleaned)
            safe_files.append(cleaned)

    if not safe_files:
        return Response("No valid PDF files found", status=400)

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for filename in safe_files:
            zf.write(PDF_DIR / filename, arcname=filename)
    zip_buffer.seek(0)

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    response = Response(zip_buffer.getvalue(), mimetype="application/zip")
    response.headers["Content-Disposition"] = f'attachment; filename="selected_pdfs_{stamp}.zip"'
    return response


@app.route("/submit", methods=["POST"])
def submit_form():
    course_value = request.form.get("course", "")
    if course_value == "Other":
        course_value = request.form.get("course_other", "").strip() or "Other"

    form_data = {
        "letterhead_no": request.form.get("letterhead_no", "").strip(),
        "student_name": request.form.get("student_name", "").strip(),
        "roll_no": request.form.get("roll_no", "").strip(),
        "phone": request.form.get("phone", "").strip(),
        "email": request.form.get("email", "").strip(),
        "course": course_value,
        "year": request.form.get("year", "").strip(),
        "date_from": request.form.get("date_from", "").strip(),
        "date_to": request.form.get("date_to", "").strip(),
        "reporting_manager": request.form.get("reporting_manager", "").strip(),
        "company_name": request.form.get("company_name", "").strip(),
    }

    for field in ["recommendation_letter"]:
        file = request.files.get(field)
        if not file or file.filename == "":
            return render_template("success.html", success=False, message=f"{field} is required.")
        if not allowed_file(file.filename):
            return render_template("success.html", success=False, message=f"Invalid file type for {field}.")
        safe_name = secure_filename(f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{file.filename}")
        target = UPLOAD_DIR / safe_name
        file.save(target)

    generate_pdf(form_data)

    return render_template(
        "success.html",
        success=True,
        message="Form submitted successfully.",
        teacher_url=url_for("teacher_dashboard"),
    )


@app.route("/noc", methods=["POST"])
def noc_form():
    course_value = request.form.get("course", "").strip()
    if course_value == "Other":
        course_value = request.form.get("course_other", "").strip() or "Other"

    recommendation_file = request.files.get("recommendation_letter")
    if not recommendation_file or recommendation_file.filename == "":
        return render_template("success.html", success=False, message="student image is required.")
    if not allowed_image_file(recommendation_file.filename):
        return render_template("success.html", success=False, message="Invalid file type for student image.")
    recommendation_file.stream.seek(0, 2)
    student_image_size = recommendation_file.stream.tell()
    recommendation_file.stream.seek(0)
    if not (MIN_STUDENT_IMAGE_BYTES <= student_image_size <= MAX_STUDENT_IMAGE_BYTES):
        return render_template("success.html", success=False, message="student image must be between 200KB and 500KB.")

    recommendation_safe_name = secure_filename(
        f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{recommendation_file.filename}"
    )
    recommendation_target = UPLOAD_DIR / recommendation_safe_name
    recommendation_file.save(recommendation_target)

    form_data = {
        "letterhead_no": request.form.get("letterhead_no", "").strip(),
        "student_name": request.form.get("student_name", "").strip(),
        "roll_no": request.form.get("roll_no", "").strip(),
        "phone": request.form.get("phone", "").strip(),
        "email": request.form.get("email", "").strip(),
        "course": course_value,
        "year": request.form.get("year", "").strip(),
        "date_from": request.form.get("date_from", "").strip(),
        "date_to": request.form.get("date_to", "").strip(),
        "reporting_manager": request.form.get("reporting_manager", "").strip(),
        "company_name": request.form.get("company_name", "").strip(),
        "recommendation_letter": recommendation_safe_name,
        "student_image_size_kb": round(student_image_size / 1024, 1),
    }
    return render_template(
        "noc_form.html",
        form_data=form_data,
        now_date=datetime.now().strftime("%d-%m-%Y"),
    )


@app.route("/submit_noc", methods=["POST"])
def submit_noc():
    form_data = {
        "letterhead_no": request.form.get("letterhead_no", "").strip(),
        "student_name": request.form.get("student_name", "").strip(),
        "roll_no": request.form.get("roll_no", "").strip(),
        "phone": request.form.get("phone", "").strip(),
        "email": request.form.get("email", "").strip(),
        "course": request.form.get("course", "").strip(),
        "year": request.form.get("year", "").strip(),
        "date_from": request.form.get("date_from", "").strip(),
        "date_to": request.form.get("date_to", "").strip(),
        "reporting_manager": request.form.get("reporting_manager", "").strip(),
        "company_name": request.form.get("company_name", "").strip(),
    }

    recommendation_letter = request.form.get("recommendation_letter", "").strip()
    student_image_size_kb = request.form.get("student_image_size_kb", "").strip()
    internship_pdf = generate_pdf(form_data)
    noc_pdf = generate_noc_pdf(form_data)

    append_submission_record(
        {
            "letterhead_no": form_data["letterhead_no"],
            "student_name": form_data["student_name"],
            "student_image_size_kb": student_image_size_kb,
            "recommendation_letter": recommendation_letter,
            "date": datetime.now().strftime("%d-%m-%Y %I:%M %p"),
            "internship_pdf": internship_pdf.name,
            "noc_pdf": noc_pdf.name,
            "created_ts": datetime.now().timestamp(),
        }
    )

    return render_template(
        "success.html",
        success=True,
        message="Form submitted successfully. Internship and NOC PDFs generated.",
        teacher_url=url_for("teacher_dashboard"),
    )


# if __name__ == "__main__":
#     app.run(debug=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)

