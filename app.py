from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from datetime import datetime
import sqlite3
import os
import json

from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

app = Flask(__name__)
app.secret_key = "change-this-secret"

UPLOAD_FOLDER = "static"
# Use the actual template present in the project
TEMPLATE_PATH = "word_templates/college_letterhead.docx"
GENERATED_FOLDER = "generated_reports"


# ---------------- DATABASE ----------------

def init_db():
    conn = sqlite3.connect("database.db")
    c = conn.cursor()

    # Users table (with email)
    c.execute(
        """
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE,
            email TEXT,
            password TEXT
        )
        """
    )

    # Ensure email column exists on old databases
    c.execute("PRAGMA table_info(users)")
    user_cols = [row[1] for row in c.fetchall()]
    if "email" not in user_cols:
        c.execute("ALTER TABLE users ADD COLUMN email TEXT")

    # Events table – stores all event metadata
    c.execute(
        """
        CREATE TABLE IF NOT EXISTS events (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT,
            date TEXT,
            venue TEXT,
            department TEXT,
            description TEXT,
            event_photo TEXT,
            academic_year TEXT,
            resource_person TEXT,
            resource_designation TEXT,
            resource_organization TEXT,
            event_time TEXT,
            event_type TEXT,
            permission_letter TEXT,
            invitation_letter TEXT,
            notice_letter TEXT,
            appreciation_letter TEXT,
            event_photos TEXT,
            attendance_photo TEXT,
            outcome_1 TEXT,
            outcome_2 TEXT,
            outcome_3 TEXT,
            feedback_data TEXT
        )
        """
    )

    # Add any missing columns for existing databases
    c.execute("PRAGMA table_info(events)")
    event_cols = [row[1] for row in c.fetchall()]
    new_event_columns = [
        ("event_photo", "TEXT"),
        ("academic_year", "TEXT"),
        ("resource_person", "TEXT"),
        ("resource_designation", "TEXT"),
        ("resource_organization", "TEXT"),
        ("event_time", "TEXT"),
        ("event_type", "TEXT"),
        ("permission_letter", "TEXT"),
        ("invitation_letter", "TEXT"),
        ("notice_letter", "TEXT"),
        ("appreciation_letter", "TEXT"),
        ("event_photos", "TEXT"),
        ("attendance_photo", "TEXT"),
        ("outcome_1", "TEXT"),
        ("outcome_2", "TEXT"),
        ("outcome_3", "TEXT"),
        ("feedback_data", "TEXT"),
    ]
    for col_name, col_type in new_event_columns:
        if col_name not in event_cols:
            c.execute(f"ALTER TABLE events ADD COLUMN {col_name} {col_type}")

    # Reports table – stores generated report info
    c.execute(
        """
        CREATE TABLE IF NOT EXISTS reports (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            event_id INTEGER,
            file_path TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        """
    )

    conn.commit()
    conn.close()


def get_db():
    conn = sqlite3.connect("database.db")
    conn.row_factory = sqlite3.Row
    return conn


# --------------- FILE HELPERS ---------------

def _save_file(file_storage, subfolder, allow_pdf=False):
    """Save an uploaded file and return its relative path, or None."""
    if not file_storage or file_storage.filename == "":
        return None

    filename = secure_filename(file_storage.filename)
    ext = os.path.splitext(filename)[1].lower()
    image_exts = {".jpg", ".jpeg", ".png", ".gif"}
    allowed_exts = image_exts | ({".pdf"} if allow_pdf else set())
    if ext not in allowed_exts:
        return None

    os.makedirs(os.path.join(UPLOAD_FOLDER, subfolder), exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    filename = f"{timestamp}_{filename}"
    rel_path = os.path.join(UPLOAD_FOLDER, subfolder, filename)
    abs_path = os.path.join(rel_path)
    file_storage.save(abs_path)
    return rel_path.replace("\\", "/")


# ---------------- AUTH ----------------

def login_required(f):
    from functools import wraps
    @wraps(f)
    def wrap(*args, **kwargs):
        if "user_id" not in session:
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return wrap


# ---------------- DOCX HELPERS ----------------

def replace_placeholders(doc, data):
    """Replace simple {{...}} placeholders in paragraphs and table cells."""

    def _replace_in_paragraph(paragraph):
        if not paragraph.text:
            return
        for k, v in data.items():
            if k in paragraph.text:
                paragraph.text = paragraph.text.replace(k, v or "")

    # Top‑level paragraphs
    for p in doc.paragraphs:
        _replace_in_paragraph(p)

    # Paragraphs inside tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    _replace_in_paragraph(p)


def insert_full_page_image(doc, marker, path):
    if not path or not os.path.exists(path):
        return
    for p in doc.paragraphs:
        if marker in p.text:
            p.clear()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run().add_picture(path, width=Inches(6.5))
            doc.add_page_break()
            return


def insert_event_photos(doc, marker, photos):
    if not photos:
        return
    for p in doc.paragraphs:
        if marker in p.text:
            p.clear()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for img in photos:
                if os.path.exists(img):
                    p.add_run().add_picture(img, width=Inches(5))
                    p.add_run().add_break()
            return


def insert_attendance(doc, marker, path):
    if not path or not os.path.exists(path):
        return
    for p in doc.paragraphs:
        if marker in p.text:
            p.clear()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            ext = os.path.splitext(path)[1].lower()
            if ext in [".jpg", ".jpeg", ".png"]:
                p.add_run().add_picture(path, width=Inches(6.5))
            else:
                p.add_run("Attendance attached as scanned document")
            return


def insert_feedback_table(doc, marker, feedback):
    if not feedback or len(feedback) < 2:
        return
    feedback = feedback[:10]
    for p in doc.paragraphs:
        if marker in p.text:
            p.clear()
            table = doc.add_table(rows=1, cols=3)
            # Use Table Grid if it exists; otherwise keep default to avoid KeyError
            try:
                table.style = "Table Grid"
            except KeyError:
                pass
            table.rows[0].cells[0].text = "Name"
            table.rows[0].cells[1].text = "Rating"
            table.rows[0].cells[2].text = "Comment"
            for fb in feedback:
                row = table.add_row().cells
                row[0].text = fb.get("name", "")
                row[1].text = fb.get("rating", "")
                row[2].text = fb.get("comment", "")
            return


# ---------------- ROUTES ----------------

@app.route("/")
def home():
    return redirect(url_for("login"))


@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        username = request.form["username"]
        email = request.form.get("email", "")
        password = request.form["password"]

        conn = get_db()
        existing = conn.execute(
            "SELECT id FROM users WHERE username=?", (username,)
        ).fetchone()
        if existing:
            conn.close()
            flash("Username already exists")
            return render_template("register.html")

        hashed = generate_password_hash(password)
        conn.execute(
            "INSERT INTO users (username, email, password) VALUES (?, ?, ?)",
            (username, email, hashed),
        )
        conn.commit()
        conn.close()
        flash("Registration successful. Please login.")
        return redirect(url_for("login"))

    return render_template("register.html")


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form["username"]
        password = request.form["password"]

        conn = get_db()
        user = conn.execute(
            "SELECT * FROM users WHERE username=?", (username,)
        ).fetchone()
        conn.close()

        if user and check_password_hash(user["password"], password):
            session["user_id"] = user["id"]
            session["username"] = user["username"]
            return redirect(url_for("events"))
        flash("Invalid login")

    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/events")
@login_required
def events():
    conn = get_db()
    events = conn.execute("SELECT * FROM events ORDER BY id DESC").fetchall()
    conn.close()
    return render_template("events.html", events=events)


@app.route("/add_event", methods=["GET", "POST"])
@login_required
def add_event():
    if request.method == "POST":
        title = request.form["title"]
        date = request.form["date"]
        venue = request.form["venue"]
        department = request.form["department"]
        description = request.form.get("description", "")
        academic_year = request.form.get("academic_year", "")
        resource_person = request.form.get("resource_person", "")
        resource_designation = request.form.get("resource_designation", "")
        resource_organization = request.form.get("resource_organization", "")
        event_time = request.form.get("event_time", "")
        event_type = request.form.get("event_type", "")
        outcome_1 = request.form.get("outcome_1", "")
        outcome_2 = request.form.get("outcome_2", "")
        outcome_3 = request.form.get("outcome_3", "")

        # Event photos (multiple)
        event_photos_paths = []
        if "event_photos" in request.files:
            for f in request.files.getlist("event_photos"):
                saved = _save_file(f, "event_photos", allow_pdf=False)
                if saved:
                    event_photos_paths.append(saved)

        event_photo_cover = event_photos_paths[0] if event_photos_paths else None

        # Scanned documents and attendance
        permission_letter = _save_file(
            request.files.get("permission_letter"), "attendance_photos", allow_pdf=True
        )
        invitation_letter = _save_file(
            request.files.get("invitation_letter"), "attendance_photos", allow_pdf=True
        )
        notice_letter = _save_file(
            request.files.get("notice_letter"), "attendance_photos", allow_pdf=True
        )
        appreciation_letter = _save_file(
            request.files.get("appreciation_letter"), "attendance_photos", allow_pdf=True
        )
        attendance_photo = _save_file(
            request.files.get("attendance_photo"), "attendance_photos", allow_pdf=True
        )

        # Feedback
        feedback_names = request.form.getlist("feedback_name[]")
        feedback_ratings = request.form.getlist("feedback_rating[]")
        feedback_comments = request.form.getlist("feedback_comment[]")
        feedback_data = []
        for i in range(min(len(feedback_names), len(feedback_ratings), len(feedback_comments))):
            if (
                feedback_names[i].strip()
                or feedback_ratings[i].strip()
                or feedback_comments[i].strip()
            ):
                feedback_data.append(
                    {
                        "name": feedback_names[i],
                        "rating": feedback_ratings[i],
                        "comment": feedback_comments[i],
                    }
                )

        conn = get_db()
        conn.execute(
            """
            INSERT INTO events (
                title, date, venue, department, description,
                event_photo, academic_year, resource_person,
                resource_designation, resource_organization,
                event_time, event_type,
                permission_letter, invitation_letter, notice_letter,
                appreciation_letter, event_photos, attendance_photo,
                outcome_1, outcome_2, outcome_3, feedback_data
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                title,
                date,
                venue,
                department,
                description,
                event_photo_cover,
                academic_year,
                resource_person,
                resource_designation,
                resource_organization,
                event_time,
                event_type,
                permission_letter,
                invitation_letter,
                notice_letter,
                appreciation_letter,
                json.dumps(event_photos_paths) if event_photos_paths else None,
                attendance_photo,
                outcome_1,
                outcome_2,
                outcome_3,
                json.dumps(feedback_data) if feedback_data else None,
            ),
        )
        conn.commit()
        conn.close()

        flash("Event added successfully.")
        return redirect(url_for("events"))

    return render_template("add_event.html")


@app.route("/generate_report/<int:event_id>")
@login_required
def generate_report(event_id):

    conn = get_db()
    event = conn.execute("SELECT * FROM events WHERE id=?", (event_id,)).fetchone()
    conn.close()

    if not event:
        flash("Event not found")
        return redirect(url_for("events"))

    doc = Document(TEMPLATE_PATH)

    replace_placeholders(doc, {
        "{{academic_year}}": event["academic_year"],
        "{{date}}": event["date"],
        "{{event_name}}": event["title"],
        "{{event_type}}": event["event_type"],
        "{{event_date}}": event["date"],
        "{{event_time}}": event["event_time"],
        "{{venue}}": event["venue"],
        "{{department}}": event["department"],
        "{{resource_person}}": event["resource_person"],
        "{{resource_designation}}": event["resource_designation"],
        "{{resource_organization}}": event["resource_organization"],
        # support both {{event_description}} and {{event description}} in template
        "{{event_description}}": event["description"],
        "{{event description}}": event["description"],
        "{{outcome_1}}": event["outcome_1"],
        "{{outcome_2}}": event["outcome_2"],
        "{{outcome_3}}": event["outcome_3"],
    })

    insert_full_page_image(doc, "<<IMAGE_PERMISSION>>", event["permission_letter"])
    insert_full_page_image(doc, "<<IMAGE_INVITATION>>", event["invitation_letter"])
    insert_full_page_image(doc, "<<IMAGE_NOTICE>>", event["notice_letter"])
    insert_full_page_image(doc, "<<IMAGE_APPRECIATION>>", event["appreciation_letter"])

    photos = json.loads(event["event_photos"]) if event["event_photos"] else []
    insert_event_photos(doc, "<<EVENT_PHOTOS>>", photos)

    insert_attendance(doc, "<<ATTENDANCE_FILE>>", event["attendance_photo"])

    feedback = json.loads(event["feedback_data"]) if event["feedback_data"] else []
    insert_feedback_table(doc, "<<FEEDBACK_TABLE>>", feedback)

    filename = event["title"].replace(" ", "_") + ".docx"
    path = os.path.join(GENERATED_FOLDER, filename)
    doc.save(path)

    # Store report metadata
    conn = get_db()
    conn.execute(
        "INSERT INTO reports (event_id, file_path) VALUES (?, ?)", (event_id, path)
    )
    conn.commit()
    conn.close()

    return send_file(path, as_attachment=True)


@app.route("/reports")
@login_required
def reports():
    conn = get_db()
    rows = conn.execute(
        """
        SELECT r.id,
               r.file_path,
               r.created_at,
               e.title AS event_title,
               e.date  AS event_date
        FROM reports r
        JOIN events e ON r.event_id = e.id
        ORDER BY r.created_at DESC
        """
    ).fetchall()
    conn.close()
    # Shape data to match template expectations
    reports_data = []
    for r in rows:
        reports_data.append(
            {
                "id": r["id"],
                "file_path": r["file_path"],
                "created_at": r["created_at"],
                "event_title": r["event_title"],
                "event_date": r["event_date"],
                "created_by_name": "",  # no separate user table linkage here
            }
        )
    return render_template("reports.html", reports=reports_data)


@app.route("/download_report/<int:report_id>")
@login_required
def download_report(report_id):
    conn = get_db()
    row = conn.execute(
        "SELECT file_path FROM reports WHERE id=?", (report_id,)
    ).fetchone()
    conn.close()
    if not row or not row["file_path"] or not os.path.exists(row["file_path"]):
        flash("Report file not found.")
        return redirect(url_for("reports"))
    return send_file(row["file_path"], as_attachment=True)


# ---------------- START ----------------

if __name__ == "__main__":
    os.makedirs(GENERATED_FOLDER, exist_ok=True)
    init_db()
    app.run(debug=True)
