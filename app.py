from flask import Flask, render_template, request, redirect, session, url_for, send_file, jsonify
import cv2
import os
import csv
import mediapipe as mp
import face_recognition
import pickle
import smtplib
import base64
import numpy as np
from email.mime.text import MIMEText
from datetime import datetime, date
from io import BytesIO, StringIO
from functools import wraps
from openpyxl import Workbook
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas as pdf_canvas
import psycopg2
from psycopg2.extras import RealDictCursor
from psycopg2 import IntegrityError
from supabase import create_client, Client

# -----------------------------
# CONFIG
# -----------------------------
app = Flask(__name__)
app.secret_key = "secret123"

DATABASE_URL = "postgresql://postgres.hegohwrmezrtoyujkdov:Sudheer%4012345@aws-1-ap-northeast-2.pooler.supabase.com:5432/postgres"

SUPABASE_URL = "https://hegohwrmezrtoyujkdov.supabase.co"
SUPABASE_KEY = "YOUR_SUPABASE_KEY"

supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

FACES_DIR = "static/faces"
UNKNOWN_DIR = "unknown_faces"
ENCODINGS_FILE = "encodings.pkl"

SENDER_EMAIL = "your_email@gmail.com"
SENDER_PASSWORD = "your_app_password"

UNKNOWN_ENCODINGS = []
LAST_UNKNOWN_SAVE_TIME = {}

UNKNOWN_COOLDOWN = 30
KNOWN_FACE_TOLERANCE = 0.55
UNKNOWN_FACE_TOLERANCE = 0.43

ATTENDANCE_FRAME_SCALE = 0.25
ATTENDANCE_PROCESS_EVERY_N_FRAMES = 3
LABEL_HOLD_FRAMES = 12

os.makedirs(FACES_DIR, exist_ok=True)
os.makedirs(UNKNOWN_DIR, exist_ok=True)

# -----------------------------
# DATABASE CONNECTION
# -----------------------------
def get_db_connection():
    return psycopg2.connect(DATABASE_URL, cursor_factory=RealDictCursor)

# -----------------------------
# MEDIAPIPE
# -----------------------------
mp_face_detection = mp.solutions.face_detection
face_detector = mp_face_detection.FaceDetection(
    model_selection=0,
    min_detection_confidence=0.6
)

# -----------------------------
# DATABASE SETUP
# -----------------------------
def create_database():
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS students (
        student_id TEXT PRIMARY KEY,
        name TEXT NOT NULL,
        department TEXT NOT NULL,
        password TEXT NOT NULL,
        parent_name TEXT,
        parent_phone TEXT,
        parent_email TEXT,
        photo TEXT,
        created_at TIMESTAMP DEFAULT NOW()
    )
    """)

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS attendance (
        id SERIAL PRIMARY KEY,
        student_id TEXT NOT NULL,
        name TEXT NOT NULL,
        department TEXT NOT NULL,
        attendance_date DATE NOT NULL,
        attendance_time TIME NOT NULL,
        status TEXT DEFAULT 'Present',
        created_at TIMESTAMP DEFAULT NOW(),
        UNIQUE(student_id, attendance_date)
    )
    """)

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS admins (
        username TEXT PRIMARY KEY,
        password TEXT NOT NULL
    )
    """)

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS absence_alerts (
        id SERIAL PRIMARY KEY,
        student_id TEXT NOT NULL,
        alert_date DATE NOT NULL,
        UNIQUE(student_id, alert_date)
    )
    """)

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS unknown_faces (
        id SERIAL PRIMARY KEY,
        image_name TEXT,
        image_url TEXT,
        captured_at TIMESTAMP DEFAULT NOW()
    )
    """)

    cursor.execute("SELECT * FROM admins WHERE username = %s", ("admin",))
    admin = cursor.fetchone()

    if not admin:
        cursor.execute(
            "INSERT INTO admins (username, password) VALUES (%s, %s)",
            ("admin", "1234")
        )

    conn.commit()
    cursor.close()
    conn.close()

create_database()

# -----------------------------
# LOGIN REQUIRED DECORATORS
# -----------------------------
def admin_login_required(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if "admin" not in session:
            return redirect(url_for("admin_login"))
        return f(*args, **kwargs)
    return wrapper

def student_login_required(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        sid = kwargs.get("sid")
        if "student_id" not in session:
            return redirect(url_for("student_login"))
        if sid and session.get("student_id") != sid:
            return "Unauthorized access"
        return f(*args, **kwargs)
    return wrapper

# -----------------------------
# CAMERA HELPER
# -----------------------------
def open_camera(camera_index=0):
    cap = cv2.VideoCapture(camera_index, cv2.CAP_DSHOW)

    if not cap.isOpened():
        cap.release()
        cap = cv2.VideoCapture(camera_index)

    if not cap.isOpened():
        return None

    cap.set(cv2.CAP_PROP_FRAME_WIDTH, 640)
    cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 480)

    try:
        cap.set(cv2.CAP_PROP_BUFFERSIZE, 1)
    except Exception:
        pass

    return cap

# -----------------------------
# EMAIL FUNCTION
# -----------------------------
def send_absent_email(parent_email, student_name):
    if not parent_email or SENDER_EMAIL == "your_email@gmail.com" or SENDER_PASSWORD == "your_app_password":
        print("Email skipped: sender email or app password not configured")
        return

    subject = "Student Absence Alert"
    message = f"""Dear Parent,

Your child {student_name} has been absent for two consecutive attendance days.

Please contact the college if needed.

Regards,
VisionDB Attendance System
"""

    msg = MIMEText(message)
    msg["Subject"] = subject
    msg["From"] = SENDER_EMAIL
    msg["To"] = parent_email

    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.sendmail(SENDER_EMAIL, parent_email, msg.as_string())
        server.quit()
        print("Email sent successfully to", parent_email)
    except Exception as e:
        print("Email error:", e)

# -----------------------------
# DATE HELPERS
# -----------------------------
def get_today_date():
    return date.today()

def get_today_date_str():
    return get_today_date().strftime("%Y-%m-%d")

def parse_date_filter(date_str):
    try:
        return datetime.strptime(date_str, "%Y-%m-%d").date()
    except Exception:
        return None

# -----------------------------
# ABSENCE CHECK
# -----------------------------
def check_absent_students():
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT student_id, name, parent_email FROM students ORDER BY student_id")
    students = cursor.fetchall()

    today = get_today_date()

    for student in students:
        student_id = student["student_id"]
        name = student["name"]
        parent_email = student["parent_email"]

        cursor.execute("""
            SELECT attendance_date, status
            FROM attendance
            WHERE student_id = %s
            ORDER BY attendance_date DESC
            LIMIT 2
        """, (student_id,))
        last_two = cursor.fetchall()

        if len(last_two) == 2:
            if last_two[0]["status"] == "Absent" and last_two[1]["status"] == "Absent":
                cursor.execute(
                    "SELECT 1 FROM absence_alerts WHERE student_id = %s AND alert_date = %s",
                    (student_id, today)
                )
                already_sent = cursor.fetchone()

                if not already_sent:
                    send_absent_email(parent_email, name)
                    try:
                        cursor.execute(
                            "INSERT INTO absence_alerts (student_id, alert_date) VALUES (%s, %s)",
                            (student_id, today)
                        )
                        conn.commit()
                    except IntegrityError:
                        conn.rollback()

    cursor.close()
    conn.close()

# -----------------------------
# STUDENT HELPERS
# -----------------------------
def student_id_exists(student_id):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT 1 FROM students WHERE student_id = %s", (student_id,))
    result = cursor.fetchone()
    cursor.close()
    conn.close()
    return result is not None

def get_student_doc(student_id):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM students WHERE student_id = %s", (student_id,))
    row = cursor.fetchone()
    cursor.close()
    conn.close()
    return dict(row) if row else None

def get_student_details(student_id):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT name, department FROM students WHERE student_id = %s", (student_id,))
    row = cursor.fetchone()
    cursor.close()
    conn.close()
    if row:
        return row["name"], row["department"]
    return None

def get_all_student_map():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT student_id, name, department FROM students")
    rows = cursor.fetchall()
    cursor.close()
    conn.close()

    result = {}
    for row in rows:
        result[row["student_id"]] = {
            "name": row["name"],
            "department": row["department"]
        }
    return result

def get_students_for_template():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT student_id, name, department, password, parent_name, parent_phone, parent_email, photo
        FROM students
        ORDER BY student_id ASC
    """)
    rows = cursor.fetchall()
    cursor.close()
    conn.close()

    return [
        (
            row["student_id"],
            row["name"],
            row["department"],
            row["password"],
            row["parent_name"],
            row["parent_phone"],
            row["parent_email"],
            row["photo"]
        )
        for row in rows
    ]

def get_student_tuple(student_id):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT student_id, name, department, password, parent_name, parent_phone, parent_email, photo
        FROM students
        WHERE student_id = %s
    """, (student_id,))
    row = cursor.fetchone()
    cursor.close()
    conn.close()

    if not row:
        return None

    return (
        row["student_id"],
        row["name"],
        row["department"],
        row["password"],
        row["parent_name"],
        row["parent_phone"],
        row["parent_email"],
        row["photo"]
    )

# -----------------------------
# FACE HELPERS
# -----------------------------
def load_known_faces():
    known_encodings = []
    known_ids = []

    if not os.path.exists(FACES_DIR):
        return known_encodings, known_ids

    for file in os.listdir(FACES_DIR):
        path = os.path.join(FACES_DIR, file)

        if not os.path.isfile(path):
            continue

        if not file.lower().endswith((".jpg", ".jpeg", ".png")):
            continue

        try:
            image = face_recognition.load_image_file(path)
            encodings = face_recognition.face_encodings(image)

            if len(encodings) == 1:
                known_encodings.append(encodings[0])
                known_ids.append(os.path.splitext(file)[0])
        except Exception:
            continue

    return known_encodings, known_ids

def generate_encodings():
    known_encodings, known_ids = load_known_faces()
    data = {"encodings": known_encodings, "ids": known_ids}

    with open(ENCODINGS_FILE, "wb") as f:
        pickle.dump(data, f)

def read_cached_encodings():
    if os.path.exists(ENCODINGS_FILE):
        try:
            with open(ENCODINGS_FILE, "rb") as f:
                data = pickle.load(f)
                return data.get("encodings", []), data.get("ids", [])
        except Exception:
            pass

    encodings, ids_ = load_known_faces()
    data = {"encodings": encodings, "ids": ids_}
    with open(ENCODINGS_FILE, "wb") as f:
        pickle.dump(data, f)
    return encodings, ids_

def is_duplicate_face(new_image_path):
    try:
        new_image = face_recognition.load_image_file(new_image_path)
        new_encoding_list = face_recognition.face_encodings(new_image)

        if len(new_encoding_list) == 0:
            return False

        new_encoding = new_encoding_list[0]

        for file in os.listdir(FACES_DIR):
            old_path = os.path.join(FACES_DIR, file)

            if not os.path.isfile(old_path):
                continue

            if os.path.abspath(old_path) == os.path.abspath(new_image_path):
                continue

            if not file.lower().endswith((".jpg", ".jpeg", ".png")):
                continue

            old_image = face_recognition.load_image_file(old_path)
            old_encoding_list = face_recognition.face_encodings(old_image)

            if len(old_encoding_list) > 0:
                match = face_recognition.compare_faces(
                    [old_encoding_list[0]], new_encoding, tolerance=0.45
                )
                if True in match:
                    return True

        return False
    except Exception:
        return False

def should_save_unknown_face(face_encoding, cooldown=30, tolerance=0.43):
    global UNKNOWN_ENCODINGS, LAST_UNKNOWN_SAVE_TIME

    current_time = datetime.now().timestamp()

    if len(UNKNOWN_ENCODINGS) == 0:
        UNKNOWN_ENCODINGS.append(face_encoding)
        LAST_UNKNOWN_SAVE_TIME[0] = current_time
        return True

    face_distances = face_recognition.face_distance(UNKNOWN_ENCODINGS, face_encoding)

    if len(face_distances) == 0:
        UNKNOWN_ENCODINGS.append(face_encoding)
        new_index = len(UNKNOWN_ENCODINGS) - 1
        LAST_UNKNOWN_SAVE_TIME[new_index] = current_time
        return True

    best_match_index = int(np.argmin(face_distances))
    best_distance = float(face_distances[best_match_index])

    if best_distance < tolerance:
        last_time = LAST_UNKNOWN_SAVE_TIME.get(best_match_index, 0)
        if current_time - last_time >= cooldown:
            LAST_UNKNOWN_SAVE_TIME[best_match_index] = current_time
            return True
        return False

    UNKNOWN_ENCODINGS.append(face_encoding)
    new_index = len(UNKNOWN_ENCODINGS) - 1
    LAST_UNKNOWN_SAVE_TIME[new_index] = current_time
    return True

def save_unknown_face(frame, top, right, bottom, left):
    top = max(0, top)
    right = max(0, right)
    bottom = max(0, bottom)
    left = max(0, left)

    face_img = frame[top:bottom, left:right]

    if face_img.size == 0:
        print("Unknown face image is empty")
        return ""

    current_time = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    filename = f"unknown_{current_time}.jpg"
    local_path = os.path.join(UNKNOWN_DIR, filename)

    saved = cv2.imwrite(local_path, face_img)
    print("Local unknown save:", saved, local_path)

    public_url = ""

    try:
        with open(local_path, "rb") as f:
            supabase.storage.from_("unknown-faces").upload(
                path=filename,
                file=f,
                file_options={"content-type": "image/jpeg"}
            )

        public_url = supabase.storage.from_("unknown-faces").get_public_url(filename)
        if isinstance(public_url, dict):
            public_url = public_url.get("publicUrl", "")

        print("Supabase upload successful:", public_url)
    except Exception as e:
        print("Supabase upload error:", e)
        public_url = ""

    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO unknown_faces (image_name, image_url, captured_at)
            VALUES (%s, %s, %s)
        """, (filename, public_url, datetime.now()))
        conn.commit()
        cursor.close()
        conn.close()
        print("Unknown face inserted into database")
    except Exception as e:
        print("Database save error for unknown face:", e)

    return filename

def clamp_box(box, frame_shape):
    h, w = frame_shape[:2]
    top, right, bottom, left = box
    top = max(0, min(top, h - 1))
    right = max(0, min(right, w - 1))
    bottom = max(0, min(bottom, h - 1))
    left = max(0, min(left, w - 1))
    return (top, right, bottom, left)

def box_center(box):
    top, right, bottom, left = box
    return ((left + right) // 2, (top + bottom) // 2)

def smooth_box(old_box, new_box, alpha=0.65):
    if old_box is None:
        return new_box

    return (
        int(alpha * old_box[0] + (1 - alpha) * new_box[0]),
        int(alpha * old_box[1] + (1 - alpha) * new_box[1]),
        int(alpha * old_box[2] + (1 - alpha) * new_box[2]),
        int(alpha * old_box[3] + (1 - alpha) * new_box[3]),
    )

def match_previous_detection(new_box, previous_detections, max_distance=80):
    if not previous_detections:
        return None

    cx, cy = box_center(new_box)
    best_item = None
    best_dist = None

    for item in previous_detections:
        px, py = box_center(item["box"])
        dist = ((cx - px) ** 2 + (cy - py) ** 2) ** 0.5
        if dist <= max_distance and (best_dist is None or dist < best_dist):
            best_dist = dist
            best_item = item

    return best_item

def build_attendance_detections(frame, known_encodings, known_ids, student_map, marked_ids, db_cursor, today):
    small_frame = cv2.resize(frame, (0, 0), fx=ATTENDANCE_FRAME_SCALE, fy=ATTENDANCE_FRAME_SCALE)
    rgb_small_frame = cv2.cvtColor(small_frame, cv2.COLOR_BGR2RGB)

    face_locations = face_recognition.face_locations(rgb_small_frame)
    face_encodings = face_recognition.face_encodings(rgb_small_frame, face_locations)

    detections = []

    for (top, right, bottom, left), encoding in zip(face_locations, face_encodings):
        top = int(top / ATTENDANCE_FRAME_SCALE)
        right = int(right / ATTENDANCE_FRAME_SCALE)
        bottom = int(bottom / ATTENDANCE_FRAME_SCALE)
        left = int(left / ATTENDANCE_FRAME_SCALE)

        top, right, bottom, left = clamp_box((top, right, bottom, left), frame.shape)

        matched_student_id = None
        label = "Unknown"
        color = (0, 0, 255)
        save_text = ""

        if len(known_encodings) > 0:
            face_distances = face_recognition.face_distance(known_encodings, encoding)

            if len(face_distances) > 0:
                best_match_index = int(np.argmin(face_distances))
                best_distance = float(face_distances[best_match_index])

                if best_distance < KNOWN_FACE_TOLERANCE:
                    matched_student_id = known_ids[best_match_index]

        if matched_student_id:
            student_info = student_map.get(matched_student_id, {})
            student_name = student_info.get("name", matched_student_id)
            department = student_info.get("department", "Unknown")

            label = f"{student_name} ({matched_student_id})"
            color = (0, 180, 0)

            if matched_student_id not in marked_ids:
                db_cursor.execute(
                    "SELECT 1 FROM attendance WHERE student_id = %s AND attendance_date = %s",
                    (matched_student_id, today)
                )
                already_marked = db_cursor.fetchone()

                if not already_marked:
                    current_time = datetime.now().strftime("%H:%M:%S")
                    db_cursor.execute("""
                        INSERT INTO attendance (
                            student_id, name, department, attendance_date, attendance_time, status
                        )
                        VALUES (%s, %s, %s, %s, %s, %s)
                    """, (
                        matched_student_id,
                        student_name,
                        department,
                        today,
                        current_time,
                        "Present"
                    ))
                    save_text = "Attendance marked"

                marked_ids.add(matched_student_id)
            else:
                save_text = "Already marked in session"

        else:
            if should_save_unknown_face(
                encoding,
                cooldown=UNKNOWN_COOLDOWN,
                tolerance=UNKNOWN_FACE_TOLERANCE
            ):
                filename = save_unknown_face(frame, top, right, bottom, left)
                save_text = f"Saved unknown: {filename}" if filename else "Unknown detected"
            else:
                save_text = "Unknown detected"

        detections.append({
            "box": (top, right, bottom, left),
            "label": label,
            "color": color,
            "extra": save_text,
            "ttl": LABEL_HOLD_FRAMES
        })

    return detections

def stabilize_detections(new_detections, previous_detections):
    stabilized = []

    for det in new_detections:
        matched_prev = match_previous_detection(det["box"], previous_detections)
        if matched_prev and matched_prev["label"] == det["label"]:
            det["box"] = smooth_box(matched_prev["box"], det["box"])
        stabilized.append(det)

    return stabilized

def decay_detections(detections):
    kept = []
    for det in detections:
        det["ttl"] -= 1
        if det["ttl"] > 0:
            kept.append(det)
    return kept

def draw_detection_boxes(frame, detections):
    for det in detections:
        top, right, bottom, left = det["box"]
        label = det["label"]
        color = det["color"]
        extra = det.get("extra", "")

        cv2.rectangle(frame, (left, top), (right, bottom), color, 2)

        label_y1 = max(0, bottom - 24)
        cv2.rectangle(frame, (left, label_y1), (right, bottom), color, cv2.FILLED)
        cv2.putText(
            frame,
            label[:40],
            (left + 6, bottom - 7),
            cv2.FONT_HERSHEY_SIMPLEX,
            0.48,
            (255, 255, 255),
            1,
            cv2.LINE_AA
        )

        if extra:
            extra_y = max(20, top - 8)
            cv2.putText(
                frame,
                extra[:45],
                (left, extra_y),
                cv2.FONT_HERSHEY_SIMPLEX,
                0.45,
                color,
                1,
                cv2.LINE_AA
            )

def draw_attendance_header(frame, marked_ids, detections):
    unknown_count = sum(1 for d in detections if d["label"] == "Unknown")
    known_count = sum(1 for d in detections if d["label"] != "Unknown")

    cv2.putText(
        frame,
        "Press Q to close attendance camera",
        (20, 30),
        cv2.FONT_HERSHEY_SIMPLEX,
        0.7,
        (0, 255, 255),
        2,
        cv2.LINE_AA
    )

    cv2.putText(
        frame,
        f"Known on screen: {known_count} | Unknown on screen: {unknown_count}",
        (20, 60),
        cv2.FONT_HERSHEY_SIMPLEX,
        0.6,
        (255, 255, 255),
        2,
        cv2.LINE_AA
    )

    cv2.putText(
        frame,
        f"Attendance marked today in this session: {len(marked_ids)}",
        (20, 90),
        cv2.FONT_HERSHEY_SIMPLEX,
        0.6,
        (255, 255, 255),
        2,
        cv2.LINE_AA
    )

# -----------------------------
# ATTENDANCE HELPERS
# -----------------------------
def get_all_attendance_dates():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT attendance_date FROM attendance ORDER BY attendance_date")
    rows = cursor.fetchall()
    cursor.close()
    conn.close()
    return [row["attendance_date"].strftime("%Y-%m-%d") for row in rows]

def get_attendance_summary_for_student(student_id):
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT DISTINCT attendance_date
        FROM attendance
        ORDER BY attendance_date
    """)
    all_date_rows = cursor.fetchall()
    all_dates = [row["attendance_date"].strftime("%Y-%m-%d") for row in all_date_rows]

    cursor.execute("""
        SELECT attendance_date, status
        FROM attendance
        WHERE student_id = %s
        ORDER BY attendance_date
    """, (student_id,))
    rows = cursor.fetchall()

    cursor.close()
    conn.close()

    student_status_map = {}
    for row in rows:
        date_str = row["attendance_date"].strftime("%Y-%m-%d")
        student_status_map[date_str] = row["status"]

    present_dates = []
    absent_dates = []

    for date_str in all_dates:
        status = student_status_map.get(date_str)

        if status == "Present":
            present_dates.append(date_str)
        else:
            absent_dates.append(date_str)

    total_days = len(all_dates)
    present_days = len(present_dates)
    absent_days = len(absent_dates)
    percentage = round((present_days / total_days) * 100, 2) if total_days > 0 else 0

    return {
        "total_days": total_days,
        "present_days": present_days,
        "absent_days": absent_days,
        "percentage": percentage,
        "present_dates": present_dates,
        "absent_dates": absent_dates
    }

def mark_absent_students_for_today():
    conn = get_db_connection()
    cursor = conn.cursor()

    today = get_today_date()
    current_time = datetime.now().strftime("%H:%M:%S")

    cursor.execute("""
        SELECT student_id, name, department
        FROM students
    """)
    all_students = cursor.fetchall()

    cursor.execute("""
        SELECT student_id
        FROM attendance
        WHERE attendance_date = %s AND status = %s
    """, (today, "Present"))
    present_rows = cursor.fetchall()

    present_ids = {row["student_id"] for row in present_rows}

    for student in all_students:
        student_id = student["student_id"]
        name = student["name"]
        department = student["department"]

        if student_id not in present_ids:
            cursor.execute("""
                SELECT 1 FROM attendance
                WHERE student_id = %s AND attendance_date = %s
            """, (student_id, today))
            already_exists = cursor.fetchone()

            if not already_exists:
                cursor.execute("""
                    INSERT INTO attendance (
                        student_id, name, department, attendance_date, attendance_time, status
                    )
                    VALUES (%s, %s, %s, %s, %s, %s)
                """, (
                    student_id,
                    name,
                    department,
                    today,
                    current_time,
                    "Absent"
                ))

    conn.commit()
    cursor.close()
    conn.close()

def get_attendance_records(date_filter=None, student_id=None):
    conn = get_db_connection()
    cursor = conn.cursor()

    query = """
        SELECT student_id, name, department, attendance_date, attendance_time, status
        FROM attendance
        WHERE 1=1
    """
    params = []

    if date_filter:
        query += " AND attendance_date = %s"
        params.append(date_filter)

    if student_id:
        query += " AND student_id = %s"
        params.append(student_id)

    query += " ORDER BY attendance_date DESC, attendance_time ASC"

    cursor.execute(query, tuple(params))
    rows = cursor.fetchall()
    cursor.close()
    conn.close()

    result = []
    for row in rows:
        result.append({
            "id": row["student_id"],
            "name": row["name"],
            "department": row["department"],
            "date": row["attendance_date"].strftime("%Y-%m-%d"),
            "time": str(row["attendance_time"]),
            "status": row["status"]
        })
    return result

def get_today_attendance_data(search_id=""):
    return get_attendance_records(date_filter=get_today_date(), student_id=search_id if search_id else None)

def get_all_attendance_records():
    records = get_attendance_records()
    return [
        [row["id"], row["name"], row["department"], row["date"], row["time"], row["status"]]
        for row in records
    ]

def build_csv_output(records):
    output = StringIO()
    writer = csv.writer(output)
    writer.writerow(["Student ID", "Name", "Department", "Date", "Time", "Status"])

    for row in records:
        writer.writerow([
            row["id"],
            row["name"],
            row["department"],
            row["date"],
            row["time"],
            row["status"]
        ])

    byte_output = BytesIO()
    byte_output.write(output.getvalue().encode("utf-8"))
    byte_output.seek(0)
    return byte_output

# -----------------------------
# ADMIN LOGIN
# -----------------------------
@app.route("/", methods=["GET", "POST"])
def admin_login():
    if request.method == "POST":
        username = request.form["username"].strip()
        password = request.form["password"].strip()

        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute(
            "SELECT * FROM admins WHERE username = %s AND password = %s",
            (username, password)
        )
        admin = cursor.fetchone()
        cursor.close()
        conn.close()

        if admin:
            session.clear()
            session["admin"] = username
            return redirect(url_for("dashboard"))

        return "Invalid admin login"

    return render_template("login.html")

# -----------------------------
# ADMIN LOGOUT
# -----------------------------
@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("admin_login"))

# -----------------------------
# STUDENT LOGIN
# -----------------------------
@app.route("/student_login", methods=["GET", "POST"])
def student_login():
    if request.method == "POST":
        student_id = request.form["student_id"].strip()
        password = request.form["password"].strip()

        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute(
            "SELECT * FROM students WHERE student_id = %s AND password = %s",
            (student_id, password)
        )
        student = cursor.fetchone()
        cursor.close()
        conn.close()

        if student:
            session.clear()
            session["student_id"] = student_id
            return redirect(url_for("student_dashboard", sid=student_id))

        return "Invalid Student Login"

    return render_template("student_login.html")

# -----------------------------
# STUDENT LOGOUT
# -----------------------------
@app.route("/student_logout")
def student_logout():
    session.pop("student_id", None)
    return redirect(url_for("student_login"))

# -----------------------------
# STUDENT DASHBOARD
# -----------------------------
@app.route("/student_dashboard/<sid>")
@student_login_required
def student_dashboard(sid):
    student = get_student_doc(sid)

    if student:
        summary = get_attendance_summary_for_student(sid)

        return render_template(
            "student_dashboard.html",
            student_id=sid,
            name=student.get("name", ""),
            department=student.get("department", ""),
            summary=summary
        )

    return "Student not found"

# -----------------------------
# ADMIN DASHBOARD
# -----------------------------
@app.route("/dashboard")
@admin_login_required
def dashboard():
    today = get_today_date()

    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT COUNT(*) AS total FROM students")
    total_students = cursor.fetchone()["total"]

    cursor.execute(
        "SELECT COUNT(DISTINCT student_id) AS total FROM attendance WHERE attendance_date = %s AND status = %s",
        (today, "Present")
    )
    total_present = cursor.fetchone()["total"]

    cursor.close()
    conn.close()

    total_absent = max(total_students - total_present, 0)

    return render_template(
        "dashboard.html",
        students=total_students,
        present=total_present,
        absent=total_absent
    )

# -----------------------------
# VIEW ALL STUDENTS
# -----------------------------
@app.route("/students")
@admin_login_required
def view_students():
    students = get_students_for_template()
    return render_template("students.html", students=students)

# -----------------------------
# VIEW ATTENDANCE RECORDS
# -----------------------------
@app.route("/attendance_records")
@admin_login_required
def view_attendance():
    mode = request.args.get("mode", "today").strip().lower()
    selected_date = request.args.get("date", "").strip()

    if selected_date:
        parsed_date = parse_date_filter(selected_date)
        records = get_attendance_records(date_filter=parsed_date)
    elif mode == "all":
        records = get_all_attendance_records()
    else:
        records = [
            [row["id"], row["name"], row["department"], row["date"], row["time"], row["status"]]
            for row in get_today_attendance_data()
        ]

    return render_template("attendance.html", data=records)

# -----------------------------
# STUDENT REPORT
# -----------------------------
@app.route("/student/<student_id>")
@admin_login_required
def student_report(student_id):
    student = get_student_tuple(student_id)

    if not student:
        return "Student not found"

    summary = get_attendance_summary_for_student(student_id)

    records = []
    for date_value in summary["present_dates"]:
        records.append((date_value, "Present"))
    for date_value in summary["absent_dates"]:
        records.append((date_value, "Absent"))

    records.sort(key=lambda x: x[0])

    return render_template(
        "student_report.html",
        student=student,
        parent=(None, None, student[4], student[5], student[6]),
        records=records,
        present=summary["present_days"],
        absent=summary["absent_days"],
        percentage=summary["percentage"]
    )

# -----------------------------
# REGISTER STUDENT
# -----------------------------
@app.route("/register", methods=["GET", "POST"])
@admin_login_required
def register():
    if request.method == "POST":
        name = request.form["name"].strip()
        student_id = request.form["student_id"].strip()
        department = request.form["department"].strip()
        password = request.form["password"].strip()
        parent_name = request.form.get("parent_name", "").strip()
        parent_phone = request.form.get("parent_phone", "").strip()
        parent_email = request.form.get("parent_email", "").strip()

        if not name or not student_id or not department or not password:
            return "All required fields must be filled"

        if student_id_exists(student_id):
            return "Student ID already exists"

        cap = open_camera(0)
        if cap is None:
            return "Unable to open camera"

        file_path = os.path.join(FACES_DIR, f"{student_id}.jpg")
        captured = False

        while True:
            ret, frame = cap.read()

            if not ret or frame is None:
                cap.release()
                cv2.destroyAllWindows()
                return "Camera not working properly"

            frame = cv2.flip(frame, 1)

            cv2.putText(
                frame,
                "Press SPACE to capture | ESC to cancel",
                (20, 30),
                cv2.FONT_HERSHEY_SIMPLEX,
                0.7,
                (0, 255, 255),
                2
            )

            cv2.imshow("Register Face", frame)
            key = cv2.waitKey(1) & 0xFF

            if key == 32:
                cv2.imwrite(file_path, frame)
                captured = True
                break
            elif key == 27:
                break

        cap.release()
        cv2.destroyAllWindows()

        if not captured:
            return "Face registration cancelled"

        try:
            image = face_recognition.load_image_file(file_path)
            face_locations = face_recognition.face_locations(image)

            if len(face_locations) == 0:
                if os.path.exists(file_path):
                    os.remove(file_path)
                return "No face detected. Please try again."

            if len(face_locations) > 1:
                if os.path.exists(file_path):
                    os.remove(file_path)
                return "Multiple faces detected. Please capture only one face."

            if is_duplicate_face(file_path):
                if os.path.exists(file_path):
                    os.remove(file_path)
                return "Duplicate face detected. Student already exists."

            conn = get_db_connection()
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO students (
                    student_id, name, department, password,
                    parent_name, parent_phone, parent_email, photo
                )
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
            """, (
                student_id,
                name,
                department,
                password,
                parent_name,
                parent_phone,
                parent_email,
                f"{student_id}.jpg"
            ))
            conn.commit()
            cursor.close()
            conn.close()

            generate_encodings()
            return redirect(url_for("dashboard"))

        except IntegrityError:
            if os.path.exists(file_path):
                os.remove(file_path)
            return "Student ID already exists"

        except Exception as e:
            if os.path.exists(file_path):
                os.remove(file_path)
            return f"Error while saving student: {str(e)}"

    return render_template("register.html")

# -----------------------------
# ATTENDANCE SYSTEM
# -----------------------------
@app.route("/attendance")
@admin_login_required
def attendance():
    known_encodings, known_ids = read_cached_encodings()

    if len(known_encodings) == 0:
        return "No registered faces found"

    cap = open_camera(0)
    if cap is None:
        return "Unable to open laptop camera"

    student_map = get_all_student_map()

    conn = get_db_connection()
    cursor = conn.cursor()

    today = get_today_date()
    marked_ids = set()

    previous_detections = []
    frame_counter = 0

    try:
        while True:
            ret, frame = cap.read()
            if not ret or frame is None:
                break

            frame = cv2.flip(frame, 1)
            frame_counter += 1

            if frame_counter % ATTENDANCE_PROCESS_EVERY_N_FRAMES == 0:
                new_detections = build_attendance_detections(
                    frame=frame,
                    known_encodings=known_encodings,
                    known_ids=known_ids,
                    student_map=student_map,
                    marked_ids=marked_ids,
                    db_cursor=cursor,
                    today=today
                )
                conn.commit()
                previous_detections = stabilize_detections(new_detections, previous_detections)
            else:
                previous_detections = decay_detections(previous_detections)

            draw_detection_boxes(frame, previous_detections)
            draw_attendance_header(frame, marked_ids, previous_detections)

            cv2.imshow("Laptop Camera Attendance", frame)

            key = cv2.waitKey(1) & 0xFF
            if key == ord("q"):
                break

    finally:
        cursor.close()
        conn.close()
        cap.release()
        cv2.destroyAllWindows()

    mark_absent_students_for_today()
    check_absent_students()
    return redirect(url_for("dashboard"))

# -----------------------------
# CAMERA PAGE
# -----------------------------
@app.route("/camera")
@admin_login_required
def camera_page():
    return render_template("camera.html")

# -----------------------------
# PROCESS ATTENDANCE
# -----------------------------
@app.route("/process_attendance", methods=["POST"])
@admin_login_required
def process_attendance():
    data = request.get_json()

    if not data or "image" not in data:
        return jsonify({"message": "No image received"})

    try:
        image_data = data["image"]
        header, encoded = image_data.split(",", 1)
        image_bytes = base64.b64decode(encoded)
        np_arr = np.frombuffer(image_bytes, np.uint8)
        frame = cv2.imdecode(np_arr, cv2.IMREAD_COLOR)

        if frame is None:
            return jsonify({"message": "Image decode failed"})

        small_frame = cv2.resize(frame, (0, 0), fx=ATTENDANCE_FRAME_SCALE, fy=ATTENDANCE_FRAME_SCALE)
        rgb_small_frame = cv2.cvtColor(small_frame, cv2.COLOR_BGR2RGB)

        known_encodings, known_ids = read_cached_encodings()

        if len(known_encodings) == 0:
            return jsonify({"message": "No registered faces found"})

        face_locations = face_recognition.face_locations(rgb_small_frame)
        face_encodings = face_recognition.face_encodings(rgb_small_frame, face_locations)

        if not face_encodings:
            return jsonify({"message": "No face detected"})

        student_map = get_all_student_map()

        conn = get_db_connection()
        cursor = conn.cursor()
        messages = []

        today = get_today_date()
        current_time = datetime.now().strftime("%H:%M:%S")

        for (top, right, bottom, left), encoding in zip(face_locations, face_encodings):
            top = int(top / ATTENDANCE_FRAME_SCALE)
            right = int(right / ATTENDANCE_FRAME_SCALE)
            bottom = int(bottom / ATTENDANCE_FRAME_SCALE)
            left = int(left / ATTENDANCE_FRAME_SCALE)

            matched_student_id = None

            face_distances = face_recognition.face_distance(known_encodings, encoding)
            if len(face_distances) > 0:
                best_match_index = int(np.argmin(face_distances))
                best_distance = float(face_distances[best_match_index])

                if best_distance < KNOWN_FACE_TOLERANCE:
                    matched_student_id = known_ids[best_match_index]

            if matched_student_id:
                student_info = student_map.get(matched_student_id, {})
                student_name = student_info.get("name", matched_student_id)
                department = student_info.get("department", "Unknown")

                cursor.execute(
                    "SELECT 1 FROM attendance WHERE student_id = %s AND attendance_date = %s",
                    (matched_student_id, today)
                )
                already_marked = cursor.fetchone()

                if not already_marked:
                    cursor.execute("""
                        INSERT INTO attendance (student_id, name, department, attendance_date, attendance_time, status)
                        VALUES (%s, %s, %s, %s, %s, %s)
                    """, (
                        matched_student_id,
                        student_name,
                        department,
                        today,
                        current_time,
                        "Present"
                    ))
                    conn.commit()
                    messages.append(f"Attendance marked for {student_name} ({matched_student_id})")
                else:
                    messages.append(f"Attendance already marked for {student_name} ({matched_student_id})")
            else:
                if should_save_unknown_face(
                    encoding,
                    cooldown=UNKNOWN_COOLDOWN,
                    tolerance=UNKNOWN_FACE_TOLERANCE
                ):
                    save_unknown_face(frame, top, right, bottom, left)
                    messages.append("Unknown face detected and saved")
                else:
                    messages.append("Unknown face detected")

        cursor.close()
        conn.close()

        return jsonify({"message": " | ".join(messages)})

    except Exception as e:
        return jsonify({"message": f"Error: {str(e)}"})

# -----------------------------
# REPORT PAGE
# -----------------------------
@app.route("/report", methods=["GET", "POST"])
@admin_login_required
def report():
    search_id = ""

    if request.method == "POST":
        search_id = request.form.get("student_id", "").strip()
    else:
        search_id = request.args.get("student_id", "").strip()

    data = get_today_attendance_data(search_id)
    summary = None

    if search_id and student_id_exists(search_id):
        summary = get_attendance_summary_for_student(search_id)

    return render_template(
        "report.html",
        records=data,
        summary=summary,
        search_id=search_id
    )

# -----------------------------
# EXPORT TO EXCEL
# -----------------------------
@app.route("/export_excel")
@admin_login_required
def export_excel():
    search_id = request.args.get("student_id", "").strip()
    data = get_today_attendance_data(search_id)

    wb = Workbook()
    ws = wb.active
    ws.title = "Attendance Report"

    ws.append(["Student ID", "Name", "Department", "Date", "Time", "Status"])

    for row in data:
        ws.append([
            row["id"],
            row["name"],
            row["department"],
            row["date"],
            row["time"],
            row["status"]
        ])

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="attendance_report.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# -----------------------------
# EXPORT TO PDF
# -----------------------------
@app.route("/export_pdf")
@admin_login_required
def export_pdf():
    search_id = request.args.get("student_id", "").strip()
    data = get_today_attendance_data(search_id)

    output = BytesIO()
    pdf = pdf_canvas.Canvas(output, pagesize=A4)
    width, height = A4

    y = height - 50

    pdf.setFont("Helvetica-Bold", 16)
    pdf.drawString(200, y, "Attendance Report")
    y -= 30

    pdf.setFont("Helvetica-Bold", 10)
    pdf.drawString(40, y, "Student ID")
    pdf.drawString(120, y, "Name")
    pdf.drawString(230, y, "Department")
    pdf.drawString(350, y, "Date")
    pdf.drawString(440, y, "Time")
    pdf.drawString(500, y, "Status")
    y -= 20

    pdf.setFont("Helvetica", 10)

    if not data:
        pdf.drawString(40, y, "No attendance records found.")
    else:
        for row in data:
            if y < 50:
                pdf.showPage()
                y = height - 50

                pdf.setFont("Helvetica-Bold", 10)
                pdf.drawString(40, y, "Student ID")
                pdf.drawString(120, y, "Name")
                pdf.drawString(230, y, "Department")
                pdf.drawString(350, y, "Date")
                pdf.drawString(440, y, "Time")
                pdf.drawString(500, y, "Status")
                y -= 20
                pdf.setFont("Helvetica", 10)

            pdf.drawString(40, y, str(row["id"])[:12])
            pdf.drawString(120, y, str(row["name"])[:18])
            pdf.drawString(230, y, str(row["department"])[:18])
            pdf.drawString(350, y, str(row["date"])[:15])
            pdf.drawString(440, y, str(row["time"])[:10])
            pdf.drawString(500, y, str(row["status"])[:10])
            y -= 20

    pdf.save()
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="attendance_report.pdf",
        mimetype="application/pdf"
    )

# -----------------------------
# DOWNLOAD TODAY ATTENDANCE CSV
# -----------------------------
@app.route("/download")
@admin_login_required
def download():
    data = get_today_attendance_data()
    if not data:
        return "No attendance records found for today"

    output = build_csv_output(data)

    return send_file(
        output,
        as_attachment=True,
        download_name=f"attendance_{get_today_date_str()}.csv",
        mimetype="text/csv"
    )

# -----------------------------
# ATTENDANCE PERCENTAGE
# -----------------------------
@app.route("/percentage")
@admin_login_required
def percentage():
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT student_id, name FROM students ORDER BY student_id ASC")
    students = cursor.fetchall()

    cursor.execute("""
        SELECT DISTINCT attendance_date
        FROM attendance
        ORDER BY attendance_date
    """)
    all_attendance_dates = cursor.fetchall()
    total_working_days = len(all_attendance_dates)

    data = []

    for student in students:
        student_id = student["student_id"]
        name = student["name"]

        cursor.execute("""
            SELECT COUNT(*) AS present_days
            FROM attendance
            WHERE student_id = %s AND status = 'Present'
        """, (student_id,))
        result = cursor.fetchone()

        present_days = result["present_days"] or 0
        total_days = total_working_days
        percentage_value = round((present_days / total_days) * 100, 2) if total_days > 0 else 0

        data.append({
            "id": student_id,
            "name": name,
            "present_days": present_days,
            "total_days": total_days,
            "percentage": percentage_value
        })

    cursor.close()
    conn.close()
    return render_template("percentage.html", records=data)

# -----------------------------
# CHANGE ADMIN PASSWORD
# -----------------------------
@app.route("/change_admin_password", methods=["GET", "POST"])
@admin_login_required
def change_admin_password():
    if request.method == "POST":
        username = request.form["username"].strip()
        old_password = request.form["old_password"].strip()
        new_password = request.form["new_password"].strip()

        if not new_password:
            return "New password cannot be empty"

        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute(
            "SELECT * FROM admins WHERE username = %s AND password = %s",
            (username, old_password)
        )
        result = cursor.fetchone()

        if result:
            cursor.execute(
                "UPDATE admins SET password = %s WHERE username = %s",
                (new_password, username)
            )
            conn.commit()
            cursor.close()
            conn.close()
            return "Admin Password Changed Successfully"

        cursor.close()
        conn.close()
        return "Old Password Incorrect"

    return render_template("change_admin_password.html")

# -----------------------------
# CHANGE STUDENT PASSWORD
# -----------------------------
@app.route("/change_student_password", methods=["GET", "POST"])
def change_student_password():
    if request.method == "POST":
        student_id = request.form["student_id"].strip()
        old_password = request.form.get("old_password", "").strip()
        new_password = request.form["new_password"].strip()

        if not new_password:
            return "New password cannot be empty"

        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute(
            "SELECT * FROM students WHERE student_id = %s AND password = %s",
            (student_id, old_password)
        )
        student = cursor.fetchone()

        if student:
            cursor.execute(
                "UPDATE students SET password = %s WHERE student_id = %s",
                (new_password, student_id)
            )
            conn.commit()
            cursor.close()
            conn.close()
            return "Student password changed successfully"

        cursor.close()
        conn.close()
        return "Student ID or old password is incorrect"

    return render_template("change_student_password.html")

# -----------------------------
# EDIT STUDENT
# -----------------------------
@app.route("/edit_student/<student_id>", methods=["GET", "POST"])
@admin_login_required
def edit_student(student_id):
    student = get_student_tuple(student_id)
    if not student:
        return "Student not found"

    if request.method == "POST":
        name = request.form["name"].strip()
        department = request.form["department"].strip()
        password = request.form["password"].strip()
        parent_name = request.form["parent_name"].strip()
        parent_phone = request.form["parent_phone"].strip()
        parent_email = request.form["parent_email"].strip()

        conn = get_db_connection()
        cursor = conn.cursor()

        cursor.execute("""
            UPDATE students
            SET name = %s, department = %s, password = %s, parent_name = %s, parent_phone = %s, parent_email = %s
            WHERE student_id = %s
        """, (
            name, department, password, parent_name, parent_phone, parent_email, student_id
        ))

        cursor.execute("""
            UPDATE attendance
            SET name = %s, department = %s
            WHERE student_id = %s
        """, (
            name, department, student_id
        ))

        conn.commit()
        cursor.close()
        conn.close()

        return redirect(url_for("view_students"))

    return render_template("edit_student.html", student=student)

# -----------------------------
# DELETE STUDENT
# -----------------------------
@app.route("/delete_student/<student_id>")
@admin_login_required
def delete_student(student_id):
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("DELETE FROM attendance WHERE student_id = %s", (student_id,))
    cursor.execute("DELETE FROM absence_alerts WHERE student_id = %s", (student_id,))
    cursor.execute("DELETE FROM students WHERE student_id = %s", (student_id,))

    conn.commit()
    cursor.close()
    conn.close()

    photo_path = os.path.join(FACES_DIR, f"{student_id}.jpg")
    if os.path.exists(photo_path):
        os.remove(photo_path)

    generate_encodings()
    return redirect(url_for("view_students"))

# -----------------------------
# AUTHOR PAGE
# -----------------------------
@app.route("/author")
@admin_login_required
def author():
    author_data = {
        "full_name": "Your Name",
        "role": "Project Developer",
        "college_name": "Your College Name",
        "department": "Your Department",
        "email": "your_email@gmail.com",
        "phone": "Your Phone Number",
        "project_title": "VisionDB - Smart Attendance System using Face Recognition",
        "guide_name": "Your Guide Name"
    }
    return render_template("author.html", author=author_data)

# -----------------------------
# RUN SERVER
# -----------------------------
if __name__ == "__main__":
    app.run(debug=True)