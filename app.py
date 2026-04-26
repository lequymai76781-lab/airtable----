from pathlib import Path
from datetime import datetime
from functools import wraps
from io import BytesIO
import os
import sqlite3

from flask import (
    Flask,
    request,
    jsonify,
    render_template,
    send_file,
    send_from_directory,
    session
)
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from werkzeug.security import generate_password_hash, check_password_hash


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DB_PATH = DATA_DIR / "choices.db"

DEFAULT_ADMIN_USERNAME = "admin"
DEFAULT_ADMIN_PASSWORD = "123456"

TUTOR_NAMES = [
    "陈积银", "窦光华", "付晓静", "付志华", "高立", "韩瑞珍", "侯小琴", "胡家浩",
    "黄海", "姜小凌", "李爱群", "李保存", "李菁", "刘娟", "刘晓丽", "吕永峰",
    "彭肇一", "邵娟", "滕姗姗", "万晓红", "汪蓓", "王创业", "王润斌", "王雷",
    "王相飞", "王真真", "王雪莲", "吴丹", "肖宁", "姚洪磊", "游迎亚", "张德胜",
    "张钢花", "周榕", "邹瑶"
]

app = Flask(
    __name__,
    template_folder=str(BASE_DIR)
)

app.secret_key = os.environ.get(
    "SECRET_KEY",
    "wti-graduate-tutor-choice-system-secret-key"
)


def now_text():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def get_conn():
    DATA_DIR.mkdir(exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    DATA_DIR.mkdir(exist_ok=True)

    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS admins (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL,
            created_at TEXT NOT NULL
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS submissions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_name TEXT NOT NULL,
            major TEXT NOT NULL,
            contact TEXT NOT NULL UNIQUE,
            first_tutor TEXT NOT NULL,
            second_tutor TEXT NOT NULL,
            third_tutor TEXT NOT NULL,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS submission_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_name TEXT NOT NULL,
            major TEXT NOT NULL,
            contact TEXT NOT NULL,
            first_tutor TEXT NOT NULL,
            second_tutor TEXT NOT NULL,
            third_tutor TEXT NOT NULL,
            action TEXT NOT NULL,
            created_at TEXT NOT NULL
        )
    """)

    admin = cur.execute(
        "SELECT id FROM admins WHERE username = ?",
        (DEFAULT_ADMIN_USERNAME,)
    ).fetchone()

    if not admin:
        cur.execute("""
            INSERT INTO admins (
                username,
                password_hash,
                created_at
            )
            VALUES (?, ?, ?)
        """, (
            DEFAULT_ADMIN_USERNAME,
            generate_password_hash(DEFAULT_ADMIN_PASSWORD),
            now_text()
        ))

    conn.commit()
    conn.close()


def admin_required(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        if session.get("admin_logged_in") is not True:
            return jsonify({
                "code": 401,
                "message": "请先登录管理员账号。"
            }), 401

        return func(*args, **kwargs)

    return wrapper


def validate_submission(data):
    student_name = str(data.get("studentName", "")).strip()
    major = str(data.get("major", "")).strip()
    contact = str(data.get("contact", "")).strip()
    first_choice = str(data.get("firstChoice", "")).strip()
    second_choice = str(data.get("secondChoice", "")).strip()
    third_choice = str(data.get("thirdChoice", "")).strip()

    if not student_name:
        return None, "请填写姓名。"

    if not major:
        return None, "请填写专业。"

    if not contact:
        return None, "请填写联系方式。"

    if not first_choice or not second_choice or not third_choice:
        return None, "请完整选择第一、第二、第三志愿导师。"

    if first_choice not in TUTOR_NAMES:
        return None, "第一志愿导师不在导师名单中。"

    if second_choice not in TUTOR_NAMES:
        return None, "第二志愿导师不在导师名单中。"

    if third_choice not in TUTOR_NAMES:
        return None, "第三志愿导师不在导师名单中。"

    if len({first_choice, second_choice, third_choice}) != 3:
        return None, "第一志愿、第二志愿、第三志愿不能选择同一位导师。"

    return {
        "student_name": student_name,
        "major": major,
        "contact": contact,
        "first_tutor": first_choice,
        "second_tutor": second_choice,
        "third_tutor": third_choice,
    }, None


def get_all_submissions():
    conn = get_conn()
    rows = conn.execute("""
        SELECT
            id,
            student_name,
            major,
            contact,
            first_tutor,
            second_tutor,
            third_tutor,
            created_at,
            updated_at
        FROM submissions
        ORDER BY updated_at DESC, id DESC
    """).fetchall()
    conn.close()
    return rows


def compute_tutor_stats():
    rows = get_all_submissions()

    stats_map = {
        tutor: {
            "tutor": tutor,
            "first": 0,
            "second": 0,
            "third": 0,
            "total": 0
        }
        for tutor in TUTOR_NAMES
    }

    for row in rows:
        first_tutor = row["first_tutor"]
        second_tutor = row["second_tutor"]
        third_tutor = row["third_tutor"]

        if first_tutor in stats_map:
            stats_map[first_tutor]["first"] += 1

        if second_tutor in stats_map:
            stats_map[second_tutor]["second"] += 1

        if third_tutor in stats_map:
            stats_map[third_tutor]["third"] += 1

    stats = []

    for item in stats_map.values():
        item["total"] = item["first"] + item["second"] + item["third"]
        stats.append(item)

    stats.sort(
        key=lambda item: (
            -item["first"],
            -item["total"],
            item["tutor"]
        )
    )

    return stats


@app.route("/")
def index():
    return render_template("app.html")


@app.route("/admin")
def admin_page():
    return render_template("admin.html")


@app.route("/武体校徽.jpeg")
def logo():
    return send_from_directory(BASE_DIR, "武体校徽.jpeg")


@app.route("/api/tutors", methods=["GET"])
def api_tutors():
    return jsonify({
        "code": 200,
        "data": TUTOR_NAMES
    })


@app.route("/api/public/tutor-stats", methods=["GET"])
def api_public_tutor_stats():
    return jsonify({
        "code": 200,
        "data": compute_tutor_stats()
    })


@app.route("/api/submit-choice", methods=["POST"])
def api_submit_choice():
    data = request.get_json(silent=True) or {}
    valid_data, error = validate_submission(data)

    if error:
        return jsonify({
            "code": 400,
            "message": error
        }), 400

    conn = get_conn()
    cur = conn.cursor()

    current_time = now_text()

    existing = cur.execute(
        "SELECT id FROM submissions WHERE contact = ?",
        (valid_data["contact"],)
    ).fetchone()

    if existing:
        cur.execute("""
            UPDATE submissions
            SET
                student_name = ?,
                major = ?,
                first_tutor = ?,
                second_tutor = ?,
                third_tutor = ?,
                updated_at = ?
            WHERE contact = ?
        """, (
            valid_data["student_name"],
            valid_data["major"],
            valid_data["first_tutor"],
            valid_data["second_tutor"],
            valid_data["third_tutor"],
            current_time,
            valid_data["contact"]
        ))

        action = "update"
        message = "志愿已更新。"
    else:
        cur.execute("""
            INSERT INTO submissions (
                student_name,
                major,
                contact,
                first_tutor,
                second_tutor,
                third_tutor,
                created_at,
                updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            valid_data["student_name"],
            valid_data["major"],
            valid_data["contact"],
            valid_data["first_tutor"],
            valid_data["second_tutor"],
            valid_data["third_tutor"],
            current_time,
            current_time
        ))

        action = "create"
        message = "志愿提交成功。"

    cur.execute("""
        INSERT INTO submission_logs (
            student_name,
            major,
            contact,
            first_tutor,
            second_tutor,
            third_tutor,
            action,
            created_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        valid_data["student_name"],
        valid_data["major"],
        valid_data["contact"],
        valid_data["first_tutor"],
        valid_data["second_tutor"],
        valid_data["third_tutor"],
        action,
        current_time
    ))

    conn.commit()
    conn.close()

    return jsonify({
        "code": 200,
        "message": message,
        "data": {
            "action": action
        }
    })


@app.route("/api/admin/login", methods=["POST"])
def api_admin_login():
    data = request.get_json(silent=True) or {}

    username = str(data.get("username", "")).strip()
    password = str(data.get("password", "")).strip()

    if not username or not password:
        return jsonify({
            "code": 400,
            "message": "请输入管理员账号和密码。"
        }), 400

    conn = get_conn()
    admin = conn.execute(
        "SELECT id, username, password_hash FROM admins WHERE username = ?",
        (username,)
    ).fetchone()
    conn.close()

    if not admin or not check_password_hash(admin["password_hash"], password):
        return jsonify({
            "code": 401,
            "message": "管理员账号或密码错误。"
        }), 401

    session["admin_logged_in"] = True
    session["admin_id"] = admin["id"]
    session["admin_username"] = admin["username"]

    return jsonify({
        "code": 200,
        "message": "管理员登录成功。",
        "data": {
            "username": admin["username"]
        }
    })


@app.route("/api/admin/logout", methods=["POST"])
def api_admin_logout():
    session.clear()

    return jsonify({
        "code": 200,
        "message": "已退出管理员登录。"
    })


@app.route("/api/admin/status", methods=["GET"])
def api_admin_status():
    return jsonify({
        "code": 200,
        "data": {
            "loggedIn": session.get("admin_logged_in") is True,
            "username": session.get("admin_username")
        }
    })


@app.route("/api/admin/submissions", methods=["GET"])
@admin_required
def api_admin_submissions():
    rows = get_all_submissions()

    data = []

    for row in rows:
        data.append({
            "id": row["id"],
            "studentName": row["student_name"],
            "major": row["major"],
            "contact": row["contact"],
            "firstChoice": row["first_tutor"],
            "secondChoice": row["second_tutor"],
            "thirdChoice": row["third_tutor"],
            "createdAt": row["created_at"],
            "updatedAt": row["updated_at"],
        })

    return jsonify({
        "code": 200,
        "data": data
    })


@app.route("/api/admin/tutor-stats", methods=["GET"])
@admin_required
def api_admin_tutor_stats():
    return jsonify({
        "code": 200,
        "data": compute_tutor_stats()
    })


@app.route("/api/admin/logs", methods=["GET"])
@admin_required
def api_admin_logs():
    conn = get_conn()
    rows = conn.execute("""
        SELECT
            id,
            student_name,
            major,
            contact,
            first_tutor,
            second_tutor,
            third_tutor,
            action,
            created_at
        FROM submission_logs
        ORDER BY created_at DESC, id DESC
        LIMIT 300
    """).fetchall()
    conn.close()

    data = []

    for row in rows:
        data.append({
            "id": row["id"],
            "studentName": row["student_name"],
            "major": row["major"],
            "contact": row["contact"],
            "firstChoice": row["first_tutor"],
            "secondChoice": row["second_tutor"],
            "thirdChoice": row["third_tutor"],
            "action": row["action"],
            "createdAt": row["created_at"],
        })

    return jsonify({
        "code": 200,
        "data": data
    })


@app.route("/api/admin/export", methods=["GET"])
@admin_required
def api_admin_export():
    stats = compute_tutor_stats()
    submissions = get_all_submissions()

    wb = Workbook()

    ws1 = wb.active
    ws1.title = "导师志愿统计"

    ws1.append([
        "序号",
        "导师姓名",
        "第一志愿人数",
        "第二志愿人数",
        "第三志愿人数",
        "总选择人数"
    ])

    for index, item in enumerate(stats, start=1):
        ws1.append([
            index,
            item["tutor"],
            item["first"],
            item["second"],
            item["third"],
            item["total"]
        ])

    ws2 = wb.create_sheet("学生提交结果")

    ws2.append([
        "序号",
        "姓名",
        "专业",
        "联系方式",
        "第一志愿",
        "第二志愿",
        "第三志愿",
        "创建时间",
        "更新时间"
    ])

    for index, row in enumerate(submissions, start=1):
        ws2.append([
            index,
            row["student_name"],
            row["major"],
            row["contact"],
            row["first_tutor"],
            row["second_tutor"],
            row["third_tutor"],
            row["created_at"],
            row["updated_at"]
        ])

    style_excel_sheet(ws1)
    style_excel_sheet(ws2)

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"武汉体育学院研究生导师选择结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


def style_excel_sheet(ws):
    header_fill = PatternFill("solid", fgColor="1B61C9")
    header_font = Font(bold=True, color="FFFFFF")
    center_alignment = Alignment(horizontal="center", vertical="center")

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_alignment

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = center_alignment

    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter

        for cell in column:
            value = str(cell.value) if cell.value is not None else ""
            max_length = max(max_length, len(value))

        ws.column_dimensions[column_letter].width = max_length + 4


if __name__ == "__main__":
    init_db()

    port = int(os.environ.get("PORT", 5000))

    app.run(
        host="0.0.0.0",
        port=port,
        debug=True
    )