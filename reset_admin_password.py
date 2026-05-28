from datetime import datetime
from pathlib import Path
import os
import sqlite3

from werkzeug.security import generate_password_hash


BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "data" / "choices.db"
ADMIN_USERNAME = "admin"
DEFAULT_PASSWORD = "123456"


def now_text():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def ensure_admin_table(cur):
    cur.execute("""
        CREATE TABLE IF NOT EXISTS admins (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL,
            created_at TEXT NOT NULL
        )
    """)


def reset_admin_password():
    if not DB_PATH.exists():
        raise FileNotFoundError(f"数据库文件不存在：{DB_PATH}")

    password = os.environ.get("ADMIN_PASSWORD") or DEFAULT_PASSWORD
    password_hash = generate_password_hash(password)

    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    ensure_admin_table(cur)

    admin = cur.execute(
        "SELECT id FROM admins WHERE username = ?",
        (ADMIN_USERNAME,),
    ).fetchone()

    if admin:
        cur.execute(
            "UPDATE admins SET password_hash = ? WHERE username = ?",
            (password_hash, ADMIN_USERNAME),
        )
    else:
        cur.execute("""
            INSERT INTO admins (
                username,
                password_hash,
                created_at
            )
            VALUES (?, ?, ?)
        """, (
            ADMIN_USERNAME,
            password_hash,
            now_text(),
        ))

    conn.commit()
    conn.close()

    print("管理员账号：admin")
    print("管理员密码已重置")
    print("数据库路径：data/choices.db")


if __name__ == "__main__":
    reset_admin_password()
