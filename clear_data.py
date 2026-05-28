from pathlib import Path
import argparse
import os
import secrets
import shutil
import sqlite3

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_DATA_DIR = BASE_DIR / "data"
DEFAULT_DB_PATH = DEFAULT_DATA_DIR / "choices.db"
FALLBACK_DB_PATH = BASE_DIR / "choices.db"


def can_write_directory(path):
    try:
        path.mkdir(exist_ok=True)
        probe = path / f".write_probe_{secrets.token_hex(4)}"
        probe.write_text("", encoding="utf-8")
        probe.unlink(missing_ok=True)
        return True
    except OSError:
        return False


def resolve_db_path():
    configured_path = os.environ.get("DATABASE_PATH")
    if configured_path:
        db_path = Path(configured_path)
        if not db_path.is_absolute():
            db_path = BASE_DIR / db_path
        return db_path.resolve()

    if DEFAULT_DB_PATH.exists() and not can_write_directory(DEFAULT_DATA_DIR):
        if not FALLBACK_DB_PATH.exists():
            try:
                shutil.copy2(DEFAULT_DB_PATH, FALLBACK_DB_PATH)
            except OSError:
                pass
        return FALLBACK_DB_PATH.resolve()

    return DEFAULT_DB_PATH.resolve()


DB_PATH = resolve_db_path()
CONFIRM_TEXT = "确认清空"

parser = argparse.ArgumentParser(
    description="清空学生提交数据和提交日志，保留导师名单和管理员账号。"
)
parser.add_argument(
    "--confirm",
    help=f"确认口令，必须填写：{CONFIRM_TEXT}",
)
args = parser.parse_args()

if args.confirm != CONFIRM_TEXT:
    raise SystemExit(
        f"为避免误删，请使用：python clear_data.py --confirm {CONFIRM_TEXT}"
    )

conn = sqlite3.connect(DB_PATH)
cur = conn.cursor()

cur.execute("DELETE FROM submissions")
cur.execute("DELETE FROM submission_logs")

conn.commit()
conn.close()

print("学生提交数据和提交日志已清空，管理员账号保留。")
