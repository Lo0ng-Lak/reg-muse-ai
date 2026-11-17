"""Kho tài khoản duy nhất: SQLite thay cho accounts.xlsx.

Một tài khoản là một dòng trong bảng accounts, email là khoá
duy nhất (không phân biệt hoa thường) nên không bao giờ có hai
dòng cùng một mail, không bao giờ lẫn cookie giữa các mail.
Đọc/ghi một dòng không động đến các dòng khác — với 2k tài
khoản vẫn tức thì, khác hẳn Excel (ghi một dòng là ghi lại cả
file, mở file trong Excel là khoá cứng).

accounts.xlsx không bị xoá: nó là ĐƯỜNG VÀO (nhập một lần từ
file cũ) và ĐƯỜNG RA (xuất ra Excel khi cần xem bằng Excel).
Kho sống là accounts.db trong cùng thư mục dữ liệu chạy.

Chỉ dùng thư viện chuẩn (sqlite3) + openpyxl (đã có) cho
nhập/xuất Excel. Mỗi thao tác mở một kết nối ngắn, WAL cho
phép web + tool đọc/ghi song song, khoá ghi để hai luồng không
giẫm lên nhau.
"""

import io
import re
import sqlite3
import threading
from datetime import datetime
from pathlib import Path

import paths

DB_PATH = paths.APP_DATA / "accounts.db"

EXCEL_HEADERS = [
    "Thời gian tạo",
    "Email Muse.ai",
    "Mật khẩu hộp thư",
    "Cookie phiên đăng nhập",
    "Mã mời",
    "Link lấy OTP",
    "Thời gian lấy cookie",
]

# Mã mời hợp lệ: 6 ký tự chữ/số (quy tắc cũ trong storage.py).
_INVITE_RE = re.compile(r"[A-Z0-9]{6}")

_WRITE_LOCK = threading.RLock()


def _connect(db_path=None):
    """Kết nối ngắn tới DB (tạo bảng nếu chưa có)."""
    path = Path(db_path) if db_path is not None else Path(DB_PATH)
    path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(path), timeout=30.0)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA busy_timeout=5000")
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS accounts (
            email TEXT PRIMARY KEY COLLATE NOCASE,
            mail_password TEXT NOT NULL DEFAULT '',
            cookie TEXT NOT NULL DEFAULT '',
            invite TEXT NOT NULL DEFAULT '',
            otp_link TEXT NOT NULL DEFAULT '',
            created_at TEXT NOT NULL DEFAULT '',
            cookie_taken_at TEXT NOT NULL DEFAULT ''
        )
        """
    )
    conn.commit()
    return conn


def _row_to_dict(rowid, row):
    """Một hàng DB -> dict cùng hình dạng với dòng Excel cũ."""
    return {
        "row": rowid,
        "email": row[0],
        "mail_password": row[1],
        "cookie": row[2],
        "invite": row[3],
        "otp_link": row[4],
        "created_at": row[5],
        "cookie_taken_at": row[6],
    }


def _merge(old, new):
    """Gộp dòng mới vào dòng cũ theo từng trường.

    Trường mới có giá trị thì thắng, trống thì giữ cũ — để một
    dòng import thiếu cookie không xoá mất cookie đang có. Ngoại
    lệ: cookie mới đi kèm mốc lấy mới của nó; "Thời gian tạo"
    thì giữ mốc đầu tiên (lúc tạo tài khoản).
    """
    merged = dict(old)
    for field in ("mail_password", "invite", "otp_link"):
        if (new.get(field) or "").strip():
            merged[field] = new[field]
    if (new.get("cookie") or "").strip():
        merged["cookie"] = new["cookie"]
        merged["cookie_taken_at"] = new.get("cookie_taken_at") or ""
    if not (merged.get("created_at") or "").strip():
        merged["created_at"] = new.get("created_at") or ""
    return merged


def upsert_account(email, mail_password="", cookie="", invite="",
                   otp_link="", created_at="", cookie_taken_at="",
                   db_path=None):
    """Thêm tài khoản mới, hoặc cập nhật đúng dòng của email đó.

    Email là khoá duy nhất nên gọi hai lần với cùng một mail
    không bao giờ tạo hai dòng (khác Excel cũ: append là thêm
    dòng mới). Trả về True.
    """
    email = (email or "").strip()
    if not email:
        raise ValueError("thiếu email")
    new = {
        "email": email,
        "mail_password": mail_password or "",
        "cookie": cookie or "",
        "invite": invite or "",
        "otp_link": otp_link or "",
        "created_at": created_at or "",
        "cookie_taken_at": cookie_taken_at or "",
    }
    with _WRITE_LOCK:
        conn = _connect(db_path)
        try:
            cur = conn.execute(
                "SELECT rowid, email, mail_password, cookie, invite,"
                " otp_link, created_at, cookie_taken_at"
                " FROM accounts WHERE email = ?",
                (email,),
            )
            found = cur.fetchone()
            if found is None:
                conn.execute(
                    "INSERT INTO accounts (email, mail_password,"
                    " cookie, invite, otp_link, created_at,"
                    " cookie_taken_at) VALUES (?, ?, ?, ?, ?, ?, ?)",
                    (
                        new["email"],
                        new["mail_password"],
                        new["cookie"],
                        new["invite"],
                        new["otp_link"],
                        new["created_at"],
                        new["cookie_taken_at"],
                    ),
                )
            else:
                old = _row_to_dict(found[0], found[1:])
                merged = _merge(old, new)
                conn.execute(
                    "UPDATE accounts SET mail_password=?, cookie=?,"
                    " invite=?, otp_link=?, created_at=?,"
                    " cookie_taken_at=? WHERE email=?",
                    (
                        merged["mail_password"],
                        merged["cookie"],
                        merged["invite"],
                        merged["otp_link"],
                        merged["created_at"],
                        merged["cookie_taken_at"],
                        email,
                    ),
                )
            conn.commit()
        finally:
            conn.close()
    return True


def list_accounts(db_path=None):
    """Mọi tài khoản, theo thứ tự thêm vào (cũ trước, mới sau)."""
    conn = _connect(db_path)
    try:
        rows = conn.execute(
            "SELECT rowid, email, mail_password, cookie, invite,"
            " otp_link, created_at, cookie_taken_at"
            " FROM accounts ORDER BY rowid"
        ).fetchall()
        return [_row_to_dict(row[0], row[1:]) for row in rows]
    finally:
        conn.close()


def get_account(email, db_path=None):
    """Một tài khoản theo email (None khi không có)."""
    target = (email or "").strip()
    if not target:
        return None
    conn = _connect(db_path)
    try:
        row = conn.execute(
            "SELECT rowid, email, mail_password, cookie, invite,"
            " otp_link, created_at, cookie_taken_at"
            " FROM accounts WHERE email = ?",
            (target,),
        ).fetchone()
        if row is None:
            return None
        return _row_to_dict(row[0], row[1:])
    finally:
        conn.close()


def update_cookie(email, cookie, taken_at, db_path=None):
    """Ghi đè cookie + mốc lấy của đúng dòng email đó.

    Trả về True nếu cập nhật được, False khi không có dòng này
    (người gọi tự thêm dòng mới khi cần).
    """
    target = (email or "").strip()
    if not target:
        return False
    with _WRITE_LOCK:
        conn = _connect(db_path)
        try:
            cur = conn.execute(
                "UPDATE accounts SET cookie=?, cookie_taken_at=?"
                " WHERE email=?",
                (cookie or "", taken_at or "", target),
            )
            conn.commit()
            return cur.rowcount > 0
        finally:
            conn.close()


def clear_cookie(email, db_path=None):
    """Xoá cookie + mốc lấy của một tài khoản (giữ lại dòng)."""
    target = (email or "").strip()
    if not target:
        return False
    with _WRITE_LOCK:
        conn = _connect(db_path)
        try:
            cur = conn.execute(
                "UPDATE accounts SET cookie='', cookie_taken_at=''"
                " WHERE email=?",
                (target,),
            )
            conn.commit()
            return cur.rowcount > 0
        finally:
            conn.close()


def count(db_path=None):
    """Số tài khoản trong kho."""
    conn = _connect(db_path)
    try:
        return conn.execute(
            "SELECT COUNT(*) FROM accounts"
        ).fetchone()[0]
    finally:
        conn.close()


def count_with_cookie(db_path=None):
    """Số tài khoản đang có cookie."""
    conn = _connect(db_path)
    try:
        return conn.execute(
            "SELECT COUNT(*) FROM accounts"
            " WHERE cookie IS NOT NULL AND TRIM(cookie) <> ''"
        ).fetchone()[0]
    finally:
        conn.close()


def last_invite(db_path=None):
    """Mã mời hợp lệ gần nhất (dòng thêm vào sau cùng có mã).

    Dùng làm mã đầu chuỗi khi ô Mã mời đang trống.
    """
    conn = _connect(db_path)
    try:
        rows = conn.execute(
            "SELECT invite FROM accounts ORDER BY rowid DESC"
        ).fetchall()
        for (invite,) in rows:
            candidate = (invite or "").strip().upper()
            if _INVITE_RE.fullmatch(candidate):
                return candidate
        return ""
    finally:
        conn.close()


def relogin_accounts(db_path=None):
    """Các mail cần đăng nhập lại: email + bí mật hộp thư.

    Dùng cho tab Đăng nhập lại và web: OTP lấy từ hộp thư của
    đúng mail đó. Provider suy ra từ domain ở tầng trên
    (storage._provider_for_email).
    """
    conn = _connect(db_path)
    try:
        rows = conn.execute(
            "SELECT email, mail_password FROM accounts"
            " ORDER BY rowid"
        ).fetchall()
        result = []
        for email, secret in rows:
            email = (email or "").strip()
            if "@" not in email:
                continue
            result.append(
                {"email": email, "secret": (secret or "").strip()}
            )
        return result
    finally:
        conn.close()


# --- Nhập / xuất Excel ----------------------------------------------


def _excel_columns(sheet):
    """Vị trí các cột của sheet theo tên tiêu đề."""
    columns = {}
    for cell in sheet[1]:
        if isinstance(cell.value, str):
            columns[cell.value.strip()] = cell.column
    return columns


def import_from_excel(excel_path, db_path=None):
    """Nhập tài khoản từ file Excel vào kho (gộp theo email).

    Đọc theo tên tiêu đề nên sheet 6 cột đời cũ vẫn nhập được
    (thiếu cột "Thời gian lấy cookie" thì mốc lấy để trống).
    File chờ accounts.pending.xlsx nhập SAU file chính để dòng
    mới hơn trong file chờ thắng (nó là bản đầy đủ nhất sau lần
    ghi bị khoá). Trả về {"imported", "updated", "skipped",
    "total"}: imported là mail chưa có trong kho, updated là
    mail đã có và được gộp thêm, skipped là dòng không có email.
    File không có / không đọc được -> toàn 0, không lỗi.
    """
    from openpyxl import load_workbook

    path = Path(excel_path) if excel_path else None
    if path is None or not path.is_file():
        return {"imported": 0, "updated": 0, "skipped": 0,
                "total": count(db_path)}
    try:
        workbook = load_workbook(path, read_only=True, data_only=True)
    except Exception:
        return {"imported": 0, "updated": 0, "skipped": 0,
                "total": count(db_path)}
    imported = 0
    updated = 0
    skipped = 0
    try:
        sheet = workbook.active
        columns = _excel_columns(sheet)
        email_col = columns.get("Email Muse.ai") or columns.get("Email")
        if not email_col:
            return {"imported": 0, "updated": 0, "skipped": 0,
                    "total": count(db_path)}

        def _get(values, header):
            column = columns.get(header)
            if not column or column > len(values):
                return ""
            value = values[column - 1]
            return "" if value is None else str(value)

        with _WRITE_LOCK:
            conn = _connect(db_path)
            try:
                for values in sheet.iter_rows(
                    min_row=2, values_only=True
                ):
                    values = values or ()
                    email = _get(values, "Email Muse.ai") or _get(
                        values, "Email"
                    )
                    email = email.strip()
                    if not email or "@" not in email:
                        skipped += 1
                        continue
                    new = {
                        "email": email,
                        "mail_password": _get(
                            values, "Mật khẩu hộp thư"
                        ),
                        "cookie": _get(
                            values, "Cookie phiên đăng nhập"
                        ),
                        "invite": _get(values, "Mã mời"),
                        "otp_link": _get(values, "Link lấy OTP"),
                        "created_at": _get(values, "Thời gian tạo"),
                        "cookie_taken_at": _get(
                            values, "Thời gian lấy cookie"
                        ),
                    }
                    existing = conn.execute(
                        "SELECT rowid, email, mail_password, cookie,"
                        " invite, otp_link, created_at,"
                        " cookie_taken_at FROM accounts WHERE email = ?",
                        (email,),
                    ).fetchone()
                    if existing is None:
                        conn.execute(
                            "INSERT INTO accounts (email,"
                            " mail_password, cookie, invite, otp_link,"
                            " created_at, cookie_taken_at)"
                            " VALUES (?, ?, ?, ?, ?, ?, ?)",
                            (
                                new["email"],
                                new["mail_password"],
                                new["cookie"],
                                new["invite"],
                                new["otp_link"],
                                new["created_at"],
                                new["cookie_taken_at"],
                            ),
                        )
                        imported += 1
                    else:
                        old = _row_to_dict(existing[0], existing[1:])
                        merged = _merge(old, new)
                        conn.execute(
                            "UPDATE accounts SET mail_password=?,"
                            " cookie=?, invite=?, otp_link=?,"
                            " created_at=?, cookie_taken_at=?"
                            " WHERE email=?",
                            (
                                merged["mail_password"],
                                merged["cookie"],
                                merged["invite"],
                                merged["otp_link"],
                                merged["created_at"],
                                merged["cookie_taken_at"],
                                email,
                            ),
                        )
                        updated += 1
                conn.commit()
            finally:
                conn.close()
    finally:
        workbook.close()
    return {
        "imported": imported,
        "updated": updated,
        "skipped": skipped,
        "total": count(db_path),
    }


def export_to_excel_bytes(db_path=None):
    """Toàn bộ kho -> bytes file Excel (7 cột, đúng thứ tự cũ)."""
    from openpyxl import Workbook

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Muse.ai"
    sheet.append(EXCEL_HEADERS)
    for row in list_accounts(db_path):
        sheet.append(
            [
                row["created_at"],
                row["email"],
                row["mail_password"],
                row["cookie"],
                row["invite"],
                row["otp_link"],
                row["cookie_taken_at"],
            ]
        )
    buffer = io.BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


def export_to_excel(excel_path, db_path=None):
    """Ghi toàn bộ kho ra file Excel (đường xuất). Trả về số dòng."""
    path = Path(excel_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(export_to_excel_bytes(db_path))
    return count(db_path)


def db_file_exists(db_path=None):
    """File DB đã có chưa (để biết có cần nhập lần đầu không)."""
    path = Path(db_path) if db_path is not None else Path(DB_PATH)
    return path.is_file()
