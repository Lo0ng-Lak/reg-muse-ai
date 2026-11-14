"""Lưu trữ: sheet Excel, keyring, file dữ liệu chạy.

Không có tkinter ở đây — module này dùng được từ test lẫn từ
coordinator mà không cần mở giao diện.
"""

import base64
import ctypes
import json
import re
import secrets
import threading
import time
from datetime import datetime
from pathlib import Path

import keyring
from openpyxl import Workbook, load_workbook

import mail_providers
import paths


SPREADSHEET_FILE = (
    paths.APP_DATA / "accounts.xlsx"
)

PENDING_SPREADSHEET_FILE = SPREADSHEET_FILE.with_name("accounts.pending.xlsx")

META_SPREADSHEET_FILE = (
    paths.APP_DATA / "meta_accounts.xlsx"
)

META_PENDING_SPREADSHEET_FILE = META_SPREADSHEET_FILE.with_name(
    "meta_accounts.pending.xlsx"
)

CAPSOLVER_SERVICE = "MuseAIHelper:capsolver"

CAPSOLVER_KEY_FIELD = "api_key"

SELFIE_PATH_FILE = "selfie.txt"

META_EXCEL_PATH_FILE = "meta_excel.txt"

MAIL_SERVICE = "MuseAIHelper:mail.tm"

CARD_SERVICE = "MuseAIHelper:card"

CARD_FIELDS = ("number", "expiry", "cvv", "zip")

CARD_SUCCESS_COUNT_FIELD = "card_success_count"

# Kho thẻ (nhiều thẻ): thẻ nằm ở FILE cards.json trong thư mục
# dữ liệu chạy — Credential Manager KHÔNG còn là database của
# thẻ (sư huynh dặn 2026-11-13: "đừng có coi nó là database,
# khi nào tôi chạy tool thì mới lấy"). CM ghi đè lẫn nhau khi
# trùng tên dịch vụ và đã lỗi 1783 rồi lỗi 8 khi nhập hàng loạt.
# File thẻ được mã hoá bằng DPAPI của Windows khi có (chỉ tài
# khoản Windows này mở được); không có DPAPI thì lưu dạng thường
# và giao diện cảnh báo rõ. Danh mục kho (id + nhãn + giới hạn +
# số lần dùng, KHÔNG có chữ số thẻ nào) nằm ở file vault.json;
# thẻ đang chọn tay nằm trong chính vault.json (trường active).
# Các hằng CARD_*_SERVICE dưới đây chỉ còn dùng lúc chuyển thẻ
# cũ từ CM sang file một lần duy nhất, rồi dọn mục CM mồ côi.
# vault.pending.json là mục chờ của bản d9800689 (lưu danh mục
# lỗi sau đợt nhập) — lần đọc sau tự hoà vào kho rồi xoá.
CARD_SERVICE_PREFIX = "MuseAIHelper:card:"
CARD_VAULT_INDEX_SERVICE = "MuseAIHelper:cards"
CARD_VAULT_INDEX_FIELD = "index"
CARD_VAULT_ACTIVE_FIELD = "active"
VAULT_INDEX_FILE = "vault.json"
VAULT_CARDS_FILE = "cards.json"
VAULT_PENDING_FILE = "vault.pending.json"

SPREADSHEET_HEADERS = [
    "Thời gian tạo",
    "Email Muse.ai",
    "Mật khẩu hộp thư",
    "Cookie phiên đăng nhập",
    "Mã mời",
    "Link lấy OTP",
]

META_SPREADSHEET_HEADERS = [
    "Thời gian tạo",
    "Email Meta",
    "Mật khẩu hộp thư",
    "Username Meta",
    "Nhà cung cấp",
    "Muse",
    "Link lấy OTP",
]

def _migrate_legacy_sheet(sheet):
    """Bỏ cột "Mật khẩu Muse.ai" khỏi sheet cũ.

    Cột đó là mật khẩu bịa từ thiết kế đầu — đăng ký bằng mail
    tạm + OTP nên không tồn tại mật khẩu Muse. Nếu hàng tiêu
    đề còn cột này ở vị trí 3, xoá nó khỏi cả tiêu đề lẫn mọi
    dòng dữ liệu (các cột sau dịch trái một ô).
    """
    if sheet.cell(row=1, column=3).value != "Mật khẩu Muse.ai":
        return
    for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row):
        values = [cell.value for cell in row]
        migrated = values[:2] + values[3:]
        for index, cell in enumerate(row):
            cell.value = migrated[index] if index < len(migrated) else None

def _append_spreadsheet_row(email, mail_password="", cookie="", invite="", otp_link=""):
    temporary_file = SPREADSHEET_FILE.with_name("accounts.tmp.xlsx")
    try:
        SPREADSHEET_FILE.parent.mkdir(parents=True, exist_ok=True)
        # File chờ (từ lần ghi bị khóa trước) là bản đầy đủ nhất:
        # lấy nó làm nền để không mất dòng nào đã giữ ở đó.
        base_file = (
            PENDING_SPREADSHEET_FILE
            if PENDING_SPREADSHEET_FILE.exists()
            else SPREADSHEET_FILE
        )
        if base_file.exists():
            workbook = load_workbook(base_file)
            sheet = workbook.active
            _migrate_legacy_sheet(sheet)
            # Sheet đời mail.tm còn tiêu đề cột 3 là "Mật khẩu
            # mail.tm" — cột này giờ giữ mật khẩu hộp thư của
            # bất kỳ dịch vụ nào nên đổi tên cho khớp tiêu đề
            # hiện tại (chỉ đổi tên ô tiêu đề, dữ liệu giữ
            # nguyên).
            if sheet.cell(row=1, column=3).value == "Mật khẩu mail.tm":
                sheet.cell(row=1, column=3).value = "Mật khẩu hộp thư"
            # Bổ sung tiêu đề cho các cột còn thiếu (sheet đời đầu
            # ít cột hơn) — chỉ ghi vào ô đang trống.
            for column, header in enumerate(SPREADSHEET_HEADERS, start=1):
                cell = sheet.cell(row=1, column=column)
                if not cell.value:
                    cell.value = header
        else:
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Muse.ai"
            sheet.append(SPREADSHEET_HEADERS)

        sheet.append(
            [
                datetime.now().astimezone().isoformat(timespec="seconds"),
                email,
                mail_password,
                cookie,
                invite,
                otp_link,
            ]
        )
        # accounts.xlsx đang mở trong Excel là khóa cứng trên
        # Windows (WinError 5 khi thay file): thử lại vài nhịp,
        # vẫn khóa thì giữ toàn bộ sheet vào accounts.pending.xlsx
        # — dòng không mất, lần ghi sau tự gộp vào sheet chính.
        written = False
        last_error = None
        for _ in range(3):
            try:
                workbook.save(temporary_file)
                temporary_file.replace(SPREADSHEET_FILE)
                written = True
                break
            except PermissionError as error:
                last_error = error
                time.sleep(1)
        if written:
            if PENDING_SPREADSHEET_FILE.exists():
                try:
                    PENDING_SPREADSHEET_FILE.unlink()
                except OSError:
                    pass
        else:
            pending_tmp = SPREADSHEET_FILE.with_name(
                "accounts.pending.tmp.xlsx"
            )
            try:
                workbook.save(pending_tmp)
                pending_tmp.replace(PENDING_SPREADSHEET_FILE)
            except Exception:
                pass
            raise RuntimeError(
                "File accounts.xlsx đang bị khóa (thường do đang "
                "mở trong Excel) nên chưa ghi được. Dòng này KHÔNG "
                "mất — đã giữ ở accounts.pending.xlsx và sẽ tự gộp "
                "vào sheet chính ở lần ghi thành công sau; bạn chỉ "
                "cần đóng file Excel đang mở."
            ) from last_error
    except RuntimeError:
        raise
    except Exception as error:
        raise RuntimeError(f"Không ghi được file Excel: {error}") from error
    finally:
        if temporary_file.exists():
            temporary_file.unlink()

def _last_invite_from_sheet():
    """Mã mời gần nhất trong accounts.xlsx (dòng cuối cùng có mã).

    Dùng làm mã đầu chuỗi khi ô Mã mời đang trống: tài khoản đầu
    tiên của lượt chạy sẽ quy đổi mã này thay vì bỏ qua bước quy
    đổi. Tìm cột theo tên tiêu đề "Mã mời" nên sheet đời nào
    cũng đọc được.
    """
    for path in (PENDING_SPREADSHEET_FILE, SPREADSHEET_FILE):
        try:
            if not path.exists():
                continue
            workbook = load_workbook(path, read_only=True)
            sheet = workbook.active
            column = None
            for cell in sheet[1]:
                if cell.value == "Mã mời":
                    column = cell.column
                    break
            if column is None:
                continue
            for row in range(sheet.max_row, 1, -1):
                value = sheet.cell(row=row, column=column).value
                if value:
                    code = str(value).strip().upper()
                    if re.fullmatch(r"[A-Z0-9]{6}", code):
                        return code
        except Exception:
            continue
    return ""

def _append_meta_row(email, mail_password="", username="", provider="", otp_link=""):
    """Ghi một tài khoản Meta đã tạo vào meta_accounts.xlsx.

    Cùng nhịp chống khoá file như _append_spreadsheet_row (thử
    lại 3 lần, vẫn khoá thì giữ vào meta_accounts.pending.xlsx
    Không có cột mật khẩu Meta — mật khẩu đăng ký Meta không
    được lưu (đăng nhập Muse sau này dùng mã OTP qua mail;
    người dùng xác nhận 2026-11-13: Meta chỉ để lấy mail).
    dùng mã OTP qua mail).
    """
    temporary_file = META_SPREADSHEET_FILE.with_name(
        "meta_accounts.tmp.xlsx"
    )
    try:
        META_SPREADSHEET_FILE.parent.mkdir(parents=True, exist_ok=True)
        base_file = (
            META_PENDING_SPREADSHEET_FILE
            if META_PENDING_SPREADSHEET_FILE.exists()
            else META_SPREADSHEET_FILE
        )
        if base_file.exists():
            workbook = load_workbook(base_file)
            sheet = workbook.active
            for column, header in enumerate(META_SPREADSHEET_HEADERS, start=1):
                cell = sheet.cell(row=1, column=column)
                if not cell.value:
                    cell.value = header
        else:
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Meta"
            sheet.append(META_SPREADSHEET_HEADERS)

        sheet.append(
            [
                datetime.now().astimezone().isoformat(timespec="seconds"),
                email,
                mail_password,
                username,
                provider,
                "",
                otp_link,
            ]
        )
        written = False
        last_error = None
        for _ in range(3):
            try:
                workbook.save(temporary_file)
                temporary_file.replace(META_SPREADSHEET_FILE)
                written = True
                break
            except PermissionError as error:
                last_error = error
                time.sleep(1)
        if written:
            if META_PENDING_SPREADSHEET_FILE.exists():
                try:
                    META_PENDING_SPREADSHEET_FILE.unlink()
                except OSError:
                    pass
        else:
            pending_tmp = META_SPREADSHEET_FILE.with_name(
                "meta_accounts.pending.tmp.xlsx"
            )
            try:
                workbook.save(pending_tmp)
                pending_tmp.replace(META_PENDING_SPREADSHEET_FILE)
            except Exception:
                pass
            raise RuntimeError(
                "File meta_accounts.xlsx đang bị khóa (thường do đang "
                "mở trong Excel) nên chưa ghi được. Dòng này KHÔNG "
                "mất — đã giữ ở meta_accounts.pending.xlsx và sẽ tự "
                "gộp vào sheet chính ở lần ghi thành công sau; bạn chỉ "
                "cần đóng file Excel đang mở."
            ) from last_error
    except RuntimeError:
        raise
    except Exception as error:
        raise RuntimeError(f"Không ghi được file Excel: {error}") from error
    finally:
        if temporary_file.exists():
            temporary_file.unlink()

def _load_capsolver_key():
    """CapSolver API key đã lưu ("" nếu chưa có). Không in ra log."""
    try:
        value = keyring.get_password(CAPSOLVER_SERVICE, CAPSOLVER_KEY_FIELD)
    except Exception:
        value = None
    return value or ""

def _save_capsolver_key(key):
    keyring.set_password(CAPSOLVER_SERVICE, CAPSOLVER_KEY_FIELD, key)

def _load_selfie_path_text():
    """Đường dẫn ảnh selfie đã lưu ("" nếu chưa chọn)."""
    try:
        return (
            paths.data_file(SELFIE_PATH_FILE)
            .read_text(encoding="utf-8")
            .strip()
        )
    except OSError:
        return ""

def _save_selfie_path_text(path):
    paths.data_file(SELFIE_PATH_FILE).write_text(
        path.strip() + "\n", encoding="utf-8"
    )

def _read_meta_username(suffix=""):
    """Username Meta helper vừa ghi (đọc xong xoá file để tài khoản
    sau không đọc nhầm username cũ)."""
    username_file = paths.data_file(f"meta_username{suffix}.txt")
    try:
        username = username_file.read_text(encoding="utf-8").strip()
    except OSError:
        return ""
    try:
        username_file.unlink()
    except OSError:
        pass
    return username.splitlines()[0].strip() if username else ""

# Domain hộp thư -> dịch vụ temp mail (để bước đăng nhập Muse
# suy ra dịch vụ của một tài khoản Meta cũ, khi file Excel không
# có cột "Nhà cung cấp"). mail.tm lấy danh sách domain qua API
# nên chỉ ghi nhận domain đang hoạt động đã biết (fldubai.com).
_PROVIDER_DOMAINS = {}
for _domain in mail_providers.TEMPMAIL_PLUS_DOMAINS:
    _PROVIDER_DOMAINS[_domain] = "tempmail.plus"
_PROVIDER_DOMAINS["maildrop.cc"] = "maildrop.cc"
_PROVIDER_DOMAINS["fldubai.com"] = "mail.tm"
_PROVIDER_DOMAINS["flubai.com"] = "mail.tm"
for _domain in (
    "guerrillamail.com",
    "guerrillamail.net",
    "guerrillamail.org",
    "guerrillamail.biz",
    "guerrillamail.info",
    "guerrillamail.de",
    "sharklasers.com",
):
    _PROVIDER_DOMAINS[_domain] = "guerrillamail"


RELOGIN_HEADERS = ["Email", "Cookie đăng nhập"]


def _relogin_accounts(excel_path):
    """Các mail cần đăng nhập lại, đọc từ file Excel người dùng chọn.

    Chỉ lấy cột email (Email Muse.ai, hoặc Email nếu file đặt tên
    vậy) và mật khẩu hộp thư — mật khẩu chỉ để đọc OTP, không ghi
    ra file xuất. Mail trùng thì giữ dòng đầu. File không có cột
    email hoặc không đọc được thì trả về [].
    """
    try:
        workbook = load_workbook(excel_path, read_only=True, data_only=True)
    except Exception:
        return []
    try:
        sheet = workbook.active
        columns = {}
        for cell in sheet[1]:
            if isinstance(cell.value, str):
                columns[cell.value.strip()] = cell.column
        email_col = columns.get("Email Muse.ai") or columns.get("Email")
        if not email_col:
            return []
        secret_col = columns.get("Mật khẩu hộp thư")
        seen = set()
        rows = []
        for row in sheet.iter_rows(min_row=2, values_only=True):
            email = ""
            if row and len(row) >= email_col and row[email_col - 1]:
                email = str(row[email_col - 1]).strip()
            if "@" not in email:
                continue
            key = email.lower()
            if key in seen:
                continue
            seen.add(key)
            secret = ""
            if secret_col and len(row) >= secret_col and row[secret_col - 1]:
                secret = str(row[secret_col - 1]).strip()
            rows.append(
                {
                    "email": email,
                    "secret": secret,
                    "provider": _provider_for_email(email),
                }
            )
        return rows
    finally:
        workbook.close()


def _write_relogin_cookie(output_path, email, cookie):
    """Ghi một dòng Email + Cookie đăng nhập vào file xuất.

    File chưa có thì tạo mới. Mail đã có trong file thì thay
    cookie, không thêm dòng trùng. Không ghi mật khẩu hộp thư.
    """
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        workbook = load_workbook(path)
        sheet = workbook.active
    else:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Cookie"
        sheet.append(list(RELOGIN_HEADERS))
    headers = {}
    for cell in sheet[1]:
        if isinstance(cell.value, str):
            headers[cell.value.strip()] = cell.column
    email_col = headers.get("Email")
    cookie_col = headers.get("Cookie đăng nhập")
    if email_col is None:
        email_col = 1
        sheet.cell(row=1, column=email_col).value = "Email"
    if cookie_col is None:
        cookie_col = 2 if email_col != 2 else 3
        sheet.cell(row=1, column=cookie_col).value = "Cookie đăng nhập"
    target = None
    for row_index in range(2, (sheet.max_row or 1) + 1):
        value = sheet.cell(row=row_index, column=email_col).value
        if str(value or "").strip().lower() == email.strip().lower():
            target = row_index
            break
    if target is None:
        target = (sheet.max_row or 1) + 1
    sheet.cell(row=target, column=email_col).value = email
    sheet.cell(row=target, column=cookie_col).value = cookie
    last_error = None
    for _ in range(3):
        try:
            workbook.save(path)
            return
        except PermissionError as error:
            last_error = error
            time.sleep(1)
    raise RuntimeError(
        f"Không ghi được {path.name} (đang bị khóa, thường do đang "
        "mở trong Excel). Đóng file đó rồi chạy lại mail này."
    ) from last_error


def _provider_for_email(email):
    """Dịch vụ temp mail của một địa chỉ (theo domain hộp thư).

    Không nhận ra domain thì về dịch vụ mặc định (tempmail.plus)
    — dịch vụ chủ tool dùng nhiều nhất.
    """
    domain = (email or "").rsplit("@", 1)[-1].strip().lower()
    return _PROVIDER_DOMAINS.get(domain, mail_providers.DEFAULT_PROVIDER)


def _meta_sheet_columns(sheet):
    """Vị trí các cột của sheet Meta (theo tên tiêu đề, 1-based).

    Thiếu cột nào thì cột đó vắng khỏi dict — file cũ (4 cột)
    không có "Nhà cung cấp"/"Muse" vẫn đọc được.
    """
    columns = {}
    for cell in sheet[1]:
        if isinstance(cell.value, str):
            columns[cell.value.strip()] = cell.column
    return columns


def _pending_meta_logins(excel_path):
    """Tài khoản Meta trong file Excel chưa đăng nhập Muse.

    Đọc file người dùng CHỌN ở tab Meta (mặc định là
    meta_accounts.xlsx trong AppData, nhưng có thể là bản sao
    khác). Một dòng là "chưa đăng nhập" khi cột "Muse" trống
    (file cũ không có cột này thì mọi dòng đều chưa đăng nhập),
    hoặc khi cột chỉ giữ dấu "đang đăng nhập" đã quá 30 phút
    (luồng cũ bị dừng/kẹt giữa chừng). Trả về list dict {email,
    secret, provider, username} theo thứ tự trong file. File
    không có/không đọc được -> [].
    """
    try:
        workbook = load_workbook(excel_path, read_only=True)
    except Exception:
        return []
    try:
        sheet = workbook.active
        columns = _meta_sheet_columns(sheet)
        email_col = columns.get("Email Meta")
        if not email_col:
            return []
        secret_col = columns.get("Mật khẩu hộp thư")
        username_col = columns.get("Username Meta")
        provider_col = columns.get("Nhà cung cấp")
        muse_col = columns.get("Muse")
        valid_providers = {pid for pid, _ in mail_providers.PROVIDERS}
        pending = []
        for row in sheet.iter_rows(min_row=2, values_only=True):
            email = (row[email_col - 1] or "").strip() if len(row) >= email_col else ""
            if not email:
                continue
            if muse_col and len(row) >= muse_col:
                muse_value = row[muse_col - 1]
                if muse_value and not _login_mark_is_stale(muse_value):
                    continue
            secret = ""
            if secret_col and len(row) >= secret_col:
                secret = str(row[secret_col - 1] or "")
            username = ""
            if username_col and len(row) >= username_col:
                username = str(row[username_col - 1] or "")
            provider = ""
            if provider_col and len(row) >= provider_col:
                provider = str(row[provider_col - 1] or "").strip()
            if provider not in valid_providers:
                provider = _provider_for_email(email)
            pending.append(
                {
                    "email": email,
                    "secret": secret,
                    "provider": provider,
                    "username": username,
                }
            )
        return pending
    finally:
        workbook.close()


# Đánh dấu "đang đăng nhập" trong cột Muse: "ĐANG ĐĂNG NHẬP <thời
# gian ISO>". Một dấu đang-đăng-nhập quá cũ (luồng bị dừng/kẹt
# giữa chừng, không kịp xoá) thì coi như chưa đăng nhập để lượt
# sau lấy lại — một lượt đăng nhập tối đa 900 giây (giới hạn của
# helper) nên 30 phút là đủ xa.
_LOGIN_IN_PROGRESS_PREFIX = "ĐANG ĐĂNG NHẬP "
_LOGIN_IN_PROGRESS_STALE_MINUTES = 30


def _login_mark_is_stale(value):
    """Dấu "đang đăng nhập" đã quá hạn chưa. Chỉ dấu
    đang-đăng-nhập mới có hạn; dấu đã-đăng-nhập (thời gian hoàn
    tất) không bao giờ hết hạn. Dấu lạ (không đúng định dạng) thì
    coi như đã đăng nhập — an toàn hơn là lấy trùng."""
    if not isinstance(value, str):
        return False
    if not value.startswith(_LOGIN_IN_PROGRESS_PREFIX):
        return False
    try:
        marked_at = datetime.fromisoformat(
            value[len(_LOGIN_IN_PROGRESS_PREFIX):].strip()
        )
    except ValueError:
        return False
    if marked_at.tzinfo is None:
        return False
    age = datetime.now().astimezone() - marked_at
    return age.total_seconds() > _LOGIN_IN_PROGRESS_STALE_MINUTES * 60


def _save_workbook_retry(workbook, excel_path):
    """Lưu workbook, thử lại 3 lần nếu file đang bị khóa (mở trong
    Excel). Vẫn không được thì lỗi RuntimeError."""
    last_error = None
    for _ in range(3):
        try:
            workbook.save(excel_path)
            return
        except PermissionError as error:
            last_error = error
            time.sleep(1)
    raise RuntimeError(
        f"Không ghi được {Path(excel_path).name} (đang bị khóa) — "
        "dòng Meta vẫn chưa được đánh dấu."
    ) from last_error


def _set_meta_muse_cell(excel_path, email, value):
    """Ghi một giá trị vào cột "Muse" của dòng có email cho trước.
    Thêm cột "Muse" nếu file cũ chưa có. Không tìm thấy dòng hoặc
    không có cột Email Meta thì không làm gì."""
    workbook = load_workbook(excel_path)
    try:
        sheet = workbook.active
        columns = _meta_sheet_columns(sheet)
        email_col = columns.get("Email Meta")
        if not email_col:
            return
        muse_col = columns.get("Muse")
        if not muse_col:
            muse_col = sheet.max_column + 1
            sheet.cell(row=1, column=muse_col, value="Muse")
        target = (email or "").strip().lower()
        for row_index in range(2, sheet.max_row + 1):
            cell_value = sheet.cell(row=row_index, column=email_col).value
            if isinstance(cell_value, str) and cell_value.strip().lower() == target:
                sheet.cell(row=row_index, column=muse_col, value=value)
                break
        else:
            return
        _save_workbook_retry(workbook, excel_path)
    finally:
        workbook.close()


def _mark_meta_logged_in(excel_path, email):
    """Đánh dấu một dòng Meta là đã đăng nhập Muse (cột "Muse" =
    thời điểm đăng nhập). Thêm cột "Muse" vào tiêu đề nếu file
    cũ chưa có. Không tìm thấy dòng thì không làm gì.
    """
    _set_meta_muse_cell(
        excel_path,
        email,
        datetime.now().astimezone().isoformat(timespec="seconds"),
    )


def _mark_meta_logging_in(excel_path, email):
    """Đánh dấu một dòng Meta là ĐANG đăng nhập Muse (cột "Muse" =
    "ĐANG ĐĂNG NHẬP <thời gian>"). Lượt đăng nhập sau bỏ qua dòng
    này trừ khi dấu đã quá 30 phút (luồng cũ bị dừng/kẹt giữa
    chừng). Thêm cột "Muse" nếu file cũ chưa có. Không tìm thấy
    dòng thì không làm gì."""
    _set_meta_muse_cell(
        excel_path,
        email,
        _LOGIN_IN_PROGRESS_PREFIX
        + datetime.now().astimezone().isoformat(timespec="seconds"),
    )


def _clear_meta_login_mark(excel_path, email):
    """Xoá dấu "đang đăng nhập" của một dòng Meta (đăng nhập thất
    bại thì lượt sau lấy lại được ngay). CHỈ xoá khi ô đang giữ
    dấu đang-đăng-nhập — không bao giờ xoá dấu đã-đăng-nhập
    (thời gian hoàn tất)."""
    workbook = load_workbook(excel_path)
    try:
        sheet = workbook.active
        columns = _meta_sheet_columns(sheet)
        email_col = columns.get("Email Meta")
        muse_col = columns.get("Muse")
        if not email_col or not muse_col:
            return
        target = (email or "").strip().lower()
        for row_index in range(2, sheet.max_row + 1):
            cell_value = sheet.cell(row=row_index, column=email_col).value
            if isinstance(cell_value, str) and cell_value.strip().lower() == target:
                current = sheet.cell(row=row_index, column=muse_col).value
                if isinstance(current, str) and current.startswith(
                    _LOGIN_IN_PROGRESS_PREFIX
                ):
                    # Gán qua .value: cell(..., value=None) là no-op
                    # trong openpyxl (None nghĩa là "không đưa giá
                    # trị"), không xoá được ô.
                    sheet.cell(row=row_index, column=muse_col).value = None
                    _save_workbook_retry(workbook, excel_path)
                return
    finally:
        workbook.close()


def _load_meta_excel_path_text():
    """Đường dẫn file Excel Meta đã chọn ("" nếu chưa chọn)."""
    try:
        return (
            paths.data_file(META_EXCEL_PATH_FILE)
            .read_text(encoding="utf-8")
            .strip()
        )
    except OSError:
        return ""


def _save_meta_excel_path_text(path):
    paths.data_file(META_EXCEL_PATH_FILE).write_text(
        path.strip() + "\n", encoding="utf-8"
    )


def _load_relogin_source_path():
    try:
        return (
            paths.data_file("relogin_source.txt")
            .read_text(encoding="utf-8")
            .strip()
        )
    except OSError:
        return ""


def _save_relogin_source_path(path):
    paths.data_file("relogin_source.txt").write_text(
        path.strip() + "\n", encoding="utf-8"
    )


def _load_relogin_output_path():
    try:
        return (
            paths.data_file("relogin_output.txt")
            .read_text(encoding="utf-8")
            .strip()
        )
    except OSError:
        return ""


def _save_relogin_output_path(path):
    paths.data_file("relogin_output.txt").write_text(
        path.strip() + "\n", encoding="utf-8"
    )


def _load_card_from_keyring():
    """Thẻ đơn cũ (nếu có) — chỉ đọc từ Credential Manager, dùng
    lúc chuyển sang file lần đầu (xem _migrate_legacy_card)."""
    card = {}
    for field in CARD_FIELDS:
        try:
            value = keyring.get_password(CARD_SERVICE, field)
        except Exception:
            value = None
        if value:
            card[field] = value
    return card if card.get("number") else {}

def _load_card_success_count():
    """Số tài khoản thành công đã dùng thẻ đơn cũ (0 nếu chưa có).

    Chỉ là một con số đếm, không chứa chữ số nào của thẻ.
    """
    try:
        value = keyring.get_password(CARD_SERVICE, CARD_SUCCESS_COUNT_FIELD)
    except Exception:
        value = None
    try:
        return max(0, int(value or 0))
    except (TypeError, ValueError):
        return 0

def _delete_card_from_keyring():
    for field in CARD_FIELDS:
        try:
            keyring.delete_password(CARD_SERVICE, field)
        except Exception:
            pass

def _vault_service(card_id):
    """Tên dịch vụ Credential Manager của một thẻ trong kho — chỉ
    còn dùng lúc chuyển thẻ cũ sang file và lúc dọn mục mồ côi."""
    return f"{CARD_SERVICE_PREFIX}{card_id}"

def _new_card_id(existing_ids):
    """Id thẻ mới: "card-" + 8 ký tự hex ngẫu nhiên, không trùng
    với id đã có. Id này không suy ra được từ số thẻ."""
    taken = set(existing_ids or ())
    for _ in range(10):
        candidate = "card-" + secrets.token_hex(4)
        if candidate not in taken:
            return candidate
    raise RuntimeError("Không tạo được id thẻ mới (trùng lặp liên tục).")

_VAULT_INDEX_LOCK = threading.RLock()


# --- DPAPI: mã hoá file thẻ bằng khoá của tài khoản Windows --------
#
# Thẻ nằm ở FILE (cards.json), không còn ở Credential Manager:
# CM ghi đè lẫn nhau khi trùng tên dịch vụ (keyring
# WinVaultKeyring: mục trùng tên bị chuyển sang dạng gộp
# {username}@{service}) và đã lỗi 1783 rồi lỗi 8 khi nhập hàng
# loạt — sư huynh dặn 2026-11-13 "đừng có coi nó là database,
# khi nào tôi chạy tool thì mới lấy". Tool đọc file lúc chạy; CM
# chỉ còn là nguồn để chuyển thẻ cũ sang file một lần duy nhất,
# rồi xoá.
#
# DPAPI (CryptProtectData) mã hoá bằng khoá của tài khoản Windows
# đang chạy tool — chỉ tài khoản đó mở được file. Thử pywin32
# (win32crypt) trước, rồi ctypes (thư viện chuẩn — keyring trên
# Windows chỉ cần pywin32-ctypes nên pywin32 thường KHÔNG có
# sẵn); không có DPAPI thì lưu dạng thường và giao diện phải cảnh
# báo rõ (xem _cards_protection_warning) — không bao giờ lặng lẽ
# hạ cấp.

_UNSET = object()
_DPAPI_BACKEND = _UNSET  # None | ("win32crypt", module) | ("ctypes", None)
_WIN32CRED = _UNSET  # None | module win32cred (để liệt kê CM)
_LAST_CARDS_READ_ERROR = ""
_MIGRATION_DONE = False
_LAST_MIGRATION = None


def _dpapi_backend():
    """Backend DPAPI khả dụng trên máy này (nhớ trong tiến trình;
    test đặt lại _DPAPI_BACKEND để đổi backend)."""
    global _DPAPI_BACKEND
    if _DPAPI_BACKEND is _UNSET:
        backend = None
        try:
            import win32crypt

            backend = ("win32crypt", win32crypt)
        except Exception:
            try:
                ctypes.windll.crypt32  # AttributeError khi không phải Windows
                backend = ("ctypes", None)
            except Exception:
                backend = None
        _DPAPI_BACKEND = backend
    return _DPAPI_BACKEND


def _dpapi_protect(data):
    """Mã hoá bytes bằng DPAPI. Trả về bytes đã mã hoá, hoặc None
    khi máy này không có DPAPI. Raise khi DPAPI lỗi."""
    backend = _dpapi_backend()
    if backend is None:
        return None
    kind, module = backend
    if kind == "win32crypt":
        return module.CryptProtectData(
            data, "MuseAIHelper card vault", None, None, None, 0
        )
    return _dpapi_protect_ctypes(data)


def _dpapi_unprotect(blob):
    """Giải mã bytes đã mã hoá bằng DPAPI. Raise khi không mở
    được (không có DPAPI / sai tài khoản Windows)."""
    backend = _dpapi_backend()
    if backend is None:
        raise RuntimeError("máy này không có DPAPI")
    kind, module = backend
    if kind == "win32crypt":
        _description, data = module.CryptUnprotectData(
            blob, None, None, None, 0
        )
        return data
    return _dpapi_unprotect_ctypes(blob)


class _DataBlob(ctypes.Structure):
    """DATA_BLOB của CryptProtectData/CryptUnprotectData."""

    _fields_ = [
        ("cbData", ctypes.c_uint32),
        ("pbData", ctypes.POINTER(ctypes.c_char)),
    ]


def _dpapi_protect_ctypes(data):
    crypt32 = ctypes.windll.crypt32
    buf_in = ctypes.create_string_buffer(data, len(data))
    blob_in = _DataBlob(
        len(data), ctypes.cast(buf_in, ctypes.POINTER(ctypes.c_char))
    )
    blob_out = _DataBlob()
    if not crypt32.CryptProtectData(
        ctypes.byref(blob_in),
        "MuseAIHelper card vault",
        None,
        None,
        None,
        0x01,  # CRYPTPROTECT_UI_FORBIDDEN: không hiện hộp thoại
        ctypes.byref(blob_out),
    ):
        raise ctypes.WinError()
    try:
        return ctypes.string_at(blob_out.pbData, blob_out.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(blob_out.pbData)


def _dpapi_unprotect_ctypes(blob):
    crypt32 = ctypes.windll.crypt32
    buf_in = ctypes.create_string_buffer(blob, len(blob))
    blob_in = _DataBlob(
        len(blob), ctypes.cast(buf_in, ctypes.POINTER(ctypes.c_char))
    )
    blob_out = _DataBlob()
    description = ctypes.c_wchar_p()
    if not crypt32.CryptUnprotectData(
        ctypes.byref(blob_in),
        ctypes.byref(description),
        None,
        None,
        None,
        0x01,
        ctypes.byref(blob_out),
    ):
        raise ctypes.WinError()
    try:
        return ctypes.string_at(blob_out.pbData, blob_out.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(blob_out.pbData)


# --- File thẻ (cards.json): số thẻ + hạn + CVV + ZIP ---------------

def _cards_file_path():
    return paths.data_file(VAULT_CARDS_FILE)


def _decode_card_entry(blob, protected):
    """Giải một mục thẻ trong cards.json -> dict 4 trường, hoặc
    None khi mục hỏng/không mở được (sai tài khoản Windows, file
    chép sang máy khác)."""
    global _LAST_CARDS_READ_ERROR
    if protected:
        if not isinstance(blob, str):
            return None
        try:
            raw = base64.b64decode(blob.encode("ascii"))
            card = json.loads(_dpapi_unprotect(raw).decode("utf-8"))
        except Exception:
            _LAST_CARDS_READ_ERROR = (
                f"không mở được thẻ đã mã hoá trong file {VAULT_CARDS_FILE} "
                "(DPAPI không có trên máy này hoặc file chép từ máy khác)"
            )
            return None
    else:
        card = blob
    if not isinstance(card, dict) or not card.get("number"):
        return None
    return {field: str(card.get(field) or "") for field in CARD_FIELDS}


def _read_cards_file():
    """Đọc file thẻ -> (dict id -> card dict, protected True/False/
    None). None = chưa có file. File hỏng JSON -> ({}, None) và ghi
    lỗi vào _LAST_CARDS_READ_ERROR để giao diện báo."""
    global _LAST_CARDS_READ_ERROR
    try:
        raw = _cards_file_path().read_text(encoding="utf-8")
    except OSError:
        return {}, None
    try:
        data = json.loads(raw)
    except (TypeError, ValueError):
        _LAST_CARDS_READ_ERROR = (
            f"file {VAULT_CARDS_FILE} hỏng (không đọc được JSON)"
        )
        return {}, None
    _LAST_CARDS_READ_ERROR = ""
    if not isinstance(data, dict):
        return {}, None
    protected = data.get("protected")
    raw_cards = data.get("cards")
    if not isinstance(raw_cards, dict):
        return {}, None
    cards = {}
    for card_id, blob in raw_cards.items():
        card = _decode_card_entry(blob, bool(protected))
        if card:
            cards[str(card_id)] = card
    return cards, bool(protected) if protected is not None else None


def _write_cards_file(cards):
    """Ghi file thẻ: ghi file tạm rồi đổi tên (không bao giờ để
    lại file dở dang), rồi đọc lại kiểm chứng. DPAPI có thì mỗi
    thẻ là một blob đã mã hoá (base64) — file không chứa chữ số
    thẻ nào ở dạng đọc được; không có DPAPI thì lưu dạng thường
    (giao diện phải cảnh báo rõ, xem _cards_protection_warning).
    Raise RuntimeError khi ghi hoặc kiểm chứng thất bại sau 3 lần
    thử."""
    backend = _dpapi_backend()
    entries = {}
    for card_id, card in cards.items():
        clean = {field: str(card.get(field, "")) for field in CARD_FIELDS}
        if backend is not None:
            blob = _dpapi_protect(
                json.dumps(clean, ensure_ascii=False).encode("utf-8")
            )
            entries[str(card_id)] = base64.b64encode(blob).decode("ascii")
        else:
            entries[str(card_id)] = clean
    payload = json.dumps(
        {"protected": backend is not None, "cards": entries},
        ensure_ascii=False,
    )
    target = _cards_file_path()
    tmp = target.with_name(target.name + ".tmp")
    last_error = None
    for _ in range(3):
        try:
            tmp.write_text(payload, encoding="utf-8")
            tmp.replace(target)
        except OSError as error:
            last_error = error
            time.sleep(0.2)
            continue
        back, _ = _read_cards_file()
        if set(back) == set(cards):
            return
        last_error = RuntimeError("đọc lại file thẻ không khớp")
        time.sleep(0.2)
    raise RuntimeError(
        f"Không lưu được file thẻ {VAULT_CARDS_FILE} ({last_error})"
    )


def _cards_protection_warning():
    """Cảnh báo khi file thẻ đang lưu KHÔNG mã hoá (máy này không
    có DPAPI) hoặc file thẻ đọc lỗi. None khi file được DPAPI bảo
    vệ, chưa có thẻ nào, hoặc chưa có file."""
    if _LAST_CARDS_READ_ERROR:
        return (
            "CẢNH BÁO: " + _LAST_CARDS_READ_ERROR + " — kiểm tra lại file "
            f"{VAULT_CARDS_FILE} trước khi chạy."
        )
    cards, protected = _read_cards_file()
    if not cards:
        return None
    if protected is False:
        return (
            f"CẢNH BÁO: file thẻ {VAULT_CARDS_FILE} đang lưu KHÔNG mã hoá "
            "(máy này không có DPAPI của Windows) — ai mở được file là "
            "đọc được số thẻ. Chỉ chạy tool trên máy của bạn, đừng copy "
            "hay đồng bộ file này đi nơi khác."
        )
    return None


# --- Danh mục kho (vault.json): nhãn + giới hạn + số lần dùng ------

def _vault_index_path():
    return paths.data_file(VAULT_INDEX_FILE)


def _vault_pending_path():
    return paths.data_file(VAULT_PENDING_FILE)


def _clean_vault_entries(entries):
    """Chuẩn hoá danh mục kho thô thành list {"id", "label",
    "limit", "used"} (giới hạn kẹp 0–100, số lần dùng >= 0,
    nhãn trống thành "Thẻ"). Mục thiếu id thì bỏ."""
    cleaned = []
    if not isinstance(entries, list):
        return cleaned
    for entry in entries:
        if not isinstance(entry, dict) or not entry.get("id"):
            continue
        try:
            limit = max(0, min(int(entry.get("limit") or 0), 100))
        except (TypeError, ValueError):
            limit = 0
        try:
            used = max(0, int(entry.get("used") or 0))
        except (TypeError, ValueError):
            used = 0
        cleaned.append(
            {
                "id": str(entry["id"]),
                "label": str(entry.get("label") or "").strip() or "Thẻ",
                "limit": limit,
                "used": used,
            }
        )
    return cleaned


def _read_vault_file():
    """Đọc vault.json -> {"active": str, "cards": [entry...]} hoặc
    None khi file thiếu/hỏng. Chấp nhận cả định dạng cũ (list
    thuần) — active mặc định ""."""
    try:
        raw = _vault_index_path().read_text(encoding="utf-8")
    except OSError:
        return None
    try:
        data = json.loads(raw)
    except (TypeError, ValueError):
        return None
    if isinstance(data, list):
        return {"active": "", "cards": _clean_vault_entries(data)}
    if not isinstance(data, dict):
        return None
    return {
        "active": str(data.get("active") or "").strip(),
        "cards": _clean_vault_entries(data.get("cards")),
    }


def _write_vault_file(data):
    """Ghi vault.json: ghi file tạm rồi đổi tên, rồi đọc lại kiểm
    chứng. Raise RuntimeError khi ghi hoặc kiểm chứng thất bại
    sau 3 lần thử."""
    cleaned = {
        "active": str(data.get("active") or "").strip(),
        "cards": _clean_vault_entries(data.get("cards")),
    }
    payload = json.dumps(cleaned, ensure_ascii=False)
    target = _vault_index_path()
    tmp = target.with_name(target.name + ".tmp")
    last_error = None
    with _VAULT_INDEX_LOCK:
        for _ in range(3):
            try:
                tmp.write_text(payload, encoding="utf-8")
                tmp.replace(target)
            except OSError as error:
                last_error = error
                time.sleep(0.2)
                continue
            back = _read_vault_file()
            if (
                back is not None
                and back["active"] == cleaned["active"]
                and [e["id"] for e in back["cards"]]
                == [e["id"] for e in cleaned["cards"]]
            ):
                return
            last_error = RuntimeError("đọc lại danh mục kho không khớp")
            time.sleep(0.2)
    raise RuntimeError(
        "Không lưu được danh mục kho thẻ vào file "
        f"{VAULT_INDEX_FILE} ({last_error})"
    )


def _migrate_vault_index_to_file():
    """Chuyển danh mục kho từ Credential Manager sang file
    vault.json (một lần duy nhất, lúc nâng cấp). File đã có nội
    dung hợp lệ thì thôi. CM còn danh mục cũ thì ghi sang file
    rồi xoá mục CM (best-effort)."""
    if _read_vault_file() is not None:
        return
    try:
        raw = keyring.get_password(
            CARD_VAULT_INDEX_SERVICE, CARD_VAULT_INDEX_FIELD
        )
    except Exception:
        raw = None
    if not raw:
        return
    try:
        data = json.loads(raw)
    except (TypeError, ValueError):
        return
    if not isinstance(data, list):
        return
    entries = _clean_vault_entries(data)
    try:
        _save_vault_index(entries)
    except Exception:
        return
    try:
        keyring.delete_password(
            CARD_VAULT_INDEX_SERVICE, CARD_VAULT_INDEX_FIELD
        )
    except Exception:
        pass


def _read_vault_pending_file():
    """Danh mục đang chờ hoà vào kho (bản d9800689 giữ lại khi
    lưu danh mục lỗi sau đợt nhập) — xem _load_vault_index_raw."""
    with _VAULT_INDEX_LOCK:
        try:
            raw = _vault_pending_path().read_text(encoding="utf-8")
        except OSError:
            return []
        try:
            data = json.loads(raw)
        except (TypeError, ValueError):
            return []
        return _clean_vault_entries(data) if isinstance(data, list) else []


def _load_vault_index_raw():
    """Danh mục kho từ file (đã hoà mục chờ trong
    vault.pending.json nếu có), KHÔNG chạy chuyển đổi từ
    Credential Manager — dùng nội bộ cho _ensure_cards_migrated
    và _load_vault_index."""
    data = _read_vault_file()
    entries = list(data["cards"]) if data else []
    pending = _read_vault_pending_file()
    if pending:
        known = {entry["id"] for entry in entries}
        merged = list(entries)
        for entry in pending:
            if entry["id"] not in known:
                merged.append(entry)
                known.add(entry["id"])
        if len(merged) != len(entries):
            entries = merged
            try:
                _save_vault_index(entries)
            except Exception:
                pass
        try:
            _vault_pending_path().unlink()
        except OSError:
            pass
    return entries


def _save_vault_index(entries):
    """Ghi danh mục kho xuống file vault.json (giữ nguyên thẻ
    đang chọn tay). Raise RuntimeError khi không ghi/kiểm chứng
    được — người gọi phải xử lý để không mất thẻ đã ghi."""
    data = _read_vault_file() or {"active": "", "cards": []}
    data["cards"] = _clean_vault_entries(entries)
    _write_vault_file(data)


def _load_vault_active():
    """Id thẻ đang chọn tay ("" là tự động xoay theo thứ tự kho)."""
    _ensure_cards_migrated()
    data = _read_vault_file()
    if not data:
        return ""
    return data.get("active", "")


def _save_vault_active(card_id):
    _ensure_cards_migrated()
    data = _read_vault_file() or {"active": "", "cards": []}
    data["active"] = card_id or ""
    _write_vault_file(data)


# --- Đọc Credential Manager (chỉ để chuyển sang file) ---------------

def _read_cm_vault_card(card_id):
    """Một thẻ trong kho còn nằm ở Credential Manager (định dạng
    cũ: 4 trường riêng) — {} khi không có. Chỉ dùng lúc chuyển
    sang file."""
    service = _vault_service(card_id)
    card = {}
    for field in CARD_FIELDS:
        try:
            value = keyring.get_password(service, field)
        except Exception:
            value = None
        if value:
            card[field] = value
    return card if card.get("number") else {}


def _read_cm_vault_card_success_count(card_id):
    """Số tài khoản thành công của một thẻ còn nằm ở Credential
    Manager (0 khi không có). Chỉ dùng lúc chuyển sang file."""
    try:
        value = keyring.get_password(
            _vault_service(card_id), CARD_SUCCESS_COUNT_FIELD
        )
    except Exception:
        value = None
    try:
        return max(0, int(value or 0))
    except (TypeError, ValueError):
        return 0


def _delete_cm_vault_card(card_id):
    """Xoá mục Credential Manager của một thẻ trong kho (cả 4
    trường + bộ đếm). Best-effort — dùng sau khi thẻ đã nằm chắc
    trong file."""
    service = _vault_service(card_id)
    for field in CARD_FIELDS + (CARD_SUCCESS_COUNT_FIELD,):
        try:
            keyring.delete_password(service, field)
        except Exception:
            pass


def _load_win32cred():
    """Module win32cred để liệt kê Credential Manager (keyring
    không có API liệt kê). Thử win32ctypes trước rồi pywin32 —
    giống cách keyring tự chọn backend trên Windows. None khi
    không có (không phải Windows)."""
    global _WIN32CRED
    if _WIN32CRED is _UNSET:
        module = None
        try:
            from win32ctypes.pywin32 import win32cred as module
        except Exception:
            try:
                import win32cred as module
            except Exception:
                module = None
        _WIN32CRED = module
    return _WIN32CRED


def _cm_service_from_target(target):
    """Tên dịch vụ CM từ một TargetName liệt kê được. keyring
    WinVaultKeyring lưu mục trùng tên dịch vụ dưới dạng gộp
    {username}@{service} nên TargetName có thể là dạng gộp —
    bóc phần sau "@" ra. Không phải dịch vụ kho thẻ thì None."""
    if not target:
        return None
    if target.startswith(CARD_SERVICE_PREFIX):
        return target
    if "@" in target:
        candidate = target.split("@", 1)[1]
        if candidate.startswith(CARD_SERVICE_PREFIX):
            return candidate
    return None


def _cleanup_cm_card_orphans(known_ids):
    """Dọn các mục CM của kho thẻ không còn trong danh mục (thẻ
    mồ côi từ đợt nhập lỗi trước — mục CM ghi được mà danh mục
    không lưu). Chỉ xoá mục khớp tiền tố kho thẻ; không đụng
    dịch vụ khác (capsolver, mail, thẻ đơn cũ). Trả về dict
    {removed_services, removed_credentials, enumerated, error}."""
    win32cred = _load_win32cred()
    empty = {
        "removed_services": 0,
        "removed_credentials": 0,
        "enumerated": False,
        "error": "",
    }
    if win32cred is None:
        empty["error"] = "không liệt kê được Credential Manager trên máy này"
        return empty
    enumerate_fn = getattr(win32cred, "CredEnumerate", None)
    if enumerate_fn is None:
        empty["error"] = "không liệt kê được Credential Manager trên máy này"
        return empty
    try:
        creds = enumerate_fn(None, 0)
    except Exception as error:
        empty["error"] = _short_error_text(error)
        return empty
    services = set()
    for cred in creds or []:
        target = cred.get("TargetName") if isinstance(cred, dict) else None
        service = _cm_service_from_target(target or "")
        if service:
            services.add(service)
    removed_services = 0
    removed_credentials = 0
    for service in sorted(services):
        card_id = service[len(CARD_SERVICE_PREFIX):]
        if card_id in known_ids:
            continue
        deleted_any = False
        for field in CARD_FIELDS + (CARD_SUCCESS_COUNT_FIELD,):
            try:
                keyring.delete_password(service, field)
                removed_credentials += 1
                deleted_any = True
            except Exception:
                pass
        if deleted_any:
            removed_services += 1
    return {
        "removed_services": removed_services,
        "removed_credentials": removed_credentials,
        "enumerated": True,
        "error": "",
    }


# --- Chuyển thẻ từ Credential Manager sang file (một lần) -----------

def _migrate_legacy_card():
    """Nhập thẻ đơn cũ (dịch vụ MuseAIHelper:card) vào kho, một
    lần duy nhất: khi danh mục kho còn trống mà thẻ đơn cũ còn
    đó, chuyển nó thành thẻ đầu tiên của kho (nhãn "Thẻ 1",
    giới hạn 0 = không giới hạn, giữ nguyên bộ đếm thành công),
    rồi xoá thẻ đơn cũ để file là nơi duy nhất giữ thẻ. Chỉ xoá
    bản cũ sau khi bản mới đã ghi xong. Trả về True nếu đã nhập.
    """
    if _read_vault_file() is not None:
        return False
    legacy = _load_card_from_keyring()
    if not legacy:
        return False
    legacy_count = _load_card_success_count()
    card_id = _new_card_id(())
    cards, _ = _read_cards_file()
    cards[card_id] = legacy
    try:
        _write_cards_file(cards)
        _save_vault_index(
            [{"id": card_id, "label": "Thẻ 1", "limit": 0, "used": legacy_count}]
        )
    except Exception:
        return False
    _delete_card_from_keyring()
    try:
        keyring.delete_password(CARD_SERVICE, CARD_SUCCESS_COUNT_FIELD)
    except Exception:
        pass
    return True


def _migrate_vault_active_to_file():
    """Chuyển thẻ đang chọn tay từ Credential Manager sang
    vault.json (một lần duy nhất). File đã có lựa chọn thì thôi."""
    data = _read_vault_file()
    if data is not None and data.get("active"):
        return
    try:
        value = keyring.get_password(
            CARD_VAULT_INDEX_SERVICE, CARD_VAULT_ACTIVE_FIELD
        )
    except Exception:
        value = None
    if not (value or "").strip():
        return
    data = data or {"active": "", "cards": []}
    data["active"] = value.strip()
    try:
        _write_vault_file(data)
    except Exception:
        return
    try:
        keyring.delete_password(
            CARD_VAULT_INDEX_SERVICE, CARD_VAULT_ACTIVE_FIELD
        )
    except Exception:
        pass


def _ensure_cards_migrated():
    """Một lần mỗi tiến trình: chuyển toàn bộ thẻ từ Credential
    Manager sang file cards.json + dọn các mục CM mồ côi của kho
    thẻ. Chạy lúc đọc kho lần đầu (giao diện chính và helper đều
    đi qua đây); sau lần đầu không còn gì để chuyển thì không ghi
    gì thêm. Kết quả xem _last_migration_report(). Lỗi trong lúc
    chuyển KHÔNG được phá lần mở app: lỗi ghi vào report["error"]
    để giao diện hiện trong log, thẻ vẫn còn ở chỗ cũ và lần sau
    thử lại."""
    global _MIGRATION_DONE, _LAST_MIGRATION
    if _MIGRATION_DONE:
        return
    _MIGRATION_DONE = True
    report = {
        "migrated": 0,
        "legacy": False,
        "counts_migrated": 0,
        "orphans_removed": 0,
        "orphan_credentials_removed": 0,
        "enumerated": False,
        "cleanup_error": "",
        "error": "",
    }
    try:
        _ensure_cards_migrated_inner(report)
    except Exception as error:
        # Chuyển đổi không được phá lần mở app: thẻ vẫn còn ở
        # chỗ cũ (CM/file), lần sau thử lại; lỗi ghi vào báo cáo
        # để giao diện hiện trong log thay vì chết lặng.
        report["error"] = _short_error_text(error)
    _LAST_MIGRATION = report


def _ensure_cards_migrated_inner(report):
    """Chuyển thẻ từ CM sang file + dọn mồ côi (phần việc thật
    của _ensure_cards_migrated). Lỗi ở đây được hàm ngoài nuốt
    vào report["error"] — không bao giờ phá lần mở app."""
    _migrate_vault_index_to_file()
    entries = _load_vault_index_raw()
    if not entries and _migrate_legacy_card():
        report["legacy"] = True
        report["migrated"] += 1
        entries = _load_vault_index_raw()
    cards, _ = _read_cards_file()
    cards_changed = False
    index_changed = False
    for entry in entries:
        card_id = entry["id"]
        if card_id not in cards:
            cm_card = _read_cm_vault_card(card_id)
            if cm_card:
                cards[card_id] = cm_card
                cards_changed = True
                report["migrated"] += 1
        cm_count = _read_cm_vault_card_success_count(card_id)
        if cm_count > 0 and entry.get("used", 0) != cm_count:
            entry["used"] = cm_count
            index_changed = True
            report["counts_migrated"] += 1
    if cards_changed:
        try:
            _write_cards_file(cards)
        except Exception:
            pass  # thẻ vẫn còn trong CM; lần sau thử lại, không mất
    if index_changed:
        try:
            _save_vault_index(entries)
        except Exception:
            pass
    # Chỉ xoá mục CM của thẻ nào đã nằm chắc trong file (đọc lại
    # file làm chuẩn — ghi file lỗi thì thẻ vẫn còn trong CM).
    cards_on_disk, _ = _read_cards_file()
    for entry in entries:
        if entry["id"] in cards_on_disk:
            _delete_cm_vault_card(entry["id"])
    _migrate_vault_active_to_file()
    cleanup = _cleanup_cm_card_orphans({entry["id"] for entry in entries})
    report["orphans_removed"] = cleanup["removed_services"]
    report["orphan_credentials_removed"] = cleanup["removed_credentials"]
    report["enumerated"] = cleanup["enumerated"]
    report["cleanup_error"] = cleanup["error"]


def _last_migration_report():
    """Báo cáo lần chuyển thẻ từ Credential Manager sang file
    (None khi chưa chạy lần nào trong tiến trình này)."""
    return _LAST_MIGRATION


# --- API kho thẻ (giao diện + coordinator + helper dùng) ------------

def _load_vault_index():
    """Danh mục kho thẻ: list dict {"id", "label", "limit", "used"}
    theo thứ tự dùng (thẻ đầu dùng trước). Không có chữ số thẻ nào
    trong danh mục. Lần đầu đọc thì chuyển thẻ cũ từ Credential
    Manager sang file (xem _ensure_cards_migrated) và hoà mục chờ
    trong vault.pending.json (đợt nhập trước của bản d9800689)."""
    _ensure_cards_migrated()
    return _load_vault_index_raw()


def _load_vault():
    """Toàn bộ kho thẻ cho giao diện: list dict {"id", "label",
    "limit", "used"} theo thứ tự kho. Không có chữ số thẻ nào
    trong kết quả — giao diện chỉ hiện nhãn + giới hạn + số lần
    dùng."""
    entries = _load_vault_index()
    return [
        {
            "id": entry["id"],
            "label": entry["label"],
            "limit": entry["limit"],
            "used": entry.get("used", 0),
        }
        for entry in entries
    ]


def _load_vault_card(card_id):
    """Một thẻ trong kho theo id ({} nếu chưa có). Đọc từ file
    cards.json; chưa chuyển sang file thì đọc dự phòng từ
    Credential Manager (định dạng cũ) để helper chạy độc lập với
    bản cũ vẫn thấy thẻ."""
    cards, _ = _read_cards_file()
    card = cards.get(card_id)
    if card:
        return dict(card)
    return _read_cm_vault_card(card_id)


def _load_vault_card_success_count(card_id):
    """Số tài khoản thành công đã dùng thẻ này (0 nếu chưa có).
    Bộ đếm nằm trong danh mục kho (vault.json), khoá theo id thẻ
    — không bao giờ theo số thẻ."""
    for entry in _load_vault_index():
        if entry["id"] == card_id:
            return entry.get("used", 0)
    return 0


def _save_vault_card_success_count(card_id, count):
    """Ghi lại số tài khoản thành công của một thẻ vào danh mục
    kho. Lỗi ghi (ổ đĩa đầy...) được nuốt — bộ đếm lệch một chút
    còn hơn làm chết lượt chạy."""
    try:
        entries = _load_vault_index()
        for entry in entries:
            if entry["id"] == card_id:
                try:
                    limit = int(entry.get("limit") or 0)
                    used = int(entry.get("used") or 0)
                except (TypeError, ValueError):
                    limit, used = 0, 0
                new_used = max(0, int(count))
                # Thẻ đã full (luồng khác vừa đánh dấu) thì không
                # ghi đè bằng số lần dùng cũ hơn.
                if limit > 0 and used >= limit and new_used < used:
                    return
                entry["used"] = new_used
                _save_vault_index(entries)
                return
    except Exception:
        pass


def _mark_vault_card_full(card_id):
    """Thẻ bị Meta từ chối vì gắn quá nhiều tài khoản quảng cáo.

    Đặt số lần dùng bằng giới hạn để mọi lượt sau bỏ qua thẻ
    này. Giới hạn 0 (không giới hạn) được đặt thành 1 và dùng
    hết 1 — danh mục kho chỉ hiểu "full" là used >= limit > 0.
    Trả về True khi ghi được. Raise khi không ghi được file.
    """
    entries = _load_vault_index()
    for entry in entries:
        if entry["id"] != card_id:
            continue
        try:
            limit = int(entry.get("limit") or 0)
        except (TypeError, ValueError):
            limit = 0
        try:
            used = int(entry.get("used") or 0)
        except (TypeError, ValueError):
            used = 0
        if limit <= 0:
            limit = max(used, 1)
        limit = max(1, min(limit, 100))
        entry["limit"] = limit
        entry["used"] = max(used, limit)
        _save_vault_index(entries)
        return True
    return False


def _delete_vault_card(card_id):
    """Xoá một thẻ khỏi kho: khỏi file thẻ + khỏi danh mục kho,
    và lựa chọn tay nếu đang chọn đúng thẻ này. Mục Credential
    Manager cũ của thẻ (nếu còn) cũng được dọn best-effort.
    Raise RuntimeError khi không ghi được file thẻ."""
    cards, _ = _read_cards_file()
    if card_id in cards:
        del cards[card_id]
        _write_cards_file(cards)
    _delete_cm_vault_card(card_id)
    entries = _load_vault_index()
    kept = [entry for entry in entries if entry["id"] != card_id]
    if len(kept) != len(entries):
        _save_vault_index(kept)
    if _load_vault_active() == card_id:
        _save_vault_active("")


def _set_vault_card_limit(card_id, limit):
    """Đặt giới hạn / thẻ cho một thẻ trong kho (0 = không giới
    hạn, kẹp 0–100). Trả về True khi thẻ có trong kho."""
    limit = max(0, min(int(limit), 100))
    entries = _load_vault_index()
    for entry in entries:
        if entry["id"] == card_id:
            entry["limit"] = limit
            _save_vault_index(entries)
            return True
    return False


def _set_all_vault_card_limits(limit):
    """Đặt MỘT giới hạn chung cho TẤT CẢ thẻ trong kho (0 = không
    giới hạn, kẹp 0–100). Ghi đè giới hạn từng thẻ — thẻ nào muốn
    khác thì đặt riêng lại sau bằng _set_vault_card_limit.
    Trả về số thẻ đã đặt (0 khi kho trống)."""
    limit = max(0, min(int(limit), 100))
    entries = _load_vault_index()
    for entry in entries:
        entry["limit"] = limit
    if entries:
        _save_vault_index(entries)
    return len(entries)

def _parse_bulk_card_line(line, taken_labels, taken_numbers):
    """Phân tích MỘT dòng nhập thẻ hàng loạt.

    Định dạng dòng (phân cách bằng |):
        số_thẻ|MM/YY|CVV|ZIP|nhãn
    - Bắt buộc: số thẻ (12–19 chữ số, khoảng trắng trong số được
      bỏ qua), hạn MM/YY (tháng 01–12), CVV (3–4 chữ số).
    - Không bắt buộc: ZIP (mặc định 10001), nhãn (mặc định "Thẻ N").
    Trả về (card_dict, None) khi hợp lệ, (None, lý_do) khi không —
    lý_do chỉ nêu loại lỗi, KHÔNG bao giờ chứa chữ số thẻ hay nội
    dung dòng. taken_labels / taken_numbers là tập nhãn và số thẻ
    đã có (trong kho + trong đợt nhập này) để chặn trùng.
    """
    parts = [part.strip() for part in line.split("|")]
    if len(parts) < 3:
        return None, "thiếu trường (cần ít nhất số thẻ|MM/YY|CVV)"
    if len(parts) > 5:
        return None, "quá nhiều trường (tối đa 5: số thẻ|MM/YY|CVV|ZIP|nhãn)"
    number = "".join(ch for ch in parts[0] if ch.isdigit())
    expiry_digits = "".join(ch for ch in parts[1] if ch.isdigit())
    cvv = "".join(ch for ch in parts[2] if ch.isdigit())
    zip_code = (
        "".join(ch for ch in parts[3] if ch.isdigit())
        if len(parts) > 3 and parts[3]
        else "10001"
    )
    if not 12 <= len(number) <= 19:
        return None, "số thẻ không hợp lệ (cần 12–19 chữ số)"
    if (
        len(expiry_digits) != 4
        or not 1 <= int(expiry_digits[:2]) <= 12
    ):
        return None, "hạn không hợp lệ (dạng MM/YY, tháng 01–12)"
    if len(cvv) not in (3, 4):
        return None, "CVV không hợp lệ (cần 3–4 chữ số)"
    if len(zip_code) != 5:
        return None, "ZIP không hợp lệ (cần 5 chữ số)"
    if number in taken_numbers:
        return None, "số thẻ đã có trong kho"
    label = parts[4].strip() if len(parts) > 4 and parts[4] else ""
    if label and label in taken_labels:
        return None, "nhãn trùng với thẻ đã có"
    card = {
        "number": number,
        "expiry": f"{expiry_digits[:2]}/{expiry_digits[2:]}",
        "cvv": cvv,
        "zip": zip_code,
        "label": label,
    }
    return card, None

def _short_error_text(error):
    """Một dòng lỗi ngắn gọn từ exception, đã lọc mọi dãy chữ
    số dài (>= 6 chữ số) để không bao giờ lộ chữ số thẻ nếu
    thông điệp lỗi vô tình chứa dữ liệu nhạy cảm. Mã lỗi ngắn
    (ví dụ 1783, 1312) vẫn giữ nguyên."""
    text = str(error).strip().splitlines()
    first = text[0] if text else type(error).__name__
    detail = f"{type(error).__name__}: {first}"
    return re.sub(r"\d{6,}", "…", detail)[:200]


def _import_bulk_cards(text):
    """Nhập nhiều thẻ từ văn bản nhiều dòng vào kho (mỗi dòng một
    thẻ, xem _parse_bulk_card_line để biết định dạng).

    Thẻ hợp lệ được ghi vào file cards.json (DPAPI nếu có) và thêm
    vào danh mục kho với giới hạn 0 (không giới hạn) — không còn
    qua Credential Manager nên không còn lỗi 1783 / lỗi 8 khi nhập
    hàng loạt. Dòng lỗi được bỏ qua — KHÔNG dừng cả đợt nhập.
    Trả về (added_labels, errors, skipped): added_labels là nhãn
    các thẻ đã thêm (theo thứ tự dòng), errors là list (số_dòng,
    lý_do) — số dòng tính theo văn bản đã dán (kể cả dòng trống),
    lý do không chứa chữ số thẻ hay nội dung dòng; skipped là số
    dòng tiêu đề/phân cách đã bỏ qua (dòng không có chữ số nào,
    không thể là thẻ). Dòng trống được bỏ qua lặng lẽ, không tính
    là lỗi. Ghi file thẻ lỗi thì không thẻ nào được thêm; lưu danh
    mục lỗi thì file thẻ được hoàn lại như cũ (không mất thẻ) và
    raise lỗi rõ ràng bằng tiếng Việt.
    """
    entries = _load_vault_index()
    cards, _ = _read_cards_file()
    taken_labels = {entry["label"] for entry in entries}
    taken_numbers = set()
    for entry in entries:
        existing = cards.get(entry["id"])
        if existing and existing.get("number"):
            taken_numbers.add(existing["number"])
    added_labels = []
    errors = []
    skipped = 0
    next_auto = len(entries) + 1
    new_entries = list(entries)
    new_cards = dict(cards)
    for line_no, raw_line in enumerate((text or "").splitlines(), start=1):
        line = raw_line.strip()
        if not line:
            continue
        if not any(ch.isdigit() for ch in line):
            # Dòng tiêu đề/phân cách (không có chữ số nào) — không
            # thể là thẻ, bỏ qua và báo số lượng, không tính lỗi.
            skipped += 1
            continue
        card, reason = _parse_bulk_card_line(
            line, taken_labels, taken_numbers
        )
        if reason is not None:
            errors.append((line_no, reason))
            continue
        label = card.pop("label")
        if not label:
            while f"Thẻ {next_auto}" in taken_labels:
                next_auto += 1
            label = f"Thẻ {next_auto}"
            next_auto += 1
        card_id = _new_card_id([entry["id"] for entry in new_entries])
        new_cards[card_id] = card
        new_entries.append({"id": card_id, "label": label, "limit": 0, "used": 0})
        taken_labels.add(label)
        taken_numbers.add(card["number"])
        added_labels.append(label)
    if added_labels:
        try:
            _write_cards_file(new_cards)
        except Exception as error:
            raise RuntimeError(
                f"Đã phân tích được {len(added_labels)} thẻ nhưng không ghi "
                f"được file thẻ {VAULT_CARDS_FILE}: {_short_error_text(error)}. "
                "Không thẻ nào được thêm — kiểm tra lại ổ đĩa / quyền ghi."
            ) from error
        try:
            _save_vault_index(new_entries)
        except Exception as error:
            try:
                _write_cards_file(cards)
            except Exception:
                pass
            raise RuntimeError(
                f"Đã ghi {len(added_labels)} thẻ vào file {VAULT_CARDS_FILE} "
                "nhưng không lưu được danh mục kho thẻ: "
                f"{_short_error_text(error)}. Đã hoàn lại file thẻ như cũ — "
                "không thẻ nào bị mất, hãy thử nhập lại."
            ) from error
    return added_labels, errors, skipped

def _write_redeem_code(code, suffix=""):
    """Ghi mã mời cần quy đổi cho tài khoản sắp chạy vào redeem.txt.

    suffix là hậu tố riêng của luồng song song ("", "-2", ...):
    mỗi luồng có file redeem riêng để các chuỗi mã mời không
    giẫm lên nhau.
    """
    redeem_file = paths.data_file(f"redeem{suffix}.txt")
    if code:
        redeem_file.write_text(code + "\n", encoding="utf-8")
    elif redeem_file.exists():
        redeem_file.unlink()

def _write_mailbox(mailbox, suffix=""):
    """Ghi hộp thư của tài khoản sắp chạy vào mailbox.txt.

    browser_helper đọc file này để biết đọc thư ở dịch vụ nào
    và bằng bí mật nào (mật khẩu mail.tm / token tempmail.lol
    / sid_token Guerrillamail). suffix là hậu tố riêng của
    luồng song song ("", "-2", ...) như proxy/redeem. File
    chứa bí mật hộp thư nên nằm trong .gitignore.
    """
    paths.data_file(f"mailbox{suffix}.txt").write_text(
        json.dumps(mailbox, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )

def _write_active_proxy(line, suffix=""):
    """Ghi proxy đang dùng cho nhóm hiện tại vào proxy.txt (helper đọc).

    suffix là hậu tố riêng của luồng song song ("", "-2", ...).
    """
    paths.data_file(f"proxy{suffix}.txt").write_text(
        line + "\n", encoding="utf-8"
    )
