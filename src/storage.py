"""Lưu trữ: sheet Excel, keyring, file dữ liệu chạy.

Không có tkinter ở đây — module này dùng được từ test lẫn từ
coordinator mà không cần mở giao diện.
"""

import json
import re
import secrets
import time
from datetime import datetime

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

# Kho thẻ (nhiều thẻ): mỗi thẻ một mục riêng trong Windows
# Credential Manager, tên dịch vụ là CARD_SERVICE_PREFIX + id
# thẻ (id ngẫu nhiên, không phải số thẻ). Danh mục kho (id +
# nhãn + giới hạn từng thẻ) và thẻ đang chọn tay nằm ở dịch vụ
# riêng CARD_VAULT_INDEX_SERVICE — trong đó KHÔNG có chữ số
# nào của thẻ; số lần dùng thành công của từng thẻ nằm trong
# chính mục của thẻ đó (trường card_success_count).
CARD_SERVICE_PREFIX = "MuseAIHelper:card:"
CARD_VAULT_INDEX_SERVICE = "MuseAIHelper:cards"
CARD_VAULT_INDEX_FIELD = "index"
CARD_VAULT_ACTIVE_FIELD = "active"

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


def _load_card_from_keyring():
    """Thẻ người dùng đã lưu (nếu có) — chỉ đọc từ Credential Manager."""
    card = {}
    for field in CARD_FIELDS:
        try:
            value = keyring.get_password(CARD_SERVICE, field)
        except Exception:
            value = None
        if value:
            card[field] = value
    return card if card.get("number") else {}

def _save_card_to_keyring(card):
    for field in CARD_FIELDS:
        keyring.set_password(CARD_SERVICE, field, card[field])

def _load_card_success_count():
    """Số tài khoản thành công đã dùng thẻ đang lưu (0 nếu chưa có).

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

def _save_card_success_count(count):
    try:
        keyring.set_password(
            CARD_SERVICE, CARD_SUCCESS_COUNT_FIELD, str(max(0, int(count)))
        )
    except Exception:
        pass

def _delete_card_from_keyring():
    for field in CARD_FIELDS:
        try:
            keyring.delete_password(CARD_SERVICE, field)
        except Exception:
            pass

def _vault_service(card_id):
    """Tên dịch vụ Credential Manager của một thẻ trong kho."""
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

def _load_vault_index():
    """Danh mục kho thẻ: list dict {"id", "label", "limit"} theo
    thứ tự dùng (thẻ đầu dùng trước). Không có chữ số thẻ nào
    trong danh mục. Lần đầu đọc mà kho trống thì thử nhập thẻ
    đơn cũ (xem _migrate_legacy_card)."""
    _migrate_legacy_card()
    try:
        raw = keyring.get_password(
            CARD_VAULT_INDEX_SERVICE, CARD_VAULT_INDEX_FIELD
        )
    except Exception:
        raw = None
    if not raw:
        return []
    try:
        entries = json.loads(raw)
    except (TypeError, ValueError):
        return []
    cleaned = []
    if isinstance(entries, list):
        for entry in entries:
            if not isinstance(entry, dict) or not entry.get("id"):
                continue
            try:
                limit = max(0, min(int(entry.get("limit") or 0), 100))
            except (TypeError, ValueError):
                limit = 0
            cleaned.append(
                {
                    "id": str(entry["id"]),
                    "label": str(entry.get("label") or "").strip()
                    or "Thẻ",
                    "limit": limit,
                }
            )
    return cleaned

def _save_vault_index(entries):
    keyring.set_password(
        CARD_VAULT_INDEX_SERVICE,
        CARD_VAULT_INDEX_FIELD,
        json.dumps(
            [
                {
                    "id": str(entry["id"]),
                    "label": str(entry.get("label") or "").strip(),
                    "limit": max(0, min(int(entry.get("limit") or 0), 100)),
                }
                for entry in entries
            ],
            ensure_ascii=False,
        ),
    )

def _load_vault_active():
    """Id thẻ đang chọn tay ("" là tự động xoay theo thứ tự kho)."""
    try:
        value = keyring.get_password(
            CARD_VAULT_INDEX_SERVICE, CARD_VAULT_ACTIVE_FIELD
        )
    except Exception:
        value = None
    return (value or "").strip()

def _save_vault_active(card_id):
    keyring.set_password(
        CARD_VAULT_INDEX_SERVICE, CARD_VAULT_ACTIVE_FIELD, card_id or ""
    )

def _load_vault_card(card_id):
    """Thẻ trong kho theo id ({} nếu chưa có) — chỉ đọc từ
    Credential Manager, cùng hình dạng với thẻ đơn cũ."""
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

def _save_vault_card(card_id, card):
    service = _vault_service(card_id)
    for field in CARD_FIELDS:
        keyring.set_password(service, field, card[field])

def _delete_vault_card(card_id):
    """Xoá một thẻ khỏi kho: mục Credential Manager của thẻ (cả
    4 trường + bộ đếm), mục của nó trong danh mục kho, và lựa
    chọn tay nếu đang chọn đúng thẻ này."""
    service = _vault_service(card_id)
    for field in CARD_FIELDS + (CARD_SUCCESS_COUNT_FIELD,):
        try:
            keyring.delete_password(service, field)
        except Exception:
            pass
    entries = _load_vault_index()
    kept = [entry for entry in entries if entry["id"] != card_id]
    if len(kept) != len(entries):
        _save_vault_index(kept)
    if _load_vault_active() == card_id:
        _save_vault_active("")

def _load_vault_card_success_count(card_id):
    """Số tài khoản thành công đã dùng thẻ này (0 nếu chưa có).

    Chỉ là một con số đếm, không chứa chữ số nào của thẻ, và
    được khóa theo id thẻ — không bao giờ theo số thẻ.
    """
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

def _save_vault_card_success_count(card_id, count):
    try:
        keyring.set_password(
            _vault_service(card_id),
            CARD_SUCCESS_COUNT_FIELD,
            str(max(0, int(count))),
        )
    except Exception:
        pass

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

def _load_vault():
    """Toàn bộ kho thẻ cho giao diện: list dict {"id", "label",
    "limit", "used"} theo thứ tự kho. Không có chữ số thẻ nào
    trong kết quả — giao diện chỉ hiện nhãn + giới hạn + số lần
    dùng. Thẻ đơn cũ (nếu có) được nhập vào kho ở lần đọc đầu
    tiên (xem _migrate_legacy_card)."""
    entries = _load_vault_index()
    vault = []
    for entry in entries:
        vault.append(
            {
                "id": entry["id"],
                "label": entry["label"],
                "limit": entry["limit"],
                "used": _load_vault_card_success_count(entry["id"]),
            }
        )
    return vault

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

def _import_bulk_cards(text):
    """Nhập nhiều thẻ từ văn bản nhiều dòng vào kho (mỗi dòng một
    thẻ, xem _parse_bulk_card_line để biết định dạng).

    Thẻ hợp lệ được ghi thẳng vào Credential Manager (mỗi thẻ một
    mục riêng) và thêm vào danh mục kho với giới hạn 0 (không giới
    hạn). Dòng lỗi được bỏ qua — KHÔNG dừng cả đợt nhập.
    Trả về (added_labels, errors): added_labels là nhãn các thẻ
    đã thêm (theo thứ tự dòng), errors là list (số_dòng, lý_do)
    — số dòng tính theo văn bản đã dán (kể cả dòng trống), lý do
    không chứa chữ số thẻ hay nội dung dòng. Dòng trống được bỏ
    qua lặng lẽ, không tính là lỗi.
    """
    entries = _load_vault_index()
    taken_labels = {entry["label"] for entry in entries}
    taken_numbers = set()
    for entry in entries:
        existing = _load_vault_card(entry["id"])
        if existing.get("number"):
            taken_numbers.add(existing["number"])
    added_labels = []
    errors = []
    next_auto = len(entries) + 1
    for line_no, raw_line in enumerate((text or "").splitlines(), start=1):
        line = raw_line.strip()
        if not line:
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
        card_id = _new_card_id([entry["id"] for entry in entries])
        try:
            _save_vault_card(card_id, card)
            _save_vault_card_success_count(card_id, 0)
        except Exception:
            errors.append((line_no, "không ghi được vào Credential Manager"))
            continue
        entries.append({"id": card_id, "label": label, "limit": 0})
        taken_labels.add(label)
        taken_numbers.add(card["number"])
        added_labels.append(label)
    if added_labels:
        _save_vault_index(entries)
    return added_labels, errors
    entries = _load_vault_index()
    vault = []
    for entry in entries:
        vault.append(
            {
                "id": entry["id"],
                "label": entry["label"],
                "limit": entry["limit"],
                "used": _load_vault_card_success_count(entry["id"]),
            }
        )
    return vault

def _migrate_legacy_card():
    """Nhập thẻ đơn cũ (dịch vụ MuseAIHelper:card) vào kho, một
    lần duy nhất: khi danh mục kho còn trống mà thẻ đơn cũ còn
    đó, chuyển nó thành thẻ đầu tiên của kho (nhãn "Thẻ 1",
    giới hạn 0 = không giới hạn, giữ nguyên bộ đếm thành công),
    rồi xoá thẻ đơn cũ để kho là nơi duy nhất giữ thẻ. Chỉ xoá
    bản cũ sau khi bản mới đã ghi xong. Trả về True nếu đã nhập.
    """
    try:
        raw = keyring.get_password(
            CARD_VAULT_INDEX_SERVICE, CARD_VAULT_INDEX_FIELD
        )
    except Exception:
        raw = None
    if raw:
        return False
    legacy = _load_card_from_keyring()
    if not legacy:
        return False
    legacy_count = _load_card_success_count()
    card_id = _new_card_id(())
    try:
        _save_vault_card(card_id, legacy)
        _save_vault_card_success_count(card_id, legacy_count)
        _save_vault_index(
            [{"id": card_id, "label": "Thẻ 1", "limit": 0}]
        )
    except Exception:
        return False
    _delete_card_from_keyring()
    try:
        keyring.delete_password(CARD_SERVICE, CARD_SUCCESS_COUNT_FIELD)
    except Exception:
        pass
    return True

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
