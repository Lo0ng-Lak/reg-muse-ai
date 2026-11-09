import json
import os
import queue
import re
import secrets
import subprocess
import sys
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, ttk

import keyring
from openpyxl import Workbook, load_workbook

import mail_providers


# Giữ lại cho tương thích (test cũ đọc danh sách này): danh
# sách domain tempmail.plus sống ở mail_providers.
TEMPMAIL_DOMAINS = mail_providers.TEMPMAIL_PLUS_DOMAINS
SPREADSHEET_FILE = (
    Path.home() / "AppData" / "Local" / "MuseAIHelper" / "accounts.xlsx"
)
PENDING_SPREADSHEET_FILE = SPREADSHEET_FILE.with_name("accounts.pending.xlsx")
MAIL_SERVICE = "MuseAIHelper:mail.tm"
CARD_SERVICE = "MuseAIHelper:card"
CARD_FIELDS = ("number", "expiry", "cvv", "zip")
HELPER_TIMEOUT_SECONDS = 900
PAUSE_BETWEEN_ACCOUNTS_SECONDS = 1


def _create_mailbox_account(provider_id=mail_providers.DEFAULT_PROVIDER):
    """Tạo một hộp thư qua dịch vụ temp mail đã chọn.

    Toàn bộ logic theo từng dịch vụ nằm ở mail_providers.
    Trả về dict {"provider", "email", "secret"} — secret là
    mật khẩu hộp thư (mail.tm) hoặc token (guerrillamail);
    tempmail.plus và maildrop.cc để trống vì
    ở đó địa chỉ tự nó là khoá hộp thư.
    """
    return mail_providers.create_mailbox(provider_id)


SPREADSHEET_HEADERS = [
    "Thời gian tạo",
    "Email Muse.ai",
    "Mật khẩu hộp thư",
    "Cookie phiên đăng nhập",
    "Mã mời",
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


def _append_spreadsheet_row(email, mail_password="", cookie="", invite=""):
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


CARD_SUCCESS_COUNT_FIELD = "card_success_count"


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


def _proxy_display(value):
    """Chỉ lấy host:port của một proxy để ghi log (che thông tin đăng nhập)."""
    display = value
    if "://" in display:
        display = display.split("://", 1)[1]
    if "@" in display:
        display = display.rsplit("@", 1)[1]
    parts = display.split(":")
    return ":".join(parts[:2]) if len(parts) >= 2 else display


def _line_burns_proxy(line):
    """Dòng tín hiệu từ browser_helper: proxy đang dùng bị Meta chặn."""
    return line.startswith("PROXY_BURN")


def _line_flags_waitlist(line):
    """Dòng tín hiệu từ browser_helper: IP thoát bị Join the waitlist."""
    return line.startswith("WAITLIST:")


BAD_PROXIES_FILE = Path(__file__).with_name("bad-proxies.txt")


def _proxy_key(value):
    """Khoá định danh một proxy = nguyên dòng proxy (gồm cả
    user/password), chuẩn hoá khoảng trắng và chữ hoa/thường.

    Các dòng proxy có thể dùng chung một cổng host:port (kiểu
    cổng xoay vòng) mà mỗi user/password là một phiên/IP thoát
    riêng — định danh theo host:port sẽ gộp chúng thành một và
    phá kho. Chỉ gộp khi hai dòng giống hệt nhau.
    """
    return value.strip().lower()


def _prepare_pool(pool, bad_keys):
    """Lọc kho proxy cho một đợt chạy: bỏ các dòng nằm trong kho
    lỗi (khớp theo nguyên dòng qua _proxy_key) và bỏ dòng trùng
    hệt nhau (giữ dòng đầu tiên)."""
    seen = set()
    result = []
    for line in pool:
        key = _proxy_key(line)
        if key in bad_keys or key in seen:
            continue
        seen.add(key)
        result.append(line)
    return result


def _load_bad_proxies():
    """Đọc kho proxy lỗi từ bad-proxies.txt → dict khoá → dòng gốc."""
    result = {}
    try:
        text = BAD_PROXIES_FILE.read_text(encoding="utf-8")
    except OSError:
        return result
    for line in text.splitlines():
        line = line.strip()
        if line:
            result[_proxy_key(line)] = line
    return result


def _save_bad_proxies(mapping):
    try:
        if mapping:
            BAD_PROXIES_FILE.write_text(
                "\n".join(mapping.values()) + "\n", encoding="utf-8"
            )
        elif BAD_PROXIES_FILE.exists():
            BAD_PROXIES_FILE.unlink()
    except OSError:
        pass


def _session_cookie_from_line(line):
    """Tách cookie phiên từ dòng tín hiệu của helper (None nếu không phải).

    Cookie là thông tin đăng nhập — dòng này bị chặn lại, không
    cho hiện vào log của ứng dụng.
    """
    prefix = "SESSION_COOKIE:"
    if line.startswith(prefix):
        return line[len(prefix):].strip()
    return None


def _invite_from_line(line):
    """Tách mã mời của tài khoản từ dòng tín hiệu (None nếu không phải)."""
    prefix = "INVITE_CODE:"
    if line.startswith(prefix):
        return line[len(prefix):].strip()
    return None


def _write_redeem_code(code, suffix=""):
    """Ghi mã mời cần quy đổi cho tài khoản sắp chạy vào redeem.txt.

    suffix là hậu tố riêng của luồng song song ("", "-2", ...):
    mỗi luồng có file redeem riêng để các chuỗi mã mời không
    giẫm lên nhau.
    """
    redeem_file = Path(__file__).with_name(f"redeem{suffix}.txt")
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
    Path(__file__).with_name(f"mailbox{suffix}.txt").write_text(
        json.dumps(mailbox, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )


class _ProxyRotation:
    """Xoay proxy theo hạn mức tài khoản; proxy bị chặn thì bỏ ngay.

    current() cho proxy của tài khoản sắp chạy. Sau mỗi tài khoản
    gọi advance_after_account(): đủ hạn mức — hoặc tài khoản đó
    báo proxy bị Meta chặn tạo tài khoản — thì tài khoản sau dùng
    proxy kế tiếp; hết kho thì quay vòng lại từ đầu. Hạn mức tính
    theo lượt thử, bất kể thành công hay thất bại.
    """

    def __init__(self, pool, per_proxy):
        self.pool = pool
        self.per_proxy = per_proxy
        self.cursor = 0
        self.used = 0
        self.cycles = 0
        self.retired = set()

    def current(self):
        return self.cursor, self.pool[self.cursor]

    def _step(self):
        if self.cursor + 1 >= len(self.pool):
            # Vừa dùng xong proxy cuối: đã đi hết một vòng kho.
            self.cycles += 1
        self.cursor = (self.cursor + 1) % len(self.pool)

    def _skip_retired(self):
        guard = 0
        while self.cursor in self.retired and guard < len(self.pool):
            self._step()
            guard += 1

    def advance_after_account(self, burned=False):
        self.used += 1
        if burned or self.used >= self.per_proxy:
            self._step()
            self.used = 0
            self._skip_retired()

    def retire_current(self):
        """Loại hẳn proxy hiện tại: bỏ phần hạn mức còn lại và
        không bao giờ quay lại proxy này, kể cả khi xoay vòng."""
        self.retired.add(self.cursor)
        self._step()
        self.used = 0
        self._skip_retired()

    def all_retired(self):
        return len(self.retired) >= len(self.pool)

    def pool_exhausted(self):
        """Đã đi hết một vòng kho proxy (mỗi proxy hết lượt một lần),
        hoặc mọi proxy đều đã bị loại vào kho lỗi."""
        return self.cycles >= 1 or self.all_retired()


def _write_active_proxy(line, suffix=""):
    """Ghi proxy đang dùng cho nhóm hiện tại vào proxy.txt (helper đọc).

    suffix là hậu tố riêng của luồng song song ("", "-2", ...).
    """
    Path(__file__).with_name(f"proxy{suffix}.txt").write_text(
        line + "\n", encoding="utf-8"
    )


class MuseAIHelper(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Muse.ai Account Helper")
        self.geometry("780x560")
        self.minsize(640, 440)
        self.events = queue.Queue()
        self.running = False
        self.stop_event = threading.Event()
        self.helper_process = None
        self.helper_processes = set()
        self._process_lock = threading.Lock()
        self._bad_lock = threading.Lock()
        self._bad_proxies = _load_bad_proxies()
        self._lane_progress = {}
        self._lane_local = threading.local()

        self._build_ui()
        self.after(100, self._process_events)
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self._append_log(
            "Sẵn sàng. Nhập số tài khoản rồi bấm 'Bắt đầu đăng ký'. "
            "Chỉ tài khoản đăng ký thành công mới được lưu vào sheet Excel."
        )
        if _load_card_from_keyring():
            self._append_log(
                "Đã có thẻ lưu sẵn — tool sẽ tự điền ở bước xác minh thẻ. "
                f"Thẻ này đã dùng cho {_load_card_success_count()} tài khoản "
                "thành công (xem / đổi ở ô 'Giới hạn / thẻ:')."
            )

    def _build_ui(self):
        root = ttk.Frame(self, padding=16)
        root.pack(fill="both", expand=True)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(7, weight=1)

        ttk.Label(
            root,
            text="Đăng ký Muse.ai theo lô",
            font=("Segoe UI", 16, "bold"),
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            root,
            text=(
                "Email tạm lấy từ dịch vụ ở ô Temp mail bên dưới, không phải Gmail, và Muse.ai có thể "
                "từ chối email tạm. Mỗi tài khoản: tạo mail → mở form đăng ký → "
                "tự lấy mã xác minh → điền hồ sơ → thành công mới lưu lại, rồi "
                "tự chuyển sang tài khoản tiếp theo."
            ),
            wraplength=720,
        ).grid(row=1, column=0, sticky="ew", pady=(6, 12))

        proxy_row = ttk.Frame(root)
        proxy_row.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        proxy_row.columnconfigure(1, weight=1)
        ttk.Label(proxy_row, text="Proxy:").grid(row=0, column=0, sticky="nw")
        self.proxy_text = tk.Text(proxy_row, height=3, width=52)
        self.proxy_text.grid(row=0, column=1, sticky="ew", padx=(6, 6))
        ttk.Label(
            proxy_row,
            text=(
                "mỗi dòng 1 proxy (host:port:user:pass) — kho proxy "
                "US của bạn, tool không kiểm tra lại IP thoát"
            ),
            wraplength=190,
        ).grid(row=0, column=2, sticky="w")
        saved_pool = ""
        for pool_name in ("proxies.txt", "proxy.txt"):
            try:
                saved_pool = (
                    Path(__file__)
                    .with_name(pool_name)
                    .read_text(encoding="utf-8")
                    .strip()
                )
            except OSError:
                saved_pool = ""
            if saved_pool:
                break
        if saved_pool:
            self.proxy_text.insert("1.0", saved_pool)

        bad_row = ttk.Frame(root)
        bad_row.grid(row=3, column=0, sticky="ew", pady=(0, 10))
        bad_row.columnconfigure(1, weight=1)
        ttk.Label(bad_row, text="Proxy lỗi:").grid(row=0, column=0, sticky="nw")
        self.bad_proxy_text = tk.Text(bad_row, height=3, width=52)
        self.bad_proxy_text.grid(row=0, column=1, sticky="ew", padx=(6, 6))
        ttk.Label(
            bad_row,
            text=(
                "proxy bị 'Join the waitlist' hoặc bị Meta chặn sẽ "
                "tự vào đây và không được dùng lại ở mọi đợt chạy "
                "sau — xoá một dòng là cho dùng lại proxy đó"
            ),
            wraplength=190,
        ).grid(row=0, column=2, sticky="w")
        if self._bad_proxies:
            self.bad_proxy_text.insert(
                "1.0", "\n".join(self._bad_proxies.values())
            )

        card_row = ttk.Frame(root)
        card_row.grid(row=4, column=0, sticky="ew", pady=(0, 10))
        ttk.Label(card_row, text="Thẻ:").pack(side="left")
        self.card_number_entry = ttk.Entry(card_row, width=24)
        self.card_number_entry.pack(side="left", padx=(6, 8))
        ttk.Label(card_row, text="Hạn:").pack(side="left")
        self.card_expiry_entry = ttk.Entry(card_row, width=7)
        self.card_expiry_entry.pack(side="left", padx=(4, 8))
        ttk.Label(card_row, text="CVV:").pack(side="left")
        self.card_cvv_entry = ttk.Entry(card_row, width=5)
        self.card_cvv_entry.pack(side="left", padx=(4, 8))
        ttk.Label(card_row, text="ZIP:").pack(side="left")
        self.card_zip_entry = ttk.Entry(card_row, width=7)
        self.card_zip_entry.pack(side="left", padx=(4, 10))
        self.save_card_button = ttk.Button(
            card_row, text="Lưu thẻ", command=self._save_card
        )
        self.save_card_button.pack(side="left")
        self.delete_card_button = ttk.Button(
            card_row, text="Xoá thẻ", command=self._delete_card
        )
        self.delete_card_button.pack(side="left", padx=(6, 0))
        saved_card = _load_card_from_keyring()
        if saved_card:
            self.card_number_entry.insert(0, saved_card.get("number", ""))
            self.card_expiry_entry.insert(0, saved_card.get("expiry", ""))
            self.card_cvv_entry.insert(0, saved_card.get("cvv", ""))
            self.card_zip_entry.insert(0, saved_card.get("zip", ""))
        else:
            self.card_zip_entry.insert(0, "10001")

        controls = ttk.Frame(root)
        controls.grid(row=5, column=0, sticky="ew")
        ttk.Label(controls, text="Số tài khoản:").pack(side="left")
        self.count_var = tk.IntVar(value=5)
        self.count_spin = ttk.Spinbox(
            controls, from_=0, to=100, width=5, textvariable=self.count_var
        )
        self.count_spin.pack(side="left", padx=(6, 12))
        ttk.Label(controls, text="Tài khoản / proxy:").pack(side="left")
        self.per_proxy_var = tk.IntVar(value=5)
        self.per_proxy_spin = ttk.Spinbox(
            controls, from_=1, to=100, width=5, textvariable=self.per_proxy_var
        )
        self.per_proxy_spin.pack(side="left", padx=(6, 12))
        ttk.Label(controls, text="Giới hạn / thẻ:").pack(side="left")
        self.card_limit_var = tk.IntVar(value=0)
        self.card_limit_spin = ttk.Spinbox(
            controls, from_=0, to=100, width=5, textvariable=self.card_limit_var
        )
        self.card_limit_spin.pack(side="left", padx=(6, 12))
        ttk.Label(controls, text="Số luồng:").pack(side="left")
        self.lanes_var = tk.IntVar(value=1)
        self.lanes_spin = ttk.Spinbox(
            controls, from_=1, to=5, width=5, textvariable=self.lanes_var
        )
        self.lanes_spin.pack(side="left", padx=(6, 12))
        ttk.Label(controls, text="Mã mời:").pack(side="left")
        self.invite_entry = ttk.Entry(controls, width=10)
        self.invite_entry.pack(side="left", padx=(6, 12))
        try:
            saved_invite = (
                Path(__file__)
                .with_name("invite.txt")
                .read_text(encoding="utf-8")
                .strip()
            )
        except OSError:
            saved_invite = ""
        if saved_invite:
            self.invite_entry.insert(0, saved_invite.splitlines()[0].strip())
        self.start_button = ttk.Button(
            controls, text="Bắt đầu đăng ký", command=self._start_registration
        )
        self.start_button.pack(side="left")
        self.stop_button = ttk.Button(
            controls,
            text="Dừng",
            command=self._stop_registration,
            state="disabled",
        )
        self.stop_button.pack(side="left", padx=(8, 0))
        self.open_sheet_button = ttk.Button(
            controls,
            text="Mở sheet Excel",
            command=self._open_spreadsheet,
            state="normal" if SPREADSHEET_FILE.exists() else "disabled",
        )
        self.open_sheet_button.pack(side="right")

        mail_row = ttk.Frame(root)
        mail_row.grid(row=6, column=0, sticky="ew", pady=(10, 0))
        ttk.Label(mail_row, text="Temp mail:").pack(side="left")
        self.provider_var = tk.StringVar(
            value=dict(mail_providers.PROVIDERS)[
                mail_providers.DEFAULT_PROVIDER
            ]
        )
        self.provider_combo = ttk.Combobox(
            mail_row,
            textvariable=self.provider_var,
            values=[label for _, label in mail_providers.PROVIDERS],
            state="readonly",
            width=16,
        )
        self.provider_combo.pack(side="left", padx=(6, 8))
        ttk.Label(
            mail_row,
            text=(
                "dịch vụ hộp thư tạm cho lượt chạy — tempmail.plus "
                "là mặc định và là dịch vụ duy nhất đã kiểm chứng "
                "nhận OTP từ Meta tới nay"
            ),
            wraplength=480,
        ).pack(side="left")

        log_frame = ttk.LabelFrame(root, text="Quá trình", padding=8)
        log_frame.grid(row=7, column=0, sticky="nsew", pady=(12, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log_text = tk.Text(log_frame, wrap="word", state="disabled")
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(
            log_frame, orient="vertical", command=self.log_text.yview
        )
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)

        self.progress = tk.StringVar(value="Sẵn sàng.")
        ttk.Label(root, textvariable=self.progress).grid(
            row=8, column=0, sticky="ew", pady=(10, 0)
        )

    def _save_card(self):
        """Lưu thẻ vào Windows Credential Manager (không ghi ra file).

        Chỉ giữ trong khoá bảo mật của Windows; không chữ số
        nào của thẻ được in ra log.
        """
        def digits(text):
            return "".join(ch for ch in text if ch.isdigit())

        number = digits(self.card_number_entry.get())
        expiry_digits = digits(self.card_expiry_entry.get())
        cvv = digits(self.card_cvv_entry.get())
        zip_code = digits(self.card_zip_entry.get()) or "10001"
        valid = (
            12 <= len(number) <= 19
            and len(expiry_digits) == 4
            and 1 <= int(expiry_digits[:2]) <= 12
            and len(cvv) in (3, 4)
            and len(zip_code) == 5
        )
        if not valid:
            messagebox.showwarning(
                "Thẻ chưa hợp lệ",
                "Kiểm tra lại: số thẻ 12–19 chữ số, hạn dạng MM/YY, "
                "CVV 3–4 chữ số, ZIP 5 chữ số.",
            )
            return
        card = {
            "number": number,
            "expiry": f"{expiry_digits[:2]}/{expiry_digits[2:]}",
            "cvv": cvv,
            "zip": zip_code,
        }
        old_card = _load_card_from_keyring()
        try:
            _save_card_to_keyring(card)
        except Exception as error:
            messagebox.showerror(
                "Không lưu được thẻ",
                f"Không ghi được vào Windows Credential Manager: {error}",
            )
            return
        if old_card.get("number") != number:
            # Đổi sang thẻ khác: giới hạn tính theo từng thẻ nên
            # bộ đếm tài khoản thành công của thẻ cũ reset về 0.
            _save_card_success_count(0)
        self.card_zip_entry.delete(0, "end")
        self.card_zip_entry.insert(0, zip_code)
        self._append_log(
            "Đã lưu thẻ vào Windows Credential Manager — tool sẽ tự "
            "điền ở bước xác minh thẻ, không cần gõ lại."
        )

    def _delete_card(self):
        _delete_card_from_keyring()
        _save_card_success_count(0)
        for entry in (
            self.card_number_entry,
            self.card_expiry_entry,
            self.card_cvv_entry,
            self.card_zip_entry,
        ):
            entry.delete(0, "end")
        self.card_zip_entry.insert(0, "10001")
        self._append_log(
            "Đã xoá thẻ đã lưu — bước thẻ sẽ chờ bạn nhập tay."
        )

    def _sync_proxy_pool(self):
        """Đồng bộ ô Proxy (nhiều dòng) thành kho proxy cho lượt chạy.

        Mỗi dòng là một proxy; dòng trống bị bỏ qua. Kho được lưu
        vào proxies.txt để lần mở sau tự điền lại, còn proxy.txt
        chỉ chứa proxy đang dùng cho nhóm tài khoản hiện tại
        (browser_helper đọc file đó). Ô trống → xoá cả hai file và
        trả về None. Log chỉ dùng _proxy_display — không bao giờ
        lộ user/password của proxy.
        """
        raw = self.proxy_text.get("1.0", "end")
        pool = [line.strip() for line in raw.splitlines() if line.strip()]
        pool_file = Path(__file__).with_name("proxies.txt")
        active_file = Path(__file__).with_name("proxy.txt")
        if not pool:
            for path in (pool_file, active_file):
                try:
                    path.unlink()
                except OSError:
                    pass
            return None
        try:
            pool_file.write_text("\n".join(pool) + "\n", encoding="utf-8")
            active_file.write_text(pool[0] + "\n", encoding="utf-8")
        except OSError as error:
            raise RuntimeError(f"Không ghi được file proxy: {error}") from error
        return pool

    def _sync_bad_proxies(self):
        """Đồng bộ ô Proxy lỗi thành kho loại trừ cho lượt chạy.

        Nội dung ô là nguồn sự thật: dòng nào còn trong ô thì
        proxy đó bị loại; xoá dòng khỏi ô là cho dùng lại. Kho
        được lưu vào bad-proxies.txt để giữ qua các lần mở app.
        Trả về tập khoá host:port bị loại.
        """
        raw = self.bad_proxy_text.get("1.0", "end")
        mapping = {}
        for line in raw.splitlines():
            line = line.strip()
            if line:
                mapping[_proxy_key(line)] = line
        self._bad_proxies = mapping
        _save_bad_proxies(mapping)
        return set(mapping)

    def _mark_proxy_bad(self, line):
        """Đưa một proxy vào kho lỗi (thread-safe, gọi từ các luồng).

        Ghi ngay xuống bad-proxies.txt và báo cho giao diện thêm
        dòng vào ô Proxy lỗi. Trả về True nếu đây là lần đầu
        proxy này bị đánh dấu.
        """
        key = _proxy_key(line)
        with self._bad_lock:
            if key in self._bad_proxies:
                return False
            self._bad_proxies[key] = line
            _save_bad_proxies(self._bad_proxies)
        self.events.put(("bad_proxy", line))
        return True

    def _sync_invite_code(self):
        """Lưu mã mời (nếu có) vào invite.txt để lần sau tự điền lại."""
        code = self.invite_entry.get().strip()
        invite_file = Path(__file__).with_name("invite.txt")
        try:
            if code:
                invite_file.write_text(code + "\n", encoding="utf-8")
            elif invite_file.exists():
                invite_file.unlink()
        except OSError:
            pass
        return code

    def _set_controls_running(self, running):
        self.start_button.configure(state="disabled" if running else "normal")
        self.count_spin.configure(state="disabled" if running else "normal")
        self.proxy_text.configure(state="disabled" if running else "normal")
        self.bad_proxy_text.configure(state="disabled" if running else "normal")
        self.per_proxy_spin.configure(state="disabled" if running else "normal")
        self.card_limit_spin.configure(state="disabled" if running else "normal")
        self.lanes_spin.configure(state="disabled" if running else "normal")
        self.invite_entry.configure(state="disabled" if running else "normal")
        self.provider_combo.configure(
            state="disabled" if running else "readonly"
        )
        for widget in (
            self.card_number_entry,
            self.card_expiry_entry,
            self.card_cvv_entry,
            self.card_zip_entry,
            self.save_card_button,
            self.delete_card_button,
        ):
            widget.configure(state="disabled" if running else "normal")
        self.stop_button.configure(state="normal" if running else "disabled")
        self.open_sheet_button.configure(
            state=(
                "disabled"
                if running
                else ("normal" if SPREADSHEET_FILE.exists() else "disabled")
            )
        )

    def _start_registration(self):
        if self.running:
            return
        try:
            total = int(self.count_var.get())
        except (tk.TclError, ValueError):
            total = 0
        if total < 0 or total > 100:
            messagebox.showwarning(
                "Số tài khoản chưa hợp lệ",
                "Nhập số tài khoản từ 0 đến 100. "
                "Để 0 là chạy liên tục tới khi dùng hết lượt của "
                "toàn bộ proxy trong danh sách thì dừng.",
            )
            return
        try:
            per_proxy = int(self.per_proxy_var.get())
        except (tk.TclError, ValueError):
            per_proxy = 5
        per_proxy = max(1, min(per_proxy, 100))
        try:
            card_limit = int(self.card_limit_var.get())
        except (tk.TclError, ValueError):
            card_limit = 0
        card_limit = max(0, min(card_limit, 100))
        try:
            lanes = int(self.lanes_var.get())
        except (tk.TclError, ValueError):
            lanes = 1
        lanes = max(1, min(lanes, 5))
        # Giới hạn tính theo thẻ ĐANG LƯU (một thẻ duy nhất): chỉ
        # đếm khi có thẻ lưu sẵn; không có thẻ lưu thì bước thẻ là
        # nhập tay nên giới hạn này không áp dụng. Để 0 là không
        # giới hạn.
        saved_card = _load_card_from_keyring()
        card_tracked = bool(saved_card) and card_limit > 0
        card_used = _load_card_success_count() if card_tracked else 0
        if card_tracked and card_used >= card_limit:
            messagebox.showwarning(
                "Thẻ đã đủ giới hạn",
                f"Thẻ đang lưu đã dùng cho {card_used} tài khoản thành "
                f"công (giới hạn {card_limit} / thẻ). Đổi thẻ ở hàng "
                "Thẻ rồi bấm Lưu thẻ để bộ đếm reset về 0, rồi chạy lại.",
            )
            return
        try:
            pool = self._sync_proxy_pool()
        except RuntimeError as error:
            messagebox.showerror("Lỗi proxy", str(error))
            return
        if not pool:
            messagebox.showwarning(
                "Thiếu proxy",
                "Tool chỉ đăng ký qua proxy. "
                "Hãy nhập ít nhất một proxy (mỗi dòng một proxy, "
                "dạng host:port:user:pass) vào ô Proxy rồi bấm lại.",
            )
            return
        bad_keys = self._sync_bad_proxies()
        if bad_keys:
            excluded = sum(
                1 for line in pool if _proxy_key(line) in bad_keys
            )
            if excluded:
                self._append_log(
                    f"Đã loại {excluded} proxy trong ô Proxy lỗi "
                    "khỏi kho cho đợt chạy này."
                )
        pool = _prepare_pool(pool, bad_keys)
        if bad_keys and not pool:
            messagebox.showwarning(
                "Toàn bộ proxy đều nằm trong kho lỗi",
                "Mọi proxy trong ô Proxy đều đã bị đánh dấu "
                "lỗi. Xoá bớt dòng trong ô Proxy lỗi để dùng "
                "lại, hoặc dán proxy mới vào ô Proxy.",
            )
            return
        invite = self._sync_invite_code()
        if not invite:
            seed = _last_invite_from_sheet()
            if seed:
                self.invite_entry.delete(0, "end")
                self.invite_entry.insert(0, seed)
                invite = self._sync_invite_code()
                self._append_log(
                    f"Ô Mã mời trống — tự lấy mã {seed} từ dòng "
                    "cuối trong sheet làm đầu chuỗi."
                )
        provider_label = self.provider_var.get()
        provider_id = next(
            (
                pid
                for pid, label in mail_providers.PROVIDERS
                if label == provider_label
            ),
            mail_providers.DEFAULT_PROVIDER,
        )
        self.running = True
        self._lane_progress = {}
        self.stop_event.clear()
        self._set_controls_running(True)
        if total == 0:
            self._append_log(
                "--- Bắt đầu lượt chạy: KHÔNG giới hạn số tài "
                "khoản — chạy tới khi dùng hết lượt của toàn bộ "
                "proxy trong danh sách (bấm Dừng để ngắt bất cứ "
                "lúc nào) ---"
            )
        else:
            self._append_log(
                f"--- Bắt đầu lượt chạy: mục tiêu {total} tài khoản ---"
            )
        self._append_log(
            f"Kho proxy: {len(pool)} proxy, mỗi proxy {per_proxy} tài "
            "khoản (tính cả thành công lẫn thất bại) — kho do bạn "
            "đảm bảo là proxy US, tool không kiểm tra lại IP thoát."
        )
        self._append_log(
            f"Dịch vụ temp mail cho lượt chạy này: {provider_label}."
        )
        if card_tracked:
            self._append_log(
                f"Giới hạn thẻ đang lưu: {card_used}/{card_limit} tài "
                "khoản thành công — đủ giới hạn là lượt chạy tự dừng, "
                "không tự đổi thẻ."
            )
        threading.Thread(
            target=self._run_loop,
            args=(total, pool, per_proxy, invite, card_limit, card_used, card_tracked, lanes, provider_id),
            daemon=True,
        ).start()

    def _stop_registration(self):
        if not self.running:
            return
        self.stop_event.set()
        self.progress.set("Đang dừng...")

    def _on_close(self):
        if self.running:
            self.stop_event.set()
            with self._process_lock:
                processes = list(self.helper_processes)
            if self.helper_process is not None:
                processes.append(self.helper_process)
            for process in processes:
                try:
                    process.terminate()
                except OSError:
                    pass
        self.destroy()

    def _run_loop(
        self, total, pool, per_proxy, invite,
        card_limit=0, card_used=0, card_tracked=False, lanes=1,
        provider_id=mail_providers.DEFAULT_PROVIDER,
    ):
        # Điều phối các luồng chạy song song. Kho proxy được
        # CHIA theo bước nhảy: luồng k dùng các proxy ở vị trí
        # k, k+lanes, k+2*lanes, ... — không hai luồng nào dùng
        # chung một proxy tại cùng thời điểm, và mỗi luồng xoay
        # vòng trong phần kho của riêng nó (proxy bị chặn cũng
        # chỉ ảnh hưởng luồng đó). Mỗi luồng có CHUỖI MÃ MỜI
        # RIÊNG, nhưng tài khoản ĐẦU TIÊN của mọi luồng đều
        # quy đổi mã trong ô Mã mời (mã chủ tool cài đặt); từ
        # tài khoản thứ hai, mỗi luồng nối tiếp mã do tài khoản
        # trước trong chính luồng đó lấy được. Thẻ
        # thì cả các luồng dùng CHUNG một thẻ đã lưu; bộ đếm
        # giới hạn thẻ tính gộp toàn cục (đếm cả số tài khoản
        # đang chạy dở để không vượt giới hạn).
        lanes = max(1, min(int(lanes), 5, len(pool)))
        if total > 0:
            lanes = min(lanes, total)
        shared = {
            "success": 0,
            "fail": 0,
            "card_used": card_used,
            "in_flight": 0,
            "card_stop_logged": False,
        }
        shared_lock = threading.Lock()
        if lanes > 1:
            self.events.put(
                (
                    "log",
                    f"Chạy {lanes} luồng song song — mỗi luồng "
                    "một phần kho proxy riêng, một chuỗi mã mời "
                    "riêng (tài khoản đầu của mọi luồng đều quy "
                    "đổi mã trong ô Mã mời); dùng chung thẻ đã "
                    "lưu, giới hạn thẻ tính gộp trên tất cả các "
                    "luồng.",
                )
            )
        threads = []
        for lane_no in range(1, lanes + 1):
            subset = pool[lane_no - 1::lanes]
            if total > 0:
                share = total // lanes + (
                    1 if lane_no <= total % lanes else 0
                )
            else:
                share = 0
            seed = invite
            lane_args = (
                lane_no, lanes, share, subset, per_proxy, seed,
                card_limit, card_tracked, shared, shared_lock,
                provider_id,
            )
            if lanes == 1:
                self._run_lane(*lane_args)
            else:
                thread = threading.Thread(
                    target=self._run_lane,
                    args=lane_args,
                    name=f"lane-{lane_no}",
                    daemon=True,
                )
                thread.start()
                threads.append(thread)
        for thread in threads:
            thread.join()
        self.events.put(
            (
                "finished",
                (
                    shared["success"],
                    shared["fail"],
                    self.stop_event.is_set(),
                ),
            )
        )

    def _run_lane(
        self, lane_no, lanes, share, subset, per_proxy, seed,
        card_limit, card_tracked, shared, shared_lock,
        provider_id=mail_providers.DEFAULT_PROVIDER,
    ):
        helper = Path(__file__).with_name("browser_helper.py")
        suffix = "" if lane_no == 1 else f"-{lane_no}"
        self._lane_local.suffix = suffix
        multi = lanes > 1
        lane_tag = f"[Luồng {lane_no}] " if multi else ""
        rotation = _ProxyRotation(subset, per_proxy)
        written_proxy_index = None
        # Chuỗi mã mời CỦA LUỒNG NÀY: tài khoản sắp chạy sẽ quy
        # đổi pending_code (mã của tài khoản trước cùng luồng);
        # mã nó tự lấy được thành pending_code cho tài khoản sau
        # cùng luồng và được ghi đè vào invite.txt để lượt chạy
        # sau có một đầu chuỗi mà nối tiếp.
        pending_code = (seed or "").strip().upper()
        # share = 0 là chế độ không giới hạn: luồng chạy tới khi
        # phần kho proxy của nó dùng hết một vòng thì dừng.
        label_total = share if share > 0 else "∞"
        index = 0
        while not self.stop_event.is_set():
            if share > 0 and index >= share:
                break
            if share == 0 and index > 0 and rotation.pool_exhausted():
                if multi:
                    message = (
                        f"{lane_tag}Đã dùng hết lượt của phần "
                        "proxy của luồng — luồng dừng."
                    )
                else:
                    message = (
                        "Đã dùng hết lượt của toàn bộ proxy trong "
                        "danh sách — dừng lượt chạy."
                    )
                self.events.put(("log", message))
                break
            if rotation.all_retired():
                self.events.put(
                    (
                        "log",
                        f"{lane_tag}Toàn bộ proxy của luồng đã bị "
                        "đưa vào kho proxy lỗi — luồng dừng.",
                    )
                )
                break
            with shared_lock:
                card_hit = (
                    card_tracked
                    and card_limit > 0
                    and shared["card_used"] + shared["in_flight"]
                    >= card_limit
                )
                log_card_stop = (
                    card_hit and not shared["card_stop_logged"]
                )
                if log_card_stop:
                    shared["card_stop_logged"] = True
                if not card_hit:
                    # Giữ chỗ ngay trong cùng một khóa với phép
                    # kiểm tra: hai luồng không thể cùng thấy
                    # "còn chỗ" rồi cùng chạy vượt giới hạn thẻ.
                    # Chỗ được trả lại ở finally cuối vòng lặp.
                    shared["in_flight"] += 1
            if card_hit:
                if log_card_stop:
                    self.events.put(
                        (
                            "log",
                            f"Thẻ đang lưu đã đủ {card_limit} tài "
                            "khoản thành công — dừng lượt chạy. "
                            "Muốn chạy tiếp thì đổi thẻ ở hàng Thẻ "
                            "rồi bấm Lưu thẻ (bộ đếm reset về 0).",
                        )
                    )
                break
            try:
                index += 1
                display = f"L{lane_no} · {index}" if multi else index
                proxy_index, line = rotation.current()
                if proxy_index != written_proxy_index:
                    written_proxy_index = proxy_index
                    try:
                        _write_active_proxy(line, suffix)
                        note = ""
                    except OSError as error:
                        note = f" (không ghi được proxy{suffix}.txt: {error})"
                    self.events.put(
                        (
                            "log",
                            f"{lane_tag}Proxy #{proxy_index + 1}/{len(subset)}: "
                            f"{_proxy_display(line)} — dùng cho các tài "
                            f"khoản tiếp theo.{note}",
                        )
                    )
                with shared_lock:
                    s_now = shared["success"]
                    f_now = shared["fail"]
                    c_now = shared["card_used"]
                card_note = (
                    f" · Thẻ: {c_now}/{card_limit}" if card_tracked else ""
                )
                self.events.put(
                    (
                        "progress",
                        (
                            lane_no,
                            index,
                            label_total,
                            s_now,
                            f_now,
                            card_note,
                            multi,
                        ),
                    )
                )
                try:
                    created = _create_mailbox_account(provider_id)
                except Exception as error:
                    with shared_lock:
                        shared["fail"] += 1
                    self.events.put(
                        ("log", f"[{display}/{label_total}] Tạo hộp thư thất bại: {error}")
                    )
                    rotation.advance_after_account()
                    continue
                if isinstance(created, dict):
                    mailbox = created
                else:
                    # Tương thích ngược (bản vá kiểm thử trả về
                    # chuỗi địa chỉ trần): coi như hộp thư không
                    # có bí mật ở dịch vụ đang chọn.
                    mailbox = {
                        "provider": provider_id,
                        "email": str(created),
                        "secret": "",
                    }
                address = mailbox["email"]
                self.events.put(
                    (
                        "log",
                        f"[{display}/{label_total}] Đã tạo mail {address} — bắt đầu đăng ký.",
                    )
                )
                try:
                    _write_redeem_code(pending_code, suffix)
                except OSError:
                    pass
                try:
                    _write_mailbox(mailbox, suffix)
                except OSError:
                    pass
                helper_ok, burned, waitlisted, cookie, own_invite = self._run_helper(
                    helper, address, display, label_total
                )
                if helper_ok:
                    saved = False
                    save_error = None
                    with shared_lock:
                        try:
                            # Cột "Mật khẩu hộp thư" của sheet:
                            # bí mật của hộp thư theo dịch vụ —
                            # mật khẩu mail.tm, token tempmail.lol
                            # / Guerrillamail; tempmail.plus và
                            # maildrop.cc để trống vì địa chỉ tự
                            # nó là khoá hộp thư.
                            mail_password = mailbox.get("secret") or ""
                            _append_spreadsheet_row(
                                address,
                                mail_password=mail_password,
                                cookie=cookie,
                                invite=own_invite,
                            )
                            saved = True
                            shared["success"] += 1
                            if card_tracked:
                                # Tài khoản chỉ được lưu sau khi qua
                                # bước thẻ, nên đây đúng là 1 lượt
                                # dùng thẻ đang lưu.
                                shared["card_used"] += 1
                                _save_card_success_count(shared["card_used"])
                        except RuntimeError as error:
                            shared["fail"] += 1
                            save_error = error
                    if saved:
                        self.events.put(
                            (
                                "log",
                                f"[{display}/{label_total}] ✔ Đăng ký thành công {address} "
                                "— đã lưu vào sheet Excel.",
                            )
                        )
                    else:
                        self.events.put(
                            (
                                "log",
                                f"[{display}/{label_total}] {address} đã đăng ký xong nhưng "
                                f"không ghi được sheet Excel: {save_error}",
                            )
                        )
                else:
                    with shared_lock:
                        shared["fail"] += 1
                    if self.stop_event.is_set():
                        self.events.put(
                            ("log", f"[{display}/{label_total}] Đã dừng theo yêu cầu.")
                        )
                    else:
                        self.events.put(
                            (
                                "log",
                                f"[{display}/{label_total}] ✘ Đăng ký thất bại cho {address} "
                                "— lý do ở dòng RESULT phía trên.",
                            )
                        )
                if helper_ok and own_invite:
                    pending_code = own_invite
                    try:
                        Path(__file__).with_name("invite.txt").write_text(
                            own_invite + "\n", encoding="utf-8"
                        )
                    except OSError:
                        pass
                    self.events.put(("invite", own_invite))
                    self.events.put(
                        (
                            "log",
                            f"[{display}/{label_total}] Chuỗi mã mời: tài khoản "
                            f"tiếp theo sẽ quy đổi mã {own_invite} — mã "
                            "của tài khoản này.",
                        )
                    )
                if burned:
                    self.events.put(
                        (
                            "log",
                            f"{lane_tag}Proxy #{proxy_index + 1} bị Meta chặn — "
                            "đổi sang proxy kế từ tài khoản sau.",
                        )
                    )
                if burned or waitlisted:
                    # Proxy này hỏng ở cấp IP: Meta chặn tạo tài
                    # khoản, hoặc IP thoát rơi vào khu vực chưa
                    # hỗ trợ (waitlist). Đưa vào kho proxy lỗi để
                    # loại khỏi mọi đợt chạy sau, và loại ngay
                    # trong đợt này (kể cả khi xoay vòng quay lại).
                    reason = (
                        "bị Meta chặn tạo tài khoản"
                        if burned
                        else "bị 'Join the waitlist' (khu vực "
                        "không hỗ trợ)"
                    )
                    if self._mark_proxy_bad(line):
                        self.events.put(
                            (
                                "log",
                                f"[{display}/{label_total}] Proxy "
                                f"{_proxy_display(line)} {reason} — "
                                "đã đưa vào ô Proxy lỗi, không dùng "
                                "lại ở các đợt sau.",
                            )
                        )
                    rotation.retire_current()
                else:
                    rotation.advance_after_account(burned=burned)
                if (share == 0 or index < share) and not self.stop_event.is_set():
                    time.sleep(PAUSE_BETWEEN_ACCOUNTS_SECONDS)
            finally:
                with shared_lock:
                    shared["in_flight"] -= 1


    def _run_helper(self, helper, address, index, total):
        suffix = getattr(self._lane_local, "suffix", "")
        command = [sys.executable, str(helper), address, "--auto"]
        if suffix:
            command.append(suffix.lstrip("-"))
        try:
            process = subprocess.Popen(
                command,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
        except OSError as error:
            self.events.put(
                ("log", f"[{index}/{total}] Không mở được trình duyệt: {error}")
            )
            return False, False, False, "", ""
        self.helper_process = process
        with self._process_lock:
            self.helper_processes.add(process)
        signals = {"burn": False, "waitlist": False, "cookie": "", "invite": ""}

        def forward_output():
            for line in process.stdout:
                line = line.strip()
                if line:
                    if _line_burns_proxy(line):
                        signals["burn"] = True
                    if _line_flags_waitlist(line):
                        signals["waitlist"] = True
                    cookie = _session_cookie_from_line(line)
                    if cookie is not None:
                        # Cookie phiên chỉ đi vào sheet, không vào log.
                        signals["cookie"] = cookie
                        continue
                    invite = _invite_from_line(line)
                    if invite is not None:
                        signals["invite"] = invite
                        continue
                    self.events.put(("log", f"[{index}/{total}] {line}"))

        reader = threading.Thread(target=forward_output, daemon=True)
        reader.start()
        deadline = time.monotonic() + HELPER_TIMEOUT_SECONDS
        while process.poll() is None:
            if self.stop_event.is_set():
                process.terminate()
                break
            if time.monotonic() > deadline:
                self.events.put(
                    (
                        "log",
                        f"[{index}/{total}] Quá "
                        f"{HELPER_TIMEOUT_SECONDS // 60} phút cho một tài "
                        "khoản — đóng trình duyệt và bỏ qua.",
                    )
                )
                process.terminate()
                break
            time.sleep(0.3)
        try:
            process.wait(timeout=10)
        except subprocess.TimeoutExpired:
            process.kill()
        reader.join(timeout=2)
        self.helper_process = None
        with self._process_lock:
            self.helper_processes.discard(process)
        return (
            process.returncode == 0 and not self.stop_event.is_set()
        ), signals["burn"], signals["waitlist"], signals["cookie"], signals["invite"]

    def _process_events(self):
        try:
            while True:
                kind, payload = self.events.get_nowait()
                if kind == "log":
                    self._append_log(payload)
                elif kind == "bad_proxy":
                    existing = {
                        _proxy_key(line)
                        for line in self.bad_proxy_text.get(
                            "1.0", "end"
                        ).splitlines()
                        if line.strip()
                    }
                    if _proxy_key(payload) not in existing:
                        if self.running:
                            self.bad_proxy_text.configure(state="normal")
                        current = self.bad_proxy_text.get(
                            "1.0", "end"
                        ).strip()
                        self.bad_proxy_text.delete("1.0", "end")
                        self.bad_proxy_text.insert(
                            "1.0",
                            (current + "\n" if current else "")
                            + payload,
                        )
                        if self.running:
                            self.bad_proxy_text.configure(state="disabled")
                elif kind == "invite":
                    # Ô Mã mời luôn phản ánh đầu chuỗi hiện tại: mã
                    # của tài khoản thành công gần nhất.
                    if self.running:
                        self.invite_entry.configure(state="normal")
                    self.invite_entry.delete(0, "end")
                    self.invite_entry.insert(0, payload)
                    if self.running:
                        self.invite_entry.configure(state="disabled")
                elif kind == "progress":
                    (
                        lane_no,
                        done,
                        label_total,
                        s_now,
                        f_now,
                        card_note,
                        multi,
                    ) = payload
                    if not multi:
                        self.progress.set(
                            f"Tài khoản {done}/{label_total} · "
                            f"Thành công: {s_now} · Thất bại: {f_now}"
                            f"{card_note}"
                        )
                    else:
                        self._lane_progress[lane_no] = done
                        suffix = "/∞" if label_total == "∞" else ""
                        tally = " · ".join(
                            f"Luồng {k}: {v}{suffix}"
                            for k, v in sorted(
                                self._lane_progress.items()
                            )
                        )
                        self.progress.set(
                            f"{tally} · Tổng thành công: {s_now} · "
                            f"Thất bại: {f_now}{card_note}"
                        )
                elif kind == "finished":
                    success, fail, stopped = payload
                    self.running = False
                    self._set_controls_running(False)
                    note = "Đã dừng theo yêu cầu. " if stopped else ""
                    self.progress.set(
                        f"{note}Hoàn tất: {success} thành công, {fail} thất bại."
                    )
                    self._append_log(
                        f"--- Kết thúc lượt chạy: {success} thành công, "
                        f"{fail} thất bại ---"
                    )
        except queue.Empty:
            pass
        self.after(100, self._process_events)

    def _append_log(self, text):
        stamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"[{stamp}] {text}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _open_spreadsheet(self):
        if not SPREADSHEET_FILE.exists():
            messagebox.showwarning("Chưa có sheet", "Chưa tạo được file Excel.")
            return
        try:
            os.startfile(str(SPREADSHEET_FILE))
        except OSError as error:
            messagebox.showerror("Không mở được Excel", str(error))


if __name__ == "__main__":
    MuseAIHelper().mainloop()
