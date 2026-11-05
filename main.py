import json
import os
import queue
import secrets
import subprocess
import sys
import threading
import time
import urllib.error
import urllib.request
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, ttk

import keyring
from openpyxl import Workbook, load_workbook


API_BASE = "https://api.mail.tm"
SPREADSHEET_FILE = (
    Path.home() / "AppData" / "Local" / "MuseAIHelper" / "accounts.xlsx"
)
MAIL_SERVICE = "MuseAIHelper:mail.tm"
MUSE_SERVICE = "MuseAIHelper:muse.ai"
CARD_SERVICE = "MuseAIHelper:card"
CARD_FIELDS = ("number", "expiry", "cvv", "zip")
HELPER_TIMEOUT_SECONDS = 900
PAUSE_BETWEEN_ACCOUNTS_SECONDS = 1


def _generate_password():
    groups = (
        "abcdefghijklmnopqrstuvwxyz",
        "ABCDEFGHIJKLMNOPQRSTUVWXYZ",
        "0123456789",
        "!@#$%_-",
    )
    characters = [secrets.choice(group) for group in groups]
    alphabet = "".join(groups)
    characters.extend(secrets.choice(alphabet) for _ in range(20 - len(characters)))
    secrets.SystemRandom().shuffle(characters)
    return "".join(characters)


def _request(method, path, payload=None, token=None):
    headers = {"Accept": "application/ld+json"}
    data = None
    if payload is not None:
        headers["Content-Type"] = "application/json"
        data = json.dumps(payload).encode("utf-8")
    if token:
        headers["Authorization"] = f"Bearer {token}"

    request = urllib.request.Request(
        f"{API_BASE}{path}", data=data, headers=headers, method=method
    )
    try:
        with urllib.request.urlopen(request, timeout=20) as response:
            body = response.read()
    except urllib.error.HTTPError as error:
        detail = error.read().decode("utf-8", errors="replace")
        try:
            parsed = json.loads(detail)
            detail = parsed.get("hydra:description") or parsed.get("detail") or detail
        except (json.JSONDecodeError, AttributeError):
            pass
        raise RuntimeError(f"mail.tm trả lỗi HTTP {error.code}: {detail}") from error
    except (urllib.error.URLError, TimeoutError) as error:
        raise RuntimeError(f"Không kết nối được mail.tm: {error}") from error

    if not body:
        return {}
    return json.loads(body.decode("utf-8"))


def _create_mailbox_account():
    """Tạo một hộp thư mail.tm mới và lưu cả hai mật khẩu vào keyring.

    Trả về (address, muse_password). Mật khẩu hộp thư chỉ sống trong
    keyring — browser_helper.py tự lấy lại từ đó khi chạy.
    """
    domains = []
    last_error = None
    for _ in range(3):
        try:
            domains = _request("GET", "/domains?page=1").get(
                "hydra:member", []
            )
            break
        except RuntimeError as error:
            last_error = error
            time.sleep(2)
    else:
        raise last_error
    active_domains = [
        domain["domain"]
        for domain in domains
        if domain.get("isActive") and not domain.get("isPrivate")
    ]
    if not active_domains:
        raise RuntimeError("mail.tm hiện không cung cấp domain khả dụng.")

    last_error = None
    attempts = max(3, len(active_domains))
    for attempt in range(attempts):
        domain = active_domains[attempt % len(active_domains)]
        address = f"muse{secrets.token_hex(5)}@{domain}"
        mailbox_password = secrets.token_urlsafe(24)
        try:
            _request(
                "POST",
                "/accounts",
                {"address": address, "password": mailbox_password},
            )
        except RuntimeError as error:
            last_error = error
            continue
        muse_password = _generate_password()
        try:
            keyring.set_password(MAIL_SERVICE, address, mailbox_password)
            keyring.set_password(MUSE_SERVICE, address, muse_password)
        except Exception as error:
            raise RuntimeError(
                "Hộp thư đã tạo nhưng không lưu được mật khẩu vào "
                f"Windows Credential Manager: {error}"
            ) from error
        return address, muse_password
    raise last_error or RuntimeError("Không tạo được hộp thư mail.tm.")


SPREADSHEET_HEADERS = [
    "Thời gian tạo",
    "Email Muse.ai",
    "Mật khẩu Muse.ai",
    "Mật khẩu mail.tm",
    "Cookie phiên đăng nhập",
    "Mã mời",
]


def _append_spreadsheet_row(
    email, password, mail_password="", cookie="", invite=""
):
    temporary_file = SPREADSHEET_FILE.with_name("accounts.tmp.xlsx")
    try:
        SPREADSHEET_FILE.parent.mkdir(parents=True, exist_ok=True)
        if SPREADSHEET_FILE.exists():
            workbook = load_workbook(SPREADSHEET_FILE)
            sheet = workbook.active
            # Sheet cũ chỉ có 3 cột: bổ sung tiêu đề cho các cột mới
            # ở hàng đầu (chỉ ghi vào ô đang trống, không đè dữ liệu).
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
                password,
                mail_password,
                cookie,
                invite,
            ]
        )
        workbook.save(temporary_file)
        temporary_file.replace(SPREADSHEET_FILE)
    except Exception as error:
        raise RuntimeError(f"Không ghi được file Excel: {error}") from error
    finally:
        if temporary_file.exists():
            temporary_file.unlink()


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


def _write_redeem_code(code):
    """Ghi mã mời cần quy đổi cho tài khoản sắp chạy vào redeem.txt."""
    redeem_file = Path(__file__).with_name("redeem.txt")
    if code:
        redeem_file.write_text(code + "\n", encoding="utf-8")
    elif redeem_file.exists():
        redeem_file.unlink()


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

    def current(self):
        return self.cursor, self.pool[self.cursor]

    def advance_after_account(self, burned=False):
        self.used += 1
        if burned or self.used >= self.per_proxy:
            self.cursor = (self.cursor + 1) % len(self.pool)
            self.used = 0


def _write_active_proxy(line):
    """Ghi proxy đang dùng cho nhóm hiện tại vào proxy.txt (helper đọc)."""
    Path(__file__).with_name("proxy.txt").write_text(
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

        self._build_ui()
        self.after(100, self._process_events)
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self._append_log(
            "Sẵn sàng. Nhập số tài khoản rồi bấm 'Bắt đầu đăng ký'. "
            "Chỉ tài khoản đăng ký thành công mới được lưu vào sheet Excel."
        )
        if _load_card_from_keyring():
            self._append_log(
                "Đã có thẻ lưu sẵn — tool sẽ tự điền ở bước xác minh thẻ."
            )

    def _build_ui(self):
        root = ttk.Frame(self, padding=16)
        root.pack(fill="both", expand=True)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(5, weight=1)

        ttk.Label(
            root,
            text="Đăng ký Muse.ai theo lô",
            font=("Segoe UI", 16, "bold"),
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            root,
            text=(
                "Mail.tm cung cấp email tạm, không phải Gmail, và Muse.ai có thể "
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

        card_row = ttk.Frame(root)
        card_row.grid(row=3, column=0, sticky="ew", pady=(0, 10))
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
        controls.grid(row=4, column=0, sticky="ew")
        ttk.Label(controls, text="Số tài khoản:").pack(side="left")
        self.count_var = tk.IntVar(value=5)
        self.count_spin = ttk.Spinbox(
            controls, from_=1, to=100, width=5, textvariable=self.count_var
        )
        self.count_spin.pack(side="left", padx=(6, 12))
        ttk.Label(controls, text="Tài khoản / proxy:").pack(side="left")
        self.per_proxy_var = tk.IntVar(value=5)
        self.per_proxy_spin = ttk.Spinbox(
            controls, from_=1, to=100, width=5, textvariable=self.per_proxy_var
        )
        self.per_proxy_spin.pack(side="left", padx=(6, 12))
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

        log_frame = ttk.LabelFrame(root, text="Quá trình", padding=8)
        log_frame.grid(row=5, column=0, sticky="nsew", pady=(12, 0))
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
            row=6, column=0, sticky="ew", pady=(10, 0)
        )

    def _save_card(self):
        """Lưu thẻ vào Windows Credential Manager (không ghi ra file).

        Chỉ giữ trong khoá bảo mật của Windows, cùng chỗ với mật
        khẩu mail.tm; không chữ số nào của thẻ được in ra log.
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
        try:
            _save_card_to_keyring(card)
        except Exception as error:
            messagebox.showerror(
                "Không lưu được thẻ",
                f"Không ghi được vào Windows Credential Manager: {error}",
            )
            return
        self.card_zip_entry.delete(0, "end")
        self.card_zip_entry.insert(0, zip_code)
        self._append_log(
            "Đã lưu thẻ vào Windows Credential Manager — tool sẽ tự "
            "điền ở bước xác minh thẻ, không cần gõ lại."
        )

    def _delete_card(self):
        _delete_card_from_keyring()
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
        self.per_proxy_spin.configure(state="disabled" if running else "normal")
        self.invite_entry.configure(state="disabled" if running else "normal")
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
        if total < 1 or total > 100:
            messagebox.showwarning(
                "Số tài khoản chưa hợp lệ",
                "Nhập số tài khoản từ 1 đến 100.",
            )
            return
        try:
            per_proxy = int(self.per_proxy_var.get())
        except (tk.TclError, ValueError):
            per_proxy = 5
        per_proxy = max(1, min(per_proxy, 100))
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
        invite = self._sync_invite_code()
        self.running = True
        self.stop_event.clear()
        self._set_controls_running(True)
        self._append_log(f"--- Bắt đầu lượt chạy: mục tiêu {total} tài khoản ---")
        self._append_log(
            f"Kho proxy: {len(pool)} proxy, mỗi proxy {per_proxy} tài "
            "khoản (tính cả thành công lẫn thất bại) — kho do bạn "
            "đảm bảo là proxy US, tool không kiểm tra lại IP thoát."
        )
        threading.Thread(
            target=self._run_loop,
            args=(total, pool, per_proxy, invite),
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
            if self.helper_process is not None:
                try:
                    self.helper_process.terminate()
                except OSError:
                    pass
        self.destroy()

    def _run_loop(self, total, pool, per_proxy, invite):
        helper = Path(__file__).with_name("browser_helper.py")
        success = 0
        fail = 0
        rotation = _ProxyRotation(pool, per_proxy)
        written_proxy_index = None
        # Chuỗi mã mời: tài khoản sắp chạy sẽ quy đổi pending_code
        # (mã của tài khoản trước, mở đầu bằng mã trong ô Mã mời);
        # mã mời nó tự lấy được sẽ thành pending_code cho tài
        # khoản sau và được ghi đè vào invite.txt để nối chuỗi
        # sang cả lượt chạy sau.
        pending_code = (invite or "").strip().upper()
        for index in range(1, total + 1):
            if self.stop_event.is_set():
                break
            proxy_index, line = rotation.current()
            if proxy_index != written_proxy_index:
                written_proxy_index = proxy_index
                try:
                    _write_active_proxy(line)
                    note = ""
                except OSError as error:
                    note = f" (không ghi được proxy.txt: {error})"
                self.events.put(
                    (
                        "log",
                        f"Proxy #{proxy_index + 1}/{len(pool)}: "
                        f"{_proxy_display(line)} — dùng cho các tài "
                        f"khoản tiếp theo.{note}",
                    )
                )
            self.events.put(
                (
                    "progress",
                    f"Tài khoản {index}/{total} · Thành công: {success} · "
                    f"Thất bại: {fail}",
                )
            )
            try:
                address, muse_password = _create_mailbox_account()
            except Exception as error:
                fail += 1
                self.events.put(
                    ("log", f"[{index}/{total}] Tạo hộp thư thất bại: {error}")
                )
                rotation.advance_after_account()
                continue
            self.events.put(
                (
                    "log",
                    f"[{index}/{total}] Đã tạo mail {address} — bắt đầu đăng ký.",
                )
            )
            try:
                _write_redeem_code(pending_code)
            except OSError:
                pass
            helper_ok, burned, cookie, own_invite = self._run_helper(
                helper, address, index, total
            )
            if helper_ok:
                try:
                    mail_password = ""
                    try:
                        mail_password = (
                            keyring.get_password(MAIL_SERVICE, address)
                            or ""
                        )
                    except Exception:
                        mail_password = ""
                    _append_spreadsheet_row(
                        address,
                        muse_password,
                        mail_password=mail_password,
                        cookie=cookie,
                        invite=own_invite,
                    )
                    success += 1
                    self.events.put(
                        (
                            "log",
                            f"[{index}/{total}] ✔ Đăng ký thành công {address} "
                            "— đã lưu vào sheet Excel.",
                        )
                    )
                except RuntimeError as error:
                    fail += 1
                    self.events.put(
                        (
                            "log",
                            f"[{index}/{total}] {address} đã đăng ký xong nhưng "
                            f"không ghi được sheet Excel: {error}",
                        )
                    )
            else:
                fail += 1
                if self.stop_event.is_set():
                    self.events.put(
                        ("log", f"[{index}/{total}] Đã dừng theo yêu cầu.")
                    )
                else:
                    self.events.put(
                        (
                            "log",
                            f"[{index}/{total}] ✘ Đăng ký thất bại cho {address} "
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
            if burned:
                self.events.put(
                    (
                        "log",
                        f"Proxy #{proxy_index + 1} bị Meta chặn — "
                        "đổi sang proxy kế từ tài khoản sau.",
                    )
                )
            rotation.advance_after_account(burned=burned)
            if index < total and not self.stop_event.is_set():
                time.sleep(PAUSE_BETWEEN_ACCOUNTS_SECONDS)
        self.events.put(
            ("finished", (success, fail, self.stop_event.is_set()))
        )

    def _run_helper(self, helper, address, index, total):
        try:
            process = subprocess.Popen(
                [sys.executable, str(helper), address, "--auto"],
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
            return False, False, "", ""
        self.helper_process = process
        signals = {"burn": False, "cookie": "", "invite": ""}

        def forward_output():
            for line in process.stdout:
                line = line.strip()
                if line:
                    if _line_burns_proxy(line):
                        signals["burn"] = True
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
        return (
            process.returncode == 0 and not self.stop_event.is_set()
        ), signals["burn"], signals["cookie"], signals["invite"]

    def _process_events(self):
        try:
            while True:
                kind, payload = self.events.get_nowait()
                if kind == "log":
                    self._append_log(payload)
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
                    self.progress.set(payload)
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
