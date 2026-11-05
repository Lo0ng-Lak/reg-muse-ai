import html
import json
import os
import queue
import re
import secrets
import subprocess
import sys
import threading
import urllib.error
import urllib.request
from datetime import datetime
from html.parser import HTMLParser
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, ttk

import keyring
from openpyxl import Workbook, load_workbook


API_BASE = "https://api.mail.tm"
APP_DIR = Path.home() / "AppData" / "Roaming" / "MuseAIHelper"
ACCOUNTS_FILE = APP_DIR / "accounts.json"
SPREADSHEET_FILE = (
    Path.home() / "AppData" / "Local" / "MuseAIHelper" / "accounts.xlsx"
)
MAIL_SERVICE = "MuseAIHelper:mail.tm"
MUSE_SERVICE = "MuseAIHelper:muse.ai"


class _TextExtractor(HTMLParser):
    def __init__(self):
        super().__init__()
        self.parts = []

    def handle_data(self, data):
        text = data.strip()
        if text:
            self.parts.append(text)


def _plain_text(value):
    parser = _TextExtractor()
    parser.feed(value or "")
    return html.unescape("\n".join(parser.parts))


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


def _load_accounts():
    try:
        with ACCOUNTS_FILE.open("r", encoding="utf-8") as file:
            data = json.load(file)
        if not isinstance(data, dict):
            raise ValueError("Dữ liệu tài khoản không hợp lệ.")
        return data
    except FileNotFoundError:
        return {}
    except (json.JSONDecodeError, OSError, ValueError) as error:
        raise RuntimeError(f"Không đọc được dữ liệu ứng dụng: {error}") from error


def _save_accounts(data):
    try:
        APP_DIR.mkdir(parents=True, exist_ok=True)
        temporary_file = ACCOUNTS_FILE.with_suffix(".tmp")
        with temporary_file.open("w", encoding="utf-8") as file:
            json.dump(data, file, ensure_ascii=False, indent=2)
        temporary_file.replace(ACCOUNTS_FILE)
    except OSError as error:
        raise RuntimeError(f"Không lưu được dữ liệu ứng dụng: {error}") from error


def _append_spreadsheet_row(email, password):
    temporary_file = SPREADSHEET_FILE.with_name("accounts.tmp.xlsx")
    try:
        SPREADSHEET_FILE.parent.mkdir(parents=True, exist_ok=True)
        if SPREADSHEET_FILE.exists():
            workbook = load_workbook(SPREADSHEET_FILE)
            sheet = workbook.active
        else:
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Muse.ai"
            sheet.append(["Thời gian tạo", "Email Muse.ai", "Mật khẩu Muse.ai"])

        sheet.append(
            [datetime.now().astimezone().isoformat(timespec="seconds"), email, password]
        )
        workbook.save(temporary_file)
        temporary_file.replace(SPREADSHEET_FILE)
    except Exception as error:
        raise RuntimeError(f"Không ghi được file Excel: {error}") from error
    finally:
        if temporary_file.exists():
            temporary_file.unlink()


class MuseAIHelper(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Muse.ai Account Helper")
        self.geometry("760x620")
        self.minsize(640, 500)
        self.events = queue.Queue()
        self.busy = False
        self.vpn_confirmed = False
        self.messages = []

        try:
            self.accounts = _load_accounts()
        except RuntimeError as error:
            messagebox.showerror("Lỗi", str(error))
            self.accounts = {}

        self._build_ui()
        self._load_saved_addresses()
        self.after(100, self._process_events)

    def _build_ui(self):
        root = ttk.Frame(self, padding=16)
        root.pack(fill="both", expand=True)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(4, weight=1)

        ttk.Label(
            root,
            text="Trợ lý đăng ký Muse.ai",
            font=("Segoe UI", 16, "bold"),
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            root,
            text=(
                "Mail.tm cung cấp email tạm, không phải Gmail. Muse.ai có thể từ chối "
                "email tạm. Bấm 'Tạo mail & đăng ký' để tạo hộp thư mới và mở ngay "
                "form đăng ký bằng email đó; CAPTCHA và các bước phát sinh vẫn do "
                "bạn xử lý."
            ),
            wraplength=700,
        ).grid(row=1, column=0, sticky="ew", pady=(6, 14))

        mailbox = ttk.LabelFrame(root, text="Hộp thư mail.tm", padding=10)
        mailbox.grid(row=2, column=0, sticky="ew")
        mailbox.columnconfigure(1, weight=1)
        ttk.Label(mailbox, text="Email:").grid(row=0, column=0, sticky="w")
        self.mail_email = ttk.Entry(mailbox)
        self.mail_email.grid(row=0, column=1, sticky="ew", padx=8)
        self.create_button = ttk.Button(
            mailbox, text="Tạo mail & đăng ký", command=self._create_mailbox
        )
        self.create_button.grid(row=0, column=2, padx=(0, 6))
        self.refresh_button = ttk.Button(
            mailbox, text="Lấy thư", command=self._refresh_messages
        )
        self.refresh_button.grid(row=0, column=3)

        signup = ttk.LabelFrame(root, text="Tài khoản Muse.ai", padding=10)
        signup.grid(row=3, column=0, sticky="ew", pady=12)
        signup.columnconfigure(1, weight=1)
        self.open_button = ttk.Button(
            signup, text="Mở trang Muse.ai", command=self._open_muse
        )
        self.open_button.grid(row=0, column=0, sticky="w")
        ttk.Label(signup, text="Email đăng ký:").grid(
            row=1, column=0, sticky="w", pady=(8, 0)
        )
        self.muse_email = ttk.Entry(signup)
        self.muse_email.grid(row=1, column=1, sticky="ew", padx=8, pady=(8, 0))
        ttk.Label(signup, text="Mật khẩu Muse.ai:").grid(
            row=2, column=0, sticky="w", pady=(8, 0)
        )
        self.muse_password = ttk.Entry(signup, show="*")
        self.muse_password.grid(row=2, column=1, sticky="ew", padx=8, pady=(8, 0))
        self.show_password = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            signup,
            text="Hiện mật khẩu",
            variable=self.show_password,
            command=self._toggle_password,
        ).grid(row=2, column=2, sticky="w", pady=(8, 0))
        self.copy_button = ttk.Button(
            signup,
            text="Sao chép email + mật khẩu",
            command=self._copy_credentials,
        )
        self.copy_button.grid(row=3, column=1, sticky="e", padx=8, pady=(10, 0))
        self.open_sheet_button = ttk.Button(
            signup,
            text="Mở sheet Excel",
            command=self._open_spreadsheet,
            state="normal" if SPREADSHEET_FILE.exists() else "disabled",
        )
        self.open_sheet_button.grid(row=3, column=0, sticky="w", pady=(10, 0))

        inbox = ttk.LabelFrame(root, text="Thư đến", padding=10)
        inbox.grid(row=4, column=0, sticky="nsew")
        inbox.columnconfigure(0, weight=1)
        inbox.rowconfigure(0, weight=1)
        self.message_list = tk.Listbox(inbox, height=8)
        self.message_list.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(
            inbox, orient="vertical", command=self.message_list.yview
        )
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.message_list.configure(yscrollcommand=scrollbar.set)
        self.message_list.bind("<Double-Button-1>", self._open_selected_message)

        self.status = tk.StringVar(
            value="Mật khẩu sẽ được lưu trong Windows Credential Manager."
        )
        ttk.Label(root, textvariable=self.status).grid(
            row=5, column=0, sticky="ew", pady=(10, 0)
        )

    def _load_saved_addresses(self):
        mailbox = self.accounts.get("mailbox_email", "")
        muse = self.accounts.get("muse_email", "")
        self.mail_email.insert(0, mailbox)
        self.muse_email.insert(0, muse or mailbox)
        saved_address = muse or mailbox
        if saved_address:
            try:
                password = keyring.get_password(MUSE_SERVICE, saved_address)
            except Exception as error:
                messagebox.showerror("Lỗi", f"Không đọc được mật khẩu đã lưu: {error}")
                return
            if password:
                self.muse_password.insert(0, password)

    def _run_background(self, task):
        if self.busy:
            return
        self.busy = True
        self._set_controls(False)
        self.status.set("Đang xử lý...")

        def run():
            try:
                result = task()
                self.events.put(("success", result))
            except Exception as error:
                self.events.put(("error", str(error)))

        threading.Thread(target=run, daemon=True).start()

    def _process_events(self):
        try:
            while True:
                kind, result = self.events.get_nowait()
                self.busy = False
                self._set_controls(True)
                if kind == "error":
                    self.status.set("Thao tác thất bại.")
                    messagebox.showerror("Lỗi", result)
                else:
                    callback, value = result
                    try:
                        callback(value)
                    except Exception as error:
                        self.status.set("Không hoàn tất được thao tác.")
                        messagebox.showerror("Lỗi", str(error))
        except queue.Empty:
            pass
        self.after(100, self._process_events)

    def _set_controls(self, enabled):
        state = "normal" if enabled else "disabled"
        for control in (
            self.create_button,
            self.refresh_button,
            self.open_button,
            self.copy_button,
        ):
            control.configure(state=state)
        self.open_sheet_button.configure(
            state=state if SPREADSHEET_FILE.exists() else "disabled"
        )

    def _open_muse(self):
        if not self._confirm_surfshark_vpn():
            return
        address = self.muse_email.get().strip()
        if not address:
            messagebox.showwarning(
                "Thiếu email",
                "Hãy tạo hoặc nhập email trước khi mở form đăng ký.",
            )
            return
        self._launch_signup_browser(address)

    def _launch_signup_browser(self, address):
        helper = Path(__file__).with_name("browser_helper.py")
        try:
            subprocess.Popen(
                [sys.executable, str(helper), address],
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
        except OSError as error:
            messagebox.showerror("Không mở được Muse.ai", str(error))
            return False
        self.status.set(
            "Chromium đang nhập email và chờ mã xác minh từ mail.tm. "
            "Hãy để cửa sổ mở; bạn xử lý các bước tiếp theo nếu được yêu cầu."
        )
        return True

    def _confirm_surfshark_vpn(self):
        if self.vpn_confirmed:
            return True
        confirmed = messagebox.askyesno(
            "Xác nhận Surfshark VPN",
            "Hãy kết nối Surfshark VPN trước khi tiếp tục.\n\n"
            "Chỉ chọn Có khi Surfshark hiển thị đã kết nối. "
            "Ứng dụng không thể tự xác minh VPN hoặc định tuyến mạng.",
        )
        if confirmed:
            self.vpn_confirmed = True
        return confirmed

    def _create_mailbox(self):
        def task():
            domains = _request("GET", "/domains?page=1").get("hydra:member", [])
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

        self._run_background(lambda: (self._mailbox_created, task()))

    def _mailbox_created(self, credentials):
        address, muse_password = credentials
        self.mail_email.delete(0, "end")
        self.mail_email.insert(0, address)
        self.muse_email.delete(0, "end")
        self.muse_email.insert(0, address)
        self.muse_password.delete(0, "end")
        self.muse_password.insert(0, muse_password)
        self.accounts["mailbox_email"] = address
        self.accounts["muse_email"] = address
        self._persist_accounts()
        spreadsheet_error = None
        try:
            _append_spreadsheet_row(address, muse_password)
            self.open_sheet_button.configure(state="normal")
        except RuntimeError as error:
            # Ghi Excel chỉ là nhật ký phụ; lỗi ở đây không được chặn
            # luồng đăng ký chính. Cảnh báo được hiển thị sau khi mở
            # trình duyệt để không ngắt quãng thao tác.
            spreadsheet_error = error
        if not self._confirm_surfshark_vpn():
            self.status.set(
                "Đã tạo hộp thư nhưng chưa mở Muse.ai vì chưa xác nhận VPN. "
                "Bấm 'Mở trang Muse.ai' để đăng ký khi đã sẵn sàng."
            )
            if spreadsheet_error is not None:
                messagebox.showerror("Lỗi ghi Excel", str(spreadsheet_error))
            return
        launched = self._launch_signup_browser(address)
        if spreadsheet_error is not None:
            if launched:
                self.status.set(
                    "Đang đăng ký bằng hộp thư mới, nhưng không ghi được "
                    "sheet Excel (thông tin vẫn còn trong ứng dụng và "
                    "Windows Credential Manager)."
                )
            messagebox.showerror("Lỗi ghi Excel", str(spreadsheet_error))

    def _get_mail_token(self, address):
        password = keyring.get_password(MAIL_SERVICE, address)
        if not password:
            raise RuntimeError(
                "Không tìm thấy mật khẩu hộp thư trong Windows Credential Manager."
            )
        response = _request(
            "POST", "/token", {"address": address, "password": password}
        )
        return response.get("token")

    def _refresh_messages(self):
        address = self.mail_email.get().strip()
        if not address:
            messagebox.showwarning("Thiếu email", "Hãy tạo hoặc nhập email mail.tm.")
            return

        def task():
            token = self._get_mail_token(address)
            if not token:
                raise RuntimeError("mail.tm không trả về token truy cập.")
            messages = _request("GET", "/messages?page=1", token=token).get(
                "hydra:member", []
            )
            return messages

        self._run_background(lambda: (self._messages_loaded, task()))

    def _messages_loaded(self, messages):
        self.messages = messages
        self.message_list.delete(0, "end")
        for item in messages:
            sender = item.get("from", {}).get("address", "không rõ người gửi")
            subject = item.get("subject") or "(không có tiêu đề)"
            self.message_list.insert("end", f"{subject} — {sender}")
        self.status.set(f"Đã tải {len(messages)} thư.")

    def _open_selected_message(self, _event):
        selection = self.message_list.curselection()
        if not selection:
            return
        message = self.messages[selection[0]]
        address = self.mail_email.get().strip()

        def task():
            token = self._get_mail_token(address)
            detail = _request("GET", f"/messages/{message['id']}", token=token)
            body = detail.get("text") or _plain_text(
                "\n".join(detail.get("html", []))
            )
            codes = list(dict.fromkeys(re.findall(r"(?<!\d)\d{4,8}(?!\d)", body)))
            return (
                detail.get("subject") or "(không có tiêu đề)",
                detail.get("from", {}).get("address", "không rõ người gửi"),
                body or detail.get("intro", "(thư không có nội dung văn bản)"),
                codes,
            )

        self._run_background(lambda: (self._message_opened, task()))

    def _message_opened(self, content):
        subject, sender, body, codes = content
        window = tk.Toplevel(self)
        window.title(subject[:80])
        window.geometry("640x420")
        if codes:
            ttk.Label(
                window,
                text=f"Mã số có thể dùng: {', '.join(codes)}",
            ).pack(anchor="w", padx=10, pady=(10, 0))
            ttk.Button(
                window,
                text=f"Sao chép mã đầu tiên ({codes[0]})",
                command=lambda: self._copy_code(codes[0]),
            ).pack(anchor="w", padx=10, pady=(6, 0))
        text = tk.Text(window, wrap="word")
        text.pack(fill="both", expand=True, padx=10, pady=10)
        text.insert("1.0", f"Từ: {sender}\n\n{body}")
        text.configure(state="disabled")

    def _toggle_password(self):
        self.muse_password.configure(show="" if self.show_password.get() else "*")

    def _copy_code(self, code):
        try:
            self.clipboard_clear()
            self.clipboard_append(code)
        except Exception as error:
            messagebox.showerror("Không sao chép được mã", str(error))
            return
        self.status.set("Đã sao chép mã vào clipboard.")

    def _copy_credentials(self):
        address = self.muse_email.get().strip()
        password = self.muse_password.get()
        if not address or not password:
            messagebox.showwarning(
                "Chưa có thông tin", "Hãy tạo email và mật khẩu trước."
            )
            return
        try:
            self.clipboard_clear()
            self.clipboard_append(f"Email: {address}\nMật khẩu: {password}")
        except Exception as error:
            messagebox.showerror("Không sao chép được thông tin", str(error))
            return
        self.status.set("Đã sao chép email và mật khẩu vào clipboard.")

    def _open_spreadsheet(self):
        if not SPREADSHEET_FILE.exists():
            messagebox.showwarning("Chưa có sheet", "Chưa tạo được file Excel.")
            return
        try:
            os.startfile(str(SPREADSHEET_FILE))
        except OSError as error:
            messagebox.showerror("Không mở được Excel", str(error))

    def _persist_accounts(self):
        _save_accounts(self.accounts)


if __name__ == "__main__":
    MuseAIHelper().mainloop()
