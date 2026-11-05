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
HELPER_TIMEOUT_SECONDS = 600
PAUSE_BETWEEN_ACCOUNTS_SECONDS = 3


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
        self.geometry("780x560")
        self.minsize(640, 440)
        self.events = queue.Queue()
        self.running = False
        self.stop_event = threading.Event()
        self.helper_process = None
        self.vpn_confirmed = False

        self._build_ui()
        self.after(100, self._process_events)
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self._append_log(
            "Sẵn sàng. Nhập số tài khoản rồi bấm 'Bắt đầu đăng ký'. "
            "Chỉ tài khoản đăng ký thành công mới được lưu vào sheet Excel."
        )

    def _build_ui(self):
        root = ttk.Frame(self, padding=16)
        root.pack(fill="both", expand=True)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(3, weight=1)

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

        controls = ttk.Frame(root)
        controls.grid(row=2, column=0, sticky="ew")
        ttk.Label(controls, text="Số tài khoản:").pack(side="left")
        self.count_var = tk.IntVar(value=5)
        self.count_spin = ttk.Spinbox(
            controls, from_=1, to=100, width=5, textvariable=self.count_var
        )
        self.count_spin.pack(side="left", padx=(6, 12))
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
        log_frame.grid(row=3, column=0, sticky="nsew", pady=(12, 0))
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
            row=4, column=0, sticky="ew", pady=(10, 0)
        )

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

    def _set_controls_running(self, running):
        self.start_button.configure(state="disabled" if running else "normal")
        self.count_spin.configure(state="disabled" if running else "normal")
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
        if not self._confirm_surfshark_vpn():
            return
        self.running = True
        self.stop_event.clear()
        self._set_controls_running(True)
        self._append_log(f"--- Bắt đầu lượt chạy: mục tiêu {total} tài khoản ---")
        threading.Thread(
            target=self._run_loop, args=(total,), daemon=True
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

    def _run_loop(self, total):
        helper = Path(__file__).with_name("browser_helper.py")
        success = 0
        fail = 0
        for index in range(1, total + 1):
            if self.stop_event.is_set():
                break
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
                continue
            self.events.put(
                (
                    "log",
                    f"[{index}/{total}] Đã tạo mail {address} — bắt đầu đăng ký.",
                )
            )
            if self._run_helper(helper, address, index, total):
                try:
                    _append_spreadsheet_row(address, muse_password)
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
            return False
        self.helper_process = process

        def forward_output():
            for line in process.stdout:
                line = line.strip()
                if line:
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
        return process.returncode == 0 and not self.stop_event.is_set()

    def _process_events(self):
        try:
            while True:
                kind, payload = self.events.get_nowait()
                if kind == "log":
                    self._append_log(payload)
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
