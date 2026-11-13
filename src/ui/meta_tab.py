"""Tab Meta: đăng ký tài khoản Meta ở meta.ai.

Đăng kí Meta chạy riêng: tạo tài khoản Meta rồi lưu vào
meta_accounts.xlsx, KHÔNG tự qua Muse. Đăng nhập Muse bằng các
tài khoản này là việc của TAB MUSE.AI: sang đó, chọn file Excel
Meta rồi bấm Đăng nhập Muse. Setup chạy (proxy, luồng, kết nối,
temp mail, số lượng) nằm ở SetupPanel riêng của tab này — kho
proxy là proxies_meta.txt / bad-proxies-meta.txt, tách khỏi kho
của tab Muse. Đăng kí Meta không dùng thẻ, không dùng mã mời.
CAPTCHA giải tự động qua CapSolver, selfie up từ file — không
có bước làm tay.
"""

import threading
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

import paths
import storage
from coordinator import RunSettings
from ui.setup_panel import SetupError, SetupPanel


class MetaTab(ttk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, padding=16)
        self.app = app
        self.coordinator = app.coordinator
        self.columnconfigure(0, weight=1)

        ttk.Label(
            self,
            text="Đăng kí tài khoản Meta (meta.ai)",
            font=("Segoe UI", 16, "bold"),
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            self,
            text=(
                "Đăng kí Meta chạy riêng: tạo tài khoản Meta ở meta.ai "
                "rồi lưu vào meta_accounts.xlsx, không qua Muse. "
                "CAPTCHA giải tự động qua CapSolver, selfie up từ "
                "file — không có bước làm tay."
            ),
            wraplength=720,
        ).grid(row=1, column=0, sticky="ew", pady=(6, 0))

        # Setup (proxy/luồng/kết nối) xuống cuối: vẫn đầy đủ mọi
        # trường, ít dùng tới thì cuộn xuống khi cần.
        self.panel = SetupPanel(
            self,
            self.coordinator,
            pool_file=paths.data_file("proxies_meta.txt"),
            bad_file=paths.data_file("bad-proxies-meta.txt"),
        )
        self.panel.grid(row=5, column=0, sticky="ew", pady=(10, 0))

        key_row = ttk.Frame(self)
        key_row.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        key_row.columnconfigure(1, weight=1)
        ttk.Label(key_row, text="CapSolver API key:").grid(
            row=0, column=0, sticky="w"
        )
        self.capsolver_entry = ttk.Entry(key_row, width=36, show="•")
        self.capsolver_entry.grid(row=0, column=1, sticky="ew", padx=(6, 6))
        self.save_capsolver_button = ttk.Button(
            key_row, text="Lưu key", command=self._save_capsolver_key
        )
        self.save_capsolver_button.grid(row=0, column=2)
        self.check_balance_button = ttk.Button(
            key_row,
            text="Số dư",
            command=self._check_capsolver_balance,
        )
        self.check_balance_button.grid(row=0, column=3, padx=(6, 0))

        selfie_row = ttk.Frame(self)
        selfie_row.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        selfie_row.columnconfigure(1, weight=1)
        ttk.Label(selfie_row, text="Ảnh selfie:").grid(
            row=0, column=0, sticky="w"
        )
        self.selfie_entry = ttk.Entry(selfie_row, width=36)
        self.selfie_entry.grid(row=0, column=1, sticky="ew", padx=(6, 6))
        self.browse_selfie_button = ttk.Button(
            selfie_row, text="Chọn ảnh…", command=self._browse_selfie
        )
        self.browse_selfie_button.grid(row=0, column=2)
        self.meta_start_button = ttk.Button(
            selfie_row,
            text="Đăng kí Meta",
            command=self._start_meta_registration,
        )
        self.meta_start_button.grid(row=0, column=3, padx=(6, 0))

        buttons = ttk.Frame(self)
        buttons.grid(row=4, column=0, sticky="ew", pady=(12, 0))
        self.open_sheet_button = ttk.Button(
            buttons,
            text="Mở sheet Meta",
            command=self._open_spreadsheet,
            state=(
                "normal"
                if storage.META_SPREADSHEET_FILE.exists()
                else "disabled"
            ),
        )
        self.open_sheet_button.pack(side="left")

        saved_key = storage._load_capsolver_key()
        if saved_key:
            self.capsolver_entry.insert(0, saved_key)
        saved_selfie = storage._load_selfie_path_text()
        if saved_selfie:
            self.selfie_entry.insert(0, saved_selfie)

    def _save_capsolver_key(self):
        """Lưu CapSolver API key vào Windows Credential Manager.

        Key không in ra log, không ghi ra file — helper đọc lại
        qua keyring khi gặp CAPTCHA trong luồng Meta.
        """
        key = self.capsolver_entry.get().strip()
        if not key:
            messagebox.showwarning(
                "Thiếu API key",
                "Nhập CapSolver API key rồi bấm Lưu key.",
            )
            return
        try:
            storage._save_capsolver_key(key)
        except Exception as error:
            messagebox.showerror(
                "Không lưu được key",
                f"Không ghi được vào Windows Credential Manager: {error}",
            )
            return
        self.app.append_log(
            "Đã lưu CapSolver API key vào Windows Credential Manager."
        )

    def _check_capsolver_balance(self):
        """Nút Số dư: hỏi số dư CapSolver trên luồng riêng.

        Nút này không bị khoá khi đang chạy (giống nút Kiểm tra
        IP) — kiểm tra giữa lượt chạy vẫn được.
        """
        key = self.capsolver_entry.get().strip() or storage._load_capsolver_key()
        if not key:
            messagebox.showwarning(
                "Thiếu API key",
                "Nhập CapSolver API key rồi bấm Lưu key trước.",
            )
            return

        def _work():
            try:
                import capsolver

                balance = capsolver.get_balance(key)
            except Exception as error:
                self.coordinator.events.put(
                    ("log", f"Kiểm tra số dư CapSolver lỗi: {error}")
                )
                return
            self.coordinator.events.put(
                ("log", f"Số dư CapSolver: ${balance:.2f}.")
            )

        threading.Thread(target=_work, daemon=True).start()

    def _browse_selfie(self):
        """Chọn file ảnh selfie cho bước up selfie của luồng Meta."""
        # Import lười để không bắt mọi test stub thêm filedialog.
        from tkinter import filedialog

        path = filedialog.askopenfilename(
            title="Chọn ảnh selfie",
            filetypes=[
                ("Ảnh", "*.jpg *.jpeg *.png *.webp *.bmp"),
                ("Tất cả", "*.*"),
            ],
        )
        if not path:
            return
        self.selfie_entry.delete(0, "end")
        self.selfie_entry.insert(0, path)
        try:
            storage._save_selfie_path_text(path)
        except OSError as error:
            messagebox.showerror(
                "Không lưu được đường dẫn", str(error)
            )

    def _start_meta_registration(self):
        """Nút Đăng kí Meta: chạy riêng lượt đăng ký tài khoản Meta.

        Chỉ đăng ký Meta ở meta.ai rồi lưu vào meta_accounts.xlsx
        — KHÔNG tự qua Muse (đăng nhập Muse là việc của tab Muse.ai).
        Dùng kho proxy / chế độ kết nối / số luồng / dịch vụ
        temp mail / số tài khoản CỦA TAB NÀY, nhưng không dùng
        thẻ và không dùng mã mời. Bắt buộc phải có CapSolver API
        key + file ảnh selfie vì luồng Meta không có bước làm
        tay nào.
        """
        if self.coordinator.running:
            return
        key = storage._load_capsolver_key()
        if not key:
            entered = self.capsolver_entry.get().strip()
            if entered:
                try:
                    storage._save_capsolver_key(entered)
                    key = entered
                except Exception:
                    key = ""
            if not key:
                messagebox.showwarning(
                    "Thiếu CapSolver API key",
                    "Đăng kí Meta giải CAPTCHA tự động qua CapSolver "
                    "— nhập API key rồi bấm Lưu key trước khi chạy.",
                )
                return
        selfie_text = self.selfie_entry.get().strip()
        if selfie_text:
            try:
                storage._save_selfie_path_text(selfie_text)
            except OSError:
                pass
        if not selfie_text or not Path(selfie_text).is_file():
            messagebox.showwarning(
                "Thiếu ảnh selfie",
                "Chọn file ảnh selfie ở ô Ảnh selfie — bước up "
                "selfie dùng đúng file này cho mọi tài khoản.",
            )
            return
        try:
            setup = self.panel.read_setup()
        except SetupError as error:
            messagebox.showwarning(error.title, error.message)
            return
        total = setup["total"]
        direct = setup["direct"]
        pool = setup["pool"]
        per_proxy = setup["per_proxy"]
        settings = RunSettings(
            total=total,
            pool=pool,
            per_proxy=per_proxy,
            lanes=setup["lanes"],
            provider_id=setup["provider_id"],
            direct=direct,
            meta=True,
            bad_file=self.panel.bad_file,
            bad_mapping=setup["bad_mapping"],
            ip_target=self.panel,
        )
        try:
            self.app.on_run_started(self.panel)
            if total == 0:
                self.app.append_log(
                    "--- Bắt đầu lượt ĐĂNG KÍ META: KHÔNG giới hạn số "
                    "tài khoản — chạy tới khi dùng hết lượt của toàn bộ "
                    "proxy trong danh sách (bấm Dừng để ngắt bất cứ lúc "
                    "nào) ---"
                )
            else:
                self.app.append_log(
                    "--- Bắt đầu lượt ĐĂNG KÍ META: mục tiêu "
                    f"{total} tài khoản ---"
                )
            if direct:
                self.app.append_log(
                    "Chế độ kết nối: IP MÁY (VPN) — chạy trực tiếp qua "
                    "IP của máy này, không qua proxy."
                )
            else:
                self.app.append_log(
                    f"Kho proxy: {len(pool)} proxy, mỗi proxy {per_proxy} tài "
                    "khoản."
                )
            self.app.append_log(
                f"Dịch vụ temp mail cho lượt chạy này: {setup['provider_label']}."
            )
            self.app.append_log(
                "Đăng kí Meta xong là lưu vào meta_accounts.xlsx, không "
                "tự qua Muse — sang tab Muse.ai, chọn file này ở ô File "
                "Excel Meta rồi bấm Đăng nhập Muse."
            )
            if not self.coordinator.start(settings):
                # Đã có lượt chạy khác xen vào: mở khoá để bấm lại được.
                self.app.set_running(False)
                self.app.append_log(
                    "⛔ Đã có một lượt chạy khác đang chạy — bấm lại sau "
                    "khi lượt đó xong."
                )
        except Exception as error:
            # Lỗi trong lúc mở khoá/khởi động (trước đây tkinter nuốt
            # lặng, nút vẫn sáng mà không chạy gì): mở khoá lại và
            # ghi rõ lỗi vào log để thấy ngay.
            try:
                self.app.set_running(False)
            except Exception:
                pass
            self.app.append_log(
                "⛔ Không bắt đầu được lượt ĐĂNG KÍ META: "
                f"{type(error).__name__}: {error}"
            )

    def _open_spreadsheet(self):
        if not storage.META_SPREADSHEET_FILE.exists():
            messagebox.showwarning("Chưa có sheet", "Chưa tạo được file Excel.")
            return
        try:
            import os

            os.startfile(str(storage.META_SPREADSHEET_FILE))
        except OSError as error:
            messagebox.showerror("Không mở được Excel", str(error))

    def set_running(self, running):
        self.panel.set_running(running)
        for widget in (
            self.capsolver_entry,
            self.save_capsolver_button,
            self.selfie_entry,
            self.browse_selfie_button,
            self.meta_start_button,
        ):
            widget.configure(state="disabled" if running else "normal")
        # Nút Số dư không bị khoá khi đang chạy (giống nút Kiểm
        # tra IP) — kiểm tra giữa lượt chạy vẫn được.
        self.open_sheet_button.configure(
            state=(
                "disabled"
                if running
                else (
                    "normal"
                    if storage.META_SPREADSHEET_FILE.exists()
                    else "disabled"
                )
            )
        )
