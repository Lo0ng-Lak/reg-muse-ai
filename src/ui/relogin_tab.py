"""Tab đăng nhập lại Muse: lấy lại cookie của mail đã có.

Đọc file Excel có cột Email Muse.ai (và Mật khẩu hộp thư để lấy
OTP). Đăng nhập đúng từng mail, không proxy, không nộp thẻ.
Cookie mới ghi ra một file Excel khác, chỉ hai cột: Email và
Cookie đăng nhập.
"""

import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

import storage
from coordinator import RunSettings


class ReloginTab(ttk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, padding=16)
        self.app = app
        self.coordinator = app.coordinator
        self.columnconfigure(0, weight=1)

        ttk.Label(
            self,
            text="Đăng nhập lại Muse — lấy cookie",
            font=("Segoe UI", 16, "bold"),
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            self,
            text=(
                "Chọn file Excel đã có mail Muse (cột Email Muse.ai). "
                "Tool đăng nhập đúng mail đó, lấy OTP từ hộp thư, "
                "rồi ghi cookie mới ra file khác — chỉ hai cột Email "
                "và Cookie đăng nhập. Không dùng proxy, không nộp thẻ."
            ),
            wraplength=720,
        ).grid(row=1, column=0, sticky="ew", pady=(6, 0))

        source = ttk.Frame(self)
        source.grid(row=2, column=0, sticky="ew", pady=(14, 0))
        source.columnconfigure(1, weight=1)
        ttk.Label(source, text="File Excel nguồn:").grid(
            row=0, column=0, sticky="w"
        )
        self.source_entry = ttk.Entry(source)
        self.source_entry.grid(row=0, column=1, sticky="ew", padx=(6, 6))
        self.browse_source_button = ttk.Button(
            source, text="Chọn file…", command=self._browse_source
        )
        self.browse_source_button.grid(row=0, column=2)

        output = ttk.Frame(self)
        output.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        output.columnconfigure(1, weight=1)
        ttk.Label(output, text="File xuất cookie:").grid(
            row=0, column=0, sticky="w"
        )
        self.output_entry = ttk.Entry(output)
        self.output_entry.grid(row=0, column=1, sticky="ew", padx=(6, 6))
        self.browse_output_button = ttk.Button(
            output, text="Chọn nơi lưu…", command=self._browse_output
        )
        self.browse_output_button.grid(row=0, column=2)

        run = ttk.Frame(self)
        run.grid(row=4, column=0, sticky="ew", pady=(14, 0))
        ttk.Label(run, text="Số mail (0 = tất cả):").pack(side="left")
        self.count_entry = ttk.Entry(run, width=8)
        self.count_entry.pack(side="left", padx=(6, 12))
        self.count_entry.insert(0, "0")
        self.start_button = ttk.Button(
            run, text="Đăng nhập lại", command=self._start
        )
        self.start_button.pack(side="left")
        self.open_output_button = ttk.Button(
            run, text="Mở file xuất", command=self._open_output
        )
        self.open_output_button.pack(side="left", padx=(8, 0))

        saved_source = storage._load_relogin_source_path()
        if saved_source:
            self.source_entry.insert(0, saved_source)
        saved_output = storage._load_relogin_output_path()
        if saved_output:
            self.output_entry.insert(0, saved_output)

    def _browse_source(self):
        from tkinter import filedialog

        path = filedialog.askopenfilename(
            title="Chọn file Excel mail Muse",
            filetypes=[("Excel", "*.xlsx"), ("Tất cả", "*.*")],
        )
        if not path:
            return
        self.source_entry.delete(0, "end")
        self.source_entry.insert(0, path)
        try:
            storage._save_relogin_source_path(path)
        except OSError as error:
            messagebox.showerror("Không lưu được đường dẫn", str(error))
        output = self.output_entry.get().strip()
        if not output:
            suggested = str(Path(path).with_name("muse_cookies.xlsx"))
            self.output_entry.insert(0, suggested)

    def _browse_output(self):
        from tkinter import filedialog

        path = filedialog.asksaveasfilename(
            title="Chọn file xuất cookie",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not path:
            return
        self.output_entry.delete(0, "end")
        self.output_entry.insert(0, path)
        try:
            storage._save_relogin_output_path(path)
        except OSError as error:
            messagebox.showerror("Không lưu được đường dẫn", str(error))

    def _start(self):
        if self.coordinator.running:
            return
        source = self.source_entry.get().strip()
        output = self.output_entry.get().strip()
        if not source or not Path(source).is_file():
            messagebox.showwarning(
                "Thiếu file Excel",
                "Chọn file Excel có cột Email Muse.ai.",
            )
            return
        if not output:
            messagebox.showwarning(
                "Thiếu file xuất",
                "Chọn file Excel để ghi Email và Cookie đăng nhập.",
            )
            return
        try:
            same = Path(source).resolve() == Path(output).resolve()
        except OSError:
            same = False
        if same:
            messagebox.showwarning(
                "File xuất trùng file nguồn",
                "Chọn một file khác để ghi cookie. File nguồn giữ "
                "nguyên mail và mật khẩu hộp thư.",
            )
            return
        try:
            count = int(self.count_entry.get().strip() or "0")
        except ValueError:
            messagebox.showwarning(
                "Số mail không hợp lệ",
                "Nhập 0 để chạy tất cả, hoặc một số nguyên dương.",
            )
            return
        if count < 0:
            messagebox.showwarning(
                "Số mail không hợp lệ",
                "Nhập 0 để chạy tất cả, hoặc một số nguyên dương.",
            )
            return
        rows = storage._relogin_accounts(source)
        if not rows:
            messagebox.showwarning(
                "Không có mail",
                "File không có cột Email Muse.ai, hoặc không có "
                "dòng mail nào.",
            )
            return
        try:
            storage._save_relogin_source_path(source)
            storage._save_relogin_output_path(output)
        except OSError:
            pass
        settings = RunSettings(
            total=count,
            direct=True,
            relogin=True,
            relogin_excel_path=source,
            relogin_output_path=output,
        )
        try:
            self.app.on_run_started(None)
            planned = len(rows) if count == 0 else min(count, len(rows))
            self.app.append_log(
                "--- Bắt đầu ĐĂNG NHẬP LẠI: "
                f"{planned} mail trong {Path(source).name} ---"
            )
            self.app.append_log(
                "Không dùng proxy. OTP lấy từ hộp thư của đúng "
                "mail đó. Cookie ghi vào "
                f"{Path(output).name} (cột Email và Cookie đăng nhập)."
            )
            if not self.coordinator.start(settings):
                self.app.set_running(False)
                self.app.append_log(
                    "⛔ Đã có một lượt chạy khác đang chạy — bấm lại sau "
                    "khi lượt đó xong."
                )
        except Exception as error:
            try:
                self.app.set_running(False)
            except Exception:
                pass
            self.app.append_log(
                "⛔ Không bắt đầu được lượt đăng nhập lại: "
                f"{type(error).__name__}: {error}"
            )

    def _open_output(self):
        path = self.output_entry.get().strip()
        if not path or not Path(path).is_file():
            messagebox.showwarning(
                "Chưa có file xuất",
                "Chưa có file cookie. Chạy đăng nhập lại trước.",
            )
            return
        try:
            import os

            os.startfile(path)
        except OSError as error:
            messagebox.showerror("Không mở được Excel", str(error))

    def set_running(self, running):
        state = "disabled" if running else "normal"
        for widget in (
            self.source_entry,
            self.browse_source_button,
            self.output_entry,
            self.browse_output_button,
            self.count_entry,
            self.start_button,
        ):
            widget.configure(state=state)
