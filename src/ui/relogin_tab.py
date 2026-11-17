"""Tab đăng nhập lại Muse: lấy lại cookie của mail đã có.

Đọc mail từ kho SQLite (accounts.db) — một dòng một tài khoản,
kèm Mật khẩu hộp thư để lấy OTP. Đăng nhập đúng từng mail,
không proxy, không nộp thẻ. Cookie mới ghi ra một file Excel
khác, chỉ hai cột: Email và Cookie đăng nhập. Muốn đưa thêm
mail từ file Excel vào kho thì bấm "Nhập từ Excel…".
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
                "Lấy lại cookie của mail đã có trong kho (accounts.db). "
                "Tool đăng nhập đúng mail đó, lấy OTP từ hộp thư, "
                "rồi ghi cookie mới ra file khác — chỉ hai cột Email "
                "và Cookie đăng nhập. Không dùng proxy, không nộp thẻ."
            ),
            wraplength=720,
        ).grid(row=1, column=0, sticky="ew", pady=(6, 0))

        source = ttk.Frame(self)
        source.grid(row=2, column=0, sticky="ew", pady=(14, 0))
        source.columnconfigure(0, weight=1)
        self.store_label = ttk.Label(source, text="")
        self.store_label.grid(row=0, column=0, sticky="w")
        self.import_button = ttk.Button(
            source, text="Nhập từ Excel…", command=self._import_from_excel
        )
        self.import_button.grid(row=0, column=1, sticky="e")
        self._refresh_store_label()

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
        ttk.Label(run, text="Số luồng:").pack(side="left")
        self.lanes_var = tk.IntVar(value=1)
        self.lanes_spin = ttk.Spinbox(
            run, from_=1, to=10, width=5, textvariable=self.lanes_var
        )
        self.lanes_spin.pack(side="left", padx=(6, 12))
        self.headless_var = tk.BooleanVar(value=False)
        self.headless_check = ttk.Checkbutton(
            run,
            text="Ẩn trình duyệt",
            variable=self.headless_var,
        )
        self.headless_check.pack(side="left", padx=(0, 12))
        self.start_button = ttk.Button(
            run, text="Đăng nhập lại", command=self._start
        )
        self.start_button.pack(side="left")
        self.open_output_button = ttk.Button(
            run, text="Mở file xuất", command=self._open_output
        )
        self.open_output_button.pack(side="left", padx=(8, 0))

        saved_output = storage._load_relogin_output_path()
        if saved_output:
            self.output_entry.insert(0, saved_output)

    def _refresh_store_label(self):
        try:
            total = len(storage._relogin_accounts())
        except Exception:
            total = 0
        self.store_label.configure(
            text=f"Kho: accounts.db — {total} tài khoản"
        )

    def _import_from_excel(self):
        from tkinter import filedialog

        path = filedialog.askopenfilename(
            title="Chọn file Excel để nhập vào kho",
            filetypes=[("Excel", "*.xlsx"), ("Tất cả", "*.*")],
        )
        if not path:
            return
        try:
            result = storage.import_accounts_from_excel(path)
        except Exception as error:
            messagebox.showerror("Không nhập được file", str(error))
            return
        try:
            storage._save_relogin_source_path(path)
        except OSError:
            pass
        self._refresh_store_label()
        messagebox.showinfo(
            "Nhập xong",
            f"Thêm {result['imported']} mail mới, cập nhật "
            f"{result['updated']} mail đã có, bỏ qua "
            f"{result['skipped']} dòng không có mail. Kho đang có "
            f"{result['total']} tài khoản.",
        )

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
        output = self.output_entry.get().strip()
        if not output:
            messagebox.showwarning(
                "Thiếu file xuất",
                "Chọn file Excel để ghi Email và Cookie đăng nhập.",
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
        try:
            lanes = int(self.lanes_var.get())
        except (tk.TclError, ValueError):
            lanes = 1
        lanes = max(1, min(lanes, 10))
        rows = storage._relogin_accounts()
        if not rows:
            messagebox.showwarning(
                "Không có mail",
                "Kho chưa có mail nào. Bấm 'Nhập từ Excel…' để đưa "
                "mail vào kho.",
            )
            return
        done = storage._relogin_done_emails(output)
        pending = [
            row for row in rows if row["email"].lower() not in done
        ]
        if not pending:
            messagebox.showinfo(
                "Đã lấy hết cookie",
                f"Cả {len(rows)} mail đã có cookie trong file xuất. "
                "Lượt này không đăng nhập lại.",
            )
            return
        try:
            storage._save_relogin_output_path(output)
        except OSError:
            pass
        settings = RunSettings(
            total=count,
            direct=True,
            lanes=lanes,
            relogin=True,
            relogin_output_path=output,
            headless=bool(self.headless_var.get()),
        )
        try:
            self.app.on_run_started(None)
            planned = (
                len(pending) if count == 0 else min(count, len(pending))
            )
            skipped = len(rows) - len(pending)
            self.app.append_log(
                "--- Bắt đầu ĐĂNG NHẬP LẠI: "
                f"{planned} mail chưa có cookie trong file xuất ---"
            )
            if skipped:
                self.app.append_log(
                    f"Bỏ qua {skipped} mail đã có cookie trong "
                    f"{Path(output).name}."
                )
            self.app.append_log(
                f"Chạy {lanes} luồng. Không dùng proxy. OTP lấy từ "
                "hộp thư của đúng mail đó. Cookie ghi vào "
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
            self.import_button,
            self.output_entry,
            self.browse_output_button,
            self.count_entry,
            self.lanes_spin,
            self.headless_check,
            self.start_button,
        ):
            widget.configure(state=state)
