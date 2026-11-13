"""Tab Muse.ai: đăng ký tài khoản Muse.ai theo lô.

Cách đăng ký chính: đăng nhập Muse bằng tài khoản Meta đã đăng
kí — chọn file Excel Meta (tab Meta tạo ra file đó), tool lấy
mail trong file đăng nhập rồi nộp thẻ đã lưu, không tạo mail
mới. Vì vậy tab này KHÔNG có ô Temp mail (mail lấy từ file).
Nút Bắt đầu đăng ký là cách cũ: tạo mail tạm mới (tempmail.plus)
rồi đăng ký trực tiếp. Setup chạy (proxy, luồng, kết nối, số
lượng) nằm ở SetupPanel riêng của tab này — kho proxy là
proxies.txt / bad-proxies.txt (file cũ). Thẻ + mã mời chỉ có ở
tab này.
"""

import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

import paths
import storage
from coordinator import RunSettings
from ui.setup_panel import SetupError, SetupPanel


class MuseTab(ttk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, padding=16)
        self.app = app
        self.coordinator = app.coordinator
        self.columnconfigure(0, weight=1)

        ttk.Label(
            self,
            text="Đăng ký Muse.ai theo lô",
            font=("Segoe UI", 16, "bold"),
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            self,
            text=(
                "Đăng nhập Muse bằng tài khoản Meta đã đăng kí: chọn file Excel "
                "Meta ở ô bên dưới (tab Meta tạo ra file đó), tool lấy mail trong "
                "file đăng nhập rồi nộp thẻ đã lưu — không tạo mail mới. Nút Bắt "
                "đầu đăng ký là cách cũ: tạo mail tạm mới rồi đăng ký trực tiếp "
                "(Muse.ai có thể từ chối email tạm)."
            ),
            wraplength=720,
        ).grid(row=1, column=0, sticky="ew", pady=(6, 12))

        self.panel = SetupPanel(
            self,
            self.coordinator,
            pool_file=paths.data_file("proxies.txt"),
            bad_file=paths.data_file("bad-proxies.txt"),
            pool_fallback_file=paths.data_file("proxy.txt"),
            show_provider=False,
        )
        self.panel.grid(row=2, column=0, sticky="ew")

        excel_row = ttk.Frame(self)
        excel_row.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        excel_row.columnconfigure(1, weight=1)
        ttk.Label(excel_row, text="File Excel Meta:").grid(
            row=0, column=0, sticky="w"
        )
        self.meta_excel_entry = ttk.Entry(excel_row, width=36)
        self.meta_excel_entry.grid(row=0, column=1, sticky="ew", padx=(6, 6))
        self.browse_excel_button = ttk.Button(
            excel_row, text="Chọn file…", command=self._browse_meta_excel
        )
        self.browse_excel_button.grid(row=0, column=2)
        self.muse_login_button = ttk.Button(
            excel_row,
            text="Đăng nhập Muse",
            command=self._start_muse_login,
        )
        self.muse_login_button.grid(row=0, column=3, padx=(6, 0))

        card_row = ttk.Frame(self)
        card_row.grid(row=4, column=0, sticky="ew", pady=(10, 0))
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
        saved_card = storage._load_card_from_keyring()
        if saved_card:
            self.card_number_entry.insert(0, saved_card.get("number", ""))
            self.card_expiry_entry.insert(0, saved_card.get("expiry", ""))
            self.card_cvv_entry.insert(0, saved_card.get("cvv", ""))
            self.card_zip_entry.insert(0, saved_card.get("zip", ""))
        else:
            self.card_zip_entry.insert(0, "10001")

        muse_row = ttk.Frame(self)
        muse_row.grid(row=5, column=0, sticky="ew", pady=(10, 0))
        ttk.Label(muse_row, text="Giới hạn / thẻ:").pack(side="left")
        self.card_limit_var = tk.IntVar(value=0)
        self.card_limit_spin = ttk.Spinbox(
            muse_row, from_=0, to=100, width=5,
            textvariable=self.card_limit_var,
        )
        self.card_limit_spin.pack(side="left", padx=(6, 12))
        ttk.Label(muse_row, text="Mã mời:").pack(side="left")
        self.invite_entry = ttk.Entry(muse_row, width=10)
        self.invite_entry.pack(side="left", padx=(6, 12))
        self.single_invite_var = tk.BooleanVar(value=False)
        self.single_invite_check = ttk.Checkbutton(
            muse_row,
            text="Chỉ dùng 1 mã mời này",
            variable=self.single_invite_var,
        )
        self.single_invite_check.pack(side="left", padx=(0, 12))
        try:
            saved_invite = paths.data_file("invite.txt").read_text(
                encoding="utf-8"
            ).strip()
        except OSError:
            saved_invite = ""
        if saved_invite:
            self.invite_entry.insert(0, saved_invite.splitlines()[0].strip())
        saved_excel = storage._load_meta_excel_path_text()
        if saved_excel:
            self.meta_excel_entry.insert(0, saved_excel)

        buttons = ttk.Frame(self)
        buttons.grid(row=6, column=0, sticky="ew", pady=(12, 0))
        self.start_button = ttk.Button(
            buttons, text="Bắt đầu đăng ký", command=self._start_registration
        )
        self.start_button.pack(side="left")
        self.open_sheet_button = ttk.Button(
            buttons,
            text="Mở sheet Excel",
            command=self._open_spreadsheet,
            state=(
                "normal"
                if storage.SPREADSHEET_FILE.exists()
                else "disabled"
            ),
        )
        self.open_sheet_button.pack(side="left", padx=(8, 0))

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
        old_card = storage._load_card_from_keyring()
        try:
            storage._save_card_to_keyring(card)
        except Exception as error:
            messagebox.showerror(
                "Không lưu được thẻ",
                f"Không ghi được vào Windows Credential Manager: {error}",
            )
            return
        if old_card.get("number") != number:
            # Đổi sang thẻ khác: giới hạn tính theo từng thẻ nên
            # bộ đếm tài khoản thành công của thẻ cũ reset về 0.
            storage._save_card_success_count(0)
        self.card_zip_entry.delete(0, "end")
        self.card_zip_entry.insert(0, zip_code)
        self.app.append_log(
            "Đã lưu thẻ vào Windows Credential Manager — tool sẽ tự "
            "điền ở bước xác minh thẻ, không cần gõ lại."
        )

    def _delete_card(self):
        storage._delete_card_from_keyring()
        storage._save_card_success_count(0)
        for entry in (
            self.card_number_entry,
            self.card_expiry_entry,
            self.card_cvv_entry,
            self.card_zip_entry,
        ):
            entry.delete(0, "end")
        self.card_zip_entry.insert(0, "10001")
        self.app.append_log(
            "Đã xoá thẻ đã lưu — bước thẻ sẽ chờ bạn nhập tay."
        )

    def _sync_invite_code(self):
        """Lưu mã mời (nếu có) vào invite.txt để lần sau tự điền lại."""
        code = self.invite_entry.get().strip()
        invite_file = paths.data_file("invite.txt")
        try:
            if code:
                invite_file.write_text(code + "\n", encoding="utf-8")
            elif invite_file.exists():
                invite_file.unlink()
        except OSError:
            pass
        return code

    def set_invite(self, code):
        """Ô Mã mời luôn phản ánh đầu chuỗi hiện tại: mã của tài
        khoản thành công gần nhất."""
        if self.coordinator.running:
            self.invite_entry.configure(state="normal")
        self.invite_entry.delete(0, "end")
        self.invite_entry.insert(0, code)
        if self.coordinator.running:
            self.invite_entry.configure(state="disabled")

    def _start_registration(self):
        if self.coordinator.running:
            return
        try:
            setup = self.panel.read_setup()
        except SetupError as error:
            messagebox.showwarning(error.title, error.message)
            return
        try:
            card_limit = int(self.card_limit_var.get())
        except (tk.TclError, ValueError):
            card_limit = 0
        card_limit = max(0, min(card_limit, 100))
        # Giới hạn tính theo thẻ ĐANG LƯU (một thẻ duy nhất): chỉ
        # đếm khi có thẻ lưu sẵn; không có thẻ lưu thì bước thẻ là
        # nhập tay nên giới hạn này không áp dụng. Để 0 là không
        # giới hạn.
        saved_card = storage._load_card_from_keyring()
        card_tracked = bool(saved_card) and card_limit > 0
        card_used = (
            storage._load_card_success_count() if card_tracked else 0
        )
        if card_tracked and card_used >= card_limit:
            messagebox.showwarning(
                "Thẻ đã đủ giới hạn",
                f"Thẻ đang lưu đã dùng cho {card_used} tài khoản thành "
                f"công (giới hạn {card_limit} / thẻ). Đổi thẻ ở hàng "
                "Thẻ rồi bấm Lưu thẻ để bộ đếm reset về 0, rồi chạy lại.",
            )
            return
        invite = self._sync_invite_code()
        if not invite:
            seed = storage._last_invite_from_sheet()
            if seed:
                self.invite_entry.delete(0, "end")
                self.invite_entry.insert(0, seed)
                invite = self._sync_invite_code()
                self.app.append_log(
                    f"Ô Mã mời trống — tự lấy mã {seed} từ dòng "
                    "cuối trong sheet làm đầu chuỗi."
                )
        single_invite = bool(self.single_invite_var.get())
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
            meta=False,
            invite=invite,
            single_invite=single_invite,
            card_limit=card_limit,
            card_used=card_used,
            card_tracked=card_tracked,
            bad_file=self.panel.bad_file,
            bad_mapping=setup["bad_mapping"],
            ip_target=self.panel,
        )
        self.app.on_run_started(self.panel)
        if total == 0:
            if direct:
                self.app.append_log(
                    "--- Bắt đầu lượt chạy: KHÔNG giới hạn số tài "
                    "khoản — chạy tới khi bấm Dừng (chế độ IP máy "
                    "không có kho proxy để dùng hết) ---"
                )
            else:
                self.app.append_log(
                    "--- Bắt đầu lượt chạy: KHÔNG giới hạn số tài "
                    "khoản — chạy tới khi dùng hết lượt của toàn bộ "
                    "proxy trong danh sách (bấm Dừng để ngắt bất cứ "
                    "lúc nào) ---"
                )
        else:
            self.app.append_log(
                f"--- Bắt đầu lượt chạy: mục tiêu {total} tài khoản ---"
            )
        if direct:
            self.app.append_log(
                "Chế độ kết nối: IP MÁY (VPN) — chạy trực tiếp qua "
                "IP của máy này, không qua proxy. Kiểm tra VPN đã "
                "bật và đang thoát ở US trước khi chạy; mọi luồng "
                "dùng chung 1 IP nên IP bị Meta chặn là cả lượt "
                "chạy dừng ngay."
            )
        else:
            self.app.append_log(
                f"Kho proxy: {len(pool)} proxy, mỗi proxy {per_proxy} tài "
                "khoản (tính cả thành công lẫn thất bại) — kho do bạn "
                "đảm bảo là proxy US, tool không kiểm tra lại IP thoát."
            )
        self.app.append_log(
            f"Dịch vụ temp mail cho lượt chạy này: {setup['provider_label']}."
        )
        if single_invite:
            if invite:
                self.app.append_log(
                    f"Chế độ mã mời: CHỈ DÙNG 1 MÃ — mọi tài khoản "
                    f"đều quy đổi mã {invite}; mã từng tài khoản tự "
                    "lấy được chỉ lưu vào sheet, không nối chuỗi."
                )
            else:
                self.app.append_log(
                    "Chế độ mã mời: CHỈ DÙNG 1 MÃ — nhưng chưa có "
                    "mã nào (ô Mã mời trống, sheet cũng không có), "
                    "nên các tài khoản sẽ không quy đổi mã mời nào."
                )
        if card_tracked:
            self.app.append_log(
                f"Giới hạn thẻ đang lưu: {card_used}/{card_limit} tài "
                "khoản thành công — đủ giới hạn là lượt chạy tự dừng, "
                "không tự đổi thẻ."
            )
        self.coordinator.start(settings)

    def _browse_meta_excel(self):
        """Chọn file Excel Meta đã đăng kí để đăng nhập Muse."""
        # Import lười để không bắt mọi test stub thêm filedialog.
        from tkinter import filedialog

        path = filedialog.askopenfilename(
            title="Chọn file Excel Meta",
            filetypes=[
                ("Excel", "*.xlsx"),
                ("Tất cả", "*.*"),
            ],
        )
        if not path:
            return
        self.meta_excel_entry.delete(0, "end")
        self.meta_excel_entry.insert(0, path)
        try:
            storage._save_meta_excel_path_text(path)
        except OSError as error:
            messagebox.showerror(
                "Không lưu được đường dẫn", str(error)
            )

    def _start_muse_login(self):
        """Nút Đăng nhập Muse: đăng nhập Muse bằng tài khoản Meta
        đã đăng kí trong file Excel đã chọn.

        Lấy email + bí mật hộp thư + dịch vụ theo từng dòng trong
        file (cột Email Meta / Mật khẩu hộp thư / Nhà cung cấp),
        KHÔNG tạo mail mới. Mỗi tài khoản đăng nhập bằng OTP rồi
        nộp thẻ đã lưu; thành công thì đánh dấu cột Muse trong
        file + ghi vào accounts.xlsx như tài khoản đăng ký mới.
        Không quy đổi mã mời. Dùng kho proxy / kết nối / số luồng
        CỦA TAB NÀY.
        """
        if self.coordinator.running:
            return
        excel_text = self.meta_excel_entry.get().strip()
        if excel_text:
            try:
                storage._save_meta_excel_path_text(excel_text)
            except OSError:
                pass
        if not excel_text or not Path(excel_text).is_file():
            messagebox.showwarning(
                "Thiếu file Excel Meta",
                "Chọn file Excel Meta đã đăng kí ở ô File Excel Meta "
                "— tool lấy mail trong đó đăng nhập Muse.",
            )
            return
        pending = storage._pending_meta_logins(excel_text)
        if not pending:
            messagebox.showwarning(
                "Không còn tài khoản",
                "File Excel đã chọn không còn tài khoản Meta nào chưa "
                "đăng nhập Muse (hoặc file không đọc được).",
            )
            return
        try:
            setup = self.panel.read_setup()
        except SetupError as error:
            messagebox.showwarning(error.title, error.message)
            return
        # Đăng nhập vẫn nộp thẻ đã lưu ở bước thẻ — chưa có thẻ thì
        # mỗi tài khoản sẽ chờ nhập tay tới 4 phút.
        if not storage._load_card_from_keyring():
            if not messagebox.askokcancel(
                "Chưa lưu thẻ",
                "Chưa có thẻ nào được lưu. Mỗi tài khoản đăng nhập "
                "xong tới bước thẻ sẽ chờ bạn nhập tay trong cửa sổ "
                "Chromium (tối đa 4 phút). Vẫn chạy?",
            ):
                return
        settings = RunSettings(
            total=setup["total"],
            pool=setup["pool"],
            per_proxy=setup["per_proxy"],
            lanes=setup["lanes"],
            provider_id=setup["provider_id"],
            direct=setup["direct"],
            meta_login=True,
            meta_excel_path=excel_text,
            card_limit=0,
            card_used=0,
            card_tracked=False,
            bad_file=self.panel.bad_file,
            bad_mapping=setup["bad_mapping"],
            ip_target=self.panel,
        )
        self.app.on_run_started(self.panel)
        self.app.append_log(
            "--- Bắt đầu lượt ĐĂNG NHẬP MUSE bằng tài khoản Meta: "
            f"{len(pending)} tài khoản chưa đăng nhập trong file "
            f"{Path(excel_text).name} ---"
        )
        if setup["total"] == 0:
            self.app.append_log(
                "Đăng nhập hết các tài khoản chưa đăng nhập trong file."
            )
        else:
            self.app.append_log(
                f"Giới hạn lượt chạy này: {setup['total']} tài khoản."
            )
        if setup["direct"]:
            self.app.append_log(
                "Chế độ kết nối: IP MÁY (VPN) — chạy trực tiếp qua "
                "IP của máy này, không qua proxy."
            )
        else:
            self.app.append_log(
                f"Kho proxy: {len(setup['pool'])} proxy, mỗi proxy "
                f"{setup['per_proxy']} tài khoản."
            )
        self.coordinator.start(settings)

    def _open_spreadsheet(self):
        if not storage.SPREADSHEET_FILE.exists():
            messagebox.showwarning("Chưa có sheet", "Chưa tạo được file Excel.")
            return
        try:
            import os

            os.startfile(str(storage.SPREADSHEET_FILE))
        except OSError as error:
            messagebox.showerror("Không mở được Excel", str(error))

    def set_running(self, running):
        self.panel.set_running(running)
        self.card_limit_spin.configure(
            state="disabled" if running else "normal"
        )
        self.invite_entry.configure(
            state="disabled" if running else "normal"
        )
        self.single_invite_check.configure(
            state="disabled" if running else "normal"
        )
        for widget in (
            self.card_number_entry,
            self.card_expiry_entry,
            self.card_cvv_entry,
            self.card_zip_entry,
            self.save_card_button,
            self.delete_card_button,
            self.start_button,
            self.meta_excel_entry,
            self.browse_excel_button,
            self.muse_login_button,
        ):
            widget.configure(state="disabled" if running else "normal")
        self.open_sheet_button.configure(
            state=(
                "disabled"
                if running
                else (
                    "normal"
                    if storage.SPREADSHEET_FILE.exists()
                    else "disabled"
                )
            )
        )
