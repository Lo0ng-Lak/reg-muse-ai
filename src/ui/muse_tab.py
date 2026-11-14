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
from coordinator import CardRunState, RunSettings
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
                "Meta ở ô bên dưới, tool lấy mail trong file đăng nhập rồi nộp "
                "thẻ đã lưu. Nút Bắt đầu đăng ký là cách cũ: tạo mail tạm mới "
                "rồi đăng ký trực tiếp."
            ),
            wraplength=720,
        ).grid(row=1, column=0, sticky="ew", pady=(6, 0))

        # Setup (proxy/luồng/kết nối) xuống cuối: vẫn đầy đủ mọi
        # trường, ít dùng tới thì cuộn xuống khi cần.
        self.panel = SetupPanel(
            self,
            self.coordinator,
            pool_file=paths.data_file("proxies.txt"),
            bad_file=paths.data_file("bad-proxies.txt"),
            pool_fallback_file=paths.data_file("proxy.txt"),
            show_provider=False,
        )
        self.panel.grid(row=5, column=0, sticky="ew", pady=(10, 0))

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

        # Kho thẻ lên đầu: bảng + ô dán + nút Nhập kho thẻ phải
        # thấy được ngay mà không cần cuộn (cửa sổ 820x980).
        vault_frame = ttk.LabelFrame(self, text="Kho thẻ", padding=10)
        vault_frame.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        vault_frame.columnconfigure(0, weight=1)
        self.vault_tree = ttk.Treeview(
            vault_frame,
            columns=("limit", "used", "status"),
            show="tree headings",
            height=5,
            selectmode="browse",
        )
        self.vault_tree.heading("#0", text="Tên thẻ")
        self.vault_tree.heading("limit", text="Giới hạn")
        self.vault_tree.heading("used", text="Đã dùng")
        self.vault_tree.heading("status", text="Trạng thái")
        self.vault_tree.column("#0", width=180)
        self.vault_tree.column("limit", width=70, anchor="center")
        self.vault_tree.column("used", width=70, anchor="center")
        self.vault_tree.column("status", width=140)
        self.vault_tree.grid(row=0, column=0, sticky="ew")
        self.vault_tree.bind("<<TreeviewSelect>>", self._on_vault_select)
        ttk.Label(
            vault_frame,
            text=(
                "Dán danh sách thẻ vào ô dưới, mỗi dòng một thẻ: "
                "số thẻ|MM/YY|CVV (ZIP mặc định 10001, nhãn tự đặt; "
                "có thể thêm ZIP|nhãn làm trường 4|5). Dòng sai báo "
                "theo số dòng, không hiện số thẻ; ô dán xoá ngay "
                "sau khi nhập."
            ),
            wraplength=700,
        ).grid(row=1, column=0, sticky="w", pady=(8, 4))
        self.bulk_card_text = tk.Text(vault_frame, height=4, width=80)
        self.bulk_card_text.grid(row=2, column=0, sticky="ew")
        self.import_bulk_button = ttk.Button(
            vault_frame,
            text="Nhập kho thẻ",
            command=self._import_bulk_cards,
        )
        self.import_bulk_button.grid(row=3, column=0, sticky="w", pady=(6, 0))
        vault_actions = ttk.Frame(vault_frame)
        vault_actions.grid(row=4, column=0, sticky="ew", pady=(8, 0))
        self.select_card_button = ttk.Button(
            vault_actions,
            text="Chọn thẻ này",
            command=self._select_vault_card,
        )
        self.select_card_button.pack(side="left")
        self.auto_rotate_button = ttk.Button(
            vault_actions,
            text="Tự động xoay",
            command=self._auto_rotate_cards,
        )
        self.auto_rotate_button.pack(side="left", padx=(6, 0))
        self.delete_vault_card_button = ttk.Button(
            vault_actions,
            text="Xoá thẻ",
            command=self._delete_vault_card,
        )
        self.delete_vault_card_button.pack(side="left", padx=(6, 0))
        ttk.Label(vault_actions, text="Giới hạn / thẻ:").pack(
            side="left", padx=(12, 0)
        )
        self.card_limit_var = tk.IntVar(value=0)
        self.card_limit_spin = ttk.Spinbox(
            vault_actions, from_=0, to=100, width=5,
            textvariable=self.card_limit_var,
        )
        self.card_limit_spin.pack(side="left", padx=(6, 6))
        self.save_limit_button = ttk.Button(
            vault_actions,
            text="Lưu giới hạn",
            command=self._save_vault_limit,
        )
        self.save_limit_button.pack(side="left")
        # Giới hạn chung: lấy số ở ô Giới hạn / thẻ rồi áp cho tất
        # cả thẻ trong kho cùng lúc (ghi đè từng thẻ) — không cần
        # chọn thẻ, không thêm hàng mới nên khung Kho thẻ vẫn gọn.
        self.apply_limit_all_button = ttk.Button(
            vault_actions,
            text="Áp dụng cho tất cả thẻ",
            command=self._apply_vault_limit_to_all,
        )
        self.apply_limit_all_button.pack(side="left", padx=(6, 0))
        self._refresh_vault()

        # Hàng chạy gọn trong một dòng: mã mời + Bắt đầu đăng ký
        # + Mở sheet Excel.
        run_row = ttk.Frame(self)
        run_row.grid(row=4, column=0, sticky="ew", pady=(8, 0))
        ttk.Label(run_row, text="Mã mời:").pack(side="left")
        self.invite_entry = ttk.Entry(run_row, width=10)
        self.invite_entry.pack(side="left", padx=(6, 12))
        self.single_invite_var = tk.BooleanVar(value=False)
        self.single_invite_check = ttk.Checkbutton(
            run_row,
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
        self.start_button = ttk.Button(
            run_row, text="Bắt đầu đăng ký", command=self._start_registration
        )
        self.start_button.pack(side="left", padx=(12, 0))
        self.open_sheet_button = ttk.Button(
            run_row,
            text="Mở sheet Excel",
            command=self._open_spreadsheet,
            state=(
                "normal"
                if storage.SPREADSHEET_FILE.exists()
                else "disabled"
            ),
        )
        self.open_sheet_button.pack(side="left", padx=(8, 0))
        saved_excel = storage._load_meta_excel_path_text()
        if saved_excel:
            self.meta_excel_entry.insert(0, saved_excel)

    def _refresh_vault(self):
        """Vẽ lại danh sách kho thẻ từ file kho thẻ.

        Chỉ hiện nhãn + giới hạn + số lần dùng + trạng thái —
        không chữ số nào của thẻ xuất hiện ở giao diện. Trạng
        thái: "Đang chọn" (thẻ chọn tay), "Đủ giới hạn", "Sẵn
        sàng (∞)" (không giới hạn), "Sẵn sàng".
        """
        for item in self.vault_tree.get_children():
            self.vault_tree.delete(item)
        active = storage._load_vault_active()
        self._vault_cards = storage._load_vault()
        self._log_vault_notices_once()
        for card in self._vault_cards:
            if card["id"] == active:
                status = "Đang chọn"
            elif card["limit"] > 0 and card["used"] >= card["limit"]:
                status = "Đủ giới hạn"
            elif card["limit"] == 0:
                status = "Sẵn sàng (∞)"
            else:
                status = "Sẵn sàng"
            limit = card["limit"] if card["limit"] > 0 else "∞"
            self.vault_tree.insert(
                "",
                "end",
                iid=card["id"],
                text=card["label"],
                values=(limit, card["used"], status),
            )

    def _log_vault_notices_once(self):
        """Ghi log một lần mỗi lần mở app: báo cáo chuyển thẻ từ
        Credential Manager sang file (nếu vừa chuyển) và cảnh báo
        khi file thẻ đang lưu không mã hoá."""
        if getattr(self, "_vault_notices_logged", False):
            return
        self._vault_notices_logged = True
        report = storage._last_migration_report()
        if report and (
            report["migrated"]
            or report["orphans_removed"]
            or report.get("error")
        ):
            parts = []
            if report["migrated"]:
                parts.append(
                    f"đã chuyển {report['migrated']} thẻ từ Credential "
                    "Manager sang file kho thẻ"
                )
            if report["orphans_removed"]:
                parts.append(
                    f"đã dọn {report['orphans_removed']} thẻ mồ côi khỏi "
                    "Credential Manager"
                )
            if not report["enumerated"] and report["cleanup_error"]:
                parts.append(
                    "không liệt kê được Credential Manager để dọn hết "
                    f"thẻ mồ côi ({report['cleanup_error']})"
                )
            if report.get("error"):
                parts.append(
                    "chuyển thẻ gặp lỗi "
                    f"({report['error']}) — thẻ vẫn còn ở chỗ cũ, "
                    "lần sau mở app sẽ thử lại"
                )
            self.app.append_log("Kho thẻ: " + "; ".join(parts) + ".")
        warning = storage._cards_protection_warning()
        if warning:
            self.app.append_log(warning)

    def _vault_selection(self):
        """Thẻ đang chọn trong danh sách kho (dict id/label/limit/
        used), hoặc None khi chưa chọn thẻ nào."""
        selected = self.vault_tree.selection()
        if not selected:
            return None
        card_id = selected[0]
        for card in getattr(self, "_vault_cards", []):
            if card["id"] == card_id:
                return card
        return None

    def _on_vault_select(self, event):
        card = self._vault_selection()
        if card is None:
            return
        self.card_limit_var.set(card["limit"])

    def _import_bulk_cards(self):
        """Nhập nhiều thẻ cùng lúc từ ô dán (mỗi dòng một thẻ).

        Định dạng dòng: số thẻ|MM/YY|CVV (ZIP mặc định 10001,
        nhãn tự đặt; có thể thêm ZIP và nhãn làm trường thứ 4
        và thứ 5). Thẻ hợp lệ ghi thẳng vào file kho thẻ
        (cards.json, mã hoá bằng DPAPI của Windows khi có). Ô dán được
        XOÁ NGAY sau khi đọc — dù thành công hay lỗi — để số
        thẻ không nằm lại trên màn hình; dòng lỗi chỉ báo theo
        số dòng, không hiện nội dung dòng hay chữ số thẻ.
        """
        if self.coordinator.running:
            return
        text = self.bulk_card_text.get("1.0", "end")
        self.bulk_card_text.delete("1.0", "end")
        if not text.strip():
            messagebox.showwarning(
                "Chưa có thẻ để nhập",
                "Dán danh sách thẻ vào ô trên — mỗi dòng một thẻ, "
                "định dạng: số thẻ|MM/YY|CVV.",
            )
            return
        try:
            added, errors, skipped = storage._import_bulk_cards(text)
        except Exception as error:
            # storage đã trả lỗi bằng tiếng Việt, ghi rõ bước
            # nào hỏng — không gắn thêm tiền tố chung chung.
            messagebox.showerror(
                "Không nhập được thẻ",
                str(error),
            )
            return
        self._refresh_vault()
        summary = []
        if added:
            summary.append(
                f"Đã thêm {len(added)} thẻ vào kho: {', '.join(added)}."
            )
        if errors:
            detail = "; ".join(
                f"dòng {line_no}: {reason}" for line_no, reason in errors
            )
            summary.append(f"{len(errors)} dòng lỗi — {detail}.")
        if skipped:
            summary.append(
                f"Đã bỏ qua {skipped} dòng tiêu đề/phân cách "
                "(không phải là thẻ)."
            )
        if not added and not errors and not skipped:
            summary.append("Không có dòng thẻ nào để nhập.")
        message = " ".join(summary)
        if added and not errors:
            self.app.append_log(
                message + " Thẻ đã lưu vào file kho thẻ."
            )
        elif added:
            messagebox.showwarning("Nhập thẻ xong (có dòng lỗi)", message)
            self.app.append_log(message)
        else:
            messagebox.showwarning("Không thêm được thẻ nào", message)
            self.app.append_log(message)

    def _select_vault_card(self):
        """Chọn tay một thẻ: lượt chạy chỉ dùng thẻ này (kể cả
        khi thẻ khác còn giới hạn)."""
        if self.coordinator.running:
            return
        card = self._vault_selection()
        if card is None:
            messagebox.showwarning(
                "Chưa chọn thẻ",
                "Chọn một thẻ trong danh sách kho thẻ trước.",
            )
            return
        storage._save_vault_active(card["id"])
        self._refresh_vault()
        self.app.append_log(
            f"Đã chọn thẻ {card['label']} — lượt chạy chỉ dùng thẻ "
            "này (đủ giới hạn là tự dừng, không tự chuyển thẻ khác)."
        )

    def _auto_rotate_cards(self):
        """Bỏ chọn tay: lượt chạy dùng các thẻ trong kho theo thứ
        tự, thẻ nào đủ giới hạn thì tự chuyển sang thẻ kế."""
        if self.coordinator.running:
            return
        storage._save_vault_active("")
        self._refresh_vault()
        self.app.append_log(
            "Đã chuyển sang tự động xoay — lượt chạy dùng các thẻ "
            "trong kho theo thứ tự, thẻ nào đủ giới hạn thì tự "
            "chuyển sang thẻ kế; hết kho là tự dừng."
        )

    def _delete_vault_card(self):
        """Xoá thẻ khỏi kho (cả file thẻ lẫn danh mục). Hộp xác nhận chỉ nêu nhãn thẻ, không nêu số thẻ."""
        if self.coordinator.running:
            return
        card = self._vault_selection()
        if card is None:
            messagebox.showwarning(
                "Chưa chọn thẻ",
                "Chọn một thẻ trong danh sách kho thẻ trước.",
            )
            return
        if not messagebox.askokcancel(
            "Xoá thẻ",
            f"Xoá thẻ {card['label']} khỏi kho thẻ? Thẻ sẽ bị xoá "
            "khỏi file kho thẻ và không dùng được nữa.",
        ):
            return
        try:
            storage._delete_vault_card(card["id"])
        except Exception as error:
            messagebox.showerror("Không xoá được thẻ", str(error))
            return
        self._refresh_vault()
        self.app.append_log(f"Đã xoá thẻ {card['label']} khỏi kho thẻ.")

    def _save_vault_limit(self):
        """Đặt giới hạn / thẻ cho thẻ đang chọn (0 = không giới hạn)."""
        if self.coordinator.running:
            return
        card = self._vault_selection()
        if card is None:
            messagebox.showwarning(
                "Chưa chọn thẻ",
                "Chọn một thẻ trong danh sách kho thẻ trước.",
            )
            return
        try:
            limit = int(self.card_limit_var.get())
        except (tk.TclError, ValueError):
            limit = 0
        limit = max(0, min(limit, 100))
        if not storage._set_vault_card_limit(card["id"], limit):
            messagebox.showerror(
                "Không lưu được giới hạn", "Thẻ này không còn trong kho."
            )
            self._refresh_vault()
            return
        self._refresh_vault()
        limit_text = f"{limit} tài khoản" if limit > 0 else "không giới hạn"
        self.app.append_log(
            f"Đã đặt giới hạn thẻ {card['label']}: {limit_text}."
        )

    def _apply_vault_limit_to_all(self):
        """Đặt MỘT giới hạn chung cho tất cả thẻ trong kho (lấy số
        ở ô Giới hạn / thẻ): ghi đè giới hạn từng thẻ. Không cần
        chọn thẻ nào trước. Thẻ nào muốn khác thì chọn thẻ đó rồi
        bấm Lưu giới hạn như cũ."""
        if self.coordinator.running:
            return
        try:
            limit = int(self.card_limit_var.get())
        except (tk.TclError, ValueError):
            limit = 0
        limit = max(0, min(limit, 100))
        count = storage._set_all_vault_card_limits(limit)
        if not count:
            messagebox.showwarning(
                "Kho thẻ đang trống",
                "Chưa có thẻ nào trong kho để đặt giới hạn.",
            )
            return
        self._refresh_vault()
        limit_text = f"{limit} tài khoản" if limit > 0 else "không giới hạn"
        self.app.append_log(
            f"Đã đặt giới hạn chung {limit_text} cho tất cả "
            f"{count} thẻ trong kho. Muốn đổi riêng một thẻ thì "
            "chọn thẻ đó rồi bấm Lưu giới hạn."
        )

    def _run_cards(self):
        """Kho thẻ của lượt chạy sắp bắt đầu.

        Trả về (cards, manual): cards là list CardRunState theo
        thứ tự kho (thẻ chọn tay thì chỉ có đúng thẻ đó);
        manual=True là chỉ dùng thẻ chọn tay, False là tự động
        xoay. Kho trống thì ([], False) — bước thẻ chờ nhập tay.
        """
        vault = storage._load_vault()
        active = storage._load_vault_active()
        if active:
            chosen = [card for card in vault if card["id"] == active]
            if not chosen:
                storage._save_vault_active("")
            else:
                return (
                    [
                        CardRunState(
                            card_id=chosen[0]["id"],
                            label=chosen[0]["label"],
                            limit=chosen[0]["limit"],
                            used=chosen[0]["used"],
                        )
                    ],
                    True,
                )
        return (
            [
                CardRunState(
                    card_id=card["id"],
                    label=card["label"],
                    limit=card["limit"],
                    used=card["used"],
                )
                for card in vault
            ],
            False,
        )

    def _check_cards_before_run(self, cards, manual):
        """Chặn chạy khi thẻ (chọn tay) hoặc toàn kho đã đủ giới
        hạn. Trả về True khi được chạy tiếp."""
        if not cards:
            return True
        if manual:
            card = cards[0]
            if card.limit > 0 and card.used >= card.limit:
                messagebox.showwarning(
                    "Thẻ đã đủ giới hạn",
                    f"Thẻ {card.label} đã dùng cho {card.used} tài "
                    f"khoản thành công (giới hạn {card.limit} / thẻ). "
                    "Chọn thẻ khác ở kho thẻ (nút Chọn thẻ này) "
                    "hoặc tăng giới hạn rồi chạy lại.",
                )
                return False
            return True
        if all(
            card.limit > 0 and card.used >= card.limit for card in cards
        ):
            messagebox.showwarning(
                "Tất cả thẻ đã đủ giới hạn",
                "Tất cả thẻ trong kho đã đủ giới hạn. Tăng giới hạn "
                "chung (ô Giới hạn / thẻ rồi bấm Áp dụng cho tất cả "
                "thẻ), tăng riêng từng thẻ, hoặc nhập thêm thẻ mới "
                "vào kho thẻ rồi chạy lại.",
            )
            return False
        return True

    def _log_run_cards(self, cards, manual):
        """Ghi log kho thẻ của lượt chạy (chỉ nhãn + số đếm)."""
        if not cards:
            return
        if manual:
            card = cards[0]
            limit = card.limit if card.limit > 0 else "∞"
            self.app.append_log(
                f"Thẻ đang chọn: {card.label} — {card.used}/{limit} "
                "tài khoản thành công; đủ giới hạn là lượt chạy tự "
                "dừng, không tự chuyển thẻ khác."
            )
        else:
            self.app.append_log(
                f"Kho thẻ: {len(cards)} thẻ — dùng theo thứ tự, thẻ "
                "nào đủ giới hạn thì tự chuyển sang thẻ kế; hết kho "
                "là lượt chạy tự dừng."
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
        # Kho thẻ: thẻ chọn tay thì chỉ dùng thẻ đó, còn không thì
        # tự động xoay theo thứ tự kho; kho trống thì bước thẻ chờ
        # nhập tay, không tính giới hạn.
        cards, manual_card = self._run_cards()
        if not self._check_cards_before_run(cards, manual_card):
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
            cards=cards,
            manual_card=manual_card,
            bad_file=self.panel.bad_file,
            bad_mapping=setup["bad_mapping"],
            ip_target=self.panel,
        )
        try:
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
            self._log_run_cards(cards, manual_card)
            if not self.coordinator.start(settings):
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
                "⛔ Không bắt đầu được lượt chạy: "
                f"{type(error).__name__}: {error}"
            )

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
        # Đăng nhập vẫn nộp thẻ trong kho ở bước thẻ — kho trống
        # thì mỗi tài khoản sẽ chờ nhập tay tới 4 phút.
        cards, manual_card = self._run_cards()
        if not cards:
            if not messagebox.askokcancel(
                "Chưa có thẻ trong kho",
                "Chưa có thẻ nào trong kho. Mỗi tài khoản đăng nhập "
                "xong tới bước thẻ sẽ chờ bạn nhập tay trong cửa sổ "
                "Chromium (tối đa 4 phút). Vẫn chạy?",
            ):
                return
        if not self._check_cards_before_run(cards, manual_card):
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
        settings = RunSettings(
            total=setup["total"],
            pool=setup["pool"],
            per_proxy=setup["per_proxy"],
            lanes=setup["lanes"],
            provider_id=setup["provider_id"],
            direct=setup["direct"],
            meta_login=True,
            meta_excel_path=excel_text,
            invite=invite,
            single_invite=single_invite,
            cards=cards,
            manual_card=manual_card,
            bad_file=self.panel.bad_file,
            bad_mapping=setup["bad_mapping"],
            ip_target=self.panel,
        )
        try:
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
            self._log_run_cards(cards, manual_card)
            if invite:
                if single_invite:
                    self.app.append_log(
                        f"Mã mời: mọi tài khoản đều nhập mã {invite}."
                    )
                else:
                    self.app.append_log(
                        f"Mã mời: tài khoản đầu nhập mã {invite}, "
                        "tài khoản sau nhập mã của tài khoản trước."
                    )
            else:
                self.app.append_log(
                    "Chưa có mã mời đầu chuỗi — tài khoản đầu không "
                    "nhập mã, tài khoản sau nhập mã của tài khoản trước."
                )
            if not self.coordinator.start(settings):
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
                "⛔ Không bắt đầu được lượt ĐĂNG NHẬP MUSE: "
                f"{type(error).__name__}: {error}"
            )

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
        # ttk.Treeview KHÔNG có tuỳ chọn -state: configure(state=...)
        # raise TclError "unknown option -state" — chính lỗi này làm
        # chết mọi nút Bắt đầu (xem main_window.on_run_started):
        # handler vỡ giữa chừng, tkinter nuốt exception, nút vẫn
        # sáng mà không chạy gì. Treeview khoá/mở bằng .state().
        self.vault_tree.state(["disabled"] if running else ["!disabled"])
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
            self.bulk_card_text,
            self.import_bulk_button,
            self.select_card_button,
            self.auto_rotate_button,
            self.delete_vault_card_button,
            self.save_limit_button,
            self.apply_limit_all_button,
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
