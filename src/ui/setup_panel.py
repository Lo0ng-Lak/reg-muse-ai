"""Khung setup chạy dùng chung cho cả hai tab.

Mỗi tab có một SetupPanel riêng với kho proxy + kho lỗi riêng:
tab Muse dùng proxies.txt / bad-proxies.txt (file cũ, giữ nguyên
để kho hiện tại không phải dán lại), tab Meta dùng
proxies_meta.txt / bad-proxies-meta.txt. Lượt chạy bấm từ tab
nào thì dùng đúng setup của tab đó.
"""

import threading
import tkinter as tk
from tkinter import ttk

import mail_providers
import paths
import proxy_pool
from coordinator import CONN_MODES, DIRECT_PROXY_LINE


class SetupError(Exception):
    """Setup chưa hợp lệ — tab bắt lỗi này để hiện cảnh báo."""

    def __init__(self, title, message):
        super().__init__(message)
        self.title = title
        self.message = message


class SetupPanel(ttk.Frame):
    """Proxy + luồng + kết nối + temp mail + số lượng cho một tab.

    show_provider=False ẩn hàng Temp mail (tab Muse: mail lấy từ
    file Excel Meta, không tạo mail mới nên không cần chọn dịch
    vụ); read_setup() khi đó trả provider mặc định cho lượt đăng
    ký trực tiếp cũ.
    """

    def __init__(
        self, parent, coordinator, pool_file, bad_file,
        pool_fallback_file=None, show_provider=True,
    ):
        super().__init__(parent)
        self.coordinator = coordinator
        self.pool_file = pool_file
        self.bad_file = bad_file
        self.pool_fallback_file = pool_fallback_file
        self.show_provider = show_provider
        self.provider_var = None
        self.provider_combo = None
        self._running = False
        self._bad_proxies = proxy_pool._load_bad_proxies(bad_file)
        self.columnconfigure(0, weight=1)
        self._build()
        self._load_initial()

    def _build(self):
        proxy_row = ttk.Frame(self)
        proxy_row.grid(row=0, column=0, sticky="ew", pady=(0, 10))
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

        bad_row = ttk.Frame(self)
        bad_row.grid(row=1, column=0, sticky="ew", pady=(0, 10))
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

        controls = ttk.Frame(self)
        controls.grid(row=2, column=0, sticky="ew")
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
        ttk.Label(controls, text="Số luồng:").pack(side="left")
        self.lanes_var = tk.IntVar(value=1)
        self.lanes_spin = ttk.Spinbox(
            controls, from_=1, to=5, width=5, textvariable=self.lanes_var
        )
        self.lanes_spin.pack(side="left", padx=(6, 12))

        if self.show_provider:
            mail_row = ttk.Frame(self)
            mail_row.grid(row=3, column=0, sticky="ew", pady=(10, 0))
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

        conn_row = ttk.Frame(self)
        conn_row.grid(row=4, column=0, sticky="ew", pady=(10, 0))
        ttk.Label(conn_row, text="Kết nối:").pack(side="left")
        self.conn_mode_var = tk.StringVar(value=CONN_MODES[0][1])
        self.conn_combo = ttk.Combobox(
            conn_row,
            textvariable=self.conn_mode_var,
            values=[label for _, label in CONN_MODES],
            state="readonly",
            width=14,
        )
        self.conn_combo.pack(side="left", padx=(6, 8))
        self.conn_combo.bind(
            "<<ComboboxSelected>>", self._on_conn_mode_change
        )
        ttk.Label(
            conn_row,
            text=(
                "IP máy (VPN): chạy trực tiếp qua IP của máy này — "
                "bật VPN trước khi bấm Bắt đầu; mọi luồng dùng "
                "chung 1 IP, IP bị chặn là dừng cả lượt chạy."
            ),
            wraplength=480,
        ).pack(side="left")

        ip_row = ttk.Frame(self)
        ip_row.grid(row=5, column=0, sticky="ew", pady=(6, 0))
        ttk.Label(ip_row, text="IP hiện tại:").pack(side="left")
        self.ip_status = tk.StringVar(value="chưa kiểm tra")
        ttk.Label(ip_row, textvariable=self.ip_status).pack(
            side="left", padx=(6, 8)
        )
        self.check_ip_button = ttk.Button(
            ip_row, text="Kiểm tra IP", command=self._on_check_ip
        )
        self.check_ip_button.pack(side="left")

    def _load_initial(self):
        saved_pool = ""
        for pool_name in (self.pool_file, self.pool_fallback_file):
            if pool_name is None:
                continue
            try:
                saved_pool = pool_name.read_text(
                    encoding="utf-8"
                ).strip()
            except OSError:
                saved_pool = ""
            if saved_pool:
                break
        if saved_pool:
            self.proxy_text.insert("1.0", saved_pool)
        if self._bad_proxies:
            self.bad_proxy_text.insert(
                "1.0", "\n".join(self._bad_proxies.values())
            )

    def _on_conn_mode_change(self, _event=None):
        """Chế độ IP máy không có kho proxy để xoay nên ô
        'Tài khoản / proxy' vô nghĩa — khoá lại cho rõ."""
        direct = self.conn_mode_var.get() == CONN_MODES[1][1]
        self.per_proxy_spin.configure(
            state="disabled" if direct else "normal"
        )

    def _on_check_ip(self):
        """Nút Kiểm tra IP: tra IP trên luồng riêng để không
        đơ giao diện; kết quả về qua hàng đợi sự kiện. Nút
        này không bị khoá khi đang chạy — kiểm tra giữa lượt
        chạy vẫn được (ví dụ nghi VPN rớt)."""
        self.check_ip_button.configure(state="disabled")
        self.ip_status.set("đang kiểm tra…")

        def _work():
            try:
                self.coordinator.report_ip(
                    proxy_pool._check_public_ip(), target=self
                )
            finally:
                self.coordinator.events.put(("ip_done", self))

        threading.Thread(target=_work, daemon=True).start()

    def set_ip_status(self, status):
        self.ip_status.set(status)

    def set_ip_done(self):
        self.check_ip_button.configure(state="normal")

    def _sync_pool(self):
        """Đồng bộ ô Proxy (nhiều dòng) thành kho proxy cho lượt chạy.

        Mỗi dòng là một proxy; dòng trống bị bỏ qua. Kho được lưu
        vào file kho của tab để lần mở sau tự điền lại, còn
        proxy.txt chỉ chứa proxy đang dùng cho nhóm tài khoản hiện
        tại (browser_helper đọc file đó). Ô trống → xoá cả hai
        file và trả về None. Log chỉ dùng _proxy_display — không
        bao giờ lộ user/password của proxy.
        """
        raw = self.proxy_text.get("1.0", "end")
        pool = [line.strip() for line in raw.splitlines() if line.strip()]
        active_file = paths.data_file("proxy.txt")
        if not pool:
            for path in (self.pool_file, active_file):
                try:
                    path.unlink()
                except OSError:
                    pass
            return None
        try:
            self.pool_file.write_text(
                "\n".join(pool) + "\n", encoding="utf-8"
            )
            active_file.write_text(pool[0] + "\n", encoding="utf-8")
        except OSError as error:
            raise RuntimeError(f"Không ghi được file proxy: {error}") from error
        return pool

    def _sync_bad(self):
        """Đồng bộ ô Proxy lỗi thành kho loại trừ cho lượt chạy.

        Nội dung ô là nguồn sự thật: dòng nào còn trong ô thì
        proxy đó bị loại; xoá dòng khỏi ô là cho dùng lại. Kho
        được lưu vào file kho lỗi của tab để giữ qua các lần mở
        app. Trả về (mapping, tập khoá bị loại).
        """
        raw = self.bad_proxy_text.get("1.0", "end")
        mapping = {}
        for line in raw.splitlines():
            line = line.strip()
            if line:
                mapping[proxy_pool._proxy_key(line)] = line
        self._bad_proxies = mapping
        proxy_pool._save_bad_proxies(mapping, self.bad_file)
        return mapping, set(mapping)

    def mark_bad(self, line):
        """Thêm một proxy vào ô Proxy lỗi (gọi từ luồng giao diện
        khi coordinator báo proxy bị chặn trong lượt chạy)."""
        existing = {
            proxy_pool._proxy_key(text)
            for text in self.bad_proxy_text.get("1.0", "end").splitlines()
            if text.strip()
        }
        if proxy_pool._proxy_key(line) in existing:
            return
        if self._running:
            self.bad_proxy_text.configure(state="normal")
        current = self.bad_proxy_text.get("1.0", "end").strip()
        self.bad_proxy_text.delete("1.0", "end")
        self.bad_proxy_text.insert(
            "1.0", (current + "\n" if current else "") + line
        )
        if self._running:
            self.bad_proxy_text.configure(state="disabled")

    def read_setup(self):
        """Đọc + kiểm tra setup của tab cho một lượt chạy.

        Trả về dict: total, per_proxy, lanes, provider_id,
        provider_label, direct, pool, bad_mapping. Lỗi setup
        thì raise SetupError (tab hiện cảnh báo, không chạy).
        """
        try:
            total = int(self.count_var.get())
        except (tk.TclError, ValueError):
            total = 0
        if total < 0 or total > 100:
            raise SetupError(
                "Số tài khoản chưa hợp lệ",
                "Nhập số tài khoản từ 0 đến 100. "
                "Để 0 là chạy liên tục tới khi dùng hết lượt của "
                "toàn bộ proxy trong danh sách thì dừng.",
            )
        try:
            per_proxy = int(self.per_proxy_var.get())
        except (tk.TclError, ValueError):
            per_proxy = 5
        per_proxy = max(1, min(per_proxy, 100))
        try:
            lanes = int(self.lanes_var.get())
        except (tk.TclError, ValueError):
            lanes = 1
        lanes = max(1, min(lanes, 5))
        conn_label = self.conn_mode_var.get()
        conn_id = next(
            (cid for cid, label in CONN_MODES if label == conn_label),
            "proxy",
        )
        direct = conn_id == "direct"
        bad_mapping = {}
        if direct:
            # Chế độ IP máy: không dùng kho proxy — pool chỉ là
            # câu đánh dấu DIRECT để helper chạy thẳng qua IP của
            # máy (người dùng tự bật VPN trước). Hạn mức "tài
            # khoản / proxy" vô nghĩa nên đặt rất lớn để vòng xoay
            # không bao giờ quay vòng — quay vòng sẽ khiến chế độ
            # không giới hạn dừng sớm sau vài tài khoản.
            pool = [DIRECT_PROXY_LINE]
            per_proxy = 10**9
        else:
            try:
                pool = self._sync_pool()
            except RuntimeError as error:
                raise SetupError("Lỗi proxy", str(error)) from error
            if not pool:
                raise SetupError(
                    "Thiếu proxy",
                    "Chế độ Proxy chỉ đăng ký qua proxy. "
                    "Hãy nhập ít nhất một proxy (mỗi dòng một proxy, "
                    "dạng host:port:user:pass) vào ô Proxy rồi bấm lại, "
                    "hoặc chọn chế độ IP máy (VPN) ở ô Kết nối.",
                )
            bad_mapping, bad_keys = self._sync_bad()
            if bad_keys:
                excluded = sum(
                    1
                    for line in pool
                    if proxy_pool._proxy_key(line) in bad_keys
                )
                if excluded:
                    self.coordinator.events.put(
                        (
                            "log",
                            f"Đã loại {excluded} proxy trong ô Proxy lỗi "
                            "khỏi kho cho đợt chạy này.",
                        )
                    )
            pool = proxy_pool._prepare_pool(pool, bad_keys)
            if bad_keys and not pool:
                raise SetupError(
                    "Toàn bộ proxy đều nằm trong kho lỗi",
                    "Mọi proxy trong ô Proxy đều đã bị đánh dấu "
                    "lỗi. Xoá bớt dòng trong ô Proxy lỗi để dùng "
                    "lại, hoặc dán proxy mới vào ô Proxy.",
                )
        if self.provider_var is None:
            # Tab không có hàng Temp mail (tab Muse): mail lấy từ
            # file Excel Meta; lượt đăng ký trực tiếp cũ (nếu còn
            # dùng) chạy với dịch vụ mặc định.
            provider_id = mail_providers.DEFAULT_PROVIDER
            provider_label = dict(mail_providers.PROVIDERS)[provider_id]
        else:
            provider_label = self.provider_var.get()
            provider_id = next(
                (
                    pid
                    for pid, label in mail_providers.PROVIDERS
                    if label == provider_label
                ),
                mail_providers.DEFAULT_PROVIDER,
            )
        return {
            "total": total,
            "per_proxy": per_proxy,
            "lanes": lanes,
            "provider_id": provider_id,
            "provider_label": provider_label,
            "direct": direct,
            "pool": pool,
            "bad_mapping": bad_mapping,
        }

    def set_running(self, running):
        self._running = running
        self.proxy_text.configure(state="disabled" if running else "normal")
        self.bad_proxy_text.configure(
            state="disabled" if running else "normal"
        )
        direct = self.conn_mode_var.get() == CONN_MODES[1][1]
        self.per_proxy_spin.configure(
            state="disabled" if (running or direct) else "normal"
        )
        self.count_spin.configure(state="disabled" if running else "normal")
        self.lanes_spin.configure(state="disabled" if running else "normal")
        if self.provider_combo is not None:
            self.provider_combo.configure(
                state="disabled" if running else "readonly"
            )
        self.conn_combo.configure(
            state="disabled" if running else "readonly"
        )
