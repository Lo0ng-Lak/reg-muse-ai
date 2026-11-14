"""Cửa sổ chính: hai tab (Muse.ai / Meta) + log và tiến độ chung.

Một lượt chạy tại một thời điểm — bấm Bắt đầu ở tab nào thì chạy
đúng setup của tab đó; tab còn lại bị khoá cho tới khi lượt chạy
kết thúc. Log, tiến độ và nút Dừng dùng chung cho cả hai tab.
"""

import queue
import tkinter as tk
import traceback
from datetime import datetime
from tkinter import ttk

import storage
from coordinator import Coordinator
from ui.meta_tab import MetaTab
from ui.muse_tab import MuseTab
from ui.relogin_tab import ReloginTab


class _ScrollableFrame(ttk.Frame):
    """Khung cuộn dọc cho nội dung tab dài hơn cửa sổ.

    Tab Muse (nhất là khung Kho thẻ) xin chiều cao ~1040px trong
    khi cửa sổ chỉ 980px — trước đây Notebook phình theo tab cao
    nhất (~1070px), chiếm hết chỗ và ép khung log xuống còn ~1px.
    Đặt mỗi tab vào khung cuộn có chiều cao cố định thì Notebook
    không phình nữa, nội dung dài cuộn bằng thanh cuộn / lăn chuột,
    khung log luôn đủ chỗ.
    """

    def __init__(self, parent, height=600):
        super().__init__(parent)
        self.canvas = tk.Canvas(self, height=height, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(
            self, orient="vertical", command=self.canvas.yview
        )
        self.inner = ttk.Frame(self.canvas)
        self._window = self.canvas.create_window(
            (0, 0), window=self.inner, anchor="nw"
        )
        self.inner.bind(
            "<Configure>",
            lambda _e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            ),
        )
        self.canvas.bind(
            "<Configure>",
            lambda e: self.canvas.itemconfigure(self._window, width=e.width),
        )
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        self.bind("<Enter>", self._bind_wheel)
        self.bind("<Leave>", self._unbind_wheel)

    def _on_mousewheel(self, event):
        if event.num == 4:
            delta = -1
        elif event.num == 5:
            delta = 1
        else:
            delta = -int(event.delta / 120)
        self.canvas.yview_scroll(delta, "units")

    def _bind_wheel(self, _event=None):
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind_all("<Button-4>", self._on_mousewheel)
        self.canvas.bind_all("<Button-5>", self._on_mousewheel)

    def _unbind_wheel(self, _event=None):
        self.canvas.unbind_all("<MouseWheel>")
        self.canvas.unbind_all("<Button-4>")
        self.canvas.unbind_all("<Button-5>")


class MainWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Muse.ai Account Helper")
        self.geometry("820x980")
        self.minsize(680, 720)
        self.coordinator = Coordinator()
        self._lane_progress = {}
        self.active_panel = None
        # Log gọi trước khi khung log dựng xong thì giữ lại rồi xả
        # ra sau (xem append_log) — MuseTab ghi log ngay lúc dựng
        # giao diện nên danh sách này phải có trước _build_ui().
        self._pending_logs = []
        self._build_ui()
        self.after(100, self._process_events)
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.append_log(
            "Sẵn sàng. Chọn tab Muse.ai, Meta, hoặc Đăng nhập lại, "
            "nhập setup rồi bấm Bắt đầu. Chỉ tài khoản đăng ký thành "
            "công mới được lưu vào sheet Excel."
        )
        vault = storage._load_vault()
        if vault:
            active = storage._load_vault_active()
            chosen = next(
                (card for card in vault if card["id"] == active), None
            )
            if chosen is not None:
                limit = chosen["limit"] if chosen["limit"] > 0 else "∞"
                self.append_log(
                    f"Kho thẻ: {len(vault)} thẻ — đang chọn thẻ "
                    f"{chosen['label']} ({chosen['used']}/{limit} tài "
                    "khoản thành công). Lượt chạy chỉ dùng thẻ này."
                )
            else:
                self.append_log(
                    f"Kho thẻ: {len(vault)} thẻ — tự động xoay theo thứ "
                    "tự, thẻ nào đủ giới hạn thì chuyển sang thẻ kế."
                )

    def _build_ui(self):
        root = ttk.Frame(self, padding=16)
        root.pack(fill="both", expand=True)
        root.columnconfigure(0, weight=1)
        # Hàng log luôn giữ ít nhất 180px: trước đây hàng này là hàng
        # duy nhất có weight nên hứng toàn bộ phần thiếu khi nội dung
        # xin cao hơn cửa sổ, khung log bị ép còn ~1px.
        root.rowconfigure(1, weight=1, minsize=180)

        # Khung log phải tạo TRƯỚC các tab: MuseTab ghi log ngay
        # lúc dựng giao diện (báo cáo chuyển thẻ sang file +
        # cảnh báo file không mã hoá, xem _log_vault_notices_once)
        # — trước bản 4f3814f8 chưa có log nào lúc dựng tab nên
        # log_text tạo sau vẫn chạy; từ khi có log lúc dựng tab,
        # append_log trong lúc dựng tab vỡ AttributeError
        # (log_text chưa có) và app chết trước khi kịp hiện cửa
        # sổ (gặp live 2026-11-13: chạy py main.py không lên
        # luôn). Tạo log trước, tab sau thì thứ tự không còn
        # quan trọng nữa. Vị trí trên lưới giữ nguyên (log ở
        # hàng 1, tab ở hàng 0) nên bố cục không đổi.
        log_frame = ttk.LabelFrame(root, text="Quá trình", padding=8)
        log_frame.grid(row=1, column=0, sticky="nsew", pady=(12, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log_text = tk.Text(
            log_frame, wrap="word", state="disabled", height=12
        )
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(
            log_frame, orient="vertical", command=self.log_text.yview
        )
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)
        for line in getattr(self, "_pending_logs", []):
            self.log_text.configure(state="normal")
            self.log_text.insert("end", line)
            self.log_text.configure(state="disabled")
        self._pending_logs = []

        notebook = ttk.Notebook(root)
        notebook.grid(row=0, column=0, sticky="ew")
        self.notebook = notebook
        muse_scroll = _ScrollableFrame(notebook)
        self.muse_tab = MuseTab(muse_scroll.inner, self)
        # Tab phải pack vào inner thì inner mới "thấy" chiều cao của
        # tab (không pack là tab chỉ cao 1px, nội dung không hiện).
        self.muse_tab.pack(fill="both", expand=True)
        meta_scroll = _ScrollableFrame(notebook)
        self.meta_tab = MetaTab(meta_scroll.inner, self)
        self.meta_tab.pack(fill="both", expand=True)
        notebook.add(muse_scroll, text="Muse.ai")
        notebook.add(meta_scroll, text="Meta (meta.ai)")
        relogin_scroll = _ScrollableFrame(notebook)
        self.relogin_tab = ReloginTab(relogin_scroll.inner, self)
        self.relogin_tab.pack(fill="both", expand=True)
        notebook.add(relogin_scroll, text="Đăng nhập lại")

        bottom = ttk.Frame(root)
        bottom.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        self.stop_button = ttk.Button(
            bottom,
            text="Dừng",
            command=self._stop_registration,
            state="disabled",
        )
        self.stop_button.pack(side="left")
        self.progress = tk.StringVar(value="Sẵn sàng.")
        ttk.Label(bottom, textvariable=self.progress).pack(
            side="left", padx=(12, 0)
        )

    def append_log(self, text):
        stamp = datetime.now().strftime("%H:%M:%S")
        line = f"[{stamp}] {text}\n"
        log_text = getattr(self, "log_text", None)
        if log_text is None:
            # Log gọi trước khi khung log được dựng (lúc dựng
            # tab): giữ lại rồi xả ra khi khung log có — không
            # được vỡ app vì một dòng log.
            pending = getattr(self, "_pending_logs", None)
            if pending is None:
                self._pending_logs = pending = []
            pending.append(line)
            return
        log_text.configure(state="normal")
        log_text.insert("end", line)
        log_text.see("end")
        log_text.configure(state="disabled")

    def on_run_started(self, panel):
        """Tab vừa bấm Bắt đầu: ghi nhận panel đang chạy (để sự
        kiện bad_proxy / ip về đúng tab) và khoá cả hai tab."""
        self.active_panel = panel
        self._lane_progress = {}
        self.set_running(True)

    def set_running(self, running):
        self.muse_tab.set_running(running)
        self.meta_tab.set_running(running)
        self.relogin_tab.set_running(running)
        self.stop_button.configure(state="normal" if running else "disabled")

    def _stop_registration(self):
        if not self.coordinator.running:
            return
        self.coordinator.stop()
        self.progress.set("Đang dừng...")

    def _on_close(self):
        if self.coordinator.running:
            self.coordinator.stop()
            self.coordinator.terminate_helpers()
        self.destroy()

    def report_callback_exception(self, exc_type, exc_value, exc_tb):
        # Tkinter mặc định chỉ in traceback ra stderr — người dùng
        # chạy bằng double-click không bao giờ thấy, nút bấm chết
        # lặng mà không rõ vì sao. Ghi cả vào khung log để lỗi hiện
        # ra ngay trong app, rồi vẫn in ra stderr như mặc định.
        try:
            text = "".join(
                traceback.format_exception(exc_type, exc_value, exc_tb)
            ).rstrip()
            self.append_log(f"⛔ Lỗi giao diện:\n{text}")
        except Exception:
            pass
        super().report_callback_exception(exc_type, exc_value, exc_tb)

    def _process_events(self):
        # Bơm sự kiện từ coordinator. Mỗi sự kiện xử lý riêng: lỗi
        # ở một sự kiện không được giết cả bơm (trước đây một
        # exception trong vòng lặp làm mất luôn lần after() hẹn lại,
        # log/tiến độ/trạng thái không bao giờ cập nhật nữa).
        try:
            while True:
                try:
                    kind, payload = self.coordinator.events.get_nowait()
                except queue.Empty:
                    break
                try:
                    if kind == "log":
                        self.append_log(payload)
                    elif kind == "ip":
                        panel, status = payload
                        target = panel or self.active_panel
                        if target is not None:
                            target.set_ip_status(status)
                    elif kind == "ip_done":
                        if payload is not None:
                            payload.set_ip_done()
                    elif kind == "bad_proxy":
                        if self.active_panel is not None:
                            self.active_panel.mark_bad(payload)
                    elif kind == "invite":
                        # Ô Mã mời luôn phản ánh đầu chuỗi hiện tại: mã
                        # của tài khoản thành công gần nhất.
                        self.muse_tab.set_invite(payload)
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
                        self.active_panel = None
                        self.set_running(False)
                        note = "Đã dừng theo yêu cầu. " if stopped else ""
                        self.progress.set(
                            f"{note}Hoàn tất: {success} thành công, {fail} thất bại."
                        )
                        self.append_log(
                            f"--- Kết thúc lượt chạy: {success} thành công, "
                            f"{fail} thất bại ---"
                        )
                except Exception as error:
                    try:
                        self.append_log(
                            f"⛔ Lỗi xử lý sự kiện '{kind}': "
                            f"{type(error).__name__}: {error}"
                        )
                    except Exception:
                        pass
        finally:
            # Luôn hẹn lại vòng bơm, kể cả khi có sự kiện lỗi —
            # bơm không bao giờ được chết lặng.
            try:
                self.after(100, self._process_events)
            except tk.TclError:
                pass  # cửa sổ đã đóng


def main():
    MainWindow().mainloop()


if __name__ == "__main__":
    main()
