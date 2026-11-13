"""Cửa sổ chính: hai tab (Muse.ai / Meta) + log và tiến độ chung.

Một lượt chạy tại một thời điểm — bấm Bắt đầu ở tab nào thì chạy
đúng setup của tab đó; tab còn lại bị khoá cho tới khi lượt chạy
kết thúc. Log, tiến độ và nút Dừng dùng chung cho cả hai tab.
"""

import queue
import tkinter as tk
from datetime import datetime
from tkinter import ttk

import storage
from coordinator import Coordinator
from ui.meta_tab import MetaTab
from ui.muse_tab import MuseTab


class MainWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Muse.ai Account Helper")
        self.geometry("820x720")
        self.minsize(680, 600)
        self.coordinator = Coordinator()
        self._lane_progress = {}
        self.active_panel = None
        self._build_ui()
        self.after(100, self._process_events)
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.append_log(
            "Sẵn sàng. Chọn tab Muse.ai hoặc Meta, nhập setup rồi bấm "
            "Bắt đầu. Chỉ tài khoản đăng ký thành công mới được lưu "
            "vào sheet Excel."
        )
        if storage._load_card_from_keyring():
            self.append_log(
                "Đã có thẻ lưu sẵn — tool sẽ tự điền ở bước xác minh thẻ. "
                f"Thẻ này đã dùng cho {storage._load_card_success_count()} tài khoản "
                "thành công (xem / đổi ở ô 'Giới hạn / thẻ:')."
            )

    def _build_ui(self):
        root = ttk.Frame(self, padding=16)
        root.pack(fill="both", expand=True)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(1, weight=1)

        notebook = ttk.Notebook(root)
        notebook.grid(row=0, column=0, sticky="ew")
        self.muse_tab = MuseTab(notebook, self)
        self.meta_tab = MetaTab(notebook, self)
        notebook.add(self.muse_tab, text="Muse.ai")
        notebook.add(self.meta_tab, text="Meta (meta.ai)")

        log_frame = ttk.LabelFrame(root, text="Quá trình", padding=8)
        log_frame.grid(row=1, column=0, sticky="nsew", pady=(12, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log_text = tk.Text(log_frame, wrap="word", state="disabled")
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(
            log_frame, orient="vertical", command=self.log_text.yview
        )
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)

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
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"[{stamp}] {text}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def on_run_started(self, panel):
        """Tab vừa bấm Bắt đầu: ghi nhận panel đang chạy (để sự
        kiện bad_proxy / ip về đúng tab) và khoá cả hai tab."""
        self.active_panel = panel
        self._lane_progress = {}
        self.set_running(True)

    def set_running(self, running):
        self.muse_tab.set_running(running)
        self.meta_tab.set_running(running)
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

    def _process_events(self):
        try:
            while True:
                kind, payload = self.coordinator.events.get_nowait()
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
        except queue.Empty:
            pass
        self.after(100, self._process_events)


def main():
    MainWindow().mainloop()


if __name__ == "__main__":
    main()
