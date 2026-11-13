"""Hồi quy cho sự cố 2026-11-13: nút Đăng kí Meta bấm không chạy,
khung log bị ép còn ~1px, luồng chạy chết lặng, bơm sự kiện chết.

- Test coordinator (lane chết, login chết) chạy mọi nơi, không cần màn hình.
- Test UI cần màn hình (chạy dưới xvfb-run); không có màn hình thì
  tự bỏ qua (SKIP) chứ không fail.
- Dùng keyring giả + thư mục dữ liệu tạm; không chạm file thật, không mạng.
"""
import json
import queue
import sys
import tempfile
import time
import types
from pathlib import Path

SRC = Path(__file__).resolve().parent / "src"
sys.path.insert(0, str(SRC))

_store = {}
_fake_keyring = types.ModuleType("keyring")
_fake_keyring.get_password = lambda s, u: _store.get((s, u))
_fake_keyring.set_password = lambda s, u, p: _store.__setitem__((s, u), p)
_fake_keyring.delete_password = lambda s, u: _store.pop((s, u), None)
sys.modules["keyring"] = _fake_keyring

import paths  # noqa: E402

TMP = Path(tempfile.mkdtemp(prefix="reg-test-ui-"))
paths.data_file = lambda name: TMP / name

import storage  # noqa: E402
from coordinator import Coordinator, RunSettings  # noqa: E402

try:
    import tkinter as tk  # noqa: E402

    _probe = tk.Tk()
    _probe.withdraw()
    _probe.destroy()
    HAVE_DISPLAY = True
except Exception:
    HAVE_DISPLAY = False

RESULTS = []


def _run(name, fn, needs_display=False):
    if needs_display and not HAVE_DISPLAY:
        RESULTS.append((name, None, "no display, skipped"))
        print(f"[SKIP] {name} (no display)")
        return
    try:
        fn()
    except Exception as error:  # noqa: BLE001
        RESULTS.append((name, False, f"{type(error).__name__}: {error}"))
        print(f"[FAIL] {name}: {type(error).__name__}: {error}")
    else:
        RESULTS.append((name, True, ""))
        print(f"[PASS] {name}")


def _drain_events(coord):
    kinds = []
    while True:
        try:
            kinds.append(coord.events.get_nowait())
        except queue.Empty:
            return kinds


def _wait_for(predicate, timeout=5.0):
    deadline = time.time() + timeout
    while time.time() < deadline:
        if predicate():
            return True
        time.sleep(0.05)
    return predicate()


def test_lane_crash_resets_running():
    """Luồng đăng ký chết giữa chừng: running phải về False và
    "finished" phải được phát (trước đây kẹt True, UI khoá cứng)."""
    coord = Coordinator()

    def _boom(*args, **kwargs):
        raise RuntimeError("boom lane")

    coord._run_lane = _boom
    settings = RunSettings(
        total=1, pool=["127.0.0.1:8080:u:p"], per_proxy=5,
        lanes=1, meta=True, bad_mapping={}, bad_file=TMP / "bad.txt",
    )
    assert coord.start(settings) is True
    assert _wait_for(lambda: coord.running is False), "running kẹt True"
    events = _drain_events(coord)
    kinds = [k for k, _ in events]
    assert "finished" in kinds, f"thiếu 'finished': {kinds}"
    logs = [p for k, p in events if k == "log"]
    assert any("dừng đột ngột" in p for p in logs), f"thiếu log lỗi: {logs}"


def test_login_loop_crash_resets_running():
    """Vòng đăng nhập chết ngay từ đầu (đọc Excel lỗi): running phải
    về False và "finished" phải được phát."""
    coord = Coordinator()
    orig = storage._pending_meta_logins
    storage._pending_meta_logins = lambda *a, **k: (_ for _ in ()).throw(
        RuntimeError("boom excel")
    )
    try:
        settings = RunSettings(
            total=1, pool=["127.0.0.1:8080:u:p"], per_proxy=5,
            lanes=1, meta_login=True, meta_excel_path="x.xlsx",
            bad_mapping={}, bad_file=TMP / "bad.txt",
        )
        assert coord.start(settings) is True
        assert _wait_for(lambda: coord.running is False), "running kẹt True"
        events = _drain_events(coord)
        kinds = [k for k, _ in events]
        assert "finished" in kinds, f"thiếu 'finished': {kinds}"
    finally:
        storage._pending_meta_logins = orig


def _make_app():
    from ui.main_window import MainWindow

    app = MainWindow()
    app.update_idletasks()
    app.update()
    return app


def test_set_running_no_raise():
    """set_running(True/False) không được raise — trước đây
    vault_tree.configure(state=...) raise TclError 'unknown option
    -state' (ttk.Treeview không có tuỳ chọn -state), làm chết mọi
    nút Bắt đầu ngay trong on_run_started."""
    app = _make_app()
    try:
        app.set_running(True)
        assert "disabled" in app.muse_tab.vault_tree.state()
        assert str(app.meta_tab.meta_start_button["state"]) == "disabled"
        app.set_running(False)
        assert "disabled" not in app.muse_tab.vault_tree.state()
        assert str(app.meta_tab.meta_start_button["state"]) == "normal"
    finally:
        app.destroy()


def test_meta_start_handler():
    """Bấm Đăng kí Meta (setup đủ) phải gọi coordinator.start và ghi
    log 'Bắt đầu lượt ĐĂNG KÍ META'."""
    app = _make_app()
    try:
        selfie = TMP / "selfie.png"
        selfie.write_bytes(b"\x89PNG\r\n\x1a\n")
        meta = app.meta_tab
        meta.capsolver_entry.delete(0, "end")
        meta.capsolver_entry.insert(0, "fake-key")
        meta.selfie_entry.delete(0, "end")
        meta.selfie_entry.insert(0, str(selfie))
        meta.panel.proxy_text.delete("1.0", "end")
        meta.panel.proxy_text.insert(
            "1.0", "127.0.0.1:8080:u:p\n127.0.0.1:8081:u:p\n"
        )
        started = {}
        app.coordinator.start = lambda s: started.__setitem__("ok", True) or True
        meta._start_meta_registration()
        assert started.get("ok") is True, "coordinator.start không được gọi"
        content = app.log_text.get("1.0", "end")
        assert "Bắt đầu lượt ĐĂNG KÍ META" in content, "thiếu log bắt đầu"
    finally:
        app.destroy()


def test_pump_survives_bad_event():
    """Một sự kiện lỗi không được giết bơm: sự kiện tốt sau đó vẫn
    phải được xử lý mà KHÔNG cần gọi _process_events bằng tay
    (tức là after(100) hẹn lại vẫn chạy)."""
    app = _make_app()
    try:
        # Chỉ đưa sự kiện lỗi vào trước: _process_events phải sống
        # sót và hẹn lại. Sự kiện tốt đưa vào SAU, và chỉ chờ bằng
        # vòng update() (không gọi _process_events bằng tay nữa) để
        # chứng minh after(100) hẹn lại thật sự chạy.
        app.coordinator.events.put(("progress", ("only-one",)))
        app._process_events()  # không được raise
        app.coordinator.events.put(("log", "PUMP-ALIVE-PROBE"))
        deadline = time.time() + 2.0
        while time.time() < deadline:
            app.update()
            if "PUMP-ALIVE-PROBE" in app.log_text.get("1.0", "end"):
                break
            time.sleep(0.05)
        content = app.log_text.get("1.0", "end")
        assert "PUMP-ALIVE-PROBE" in content, "bơm chết sau sự kiện lỗi"
        assert "Lỗi xử lý sự kiện" in content, "thiếu log báo sự kiện lỗi"
    finally:
        app.destroy()


def test_log_frame_height():
    """Khung log phải đủ cao và tab phải cuộn được: trước đây cửa sổ
    xin cao 1608px, Notebook chiếm ~1070px, khung log bị ép còn 1px."""
    app = _make_app()
    try:
        log_frame = app.log_text.master
        log_h = log_frame.winfo_height()
        assert log_h >= 180, f"khung log chỉ {log_h}px"
        notebook_h = app.notebook.winfo_height()
        assert notebook_h <= 700, f"notebook cao {notebook_h}px"
        # Nội dung tab Muse vẫn tới được bằng cách cuộn.
        muse_scroll = app.notebook.nametowidget(
            app.notebook.tabs()[0]
        )
        region = muse_scroll.canvas.cget("scrollregion")
        _x1, _y1, _x2, y2 = (float(v) for v in region.split())
        tab_req = app.muse_tab.winfo_reqheight()
        assert y2 >= tab_req - 5, (
            f"vùng cuộn {y2}px < chiều cao tab {tab_req}px"
        )
    finally:
        app.destroy()


def test_vault_and_run_visible_without_scrolling():
    """Sau khi sắp xếp lại (2026-11-13): ở cửa sổ mặc định
    820x980, Kho thẻ (bảng + ô dán + nút Nhập kho thẻ) và các
    nút chạy phải thấy được mà KHÔNG cần cuộn; khung log vẫn
    >= 180px. Tab Meta: nút Đăng kí Meta cũng phải thấy được."""
    app = _make_app()
    try:
        log_h = app.log_text.master.winfo_height()
        assert log_h >= 180, f"khung log chỉ {log_h}px"

        def _check_visible(canvas, widget, name):
            top = canvas.winfo_rooty()
            bottom = top + canvas.winfo_height()
            y = widget.winfo_rooty()
            h = widget.winfo_height()
            assert y >= top and y + h <= bottom, (
                f"{name} ngoài vùng nhìn thấy: y={y}..{y + h}, "
                f"viewport {top}..{bottom} (cần cuộn)"
            )

        app.notebook.select(0)
        app.update_idletasks()
        app.update()
        muse_scroll = app.notebook.nametowidget(app.notebook.tabs()[0])
        canvas = muse_scroll.canvas
        for name, widget in (
            ("bảng Kho thẻ", app.muse_tab.vault_tree),
            ("ô dán thẻ", app.muse_tab.bulk_card_text),
            ("nút Nhập kho thẻ", app.muse_tab.import_bulk_button),
            ("nút Đăng nhập Muse", app.muse_tab.muse_login_button),
            ("nút Bắt đầu đăng ký", app.muse_tab.start_button),
        ):
            _check_visible(canvas, widget, name)

        app.notebook.select(1)
        app.update_idletasks()
        app.update()
        meta_scroll = app.notebook.nametowidget(app.notebook.tabs()[1])
        _check_visible(
            meta_scroll.canvas,
            app.meta_tab.meta_start_button,
            "nút Đăng kí Meta",
        )
    finally:
        app.destroy()


def test_first_launch_migration_opens_window():
    """Hồi quy cho sự cố 2026-11-13 18:43 (bản 4f3814f8): lần mở
    đầu sau nâng cấp, migration chuyển thẻ sang file + dọn thẻ mồ
    côi, rồi MuseTab ghi log báo cáo NGAY LÚC DỰNG giao diện —
    nhưng log_text chỉ được tạo SAU các tab trong _build_ui nên
    append_log vỡ AttributeError và app chết trước khi hiện cửa
    sổ ("chạy py main.py k lên luôn"). Mô phỏng đúng trạng thái
    máy Long: vault.json có Thẻ 1, vault.pending.json còn 2 thẻ
    chờ, CM còn Thẻ 1 + 80 thẻ mồ côi + capsolver/mail. App phải
    mở được, log phải có báo cáo chuyển/dọn, kho phải có 3 thẻ,
    capsolver/mail không bị đụng."""
    _store.clear()
    for name in (
        "vault.json",
        "vault.pending.json",
        "cards.json",
        "cards.json.tmp",
        "vault.json.tmp",
    ):
        try:
            (TMP / name).unlink()
        except OSError:
            pass
    main_id = "card-aaaabbbb"
    (TMP / "vault.json").write_text(
        json.dumps(
            {
                "active": "",
                "cards": [
                    {
                        "id": main_id,
                        "label": "Thẻ 1",
                        "limit": 0,
                        "used": 0,
                    }
                ],
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    (TMP / "vault.pending.json").write_text(
        json.dumps(
            [
                {"id": "card-ccccdddd", "label": "Thẻ 2", "limit": 0,
                 "used": 0},
                {"id": "card-eeeeffff", "label": "Thẻ 3", "limit": 0,
                 "used": 0},
            ],
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    def _put_card(service, number):
        for field, value in (
            ("number", number),
            ("expiry", "11/27"),
            ("cvv", "122"),
            ("zip", "10001"),
        ):
            _store[(service, field)] = value
        _store[(service, "card_success_count")] = "0"

    prefix = "MuseAIHelper:card:"
    _put_card(f"{prefix}{main_id}", "4111111111111111")
    for i in range(80):
        _put_card(f"{prefix}card-orphan{i:04d}", "4000056655665556")
    _store[("MuseAIHelper:capsolver", "api_key")] = "fake-key"
    _store[("MuseAIHelper:mail.tm", "x")] = "y"

    # TargetName như CM thật ghi (WinVaultKeyring): trường ghi
    # cuối nằm ở TargetName=service, còn lại dạng gộp
    # {username}@{service}.
    targets = []
    by_service = {}
    for service, field in list(_store):
        by_service.setdefault(service, []).append(field)
    for service, fields in by_service.items():
        if service.startswith(prefix):
            targets.append(service)
            for field in fields[:-1]:
                targets.append(f"{field}@{service}")
        else:
            targets.append(service)
    fake_win32cred = types.ModuleType("win32cred")
    fake_win32cred.CredEnumerate = (
        lambda f, fl: [{"TargetName": t} for t in targets]
    )
    fake_pywin32 = types.ModuleType("win32ctypes.pywin32")
    fake_pywin32.win32cred = fake_win32cred
    sys.modules["win32ctypes"] = types.ModuleType("win32ctypes")
    sys.modules["win32ctypes.pywin32"] = fake_pywin32
    sys.modules["win32ctypes.pywin32.win32cred"] = fake_win32cred
    storage._MIGRATION_DONE = False
    storage._LAST_MIGRATION = None
    storage._WIN32CRED = storage._UNSET

    app = _make_app()
    try:
        # App mở được (trước đây vỡ AttributeError ở đây).
        assert app.winfo_exists(), "cửa sổ không mở được"
        report = storage._last_migration_report()
        assert report is not None, "không có báo cáo migration"
        assert report["migrated"] == 1, f"migrated={report}"
        assert report["orphans_removed"] == 80, f"orphans={report}"
        content = app.log_text.get("1.0", "end")
        assert "đã chuyển 1 thẻ" in content, f"thiếu log chuyển: {content}"
        assert "đã dọn 80 thẻ mồ côi" in content, f"thiếu log dọn: {content}"
        children = app.muse_tab.vault_tree.get_children()
        assert len(children) == 3, f"kho phải có 3 thẻ, có {len(children)}"
        assert (TMP / "cards.json").exists(), "thiếu file cards.json"
        assert _store.get(("MuseAIHelper:capsolver", "api_key")) == (
            "fake-key"
        ), "capsolver bị đụng khi dọn mồ côi"
        assert _store.get(("MuseAIHelper:mail.tm", "x")) == "y", (
            "mail bị đụng khi dọn mồ côi"
        )
    finally:
        app.destroy()


if __name__ == "__main__":
    _run("lane crash resets running + finished", test_lane_crash_resets_running)
    _run("login loop crash resets running + finished",
         test_login_loop_crash_resets_running)
    _run("set_running does not raise (vault Treeview)",
         test_set_running_no_raise, needs_display=True)
    _run("Meta start handler calls start + logs",
         test_meta_start_handler, needs_display=True)
    _run("event pump survives bad event + reschedules",
         test_pump_survives_bad_event, needs_display=True)
    _run("log frame tall, notebook bounded, tabs scrollable",
         test_log_frame_height, needs_display=True)
    _run("vault + run buttons visible without scrolling",
         test_vault_and_run_visible_without_scrolling, needs_display=True)
    _run("first launch migration opens window + logs report",
         test_first_launch_migration_opens_window, needs_display=True)
    failed = [n for n, ok, _ in RESULTS if ok is False]
    skipped = [n for n, ok, _ in RESULTS if ok is None]
    print(f"--- {len(RESULTS) - len(failed) - len(skipped)}/"
          f"{len(RESULTS) - len(skipped)} passed, "
          f"{len(skipped)} skipped, failures: {failed}")
    sys.exit(1 if failed else 0)
