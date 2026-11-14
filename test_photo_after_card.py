"""Kiểm tra bước up ảnh SAU bước nhập thẻ trong luồng Muse
(có điều kiện — sư huynh dặn 2026-11-13).

- Đã vào được Muse → bỏ qua bước ảnh, đi tiếp tới các bước mã.
- Muse vẫn đòi ảnh (hiện prompt up ảnh) → up file selfie đã lưu
  (hoặc chờ up tay khi chưa có file) rồi bấm tiếp.
- Không có prompt ảnh mà vẫn chưa vào được → trả về ngay để
  _confirm_final_step xử lý tiếp, không kẹt.

Chạy:  python test_photo_after_card.py
Không cần trình duyệt thật: dùng trang giả (fake page) mô phỏng
URL + các dấu hiệu trang (ô file, nút Mời, nút Get started).
Real-browser flow KHÔNG test được từ Linux — Long test trên
Windows.
"""

import sys
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

import paths

TMP = Path(tempfile.mkdtemp(prefix="reg-test-photo-"))
paths.data_file = lambda name: TMP / name


class _FakeKeyring:
    def __init__(self):
        self.store = {}

    def get_password(self, service, field):
        return self.store.get((service, field))

    def set_password(self, service, field, value):
        self.store[(service, field)] = value

    def delete_password(self, service, field):
        self.store.pop((service, field), None)


sys.modules["keyring"] = _FakeKeyring()

import browser_helper  # noqa: E402
import storage  # noqa: E402


class _FakeElement:
    """Một phần tử trang giả: hiện/bấm được, bấm thì đổi trạng thái."""

    def __init__(self, page, kind, name=None):
        self.page = page
        self.kind = kind
        self.name = name

    def is_visible(self):
        return True

    def is_enabled(self):
        return True

    def bounding_box(self):
        return {"x": 10, "y": 10, "width": 50, "height": 20}

    def click(self, timeout=None):
        self.page.clicked.append(self.name or self.kind)
        # Bấm Continue ở trang ảnh / Get started ở cổng → vào Muse
        # (URL đổi sang giao diện như trình duyệt thật).
        if self.kind == "button" and self.name in (
            "Continue",
            "Confirm",
            "Submit",
            "Get started",
            "Bắt đầu",
        ):
            self.page.state = "muse"
            self.page.url = "https://muse.ai/"

    def set_input_files(self, path):
        self.page.uploaded_files.append(str(path))


class _FakeLocator:
    def __init__(self, elements):
        self._elements = list(elements)

    def count(self):
        return len(self._elements)

    @property
    def first(self):
        if self._elements:
            return self._elements[0]
        return _FakeElement(None, "missing")

    def nth(self, index):
        return self._elements[index]


class _FakePage:
    """Trang giả với 5 trạng thái: muse (đã vào giao diện, có nút
    Mời), muse-chat (đã vào giao diện chat nhưng KHÔNG có nút
    Mời, ô chat hiện, ô file đính kèm của khung chat hiện — đúng
    trường hợp thật làm tool up nhầm ảnh vào khung chat),
    photo (đang ở trang up ảnh), gate (đang ở cổng Get started),
    loading (chưa rõ màn nào)."""

    def __init__(self, url, state="loading"):
        self.url = url
        self.state = state
        self.closed = False
        self.uploaded_files = []
        self.clicked = []
        self.viewport_size = {"width": 1280, "height": 720}

    def is_closed(self):
        return self.closed

    def wait_for_timeout(self, ms):
        pass

    def evaluate(self, js):
        if 'input[type="file"]' in js:
            # _verification_file_input_index: ô file trong bối cảnh
            # xác minh chỉ có ở trang up ảnh; khung chat chỉ có ô
            # đính kèm nên trả -1.
            return 0 if self.state == "photo" else -1
        return False

    def locator(self, selector):
        if selector == 'input[type="file"]' and self.state in (
            "photo",
            "muse-chat",
        ):
            return _FakeLocator([_FakeElement(self, "file")])
        return _FakeLocator([])

    def get_by_placeholder(self, text, exact=False):
        placeholder = "Ask Meta AI anything..."
        if self.state in ("muse", "muse-chat"):
            if hasattr(text, "search"):
                matched = bool(text.search(placeholder))
            else:
                matched = str(text).lower() in placeholder.lower()
            if matched:
                return _FakeLocator(
                    [_FakeElement(self, "textbox", "composer")]
                )
        return _FakeLocator([])

    def get_by_text(self, label, exact=True):
        if self.state == "photo" and label in (
            "Upload a verification selfie",
            "Upload image",
        ):
            return _FakeLocator([_FakeElement(self, "text", label)])
        return _FakeLocator([])

    def get_by_role(self, role, name=None, exact=True):
        if role != "button":
            return _FakeLocator([])
        if name in ("Mời", "Invite") and self.state == "muse":
            return _FakeLocator([_FakeElement(self, "button", name)])
        if self.state == "gate" and name in (
            "Get started",
            "Confirm",
            "Continue",
            "Bắt đầu",
        ):
            return _FakeLocator([_FakeElement(self, "button", name)])
        if self.state == "photo" and name in (
            "Continue",
            "Confirm",
            "Submit",
        ):
            return _FakeLocator([_FakeElement(self, "button", name)])
        return _FakeLocator([])


def _muse_page():
    return _FakePage("https://muse.ai/", state="muse")


def _muse_chat_page():
    return _FakePage("https://muse.ai/", state="muse-chat")


def _photo_page():
    return _FakePage("https://muse.ai/age-verification", state="photo")


def _gate_page():
    return _FakePage("https://muse.ai/age-verification", state="gate")


def _loading_page():
    return _FakePage("https://muse.ai/age-verification", state="loading")


def _capture_status():
    captured = []
    orig = browser_helper._status
    browser_helper._status = captured.append
    return captured, orig


def test_decision_skip_when_in_muse():
    assert browser_helper._photo_after_card_decision(_muse_page()) == "skip"


def test_decision_photo_when_prompt_visible():
    assert (
        browser_helper._photo_after_card_decision(_photo_page()) == "photo"
    )


def test_decision_none_when_neither():
    assert (
        browser_helper._photo_after_card_decision(_loading_page()) == "none"
    )
    assert browser_helper._photo_after_card_decision(_gate_page()) == "none"


def test_handle_skip_when_in_muse():
    captured, orig = _capture_status()
    try:
        result = browser_helper._handle_photo_after_card(_muse_page())
    finally:
        browser_helper._status = orig
    assert result == "skipped", result
    assert any("bỏ qua bước up ảnh" in s for s in captured), captured


def test_handle_photo_uploads_saved_selfie():
    selfie = TMP / "selfie.png"
    selfie.write_bytes(b"\x89PNG\r\n\x1a\n")
    (TMP / "selfie.txt").write_text(str(selfie) + "\n", encoding="utf-8")
    page = _photo_page()
    captured, orig = _capture_status()
    try:
        result = browser_helper._handle_photo_after_card(page)
    finally:
        browser_helper._status = orig
    assert result == "photo-done", result
    assert page.uploaded_files == [str(selfie)], page.uploaded_files
    assert any("Đã up ảnh selfie" in s for s in captured), captured
    assert page.state == "muse", "sau bước ảnh phải vào được Muse"


def test_handle_photo_without_saved_selfie_waits_manual():
    try:
        (TMP / "selfie.txt").unlink()
    except OSError:
        pass
    page = _photo_page()
    captured, orig = _capture_status()
    try:
        result = browser_helper._handle_photo_after_card(page)
    finally:
        browser_helper._status = orig
    assert result == "photo-done", result
    assert page.uploaded_files == [], "không có file thì không tự up"
    assert any("up tay" in s for s in captured), captured
    assert "Continue" in page.clicked, page.clicked


def test_handle_none_returns_fast_without_photo():
    """Không có prompt ảnh mà vẫn chưa vào được: trả về ngay để
    _confirm_final_step xử lý tiếp — không được kẹt chờ hết giờ."""
    captured, orig = _capture_status()
    try:
        started = time.monotonic()
        result = browser_helper._handle_photo_after_card(_gate_page())
        elapsed = time.monotonic() - started
    finally:
        browser_helper._status = orig
    assert result == "none", result
    assert elapsed < 2.0, f"kẹt {elapsed:.1f}s ở bước ảnh"


def test_handle_none_after_settle_when_loading():
    captured, orig = _capture_status()
    try:
        started = time.monotonic()
        result = browser_helper._handle_photo_after_card(
            _loading_page(), timeout_seconds=30, settle_seconds=0.2
        )
        elapsed = time.monotonic() - started
    finally:
        browser_helper._status = orig
    assert result == "none", result
    assert elapsed < 5.0, f"kẹt {elapsed:.1f}s ở bước ảnh"


def test_handle_closed_page():
    page = _muse_page()
    page.closed = True
    assert browser_helper._handle_photo_after_card(page) == "closed"


def test_decision_skip_when_in_muse_chat_without_invite_button():
    """Đúng lỗi thật 2026-11-14: đã vào giao diện chat (ô chat
    hiện) nhưng nút Mời không hiện — vẫn phải nhận ra đã vào
    Muse và bỏ qua bước ảnh, không được tính ô file đính kèm
    của khung chat là trang up ảnh."""
    page = _muse_chat_page()
    assert browser_helper._is_muse_chat_ui(page) is True
    assert (
        browser_helper._photo_after_card_decision(page) == "skip"
    ), "ô file của khung chat không được tính là trang up ảnh"


def test_selfie_page_not_visible_on_chat_ui():
    assert (
        browser_helper._meta_selfie_page_visible(_muse_chat_page())
        is False
    ), "ô đính kèm của khung chat không phải ô up ảnh xác minh"


def test_upload_refuses_chat_ui():
    selfie = TMP / "selfie.png"
    selfie.write_bytes(b"\x89PNG\r\n\x1a\n")
    page = _muse_chat_page()
    ok, reason = browser_helper._meta_upload_selfie(page, selfie)
    assert ok is False, "không được up ảnh khi đang ở giao diện chat"
    assert "khung chat" in reason, reason
    assert page.uploaded_files == [], "không được đặt file vào ô chat"


def test_handle_skips_photo_in_muse_chat():
    page = _muse_chat_page()
    captured, orig = _capture_status()
    try:
        result = browser_helper._handle_photo_after_card(page)
    finally:
        browser_helper._status = orig
    assert result == "skipped", result
    assert page.uploaded_files == [], "không được up gì vào khung chat"
    assert any("bỏ qua bước up ảnh" in s for s in captured), captured


TESTS = [
    test_decision_skip_when_in_muse,
    test_decision_photo_when_prompt_visible,
    test_decision_none_when_neither,
    test_handle_skip_when_in_muse,
    test_handle_photo_uploads_saved_selfie,
    test_handle_photo_without_saved_selfie_waits_manual,
    test_handle_none_returns_fast_without_photo,
    test_handle_none_after_settle_when_loading,
    test_handle_closed_page,
    test_decision_skip_when_in_muse_chat_without_invite_button,
    test_selfie_page_not_visible_on_chat_ui,
    test_upload_refuses_chat_ui,
    test_handle_skips_photo_in_muse_chat,
]


def main():
    passed = 0
    failed = 0
    for test in TESTS:
        try:
            (TMP / "selfie.txt").unlink()
        except OSError:
            pass
        storage._MIGRATION_DONE = False
        storage._LAST_MIGRATION = None
        try:
            test()
        except Exception as error:
            failed += 1
            print(f"FAIL {test.__name__}: {error!r}")
        else:
            passed += 1
            print(f"PASS {test.__name__}")
    print(f"\n{passed} passed, {failed} failed")
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
