"""Kiểm tra quy tắc proxy hỏng mới: mở meta.ai lỗi tải trang
2 lần liên tiếp thì đổi proxy + đưa vào kho proxy lỗi.

Chạy:  python test_proxy_load_fail.py
Không cần Windows, trình duyệt hay proxy thật: page.goto được
thay bằng một page giả lập theo kịch bản lỗi. Không dùng proxy
thật — chỉ dùng chuỗi lỗi Playwright mẫu.
"""

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))


class _FakeKeyring:
    """Bản giả tối thiểu của keyring, giống test_card_vault.py."""

    def __init__(self):
        self.store = {}

    def get_password(self, service, field):
        return self.store.get((service, field))

    def set_password(self, service, field, value):
        self.store[(service, field)] = value

    def delete_password(self, service, field):
        self.store.pop((service, field), None)


sys.modules["keyring"] = _FakeKeyring()

import browser_helper
import proxy_pool

_ORIGINAL_STATUS = browser_helper._status


def reset():
    browser_helper._status = _ORIGINAL_STATUS


def _capture_status():
    captured = []
    browser_helper._status = captured.append
    return captured


def _load_fail(kind="net::ERR_HTTP_RESPONSE_CODE_FAILURE"):
    return Exception(f"Page.goto: {kind} at https://www.meta.ai/")


def _other_error():
    return Exception("Target page, context or browser has been closed")


class _FakePage:
    """Page giả: goto() đi theo kịch bản — mỗi phần tử là None
    (mở thành công) hoặc một Exception sẽ raise. Hết kịch bản
    thì mở thành công."""

    def __init__(self, script):
        self.script = list(script)
        self.goto_calls = 0

    def goto(self, url, wait_until=None, timeout=None):
        self.goto_calls += 1
        outcome = self.script.pop(0) if self.script else None
        if outcome is not None:
            raise outcome

    def wait_for_timeout(self, ms):
        pass


def test_is_page_load_error():
    assert browser_helper._is_page_load_error(
        _load_fail("net::ERR_HTTP_RESPONSE_CODE_FAILURE")
    )
    assert browser_helper._is_page_load_error(
        _load_fail("net::ERR_TUNNEL_CONNECTION_FAILED")
    )
    assert browser_helper._is_page_load_error(
        _load_fail("net::ERR_PROXY_CONNECTION_FAILED")
    )
    assert browser_helper._is_page_load_error(
        _load_fail("net::ERR_CONNECTION_REFUSED")
    )
    assert browser_helper._is_page_load_error(
        Exception("Page.goto: Timeout 30000ms exceeded.")
    )
    assert browser_helper._is_page_load_error(
        Exception(
            "Page.goto: net::ERR_HTTP_RESPONSE_CODE_FAILURE "
            "at https://www.meta.ai/\nCall log:\n  - navigating"
        )
    )
    assert not browser_helper._is_page_load_error(_other_error())
    assert not browser_helper._is_page_load_error(
        Exception("Something went wrong")
    )
    assert not browser_helper._is_page_load_error(Exception(""))


def test_open_success_first_try():
    page = _FakePage([None])
    captured = _capture_status()
    browser_helper._open_meta_ai(page)
    assert page.goto_calls == 1, page.goto_calls
    assert not any(
        line.startswith("PROXY_LOAD_FAIL") for line in captured
    ), captured


def test_open_one_load_fail_then_success():
    # Một lần lỗi đơn lẻ: thử lại, mở được, KHÔNG đánh dấu proxy.
    page = _FakePage([_load_fail(), None])
    captured = _capture_status()
    browser_helper._open_meta_ai(page)
    assert page.goto_calls == 2, page.goto_calls
    assert not any(
        line.startswith("PROXY_LOAD_FAIL") for line in captured
    ), captured


def test_open_two_load_fails_signals():
    # Hai lần lỗi tải trang liên tiếp: báo tín hiệu + dừng ngay
    # (không đốt thêm 3 lần thử còn lại).
    page = _FakePage([_load_fail(), _load_fail()])
    captured = _capture_status()
    try:
        browser_helper._open_meta_ai(page)
    except RuntimeError as error:
        assert "2 lần thử liên tiếp" in str(error), str(error)
    else:
        raise AssertionError("expected RuntimeError")
    assert page.goto_calls == 2, page.goto_calls
    assert any(
        line.startswith("PROXY_LOAD_FAIL") for line in captured
    ), captured


def test_open_streak_resets_on_other_error():
    # Lỗi khác chen giữa thì chuỗi lỗi tải trang bị ngắt: lần
    # lỗi tải trang sau đó chỉ tính là lần 1, không báo tín hiệu.
    page = _FakePage([_load_fail(), _other_error(), None])
    captured = _capture_status()
    browser_helper._open_meta_ai(page)
    assert page.goto_calls == 3, page.goto_calls
    assert not any(
        line.startswith("PROXY_LOAD_FAIL") for line in captured
    ), captured


def test_open_other_errors_retry_five_times():
    # Lỗi ngoài lớp tải trang: vẫn thử lại đủ 5 lần như cũ,
    # không báo tín hiệu proxy hỏng.
    page = _FakePage([_other_error()] * 6)
    captured = _capture_status()
    try:
        browser_helper._open_meta_ai(page)
    except RuntimeError as error:
        assert "5 lần thử lại" in str(error), str(error)
    else:
        raise AssertionError("expected RuntimeError")
    assert page.goto_calls == 6, page.goto_calls
    assert not any(
        line.startswith("PROXY_LOAD_FAIL") for line in captured
    ), captured


def test_open_all_load_fails_stops_at_two():
    page = _FakePage([_load_fail()] * 6)
    captured = _capture_status()
    try:
        browser_helper._open_meta_ai(page)
    except RuntimeError:
        pass
    else:
        raise AssertionError("expected RuntimeError")
    assert page.goto_calls == 2, page.goto_calls
    assert any(
        line.startswith("PROXY_LOAD_FAIL") for line in captured
    ), captured


def test_open_timeout_counts_as_load_error():
    page = _FakePage(
        [
            Exception("Page.goto: Timeout 30000ms exceeded."),
            Exception("Page.goto: Timeout 30000ms exceeded."),
        ]
    )
    captured = _capture_status()
    try:
        browser_helper._open_meta_ai(page)
    except RuntimeError:
        pass
    else:
        raise AssertionError("expected RuntimeError")
    assert page.goto_calls == 2, page.goto_calls
    assert any(
        line.startswith("PROXY_LOAD_FAIL") for line in captured
    ), captured


def test_line_flags_load_fail():
    assert proxy_pool._line_flags_load_fail(
        "PROXY_LOAD_FAIL: mở meta.ai lỗi tải trang 2 lần liên tiếp"
    )
    assert not proxy_pool._line_flags_load_fail(
        "PROXY_BURN: Meta chặn tạo tài khoản trên IP này"
    )
    assert not proxy_pool._line_flags_load_fail(
        "WAITLIST: IP thoát rơi vào khu vực chưa được hỗ trợ"
    )
    assert not proxy_pool._line_flags_load_fail(
        "Meta: mở meta.ai lỗi (Page.goto: net::ERR_HTTP_"
        "RESPONSE_CODE_FAILURE at https://www.meta.ai/) — "
        "thử lại lần 1/5..."
    )


TESTS = [
    test_is_page_load_error,
    test_open_success_first_try,
    test_open_one_load_fail_then_success,
    test_open_two_load_fails_signals,
    test_open_streak_resets_on_other_error,
    test_open_other_errors_retry_five_times,
    test_open_all_load_fails_stops_at_two,
    test_open_timeout_counts_as_load_error,
    test_line_flags_load_fail,
]


def main():
    passed = 0
    failed = 0
    for test in TESTS:
        reset()
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
