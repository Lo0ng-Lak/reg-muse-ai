"""Kiểm tra quy tắc strike proxy hỏng: mở meta.ai lỗi tải trang
hết 5 lần thử trong một lượt chạy helper = 1 strike; đủ 2
strike (2 lượt chạy liên tiếp đều hỏng) mới đổi proxy + đưa
vào kho proxy lỗi; strike 1 thì lượt sau vẫn dùng đúng proxy
đó; mở được meta.ai thì xoá strike.

Chạy:  python test_proxy_load_fail.py
Không cần Windows, trình duyệt hay proxy thật: page.goto được
thay bằng một page giả lập theo kịch bản lỗi. Không dùng proxy
thật — chỉ dùng chuỗi lỗi Playwright mẫu và response giả.
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


class _FakeResponse:
    """Response giả của page.goto: chỉ cần thuộc tính status."""

    def __init__(self, status):
        self.status = status


class _FakePage:
    """Page giả: goto() đi theo kịch bản — mỗi phần tử là None
    (mở thành công, trả response None), một Exception sẽ raise,
    hoặc một _FakeResponse sẽ được trả về. Hết kịch bản thì mở
    thành công."""

    def __init__(self, script):
        self.script = list(script)
        self.goto_calls = 0

    def goto(self, url, wait_until=None, timeout=None):
        self.goto_calls += 1
        outcome = self.script.pop(0) if self.script else None
        if isinstance(outcome, _FakeResponse):
            return outcome
        if outcome is not None:
            raise outcome
        return None

    def wait_for_timeout(self, ms):
        pass


def _has_load_fail_signal(captured):
    return any(line.startswith("PROXY_LOAD_FAIL") for line in captured)


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
        _load_fail("net::ERR_PROXY_AUTH_UNSUPPORTED")
    )
    assert browser_helper._is_page_load_error(
        Exception("Page.goto: HTTP ERROR 407 at https://www.meta.ai/")
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
    assert not _has_load_fail_signal(captured), captured


def test_open_one_load_fail_then_success():
    # Một lần lỗi đơn lẻ: thử lại, mở được, KHÔNG tính strike.
    page = _FakePage([_load_fail(), None])
    captured = _capture_status()
    browser_helper._open_meta_ai(page)
    assert page.goto_calls == 2, page.goto_calls
    assert not _has_load_fail_signal(captured), captured


def test_open_two_load_fails_then_success():
    # Hai lần lỗi tải trang rồi mở được: vẫn thử lại đủ, KHÔNG
    # tính strike — quy tắc cũ "bỏ proxy sau 2 lần thử" đã bị
    # chủ tool sửa lại, chỉ hết 5 lần thử mới tính 1 strike.
    page = _FakePage([_load_fail(), _load_fail(), None])
    captured = _capture_status()
    browser_helper._open_meta_ai(page)
    assert page.goto_calls == 3, page.goto_calls
    assert not _has_load_fail_signal(captured), captured


def test_open_mixed_errors_then_success():
    # Lỗi tải trang xen lỗi khác rồi mở được: không tính strike.
    page = _FakePage([_load_fail(), _other_error(), None])
    captured = _capture_status()
    browser_helper._open_meta_ai(page)
    assert page.goto_calls == 3, page.goto_calls
    assert not _has_load_fail_signal(captured), captured


def test_open_other_errors_retry_five_times():
    # Lỗi ngoài lớp tải trang: vẫn thử lại đủ 5 lần như cũ
    # (1 lần mở + 5 lần thử lại = 6 lần goto), không báo tín
    # hiệu proxy hỏng.
    page = _FakePage([_other_error()] * 6)
    captured = _capture_status()
    try:
        browser_helper._open_meta_ai(page)
    except RuntimeError as error:
        assert "5 lần thử lại" in str(error), str(error)
    else:
        raise AssertionError("expected RuntimeError")
    assert page.goto_calls == 6, page.goto_calls
    assert not _has_load_fail_signal(captured), captured


def test_open_all_load_fails_signals_after_five():
    # Hết 5 lần thử đều lỗi tải trang: thử đủ 6 lần goto rồi
    # mới báo PROXY_LOAD_FAIL (1 strike) — không dừng sớm.
    page = _FakePage([_load_fail()] * 6)
    captured = _capture_status()
    try:
        browser_helper._open_meta_ai(page)
    except RuntimeError as error:
        assert "1 strike" in str(error), str(error)
    else:
        raise AssertionError("expected RuntimeError")
    assert page.goto_calls == 6, page.goto_calls
    assert _has_load_fail_signal(captured), captured


def test_open_timeout_counts_as_load_error():
    page = _FakePage(
        [Exception("Page.goto: Timeout 30000ms exceeded.")] * 6
    )
    captured = _capture_status()
    try:
        browser_helper._open_meta_ai(page)
    except RuntimeError:
        pass
    else:
        raise AssertionError("expected RuntimeError")
    assert page.goto_calls == 6, page.goto_calls
    assert _has_load_fail_signal(captured), captured


def test_open_407_response_signals():
    # Proxy trả response 407 mà goto không raise: vẫn tính là
    # lỗi tải trang, hết 5 lần thử thì báo PROXY_LOAD_FAIL.
    page = _FakePage([_FakeResponse(407)] * 6)
    captured = _capture_status()
    try:
        browser_helper._open_meta_ai(page)
    except RuntimeError as error:
        assert "1 strike" in str(error), str(error)
    else:
        raise AssertionError("expected RuntimeError")
    assert page.goto_calls == 6, page.goto_calls
    assert _has_load_fail_signal(captured), captured


def test_open_407_then_success():
    # Response 407 một lần rồi mở được (response 200): không
    # tính strike.
    page = _FakePage([_FakeResponse(407), _FakeResponse(200)])
    captured = _capture_status()
    browser_helper._open_meta_ai(page)
    assert page.goto_calls == 2, page.goto_calls
    assert not _has_load_fail_signal(captured), captured


def test_line_flags_facebook():
    assert proxy_pool._line_flags_facebook(
        "PROXY_FACEBOOK: proxy này đẩy sang Facebook "
        "(facebook.com) — bỏ proxy, không dùng lại."
    )
    assert not proxy_pool._line_flags_facebook(
        "PROXY_BURN: Meta chặn tạo tài khoản trên IP này"
    )
    assert not proxy_pool._line_flags_load_fail(
        "PROXY_FACEBOOK: proxy này đẩy sang Facebook"
    )


def test_line_flags_load_fail():
    assert proxy_pool._line_flags_load_fail(
        "PROXY_LOAD_FAIL: mở meta.ai lỗi tải trang hết 5 lần thử"
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


def _make_rotation():
    pool = [
        "http://user1:pass1@1.2.3.4:8080",
        "http://user2:pass2@5.6.7.8:8080",
        "http://user3:pass3@9.10.11.12:8080",
    ]
    return pool, proxy_pool._ProxyRotation(pool, per_proxy=5)


def test_rotation_strikes_count_per_proxy():
    # Strike đếm theo từng proxy (khóa = nguyên dòng proxy):
    # proxy khác nhau đếm riêng.
    pool, rotation = _make_rotation()
    assert rotation.record_load_fail_strike(pool[0]) == 1
    assert rotation.record_load_fail_strike(pool[0]) == 2
    assert rotation.record_load_fail_strike(pool[1]) == 1
    assert rotation.record_load_fail_strike(pool[1]) == 2
    # Cùng một dòng proxy dù viết hoa/thừa khoảng trắng vẫn
    # chung một khóa (chuẩn hoá qua _proxy_key).
    assert (
        rotation.record_load_fail_strike("  HTTP://USER2:PASS2@5.6.7.8:8080 ")
        == 3
    )


def test_rotation_strikes_clear_on_success():
    # Mở được meta.ai thì xoá strike: lần hỏng sau đó chỉ tính
    # lại từ 1 (strike chỉ tính các lần hỏng LIÊN TIẾP).
    pool, rotation = _make_rotation()
    assert rotation.record_load_fail_strike(pool[0]) == 1
    rotation.clear_load_fail_strikes(pool[0])
    assert rotation.record_load_fail_strike(pool[0]) == 1
    # Xoá strike của proxy này không động tới proxy khác.
    rotation.record_load_fail_strike(pool[1])
    rotation.clear_load_fail_strikes(pool[0])
    assert rotation.record_load_fail_strike(pool[1]) == 2


def test_strike_flow_hold_then_drop():
    # Mô phỏng đúng nhịp lane dùng rotation: lượt 1 hỏng hết 5
    # lần (strike 1) → GIỮ proxy (current() không đổi, chưa
    # retire); lượt 2 hỏng tiếp (strike 2) → BỎ proxy
    # (retire_current đẩy sang proxy kế, proxy cũ không quay
    # lại khi xoay vòng).
    pool, rotation = _make_rotation()
    _, line = rotation.current()
    assert line == pool[0]
    strikes = rotation.record_load_fail_strike(line)
    assert strikes == 1
    assert rotation.current()[1] == pool[0], "strike 1 phải giữ proxy"
    assert not rotation.all_retired()
    strikes = rotation.record_load_fail_strike(line)
    assert strikes == 2
    rotation.retire_current()
    assert rotation.current()[1] == pool[1], "strike 2 phải đổi proxy"
    assert 0 in rotation.retired


def test_strike_flow_failure_then_success():
    # Lượt 1 hỏng (strike 1), lượt 2 mở được meta.ai (xoá
    # strike), lượt 3 hỏng lại chỉ tính strike 1 — proxy KHÔNG
    # bị bỏ, vẫn là proxy hiện tại.
    pool, rotation = _make_rotation()
    _, line = rotation.current()
    assert rotation.record_load_fail_strike(line) == 1
    rotation.clear_load_fail_strikes(line)
    assert rotation.record_load_fail_strike(line) == 1
    assert rotation.current()[1] == pool[0], "proxy không bị bỏ"
    assert not rotation.all_retired()


TESTS = [
    test_is_page_load_error,
    test_open_success_first_try,
    test_open_one_load_fail_then_success,
    test_open_two_load_fails_then_success,
    test_open_mixed_errors_then_success,
    test_open_other_errors_retry_five_times,
    test_open_all_load_fails_signals_after_five,
    test_open_timeout_counts_as_load_error,
    test_open_407_response_signals,
    test_open_407_then_success,
    test_line_flags_facebook,
    test_line_flags_load_fail,
    test_rotation_strikes_count_per_proxy,
    test_rotation_strikes_clear_on_success,
    test_strike_flow_hold_then_drop,
    test_strike_flow_failure_then_success,
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
