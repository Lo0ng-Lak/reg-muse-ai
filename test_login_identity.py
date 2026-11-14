"""Kiểm tra luồng đăng nhập Muse luôn vào ĐÚNG tài khoản
(sư huynh dặn 2026-11-14: đi qua trang khác rất nhiều, đăng
nhập lại tài khoản cũ rất nhiều; và mỗi lần mở trình duyệt phải
là trình duyệt sạch mặc dù cùng proxy, xoá hết cookie).

- Mỗi lần mở trình duyệt là context SẠCH hoàn toàn (không
  cookie/session cũ) kể cả khi dùng cùng một proxy.
- Hộp thư OTP phải khớp đúng email cần chạy (file mailbox.txt),
  không đọc mã từ hộp thư sai.
- Màn chọn tài khoản: chỉ bấm khi thấy đúng email cần đăng nhập;
  không thấy thì dừng, không bấm nhầm tài khoản cũ.
- Vào giao diện rồi phải xác minh tài khoản khớp email cần đăng
  nhập trước khi lấy mã mời + cookie; nhầm tài khoản thì đăng
  xuất + xoá phiên rồi lỗi rõ (không lưu).

Chạy:  python test_login_identity.py
Không cần trình duyệt thật: dùng trang giả (fake page/context).
Real-browser flow KHÔNG test được từ Linux — Long test trên
Windows.
"""

import json
import re
import sys
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

import paths

TMP = Path(tempfile.mkdtemp(prefix="reg-test-login-"))
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

browser_helper.LANE_SUFFIX = ""

INTENDED = "new-account-123@tempmail.plus"
OTHER = "old-account-999@tempmail.plus"


class _FakeElement:
    """Một phần tử trang giả: hiện/bấm được, bấm thì đổi trạng thái."""

    def __init__(self, page, kind, name=None, text=None):
        self.page = page
        self.kind = kind
        self.name = name
        self.text = text if text is not None else (name or "")
        self.value = ""

    def is_visible(self):
        return True

    def is_enabled(self):
        return True

    def wait_for(self, state=None, timeout=None):
        return None

    def fill(self, value, timeout=None):
        self.value = value
        filled = getattr(self.page, "filled", None)
        if isinstance(filled, list):
            filled.append(value)
        return None

    def press(self, key):
        return None

    def input_value(self, timeout=None):
        return self.value

    def click(self, timeout=None):
        self.page.clicked.append(self.name or self.kind)
        # Bấm nút avatar → mở menu tài khoản.
        if self.kind == "button" and self.name in ("Account", "Profile"):
            self.page.menu_open = True
        # Bấm mục tài khoản trên màn chọn → vào màn hình tiếp theo.
        if self.kind == "text" and self.text == self.page.intended_email:
            self.page.state = "card"
            self.page.body_text = (
                "Enter your card number to continue with Muse."
            )
        # Màn "Continue with your Meta Account": bấm Continue khi
        # email khớp → tới bước thẻ; bấm "Create new account"
        # khi không khớp (đang tạo tài khoản mới) → về form hồ sơ.
        # Tài khoản đã đăng nhập sẵn (after_link_continue =
        # "chat-ui") thì bấm Continue vào thẳng giao diện chat.
        if self.kind == "button" and self.name == "Continue":
            if getattr(self.page, "hold_link_screen", False):
                return
            if self.page.state == "meta-link-match":
                if getattr(self.page, "after_link_continue", "card") == "chat-ui":
                    self.page.state = "chat-ui"
                    self.page.body_text = (
                        "Ask Meta AI anything...\n"
                        "Welcome to Muse. Start a new chat."
                    )
                else:
                    self.page.state = "card"
                    self.page.body_text = (
                        "Enter your card number to continue with Muse."
                    )
        # Bấm Confirm trên form thẻ → bước thẻ hoàn tất, vào
        # giao diện Muse. reject_next_confirm: lần Confirm đầu
        # hiện hộp thoại "Unable to Add Card" (thẻ gắn quá nhiều
        # tài khoản quảng cáo) thay vì vào Muse.
        if self.kind == "button" and self.name == "Confirm":
            if (
                getattr(self.page, "reject_next_confirm", False)
                and self.page.state == "card"
            ):
                self.page.reject_next_confirm = False
                self.page.body_text = (
                    "Unable to Add Card\n"
                    "Your card is associated with too many ad accounts. "
                    "Please select a different payment method and try again.\n"
                    "OK\nLearn more"
                )
                self.page.buttons.append(
                    _FakeElement(self.page, "button", name="OK")
                )
                self.page.buttons.append(
                    _FakeElement(self.page, "button", name="Learn more")
                )
                return
            if (
                getattr(self.page, "too_many_attempts_once", False)
                and self.page.state == "card"
            ):
                self.page.too_many_attempts_once = False
                self.page.restore_card_on_reload = True
                self.page.body_text = (
                    "Too many attempts\n"
                    "Too many failed attempts. Please try again later.\n"
                    "Confirm your age\n"
                    "OK"
                )
                self.page.buttons.append(
                    _FakeElement(self.page, "button", name="OK")
                )
                return
            if self.page.state == "card":
                self.page.state = "muse"
                self.page.body_text = "Welcome to Muse."
        if self.kind == "button" and self.name == "OK":
            if "unable to add card" in (self.page.body_text or "").lower():
                self.page.body_text = (
                    "Enter your card number to continue with Muse."
                )
        if self.kind == "text" and self.name == "Create new account":
            if self.page.state == "meta-link-mismatch":
                self.page.state = "profile"
                self.page.body_text = (
                    "Finish creating your account — your birthday"
                )

    def inner_text(self, timeout=None):
        return self.text

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

    def inner_text(self, timeout=None):
        return "\n".join(el.inner_text() for el in self._elements)

    def wait_for(self, state=None, timeout=None):
        return None

    def is_enabled(self):
        return bool(self._elements) and all(
            el.is_enabled() for el in self._elements
        )

    def click(self, timeout=None):
        if self._elements:
            self._elements[0].click(timeout=timeout)


class _FakeKeyboard:
    def __init__(self, page):
        self.page = page
        self.pressed = []

    def press(self, key):
        self.pressed.append(key)
        if key == "Escape":
            self.page.menu_open = False


class _FakeContext:
    """Context trình duyệt giả: giữ cookie, mở trang, xoá cookie."""

    def __init__(self, browser):
        self.browser = browser
        self.cookie_list = []
        self.cleared = 0
        self.closed = False
        self.pages = []

    def new_page(self):
        page = _FakePage(self, "https://muse.ai/")
        self.pages.append(page)
        return page

    def clear_cookies(self):
        self.cookie_list = []
        self.cleared += 1

    def cookies(self):
        return list(self.cookie_list)

    def close(self):
        self.closed = True


class _FakeBrowser:
    """Trình duyệt giả: mỗi lần new_context() là một context mới.
    new_page() ở cấp browser KHÔNG được dùng nữa — gọi là lỗi."""

    def __init__(self):
        self.contexts = []
        self.new_page_calls = 0

    def new_context(self):
        context = _FakeContext(self)
        self.contexts.append(context)
        return context

    def new_page(self):
        self.new_page_calls += 1
        raise AssertionError(
            "browser.new_page() không được dùng — phải mở trang "
            "trong context mới qua _open_clean_page."
        )


class _FakePage:
    """Trang giả: URL + trạng thái màn hình + menu tài khoản."""

    def __init__(self, context, url, state="muse"):
        self.context = context
        self.url = url
        self.state = state
        self.closed = False
        self.body_text = ""
        self.menu_open = False
        self.menu_email = None
        self.intended_email = INTENDED
        self.clicked = []
        self.evaluated = []
        self.uploaded_files = []
        self.keyboard = _FakeKeyboard(self)
        # Các nút đăng ký theo (role, tên) để get_by_role tìm.
        self.buttons = []
        # Các ô nhập (theo placeholder) — giữ nguyên đối tượng
        # qua các lần gọi để giá trị đã điền còn đọc lại được.
        self.fields = {}
        # Bấm Continue ở màn liên kết thì đi đâu: "card" (tới
        # bước thẻ, mặc định) hoặc "chat-ui" (tài khoản đã đăng
        # nhập sẵn, vào thẳng giao diện chat).
        self.after_link_continue = "card"

    def is_closed(self):
        return self.closed

    def wait_for_timeout(self, ms):
        pass

    def goto(self, url, wait_until=None, timeout=None):
        self.url = url

    def reload(self, wait_until=None, timeout=None):
        self.reloads = getattr(self, "reloads", 0) + 1
        if getattr(self, "restore_card_on_reload", False):
            self.restore_card_on_reload = False
            self.state = "card"
            self.body_text = (
                "Enter your card number to continue with Muse."
            )

    def evaluate(self, script):
        self.evaluated.append(script)
        if "data-muse-card-continue" in script:
            if getattr(self, "card_button", None) is not None:
                return "ok"
            return ""
        if "getBoundingClientRect" in script:
            # Nút avatar ở góc dưới trái thanh bên.
            for button in self.buttons:
                if button.name in ("Account", "Profile"):
                    button.click()
                    return True
            return False
        return None

    def _match_name(self, wanted, name):
        if hasattr(wanted, "search"):
            return bool(wanted.search(name or ""))
        if wanted is None:
            return True
        return (name or "") == wanted

    def get_by_role(self, role, name=None, exact=False):
        found = [
            button
            for button in self.buttons
            if button.kind == "button"
            and self._match_name(name, button.name)
        ]
        return _FakeLocator(found)

    def get_by_text(self, text, exact=False):
        found = []
        if hasattr(text, "search"):
            # Regex (ví dụ _ACCOUNT_EMAIL_RE): trả một phần tử cho
            # mỗi email tìm thấy trong nội dung trang.
            for email in set(text.findall(self.body_text or "")):
                found.append(
                    _FakeElement(self, "text", name=email, text=email)
                )
            return _FakeLocator(found)
        if text in self.body_text:
            found.append(
                _FakeElement(self, "text", name=text, text=text)
            )
        return _FakeLocator(found)

    def get_by_placeholder(self, text, exact=False):
        wanted = text.pattern if hasattr(text, "pattern") else str(text)
        lowered = wanted.lower()
        # Khung chat của giao diện Muse ("Ask Meta AI...").
        if self.state == "chat-ui" and "ask meta ai" in lowered:
            return _FakeLocator(
                [_FakeElement(self, "field", "Ask Meta AI anything...")]
            )
        # Các ô nhập của form thẻ.
        if self.state == "card" and lowered in (
            "card number",
            "mm/yy",
            "cvv",
            "zip code",
        ):
            if lowered not in self.fields:
                self.fields[lowered] = _FakeElement(
                    self, "field", text
                )
            return _FakeLocator([self.fields[lowered]])
        return _FakeLocator([])

    def locator(self, selector):
        if "data-muse-card-continue" in selector:
            button = getattr(self, "card_button", None)
            return _FakeLocator([button] if button is not None else [])
        if selector == "body":
            body = _FakeElement(self, "body", text=self.body_text)
            return _FakeLocator([body])
        if selector in (
            '[role="menu"]',
            '[role="dialog"]',
            '[class*="menu"]',
            '[class*="popover"]',
            '[class*="dropdown"]',
        ):
            if self.menu_open and self.menu_email:
                menu = _FakeElement(
                    self, "menu", text=f"Account\n{self.menu_email}"
                )
                return _FakeLocator([menu])
            return _FakeLocator([])
        return _FakeLocator([])


def _page_with_menu(email):
    page = _FakePage(_FakeContext(_FakeBrowser()), "https://muse.ai/")
    page.buttons.append(_FakeElement(page, "button", name="Account"))
    page.menu_email = email
    page.body_text = "Welcome to Muse."
    return page


def _meta_link_page(shown_email, intended_email=INTENDED):
    """Trang giả màn 'Continue with your Meta Account': thẻ tài
    khoản hiện email shown_email, nút Continue, link 'Create new
    account'."""
    page = _FakePage(_FakeContext(_FakeBrowser()), "https://muse.ai/")
    page.intended_email = intended_email
    page.state = (
        "meta-link-match"
        if shown_email == intended_email
        else "meta-link-mismatch"
    )
    page.body_text = (
        "Continue with your Meta Account\n"
        "We found a Meta Account associated with your info. You "
        "can use Muse with your existing account, or create a new "
        "one instead.\n"
        f"{shown_email}\n"
        "1 profile will have access to Muse\n"
        "Continue\n"
        "Create new account"
    )
    button = _FakeElement(page, "button", name="Continue")
    page.buttons.append(button)
    page.card_button = button
    return page


# --- 1. Trình duyệt sạch mỗi lần mở -------------------------------

def test_open_clean_page_uses_fresh_context():
    browser = _FakeBrowser()
    context, page = browser_helper._open_clean_page(browser)
    assert len(browser.contexts) == 1, "phải tạo context mới"
    assert browser.new_page_calls == 0, "không được dùng browser.new_page()"
    assert page.context is context, "trang phải thuộc context mới"
    assert context.cleared >= 1, "phải xoá cookie ngay khi mở"
    assert context.cookies() == [], "context mới phải có 0 cookie"
    assert any(
        "localStorage.clear" in script for script in page.evaluated
    ), "phải xoá bộ nhớ trang"
    assert any(
        "sessionStorage.clear" in script for script in page.evaluated
    ), "phải xoá sessionStorage"


def test_open_clean_page_same_proxy_still_clean():
    # Cùng một proxy (cùng một browser) nhưng mở lần 2 vẫn sạch:
    # context thứ 2 không thấy cookie của context thứ 1.
    browser = _FakeBrowser()
    context1, _page1 = browser_helper._open_clean_page(browser)
    context1.cookie_list.append({"name": "session", "value": "old-account"})
    context2, page2 = browser_helper._open_clean_page(browser)
    assert context2 is not context1, "mỗi lần mở là context khác nhau"
    assert context2.cookies() == [], "context lần 2 không có cookie cũ"
    assert context2.cleared >= 1, "lần 2 vẫn phải xoá cookie"
    assert page2.context is context2


def test_clean_session_clears_seeded_cookies():
    browser = _FakeBrowser()
    context = browser.new_context()
    context.cookie_list = [
        {"name": "a", "value": "1"},
        {"name": "b", "value": "2"},
        {"name": "c", "value": "3"},
    ]
    page = context.new_page()
    browser_helper._ensure_clean_login_session(page)
    assert context.cookies() == [], "phải xoá hết cookie đã có"
    assert context.cleared == 1
    assert any(
        "localStorage.clear" in script for script in page.evaluated
    )


# --- 2. Hộp thư OTP phải khớp email -------------------------------

def test_mailbox_file_matches():
    (TMP / "mailbox.txt").write_text(
        json.dumps(
            {
                "provider": "tempmail.plus",
                "email": INTENDED,
                "secret": "",
            }
        ),
        encoding="utf-8",
    )
    assert browser_helper._mailbox_file_matches(INTENDED) is True
    assert browser_helper._mailbox_file_matches(OTHER) is False


def test_mailbox_file_missing_or_bad():
    try:
        (TMP / "mailbox.txt").unlink()
    except OSError:
        pass
    assert browser_helper._mailbox_file_matches(INTENDED) is False
    (TMP / "mailbox.txt").write_text("không phải json", encoding="utf-8")
    assert browser_helper._mailbox_file_matches(INTENDED) is False
    (TMP / "mailbox.txt").write_text(
        json.dumps(
            {"provider": "tempmail.plus", "email": OTHER, "secret": ""}
        ),
        encoding="utf-8",
    )
    assert browser_helper._mailbox_file_matches(INTENDED) is False


# --- 3. Xác minh tài khoản ở giao diện ------------------------------

def test_verify_matching_account():
    page = _page_with_menu(INTENDED)
    browser_helper._verify_muse_account(page, INTENDED)


def test_verify_mismatched_account_logs_out_and_fails():
    page = _page_with_menu(OTHER)
    page.context.cookie_list.append({"name": "session", "value": "old"})
    try:
        browser_helper._verify_muse_account(page, INTENDED)
    except RuntimeError as error:
        message = str(error)
    else:
        raise AssertionError("tài khoản khác phải lỗi rõ")
    assert "không khớp" in message, "lỗi phải nói rõ không khớp"
    assert OTHER not in message, "không in email tìm thấy ra log"
    assert INTENDED not in message or "email cần đăng nhập" in message
    assert page.context.cookies() == [], "phải xoá phiên tài khoản nhầm"
    assert page.context.cleared >= 1


def test_verify_no_email_readable_continues():
    page = _page_with_menu(None)
    browser_helper._verify_muse_account(page, INTENDED)


# --- 4. Màn chọn tài khoản ------------------------------------------

def test_chooser_clicks_intended_account():
    page = _FakePage(_FakeContext(_FakeBrowser()), "https://muse.ai/")
    page.body_text = (
        "Continue as\n" + INTENDED + "\nNot you? Switch account"
    )
    assert (
        browser_helper._wait_for_profile_or_card(page, INTENDED, 2000)
        == "card"
    ), "thấy đúng email thì bấm vào rồi đi tiếp tới bước thẻ"
    assert INTENDED in page.clicked


def test_chooser_without_intended_email_returns_chooser():
    page = _FakePage(_FakeContext(_FakeBrowser()), "https://muse.ai/")
    page.body_text = (
        "Continue as\n" + OTHER + "\nNot you? Switch account"
    )
    assert (
        browser_helper._wait_for_profile_or_card(page, INTENDED, 2000)
        == "chooser"
    ), "không thấy email cần đăng nhập thì dừng, không bấm nhầm"
    assert page.clicked == [], "không được bấm gì trên màn chọn sai"


def test_no_chooser_normal_flow_unaffected():
    page = _FakePage(_FakeContext(_FakeBrowser()), "https://muse.ai/")
    page.body_text = "Finish creating your account — your birthday"
    assert (
        browser_helper._wait_for_profile_or_card(page, INTENDED, 2000)
        == "profile"
    )
    page.body_text = "Enter your card number to continue."
    assert (
        browser_helper._wait_for_profile_or_card(page, INTENDED, 2000)
        == "card"
    )


# --- 5. Màn "Continue with your Meta Account" ----------------------
# (sư huynh gửi ảnh 2026-11-14: màn này hiện ra dù cookie đã
# sạch — Meta liên kết tài khoản phía server nên phải so khớp
# email trên thẻ trước khi bấm Continue.)

def test_link_screen_detects_shown_email():
    page = _meta_link_page(OTHER)
    assert browser_helper._meta_account_link_screen(page) == OTHER
    page = _meta_link_page(INTENDED)
    assert browser_helper._meta_account_link_screen(page) == INTENDED
    normal = _FakePage(_FakeContext(_FakeBrowser()), "https://muse.ai/")
    normal.body_text = "Finish creating your account — your birthday"
    assert browser_helper._meta_account_link_screen(normal) is None


def test_login_cookie_requires_hatch_sess():
    """Cookie lúc đang đăng nhập không có hatch_sess thì chưa dùng được."""
    pending = (
        "datr=1; _fbp=2; fr=3; hatch_native_auth_csrf=a; "
        "hatch_native_auth_temp=b; hatch_native_auth_state=c; "
        "hatch_native_auth_device=d"
    )
    assert browser_helper._cookie_has_hatch_sess(pending) is False
    assert browser_helper._cookie_has_hatch_sess("") is False
    assert browser_helper._cookie_has_hatch_sess(
        "datr=1; hatch_sess=abc"
    ) is True
    assert browser_helper._cookie_has_hatch_sess("hatch_sess=abc") is True
    logged_in = (
        "datr=abc; wd=1280x720; dpr=1; "
        "hatch_native_auth_temp=v1; _fbp=fb.1; locale=en_US; "
        "hatch_sess=57325bd8.v1%3Atoken; "
        "theme=%7B%22chatThemeCache%22%3A%7B%7D%7D"
    )
    assert browser_helper._cookie_has_hatch_sess(logged_in) is True


def test_blocked_request_relogin_asks_for_new_browser():
    """Đăng nhập lại lấy cookie: lỗi này thì bỏ trình duyệt hiện tại."""
    page = _FakePage(_FakeContext(_FakeBrowser()), "https://muse.ai/")
    page.body_text = (
        "Log in or create an account\n"
        "We couldn't complete that request. Try again.\n"
        "Continue"
    )
    captured = []
    orig = browser_helper._status
    browser_helper._status = captured.append
    try:
        try:
            browser_helper._reject_blocked_request_browser(page)
        except browser_helper._NeedFreshBrowser as error:
            message = str(error)
        else:
            raise AssertionError("phải đóng trình duyệt khi Muse từ chối")
    finally:
        browser_helper._status = orig
    assert any("trình duyệt mới" in line for line in captured), captured
    assert "trình duyệt" in message.lower()
    clean = _FakePage(_FakeContext(_FakeBrowser()), "https://muse.ai/")
    clean.body_text = "Enter the code we sent"
    browser_helper._reject_blocked_request_browser(clean)


def test_blocked_request_drops_proxy_on_muse_login():
    """Form email báo không hoàn tất được yêu cầu thì bỏ proxy,
    không chờ OTP trên IP đó."""
    page = _FakePage(_FakeContext(_FakeBrowser()), "https://muse.ai/")
    page.body_text = (
        "Log in or create an account\n"
        "muse50313da45f@merepost.com\n"
        "We couldn't complete that request. Try again.\n"
        "Continue"
    )
    captured = []
    orig = browser_helper._status
    browser_helper._status = captured.append
    try:
        try:
            browser_helper._reject_blocked_request_proxy(page)
        except RuntimeError as error:
            message = str(error)
        else:
            raise AssertionError("phải bỏ proxy khi Muse từ chối yêu cầu")
    finally:
        browser_helper._status = orig
    assert any(line.startswith("PROXY_BURN") for line in captured), captured
    assert "proxy" in message.lower()
    clean = _FakePage(_FakeContext(_FakeBrowser()), "https://muse.ai/")
    clean.body_text = "Enter the code we sent"
    browser_helper._reject_blocked_request_proxy(clean)


def test_email_too_many_attempts_drops_proxy():
    """Form email 'Too many attempts. Try again later.' thì bỏ proxy.

    Câu 'Too many failed attempts' ở bước thẻ không được đụng:
    bước thẻ vẫn tải lại và điền lại đúng thẻ.
    """
    page = _FakePage(_FakeContext(_FakeBrowser()), "https://muse.ai/")
    page.body_text = (
        "Log in or create an account\n"
        "muse26c95893ca@fexbox.org\n"
        "Too many attempts. Try again later.\n"
        "Continue"
    )
    captured = []
    orig = browser_helper._status
    browser_helper._status = captured.append
    try:
        try:
            browser_helper._wait_for_verification_code(
                page, {"provider": "tempmail.plus", "address": "a@b.c"},
                set(), reject_blocked=False,
            )
        except RuntimeError as error:
            message = str(error)
        else:
            raise AssertionError("phải bỏ proxy khi form email bị giới hạn")
    finally:
        browser_helper._status = orig
    assert any(line.startswith("PROXY_BURN") for line in captured), captured
    assert "Too many attempts" in " ".join(captured)
    assert "proxy" in message.lower()
    card = _FakePage(_FakeContext(_FakeBrowser()), "https://muse.ai/")
    card.body_text = (
        "Too many failed attempts. Please try again later.\nOK"
    )
    browser_helper._reject_blocked_request_proxy(card)


def test_facebook_page_signals_drop_proxy():
    """Proxy đẩy sang facebook.com thì báo bỏ proxy và dừng,
    không bấm form Log into Facebook."""
    page = _FakePage(
        _FakeContext(_FakeBrowser()),
        "https://www.facebook.com/login/?next=https%3A%2F%2Fwww.facebook.com",
    )
    page.body_text = "Log into Facebook\nEmail or mobile number\nPassword"
    captured = []
    orig = browser_helper._status
    browser_helper._status = captured.append
    try:
        try:
            browser_helper._reject_facebook_proxy(page)
        except RuntimeError as error:
            message = str(error)
        else:
            raise AssertionError("phải dừng khi ra Facebook")
    finally:
        browser_helper._status = orig
    assert any(line.startswith("PROXY_FACEBOOK") for line in captured), captured
    assert "Facebook" in message
    assert page.clicked == []
    muse = _FakePage(_FakeContext(_FakeBrowser()), "https://muse.ai/")
    browser_helper._reject_facebook_proxy(muse)


def test_link_screen_clicks_card_button_not_login_continue():
    """Nút xanh trên thẻ đôi khi chưa có chữ. Continue của form
    email vẫn còn trên trang — chỉ bấm nút trên thẻ, không bấm
    Continue của form (nút đó kéo về màn đăng nhập)."""
    page = _meta_link_page(INTENDED)
    login_continue = _FakeElement(page, "button", name="Continue")
    card = _FakeElement(page, "button", name="")
    page.buttons = [login_continue]
    page.card_button = card
    assert browser_helper._click_linked_meta_continue(page) is True
    assert page.clicked == ["button"], page.clicked


def test_link_screen_continue_is_clicked_only_once():
    page = _meta_link_page(INTENDED)
    page.hold_link_screen = True
    browser_helper._wait_for_profile_or_card(page, INTENDED, 1500)
    assert page.clicked.count("Continue") == 1, page.clicked


def test_wait_returns_checkpoint_before_card():
    page = _FakePage(_FakeContext(_FakeBrowser()), "https://auth.meta.com/checkpoints/")
    page.body_text = "Confirm you're human to use your account"
    page.get_by_text = lambda text, exact=False: (
        _FakeLocator([_FakeElement(page, "text", name=text, text=text)])
        if text == "Confirm you're human" or (
            hasattr(text, "search") and text.search(page.body_text)
        )
        else _FakeLocator([])
    )
    assert (
        browser_helper._wait_for_profile_or_card(page, INTENDED, 2000)
        == "checkpoint"
    )


def test_link_screen_match_clicks_continue():
    page = _meta_link_page(INTENDED)
    captured = []
    orig = browser_helper._status
    browser_helper._status = captured.append
    try:
        result = browser_helper._handle_meta_account_link(
            page, INTENDED, allow_create_new=False
        )
    finally:
        browser_helper._status = orig
    assert result == "continue", result
    assert "Continue" in page.clicked, page.clicked
    assert any("khớp" in s for s in captured), captured
    assert page.state == "card", "bấm Continue đúng email thì đi tiếp"


def test_link_screen_mismatch_login_stops_without_click():
    page = _meta_link_page(OTHER)
    captured = []
    orig = browser_helper._status
    browser_helper._status = captured.append
    try:
        result = browser_helper._handle_meta_account_link(
            page, INTENDED, allow_create_new=False
        )
    finally:
        browser_helper._status = orig
    assert result == "stop", result
    assert "Continue" not in page.clicked, (
        "email không khớp thì TUYỆT ĐỐI không bấm Continue, "
        f"đã bấm: {page.clicked}"
    )
    assert any("KHÔNG khớp" in s for s in captured), captured


def test_link_screen_mismatch_register_clicks_create_new():
    page = _meta_link_page(OTHER)
    result = browser_helper._handle_meta_account_link(
        page, INTENDED, allow_create_new=True
    )
    assert result == "create-new", result
    assert "Create new account" in page.clicked, page.clicked
    assert "Continue" not in page.clicked, page.clicked
    assert page.state == "profile", "tạo tài khoản mới thì về form hồ sơ"


def test_wait_for_profile_or_card_stops_on_mismatched_link():
    page = _meta_link_page(OTHER)
    assert (
        browser_helper._wait_for_profile_or_card(page, INTENDED, 2000)
        == "chooser"
    ), "màn liên kết sai email thì dừng như màn chọn sai"
    assert page.clicked == [], "không được bấm gì trên màn liên kết sai"


def test_wait_for_profile_or_card_continues_on_matched_link():
    page = _meta_link_page(INTENDED)
    assert (
        browser_helper._wait_for_profile_or_card(page, INTENDED, 2000)
        == "card"
    ), "màn liên kết đúng email thì bấm Continue rồi đi tiếp"
    assert "Continue" in page.clicked, page.clicked


def test_confirm_final_step_stops_on_mismatched_link():
    page = _meta_link_page(OTHER)
    try:
        browser_helper._confirm_final_step(
            page, timeout_seconds=5, email=INTENDED,
            allow_create_new=False,
        )
    except RuntimeError as error:
        message = str(error)
    else:
        raise AssertionError(
            "màn liên kết sai email sau bước thẻ phải lỗi rõ"
        )
    assert "không khớp" in message, message
    assert "Continue" not in page.clicked, (
        "không được bấm Continue vào tài khoản cũ, "
        f"đã bấm: {page.clicked}"
    )


def test_confirm_final_step_continues_on_matched_link():
    page = _meta_link_page(INTENDED)
    browser_helper._confirm_final_step(
        page, timeout_seconds=5, email=INTENDED, allow_create_new=False
    )
    assert "Continue" in page.clicked, page.clicked
    assert page.state == "card", "đúng email thì bấm Continue rồi đi tiếp"


# --- 6. Bấm Continue ở màn liên kết mà tài khoản ĐÃ đăng nhập
# sẵn: vào thẳng giao diện Muse — KHÔNG đăng nhập lại. ------
# (sư huynh 2026-11-14: "tài khoản đó như vậy là đã được đăng
# nhập và lấy cookie rồi, bỏ đi còn đăng nhập lại làm gì")

def _meta_link_page_to_chat_ui():
    """Màn liên kết mà bấm Continue thì vào thẳng giao diện chat
    (tài khoản đã đăng nhập sẵn)."""
    page = _meta_link_page(INTENDED)
    page.after_link_continue = "chat-ui"
    return page


def test_wait_for_profile_or_card_link_match_to_chat_ui_logged_in():
    page = _meta_link_page_to_chat_ui()
    assert (
        browser_helper._wait_for_profile_or_card(page, INTENDED, 2000)
        == "logged-in"
    ), (
        "bấm Continue ở màn liên kết mà tài khoản đã vào được "
        "Muse thì phải báo đã đăng nhập xong, không chờ form "
        "hồ sơ hay form thẻ nữa"
    )
    assert "Continue" in page.clicked, page.clicked
    assert page.state == "chat-ui"


def test_ensure_profile_form_link_match_to_chat_ui_logged_in():
    page = _meta_link_page_to_chat_ui()
    assert (
        browser_helper._ensure_profile_form(page, email=INTENDED)
        == "logged-in"
    ), (
        "đăng ký mà bấm Continue ở màn liên kết rồi vào thẳng "
        "giao diện thì không điền hồ sơ nữa"
    )
    assert "Continue" in page.clicked, page.clicked


def test_confirm_final_step_link_match_to_chat_ui_returns():
    page = _meta_link_page_to_chat_ui()
    browser_helper._confirm_final_step(
        page, timeout_seconds=5, email=INTENDED, allow_create_new=False,
    )
    assert page.clicked == ["Continue"], (
        "vào thẳng giao diện thì chỉ bấm Continue ở màn liên "
        f"kết, không bấm thêm nút nào: {page.clicked}"
    )


def test_do_card_step_if_present_chat_ui_skips():
    page = _FakePage(_FakeContext(_FakeBrowser()), "https://muse.ai/")
    page.state = "chat-ui"
    page.body_text = "Ask Meta AI anything...\nWelcome to Muse."
    assert (
        browser_helper._do_card_step_if_present(page, None) is False
    ), (
        "đã vào giao diện chat thì không có form thẻ — bỏ qua "
        "bước thẻ, không chờ, không điền"
    )
    assert page.clicked == [], "không được bấm gì khi bỏ qua bước thẻ"


def test_do_card_step_if_present_card_form_fills_saved_card():
    # Phiên chưa vào giao diện mà form thẻ hiện: bước thẻ VẪN
    # chạy như thường — chỉ bỏ qua khi đã vào được Muse.
    storage._import_bulk_cards("4111111111111111|12/28|123")
    page = _FakePage(_FakeContext(_FakeBrowser()), "https://muse.ai/")
    page.state = "card"
    page.body_text = "Enter your card number to continue with Muse."
    page.buttons.append(_FakeElement(page, "button", name="Confirm"))
    assert (
        browser_helper._do_card_step_if_present(page, None) is True
    ), "form thẻ hiện thì phải điền thẻ đã lưu và hoàn tất bước thẻ"
    assert "Confirm" in page.clicked, page.clicked
    assert page.state == "muse", "điền thẻ xong phải vào giao diện"


def test_relogin_waitlist_drops_account_without_clicking():
    # muse.ai/access "Join the waitlist": bỏ mail, không bấm nút.
    page = _FakePage(_FakeContext(_FakeBrowser()), "https://muse.ai/access")
    page.body_text = (
        "Join the waitlist for Muse\n"
        "Muse isn't available in your country or region yet.\n"
        "Join waitlist"
    )
    page.buttons.append(_FakeElement(page, "button", name="Join waitlist"))
    try:
        browser_helper._open_signup_and_submit_email(
            page, "muse111@flubai.com", abort_on_waitlist=True
        )
    except RuntimeError as error:
        assert "waitlist" in str(error).lower(), error
    else:
        raise AssertionError("waitlist phải bỏ tài khoản")
    assert page.clicked == [], page.clicked


def test_too_many_attempts_reloads_and_fills_same_card():
    # "Too many attempts" khi điền lần 2: bấm OK, tải lại,
    # điền lại đúng thẻ đó. Không đánh dấu full.
    storage._import_bulk_cards("4111111111111111|12/28|123|10001|Thẻ A")
    entry = next(
        item for item in storage._load_vault() if item["label"] == "Thẻ A"
    )
    page = _FakePage(_FakeContext(_FakeBrowser()), "https://muse.ai/")
    page.state = "card"
    page.body_text = "Enter your card number to continue with Muse."
    page.buttons.append(_FakeElement(page, "button", name="Confirm"))
    page.too_many_attempts_once = True
    page.reloads = 0
    assert browser_helper._do_card_step(page, entry["id"]) is True
    assert page.reloads == 1, page.reloads
    assert page.clicked.count("OK") == 1, page.clicked
    assert page.clicked.count("Confirm") == 2, page.clicked
    saved = next(
        item for item in storage._load_vault() if item["id"] == entry["id"]
    )
    assert saved["used"] == 0 and saved["limit"] == 0, saved


def test_already_full_card_is_not_typed_again():
    # Thẻ đã full (luồng khác vừa đánh dấu): không gõ lại số
    # thẻ đó, điền thẻ kế rồi bấm Confirm để đi tiếp.
    storage._import_bulk_cards(
        "4111111111111111|12/28|123|10001|Thẻ A\n"
        "4222222222222222|11/29|456|10001|Thẻ B"
    )
    first = next(
        entry for entry in storage._load_vault() if entry["label"] == "Thẻ A"
    )
    storage._mark_vault_card_full(first["id"])
    page = _FakePage(_FakeContext(_FakeBrowser()), "https://muse.ai/")
    page.state = "card"
    page.body_text = "Enter your card number to continue with Muse."
    page.buttons.append(_FakeElement(page, "button", name="Confirm"))
    page.filled = []
    assert browser_helper._do_card_step(page, first["id"]) is True
    assert "4111111111111111" not in page.filled, page.filled
    assert "4222222222222222" in page.filled, page.filled
    assert "Confirm" in page.clicked, page.clicked


def test_card_blocked_dialog_marks_full_and_fills_next():
    # "Unable to Add Card" / too many ad accounts: bấm OK, đánh
    # dấu thẻ đó full, điền thẻ kế. Không bấm Learn more.
    storage._import_bulk_cards(
        "4111111111111111|12/28|123|10001|Thẻ A\n"
        "4222222222222222|11/29|456|10001|Thẻ B"
    )
    first = next(
        entry for entry in storage._load_vault() if entry["label"] == "Thẻ A"
    )
    page = _FakePage(_FakeContext(_FakeBrowser()), "https://muse.ai/")
    page.state = "card"
    page.body_text = "Enter your card number to continue with Muse."
    page.buttons.append(_FakeElement(page, "button", name="Confirm"))
    page.reject_next_confirm = True
    assert browser_helper._do_card_step(page, first["id"]) is True
    assert page.clicked.count("OK") == 1, page.clicked
    assert "Learn more" not in page.clicked, page.clicked
    assert page.clicked.count("Confirm") == 2, page.clicked
    assert page.fields["card number"].value == "4222222222222222"
    full = next(
        entry for entry in storage._load_vault() if entry["label"] == "Thẻ A"
    )
    other = next(
        entry for entry in storage._load_vault() if entry["label"] == "Thẻ B"
    )
    assert full["limit"] > 0 and full["used"] >= full["limit"], full
    assert other["used"] == 0, other


TESTS = [
    test_open_clean_page_uses_fresh_context,
    test_open_clean_page_same_proxy_still_clean,
    test_clean_session_clears_seeded_cookies,
    test_mailbox_file_matches,
    test_mailbox_file_missing_or_bad,
    test_verify_matching_account,
    test_verify_mismatched_account_logs_out_and_fails,
    test_verify_no_email_readable_continues,
    test_chooser_clicks_intended_account,
    test_chooser_without_intended_email_returns_chooser,
    test_no_chooser_normal_flow_unaffected,
    test_link_screen_detects_shown_email,
    test_login_cookie_requires_hatch_sess,
    test_blocked_request_relogin_asks_for_new_browser,
    test_blocked_request_drops_proxy_on_muse_login,
    test_email_too_many_attempts_drops_proxy,
    test_facebook_page_signals_drop_proxy,
    test_link_screen_clicks_card_button_not_login_continue,
    test_link_screen_continue_is_clicked_only_once,
    test_wait_returns_checkpoint_before_card,
    test_link_screen_match_clicks_continue,
    test_link_screen_mismatch_login_stops_without_click,
    test_link_screen_mismatch_register_clicks_create_new,
    test_wait_for_profile_or_card_stops_on_mismatched_link,
    test_wait_for_profile_or_card_continues_on_matched_link,
    test_confirm_final_step_stops_on_mismatched_link,
    test_confirm_final_step_continues_on_matched_link,
    test_wait_for_profile_or_card_link_match_to_chat_ui_logged_in,
    test_ensure_profile_form_link_match_to_chat_ui_logged_in,
    test_confirm_final_step_link_match_to_chat_ui_returns,
    test_do_card_step_if_present_chat_ui_skips,
    test_do_card_step_if_present_card_form_fills_saved_card,
    test_relogin_waitlist_drops_account_without_clicking,
    test_too_many_attempts_reloads_and_fills_same_card,
    test_already_full_card_is_not_typed_again,
    test_card_blocked_dialog_marks_full_and_fills_next,
]


def main():
    passed = 0
    failed = 0
    for test in TESTS:
        try:
            (TMP / "mailbox.txt").unlink()
        except OSError:
            pass
        # Kho thẻ của test trước không được rò sang test sau.
        for vault_name in ("cards.json", "vault.json"):
            try:
                (TMP / vault_name).unlink()
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
