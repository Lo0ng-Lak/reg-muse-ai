import html
import json
import re
import secrets
import sys
import time
import tkinter as tk
import urllib.error
import urllib.request
from datetime import date
from tkinter import messagebox
from html.parser import HTMLParser

import keyring


# Ép stdout/stderr sang UTF-8. Trên Windows, print() tiếng Việt vào pipe
# mặc định dùng bảng mã console (charmap) và sập ngay ở dòng trạng thái
# đầu tiên — trước cả khi kịp điền email vào form.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


SIGNUP_URL = "https://muse.ai/"
API_BASE = "https://api.mail.tm"
MAIL_SERVICE = "MuseAIHelper:mail.tm"
CODE_TIMEOUT_SECONDS = 180
PAGE_TIMEOUT_MS = 45000
BUTTON_TIMEOUT_MS = 90000
UI_SETTLE_MS = 1500


def _status(message):
    print(message, flush=True)


def _classify_after_profile(page, timeout_seconds=45):
    """Phân loại màn hình xuất hiện sau khi gửi form hồ sơ.

    Trả về:
    - "waitlist": màn hình "Join the waitlist" — Muse chưa khả dụng ở
      khu vực của IP hiện tại, đăng ký coi như THẤT BẠI.
    - "card": màn hình nhập thẻ — đăng ký đi đúng hướng; phải hoàn
      tất bước nhập thẻ nữa mới tính là THÀNH CÔNG.
    - "unknown": hết thời gian chờ mà không thấy hai màn hình trên.
    """
    waitlist_markers = ("join the waitlist", "available in your country")
    card_markers = (
        "card number",
        "card details",
        "enter your card",
        "add a card",
        "payment method",
    )
    deadline = time.monotonic() + timeout_seconds
    while time.monotonic() < deadline:
        if page.is_closed():
            return "unknown"
        try:
            text = page.locator("body").inner_text(timeout=3000).lower()
        except Exception:
            text = ""
        if any(marker in text for marker in waitlist_markers):
            return "waitlist"
        if any(marker in text for marker in card_markers):
            return "card"
        page.wait_for_timeout(2000)
    return "unknown"


def _wait_for_card_step_done(page, timeout_seconds=240):
    """Chờ người dùng tự nhập thẻ + ZIP US trong cửa sổ Chromium.

    Tool không nhận và không tự điền thông tin thẻ. Bước thẻ coi như
    xong khi các dấu hiệu của form thẻ biến mất khỏi trang qua hai lần
    kiểm tra liên tiếp (tránh tính nhầm lúc trang đang chuyển cảnh).
    Trả về "done", "waitlist", "closed" hoặc "timeout".
    """
    waitlist_markers = ("join the waitlist", "available in your country")
    card_markers = (
        "card number",
        "card details",
        "enter your card",
        "add a card",
        "payment method",
    )
    deadline = time.monotonic() + timeout_seconds
    clear_polls = 0
    while time.monotonic() < deadline:
        if page.is_closed():
            return "closed"
        try:
            text = page.locator("body").inner_text(timeout=3000).lower()
        except Exception:
            text = ""
        if any(marker in text for marker in waitlist_markers):
            return "waitlist"
        if text and not any(marker in text for marker in card_markers):
            clear_polls += 1
            if clear_polls >= 2:
                return "done"
        else:
            clear_polls = 0
        page.wait_for_timeout(2000)
    return "timeout"


def _wait_for_profile_form(page, timeout_ms):
    """Chờ form hồ sơ xuất hiện — form này có HAI biến thể:

    - "Finish creating your account": có cả ô tên và ngày sinh.
    - "What's your birthday?": chỉ hỏi ngày sinh (ví dụ khi qua IP Mỹ).

    Trả về True ngay khi một trong hai biến thể hiện ra.
    """
    deadline = time.monotonic() + timeout_ms / 1000
    while time.monotonic() < deadline:
        if page.is_closed():
            return False
        for name in (
            "Finish creating your account",
            "What's your birthday?",
        ):
            heading = page.get_by_role("heading", name=name)
            try:
                if heading.count() and heading.first.is_visible():
                    return True
            except Exception:
                pass
        page.wait_for_timeout(500)
    return False


def main():
    if len(sys.argv) not in (2, 3):
        raise ValueError("Thiếu địa chỉ email cần điền.")
    email = sys.argv[1]
    auto = len(sys.argv) == 3 and sys.argv[2] == "--auto"
    mailbox_password = keyring.get_password(MAIL_SERVICE, email)
    if not mailbox_password:
        detail = (
            "Không có mật khẩu mail.tm trong Windows Credential Manager cho "
            f"{email}. Hãy tạo lại hộp thư bằng ứng dụng."
        )
        if auto:
            _status(f"RESULT: FAIL {detail}")
            sys.exit(1)
        _show_error("Không tìm thấy hộp thư", detail)
        return

    try:
        from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
        from playwright.sync_api import sync_playwright
    except ImportError as error:
        detail = (
            "Cài các gói phụ thuộc và Chromium bằng lệnh:\n"
            "`py -m pip install -r requirements.txt`\n"
            "`py -m playwright install chromium`\n\n"
            f"{error}"
        )
        if auto:
            _status(f"RESULT: FAIL Thiếu Playwright. {error}")
            sys.exit(1)
        _show_error("Thiếu Playwright", detail)
        return

    playwright = sync_playwright().start()
    browser = None
    page = None
    try:
        token = _get_mail_token(email, mailbox_password)
        browser = playwright.chromium.launch(headless=False)
        page = browser.new_page()
        page.set_default_timeout(PAGE_TIMEOUT_MS)
        page.goto(
            SIGNUP_URL,
            wait_until="domcontentloaded",
            timeout=PAGE_TIMEOUT_MS,
        )
        page.wait_for_timeout(UI_SETTLE_MS)
        _status(f"Đã mở trang Muse.ai, đang điền email {email}...")

        email_field = page.get_by_role(
            "textbox", name="Mobile number or email"
        )
        try:
            email_field.wait_for(state="visible", timeout=5000)
        except PlaywrightTimeoutError:
            login_button = page.get_by_role("button", name="Log in")
            _click_when_ready(page, login_button, "Log in")
            email_field.wait_for(state="visible", timeout=PAGE_TIMEOUT_MS)

        continue_button = page.get_by_role(
            "button", name="Continue", exact=True
        )
        continue_button.wait_for(state="visible", timeout=BUTTON_TIMEOUT_MS)
        email_field.fill(email)
        email_field.wait_for(state="visible", timeout=BUTTON_TIMEOUT_MS)
        if email_field.input_value() != email:
            raise RuntimeError(
                "Email chưa được giữ trong form Muse.ai; chưa bấm Continue."
            )
        page.wait_for_timeout(UI_SETTLE_MS)
        existing_ids = _get_message_ids(token)
        _click_when_ready(page, continue_button, "Continue")
        _status("Đã gửi email, đang chờ mã xác minh từ mail.tm...")

        code = _wait_for_verification_code(page, token, existing_ids)
        _status("Đã nhận mã xác minh, đang điền mã...")
        _fill_verification_code(page, code)
        page.wait_for_timeout(UI_SETTLE_MS)
        # Form mã của Muse.ai TỰ GỬI ngay khi nhập đủ số: trang tự nhảy
        # sang bước hồ sơ mà không cần bấm Confirm. TUYỆT ĐỐI không bấm
        # nút "Confirm" nào khi form hồ sơ đã hiện mà chưa điền: nút
        # Confirm của form hồ sơ luôn bấm được, bấm ngay lúc này là gửi
        # ngày sinh MẶC ĐỊNH (ngày hôm nay — tuổi 0) và bị chặn vì dưới
        # 18 tuổi. Vì vậy: chờ form hồ sơ (cả hai biến thể) tự hiện
        # trước; chỉ khi nó không hiện mới bấm Confirm của form mã rồi
        # chờ tiếp.
        if not _wait_for_profile_form(page, 15000):
            confirm_button = page.get_by_role(
                "button", name="Confirm", exact=True
            )
            _click_when_ready(page, confirm_button, "Confirm")
            if not _wait_for_profile_form(page, 15000):
                raise RuntimeError(
                    "Không thấy form hồ sơ sau khi xác nhận mã xác minh."
                )
        _status("Mã đã xác nhận, đang điền hồ sơ...")
        _fill_profile(page)
        finish_button = page.get_by_role(
            "button", name="Confirm", exact=True
        )
        _click_when_ready(page, finish_button, "Confirm")
        if auto:
            _status("Đã gửi hồ sơ, đang chờ màn hình tiếp theo...")
            outcome = _classify_after_profile(page)
            if outcome == "waitlist":
                raise RuntimeError(
                    "Ra màn hình 'Join the waitlist' — Muse chưa khả "
                    "dụng ở khu vực của IP hiện tại. Hãy chọn location "
                    "VPN ở nước được hỗ trợ (ví dụ Mỹ) rồi chạy lại."
                )
            if outcome != "card":
                raise RuntimeError(
                    "Sau hồ sơ không ra bước nhập thẻ, cũng không ra "
                    "waitlist — bước tiếp theo chưa rõ, hãy chụp màn "
                    "hình gửi lại để bổ sung vào tool."
                )
            _status(
                "Đã tới bước nhập thẻ — hãy nhập thẻ và ZIP US trong "
                "cửa sổ Chromium. Tool chờ tối đa 4 phút; thẻ được "
                "chấp nhận thì mới lưu tài khoản."
            )
            card_outcome = _wait_for_card_step_done(page)
            if card_outcome != "done":
                raise RuntimeError(
                    "Chưa hoàn tất bước nhập thẻ "
                    f"(kết quả chờ: {card_outcome}) — tài khoản "
                    "không được lưu."
                )
            _status("Đã qua bước nhập thẻ — tài khoản được lưu.")
            _status("RESULT: SUCCESS")
        else:
            page.wait_for_event("close")
    except PlaywrightTimeoutError as error:
        detail = (
            "Trang Muse.ai chưa phản hồi đúng thời gian, hoặc form đã "
            f"thay đổi. {error}"
        )
        if auto:
            _status(f"RESULT: FAIL {detail}")
            sys.exit(1)
        _show_error(
            "Không hoàn tất được bước trình duyệt",
            detail
            + "\n\nCửa sổ Chromium vẫn được giữ nguyên để bạn kiểm tra/thử lại.",
        )
        _wait_for_browser_close(page)
    except Exception as error:
        if auto:
            _status(f"RESULT: FAIL {error}")
            sys.exit(1)
        _show_error(
            "Không hoàn tất được bước trình duyệt",
            "Cửa sổ Chromium vẫn được giữ nguyên để bạn kiểm tra/thử lại.\n\n"
            f"{error}",
        )
        _wait_for_browser_close(page)
    finally:
        if browser and browser.is_connected():
            browser.close()
        playwright.stop()


def _random_birth_date():
    """Ngày sinh ngẫu nhiên với tuổi luôn trên 18.

    Chọn tuổi mục tiêu 20–45 rồi lùi về năm sinh tương ứng; dù sinh
    nhật của năm hiện tại chưa tới thì tuổi thật vẫn từ 19 trở lên.
    Ngày luôn lấy trong 1–28 để tháng nào cũng hợp lệ.
    """
    today = date.today()
    age = secrets.randbelow(26) + 20
    return date(
        today.year - age,
        secrets.randbelow(12) + 1,
        secrets.randbelow(28) + 1,
    )


def _fill_profile(page):
    if not _wait_for_profile_form(page, BUTTON_TIMEOUT_MS):
        raise RuntimeError(
            "Không thấy form hồ sơ (cả hai biến thể tên/ngày sinh)."
        )
    first_name = page.get_by_placeholder("First name", exact=True)
    last_name = page.get_by_placeholder("Last name", exact=True)
    try:
        first_name.wait_for(state="visible", timeout=5000)
        last_name.wait_for(state="visible", timeout=5000)
    except Exception:
        # Một số khu vực (ví dụ IP Mỹ) form không hỏi tên — bỏ qua
        # phần tên, chỉ điền ngày sinh.
        _status("Form không yêu cầu tên — bỏ qua, chỉ điền ngày sinh.")
    else:
        first_name.fill("Long")
        last_name.fill("long")
        _wait_for_input_value(page, first_name, "Long", "tên")
        _wait_for_input_value(page, last_name, "long", "họ")
        _status("Đã điền tên.")
    birth_date = _random_birth_date()
    _select_birth_date(page, birth_date)
    _status(
        "Đã chọn ngày sinh ngẫu nhiên: "
        f"{birth_date.day:02d}/{birth_date.month:02d}/{birth_date.year} "
        "(tuổi trên 18)."
    )


def _wait_for_input_value(page, field, expected, label):
    deadline = time.monotonic() + 10
    while field.input_value() != expected:
        if time.monotonic() >= deadline:
            raise RuntimeError(f"Không xác nhận được {label} trong form Muse.ai.")
        page.wait_for_timeout(250)


def _select_birth_date(page, birth_date):
    month_names = (
        "January",
        "February",
        "March",
        "April",
        "May",
        "June",
        "July",
        "August",
        "September",
        "October",
        "November",
        "December",
    )
    targets = (
        ("Month", month_names[birth_date.month - 1], str(birth_date.month)),
        ("Day", str(birth_date.day), str(birth_date.day)),
        ("Year", str(birth_date.year), str(birth_date.year)),
    )
    for label, value, option_value in targets:
        native_select = page.locator(
            f'select[autocomplete="bday-{label.lower()}"]'
        )
        control = page.get_by_role(
            "combobox", name=re.compile(f"^{label}$", re.I)
        )
        if control.count():
            control = control.first
            control.wait_for(state="visible", timeout=BUTTON_TIMEOUT_MS)
            tag_name = control.evaluate("(element) => element.tagName.toLowerCase()")
            if tag_name == "select":
                control.select_option(value=option_value, timeout=BUTTON_TIMEOUT_MS)
                _wait_for_birthday_value(
                    page, control, option_value, value, label
                )
                continue
            _click_when_ready(page, control, label)
        else:
            control = page.get_by_role(
                "button", name=re.compile(label, re.I)
            ).first
            _click_when_ready(page, control, label)

        option = page.get_by_role("option", name=value, exact=True)
        try:
            from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

            option.wait_for(state="visible", timeout=1000)
            option.click(timeout=BUTTON_TIMEOUT_MS)
        except PlaywrightTimeoutError:
            if not native_select.count():
                raise RuntimeError(
                    f"Không tìm thấy mục {value} trong dropdown {label.lower()}."
                )
            native_select.select_option(
                value=option_value, timeout=BUTTON_TIMEOUT_MS
            )
        _wait_for_birthday_value(page, control, option_value, value, label)


def _wait_for_birthday_value(page, control, option_value, value, label):
    deadline = time.monotonic() + 10
    while True:
        if _birthday_value_matches(control, option_value, value):
            return
        if time.monotonic() >= deadline:
            raise RuntimeError(
                f"Không xác nhận được {value} cho {label.lower()} sinh."
            )
        page.wait_for_timeout(250)


def _birthday_value_matches(control, option_value, value):
    tag_name = control.evaluate("(element) => element.tagName.toLowerCase()")
    if tag_name == "select":
        selected_value = control.input_value()
        displayed_value = control.locator("option:checked").inner_text().strip()
    else:
        selected_text = control.locator('[data-slot="select-value"]')
        displayed_value = (
            selected_text.inner_text().strip()
            if selected_text.count()
            else control.inner_text().strip()
        )
        selected_value = control.get_attribute("value")
    return (
        (selected_value is None or selected_value == option_value)
        and displayed_value == value
    )


def _wait_for_browser_close(page):
    if page and not page.is_closed():
        page.wait_for_event("close")


def _click_when_ready(page, locator, label):
    locator.wait_for(state="visible", timeout=BUTTON_TIMEOUT_MS)
    deadline = time.monotonic() + BUTTON_TIMEOUT_MS / 1000
    while not locator.is_enabled():
        if time.monotonic() >= deadline:
            raise RuntimeError(
                f"Nút {label} vẫn chưa bật sau "
                f"{BUTTON_TIMEOUT_MS // 1000} giây."
            )
        page.wait_for_timeout(250)
    page.wait_for_timeout(UI_SETTLE_MS)
    locator.click(timeout=BUTTON_TIMEOUT_MS)


def _mail_request(method, path, payload=None, token=None):
    headers = {"Accept": "application/ld+json"}
    data = None
    if payload is not None:
        headers["Content-Type"] = "application/json"
        data = json.dumps(payload).encode("utf-8")
    if token:
        headers["Authorization"] = f"Bearer {token}"

    request = urllib.request.Request(
        f"{API_BASE}{path}", data=data, headers=headers, method=method
    )
    try:
        with urllib.request.urlopen(request, timeout=20) as response:
            body = response.read()
    except urllib.error.HTTPError as error:
        detail = error.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"mail.tm trả lỗi HTTP {error.code}: {detail}") from error
    except (urllib.error.URLError, TimeoutError) as error:
        raise RuntimeError(f"Không kết nối được mail.tm: {error}") from error

    return json.loads(body.decode("utf-8")) if body else {}


def _get_mail_token(email, password):
    response = _mail_request(
        "POST", "/token", {"address": email, "password": password}
    )
    token = response.get("token")
    if not token:
        raise RuntimeError("mail.tm không trả về token truy cập.")
    return token


def _get_message_ids(token):
    messages = _mail_request("GET", "/messages?page=1", token=token).get(
        "hydra:member", []
    )
    return {message["id"] for message in messages if message.get("id")}


def _message_text(detail):
    text = detail.get("text")
    if text:
        return text
    html_body = detail.get("html") or []
    if not html_body:
        return detail.get("intro", "")
    parser = _HtmlTextExtractor()
    parser.feed("\n".join(html_body) if isinstance(html_body, list) else html_body)
    return html.unescape("\n".join(parser.parts))


class _HtmlTextExtractor(HTMLParser):
    def __init__(self):
        super().__init__()
        self.parts = []

    def handle_data(self, data):
        value = data.strip()
        if value:
            self.parts.append(value)


def _verification_code_from_new_messages(token, existing_ids):
    messages = _mail_request("GET", "/messages?page=1", token=token).get(
        "hydra:member", []
    )
    for message in messages:
        message_id = message.get("id")
        if not message_id or message_id in existing_ids:
            continue
        detail = _mail_request(
            "GET", f"/messages/{message_id}", token=token
        )
        codes = re.findall(r"(?<!\d)\d{4,8}(?!\d)", _message_text(detail))
        if codes:
            return codes[0]
    return None


def _verification_input(page):
    code_field = page.get_by_role(
        "textbox", name="6-digit security code"
    )
    if code_field.count():
        return code_field

    selectors = (
        'input[autocomplete="one-time-code"]',
        'input[name*="code" i]',
        'input[name*="otp" i]',
        'input[aria-label*="code" i]',
        'input[aria-label*="verification" i]',
        'input[placeholder*="code" i]',
        'input[placeholder*="verification" i]',
        'input[inputmode="numeric"]',
    )
    for selector in selectors:
        fields = page.locator(f"{selector}:visible")
        if fields.count():
            return fields
    return None


def _wait_for_verification_code(page, token, existing_ids):
    deadline = time.monotonic() + CODE_TIMEOUT_SECONDS
    code = None
    while time.monotonic() < deadline:
        if page.is_closed():
            raise RuntimeError("Cửa sổ Chromium đã đóng trước khi xác minh.")
        if code is None:
            code = _verification_code_from_new_messages(token, existing_ids)
        fields = _verification_input(page)
        if code and fields:
            return code
        page.wait_for_timeout(3000)
    raise TimeoutError(
        f"Không tìm thấy mã xác minh mới trong mail.tm sau "
        f"{CODE_TIMEOUT_SECONDS} giây."
    )


def _fill_verification_code(page, code):
    fields = _verification_input(page)
    if not fields:
        raise RuntimeError("Không tìm thấy ô nhập mã xác minh trên Muse.ai.")
    if fields.count() == 1:
        fields.first.fill(code)
        return
    if fields.count() == len(code):
        for index, digit in enumerate(code):
            fields.nth(index).fill(digit)
        return
    raise RuntimeError(
        "Form mã xác minh của Muse.ai không khớp với định dạng mã nhận được."
    )


def _show_error(title, detail):
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror(title, detail, parent=root)
    root.destroy()


if __name__ == "__main__":
    main()
