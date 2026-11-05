import html
import json
import re
import secrets
import sys
import time
import tkinter as tk
import urllib.error
import urllib.parse
import urllib.request
from datetime import date
from pathlib import Path
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
CARD_SERVICE = "MuseAIHelper:card"
CODE_TIMEOUT_SECONDS = 180
PAGE_TIMEOUT_MS = 45000
BUTTON_TIMEOUT_MS = 90000
BUTTON_READY_TIMEOUT_MS = 30000
UI_SETTLE_MS = 700


def _status(message):
    print(message, flush=True)


def _load_proxy():
    """Đọc proxy tuỳ chọn từ file proxy.txt nằm cạnh script.

    Một dòng duy nhất, một trong hai dạng:
    host:port:username:password  hoặc  http://username:password@host:port
    Không có file (hoặc file trống) → None, và main() sẽ huỷ lượt chạy
    vì đăng ký bắt buộc phải qua proxy. Chỉ lưu lượng Chromium
    (muse.ai) đi qua proxy; mail.tm vẫn đi thẳng.
    """
    try:
        raw = Path(__file__).with_name("proxy.txt").read_text(
            encoding="utf-8"
        ).strip()
    except OSError:
        return None
    if not raw:
        return None
    line = raw.splitlines()[0].strip()
    if line.startswith(("http://", "https://")):
        parsed = urllib.parse.urlparse(line)
        if not parsed.hostname or not parsed.port:
            return None
        proxy = {
            "server": f"{parsed.scheme}://{parsed.hostname}:{parsed.port}"
        }
        if parsed.username:
            proxy["username"] = urllib.parse.unquote(parsed.username)
        if parsed.password:
            proxy["password"] = urllib.parse.unquote(parsed.password)
        return proxy
    parts = line.split(":")
    if len(parts) < 2:
        return None
    proxy = {"server": f"http://{parts[0]}:{parts[1]}"}
    if len(parts) >= 4:
        proxy["username"] = parts[2]
        proxy["password"] = ":".join(parts[3:])
    return proxy


def _short_error(error):
    """Rút gọn lỗi Playwright thành một dòng dễ đọc trong log."""
    text = str(error).split("Call log:")[0].strip().splitlines()
    first = text[0] if text else error.__class__.__name__
    return first[:200]


def _classify_after_profile(page, timeout_seconds=45):
    """Phân loại màn hình xuất hiện sau khi gửi form hồ sơ.

    Trả về:
    - "waitlist": màn hình "Join the waitlist" — Muse chưa khả dụng ở
      khu vực của IP hiện tại, đăng ký coi như THẤT BẠI.
    - "card": màn hình nhập thẻ — đăng ký đi đúng hướng; phải hoàn
      tất bước nhập thẻ nữa mới tính là THÀNH CÔNG.
    - "error": hộp thoại "Something went wrong — You can't create a
      new account right now": IP hiện tại đang bị Meta chặn tạo tài
      khoản. THẤT BẠI ngay, và proxy đó bị đánh dấu bỏ.
    - "unknown": hết thời gian chờ mà không thấy các màn hình trên.
    """
    waitlist_markers = ("join the waitlist", "available in your country")
    card_markers = (
        "card number",
        "card details",
        "enter your card",
        "add a card",
        "payment method",
    )
    error_markers = (
        "something went wrong",
        "create a new account right now",
    )
    deadline = time.monotonic() + timeout_seconds
    while time.monotonic() < deadline:
        if page.is_closed():
            return "unknown"
        try:
            text = page.locator("body").inner_text(timeout=3000).lower()
        except Exception:
            text = ""
        if any(marker in text for marker in error_markers):
            return "error"
        if any(marker in text for marker in waitlist_markers):
            return "waitlist"
        if any(marker in text for marker in card_markers):
            return "card"
        page.wait_for_timeout(1200)
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
        page.wait_for_timeout(1500)
    return "timeout"


def _profile_form_present(page):
    """Nhận diện form hồ sơ bằng NHIỀU dấu hiệu độc lập.

    Chỉ dựa vào role "heading" là không đủ tin cậy: trang có thể
    render tiêu đề bằng thẻ thường không có role heading, hoặc dấu
    nháy trong "What's" là nháy cong (’) khác với nháy thẳng trong
    code — cả hai đều khiến tool không thấy form dù nó đang hiển
    thị ngay trên màn hình. Vì vậy kiểm tra cả nội dung text của
    trang (không lệ thuộc dấu nháy hay loại thẻ) lẫn sự hiện diện
    của các ô nhập đặc trưng của form hồ sơ.
    """
    try:
        text = page.locator("body").inner_text(timeout=2000).lower()
    except Exception:
        text = ""
    if "finish creating your account" in text or "your birthday" in text:
        return True
    probes = (
        lambda: page.get_by_placeholder("First name", exact=True).count(),
        lambda: page.locator('select[autocomplete^="bday"]').count(),
        lambda: page.get_by_role(
            "combobox", name=re.compile(r"^(Month|Day|Year)$", re.I)
        ).count(),
    )
    for probe in probes:
        try:
            if probe():
                return True
        except Exception:
            pass
    return False


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
        if _profile_form_present(page):
            return True
        page.wait_for_timeout(500)
    return False


def _code_form_present(page):
    """Form nhập mã còn trên màn hình không (còn ô nhập mã là còn)."""
    try:
        fields = _verification_input(page)
        return bool(fields) and fields.count() > 0
    except Exception:
        return False


def _ensure_profile_form(page):
    """Đưa trang tới form hồ sơ sau khi điền mã — KHÔNG bấm mò.

    Form mã tự gửi khi nhập đủ số nên thường trang tự sang form hồ
    sơ. Chỉ bấm Confirm khi form mã VẪN còn trên màn hình: nếu form
    mã đã biến mất mà form hồ sơ chưa nhận diện được, nút Confirm
    duy nhất còn lại trên màn hình là của form hồ sơ — bấm vào là
    gửi form hồ sơ còn trống (ngày sinh mặc định = ngày hôm nay,
    tuổi 0) và bị chặn dưới 18 tuổi. Trường hợp đó phải dừng và
    báo rõ, tuyệt đối không bấm.
    """
    if _wait_for_profile_form(page, 15000):
        return
    if not _code_form_present(page):
        raise RuntimeError(
            "Form nhập mã đã biến mất nhưng chưa nhận diện được "
            "form hồ sơ — dừng lại để tránh bấm nhầm Confirm của "
            "form hồ sơ khi chưa điền. Hãy chụp màn hình hiện tại "
            "gửi lại để bổ sung cách nhận diện."
        )
    confirm_button = page.get_by_role(
        "button", name="Confirm", exact=True
    )
    _click_when_ready(page, confirm_button, "Confirm")
    if not _wait_for_profile_form(page, 15000):
        raise RuntimeError(
            "Không thấy form hồ sơ sau khi xác nhận mã xác minh."
        )


def _saved_card():
    """Thẻ người dùng đã lưu trong ứng dụng (Windows Credential Manager)."""
    card = {}
    for field in ("number", "expiry", "cvv", "zip"):
        try:
            value = keyring.get_password(CARD_SERVICE, field)
        except Exception:
            value = None
        if value:
            card[field] = value
    return card if card.get("number") else {}


def _fill_card_form(page, card):
    """Điền thẻ đã lưu vào form "Confirm your age" rồi bấm Confirm.

    Không chữ số nào của thẻ được in ra log. Điền xong đọc lại số
    thẻ trong form để đối chiếu (form tự chèn khoảng trắng khi
    format); lệch là dừng, không bấm Confirm.
    """
    def digits(text):
        return re.sub(r"\D", "", str(text or ""))

    number = digits(card.get("number"))
    expiry_digits = digits(card.get("expiry"))
    expiry = (
        f"{expiry_digits[:2]}/{expiry_digits[2:]}"
        if len(expiry_digits) == 4
        else str(card.get("expiry", "")).strip()
    )
    cvv = digits(card.get("cvv"))
    zip_code = digits(card.get("zip")) or "10001"
    if not number or not expiry or not cvv:
        raise RuntimeError("Thẻ đã lưu thiếu thông tin.")
    _status("Đã tới bước thẻ — đang điền thẻ đã lưu...")
    for placeholder, value in (
        ("Card number", number),
        ("MM/YY", expiry),
        ("CVV", cvv),
        ("ZIP code", zip_code),
    ):
        field = page.get_by_placeholder(placeholder).first
        field.wait_for(state="visible", timeout=15000)
        field.fill(value)
        page.wait_for_timeout(250)
    entered = digits(
        page.get_by_placeholder("Card number").first.input_value()
    )
    if entered != number:
        raise RuntimeError("Số thẻ trong form không khớp với thẻ đã lưu.")
    _status("Đã điền thẻ — bấm Confirm.")
    confirm = page.get_by_role("button", name="Confirm", exact=True)
    _click_when_ready(page, confirm, "Confirm")


def _in_muse_interface(page):
    """Đã vào giao diện Muse chưa (URL rời khỏi luồng đăng ký/xác minh)."""
    try:
        url = page.url.lower()
    except Exception:
        return False
    if "muse.ai" not in url:
        return False
    return not any(
        part in url
        for part in (
            "verification",
            "signup",
            "sign-up",
            "checkpoint",
            "age-verification",
            "onboarding",
            "welcome",
        )
    )


def _confirm_final_step(page, timeout_seconds=30):
    """Bấm nút xác nhận cuối sau bước thẻ để vào giao diện Muse.

    Sau khi thẻ được chấp nhận, Muse.ai còn một màn hình nữa với
    nút "Get started" (đôi khi là Confirm/Continue) — bấm nút đó
    mới vào được giao diện. Nếu có nhiều màn hình nối nhau thì
    bấm tiếp (tối đa 3 lần). Đã ở trong giao diện rồi — hoặc hết
    giờ mà không thấy nút — thì thôi, bước lấy cookie vẫn chạy.
    """
    deadline = time.monotonic() + timeout_seconds
    clicks = 0
    last_click = None
    interface_since = None
    while time.monotonic() < deadline:
        if page.is_closed():
            return
        if clicks >= 1 and _in_muse_interface(page):
            # Đã bấm qua cổng và đã ở trong giao diện: phần còn
            # lại là giao diện thật, không bấm thêm gì nữa.
            return
        clicked = False
        for name in ("Get started", "Confirm", "Continue", "Bắt đầu"):
            button = page.get_by_role("button", name=name, exact=True)
            try:
                visible = (
                    button.count()
                    and button.first.is_visible()
                    and button.first.is_enabled()
                )
            except Exception:
                visible = False
            if not visible:
                # Phòng khi nút không khai role=button: dò theo chữ.
                try:
                    text_hit = page.get_by_text(name, exact=True)
                    if text_hit.count() and text_hit.first.is_visible():
                        button = text_hit
                        visible = True
                except Exception:
                    pass
            if not visible:
                continue
            url_before = page.url
            if last_click == (name, url_before):
                # Đã bấm đúng nút này ở đúng URL này mà nó vẫn còn
                # nguyên: đó là nút cố định của giao diện (checklist,
                # banner...), không phải cổng vào — dừng lại.
                return
            _click_when_ready(page, button, name)
            last_click = (name, url_before)
            clicks += 1
            clicked = True
            _status(
                f"Đã bấm nút '{name}' sau bước thẻ ({clicks}/3)."
            )
            break
        if clicks >= 3:
            return
        if not clicked:
            if _in_muse_interface(page):
                # URL giống giao diện chưa đủ tin — lúc chuyển
                # cảnh nó đổi trước khi màn hình cổng kịp hiện
                # (đã trượt thật ở tài khoản 2, 16:36). Chỉ dừng
                # khi thấy khung giao diện thật, hoặc URL đã yên
                # ở dạng giao diện suốt 12 giây mà cổng không hiện.
                if _app_shell_visible(page):
                    return
                if interface_since is None:
                    interface_since = time.monotonic()
                elif time.monotonic() - interface_since >= 12:
                    return
            else:
                interface_since = None
        page.wait_for_timeout(1200 if clicked else 1000)


def _wait_for_muse_interface(page, timeout_seconds=30):
    """Chờ URL rời khỏi luồng xác minh rồi thêm một nhịp ổn định."""
    deadline = time.monotonic() + timeout_seconds
    while time.monotonic() < deadline:
        if page.is_closed():
            return
        if _in_muse_interface(page):
            break
        page.wait_for_timeout(1000)
    page.wait_for_timeout(1500)


def _app_shell_visible(page):
    """Khung giao diện đã dựng chưa: nút Mời hiện hoặc thanh trái có nút.

    Chỉ nhìn URL thì không đủ tin: lúc chuyển cảnh, URL đã là
    trang chủ trong khi màn hình cổng Get started còn chưa kịp
    hiện — phải thấy bằng chứng của giao diện thật mới tin.
    """
    # Giao diện trong trình duyệt của tool là tiếng Anh (nút
    # "Invite"), trên máy người dùng là tiếng Việt (nút "Mời"):
    # dò cả hai.
    for name in ("Mời", "Invite"):
        invite = page.get_by_role("button", name=name, exact=False)
        try:
            if invite.count() and invite.first.is_visible():
                return True
        except Exception:
            pass
    try:
        return bool(
            page.evaluate(
                """() => [...document.querySelectorAll('button')]
                    .some((el) => {
                        const r = el.getBoundingClientRect();
                        return (
                            r.width > 0 &&
                            r.width <= 90 &&
                            r.left < 90 &&
                            r.bottom > window.innerHeight - 100
                        );
                    })"""
            )
        )
    except Exception:
        return False


def _wait_for_app_shell(page, timeout_seconds=20):
    """Chờ khung giao diện Muse dựng xong (nút Mời hoặc thanh trái).

    URL đổi sang giao diện trước khi các nút bấm sẵn sàng — qua
    proxy chậm, quy đổi/lấy mã chạy ngay lúc này sẽ trượt hết.
    Thấy nút Mời ở góc phải trên, hoặc thanh icon bên trái đã
    dựng, thì mới coi là sẵn sàng; hết giờ vẫn đi tiếp.
    """
    deadline = time.monotonic() + timeout_seconds
    while time.monotonic() < deadline:
        if page.is_closed():
            return
        if _app_shell_visible(page):
            return
        page.wait_for_timeout(800)


def _session_cookie_header(page):
    """Cookie phiên đăng nhập, dạng chuỗi "name=value; name=value".

    Đúng định dạng Cookie header như trình duyệt hiển thị — dán
    thẳng vào cookie editor hay HTTP client là dùng được. Chỉ
    lấy cookie thuộc muse.ai / meta.com / facebook.com (phiên
    đăng nhập Meta rải trên cả ba). Dòng SESSION_COOKIE được
    ứng dụng chính chặn lại để ghi vào sheet — không hiển thị
    trong log.
    """
    try:
        cookies = page.context.cookies()
    except Exception:
        return ""
    parts = []
    for cookie in cookies:
        domain = cookie.get("domain") or ""
        if not any(
            host in domain
            for host in ("muse.ai", "meta.com", "facebook.com")
        ):
            continue
        name = cookie.get("name") or ""
        if name:
            parts.append(f"{name}={cookie.get('value') or ''}")
    return "; ".join(parts)


def _load_redeem_code():
    """Mã mời cần quy đổi cho tài khoản này (app chính ghi redeem.txt).

    Đó là mã mời của tài khoản trước trong chuỗi. File trống hoặc
    không có = không cần quy đổi (tài khoản đầu chuỗi).
    """
    try:
        text = (
            Path(__file__)
            .with_name("redeem.txt")
            .read_text(encoding="utf-8")
        )
    except OSError:
        return ""
    code = text.strip().upper()
    return code if re.fullmatch(r"[A-Z0-9]{6}", code) else ""


REDEEM_ROW_TEXTS = (
    "Quy đổi mã lời mời",
    "Redeem invite code",
    "Redeem an invite code",
    "Redeem invitation code",
    "Redeem a code",
)


def _settings_open(page):
    """Cài đặt đã mở chưa — mục quy đổi nằm ngay tab Chung/General."""
    for text in REDEEM_ROW_TEXTS:
        probe = page.get_by_text(text, exact=False)
        try:
            if probe.count():
                return True
        except Exception:
            pass
    return False


def _open_settings(page):
    """Mở Cài đặt: phím tắt Ctrl+, trước, menu góc trái dưới sau.

    Menu của Muse ghi rõ Cài đặt = Ctrl+, — và theo người
    dùng, bấm chuột mở Cài đặt không ăn, phím tắt mới là đường
    chính. Hai việc trước khi bấm phím: giành focus cho trang
    (focus đang nằm ở ô nhập thì phím tắt rơi vào hư không),
    và thử cả hai dạng tên phím. Sau mỗi lần bấm đều chờ xác
    nhận Cài đặt đã mở thật mới đi tiếp.
    """
    try:
        page.mouse.click(200, 200)
        page.wait_for_timeout(300)
    except Exception:
        pass
    for keys in ("Control+,", "Control+Comma"):
        try:
            page.keyboard.press(keys)
        except Exception:
            pass
        deadline = time.monotonic() + 3
        while time.monotonic() < deadline:
            if page.is_closed():
                return False
            if _settings_open(page):
                return True
            page.wait_for_timeout(500)
    for name in ("Cài đặt", "Settings"):
        item = page.get_by_text(name, exact=True)
        try:
            if item.count() and item.first.is_visible():
                item.first.click(timeout=5000)
                return True
        except Exception:
            pass
    try:
        opened = page.evaluate(
            """() => {
                const els = [
                    ...document.querySelectorAll(
                        'button, [role="button"], a'
                    ),
                ];
                const cands = els.filter((el) => {
                    const r = el.getBoundingClientRect();
                    return (
                        r.width > 0 &&
                        r.width <= 90 &&
                        r.left < 90 &&
                        r.bottom > window.innerHeight - 100
                    );
                });
                if (!cands.length) return false;
                cands.sort(
                    (a, b) =>
                        b.getBoundingClientRect().bottom -
                        a.getBoundingClientRect().bottom
                );
                cands[0].click();
                return true;
            }"""
        )
    except Exception:
        opened = False
    if not opened:
        return False
    # Menu trượt lên cũng có thể chậm: chờ mục Cài đặt xuất hiện.
    deadline = time.monotonic() + 8
    while time.monotonic() < deadline:
        if page.is_closed():
            return False
        for name in ("Cài đặt", "Settings"):
            item = page.get_by_text(name, exact=True)
            try:
                if item.count() and item.first.is_visible():
                    item.first.click(timeout=5000)
                    return True
            except Exception:
                pass
        page.wait_for_timeout(600)
    return False


def _redeem_dialog_present(page):
    """Hộp thoại nhập mã đã mở chưa.

    Tiêu đề khớp một phần (tiếng Anh có thể là "Redeem invite
    code", tiếng Việt "Quy đổi mã...") NHƯNG phải kèm ít nhất
    một ô nhập đang hiện — mục quy đổi trong Cài đặt cũng chứa
    các chữ đó, chỉ khác là nó không có ô nhập nào.
    """
    for title in (
        "Quy đổi mã",
        "Redeem invite code",
        "Redeem a code",
        "Redeem code",
    ):
        heading = page.get_by_text(title, exact=False)
        try:
            if heading.count() and heading.first.is_visible():
                if page.locator("input:visible").count() >= 1:
                    return True
        except Exception:
            pass
    return False


def _fill_redeem_code(page, code):
    """Điền mã 6 ký tự vào hộp thoại quy đổi, có đọc lại kiểm chứng.

    Lần lượt thử: điền từng ô (6 ô riêng) → gõ phím như người
    dùng (bấm ô đầu rồi gõ, các ô tự nhảy) → điền một ô duy nhất
    chứa cả mã. Cách nào cho tổng số ký tự đọc lại đúng bằng độ
    dài mã thì coi như xong.
    """
    boxes = page.locator("input:visible")
    try:
        count = boxes.count()
    except Exception:
        return False
    if count == 0:
        return False

    def total_filled():
        try:
            return sum(
                len(boxes.nth(index).input_value() or "")
                for index in range(boxes.count())
            )
        except Exception:
            return -1

    if count == len(code) and count > 1:
        try:
            for index, char in enumerate(code):
                boxes.nth(index).fill(char)
        except Exception:
            pass
        if total_filled() == len(code):
            return True
    try:
        boxes.first.click(timeout=5000)
        page.keyboard.type(code, delay=80)
    except Exception:
        pass
    if total_filled() == len(code):
        return True
    if count == 1:
        try:
            boxes.first.fill(code)
        except Exception:
            pass
        return total_filled() == len(code)
    return False


def _close_settings(page):
    try:
        page.keyboard.press("Escape")
        page.wait_for_timeout(400)
        page.keyboard.press("Escape")
        page.wait_for_timeout(400)
    except Exception:
        pass


def _redeem_invite_code(page, code):
    """Quy đổi mã mời của tài khoản trước trong Cài đặt.

    Luồng đúng theo giao diện: menu góc trái dưới → Cài đặt →
    "Quy đổi mã lời mời" → điền 6 ô mã → Xác nhận. Trả True khi
    hộp thoại quy đổi đóng lại. Lỗi ở bất kỳ bước nào thì False —
    tài khoản vẫn tính là tạo thành công.
    """
    try:
        if not _open_settings(page):
            _status("Quy đổi mã: không mở được Cài đặt.")
            return False
        # Hàng quy đổi nằm khá sâu trong Cài đặt và hiện hơi
        # muộn (trang nạp chậm): dò theo sự tồn tại tới 20 giây
        # (không đòi hỏi đã nhìn thấy), bấm thì Playwright tự
        # cuộn tới.
        row = None
        deadline = time.monotonic() + 20
        while row is None and time.monotonic() < deadline:
            if page.is_closed():
                return False
            for text in REDEEM_ROW_TEXTS:
                candidate = page.get_by_text(text, exact=False)
                try:
                    if candidate.count():
                        row = candidate.first
                        break
                except Exception:
                    pass
            if row is None:
                page.wait_for_timeout(600)
        if row is None:
            _status(
                "Quy đổi mã: không thấy mục quy đổi mã trong Cài đặt."
            )
            return False
        row.scroll_into_view_if_needed(timeout=5000)
        row.click(timeout=5000)
        deadline = time.monotonic() + 15
        while not _redeem_dialog_present(page):
            if page.is_closed() or time.monotonic() >= deadline:
                _status(
                    "Quy đổi mã: bấm mục quy đổi nhưng hộp thoại "
                    "không mở."
                )
                return False
            page.wait_for_timeout(500)
        if not _fill_redeem_code(page, code):
            _status("Quy đổi mã: không điền được mã vào hộp thoại.")
            return False
        confirm = page.get_by_role(
            "button", name="Xác nhận", exact=True
        )
        if not (confirm.count() and confirm.first.is_visible()):
            confirm = page.get_by_role(
                "button", name="Confirm", exact=True
            )
        if not (confirm.count() and confirm.first.is_visible()):
            _status("Quy đổi mã: không thấy nút Xác nhận.")
            return False
        button = confirm.first
        deadline = time.monotonic() + 10
        while not button.is_enabled():
            if time.monotonic() >= deadline:
                _status(
                    "Quy đổi mã: nút Xác nhận không sáng lên "
                    "(mã chưa đủ 6 ký tự?)."
                )
                return False
            page.wait_for_timeout(300)
        _click_when_ready(page, button, "Xác nhận")
        deadline = time.monotonic() + 15
        while _redeem_dialog_present(page):
            if page.is_closed():
                return True
            if time.monotonic() >= deadline:
                _status(
                    "Quy đổi mã: đã bấm Xác nhận nhưng hộp thoại "
                    "không đóng — mã có thể sai hoặc mạng chậm."
                )
                return False
            page.wait_for_timeout(500)
        return True
    except Exception:
        return False
    finally:
        _close_settings(page)


_HARVEST_CODE_JS = """() => {
    const panels = [
        ...document.querySelectorAll(
            'div, section, aside, [role="dialog"]'
        ),
    ].filter((el) => {
        const r = el.getBoundingClientRect();
        return (
            r.width > 0 &&
            /Mời bạn bè|Invite friends/.test(el.textContent || '')
        );
    });
    if (!panels.length) return '';
    panels.sort(
        (a, b) =>
            (a.textContent || '').length - (b.textContent || '').length
    );
    const panel = panels[0];
    for (const el of panel.querySelectorAll('*')) {
        if (
            el.tagName === 'INPUT' &&
            /^[A-Z0-9]{6}$/.test((el.value || '').trim())
        ) {
            return el.value.trim();
        }
        if (el.childElementCount === 0) {
            const t = (el.textContent || '').trim();
            if (/^[A-Z0-9]{6}$/.test(t)) return t;
        }
    }
    const stripped = (s) => (s || '').replace(/\\s+/g, '');
    for (const el of panel.querySelectorAll('*')) {
        if (
            el.childElementCount <= 8 &&
            /^[A-Z0-9]{6}$/.test(stripped(el.textContent))
        ) {
            return stripped(el.textContent);
        }
    }
    const m = (panel.textContent || '').match(/[A-Z0-9]{6}/);
    return m ? m[0] : '';
}"""


def _read_clipboard(page):
    try:
        page.context.grant_permissions(
            ["clipboard-read", "clipboard-write"],
            origin="https://muse.ai",
        )
    except Exception:
        pass
    try:
        return (
            page.evaluate(
                "() => navigator.clipboard.readText().catch(() => '')"
            )
            or ""
        )
    except Exception:
        return ""


def _invite_code_from_text(text):
    """Bóc mã mời 6 ký tự (chữ hoa/số) khỏi một đoạn text hoặc URL."""
    if not text:
        return ""
    text = text.strip()
    if re.fullmatch(r"[A-Z0-9]{6}", text):
        return text
    # Mẫu tin chia sẻ của Muse có sẵn nhãn: "Mã: ATGXDX
    # https://muse.ai/join" — bám nhãn là chắc chắn nhất.
    anchored = re.search(
        r"(?:Mã|Code)\s*:\s*([A-Z0-9]{6})", text, re.IGNORECASE
    )
    if anchored:
        return anchored.group(1).upper()
    match = re.search(r"(?<![A-Za-z0-9])[A-Z0-9]{6}(?![A-Za-z0-9])", text)
    return match.group(0) if match else ""


def _harvest_via_copy_link(page):
    """Bấm "Copy referral link" rồi bóc mã khỏi nội dung vừa copy.

    Đường lấy mã chính — nhưng chỉ sau khi đã cào lớp phủ
    "Scratch to reveal" (chưa cào thì copy không ra mã). Nội
    dung copy chứa mã trong tin chia sẻ "Mã: XXXXXX" hoặc
    trong link. Panel mở chậm nên chờ nút copy xuất hiện tới
    10 giây thay vì thử đúng một lần.
    """
    deadline = time.monotonic() + 10
    while time.monotonic() < deadline:
        if page.is_closed():
            return ""
        for name in (
            "Sao chép liên kết giới thiệu",
            "Copy referral link",
            "Sao chép",
            "Copy link",
        ):
            button = page.get_by_role("button", name=name, exact=False)
            try:
                if button.count() and button.first.is_visible():
                    button.first.click(timeout=5000)
                    page.wait_for_timeout(700)
                    return _invite_code_from_text(
                        _read_clipboard(page)
                    )
            except Exception:
                pass
        page.wait_for_timeout(700)
    return ""


_SCRATCH_RECT_JS = """() => {
    const hits = [
        ...document.querySelectorAll(
            'div, span, button, canvas, section'
        ),
    ].filter((el) => {
        const r = el.getBoundingClientRect();
        return (
            r.width >= 40 &&
            r.height >= 8 &&
            /scratch to reveal|cào để/i.test(el.textContent || '')
        );
    });
    if (!hits.length) return null;
    hits.sort((a, b) => {
        const ra = a.getBoundingClientRect();
        const rb = b.getBoundingClientRect();
        return ra.width * ra.height - rb.width * rb.height;
    });
    const r = hits[0].getBoundingClientRect();
    return {x: r.x, y: r.y, width: r.width, height: r.height};
}"""


def _scratch_box(page, box):
    """Kéo chuột cào qua ô "Scratch to reveal" (vài đường ngang)."""
    try:
        left = box["x"] + 4
        right = box["x"] + box["width"] - 4
        for fraction in (0.5, 0.3, 0.7):
            y = box["y"] + box["height"] * fraction
            page.mouse.move(left, y)
            page.mouse.down()
            page.mouse.move(right, y, steps=20)
            page.mouse.up()
        page.wait_for_timeout(700)
        return True
    except Exception:
        return False


def _prepare_panel(page):
    """Chờ panel Invite sẵn sàng, cào lớp phủ che mã nếu có.

    Thứ tự đúng theo giao diện thật: mã của tài khoản mới nằm
    sau lớp "Scratch to reveal" — phải cào xong mới bấm Copy
    referral link lấy mã được. Trong lúc chờ lớp cào hiện (tối
    đa 10 giây), nếu quét thấy mã luôn (tài khoản mã đã hiện
    sẵn) thì trả mã ngay.
    """
    deadline = time.monotonic() + 10
    while time.monotonic() < deadline:
        if page.is_closed():
            return ""
        try:
            found = page.evaluate(_HARVEST_CODE_JS) or ""
        except Exception:
            found = ""
        found = str(found).strip().upper()
        if re.fullmatch(r"[A-Z0-9]{6}", found):
            return found
        try:
            box = page.evaluate(_SCRATCH_RECT_JS)
        except Exception:
            box = None
        if box:
            if _scratch_box(page, box):
                _status("Đã cào lớp phủ 'Scratch to reveal' che mã mời.")
            return ""
        page.wait_for_timeout(700)
    return ""


def _scan_panel_code(page, seconds):
    """Quét chữ trong panel Mời tìm mã 6 ký tự, chờ tối đa vài giây."""
    deadline = time.monotonic() + seconds
    while time.monotonic() < deadline:
        if page.is_closed():
            return ""
        try:
            found = page.evaluate(_HARVEST_CODE_JS) or ""
        except Exception:
            found = ""
        found = str(found).strip().upper()
        if re.fullmatch(r"[A-Z0-9]{6}", found):
            return found
        page.wait_for_timeout(700)
    return ""


def _harvest_invite_code(page):
    """Lấy mã mời của chính tài khoản này từ panel "Mời bạn bè".

    Bấm nút Mời/Invite ở góc phải trên → cào lớp "Scratch to
    reveal" che mã (bắt buộc trước, chưa cào thì copy không
    ra mã) → bấm "Copy referral link" và bóc mã khỏi nội dung
    copy. Dự phòng cuối: quét chữ trong panel. Không lấy được
    thì trả "" — tài khoản vẫn được lưu bình thường.
    """
    code = ""
    try:
        # Nút Mời có thể hiện muộn vài giây sau khi vào giao
        # diện (nhất là qua proxy chậm): chờ nó xuất hiện thay
        # vì thử đúng một lần rồi bỏ cuộc.
        button = None
        deadline = time.monotonic() + 20
        while button is None and time.monotonic() < deadline:
            if page.is_closed():
                return ""
            for name in ("Mời", "Invite"):
                candidate = page.get_by_role(
                    "button", name=name, exact=False
                )
                try:
                    if (
                        candidate.count()
                        and candidate.first.is_visible()
                    ):
                        button = candidate.first
                        break
                except Exception:
                    pass
            if button is None:
                page.wait_for_timeout(800)
        if button is None:
            return ""
        button.click(timeout=5000)
        # Cào lớp phủ trước (hoặc nhận mã ngay nếu đã hiện sẵn),
        # rồi mới copy — chưa cào thì copy không ra mã.
        code = _prepare_panel(page)
        if not code:
            code = _harvest_via_copy_link(page)
        if not code:
            code = _scan_panel_code(page, 6)
    except Exception:
        return ""
    finally:
        try:
            page.keyboard.press("Escape")
        except Exception:
            pass
    return code


def _locator_visible(locator):
    try:
        return locator.count() > 0 and locator.first.is_visible()
    except Exception:
        return False


def _locator_enabled(locator):
    try:
        return locator.count() > 0 and locator.first.is_enabled()
    except Exception:
        return False


def _open_signup_and_submit_email(page, email):
    """Mở trang đăng ký, chờ form HIỆN HẲN rồi mới điền email.

    Không điền ngay khi DOM vừa có ô nhập: trang qua proxy nạp
    chậm, form có thể hiện từng phần hoặc bị render lại ngay sau
    đó — điền sớm là email bị nuốt hoặc form bị thay giữa chừng.
    Điều kiện để điền: ô email đã hiển thị và nút Continue đã
    hiện VÀ đã bật (đổi sang màu sẵn sàng — nút còn xám là trang
    chưa nạp xong), cộng thêm một nhịp chờ cho form ổn định.
    Điền xong còn kiểm tra lại giá trị; bị mất thì điền lại một
    lần, vẫn mất mới báo lỗi. Trả về nút Continue để main() bấm
    sau khi chụp mốc thư cũ.
    """
    page.goto(
        SIGNUP_URL,
        wait_until="domcontentloaded",
        timeout=PAGE_TIMEOUT_MS,
    )
    _status("Đã mở trang Muse.ai — chờ form hiển thị đầy đủ...")
    email_field = page.get_by_role(
        "textbox", name="Mobile number or email"
    )
    login_button = page.get_by_role("button", name="Log in")
    continue_button = page.get_by_role(
        "button", name="Continue", exact=True
    )
    deadline = time.monotonic() + PAGE_TIMEOUT_MS / 1000
    while not _locator_visible(email_field):
        if page.is_closed():
            raise RuntimeError(
                "Cửa sổ Chromium đã đóng trước khi form hiện."
            )
        if _locator_visible(login_button):
            _click_when_ready(page, login_button, "Log in")
            break
        if time.monotonic() >= deadline:
            raise RuntimeError(
                "Chờ mãi không thấy form nhập email của Muse.ai."
            )
        page.wait_for_timeout(500)
    email_field.wait_for(state="visible", timeout=PAGE_TIMEOUT_MS)
    # Nút Continue phải HIỆN và BẬT (đổi sang màu sẵn sàng) rồi mới
    # được điền email — nút còn xám/chưa bật nghĩa là trang chưa
    # nạp xong, kể cả khi các ô đã có mặt trong DOM.
    continue_button.wait_for(state="visible", timeout=BUTTON_TIMEOUT_MS)
    ready_deadline = time.monotonic() + BUTTON_READY_TIMEOUT_MS / 1000
    button_ready = False
    while time.monotonic() < ready_deadline:
        if page.is_closed():
            raise RuntimeError(
                "Cửa sổ Chromium đã đóng trước khi form sẵn sàng."
            )
        if _locator_enabled(continue_button):
            button_ready = True
            break
        page.wait_for_timeout(500)
    if button_ready:
        _status("Nút Continue đã bật — form đã sẵn sàng.")
    else:
        _status(
            "Nút Continue vẫn chưa bật sau "
            f"{BUTTON_READY_TIMEOUT_MS // 1000} giây — điền email "
            "trước để kích hoạt nút."
        )
    page.wait_for_timeout(UI_SETTLE_MS)
    email_field.wait_for(state="visible", timeout=PAGE_TIMEOUT_MS)
    _status(f"Form đã hiển thị đầy đủ — đang điền email {email}...")
    for _ in range(2):
        email_field.fill(email)
        page.wait_for_timeout(400)
        if email_field.input_value() == email:
            return continue_button
    raise RuntimeError(
        "Email chưa được giữ trong form Muse.ai; chưa bấm Continue."
    )


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

    proxy = _load_proxy()
    if not proxy:
        detail = (
            "Chưa cấu hình proxy — tool chỉ đăng ký qua proxy. "
            "Hãy nhập proxy trong ứng dụng rồi chạy lại."
        )
        if auto:
            _status(f"RESULT: FAIL {detail}")
            sys.exit(1)
        _show_error("Chưa cấu hình proxy", detail)
        return

    playwright = sync_playwright().start()
    browser = None
    page = None
    try:
        token = _get_mail_token(email, mailbox_password)
        browser = playwright.chromium.launch(
            headless=False, proxy=proxy
        )
        page = browser.new_page()
        page.set_default_timeout(PAGE_TIMEOUT_MS)
        # Không kiểm tra IP thoát nữa: kho proxy do người dùng đảm
        # bảo toàn bộ là proxy US. Nếu một proxy nào đó thoát sai
        # khu vực, màn hình waitlist sau bước hồ sơ vẫn chặn lại.
        continue_button = _open_signup_and_submit_email(page, email)
        try:
            existing_ids = _get_message_ids(token)
        except Exception:
            # Hộp thư mới luôn trống; không lấy được danh sách thư cũ
            # thì dùng tập rỗng làm mốc, đừng chặn cả lượt đăng ký.
            existing_ids = set()
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
        _ensure_profile_form(page)
        _status("Mã đã xác nhận, đang điền hồ sơ...")
        _fill_profile(page)
        finish_button = page.get_by_role(
            "button", name="Confirm", exact=True
        )
        _click_when_ready(page, finish_button, "Confirm")
        if auto:
            _status("Đã gửi hồ sơ, đang chờ màn hình tiếp theo...")
            outcome = _classify_after_profile(page)
            if outcome == "error":
                _status(
                    "PROXY_BURN: Meta chặn tạo tài khoản trên IP này "
                    "('Something went wrong')."
                )
                raise RuntimeError(
                    "Meta báo không thể tạo tài khoản lúc này — IP "
                    "của proxy đang dùng đã bị chặn; đổi proxy khác."
                )
            if outcome == "waitlist":
                raise RuntimeError(
                    "Ra màn hình 'Join the waitlist' — Muse chưa khả "
                    "dụng ở IP thoát hiện tại. Hãy đổi sang proxy US "
                    "khác rồi chạy lại."
                )
            if outcome != "card":
                raise RuntimeError(
                    "Sau hồ sơ không ra bước nhập thẻ, cũng không ra "
                    "waitlist — bước tiếp theo chưa rõ, hãy chụp màn "
                    "hình gửi lại để bổ sung vào tool."
                )
            card = _saved_card()
            auto_filled = False
            if card:
                try:
                    _fill_card_form(page, card)
                    auto_filled = True
                except Exception as error:
                    _status(
                        "Tự điền thẻ không xong ("
                        + _short_error(error)
                        + ") — chuyển sang chờ bạn nhập tay."
                    )
            if auto_filled:
                card_outcome = _wait_for_card_step_done(
                    page, timeout_seconds=120
                )
            else:
                _status(
                    "Đã tới bước nhập thẻ — hãy nhập thẻ và ZIP US "
                    "trong cửa sổ Chromium. Tool chờ tối đa 4 phút; "
                    "thẻ được chấp nhận thì mới lưu tài khoản."
                )
                card_outcome = _wait_for_card_step_done(page)
            if card_outcome != "done":
                raise RuntimeError(
                    "Chưa hoàn tất bước nhập thẻ "
                    f"(kết quả chờ: {card_outcome}) — tài khoản "
                    "không được lưu."
                )
            _status(
                "Đã qua bước nhập thẻ — xử lý bước xác nhận cuối "
                "để vào giao diện Muse..."
            )
            _confirm_final_step(page)
            _wait_for_muse_interface(page)
            _wait_for_app_shell(page)
            redeem_code = _load_redeem_code()
            if redeem_code:
                if _redeem_invite_code(page, redeem_code):
                    _status(
                        f"Đã quy đổi mã mời {redeem_code} của tài "
                        "khoản trước."
                    )
                else:
                    _status(
                        f"Không quy đổi được mã mời {redeem_code} — "
                        "tài khoản vẫn được lưu; bạn có thể tự quy "
                        "đổi trong Cài đặt trong vòng 48 giờ."
                    )
            own_code = _harvest_invite_code(page)
            if own_code:
                _status(f"INVITE_CODE: {own_code}")
                _status(
                    f"Đã lấy mã mời của tài khoản này: {own_code} — "
                    "tài khoản sau sẽ quy đổi mã này."
                )
            else:
                _status(
                    "Không lấy được mã mời của tài khoản này — "
                    "tài khoản sau sẽ dùng tiếp mã đang chờ hiện tại."
                )
            cookie_header = _session_cookie_header(page)
            if cookie_header:
                _status(f"SESSION_COOKIE: {cookie_header}")
                _status(
                    "Đã vào giao diện Muse, đã lấy cookie phiên "
                    "đăng nhập — tài khoản được lưu."
                )
            else:
                _status(
                    "Đã qua bước thẻ nhưng không lấy được cookie "
                    "phiên đăng nhập nào — tài khoản vẫn được lưu, "
                    "cột cookie sẽ trống."
                )
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
    last_error = None
    for _ in range(3):
        try:
            response = _mail_request(
                "POST", "/token", {"address": email, "password": password}
            )
        except Exception as error:
            last_error = error
            time.sleep(2)
            continue
        token = response.get("token")
        if not token:
            raise RuntimeError("mail.tm không trả về token truy cập.")
        return token
    raise last_error


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
        try:
            detail = _mail_request(
                "GET", f"/messages/{message_id}", token=token
            )
        except Exception:
            continue
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
    warned = False
    while time.monotonic() < deadline:
        if page.is_closed():
            raise RuntimeError("Cửa sổ Chromium đã đóng trước khi xác minh.")
        if code is None:
            try:
                code = _verification_code_from_new_messages(
                    token, existing_ids
                )
            except Exception:
                # Mạng qua VPN hay chập chờn: một lần rớt kết nối tới
                # mail.tm không được giết cả tài khoản — bỏ qua lượt
                # kiểm tra này và thử lại ở vòng chờ tiếp theo.
                if not warned:
                    _status(
                        "Kiểm tra hộp thư bị rớt kết nối — đang thử lại..."
                    )
                    warned = True
        fields = _verification_input(page)
        if code and fields:
            return code
        page.wait_for_timeout(2000)
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
