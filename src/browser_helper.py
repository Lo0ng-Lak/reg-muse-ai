import json
import re
import secrets
import string
import sys
import time
import tkinter as tk
import urllib.error
import urllib.parse
import urllib.request
import uuid
from datetime import date
from pathlib import Path
from tkinter import messagebox

import keyring

import capsolver
import mail_providers
import paths


# Ép stdout/stderr sang UTF-8. Trên Windows, print() tiếng Việt vào pipe
# mặc định dùng bảng mã console (charmap) và sập ngay ở dòng trạng thái
# đầu tiên — trước cả khi kịp điền email vào form.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


SIGNUP_URL = "https://muse.ai/"
# Đăng ký tài khoản Meta (chế độ --meta): vào meta.ai rồi đăng
# ký qua auth.meta.com, KHÔNG qua Muse.ai. Luồng này chạy hoàn
# toàn tự động (CAPTCHA giải qua CapSolver, selfie up từ file)
# theo thứ tự sư huynh đi tay và ghi lại 2026-11-13.
META_URL = "https://www.meta.ai/"
META_AUTH_HOST = "auth.meta.com"
# CapSolver API key lưu trong Windows Credential Manager (app
# chính lưu qua ô CapSolver, helper chỉ đọc — không in ra log).
CAPSOLVER_SERVICE = "MuseAIHelper:capsolver"
CAPSOLVER_KEY_FIELD = "api_key"
# Thời gian chờ tối đa cho chặng sau OTP trong luồng Meta
# (CAPTCHA + selfie + vào tài khoản): 7 phút.
META_POST_OTP_TIMEOUT_SECONDS = 420
# API native của muse.ai (phần auth mà trang đăng ký tự gọi).
# Chỉ gọi được TỪ TRONG trình duyệt (fetch cùng origin) — gọi
# từ ngoài bằng curl bị 403 Forbidden ngay từ send-otp (đã thử
# 2026-11-12).
NATIVE_AUTH_BASE = "https://muse.ai/api/auth/native"
# Hậu tố file riêng của luồng song song: "" = luồng duy
# nhất/luồng 1 (proxy.txt, redeem.txt như cũ); "-2".."-5"
# khi app chính chạy nhiều luồng (proxy-2.txt, ...).
# main() gán từ tham số dòng lệnh thứ 4.
LANE_SUFFIX = ""
# Chế độ IP máy (VPN): proxy.txt chứa câu "DIRECT" thì helper
# chạy Chromium thẳng, không qua proxy. main() gán từ _load_proxy().
DIRECT_MODE = False
CARD_SERVICE = "MuseAIHelper:card"
CODE_TIMEOUT_SECONDS = 180
PAGE_TIMEOUT_MS = 45000
BUTTON_TIMEOUT_MS = 90000
BUTTON_READY_TIMEOUT_MS = 30000
UI_SETTLE_MS = 700


def _status(message):
    print(message, flush=True)


def _short_error(error):
    """Dòng lỗi NGẮN một dòng cho log.

    Lỗi Playwright mang theo cả một "Call log" nhiều dòng; in
    nguyên văn là lý do thật bị chôn ở cuối và người dùng copy
    log không thấy được. Chỉ lấy dòng đầu tiên, cắt 200 ký tự.
    """
    text = str(error).strip()
    first = text.splitlines()[0] if text else repr(error)
    return first[:200]


def _load_proxy():
    """Đọc proxy tuỳ chọn từ file proxy.txt nằm cạnh script.

    Một dòng duy nhất, một trong ba dạng:
    host:port:username:password, http://username:password@host:port,
    hoặc câu đánh dấu DIRECT (chế độ IP máy: chạy thẳng qua IP
    của máy, không qua proxy — trả về chuỗi "DIRECT" để main()
    phân biệt với "không có proxy").
    Không có file (hoặc file trống / dòng không đọc được) → None,
    và main() sẽ huỷ lượt chạy vì đăng ký bắt buộc phải qua proxy
    hoặc qua IP máy. Chỉ lưu lượng Chromium (muse.ai) đi qua
    proxy; dịch vụ temp mail vẫn đi thẳng.
    """
    try:
        raw = paths.data_file(f"proxy{LANE_SUFFIX}.txt").read_text(
            encoding="utf-8"
        ).strip()
    except OSError:
        return None
    if not raw:
        return None
    line = raw.splitlines()[0].strip()
    if line.upper() == "DIRECT":
        return "DIRECT"
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
            paths.data_file(f"redeem{LANE_SUFFIX}.txt")
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


def _dialog_scope(page):
    """Phạm vi hộp thoại đang mở, để không đụng trang phía sau.

    Mọi thao tác với ô nhập mã và nút xác nhận phải giới hạn
    trong hộp thoại: trang phía sau vẫn còn khung chat và các
    nút khác — điền/tìm ở phạm vi toàn trang là mã có thể bị
    gõ nhầm vào khung chat. Không có phần tử dialog (hoặc đối
    tượng trang không hỗ trợ truy vấn lồng nhau) thì dùng cả
    trang như cũ.
    """
    try:
        dialogs = page.locator('[role="dialog"]')
        if dialogs.count() >= 1:
            scope = dialogs.first
            scope.locator("input")
            return scope
    except Exception:
        pass
    return page


def _redeem_dialog_present(page):
    """Hộp thoại nhập mã đã mở chưa.

    Trong phạm vi hộp thoại phải có cùng lúc: tiêu đề quy đổi
    (khớp một phần — Anh/Việt đều được), ít nhất một ô nhập,
    và nút xác nhận. Mục quy đổi trong Cài đặt chỉ có chữ,
    không có ô nhập lẫn nút xác nhận, nên không bị nhận nhầm.
    """
    scope = _dialog_scope(page)
    try:
        heading_found = False
        for title in (
            "Quy đổi mã",
            "Redeem invite code",
            "Redeem a code",
            "Redeem code",
        ):
            heading = scope.get_by_text(title, exact=False)
            if heading.count() and heading.first.is_visible():
                heading_found = True
                break
        if not heading_found:
            return False
        if scope.locator("input:visible").count() < 1:
            return False
        for name in ("Xác nhận", "Confirm", "Redeem"):
            button = scope.get_by_role(
                "button", name=name, exact=True
            )
            if button.count() and button.first.is_visible():
                return True
        return False
    except Exception:
        return False


def _fill_redeem_code(page, code):
    """Điền mã 6 ký tự vào hộp thoại quy đổi, có đọc lại kiểm chứng.

    Lần lượt thử: điền từng ô (6 ô riêng) → gõ phím như người
    dùng (bấm ô đầu rồi gõ, các ô tự nhảy) → điền một ô duy nhất
    chứa cả mã. Cách nào cho tổng số ký tự đọc lại đúng bằng độ
    dài mã thì coi như xong.
    """
    boxes = _dialog_scope(page).locator("input:visible")
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
        scope = _dialog_scope(page)
        button = None
        for name in ("Xác nhận", "Confirm", "Redeem"):
            candidate = scope.get_by_role(
                "button", name=name, exact=True
            )
            try:
                if candidate.count() and candidate.first.is_visible():
                    button = candidate.first
                    break
            except Exception:
                pass
        if button is None:
            _status("Quy đổi mã: không thấy nút Xác nhận.")
            return False
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
    const rectOf = (el) => {
        const r = el.getBoundingClientRect();
        return {x: r.x, y: r.y, width: r.width, height: r.height};
    };
    const isCode = (s) => /^[A-Z0-9]{6}$/.test((s || '').trim());
    // 1. Theo cau truc, khong theo ngon ngu: canvas nam cung
    // container voi span mang ma (title == chu hien thi).
    for (const el of document.querySelectorAll('[title]')) {
        const t = (el.getAttribute('title') || '').trim();
        if (!isCode(t)) continue;
        let node = el;
        for (let i = 0; i < 4 && node; i++) {
            const c = node.querySelector
                ? node.querySelector('canvas')
                : null;
            if (c) {
                const r = c.getBoundingClientRect();
                if (r.width > 0) return rectOf(c);
            }
            node = node.parentElement;
        }
    }
    // 2. Canvas dong vai tro nut trong panel Moi (role=button,
    // phu kin thanh ma) — dung cau truc HTML that nguoi dung gui.
    for (const c of document.querySelectorAll('canvas[role="button"]')) {
        const r = c.getBoundingClientRect();
        if (r.width > 0 && r.height >= 8) return rectOf(c);
    }
    // 3. Theo aria-label (ban b2338133) — nhan tieng Anh co the
    // khac tieng Viet nen chi la duong du phong.
    for (const c of document.querySelectorAll('canvas')) {
        if (
            /mã lời mời|invite code/i.test(
                c.getAttribute('aria-label') || ''
            )
        ) {
            const r = c.getBoundingClientRect();
            if (r.width > 0) return rectOf(c);
        }
    }
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
    const rectOfEl = (el) => el.getBoundingClientRect();
    // O cao that la mot thanh ngang (rong 150-650, cao 24-90);
    // phan tu khop nho nhat chi la mau chu ben trong — cao
    // dung mau chu thi phu khong du dien tich, ma khong hien.
    const bars = hits.filter((el) => {
        const r = rectOfEl(el);
        return r.width >= 150 && r.width <= 650 &&
            r.height >= 24 && r.height <= 90;
    });
    const pool = bars.length ? bars : hits;
    pool.sort((a, b) => {
        const ra = rectOfEl(a);
        const rb = rectOfEl(b);
        return bars.length
            ? rb.width * rb.height - ra.width * ra.height
            : ra.width * ra.height - rb.width * rb.height;
    });
    const r = rectOfEl(pool[0]);
    return {x: r.x, y: r.y, width: r.width, height: r.height};
}
"""


def _paint_slow(page, x_from, x_to, y):
    """Quét chuột CHẬM từ x_from tới x_to trên hàng y.

    Trang xử lý vị trí chuột theo từng khung hình (~60fps):
    quét nhanh là giữa hai khung hình chuột nhảy cả trăm px,
    cọ chỉ chấm được vài điểm rời — nhìn như tô mà lớp phủ
    gần như nguyên vẹn. Mỗi đoạn ~10px nghỉ ~15ms thì khung
    nào cũng bắt được điểm tô, như tay người tô thật.
    """
    distance = abs(x_to - x_from)
    steps = max(1, int(distance // 10))
    for index in range(1, steps + 1):
        page.mouse.move(x_from + (x_to - x_from) * index / steps, y)
        page.wait_for_timeout(15)


def _paint_reveal(page, box):
    """Một lượt tô mở lớp phủ "Scratch to reveal".

    Đưa chuột vào ô, con trỏ đổi thành hình tròn như một cây
    cọ; nhấn giữ rồi TÔ phủ kín cả ô — ngang từng hàng chậm
    rãi, rồi dọc từng cột phủ chéo. Lớp phủ bị xoá ở nơi cọ
    quét qua; tô đủ thì mã mời hiện ra và nút Copy referral
    link chuyển xanh, bấm được.
    """
    try:
        left = box["x"] + 4
        right = box["x"] + box["width"] - 4
        top = box["y"] + 4
        bottom = box["y"] + box["height"] - 4
        # Đưa chuột vào giữa ô trước để con trỏ cọ xuất hiện.
        page.mouse.move((left + right) / 2, (top + bottom) / 2)
        page.wait_for_timeout(400)
        page.mouse.down()
        page.wait_for_timeout(200)
        rows = max(4, int((bottom - top) // 6) + 1)
        for index in range(rows):
            if rows > 1:
                y = top + (bottom - top) * index / (rows - 1)
            else:
                y = (top + bottom) / 2
            if index % 2 == 0:
                page.mouse.move(left, y)
                _paint_slow(page, left, right, y)
            else:
                page.mouse.move(right, y)
                _paint_slow(page, right, left, y)
        # Phủ chéo bằng các cột dọc để không sót vệt nào.
        columns = max(3, int((right - left) // 60) + 1)
        for index in range(columns):
            if columns > 1:
                x = left + (right - left) * index / (columns - 1)
            else:
                x = (left + right) / 2
            page.mouse.move(x, top)
            steps = max(1, int((bottom - top) // 10))
            for step in range(1, steps + 1):
                page.mouse.move(
                    x, top + (bottom - top) * step / steps
                )
                page.wait_for_timeout(15)
        page.wait_for_timeout(200)
        page.mouse.up()
        page.wait_for_timeout(600)
        return True
    except Exception:
        return False


_INVITE_CODE_DOM_JS = """() => {
    const isCode = (s) => /^[A-Z0-9]{6}$/.test((s || '').trim());
    // Đúng cấu trúc thật của panel (người dùng gửi HTML): mã
    // nằm trong một span có title chính là mã, chữ hiển thị
    // cũng là mã — lớp canvas chỉ phủ lên trên để che mắt.
    for (const el of document.querySelectorAll('[title]')) {
        const t = (el.getAttribute('title') || '').trim();
        if (isCode(t) && (el.textContent || '').trim() === t) {
            return t;
        }
    }
    // Dòng trạng thái cho trình đọc màn hình khi mã đã lộ:
    // "Đã tiết lộ mã lời mời: ATGXDX" / "Invite code revealed:".
    for (const el of document.querySelectorAll('span')) {
        if (el.childElementCount > 0) continue;
        const m = (el.textContent || '').match(
            /(?:mã lời mời|invite code)[^A-Z0-9]{0,12}([A-Z0-9]{6})/i
        );
        if (m) return m[1].toUpperCase();
    }
    for (const el of document.querySelectorAll(
        '[data-pel-impression*="invite_code"]'
    )) {
        const t = (el.textContent || '').trim();
        if (isCode(t)) return t;
        const title = (el.getAttribute('title') || '').trim();
        if (isCode(title)) return title;
    }
    // Cả thân trang: dòng thông báo tiết lộ mã có thể nằm ở
    // phần tử không phải span lá.
    const bodyText = document.body
        ? (document.body.textContent || '')
        : '';
    const ann = bodyText.match(
        /(?:mã lời mời|invite code)[^A-Z0-9]{0,20}([A-Z0-9]{6})/i
    );
    if (ann) return ann[1].toUpperCase();
    return '';
}"""


def _scan_once(page):
    """Đọc mã mời một lần: đọc thẳng từ DOM trước, quét panel sau.

    Mã nằm sẵn trong DOM của panel (span mang title là mã) —
    đọc được là khỏi cần copy hay tô gì cả. Bộ quét panel là
    dự phòng cho cấu trúc khác.
    """
    for script in (_INVITE_CODE_DOM_JS, _HARVEST_CODE_JS):
        try:
            found = page.evaluate(script) or ""
        except Exception:
            found = ""
        found = str(found).strip().upper()
        if re.fullmatch(r"[A-Z0-9]{6}", found):
            return found
    return ""


def _wait_panel_ready(page):
    """Chờ panel Invite mở xong; trả mã ngay nếu quét thấy sẵn.

    Panel sẵn sàng khi thấy lớp phủ che mã hoặc nút copy.
    Tối đa 10 giây (panel mở chậm qua proxy).
    """
    deadline = time.monotonic() + 10
    while time.monotonic() < deadline:
        if page.is_closed():
            return ""
        code = _scan_once(page)
        if code:
            return code
        try:
            if page.evaluate(_SCRATCH_RECT_JS):
                return ""
        except Exception:
            pass
        for name in (
            "Copy referral link",
            "Sao chép liên kết giới thiệu",
        ):
            button = page.get_by_role(
                "button", name=name, exact=False
            )
            try:
                if button.count() and button.first.is_visible():
                    return ""
            except Exception:
                pass
        page.wait_for_timeout(700)
    return ""


def _reveal_by_painting(page):
    """Tô mở lớp phủ cho tới khi lấy được mã (tối đa 4 lượt).

    Mỗi lượt tô xong đều kiểm chứng ngay: quét panel tìm mã
    (mã hiện ra sau lớp phủ là chữ đọc được), rồi bấm Copy
    referral link (nút đã xanh khi mã hiện) và bóc mã khỏi
    nội dung copy.
    """
    for attempt in range(4):
        try:
            box = page.evaluate(_SCRATCH_RECT_JS)
        except Exception:
            box = None
        if not box:
            return ""
        # Cuộn ô cào vào giữa tầm nhìn trước khi tô — ô nằm
        # lệch khỏi màn hình thì chuột tô vào hư không.
        try:
            page.evaluate(
                """(pt) => {
                    const el = document.elementFromPoint(
                        pt.x, pt.y
                    );
                    if (el) {
                        el.scrollIntoView({
                            block: 'center',
                            inline: 'center',
                        });
                    }
                }""",
                {
                    "x": box["x"] + box["width"] / 2,
                    "y": box["y"] + box["height"] / 2,
                },
            )
            page.wait_for_timeout(400)
            box = page.evaluate(_SCRATCH_RECT_JS) or box
        except Exception:
            pass
        if not _paint_reveal(page, box):
            return ""
        _status(
            "Đã tô lớp phủ che mã mời "
            f"(lần {attempt + 1}, ô {int(box['width'])}x"
            f"{int(box['height'])})."
        )
        code = _scan_once(page)
        if code:
            return code
        code = _harvest_via_copy_link(page)
        if code:
            return code
    return ""


def _scan_panel_code(page, seconds):
    """Quét chữ trong panel Mời tìm mã 6 ký tự, chờ tối đa vài giây."""
    deadline = time.monotonic() + seconds
    while time.monotonic() < deadline:
        if page.is_closed():
            return ""
        code = _scan_once(page)
        if code:
            return code
        page.wait_for_timeout(700)
    return ""


_INVITE_DEBUG_JS = """() => {
    const info = {url: location.href};
    const canvas = [...document.querySelectorAll('canvas')].find(
        (c) => /mã lời mời|invite code/i.test(
            c.getAttribute('aria-label') || ''
        )
    ) || [...document.querySelectorAll('canvas[role="button"]')][0];
    info.canvas_found = !!canvas;
    let node = canvas;
    for (let i = 0; i < 4 && node && node.parentElement; i++) {
        node = node.parentElement;
    }
    if (!node) {
        const anchor = [...document.querySelectorAll('button')].find(
            (b) => /copy referral link|sao chép liên kết/i.test(
                b.textContent || ''
            )
        ) || [...document.querySelectorAll('*')].find(
            (e) => e.childElementCount <= 3 &&
                /invite friends|mời bạn bè/i.test(e.textContent || '')
        );
        node = anchor;
        for (let i = 0; i < 6 && node && node.parentElement; i++) {
            node = node.parentElement;
        }
    }
    info.panel_html = node ? node.outerHTML.slice(0, 20000) : '';
    return JSON.stringify(info);
}
"""


def _dump_invite_debug(page):
    """Lưu cấu trúc panel Mời khi lấy mã thất bại để chẩn đoán.

    File invite_debug.html nằm cùng thư mục với accounts.xlsx.
    Cấu trúc panel của tài khoản MỚI (chưa tiết lộ mã) chỉ xem
    được ở máy người dùng — có nó mới sửa trúng được thay vì
    đoán mò theo ảnh chụp.
    """
    try:
        raw = page.evaluate(_INVITE_DEBUG_JS) or ""
        data_dir = (
            Path.home() / "AppData" / "Local" / "MuseAIHelper"
        )
        data_dir.mkdir(parents=True, exist_ok=True)
        target = data_dir / "invite_debug.html"
        target.write_text(str(raw), encoding="utf-8")
        _status(
            "Không lấy được mã mời — đã lưu cấu trúc panel vào "
            "invite_debug.html (cùng thư mục với accounts.xlsx); "
            "gửi file đó để chẩn đoán đúng chỗ."
        )
    except Exception:
        pass


def _harvest_invite_code(page):
    """Lấy mã mời của chính tài khoản này từ panel "Mời bạn bè".

    Bấm nút Mời/Invite ở góc phải trên → tô mở lớp phủ
    "Scratch to reveal" che mã (chưa tô thì nút copy còn mờ,
    copy không ra mã) → bấm "Copy referral link" và bóc mã
    khỏi nội dung copy. Dự phòng cuối: quét chữ trong panel.
    Không lấy được thì trả "" — tài khoản vẫn được lưu bình
    thường.
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
        # Chờ panel mở xong (hoặc nhận mã ngay nếu đã hiện sẵn),
        # thử copy — tài khoản đã mở mã thì nút copy hiệu lực sẵn.
        code = _wait_panel_ready(page)
        if not code:
            code = _harvest_via_copy_link(page)
        if not code:
            # Chưa tô mở lớp phủ thì nút copy còn mờ, copy
            # không ra mã: đưa chuột vào ô "Scratch to reveal",
            # nhấn giữ và tô phủ kín cho mã hiện ra, rồi copy
            # lại sau mỗi lượt tô.
            code = _reveal_by_painting(page)
        if not code:
            code = _scan_panel_code(page, 6)
    except Exception:
        code = ""
    finally:
        try:
            page.keyboard.press("Escape")
        except Exception:
            pass
    if not code:
        _dump_invite_debug(page)
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


# Gọi API native: các header đúng như hàm gọi API của trang tự
# làm (aT trong JS của muse.ai) — CSRF token, X-Hatch-Waterfall-Id
# (một UUID cho cả lượt thử, trang tự sinh khi bắt đầu lượt đăng
# ký) và X-Hatch-Caa-Reg-Entry-Point: login_home (điểm vào mặc
# định khi đăng ký từ trang chính). Header activation-source trang
# chỉ gửi khi có nguồn kích hoạt (luồng đăng ký thường không có)
# nên không gửi.
_API_FETCH_JS = """([url, payload, csrfToken, waterfallId]) => {
    const headers = {
        "Content-Type": "application/json",
        "Accept": "application/json",
        "X-Hatch-Caa-Reg-Entry-Point": "login_home"
    };
    if (csrfToken) {
        headers["X-Hatch-CSRF-Token"] = csrfToken;
    }
    if (waterfallId) {
        headers["X-Hatch-Waterfall-Id"] = waterfallId;
    }
    return fetch(url, {
        method: "POST",
        headers: headers,
        credentials: "same-origin",
        body: JSON.stringify(payload)
    }).then(async (response) => ({
        status: response.status,
        text: await response.text()
    }));
}"""


# Lấy CSRF token: POST /restart (không cần token) rồi đọc
# csrf_token trong JSON trả về — đúng cách trang đăng ký tự làm
# khi khởi tạo luồng auth (đã đọc trong JS của muse.ai).
_API_RESTART_JS = """([url, waterfallId]) => {
    const headers = {
        "Content-Type": "application/json",
        "Accept": "application/json"
    };
    if (waterfallId) {
        headers["X-Hatch-Waterfall-Id"] = waterfallId;
    }
    return fetch(url, {
    method: "POST",
    headers: headers,
    credentials: "same-origin",
    cache: "no-store",
    body: "{}"
}).then(async (response) => ({
    status: response.status,
    text: await response.text()
}))}"""


def _api_post(page, path, payload, csrf_token=None, waterfall_id=None):
    """Gọi một API native của muse.ai TỪ TRONG trình duyệt.

    fetch cùng origin nên request mang đúng cookie/phiên của
    trang — khác hẳn gọi từ ngoài bằng curl (bị 403 Forbidden
    ngay từ send-otp, đã thử 2026-11-12). csrf_token (lấy qua
    /restart) được gửi trong header X-Hatch-CSRF-Token — thiếu
    nó server trả 403 csrfFailed (gặp live 2026-11-12).
    waterfall_id (một UUID cho cả lượt thử) và điểm vào
    login_home được gửi như trang tự làm. Trả (status, data,
    raw_text); data là JSON đã parse hoặc None khi thân không
    phải JSON. Lỗi evaluate (trang đóng, fetch bị chặn) ném
    RuntimeError để phía gọi rơi về luồng DOM.
    """
    try:
        result = page.evaluate(
            _API_FETCH_JS,
            [NATIVE_AUTH_BASE + path, payload, csrf_token, waterfall_id],
        )
    except Exception as error:
        raise RuntimeError(
            f"Không gọi được API {path} từ trình duyệt "
            f"({_short_error(error)})"
        ) from error
    status = result.get("status", 0) if isinstance(result, dict) else 0
    text = result.get("text", "") if isinstance(result, dict) else ""
    data = None
    if text:
        try:
            data = json.loads(text)
        except (json.JSONDecodeError, ValueError):
            data = None
    return status, data, text


def _fetch_csrf_token(page, waterfall_id=None):
    """Lấy CSRF token cho luồng auth native qua POST /restart.

    Trả token (str) hoặc None khi lỗi (đã ghi log lý do). Token
    này trang đăng ký cũng tự lấy khi khởi tạo luồng auth — đọc
    từ JSON trả về của /restart, không đọc cookie. waterfall_id
    (nếu có) được gửi kèm như trang tự làm.
    """
    try:
        result = page.evaluate(
            _API_RESTART_JS,
            [NATIVE_AUTH_BASE + "/restart", waterfall_id],
        )
    except Exception as error:
        _status(
            "Không gọi được /restart để lấy CSRF token "
            f"({_short_error(error)})."
        )
        return None
    status = result.get("status", 0) if isinstance(result, dict) else 0
    text = result.get("text", "") if isinstance(result, dict) else ""
    if not 200 <= status < 300:
        _status(
            f"Gọi /restart trả HTTP {status} ({_preview_body(text)})."
        )
        return None
    try:
        data = json.loads(text)
    except (json.JSONDecodeError, ValueError):
        _status("Gọi /restart trả về không phải JSON.")
        return None
    token = data.get("csrf_token") if isinstance(data, dict) else None
    if not isinstance(token, str) or not token:
        _status("Gọi /restart không trả về csrf_token.")
        return None
    return token


# register-completion trả 403 registrationBlocked: Meta không cho
# tạo tài khoản trên IP này — cùng bản chất với hộp thoại
# "Something went wrong" ở luồng bấm nút, nên được báo PROXY_BURN
# như nhau (gặp live 2026-11-12: 4/5 luồng cùng lúc).
REGISTRATION_BLOCKED = "registration_blocked"


def _api_post_authed(page, path, payload, token_box):
    """Gọi API native kèm CSRF token trong token_box (dict có khóa
    "token" và "waterfall_id"). Gặp lỗi csrfFailed thì lấy token
    mới qua /restart rồi thử lại đúng một lần (server tự gợi ý
    "Refresh the page and try again")."""
    status, data, text = _api_post(
        page, path, payload, token_box.get("token"),
        token_box.get("waterfall_id"),
    )
    if status == 403 and "csrfFailed" in (text or ""):
        _status("CSRF token hết hiệu lực — lấy token mới rồi thử lại...")
        fresh = _fetch_csrf_token(page, token_box.get("waterfall_id"))
        if fresh:
            token_box["token"] = fresh
            status, data, text = _api_post(
                page, path, payload, fresh, token_box.get("waterfall_id")
            )
    return status, data, text


# Đồng bộ phiên web sau khi hoàn tất hồ sơ qua API — đúng cách
# trang đăng ký tự làm (đã đọc trong JS của muse.ai): POST form ẩn
# tới /hatch/session/bootstrap/ kèm frl_access_token, kèm đường
# dẫn đích trong sessionStorage. Server đặt cookie phiên rồi
# chuyển tới trang tiếp theo (checkpoint nhập thẻ hoặc app). Không
# có bước này thì trình duyệt chưa có phiên — mở lại trang chỉ thấy
# khung đăng nhập (gặp live 2026-11-12).
_API_BOOTSTRAP_JS = """([frlToken, returnTo]) => {
    if (sessionStorage.getItem("auth_meta_session_bootstrap_attempted") === "1") {
        return "already_attempted";
    }
    sessionStorage.setItem("auth_meta_session_bootstrap_attempted", "1");
    if (returnTo && returnTo !== "/") {
        sessionStorage.setItem("auth_meta_post_login_return_to", returnTo);
    } else {
        sessionStorage.removeItem("auth_meta_post_login_return_to");
    }
    const form = document.createElement("form");
    form.method = "POST";
    form.action = new URL("/hatch/session/bootstrap/", window.location.origin).toString();
    form.enctype = "application/x-www-form-urlencoded";
    form.hidden = true;
    const input = document.createElement("input");
    input.type = "hidden";
    input.name = "frl_access_token";
    input.value = frlToken;
    form.appendChild(input);
    document.body.appendChild(form);
    form.submit();
    return "submitted";
}"""


def _resolve_redirect_path(redirect_to):
    """Chuẩn hoá redirect_to thành đường dẫn nội bộ, như code web
    tự làm: phải là chuỗi bắt đầu bằng "/" (không phải "//"), dài
    ≤ 4096, không chứa "\\"; còn không thì "/"."""
    if not isinstance(redirect_to, str):
        return "/"
    value = redirect_to.strip()
    if (
        not value
        or len(value) > 4096
        or "\\" in value
        or not value.startswith("/")
        or value.startswith("//")
    ):
        return "/"
    return value


def _bootstrap_session(page, frl_token, redirect_to):
    """Đồng bộ phiên web bằng form bootstrap (xem _API_BOOTSTRAP_JS).
    Trả True khi đã gửi form (trang sẽ tự chuyển tới bước tiếp
    theo), False khi lỗi hoặc đã đồng bộ rồi trong tab này (phía
    gọi mở thẳng trang đích thay thế). Token phiên không bao giờ
    được ghi ra log."""
    path = _resolve_redirect_path(redirect_to)
    try:
        result = page.evaluate(_API_BOOTSTRAP_JS, [frl_token, path])
    except Exception as error:
        _status(
            "Không gửi được form đồng bộ phiên "
            f"({_short_error(error)})."
        )
        return False
    if result != "submitted":
        return False
    _status("Đã gửi form đồng bộ phiên — đang chờ trang tiếp theo...")
    page.wait_for_timeout(1500)
    try:
        page.wait_for_load_state("domcontentloaded", timeout=30000)
    except Exception:
        pass
    return True


def _preview_body(text, limit=200):
    """Một dòng xem trước ngắn của thân phản hồi lỗi (để ghi log).

    Chỉ dùng cho phản hồi LỖI (thân thường là mã/thông báo lỗi).
    Phản hồi thành công không bao giờ ghi thân ra log vì có thể
    chứa token — chỉ ghi tên trường.
    """
    return (text or "").replace("\n", " ").replace("\r", " ")[:limit]


def _api_response_has_challenge(data, text):
    """API có yêu cầu kiểm tra thêm (captcha/thử thách) không?

    Chỉ soi tên trường và đoạn text ngắn — không đọc giá trị nhạy
    cảm. Gặp thử thách thì luồng API dừng lại để luồng DOM mở
    trình duyệt cho người dùng xử lý bằng tay, không bỏ qua.
    """
    haystack = ""
    if isinstance(data, dict):
        haystack += " ".join(str(key) for key in data.keys()).lower()
        for key in ("error", "message", "code"):
            value = data.get(key)
            if isinstance(value, str):
                haystack += " " + value.lower()
    haystack += " " + (text or "")[:500].lower()
    return any(
        word in haystack
        for word in (
            "captcha",
            "challenge",
            "verify you are human",
            "human verification",
        )
    )


def _wait_for_code_from_mail(page, mailbox, existing_ids):
    """Chờ mã OTP trong hộp thư, không cần form mã trên trang
    (dùng cho luồng API — trang vẫn đang ở bước email). Nhịp và
    thời gian chờ giống hệt _wait_for_verification_code."""
    deadline = time.monotonic() + CODE_TIMEOUT_SECONDS
    warned = False
    while time.monotonic() < deadline:
        if page.is_closed():
            raise RuntimeError("Cửa sổ Chromium đã đóng trước khi xác minh.")
        try:
            code = _verification_code_from_new_messages(mailbox, existing_ids)
        except Exception:
            if not warned:
                _status("Kiểm tra hộp thư bị rớt kết nối — đang thử lại...")
                warned = True
            code = None
        if code:
            return code
        page.wait_for_timeout(4000)
    raise TimeoutError(
        f"Không tìm thấy mã xác minh mới trong hộp thư "
        f"{mailbox['provider']} sau {CODE_TIMEOUT_SECONDS} giây."
    )


def _try_send_otp_api(page, email, token_box):
    """Gửi OTP qua API native. Trả True khi đã gửi, False khi lỗi
    (ghi log lý do) để phía gọi bấm nút Continue trên form đã điền
    sẵn. Form email phải được điền trước khi gọi hàm này."""
    try:
        status, data, text = _api_post_authed(
            page, "/send-otp", {"contact_point": email, "region": "US"},
            token_box,
        )
    except Exception as error:
        _status(
            "Gọi API gửi OTP không được "
            f"({_short_error(error)}) — dùng nút bấm."
        )
        return False
    if _api_response_has_challenge(data, text):
        _status(
            "API gửi OTP yêu cầu kiểm tra thêm — dùng nút bấm để "
            "bạn xử lý trên trình duyệt."
        )
        return False
    if not 200 <= status < 300:
        _status(
            f"API gửi OTP trả HTTP {status} "
            f"({_preview_body(text)}) — dùng nút bấm thay thế."
        )
        return False
    if isinstance(data, dict):
        keys = ", ".join(sorted(data.keys())) or "trống"
        _status(f"API gửi OTP OK (trả về: {keys}).")
    return True


def _try_confirm_otp_api(page, code, token_box):
    """Xác nhận OTP qua API. Trả (tos_content_version, requires_name)
    khi ra bước hoàn tất hồ sơ, None khi lỗi (ghi log lý do) để phía
    gọi rơi về điền mã vào form (form tự gửi khi đủ số).
    requires_name: server có bắt nhập họ tên ở bước hoàn tất hồ sơ
    không — trang tự đọc cờ này từ kết quả xác nhận OTP và chỉ gửi
    first_name/last_name khi cờ bật (thiếu tên khi bị bắt thì
    register-completion trả 400 invalidRequest)."""
    try:
        status, data, text = _api_post_authed(
            page, "/confirm-otp",
            {"otp_code": code, "remember_device": True}, token_box,
        )
    except Exception as error:
        _status(
            "Gọi API xác nhận OTP không được "
            f"({_short_error(error)}) — dùng nút bấm."
        )
        return None
    if _api_response_has_challenge(data, text):
        _status(
            "API xác nhận OTP yêu cầu kiểm tra thêm — dùng nút bấm "
            "để bạn xử lý trên trình duyệt."
        )
        return None
    if not 200 <= status < 300:
        _status(
            f"API xác nhận OTP trả HTTP {status} "
            f"({_preview_body(text)}) — dùng nút bấm thay thế."
        )
        return None
    if not isinstance(data, dict):
        _status("API xác nhận OTP trả về không phải JSON — dùng nút bấm.")
        return None
    state = data.get("registration_state")
    tos = data.get("tos") or {}
    tos_version = tos.get("content_version") if isinstance(tos, dict) else None
    requires_name = bool(data.get("requires_name"))
    _status(
        "API xác nhận OTP OK (trạng thái: "
        f"{state}; cần tên: {'có' if requires_name else 'không'})."
    )
    if state != "completion_required" or not tos_version:
        _status(
            "Xác nhận OTP xong nhưng không ra bước hoàn tất hồ "
            "sơ — dùng nút bấm."
        )
        return None
    return tos_version, requires_name


def _try_register_completion_api(page, tos_version, token_box,
                                 birth_date=None, include_names=False):
    """Hoàn tất hồ sơ qua API. Trả dict JSON trả về khi xong (phía
    gọi đọc redirect_to / checkpoint_url / auth_meta_session trong
    đó để đồng bộ phiên web), REGISTRATION_BLOCKED khi Meta chặn
    tạo tài khoản trên IP này, None khi lỗi khác (ghi log lý do).
    birth_date: ngày sinh dùng cho API; không có thì tự sinh ngẫu
    nhiên (trên 18 tuổi). include_names: kèm first_name/last_name
    ("Long"/"long", như form điền tay) — chỉ bật khi kết quả xác
    nhận OTP báo requires_name, đúng như trang tự làm; thiếu tên
    khi bị bắt thì server trả 400 invalidRequest (gặp live
    2026-11-12). Luồng API KHÔNG điền form hồ sơ trước — tải lại
    trang để lấy form làm mất trạng thái luồng (chưa có phiên) và
    trang về khung đăng nhập (gặp live 2026-11-12)."""
    if birth_date is None:
        birth_date = _random_birth_date()
    _status(
        "Đang hoàn tất hồ sơ qua API (ngày sinh "
        f"{birth_date.day:02d}/{birth_date.month:02d}/{birth_date.year}"
        f"{'; kèm tên' if include_names else ''})..."
    )
    payload = {
        "date_of_birth": birth_date.isoformat(),
        "tos_content_version": tos_version,
        "remember_device": True,
    }
    if include_names:
        payload["first_name"] = "Long"
        payload["last_name"] = "long"
    try:
        status, data, text = _api_post_authed(
            page,
            "/register-completion",
            payload,
            token_box,
        )
    except Exception as error:
        _status(
            "Gọi API hoàn tất hồ sơ không được "
            f"({_short_error(error)})."
        )
        return None
    if _api_response_has_challenge(data, text):
        _status(
            "API hoàn tất hồ sơ yêu cầu kiểm tra thêm — mở lại trang "
            "để bạn xử lý trên trình duyệt."
        )
        return None
    if not 200 <= status < 300:
        if (
            status == 403
            and isinstance(data, dict)
            and data.get("type") == "registrationBlocked"
        ):
            _status(
                "API hoàn tất hồ sơ bị chặn tạo tài khoản "
                "(registrationBlocked) — IP này bị Meta chặn."
            )
            return REGISTRATION_BLOCKED
        _status(
            f"API hoàn tất hồ sơ trả HTTP {status} "
            f"({_preview_body(text)})."
        )
        return None
    if isinstance(data, dict) and data:
        keys = ", ".join(sorted(data.keys()))
        nxt = data.get("redirect_to") or "(không có)"
        _status(f"API hoàn tất hồ sơ OK (trả về: {keys}; chuyển tới: {nxt}).")
    _status("Đã hoàn tất hồ sơ qua API.")
    return data if isinstance(data, dict) else {}


def _try_fill_profile_form(page, timeout_ms=20000):
    """Cố điền form hồ sơ sau khi OTP đã xác nhận qua API: mở lại
    trang để app hiện form, điền ngày sinh (+ tên nếu hỏi). Trả
    ngày sinh đã điền khi xong, None khi form không hiện hoặc điền
    lỗi (lúc đó luồng API gửi register-completion mà không cần
    form). Không ném lỗi.

    KHÔNG DÙNG NỮA trong luồng API (từ 2026-11-12): tải lại trang
    ở bước này làm mất trạng thái luồng completion_required trong
    bộ nhớ của app — tài khoản chưa có phiên nên trang khởi động
    lại như chưa đăng nhập và chỉ còn khung đăng nhập, form hồ sơ
    không hiện để mà điền. Giữ hàm lại để khỏi vỡ code/test cũ
    đang tham chiếu tới nó."""
    try:
        page.goto(
            SIGNUP_URL,
            wait_until="domcontentloaded",
            timeout=PAGE_TIMEOUT_MS,
        )
    except Exception as error:
        _status(
            "Không mở lại được trang để điền hồ sơ "
            f"({_short_error(error)}) — gửi API luôn."
        )
        return None
    try:
        return _fill_profile(page, timeout_ms=timeout_ms)
    except Exception as error:
        _status(
            "Không điền được form hồ sơ "
            f"({_short_error(error)}) — gửi API luôn."
        )
        return None


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
    try:
        page.goto(
            SIGNUP_URL,
            wait_until="domcontentloaded",
            timeout=PAGE_TIMEOUT_MS,
        )
    except Exception as error:
        # Mọi lỗi ở riêng bước mở trang đều có nghĩa là đường
        # mạng hiện tại không tải nổi muse.ai (proxy quá chậm /
        # đã chết, hoặc VPN mất kết nối ở chế độ IP máy). Dịch
        # sang một câu ngắn gọn để dòng RESULT đọc được ngay,
        # kèm dòng lỗi gốc đầu tiên.
        if DIRECT_MODE:
            detail = (
                "Không tải được trang muse.ai — mạng/VPN quá chậm "
                "hoặc mất kết nối."
            )
        else:
            detail = (
                "Không tải được trang muse.ai qua proxy — proxy quá "
                "chậm hoặc đã chết."
            )
        raise RuntimeError(
            f"{detail} ({_short_error(error)})"
        ) from error
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


def _capsolver_key():
    """CapSolver API key đã lưu (Windows Credential Manager).

    Không bao giờ in key ra log. Trả "" khi chưa lưu.
    """
    try:
        value = keyring.get_password(CAPSOLVER_SERVICE, CAPSOLVER_KEY_FIELD)
    except Exception:
        value = None
    return value or ""


def _load_selfie_path():
    """Đường dẫn ảnh selfie (app chính ghi selfie.txt).

    Trả về Path khi file tồn tại, None khi chưa chọn hoặc file
    không còn. Một ảnh dùng cho mọi tài khoản trong lượt chạy.
    """
    try:
        raw = paths.data_file("selfie.txt").read_text(
            encoding="utf-8"
        ).strip()
    except OSError:
        return None
    if not raw:
        return None
    path = Path(raw.splitlines()[0].strip())
    return path if path.is_file() else None


def _random_meta_password(length=16):
    """Mật khẩu ngẫu nhiên cho form đăng ký Meta.

    Mật khẩu này CHỈ để điền form — không lưu ở đâu cả (theo
    quyết định của chủ tool: đăng nhập Muse sau này dùng mã
    OTP qua mail, không cần mật khẩu) và không in ra log.
    """
    alphabet = string.ascii_letters + string.digits + "!@#$%^&*"
    while True:
        password = "".join(
            secrets.choice(alphabet) for _ in range(length)
        )
        if (
            any(char.islower() for char in password)
            and any(char.isupper() for char in password)
            and any(char.isdigit() for char in password)
            and any(char in "!@#$%^&*" for char in password)
        ):
            return password


def _meta_write_username(username):
    """Ghi username Meta tự sinh ra file để app chính lưu vào sheet."""
    if not username:
        return
    try:
        paths.data_file(
            f"meta_username{LANE_SUFFIX}.txt"
        ).write_text(username + "\n", encoding="utf-8")
    except OSError:
        pass


def _meta_logged_in(page):
    """Đã đăng nhập Meta AI chưa.

    Tiêu chí: đang ở meta.ai (không còn ở auth.meta.com) và
    không còn nút "Log in" — khi đăng nhập, nút này được thay
    bằng avatar/menu người dùng. Ô chat "Ask Meta AI anything..."
    KHÔNG dùng làm tín hiệu vì trang chưa đăng nhập cũng có nó.
    """
    try:
        url = page.url.lower()
    except Exception:
        return False
    if META_AUTH_HOST in url or "meta.ai" not in url:
        return False
    try:
        login = page.get_by_role("button", name="Log in", exact=True)
        if login.count() and login.first.is_visible():
            return False
    except Exception:
        pass
    return True


def _page_closed(page):
    """Cửa sổ Chromium đã đóng chưa (người dùng đóng tay / bị giết).

    Các vòng chờ của luồng Meta phải fail nhanh khi cửa sổ đã
    đóng, thay vì chờ hết 90 giây rồi báo "nút không hiện".
    """
    try:
        return page.is_closed()
    except Exception:
        return True


def _visible_candidates(locator):
    """Các phần tử khớp đang hiện + bấm được (locator đơn từng cái).

    Playwright strict mode: wait_for/is_enabled/click trên
    locator khớp NHIỀU phần tử đều lỗi ngay — trang meta.ai có
    2 nút "Log in", và các trang auth.meta.com cũng có nút trùng.
    Quét từng phần tử khớp, chỉ giữ cái visible + enabled, trả
    về locator đơn (.nth) để mọi thao tác sau không bao giờ gặp
    strict mode nữa.
    """
    try:
        count = locator.count()
    except Exception:
        return []
    found = []
    for index in range(count):
        candidate = locator.nth(index)
        try:
            if candidate.is_visible() and candidate.is_enabled():
                found.append(candidate)
        except Exception:
            continue
    return found


def _first_visible(locator):
    """Phần tử đầu tiên đang hiện + bấm được, None khi không có."""
    found = _visible_candidates(locator)
    return found[0] if found else None


def _in_viewport_candidates(page, locator):
    """Các phần tử khớp đang hiện, bấm được VÀ nằm trong khung hình.

    is_visible() của Playwright chỉ kiểm tra hộp bao + kiểu hiển
    thị — nút mobile của meta.ai nằm ngoài khung hình (off-canvas)
    vẫn "visible", bấm vào không chuyển trang (live 2026-11-13
    09:45: bấm xong chờ 45 giây không sang auth.meta.com). Lọc
    thêm bằng bounding_box() trong viewport để chỉ chọn nút thật
    sự đang hiện trên màn hình.
    """
    try:
        viewport = page.viewport_size or {"width": 1280, "height": 720}
    except Exception:
        viewport = {"width": 1280, "height": 720}
    found = []
    for candidate in _visible_candidates(locator):
        try:
            box = candidate.bounding_box()
        except Exception:
            continue
        if not box:
            continue
        if (
            box["x"] >= 0
            and box["y"] >= 0
            and box["x"] + box["width"] <= viewport["width"]
            and box["y"] + box["height"] <= viewport["height"]
        ):
            found.append(candidate)
    return found


def _first_in_viewport(page, locator):
    """Phần tử đầu tiên đang hiện trong khung hình, None khi không có."""
    found = _in_viewport_candidates(page, locator)
    return found[0] if found else None


def _meta_click(page, locator, label):
    """Bấm một nút của luồng Meta, chịu được nút trùng.

    Như _click_when_ready nhưng không gọi wait_for/click trên
    locator gốc (khớp nhiều phần tử là lỗi strict mode) — chờ
    tới khi có phần tử hiện + bấm được rồi bấm đúng phần tử đó.
    Ưu tiên nút đang nằm trong khung hình, fallback sang nút
    visible ngoài khung hình (nút dưới nếp gấp trang — click sẽ
    tự cuộn tới). Dùng cho các bước bắt buộc của luồng Meta
    (Continue, Confirm...), nơi nút PHẢI xuất hiện.
    """
    deadline = time.monotonic() + BUTTON_TIMEOUT_MS / 1000
    while True:
        if _page_closed(page):
            raise RuntimeError(
                "Cửa sổ Chromium đã đóng trước khi bấm được "
                f"nút {label}."
            )
        target = _first_in_viewport(page, locator) or _first_visible(
            locator
        )
        if target is not None:
            page.wait_for_timeout(UI_SETTLE_MS)
            target.click(timeout=BUTTON_TIMEOUT_MS)
            return
        if time.monotonic() >= deadline:
            raise RuntimeError(
                f"Nút {label} không hiện sau "
                f"{BUTTON_TIMEOUT_MS // 1000} giây."
            )
        page.wait_for_timeout(250)


def _meta_try_login_buttons(page, candidates):
    """Thử bấm từng nút Log in tới khi trang chuyển sang auth.meta.com.

    Trả về True khi chuyển trang thành công, False khi đã thử hết
    mà không nút nào chuyển được. Mỗi nút được 8 giây để chuyển
    trang — bấm đúng nút thì chuyển gần như ngay.
    """
    for target in candidates:
        try:
            page.wait_for_timeout(UI_SETTLE_MS)
            target.click(timeout=BUTTON_TIMEOUT_MS)
        except Exception:
            continue
        try:
            page.wait_for_url(f"**{META_AUTH_HOST}**", timeout=8000)
            return True
        except Exception:
            continue
    return False


def _meta_testid_login(page):
    """Thử bấm nút Log in theo data-testid="login-button".

    Nút Log in thật trong header meta.ai mang testid này (quan
    sát DOM 2026-11-13). Thử TỪNG nút khớp testid đang hiện +
    bấm được (không chỉ .first — phòng khi có bản mobile trùng
    testid đứng trước). Trả về True khi bấm xong trang chuyển
    sang auth.meta.com, False khi không nút nào chuyển được.
    """
    try:
        testid = page.get_by_test_id("login-button")
        count = testid.count()
    except Exception:
        return False
    for index in range(count):
        try:
            target = testid.nth(index)
            if not (target.is_visible() and target.is_enabled()):
                continue
        except Exception:
            continue
        try:
            page.wait_for_timeout(UI_SETTLE_MS)
            target.click(timeout=BUTTON_TIMEOUT_MS)
            page.wait_for_url(f"**{META_AUTH_HOST}**", timeout=8000)
            return True
        except Exception:
            continue
    return False


def _meta_click_login(page):
    """Bấm Log in ở meta.ai: thử đi thử lại tới khi chuyển trang.

    Nút Log in thật trong header có data-testid="login-button"
    (quan sát DOM 2026-11-13) — thử nó trước mỗi vòng; nếu
    testid không còn thì thử từng nút trùng tên trong khung
    hình. Bấm xong chờ chuyển sang auth.meta.com (mỗi nút 8
    giây).

    Vì sao phải thử lại: nút này là <button> thường, không có
    href — chuyển trang nhờ handler JS của React. Giao diện
    meta.ai hiện ra TRƯỚC khi JS hydrate xong, bấm lúc đó nút
    chưa có tác dụng (live 2026-11-13 10:08: mở trang 12 giây,
    bấm 1 nút, không chuyển). Qua proxy chậm thì hydrate càng
    lâu. Nên thử lại cả vòng tới khi hết 90 giây thay vì bỏ
    cuộc sau lần đầu — lần bấm sau khi hydrate xong mới là lần
    bấm thật.
    """
    locator = page.get_by_role("button", name="Log in", exact=True)
    deadline = time.monotonic() + BUTTON_TIMEOUT_MS / 1000
    attempts = 0
    while time.monotonic() < deadline:
        if _page_closed(page):
            raise RuntimeError(
                "Cửa sổ Chromium đã đóng trước khi bấm được Log in."
            )
        attempts += 1
        if _meta_testid_login(page):
            break
        candidates = _in_viewport_candidates(
            page, locator
        ) or _visible_candidates(locator)
        if candidates and _meta_try_login_buttons(page, candidates):
            break
        page.wait_for_timeout(3000)
    else:
        raise RuntimeError(
            f"Đã thử Log in {attempts} lần trong "
            f"{BUTTON_TIMEOUT_MS // 1000} giây nhưng trang không "
            "chuyển sang auth.meta.com — có thể trang meta.ai "
            "chưa tải xong hoặc đã đổi bố cục."
        )
    if attempts > 1:
        _status(f"Meta: đã bấm Log in (thử lần {attempts}).")


def _meta_wait_visible(page, locator, label):
    """Chờ một phần tử của luồng Meta hiện ra, chịu được trùng.

    Như .wait_for(state="visible") nhưng không lỗi strict mode
    khi locator khớp nhiều phần tử — trả về phần tử đầu tiên
    đang hiện (tiêu đề, ô nhập...). Dùng cho các điểm chờ của
    luồng Meta.
    """
    deadline = time.monotonic() + BUTTON_TIMEOUT_MS / 1000
    while True:
        if _page_closed(page):
            raise RuntimeError(
                "Cửa sổ Chromium đã đóng trong lúc chờ "
                f"{label} hiện ra."
            )
        try:
            count = locator.count()
        except Exception:
            count = 0
        for index in range(count):
            candidate = locator.nth(index)
            try:
                if candidate.is_visible():
                    return candidate
            except Exception:
                continue
        if time.monotonic() >= deadline:
            raise RuntimeError(
                f"Không thấy {label} sau {BUTTON_TIMEOUT_MS // 1000} giây."
            )
        page.wait_for_timeout(250)


def _meta_click_button(page, names):
    """Bấm nút đầu tiên trong danh sách đang hiện và bấm được.

    Trả về tên nút đã bấm, None khi không có nút nào. Dùng cho
    các bước sau CAPTCHA/selfie — nhãn nút tiếp theo chưa biết
    chắc nên thử lần lượt. Không chờ lâu từng tên: tên nào không
    có nút hiện thì bỏ qua ngay. Ưu tiên nút trong khung hình.
    """
    for name in names:
        button = page.get_by_role("button", name=name, exact=True)
        target = _first_in_viewport(page, button) or _first_visible(
            button
        )
        if target is None:
            continue
        try:
            page.wait_for_timeout(UI_SETTLE_MS)
            target.click(timeout=BUTTON_TIMEOUT_MS)
            return name
        except Exception:
            continue
    return None


def _meta_fill_code(page, code):
    """Điền mã xác minh vào form "Confirm your account" của Meta.

    Form Meta là một ô duy nhất (nhãn "Code"), không phải ô
    tách như Muse.ai, và KHÔNG tự gửi — phía gọi vẫn bấm
    Continue sau khi điền. Ô nhập của Meta dùng nhãn nổi
    (floating label) nên thử placeholder -> label -> textbox
    theo tên, giống ô email. Mã Meta luôn 6 số — mã lấy được
    mà không đủ 6 số thì lỗi ngay tại đây thay vì điền mã sai
    rồi kẹt 7 phút ở bước sau.
    """
    if len(code) != 6 or not code.isdigit():
        raise RuntimeError(
            f"Mã xác minh lấy được không đúng 6 số "
            f"({len(code)} ký tự) — không điền để khỏi kẹt ở "
            "form xác nhận."
        )
    for locator in (
        page.get_by_placeholder("Code", exact=True),
        page.get_by_label("Code"),
        page.get_by_role("textbox", name="Code"),
    ):
        field = _first_visible(locator)
        if field is not None:
            field.fill(code)
            return
    _fill_verification_code(page, code)


def _meta_app_error_visible(page):
    """Màn hình "Application Error" của Meta (React error
    boundary, live 2026-11-13) — trang đã chết, chỉ tải lại
    mới cứu được."""
    try:
        heading = page.get_by_text("Application Error", exact=True)
        return heading.count() > 0 and heading.first.is_visible()
    except Exception:
        return False


def _meta_reload_on_app_error(page, max_reloads=3):
    """Tải lại meta.ai nếu vừa mở đã gặp màn hình "Application
    Error" (React error boundary — live 2026-11-13, lỗi ngay lúc
    mới vào trang, không phải lỗi ở bước sau). Chờ 2 giây sau
    mỗi lần tải lại để React kịp vẽ lại (hoặc lỗi lại). Quá số
    lần cho phép mà vẫn lỗi thì lỗi rõ.
    """
    for attempt in range(max_reloads + 1):
        page.wait_for_timeout(2000)
        if not _meta_app_error_visible(page):
            return
        if attempt >= max_reloads:
            raise RuntimeError(
                "Trang meta.ai báo lỗi ứng dụng (Application Error) "
                f"sau {max_reloads} lần tải lại."
            )
        _status("Meta: vừa mở meta.ai đã lỗi ứng dụng — tải lại trang...")
        try:
            page.reload(wait_until="domcontentloaded")
        except Exception:
            pass


def _meta_checkpoint_visible(page):
    """Màn hình checkpoint "Confirm you're human to use your
    account" (auth.meta.com/checkpoints/..., live 2026-11-13)
    — bấm Continue thì mới hiện CAPTCHA thật."""
    try:
        heading = page.get_by_text("Confirm you're human")
        return heading.count() > 0 and heading.first.is_visible()
    except Exception:
        return False


def _meta_finish_after_otp(page, api_key, selfie_path):
    """Sau bước OTP: xử lý CAPTCHA + selfie rồi đợi vào tài khoản.

    Hai biến thể (quan sát tay 2026-11-13): (A) trang về màn
    hình đầu meta.ai — bấm Log in lại thì gặp CAPTCHA; (B)
    CAPTCHA hiện ngay sau OTP. Sau CAPTCHA là bước up ảnh
    selfie, rồi mới vào tài khoản. Vòng lặp thăm dò lần lượt:
    trang lỗi ứng dụng (tải lại) → đã đăng nhập chưa → có
    checkpoint "Confirm you're human" không (bấm Continue để
    hiện CAPTCHA) → có CAPTCHA không (giải qua CapSolver)
    → có ô up selfie không → có phải về trang chủ không (bấm
    Log in lại). Hết giờ thì lỗi rõ kèm URL hiện tại để chụp
    màn hình bổ sung.
    """
    deadline = time.monotonic() + META_POST_OTP_TIMEOUT_SECONDS
    selfie_done = False
    relogin_tried = False
    error_reloads = 0
    while time.monotonic() < deadline:
        if page.is_closed():
            raise RuntimeError(
                "Cửa sổ Chromium đã đóng trước khi xong đăng ký Meta."
            )
        if _meta_app_error_visible(page):
            if error_reloads >= 3:
                raise RuntimeError(
                    "Trang Meta báo lỗi ứng dụng (Application Error) "
                    "sau 3 lần tải lại."
                )
            error_reloads += 1
            _status("Meta: trang báo lỗi ứng dụng — tải lại trang...")
            try:
                page.reload(wait_until="domcontentloaded")
            except Exception:
                pass
            page.wait_for_timeout(2000)
            continue
        if _meta_logged_in(page):
            # Xác nhận lần hai sau 2 giây để loại lần đọc lúc
            # trang đang chuyển (nút Log in chưa kịp hiện lại).
            page.wait_for_timeout(2000)
            if _meta_logged_in(page):
                return
            continue
        if _meta_checkpoint_visible(page):
            clicked = _meta_click_button(page, ("Continue",))
            if clicked:
                _status(
                    "Meta: đã bấm Continue ở trang kiểm tra "
                    "'Confirm you're human'."
                )
                page.wait_for_timeout(2000)
                continue
            # Không bấm được Continue (nút đã biến mất) -> rơi
            # xuống kiểm tra CAPTCHA/selfie bên dưới.
        try:
            kind = capsolver.solve_page_captcha(page, api_key)
        except (capsolver.CapSolverError, TimeoutError) as error:
            raise RuntimeError(
                "Giải CAPTCHA qua CapSolver lỗi: "
                f"{_short_error(error)}"
            )
        if kind:
            _status(f"Meta: đã giải CAPTCHA ({kind}) qua CapSolver.")
            page.wait_for_timeout(1500)
            clicked = _meta_click_button(
                page, ("Continue", "Verify", "Submit", "Confirm")
            )
            if clicked:
                _status(f"Meta: đã bấm {clicked} sau CAPTCHA.")
            page.wait_for_timeout(2000)
            continue
        if not selfie_done:
            file_input = page.locator('input[type="file"]')
            try:
                visible = (
                    file_input.count() and file_input.first.is_visible()
                )
            except Exception:
                visible = False
            if visible:
                _status("Meta: đang up ảnh selfie...")
                file_input.first.set_input_files(str(selfie_path))
                selfie_done = True
                page.wait_for_timeout(2000)
                clicked = _meta_click_button(
                    page, ("Continue", "Upload", "Confirm", "Submit")
                )
                if clicked:
                    _status(f"Meta: đã bấm {clicked} sau khi up selfie.")
                page.wait_for_timeout(2000)
                continue
        if not relogin_tried:
            login = page.get_by_role("button", name="Log in", exact=True)
            candidates = _in_viewport_candidates(
                page, login
            ) or _visible_candidates(login)
            try:
                testid_present = (
                    page.get_by_test_id("login-button").count() > 0
                )
            except Exception:
                testid_present = False
            if testid_present or candidates:
                _status("Meta: trang về màn hình đầu — bấm Log in lại...")
                if _meta_testid_login(page) or _meta_try_login_buttons(
                    page, candidates
                ):
                    relogin_tried = True
                page.wait_for_timeout(2000)
                continue
        page.wait_for_timeout(3000)
    raise RuntimeError(
        "Hết "
        f"{META_POST_OTP_TIMEOUT_SECONDS // 60} phút sau bước OTP mà "
        "chưa vào được tài khoản Meta "
        f"(URL hiện tại: {page.url[:120]})."
    )


def _meta_find_email_field(page):
    """Tìm ô nhập email ở auth.meta.com (nhiều cách định vị).

    Ô email của Meta dùng nhãn nổi (floating label): chữ
    "Mobile number or email" là <label> nằm phía trên ô nhập,
    không phải thuộc tính placeholder của <input> — chỉ tìm
    bằng placeholder thì không thấy gì dù form đã hiện (live
    2026-11-13 10:20: qua được form mà tool không điền mail).
    Thử lần lượt placeholder -> label -> textbox theo tên; trả
    về locator đơn đầu tiên hiện ra, None khi không có ô nào.
    """
    for locator in (
        page.get_by_placeholder("Mobile number or email"),
        page.get_by_label("Mobile number or email"),
        page.get_by_role("textbox", name="Mobile number or email"),
    ):
        found = _first_visible(locator)
        if found is not None:
            return found
    return None


def _meta_wait_email_field(page):
    """Chờ ô nhập email ở auth.meta.com hiện ra rồi trả về nó."""
    deadline = time.monotonic() + BUTTON_TIMEOUT_MS / 1000
    while True:
        if _page_closed(page):
            raise RuntimeError(
                "Cửa sổ Chromium đã đóng trong lúc chờ ô email."
            )
        found = _meta_find_email_field(page)
        if found is not None:
            return found
        if time.monotonic() >= deadline:
            raise RuntimeError(
                "Không thấy ô Mobile number or email sau "
                f"{BUTTON_TIMEOUT_MS // 1000} giây."
            )
        page.wait_for_timeout(250)


def _meta_pick_email_signup(page):
    """Vào form nhập email ở auth.meta.com (2 biến thể).

    Sau khi bấm Log in, auth.meta.com có 2 dạng: (A) màn hình
    chọn phương thức — phải bấm "Use mobile number or email"
    (quan sát tay 2026-11-13); (B) vào thẳng form "What's your
    mobile number or email?" (live 2026-11-13 10:20 — không có
    màn hình chọn). Chờ tới khi một trong hai hiện ra: thấy ô
    email thì khỏi bấm chọn, thấy nút chọn thì bấm. Lỗi rõ ràng
    nếu 30 giây không thấy bên nào.
    """
    chooser = page.get_by_text("Use mobile number or email", exact=True)
    deadline = time.monotonic() + 30
    while True:
        if _page_closed(page):
            raise RuntimeError(
                "Cửa sổ Chromium đã đóng trước khi vào form email."
            )
        if _meta_find_email_field(page) is not None:
            return
        if _first_visible(chooser) is not None:
            _meta_click(page, chooser, "Use mobile number or email")
            return
        if time.monotonic() >= deadline:
            raise RuntimeError(
                "Trang auth.meta.com không hiện form email hay nút "
                "chọn đăng ký bằng email sau 30 giây."
            )
        page.wait_for_timeout(250)


def _run_meta_registration(page, email, mailbox, existing_ids):
    """Đăng ký một tài khoản Meta ở meta.ai (chế độ --meta).

    Thứ tự theo quan sát tay của chủ tool (2026-11-13):
    meta.ai -> Log in -> "Use mobile number or email" (có khi
    auth.meta.com vào thẳng form email, không có màn hình chọn)
    -> nhập temp mail -> ngày sinh (trên 18) + mật khẩu ngẫu
    nhiên (không lưu) -> giữ nguyên username tự sinh, bấm
    Confirm -> nhập mã OTP từ hộp thư -> CAPTCHA (CapSolver) +
    up ảnh selfie -> vào tài khoản. Xong tới đây thì tài khoản
    Meta đã tạo xong; app chính lưu lại (email + bí mật hộp thư
    + username) và DỪNG — không qua Muse (đăng nhập Muse là
    bước riêng sau này).
    """
    api_key = _capsolver_key()
    if not api_key:
        raise RuntimeError(
            "Chưa lưu CapSolver API key — nhập key ở app chính "
            "rồi chạy lại."
        )
    selfie_path = _load_selfie_path()
    if not selfie_path:
        raise RuntimeError(
            "Chưa chọn ảnh selfie (hoặc file ảnh không còn) — "
            "chọn lại ở app chính rồi chạy lại."
        )
    _status("Meta: đang mở meta.ai...")
    page.goto(META_URL, wait_until="domcontentloaded", timeout=PAGE_TIMEOUT_MS)
    _meta_reload_on_app_error(page)
    _meta_click_login(page)
    _status("Meta: đang chọn đăng ký bằng email...")
    _meta_pick_email_signup(page)
    email_field = _meta_wait_email_field(page)
    email_field.fill(email)
    _meta_click(
        page,
        page.get_by_role("button", name="Continue", exact=True),
        "Continue",
    )
    _status("Meta: đang điền ngày sinh + mật khẩu...")
    _meta_wait_visible(
        page,
        page.get_by_text("Get started on Meta AI", exact=True),
        "Get started on Meta AI",
    )
    birth_date = _random_birth_date()
    _select_birth_date(page, birth_date)
    password_field = page.get_by_placeholder("Password")
    if not password_field.count():
        password_field = page.locator('input[type="password"]')
    password_field.first.fill(_random_meta_password())
    _meta_click(
        page,
        page.get_by_role("button", name="Continue", exact=True),
        "Continue",
    )
    _status("Meta: đang xác nhận tạo tài khoản...")
    _meta_wait_visible(
        page,
        page.get_by_text("Edit your Meta AI details", exact=True),
        "Edit your Meta AI details",
    )
    username = ""
    try:
        user_field = page.get_by_label("Username")
        if user_field.count():
            username = user_field.first.input_value().strip()
    except Exception:
        username = ""
    if not username:
        # Nhãn nổi không gắn với ô input: lấy ô text đầu tiên
        # đang có giá trị (ô Username tự điền, ô Name để trống).
        try:
            for index in range(page.locator("input").count()):
                candidate = page.locator("input").nth(index)
                if candidate.is_visible():
                    value = (candidate.input_value() or "").strip()
                    if value:
                        username = value
                        break
        except Exception:
            username = ""
    if username:
        _meta_write_username(username)
        _status(f"META_USERNAME: {username}")
    _meta_click(
        page,
        page.get_by_role("button", name="Confirm", exact=True),
        "Confirm",
    )
    _status(
        "Meta: đang chờ mã xác minh từ "
        f"{mailbox['provider']}..."
    )
    _meta_wait_visible(
        page,
        page.get_by_text("Confirm your account", exact=True),
        "Confirm your account",
    )
    code = _wait_for_code_from_mail(page, mailbox, existing_ids)
    _status("Meta: đã nhận mã xác minh, đang điền mã...")
    _meta_fill_code(page, code)
    _meta_click(
        page,
        page.get_by_role("button", name="Continue", exact=True),
        "Continue",
    )
    _meta_finish_after_otp(page, api_key, selfie_path)
    _status("Meta: đã đăng ký xong tài khoản Meta.")


def main():
    global LANE_SUFFIX, DIRECT_MODE
    if len(sys.argv) < 2:
        raise ValueError("Thiếu địa chỉ email cần điền.")
    email = sys.argv[1]
    flags = sys.argv[2:]
    auto = "--auto" in flags
    # --meta: đăng ký tài khoản Meta ở meta.ai thay vì đăng ký
    # Muse.ai (app chính chạy với nút "Đăng kí Meta" riêng).
    meta = "--meta" in flags
    for arg in flags:
        if arg.isdigit():
            # Luồng song song: đọc đúng file proxy/redeem của luồng này.
            LANE_SUFFIX = "-" + arg
            break
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
            "Chưa cấu hình proxy — tool chỉ đăng ký qua proxy hoặc "
            "qua IP máy (chế độ IP máy). Hãy nhập proxy trong ứng "
            "dụng, hoặc chọn chế độ IP máy rồi chạy lại."
        )
        if auto:
            _status(f"RESULT: FAIL {detail}")
            sys.exit(1)
        _show_error("Chưa cấu hình proxy", detail)
        return
    DIRECT_MODE = proxy == "DIRECT"
    if DIRECT_MODE:
        _status(
            "Chế độ IP máy (VPN): chạy trực tiếp qua IP của máy, "
            "không qua proxy."
        )

    playwright = sync_playwright().start()
    browser = None
    page = None
    try:
        browser = playwright.chromium.launch(
            headless=False, proxy=None if DIRECT_MODE else proxy
        )
        page = browser.new_page()
        page.set_default_timeout(PAGE_TIMEOUT_MS)
        # Không kiểm tra IP thoát nữa: kho proxy do người dùng đảm
        # bảo toàn bộ là proxy US. Nếu một proxy nào đó thoát sai
        # khu vực, màn hình waitlist sau bước hồ sơ vẫn chặn lại.
        mailbox = _load_mailbox(email)
        try:
            existing_ids = _get_message_ids(mailbox)
        except Exception:
            # Hộp thư mới luôn trống; không lấy được danh sách thư cũ
            # thì dùng tập rỗng làm mốc, đừng chặn cả lượt đăng ký.
            existing_ids = set()
        if meta:
            # Chế độ đăng ký Meta: chạy xong luồng Meta thì dừng —
            # không qua Muse. Lỗi ở bước nào thì khối except bên
            # ngoài xử lý như luồng Muse (RESULT: FAIL + exit 1
            # khi chạy tự động). return trong try vẫn chạy finally
            # (đóng trình duyệt) như thường.
            _run_meta_registration(page, email, mailbox, existing_ids)
            if auto:
                _status("RESULT: META_OK")
            else:
                page.wait_for_event("close")
            return
        # API thay nút bấm, form vẫn điền như cũ: điền form email
        # trước rồi gửi OTP qua API; xác nhận mã qua API; rồi gửi
        # hoàn tất hồ sơ qua API LUÔN với ngày sinh ngẫu nhiên —
        # KHÔNG mở lại trang để điền form hồ sơ trước, vì tải lại
        # lúc chưa có phiên làm mất trạng thái luồng và trang về
        # khung đăng nhập (gặp live 2026-11-12). Riêng form mã thì
        # KHÔNG điền trước được — nó tự gửi ngay khi nhập đủ số,
        # nên bước mã chỉ gọi API (form để trống); API lỗi mới
        # điền form để nó tự gửi. Mọi lệnh gọi API đều kèm CSRF
        # token lấy qua POST /restart (thiếu token server trả 403
        # csrfFailed — gặp live 2026-11-12). Hoàn tất hồ sơ xong
        # thì đồng bộ phiên web bằng form bootstrap (như trang tự
        # làm) rồi mới tới bước tiếp theo. Trục trặc ở bước nào
        # thì rơi về bấm nút DOM ở bước đó. Từ bước nhập thẻ trở
        # đi mọi thứ giữ nguyên như cũ.
        continue_button = _open_signup_and_submit_email(page, email)
        # Một UUID "waterfall" cho cả lượt thử này — trang tự sinh
        # một UUID khi bắt đầu lượt đăng ký (beginHatchAuthAttempt)
        # và gửi kèm trên /restart lẫn mọi lệnh auth sau đó.
        waterfall_id = str(uuid.uuid4())
        token_box = {
            "token": _fetch_csrf_token(page, waterfall_id),
            "waterfall_id": waterfall_id,
        }
        if token_box["token"] is None:
            _status("Không lấy được CSRF token — dùng nút bấm.")
            use_api = False
        else:
            use_api = _try_send_otp_api(page, email, token_box)
        tos_version = None
        if use_api:
            _status(
                "Đã gửi OTP qua API, đang chờ mã xác minh từ "
                f"{mailbox['provider']}..."
            )
            try:
                code = _wait_for_code_from_mail(page, mailbox, existing_ids)
            except TimeoutError:
                _status(
                    "Chờ mã OTP qua API hết giờ — bấm Continue để "
                    "gửi lại mã rồi điền như cũ."
                )
                use_api = False
            else:
                _status("Đã nhận mã xác minh, đang xác nhận qua API...")
                confirmed = _try_confirm_otp_api(page, code, token_box)
                if confirmed is None:
                    use_api = False
                else:
                    tos_version, requires_name = confirmed
        if not use_api:
            _click_when_ready(page, continue_button, "Continue")
            # Mốc thư cũ chụp lại sau khi bấm: mã OTP của lần gửi
            # qua API (nếu có) coi như thư cũ, chỉ chờ mã mới gửi
            # lại qua nút bấm.
            try:
                existing_ids = _get_message_ids(mailbox)
            except Exception:
                pass
            _status(
                "Đã gửi email, đang chờ mã xác minh từ "
                f"{mailbox['provider']}..."
            )
            code = _wait_for_verification_code(page, mailbox, existing_ids)
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
        else:
            # API đã xác nhận OTP: gửi hoàn tất hồ sơ qua API
            # LUÔN, không mở lại trang để điền form hồ sơ trước.
            # Lý do (gặp live 2026-11-12): ở bước này tài khoản
            # chưa có phiên — phiên chỉ được tạo khi
            # register-completion thành công (trả về redirect_to
            # + auth_meta_session). Tải lại trang lúc này làm
            # mất trạng thái luồng trong bộ nhớ của app, trang
            # khởi động lại như chưa đăng nhập và chỉ còn khung
            # đăng nhập; form hồ sơ cũng không hiện để mà điền.
            # Ngày sinh dùng ngẫu nhiên (trên 18 tuổi) như luồng
            # API ban đầu; họ tên kèm theo khi kết quả xác nhận
            # OTP báo requires_name (đúng như trang tự làm).
            completion = _try_register_completion_api(
                page, tos_version, token_box,
                include_names=requires_name,
            )
            if completion == REGISTRATION_BLOCKED:
                # Meta chặn tạo tài khoản trên IP này — cùng bản
                # chất với hộp thoại "Something went wrong" ở luồng
                # bấm nút: báo PROXY_BURN để lượt chạy bỏ proxy này
                # thay vì thử tiếp 50 tài khoản trên cùng một IP
                # bị chặn (gặp live 2026-11-12: 4/5 luồng cùng lúc).
                _status(
                    "PROXY_BURN: Meta chặn tạo tài khoản trên IP này "
                    "(registrationBlocked ở API hoàn tất hồ sơ)."
                )
                if DIRECT_MODE:
                    detail = (
                        "Meta báo không thể tạo tài khoản lúc này — IP "
                        "của máy (VPN) đã bị chặn; đổi server VPN "
                        "khác rồi chạy lại."
                    )
                else:
                    detail = (
                        "Meta báo không thể tạo tài khoản lúc này — IP "
                        "của proxy đang dùng đã bị chặn; đổi proxy khác."
                    )
                raise RuntimeError(detail)
            if completion is None:
                # API lỗi: KHÔNG tải lại trang — tải lại cũng về
                # khung đăng nhập và mất luôn màn hình mã/hồ sơ
                # đang mở, trong khi phân loại bên dưới vẫn chạy
                # được trên trang hiện tại.
                _status(
                    "API hoàn tất hồ sơ lỗi — giữ nguyên trang hiện "
                    "tại để kiểm tra màn hình thật."
                )
            else:
                # Hoàn tất xong: đồng bộ phiên web bằng form
                # bootstrap (đúng cách trang tự làm) để trình duyệt
                # nhận cookie phiên rồi tới bước tiếp theo (thường
                # là checkpoint nhập thẻ). Không có token phiên,
                # hoặc gửi form lỗi, thì mở thẳng redirect_to
                # server trả về; không có nữa mới mở lại trang chính.
                frl_token = None
                session = completion.get("auth_meta_session")
                if isinstance(session, dict):
                    frl_token = session.get("frl_access_token")
                redirect_to = completion.get("redirect_to")
                if isinstance(frl_token, str) and frl_token:
                    _status(
                        "Đã hoàn tất hồ sơ qua API — đang đồng bộ "
                        "phiên web..."
                    )
                    bootstrapped = _bootstrap_session(
                        page, frl_token, redirect_to
                    )
                else:
                    _status(
                        "Hoàn tất hồ sơ xong nhưng không có token "
                        "phiên — mở thẳng trang đích thay thế."
                    )
                    bootstrapped = False
                if not bootstrapped:
                    path = _resolve_redirect_path(redirect_to)
                    _status(
                        "Đã hoàn tất hồ sơ qua API — tới bước tiếp "
                        f"theo ({path})..."
                    )
                    page.goto(
                        SIGNUP_URL.rstrip("/") + path,
                        wait_until="domcontentloaded",
                        timeout=PAGE_TIMEOUT_MS,
                    )
        if auto:
            _status("Đã gửi hồ sơ, đang chờ màn hình tiếp theo...")
            outcome = _classify_after_profile(page)
            if outcome == "error":
                _status(
                    "PROXY_BURN: Meta chặn tạo tài khoản trên IP này "
                    "('Something went wrong')."
                )
                if DIRECT_MODE:
                    detail = (
                        "Meta báo không thể tạo tài khoản lúc này — IP "
                        "của máy (VPN) đã bị chặn; đổi server VPN "
                        "khác rồi chạy lại."
                    )
                else:
                    detail = (
                        "Meta báo không thể tạo tài khoản lúc này — IP "
                        "của proxy đang dùng đã bị chặn; đổi proxy khác."
                    )
                raise RuntimeError(detail)
            if outcome == "waitlist":
                _status(
                    "WAITLIST: IP thoát rơi vào khu vực chưa được "
                    "hỗ trợ ('Join the waitlist')."
                )
                if DIRECT_MODE:
                    detail = (
                        "Ra màn hình 'Join the waitlist' — Muse chưa "
                        "khả dụng ở IP hiện tại. Hãy đổi sang server "
                        "VPN US khác rồi chạy lại."
                    )
                else:
                    detail = (
                        "Ra màn hình 'Join the waitlist' — Muse chưa khả "
                        "dụng ở IP thoát hiện tại. Hãy đổi sang proxy US "
                        "khác rồi chạy lại."
                    )
                raise RuntimeError(detail)
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
        site = "Meta" if meta else "Muse.ai"
        detail = (
            f"Trang {site} chưa phản hồi đúng thời gian, hoặc form đã "
            f"thay đổi. {_short_error(error)}"
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
            _status(f"RESULT: FAIL {_short_error(error)}")
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


def _fill_profile(page, timeout_ms=BUTTON_TIMEOUT_MS):
    """Điền form hồ sơ, trả về ngày sinh đã chọn.

    timeout_ms: thời gian chờ form hiện — luồng DOM chờ lâu
    (mặc định), luồng API chờ ngắn vì form có thể không hiện
    (lúc đó API gửi luôn mà không cần form).
    """
    if not _wait_for_profile_form(page, timeout_ms):
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
    return birth_date


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
        # Tên combobox: Muse.ai đặt đúng "Month"/"Day"/"Year",
        # còn Meta đặt "Select month"/"Select day"/"Select year"
        # (quan sát DOM 2026-11-13) — chấp nhận cả hai.
        control = page.get_by_role(
            "combobox", name=re.compile(f"^(Select )?{label}$", re.I)
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

            option.first.wait_for(state="visible", timeout=1000)
            option.first.click(timeout=BUTTON_TIMEOUT_MS)
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
        if _birthday_value_matches(control, option_value, value, label):
            return
        if time.monotonic() >= deadline:
            raise RuntimeError(
                f"Không xác nhận được {value} cho {label.lower()} sinh."
            )
        page.wait_for_timeout(250)


def _birthday_value_matches(control, option_value, value, label=None):
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
    if selected_value is not None and selected_value != option_value:
        return False
    if displayed_value == value:
        return True
    # Combobox tự vẽ (Meta): inner_text gồm cả nhãn ("Month")
    # lẫn giá trị ("November") — "Month\nNovember" hoặc dính
    # liền "MonthNovember". Bỏ phần nhãn ở đầu rồi so, hoặc
    # giá trị là một dòng riêng trong đó. Không dùng endswith:
    # ngày "2" mà hiển thị "12" thì endswith vẫn khớp sai.
    if label and displayed_value.startswith(label):
        if displayed_value[len(label):].strip() == value:
            return True
    return value in [
        line.strip() for line in displayed_value.splitlines()
    ]


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


def _load_mailbox(email):
    """Hộp thư của tài khoản này (app chính ghi mailbox.txt).

    App chính ghi {"provider", "email", "secret"} vào
    mailbox{LANE_SUFFIX}.txt trước khi mở helper — helper đọc
    thư đúng ở dịch vụ đó. Chạy helper độc lập (không có file,
    file hỏng, hoặc địa chỉ trong file khác email được giao)
    thì quay về tempmail.plus theo email như hành vi cũ.
    """
    try:
        raw = paths.data_file(
            f"mailbox{LANE_SUFFIX}.txt"
        ).read_text(encoding="utf-8")
        data = json.loads(raw)
    except (OSError, ValueError):
        data = None
    if (
        isinstance(data, dict)
        and data.get("email") == email
        and data.get("provider")
    ):
        return {
            "provider": data["provider"],
            "email": email,
            "secret": data.get("secret") or "",
        }
    return {
        "provider": mail_providers.DEFAULT_PROVIDER,
        "email": email,
        "secret": "",
    }


def _message_id(message):
    return message.get("id")


def _get_message_ids(mailbox):
    return {
        message_id
        for message_id in (
            _message_id(message)
            for message in mail_providers.list_messages(mailbox)
        )
        if message_id
    }


def _extract_strict_code(text):
    """Mã xác minh 6 số trong một đoạn văn bản.

    Ưu tiên 6 số liền nhau, rồi 3+3 cách nhau bằng khoảng
    trắng/gạch ngang ("482 891"). Trả về None khi không có.
    """
    found = re.findall(r"(?<!\d)\d{6}(?!\d)", text)
    if found:
        return found[0]
    spaced = re.search(r"(?<!\d)(\d{3})[ \u00a0\-](\d{3})(?!\d)", text)
    if spaced:
        return spaced.group(1) + spaced.group(2)
    return None


def _extract_loose_code(text):
    """Số 4-8 chữ số đầu tiên trong văn bản (fallback khi không
    có mã 6 số — tương thích mã 4/8 số nếu dịch vụ nào dùng)."""
    found = re.findall(r"(?<!\d)\d{4,8}(?!\d)", text)
    return found[0] if found else None


def _verification_code_from_new_messages(mailbox, existing_ids):
    """Tìm mã xác minh trong các thư MỚI của hộp thư.

    Ưu tiên mã 6 số trên TOÀN BỘ thư mới trước (tiêu đề tất cả
    thư, rồi toàn văn tất cả thư), vì thư của Meta chứa số 4
    chữ số khác đứng trước mã thật — lấy số đầu tiên 4-8 chữ
    số sẽ vớ nhầm (live 2026-11-13: lấy "4828" thay vì mã 6
    số, form báo "should have 6 digits"). Chỉ khi không có mã
    6 số nào mới lấy số 4-8 chữ số đầu tiên.
    """
    new_messages = []
    for message in mail_providers.list_messages(mailbox):
        message_id = _message_id(message)
        if not message_id or message_id in existing_ids:
            continue
        subject = message.get("subject") or ""
        try:
            body = mail_providers.read_message(mailbox, message_id)
        except Exception:
            body = ""
        new_messages.append((subject, body or ""))
    for subject, _body in new_messages:
        code = _extract_strict_code(subject)
        if code:
            return code
    for _subject, body in new_messages:
        code = _extract_strict_code(body)
        if code:
            return code
    for subject, body in new_messages:
        code = _extract_loose_code(subject)
        if code:
            return code
        code = _extract_loose_code(body)
        if code:
            return code
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


def _wait_for_verification_code(page, mailbox, existing_ids):
    deadline = time.monotonic() + CODE_TIMEOUT_SECONDS
    code = None
    warned = False
    while time.monotonic() < deadline:
        if page.is_closed():
            raise RuntimeError("Cửa sổ Chromium đã đóng trước khi xác minh.")
        if code is None:
            try:
                code = _verification_code_from_new_messages(
                    mailbox, existing_ids
                )
            except Exception:
                # Một lần rớt kết nối tới dịch vụ mail không
                # được giết cả tài khoản — bỏ qua lượt kiểm tra
                # này và thử lại ở vòng chờ tiếp theo.
                if not warned:
                    _status(
                        "Kiểm tra hộp thư bị rớt kết nối — đang thử lại..."
                    )
                    warned = True
        fields = _verification_input(page)
        if code and fields:
            return code
        page.wait_for_timeout(4000)
    raise TimeoutError(
        f"Không tìm thấy mã xác minh mới trong hộp thư "
        f"{mailbox['provider']} sau {CODE_TIMEOUT_SECONDS} giây."
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
