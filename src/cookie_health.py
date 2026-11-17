"""Kiểm tra cookie Muse còn hoạt động hay đã die.

Không chỉ "có cookie là được": một cookie có thể còn trong
hạn 7 ngày nhưng đã chết (đăng xuất ở nơi khác, phiên bị thu
hồi...). Hàm check_cookie gọi thẳng trang chủ Muse kèm cookie
đó và đọc dấu hiệu trên trang trả về:

- Sống: trang trả về giao diện chat (ô chat "Ask Meta AI").
- Die: trang trả về màn đăng nhập (URL chứa login/auth, hoặc
  trang có nút Log in mà không có giao diện chat).
- Không rõ: trang không nhận diện được, hoặc mạng lỗi — không
  đoán mò sống/chết khi không chắc.

Cookie KHÔNG bao giờ được ghi vào log hay trả về cùng kết quả
kiểm tra (chỉ trả trạng thái + ghi chú ngắn). Kết quả được cache
theo email trong cookie_health.json (cạnh accounts.xlsx) để
dashboard không phải gọi lại mỗi lần mở.
"""

import json
import urllib.error
import urllib.request
from datetime import datetime

import paths


MUSE_HOME_URL = "https://www.meta.ai/"
CHECK_TIMEOUT_SECONDS = 10
HEALTH_CACHE_FILE = "cookie_health.json"

# Dấu hiệu giao diện chat đã đăng nhập (ô chat của Muse — cùng
# dấu hiệu helper dùng để nhận diện giao diện, xem
# browser_helper._is_muse_chat_ui).
ALIVE_MARKERS = ("Ask Meta AI",)

STATUS_ALIVE = "song"
STATUS_DEAD = "chet"
STATUS_UNKNOWN = "khong-ro"

STATUS_LABELS = {
    STATUS_ALIVE: "Hoạt động",
    STATUS_DEAD: "Die",
    STATUS_UNKNOWN: "Chưa kiểm tra",
}

_BROWSER_UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/120.0.0.0 Safari/537.36"
)


def _default_fetch(url, cookie, timeout):
    """GET trang Muse kèm cookie -> (url cuối, body text).

    urllib tự theo redirect; url cuối cho biết có bị đẩy về
    màn đăng nhập không.
    """
    request = urllib.request.Request(
        url,
        headers={"Cookie": cookie, "User-Agent": _BROWSER_UA},
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:
        body = response.read().decode("utf-8", errors="replace")
        return response.geturl(), body


def check_cookie(cookie, fetcher=None, url=MUSE_HOME_URL):
    """Kiểm tra một cookie -> (status, detail).

    fetcher(url, cookie, timeout) -> (final_url, body) — dùng
    mặc định (urllib) khi không đưa vào; test đưa fetcher giả
    để không cần mạng thật. Không có hatch_sess thì cookie
    chưa bao giờ là phiên đăng nhập dùng được (cùng tiêu chuẩn
    _relogin_done_emails dùng để tính mail đã lấy cookie).
    """
    text = (cookie or "").strip()
    if not text:
        return STATUS_DEAD, "không có cookie"
    if "hatch_sess=" not in text:
        return STATUS_DEAD, "không có hatch_sess"
    fetch = fetcher or _default_fetch
    try:
        final_url, body = fetch(url, text, CHECK_TIMEOUT_SECONDS)
    except (urllib.error.URLError, OSError, TimeoutError) as error:
        return STATUS_UNKNOWN, f"không gọi được Muse ({error})"
    except Exception as error:
        return STATUS_UNKNOWN, f"không kiểm tra được ({error})"
    lowered_url = (final_url or "").lower()
    lowered_body = (body or "").lower()
    if any(marker.lower() in lowered_body for marker in ALIVE_MARKERS):
        return STATUS_ALIVE, "vào được giao diện chat"
    if "login" in lowered_url or "auth.meta.com" in lowered_url:
        return STATUS_DEAD, "bị đẩy về màn đăng nhập"
    if "log in" in lowered_body:
        return STATUS_DEAD, "trang hiện màn đăng nhập"
    return STATUS_UNKNOWN, "trang trả về không nhận diện được"


def _health_cache_path():
    return paths.APP_DATA / HEALTH_CACHE_FILE


def load_health_cache():
    """Cache kiểm tra: {email_lower: {status, checked_at, detail}}."""
    try:
        data = json.loads(
            _health_cache_path().read_text(encoding="utf-8")
        )
    except (OSError, ValueError):
        return {}
    return data if isinstance(data, dict) else {}


def save_health_cache(cache):
    try:
        path = _health_cache_path()
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(
            json.dumps(cache, ensure_ascii=False), encoding="utf-8"
        )
    except OSError:
        pass


def check_and_cache(email, cookie, fetcher=None):
    """Kiểm tra cookie của một email rồi lưu kết quả vào cache.

    Trả về dict {status, checked_at, detail}.
    """
    status, detail = check_cookie(cookie, fetcher=fetcher)
    entry = {
        "status": status,
        "checked_at": datetime.now().astimezone().isoformat(
            timespec="seconds"
        ),
        "detail": detail,
    }
    cache = load_health_cache()
    cache[(email or "").strip().lower()] = entry
    save_health_cache(cache)
    return entry
