"""Kiểm tra cookie Muse còn hoạt động hay đã die.

Không chỉ "có cookie là được": một cookie có thể còn trong
hạn 7 ngày nhưng đã chết (đăng xuất ở nơi khác, phiên bị thu
hồi...). Kiểm tra thật qua BrowserChecker: trình duyệt Chromium
headless (không mở cửa sổ), gắn cookie vào rồi mở trang chủ
Muse và đọc dấu hiệu trên trang:

- Sống: trang trả về giao diện chat (ô chat "Ask Meta AI").
- Die: trang trả về màn đăng nhập (URL chứa login/auth, hoặc
  trang có nút Log in mà không có giao diện chat).
- Không rõ: trang không nhận diện được, hoặc mạng lỗi — không
  đoán mò sống/chết khi không chắc.

Kiểm tra bằng urllib (check_cookie) KHÔNG dùng được nữa:
meta.ai trả 403 "Client challenge" cho client không phải trình
duyệt — hàm đó giữ lại cho test/tham chiếu. Kết quả được cache
theo email trong cookie_health.json (cạnh accounts.xlsx) để
dashboard không phải gọi lại mỗi lần mở.
"""

import json
import threading
import time
import urllib.error
import urllib.request
from datetime import datetime

import paths


# Khoá ghi cache kiểm tra: job kiểm tra chạy nhiều luồng, mỗi
# luồng đọc-sửa-ghi cookie_health.json — không khoá thì hai
# luồng đọc cùng một bản rồi ghi đè lên nhau, mất kết quả.
_CACHE_LOCK = threading.Lock()


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


def check_and_cache(email, cookie, fetcher=None, checker=None):
    """Kiểm tra cookie của một email rồi lưu kết quả vào cache.

    Trả về dict {status, checked_at, detail}. `checker` là một
    BrowserChecker đã mở (job kiểm tra dùng chung một browser
    cho cả lượt); không đưa thì dùng check_cookie qua urllib.
    """
    if checker is not None:
        status, detail = checker.check(cookie)
    else:
        status, detail = check_cookie(cookie, fetcher=fetcher)
    entry = {
        "status": status,
        "checked_at": datetime.now().astimezone().isoformat(
            timespec="seconds"
        ),
        "detail": detail,
    }
    with _CACHE_LOCK:
        cache = load_health_cache()
        cache[(email or "").strip().lower()] = entry
        save_health_cache(cache)
    return entry


BROWSER_CHECK_TIMEOUT_MS = 20000


def _is_valid_cookie_name(name):
    """Tên cookie hợp lệ theo RFC 7230 (token)."""
    if not name:
        return False
    return all(
        33 <= ord(ch) <= 126 and ch not in "()<>@,;:\\\"/[]?={} \t"
        for ch in name
    )


def _is_valid_cookie_value(value):
    """Giá trị cookie hợp lệ: ASCII in được, trừ dấu cách,
    dấu phẩy, chấm phẩy, nháy kép, gạch chéo ngược (Chromium
    từ chối cả lệnh gắn cookie khi một giá trị chứa các ký tự
    này — lỗi 'Invalid cookie fields' 2026-11-17)."""
    return all(
        33 <= ord(ch) <= 126 and ch not in '",;\\' for ch in value
    )


def _is_challenge_page(body):
    """Trang chặn bot (Cloudflare và tương tự): chưa phải giao
    diện thật, chờ thêm cũng chỉ thấy màn này."""
    return any(
        marker in (body or "").lower()
        for marker in (
            "checking your browser",
            "verify you are human",
            "attention required",
            "just a moment",
            "cloudflare",
        )
    )


def clean_cookie_pairs(pairs):
    """Lọc các cặp cookie trước khi gắn vào trình duyệt: bỏ tên
    rỗng, giá trị chứa ký tự cấm, trùng tên thì giữ cái sau.
    Cookie thật của Meta có thể kèm vài cookie phụ có giá trị
    lạ — bỏ chúng đi chứ không để cả lệnh gắn cookie sập theo.
    """
    cleaned = {}
    for name, value in pairs:
        if not _is_valid_cookie_name(name):
            continue
        if not _is_valid_cookie_value(value):
            continue
        cleaned[name] = value
    return list(cleaned.items())


def parse_cookie_string(text):
    """Chuỗi cookie 'a=1; b=2' -> [(tên, giá trị)]."""
    pairs = []
    for part in (text or "").split(";"):
        part = part.strip()
        if not part or "=" not in part:
            continue
        name, value = part.split("=", 1)
        name, value = name.strip(), value.strip()
        if name:
            pairs.append((name, value))
    return pairs


class BrowserChecker:
    """Kiểm tra cookie bằng trình duyệt thật, headless (không mở
    cửa sổ): một browser dùng chung cho cả lượt kiểm tra, mỗi
    cookie một context sạch rồi đóng.

    meta.ai chặn kiểm tra bằng urllib (403 "Client challenge" —
    đòi trình duyệt chạy JavaScript), nên kiểm tra sống/chết
    phải qua trình duyệt thật. Chậm hơn nhiều (vài giây một
    cookie) nhưng mới có kết quả thật.
    """

    def __init__(self, launcher=None, timeout_ms=None):
        # launcher: hàm () -> browser (để test đưa browser giả).
        # timeout_ms: tối đa chờ một cookie (mặc định 20 giây).
        self._launcher = launcher
        self._timeout_ms = (
            BROWSER_CHECK_TIMEOUT_MS
            if timeout_ms is None
            else timeout_ms
        )
        self._playwright = None
        self._browser = None

    def start(self):
        """Mở trình duyệt headless.

        Lỗi -> RuntimeError với hướng dẫn cài đặt (không để job
        chạy cả nghìn lần vô ích).
        """
        if self._launcher is not None:
            try:
                self._browser = self._launcher()
            except ImportError as error:
                raise RuntimeError(
                    "Thiếu Playwright — kiểm tra sống/chết cần "
                    "trình duyệt thật. Cài bằng lệnh:\n"
                    "`py -m pip install -r requirements.txt`\n"
                    "`py -m playwright install chromium`\n\n"
                    f"{error}"
                ) from error
            return
        try:
            from playwright.sync_api import sync_playwright
        except ImportError as error:
            raise RuntimeError(
                "Thiếu Playwright — kiểm tra sống/chết cần trình "
                "duyệt thật. Cài bằng lệnh:\n"
                "`py -m pip install -r requirements.txt`\n"
                "`py -m playwright install chromium`\n\n"
                f"{error}"
            ) from error
        try:
            self._playwright = sync_playwright().start()
            self._browser = self._playwright.chromium.launch(
                headless=True
            )
        except Exception as error:
            self.close()
            raise RuntimeError(
                "Không mở được trình duyệt kiểm tra: "
                f"{error}"
            ) from error

    def smoke_test(self):
        """Mở thử một trang trắng: trình duyệt mở được nhưng
        không tạo được trang thì cũng không kiểm tra được gì —
        để job dừng ngay với lý do thật thay vì chết lặng."""
        context = self._browser.new_context()
        try:
            page = context.new_page()
            try:
                page.goto("about:blank", timeout=10000)
            finally:
                page.close()
        finally:
            context.close()

    def check(self, cookie):
        """Kiểm tra một cookie -> (status, detail).

        Gắn cookie bằng `url` (dạng Playwright khuyên dùng) sau
        khi lọc bỏ cặp lỗi/trùng — một cookie phụ có giá trị lạ
        không được làm sập cả lần kiểm tra (lỗi 2026-11-17:
        add_cookies ném 'Invalid cookie fields' và giết chết
        cả 5 luồng kiểm tra). Gắn cookie lỗi -> chưa rõ, không
        ném ra ngoài.
        """
        text = (cookie or "").strip()
        if not text or "hatch_sess=" not in text:
            return STATUS_DEAD, "không có hatch_sess"
        pairs = clean_cookie_pairs(parse_cookie_string(text))
        if not any(name == "hatch_sess" for name, _ in pairs):
            return STATUS_DEAD, "không có hatch_sess"
        context = self._browser.new_context()
        try:
            try:
                context.add_cookies(
                    [
                        {
                            "name": name,
                            "value": value,
                            "url": MUSE_HOME_URL,
                        }
                        for name, value in pairs
                    ]
                )
            except Exception as error:
                return (
                    STATUS_UNKNOWN,
                    f"không gắn được cookie ({error})",
                )
            page = context.new_page()
            try:
                try:
                    # domcontentloaded thay vì load: chỉ cần HTML
                    # về là vòng chờ ổn định bên dưới lo phần còn
                    # lại — không chờ hình/font phụ tải xong
                    # (mỗi cookie nhanh hơn).
                    page.goto(
                        MUSE_HOME_URL,
                        timeout=self._timeout_ms,
                        wait_until="domcontentloaded",
                    )
                except Exception as error:
                    return (
                        STATUS_UNKNOWN,
                        f"không mở được meta.ai ({error})",
                    )
                # Chờ trang ổn định rồi mới đọc: meta.ai là SPA,
                # load xong vẫn còn điều hướng/challenge/render —
                # đọc ngay thì gặp 'page is navigating' hoặc
                # trang rỗng/chưa nhận diện được (lỗi 2026-11-17).
                deadline = (
                    time.monotonic() + self._timeout_ms / 1000
                )
                while True:
                    url = (page.url or "").lower()
                    if "login" in url or "auth.meta.com" in url:
                        return STATUS_DEAD, "bị đẩy về màn đăng nhập"
                    try:
                        body = (page.content() or "").lower()
                    except Exception:
                        # Trang đang điều hướng: chưa đọc được,
                        # chờ tiếp thay vì kết luận.
                        body = ""
                    if any(
                        marker.lower() in body
                        for marker in ALIVE_MARKERS
                    ):
                        return STATUS_ALIVE, "vào được giao diện chat"
                    if "log in" in body or "sign in" in body:
                        return STATUS_DEAD, "trang hiện màn đăng nhập"
                    if time.monotonic() >= deadline:
                        if _is_challenge_page(body):
                            return (
                                STATUS_UNKNOWN,
                                "trang đang kiểm tra bot "
                                "(challenge)",
                            )
                        if not body.strip():
                            return (
                                STATUS_UNKNOWN,
                                "trang tải chưa xong nội dung",
                            )
                        return (
                            STATUS_UNKNOWN,
                            "trang trả về không nhận diện được",
                        )
                    page.wait_for_timeout(250)
            except Exception as error:
                return (
                    STATUS_UNKNOWN,
                    f"không kiểm tra được ({error})",
                )
            finally:
                page.close()
        finally:
            context.close()

    def close(self):
        if self._browser is not None:
            try:
                self._browser.close()
            except Exception:
                pass
            self._browser = None
        if self._playwright is not None:
            try:
                self._playwright.stop()
            except Exception:
                pass
            self._playwright = None
