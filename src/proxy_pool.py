"""Kho proxy: định danh, xoay vòng, kho lỗi, tra IP máy.

Không có tkinter ở đây.
"""

import json
import urllib.request

import paths


# Kho lỗi mặc định (tab Muse) — tab Meta dùng file riêng, truyền
# qua tham số path của _load_bad_proxies / _save_bad_proxies.
DEFAULT_BAD_PROXIES_FILE = paths.data_file("bad-proxies.txt")


def _proxy_display(value):
    """Chỉ lấy host:port của một proxy để ghi log (che thông tin đăng nhập)."""
    display = value
    if "://" in display:
        display = display.split("://", 1)[1]
    if "@" in display:
        display = display.rsplit("@", 1)[1]
    parts = display.split(":")
    return ":".join(parts[:2]) if len(parts) >= 2 else display

def _line_burns_proxy(line):
    """Dòng tín hiệu từ browser_helper: proxy đang dùng bị Meta chặn."""
    return line.startswith("PROXY_BURN")

def _line_flags_waitlist(line):
    """Dòng tín hiệu từ browser_helper: IP thoát bị Join the waitlist."""
    return line.startswith("WAITLIST:")

def _line_flags_facebook(line):
    """Dòng tín hiệu từ browser_helper: proxy này đẩy trình duyệt
    sang facebook.com. Bỏ proxy ngay, không dùng lại."""
    return line.startswith("PROXY_FACEBOOK")

def _line_flags_load_fail(line):
    """Dòng tín hiệu từ browser_helper: một lượt chạy helper mở
    meta.ai lỗi tải trang hết 5 lần thử trên cùng một proxy
    (1 strike cho proxy đó; đủ 2 strike mới bỏ proxy)."""
    return line.startswith("PROXY_LOAD_FAIL")

def _proxy_key(value):
    """Khoá định danh một proxy = nguyên dòng proxy (gồm cả
    user/password), chuẩn hoá khoảng trắng và chữ hoa/thường.

    Các dòng proxy có thể dùng chung một cổng host:port (kiểu
    cổng xoay vòng) mà mỗi user/password là một phiên/IP thoát
    riêng — định danh theo host:port sẽ gộp chúng thành một và
    phá kho. Chỉ gộp khi hai dòng giống hệt nhau.
    """
    return value.strip().lower()

def _check_public_ip():
    """Tra IP công khai hiện tại của máy (IP thoát ra internet).

    Dùng ipinfo.io (HTTPS) trước, ip-api.com (HTTP) dự phòng;
    cả hai đều miễn phí, không cần key. Trả về dict {"ip",
    "country_code", "country", "city", "region", "org"} hoặc
    None khi cả hai đều lỗi. Chỉ đọc — không gửi gì ngoài chính
    request tra IP. Dùng để hiển thị và cảnh báo, KHÔNG chặn
    chạy: lượt chạy vẫn tiếp tục dù IP không phải US hay tra
    không được.
    """
    request = urllib.request.Request(
        "https://ipinfo.io/json",
        headers={"User-Agent": "Mozilla/5.0"},
    )
    try:
        with urllib.request.urlopen(request, timeout=10) as response:
            data = json.loads(response.read().decode("utf-8"))
        ip = data.get("ip")
        if ip:
            return {
                "ip": ip,
                "country_code": (data.get("country") or "").upper(),
                "country": data.get("country") or "?",
                "city": data.get("city") or "",
                "region": data.get("region") or "",
                "org": data.get("org") or "",
            }
    except Exception:
        pass
    try:
        request = urllib.request.Request(
            "http://ip-api.com/json/",
            headers={"User-Agent": "Mozilla/5.0"},
        )
        with urllib.request.urlopen(request, timeout=10) as response:
            data = json.loads(response.read().decode("utf-8"))
        ip = data.get("query")
        if ip:
            return {
                "ip": ip,
                "country_code": (data.get("countryCode") or "").upper(),
                "country": data.get("country") or "?",
                "city": data.get("city") or "",
                "region": data.get("regionName") or "",
                "org": data.get("isp") or "",
            }
    except Exception:
        pass
    return None

def _format_ip_status(info):
    """Một dòng ngắn gọn cho nhãn IP: '1.2.3.4 (US — Los Angeles,
    California)'."""
    where = info["country_code"] or info["country"]
    if info["city"]:
        where += f" — {info['city']}"
        if info["region"]:
            where += f", {info['region']}"
    return f"{info['ip']} ({where})"

def _prepare_pool(pool, bad_keys):
    """Lọc kho proxy cho một đợt chạy: bỏ các dòng nằm trong kho
    lỗi (khớp theo nguyên dòng qua _proxy_key) và bỏ dòng trùng
    hệt nhau (giữ dòng đầu tiên)."""
    seen = set()
    result = []
    for line in pool:
        key = _proxy_key(line)
        if key in bad_keys or key in seen:
            continue
        seen.add(key)
        result.append(line)
    return result

class _ProxyRotation:
    """Xoay proxy theo hạn mức tài khoản; proxy bị chặn thì bỏ ngay.

    current() cho proxy của tài khoản sắp chạy. Sau mỗi tài khoản
    gọi advance_after_account(): đủ hạn mức — hoặc tài khoản đó
    báo proxy bị Meta chặn tạo tài khoản — thì tài khoản sau dùng
    proxy kế tiếp; hết kho thì quay vòng lại từ đầu. Hạn mức tính
    theo lượt thử, bất kể thành công hay thất bại.

    Strike lỗi tải trang (2026-11-13): một lượt chạy helper mở
    meta.ai lỗi hết 5 lần thử = 1 strike cho proxy đó
    (record_load_fail_strike). Đủ 2 strike liên tiếp mới bỏ
    proxy; mở được meta.ai thì xoá strike
    (clear_load_fail_strikes). Bộ đếm chỉ sống trong một lượt
    chạy (tạo mới mỗi lần _run_lane/_run_login_lane chạy), khoá
    theo nguyên dòng proxy qua _proxy_key.
    """

    def __init__(self, pool, per_proxy):
        self.pool = pool
        self.per_proxy = per_proxy
        self.cursor = 0
        self.used = 0
        self.cycles = 0
        self.retired = set()
        self.load_fail_strikes = {}

    def record_load_fail_strike(self, line):
        """Ghi 1 strike cho proxy (một lượt chạy helper mở meta.ai
        lỗi tải trang hết 5 lần thử). Trả về tổng số strike của
        proxy này trong lượt chạy."""
        key = _proxy_key(line)
        strikes = self.load_fail_strikes.get(key, 0) + 1
        self.load_fail_strikes[key] = strikes
        return strikes

    def clear_load_fail_strikes(self, line):
        """Xoá strike của proxy — gọi khi một lượt chạy helper mở
        được meta.ai (proxy đã chứng minh còn tải được trang):
        strike chỉ tính các lần hỏng LIÊN TIẾP."""
        self.load_fail_strikes.pop(_proxy_key(line), None)

    def current(self):
        return self.cursor, self.pool[self.cursor]

    def _step(self):
        if self.cursor + 1 >= len(self.pool):
            # Vừa dùng xong proxy cuối: đã đi hết một vòng kho.
            self.cycles += 1
        self.cursor = (self.cursor + 1) % len(self.pool)

    def _skip_retired(self):
        guard = 0
        while self.cursor in self.retired and guard < len(self.pool):
            self._step()
            guard += 1

    def advance_after_account(self, burned=False):
        self.used += 1
        if burned or self.used >= self.per_proxy:
            self._step()
            self.used = 0
            self._skip_retired()

    def retire_current(self):
        """Loại hẳn proxy hiện tại: bỏ phần hạn mức còn lại và
        không bao giờ quay lại proxy này, kể cả khi xoay vòng."""
        self.retired.add(self.cursor)
        self._step()
        self.used = 0
        self._skip_retired()

    def all_retired(self):
        return len(self.retired) >= len(self.pool)

    def pool_exhausted(self):
        """Đã đi hết một vòng kho proxy (mỗi proxy hết lượt một lần),
        hoặc mọi proxy đều đã bị loại vào kho lỗi."""
        return self.cycles >= 1 or self.all_retired()

def _load_bad_proxies(path=None):
    """Đọc kho proxy lỗi từ file → dict khoá → dòng gốc."""
    result = {}
    try:
        text = (path or DEFAULT_BAD_PROXIES_FILE).read_text(encoding="utf-8")
    except OSError:
        return result
    for line in text.splitlines():
        line = line.strip()
        if line:
            result[_proxy_key(line)] = line
    return result


def _save_bad_proxies(mapping, path=None):
    target = path or DEFAULT_BAD_PROXIES_FILE
    try:
        if mapping:
            target.write_text(
                "\n".join(mapping.values()) + "\n", encoding="utf-8"
            )
        elif target.exists():
            target.unlink()
    except OSError:
        pass
