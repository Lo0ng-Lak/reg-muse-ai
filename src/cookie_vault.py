"""Kho cookie: hạn sống, trạng thái, che cookie.

Logic thuần (không có HTTP, không chạy trình duyệt) để test
được: một cookie sống 7 ngày kể từ lúc lấy (quy tắc của sư
huynh, 2026-11-17). Mốc lấy là cột "Thời gian lấy cookie"
trong accounts.xlsx; dòng cũ chưa có cột này thì lấy "Thời
gian tạo" làm mốc dự phòng (dòng được append đúng lúc lấy
cookie nên hai mốc trùng nhau trong trường hợp đó).
"""

from datetime import datetime, timedelta

import storage
import cookie_health


COOKIE_TTL_DAYS = 7
# Còn dưới ngưỡng này thì cảnh báo "sắp hết" để đăng nhập lại
# trước khi cookie chết.
EXPIRING_SOON_SECONDS = 24 * 3600

STATUS_ALIVE = "song"
STATUS_EXPIRING = "sap-het"
STATUS_EXPIRED = "het-han"
STATUS_UNKNOWN = "khong-ro"

STATUS_LABELS = {
    STATUS_ALIVE: "Còn sống",
    STATUS_EXPIRING: "Sắp hết",
    STATUS_EXPIRED: "Hết hạn",
    STATUS_UNKNOWN: "Không rõ",
}


def parse_dt(value):
    """Chuỗi ISO / datetime -> datetime có múi giờ, hoặc None."""
    if value is None:
        return None
    if isinstance(value, datetime):
        if value.tzinfo is None:
            return None
        return value
    if not isinstance(value, str):
        return None
    text = value.strip()
    if not text:
        return None
    try:
        parsed = datetime.fromisoformat(text)
    except ValueError:
        return None
    if parsed.tzinfo is None:
        return None
    return parsed


def cookie_status(taken_at, now=None):
    """Trạng thái của một cookie từ mốc lấy.

    Trả về (status, remaining_seconds, expires_at):
    - song: còn hơn 24 giờ.
    - sap-het: còn 0–24 giờ.
    - het-han: đã quá hạn 7 ngày.
    - khong-ro: không có mốc lấy (remaining None, expires None).
    """
    now = now or datetime.now().astimezone()
    taken = parse_dt(taken_at)
    if taken is None:
        return STATUS_UNKNOWN, None, None
    expires = taken + timedelta(days=COOKIE_TTL_DAYS)
    remaining = (expires - now).total_seconds()
    if remaining <= 0:
        return STATUS_EXPIRED, remaining, expires
    if remaining <= EXPIRING_SOON_SECONDS:
        return STATUS_EXPIRING, remaining, expires
    return STATUS_ALIVE, remaining, expires


def mask_cookie(cookie):
    """Che cookie để hiện trong bảng: 10 ký tự đầu + … + 4 cuối.

    Cookie đầy đủ chỉ lấy qua nút Sao chép (gọi API riêng) —
    bảng danh sách không bao giờ render nguyên cookie.
    """
    text = (cookie or "").strip()
    if not text:
        return ""
    if len(text) <= 16:
        return text[:4] + "…"
    return text[:10] + "…" + text[-4:]


def list_cookies(now=None):
    """Kho cookie: một mục một email (dòng MỚI NHẤT thắng).

    Mỗi mục: email, cookie_masked, has_cookie, invite,
    created_at, cookie_taken_at (ISO hoặc ""), expires_at (ISO
    hoặc None), remaining_seconds (None khi không rõ), status,
    health ("song" / "chet" / "khong-ro" — kết quả kiểm tra
    sống/chết gần nhất, xem cookie_health.py), health_checked_at,
    health_detail. Sắp xếp: hết hạn/sắp hết lên trước, rồi theo thời gian còn
    lại tăng dần.
    """
    now = now or datetime.now().astimezone()
    health_cache = cookie_health.load_health_cache()
    latest = {}
    for row in storage._read_cookie_rows():
        latest[row["email"]] = row
    cookies = []
    for email in sorted(latest):
        row = latest[email]
        taken_at = row.get("cookie_taken_at") or ""
        if not taken_at and row.get("cookie"):
            # Dòng cũ chưa có cột mốc lấy: "Thời gian tạo" là
            # lúc append dòng = lúc lấy cookie.
            taken_at = row.get("created_at") or ""
        status, remaining, expires = cookie_status(taken_at, now)
        health = health_cache.get(email.lower(), {})
        cookies.append(
            {
                "email": email,
                "cookie_masked": mask_cookie(row.get("cookie") or ""),
                "has_cookie": bool((row.get("cookie") or "").strip()),
                "invite": row.get("invite") or "",
                "created_at": row.get("created_at") or "",
                "cookie_taken_at": (
                    parse_dt(taken_at).isoformat(timespec="seconds")
                    if parse_dt(taken_at)
                    else ""
                ),
                "expires_at": (
                    expires.isoformat(timespec="seconds")
                    if expires
                    else None
                ),
                "remaining_seconds": (
                    int(remaining) if remaining is not None else None
                ),
                "status": status,
                "health": health.get("status", cookie_health.STATUS_UNKNOWN),
                "health_checked_at": health.get("checked_at", ""),
                "health_detail": health.get("detail", ""),
            }
        )

    def _sort_key(item):
        order = {
            STATUS_EXPIRED: 0,
            STATUS_EXPIRING: 1,
            STATUS_UNKNOWN: 2,
            STATUS_ALIVE: 3,
        }
        remaining = item["remaining_seconds"]
        return (
            order[item["status"]],
            remaining if remaining is not None else 10**18,
        )

    cookies.sort(key=_sort_key)
    return cookies


def cookie_stats(cookies):
    """Đếm theo trạng thái cho thẻ thống kê trên dashboard."""
    stats = {
        "total": len(cookies),
        "song": 0,
        "sap-het": 0,
        "het-han": 0,
        "khong-ro": 0,
        "khong-co-cookie": 0,
        "chet": 0,
        "hoat-dong": 0,
    }
    for item in cookies:
        stats[item["status"]] = stats.get(item["status"], 0) + 1
        if not item["has_cookie"]:
            stats["khong-co-cookie"] += 1
        if item.get("health") == cookie_health.STATUS_DEAD:
            stats["chet"] += 1
        elif item.get("health") == cookie_health.STATUS_ALIVE:
            stats["hoat-dong"] += 1
    return stats
