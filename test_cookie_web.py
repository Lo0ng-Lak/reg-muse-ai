"""Kiểm tra web quản lý kho cookie (cookie_vault + cookie_web + storage).

Chạy:  python test_cookie_web.py
Không cần Windows hay trình duyệt thật: keyring được thay bằng
bản giả trong bộ nhớ; kho SQLite (accounts.db), accounts.xlsx và
lịch sử cookie trỏ sang thư mục tạm. Luồng đăng nhập lại thật
(Playwright) KHÔNG được kiểm tra ở đây — cần chạy trên máy
Windows của Long.
"""

import json
import sys
import tempfile
import threading
import time
import urllib.request
from datetime import datetime, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

TMP = Path(tempfile.mkdtemp(prefix="reg-test-cookie-"))


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

import paths

paths.APP_DATA = TMP
paths.data_file = lambda name: TMP / name

import storage
import cookie_vault
import cookie_web
import cookie_health

storage.SPREADSHEET_FILE = TMP / "accounts.xlsx"
storage.PENDING_SPREADSHEET_FILE = TMP / "accounts.pending.xlsx"
storage.ACCOUNTS_DB_FILE = TMP / "accounts.db"

COOKIE_A = "sessionid=aaa111; c_user=1001; sb=bbb222"
COOKIE_B = "sessionid=ccc333; c_user=1002; sb=ddd444"
# Cookie có phiên đăng nhập thật (hatch_sess) — dùng cho test
# kiểm tra sống/chết.
COOKIE_LIVE = "hatch_sess=live123; c_user=1001; sb=bbb222"
COOKIE_NO_SESS = "c_user=1001; sb=bbb222"


def _now():
    return datetime.now().astimezone()


def _iso(dt):
    return dt.isoformat(timespec="seconds")


def reset():
    # Đặt lại trạng thái nhập kho: mỗi test tự nhập từ fixture
    # Excel của mình (không thì test đầu tiên nhập xong sẽ làm
    # các test sau bỏ qua bước nhập).
    storage._ACCOUNTS_MIGRATION_STATE = "idle"
    storage._ACCOUNTS_MIGRATION_RESULT = None
    for name in (
        "accounts.xlsx",
        "accounts.tmp.xlsx",
        "accounts.pending.xlsx",
        "accounts.pending.tmp.xlsx",
        "accounts.db",
        "accounts.db-wal",
        "accounts.db-shm",
        "cookie_history.json",
        "cookie_health.json",
        "meta_accounts.xlsx",
    ):
        try:
            (TMP / name).unlink()
        except OSError:
            pass


# --- Hạn sống 7 ngày ------------------------------------------------


def test_status_alive():
    status, remaining, expires = cookie_vault.cookie_status(_iso(_now()))
    assert status == cookie_vault.STATUS_ALIVE, status
    assert 6 * 86400 < remaining <= 7 * 86400, remaining
    assert expires is not None


def test_status_expiring_soon():
    taken = _now() - timedelta(days=6, hours=12)
    status, remaining, _ = cookie_vault.cookie_status(_iso(taken))
    assert status == cookie_vault.STATUS_EXPIRING, status
    assert 0 < remaining <= 24 * 3600, remaining


def test_status_expired():
    taken = _now() - timedelta(days=8)
    status, remaining, _ = cookie_vault.cookie_status(_iso(taken))
    assert status == cookie_vault.STATUS_EXPIRED, status
    assert remaining < 0, remaining


def test_status_boundary_24h():
    # Còn đúng 24 giờ -> sắp hết; còn 25 giờ -> còn sống.
    taken_24h = _now() - timedelta(days=6)
    status, _, _ = cookie_vault.cookie_status(_iso(taken_24h))
    assert status == cookie_vault.STATUS_EXPIRING, status
    taken_25h = _now() - timedelta(days=5, hours=23)
    status, _, _ = cookie_vault.cookie_status(_iso(taken_25h))
    assert status == cookie_vault.STATUS_ALIVE, status


def test_status_unknown():
    status, remaining, expires = cookie_vault.cookie_status("")
    assert status == cookie_vault.STATUS_UNKNOWN, status
    assert remaining is None and expires is None
    status, _, _ = cookie_vault.cookie_status("không phải ngày")
    assert status == cookie_vault.STATUS_UNKNOWN, status


def test_mask_cookie():
    masked = cookie_vault.mask_cookie(COOKIE_A)
    assert masked.startswith(COOKIE_A[:10]) and masked.endswith(
        COOKIE_A[-4:]
    ), masked
    assert COOKIE_A[11:-4] not in masked
    assert cookie_vault.mask_cookie("") == ""
    assert cookie_vault.mask_cookie("ngan") == "ngan…"


# --- Sheet Excel ----------------------------------------------------


def test_append_row_records_taken_at():
    storage._append_spreadsheet_row(
        "a@fexpost.com", "secret", cookie=COOKIE_A, invite="ABC123"
    )
    rows = storage._read_cookie_rows()
    assert len(rows) == 1, rows
    assert rows[0]["email"] == "a@fexpost.com"
    assert rows[0]["cookie"] == COOKIE_A
    taken = cookie_vault.parse_dt(rows[0]["cookie_taken_at"])
    assert taken is not None, rows[0]
    assert abs((_now() - taken).total_seconds()) < 120


def test_append_row_without_cookie_leaves_taken_at_empty():
    storage._append_spreadsheet_row("b@fexpost.com", "secret")
    rows = storage._read_cookie_rows()
    assert len(rows) == 1, rows
    assert rows[0]["cookie"] == ""
    assert rows[0]["cookie_taken_at"] == ""


def test_old_six_column_sheet_still_reads():
    from openpyxl import Workbook

    workbook = Workbook()
    sheet = workbook.active
    sheet.append(
        [
            "Thời gian tạo",
            "Email Muse.ai",
            "Mật khẩu hộp thư",
            "Cookie phiên đăng nhập",
            "Mã mời",
            "Link lấy OTP",
        ]
    )
    sheet.append(
        ["2026-11-10T10:00:00+07:00", "old@fexpost.com", "s", COOKIE_A, "", ""]
    )
    workbook.save(storage.SPREADSHEET_FILE)
    rows = storage._read_cookie_rows()
    assert len(rows) == 1, rows
    assert rows[0]["email"] == "old@fexpost.com"
    assert rows[0]["cookie"] == COOKIE_A
    assert rows[0]["cookie_taken_at"] == ""
    # Và append thêm dòng mới vẫn khớp cột cũ (mã mời không lệch).
    storage._append_spreadsheet_row(
        "new@fexpost.com", "s", cookie=COOKIE_B, invite="XYZ789"
    )
    rows = storage._read_cookie_rows()
    assert len(rows) == 2, rows
    assert rows[1]["invite"] == "XYZ789"
    assert rows[1]["cookie"] == COOKIE_B


def test_update_cookie_row_in_place():
    storage._append_spreadsheet_row(
        "a@fexpost.com", "s", cookie=COOKIE_A, invite="ABC123"
    )
    # Ghi lại cùng email: kho SQLite chỉ giữ một dòng, lần ghi
    # sau gộp vào dòng cũ thay vì thêm dòng mới (khác Excel cũ:
    # append là thêm dòng).
    storage._append_spreadsheet_row(
        "a@fexpost.com", "s", cookie=COOKIE_A, invite="ABC123"
    )
    rows = storage._read_cookie_rows()
    assert len(rows) == 1, rows
    new_taken = _iso(_now())
    assert storage._update_cookie_row("a@fexpost.com", COOKIE_B, new_taken)
    rows = storage._read_cookie_rows()
    assert len(rows) == 1, rows
    assert rows[0]["cookie"] == COOKIE_B
    assert rows[0]["cookie_taken_at"] == new_taken
    assert rows[0]["invite"] == "ABC123"  # mã mời giữ nguyên
    assert not storage._update_cookie_row("nope@fexpost.com", "x", new_taken)


def test_clear_cookie_row():
    storage._append_spreadsheet_row(
        "a@fexpost.com", "s", cookie=COOKIE_A, invite="ABC123"
    )
    assert storage._clear_cookie_row("a@fexpost.com")
    rows = storage._read_cookie_rows()
    assert rows[0]["cookie"] == ""
    assert rows[0]["cookie_taken_at"] == ""
    assert rows[0]["invite"] == "ABC123"  # dòng vẫn giữ lại
    assert not storage._clear_cookie_row("nope@fexpost.com")


# --- Kho cookie (vault) ----------------------------------------------


def test_list_cookies_dedupes_and_sorts():
    storage._append_spreadsheet_row(
        "alive@fexpost.com",
        "s",
        cookie=COOKIE_A,
        cookie_taken_at=_iso(_now()),
    )
    storage._append_spreadsheet_row(
        "soon@fexpost.com",
        "s",
        cookie=COOKIE_A,
        cookie_taken_at=_iso(_now() - timedelta(days=6, hours=12)),
    )
    storage._append_spreadsheet_row(
        "dead@fexpost.com",
        "s",
        cookie=COOKIE_A,
        cookie_taken_at=_iso(_now() - timedelta(days=9)),
    )
    # Trùng email: dòng mới thắng.
    storage._append_spreadsheet_row(
        "alive@fexpost.com",
        "s",
        cookie=COOKIE_B,
        cookie_taken_at=_iso(_now() - timedelta(days=6, hours=12)),
    )
    cookies = cookie_vault.list_cookies()
    assert [item["email"] for item in cookies] == [
        "dead@fexpost.com",
        "alive@fexpost.com",
        "soon@fexpost.com",
    ], cookies
    alive = next(
        item for item in cookies if item["email"] == "alive@fexpost.com"
    )
    assert alive["status"] == cookie_vault.STATUS_EXPIRING
    assert alive["cookie_masked"].startswith(COOKIE_B[:10])
    stats = cookie_vault.cookie_stats(cookies)
    assert stats["total"] == 3
    assert stats["het-han"] == 1 and stats["sap-het"] == 2


def test_history_append_read():
    storage._append_cookie_history("lay", "a@fexpost.com", "lần đầu")
    storage._append_cookie_history("lam-moi", "a@fexpost.com", "lấy lại")
    history = storage._read_cookie_history()
    assert len(history) == 2, history
    assert history[0]["action"] == "lam-moi"  # mới nhất trước
    assert history[1]["action"] == "lay"


# --- HTTP ------------------------------------------------------------


def _http(method, url, payload=None):
    data = (
        json.dumps(payload).encode("utf-8") if payload is not None else None
    )
    request = urllib.request.Request(url, data=data, method=method)
    if data is not None:
        request.add_header("Content-Type", "application/json")
    try:
        with urllib.request.urlopen(request, timeout=10) as response:
            return response.status, response.read()
    except urllib.error.HTTPError as error:
        return error.code, error.read()


def test_http_api():
    storage._append_spreadsheet_row(
        "a@fexpost.com",
        "s",
        cookie=COOKIE_A,
        invite="ABC123",
        cookie_taken_at=_iso(_now()),
    )
    httpd, url = cookie_web.start_server(port=0)
    try:
        status, body = _http("GET", url + "/")
        assert status == 200, status
        assert "Kho Cookie" in body.decode("utf-8")
        status, body = _http("GET", url + "/api/cookies")
        assert status == 200, status
        data = json.loads(body)
        assert data["stats"]["total"] == 1, data
        assert data["cookies"][0]["email"] == "a@fexpost.com"
        assert data["cookies"][0]["status"] == "song"
        assert COOKIE_A not in body.decode("utf-8")  # bảng không lộ cookie
        status, body = _http(
            "GET", url + "/api/cookie?email=a%40fexpost.com"
        )
        assert status == 200, status
        assert json.loads(body)["cookie"] == COOKIE_A
        status, body = _http(
            "POST", url + "/api/delete", {"email": "a@fexpost.com"}
        )
        assert status == 200, status
        assert json.loads(body)["ok"] is True
        status, body = _http("GET", url + "/api/cookies")
        assert json.loads(body)["cookies"][0]["has_cookie"] is False
    finally:
        httpd.shutdown()
        httpd.server_close()


def test_http_refresh_unknown_email_fails_cleanly():
    httpd, url = cookie_web.start_server(port=0)
    try:
        status, body = _http(
            "POST",
            url + "/api/refresh",
            {"email": "khong-co@fexpost.com", "direct": False},
        )
        assert status == 202, (status, body)
        job_id = json.loads(body)["job_id"]
        deadline = time.time() + 15
        job = None
        while time.time() < deadline:
            _, body = _http("GET", url + "/api/jobs")
            jobs = json.loads(body)["jobs"]
            job = next(
                (item for item in jobs if item["id"] == job_id), None
            )
            if job and job["status"] == "done":
                break
            time.sleep(0.5)
        assert job is not None and job["status"] == "done", job
        assert "0 thành công, 1 thất bại" in job["result"], job
        assert any(
            "không tìm thấy mail" in line for line in job["log"]
        ), job["log"]
    finally:
        httpd.shutdown()
        httpd.server_close()


# --- Kiểm tra sống/chết (cookie_health) ----------------------------


def _fake_fetch(final_url, body):
    def fetch(url, cookie, timeout):
        return final_url, body

    return fetch


def test_health_alive_when_chat_ui():
    status, _detail = cookie_health.check_cookie(
        COOKIE_LIVE,
        fetcher=_fake_fetch(
            "https://www.meta.ai/",
            '<html><body><textarea placeholder="Ask Meta AI">'
            "</textarea></body></html>",
        ),
    )
    assert status == cookie_health.STATUS_ALIVE, status


def test_health_dead_when_redirected_to_login():
    status, _detail = cookie_health.check_cookie(
        COOKIE_LIVE,
        fetcher=_fake_fetch(
            "https://www.meta.ai/login",
            "<html><body><button>Log in</button></body></html>",
        ),
    )
    assert status == cookie_health.STATUS_DEAD, status


def test_health_dead_when_login_button_without_chat():
    status, _detail = cookie_health.check_cookie(
        COOKIE_LIVE,
        fetcher=_fake_fetch(
            "https://www.meta.ai/",
            "<html><body><h1>Log in to Meta AI</h1></body></html>",
        ),
    )
    assert status == cookie_health.STATUS_DEAD, status


def test_health_unknown_when_page_unrecognised():
    status, _detail = cookie_health.check_cookie(
        COOKIE_LIVE,
        fetcher=_fake_fetch(
            "https://www.meta.ai/", "<html><body>hello</body></html>"
        ),
    )
    assert status == cookie_health.STATUS_UNKNOWN, status


def test_health_dead_without_hatch_sess():
    status, detail = cookie_health.check_cookie(COOKIE_NO_SESS)
    assert status == cookie_health.STATUS_DEAD, status
    assert "hatch_sess" in detail


def test_health_unknown_on_network_error():
    import urllib.error

    def fetch(url, cookie, timeout):
        raise urllib.error.URLError("no route to host")

    status, _detail = cookie_health.check_cookie(
        COOKIE_LIVE, fetcher=fetch
    )
    assert status == cookie_health.STATUS_UNKNOWN, status


def test_health_cache_roundtrip_and_list_enrichment():
    entry = cookie_health.check_and_cache(
        "a@fexpost.com",
        COOKIE_LIVE,
        fetcher=_fake_fetch(
            "https://www.meta.ai/",
            '<html><body><textarea placeholder="Ask Meta AI">'
            "</textarea></body></html>",
        ),
    )
    assert entry["status"] == cookie_health.STATUS_ALIVE, entry
    assert entry["checked_at"]
    cache = cookie_health.load_health_cache()
    assert cache["a@fexpost.com"]["status"] == "song", cache
    storage._append_spreadsheet_row(
        "a@fexpost.com",
        "s",
        cookie=COOKIE_LIVE,
        cookie_taken_at=_iso(_now()),
    )
    cookies = cookie_vault.list_cookies()
    assert cookies[0]["health"] == "song", cookies[0]
    assert cookies[0]["health_checked_at"] == entry["checked_at"]


def test_http_check_endpoint():
    storage._append_spreadsheet_row(
        "a@fexpost.com",
        "s",
        cookie=COOKIE_LIVE,
        cookie_taken_at=_iso(_now()),
    )
    orig = cookie_health._default_fetch
    cookie_health._default_fetch = _fake_fetch(
        "https://www.meta.ai/",
        '<html><body><textarea placeholder="Ask Meta AI">'
        "</textarea></body></html>",
    )
    httpd, url = cookie_web.start_server(port=0)
    try:
        status, body = _http(
            "POST", url + "/api/check", {"email": "a@fexpost.com"}
        )
        assert status == 200, (status, body)
        data = json.loads(body)
        assert data["status"] == "song", data
        status, body = _http("GET", url + "/api/cookies")
        assert json.loads(body)["cookies"][0]["health"] == "song"
    finally:
        cookie_health._default_fetch = orig
        httpd.shutdown()
        httpd.server_close()


def test_http_check_all_job():
    for email in ("a@fexpost.com", "b@fexpost.com"):
        storage._append_spreadsheet_row(
            email,
            "s",
            cookie=COOKIE_LIVE,
            cookie_taken_at=_iso(_now()),
        )
    orig = cookie_health._default_fetch
    cookie_health._default_fetch = _fake_fetch(
        "https://www.meta.ai/login",
        "<html><body><button>Log in</button></body></html>",
    )
    httpd, url = cookie_web.start_server(port=0)
    try:
        status, body = _http("POST", url + "/api/check-all", {})
        assert status == 202, (status, body)
        job_id = json.loads(body)["job_id"]
        deadline = time.time() + 15
        job = None
        while time.time() < deadline:
            _, body = _http("GET", url + "/api/jobs")
            jobs = json.loads(body)["jobs"]
            job = next(
                (item for item in jobs if item["id"] == job_id), None
            )
            if job and job["status"] == "done":
                break
            time.sleep(0.5)
        assert job is not None and job["status"] == "done", job
        assert "2 die" in job["result"], job
    finally:
        cookie_health._default_fetch = orig
        httpd.shutdown()
        httpd.server_close()


def test_read_many_rows_complete():
    # Vòng đọc mới (iter_rows một lượt xuôi) phải đọc đủ mọi
    # dòng, đúng thứ tự, không bỏ dòng cuối. Vòng cũ dùng
    # sheet.cell() trong read_only nên với 2k dòng mất hàng phút.
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.append(list(storage.SPREADSHEET_HEADERS))
    for i in range(300):
        ws.append(
            [
                _iso(_now()),
                f"acc{i:03d}@fexpost.com",
                "s",
                f"hatch_sess=c{i:03d}",
                f"INV{i:03d}",
                "",
                _iso(_now()),
            ]
        )
    wb.save(storage.SPREADSHEET_FILE)
    rows = storage._read_cookie_rows()
    assert len(rows) == 300, len(rows)
    assert rows[0]["email"] == "acc000@fexpost.com"
    assert rows[-1]["email"] == "acc299@fexpost.com"
    assert rows[-1]["cookie"] == "hatch_sess=c299"
    assert rows[150]["invite"] == "INV150"


def test_last_invite_from_sheet_reads_last_code():
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.append(list(storage.SPREADSHEET_HEADERS))
    for email, invite in [
        ("a@fexpost.com", "ABC123"),
        ("b@fexpost.com", ""),
        ("c@fexpost.com", "không phải mã"),
        ("d@fexpost.com", "ZZZ999"),
    ]:
        ws.append([_iso(_now()), email, "s", "", invite, "", ""])
    wb.save(storage.SPREADSHEET_FILE)
    assert storage._last_invite_from_sheet() == "ZZZ999"


def test_http_check_all_expiring_scope():
    from datetime import timedelta

    for email, age_days in [
        ("song@fexpost.com", 0),
        ("saphet@fexpost.com", 6.5),
        ("hethan@fexpost.com", 8),
    ]:
        storage._append_spreadsheet_row(
            email,
            "s",
            cookie=COOKIE_LIVE,
            cookie_taken_at=_iso(_now() - timedelta(days=age_days)),
        )
    orig = cookie_health._default_fetch
    cookie_health._default_fetch = _fake_fetch(
        "https://www.meta.ai/",
        '<html><body><textarea placeholder="Ask Meta AI">'
        "</textarea></body></html>",
    )
    httpd, url = cookie_web.start_server(port=0)
    try:
        status, body = _http(
            "POST",
            url + "/api/check-all",
            {"scope": "expiring"},
        )
        assert status == 202, (status, body)
        job_id = json.loads(body)["job_id"]
        assert json.loads(body)["count"] == 2, body
        deadline = time.time() + 15
        job = None
        while time.time() < deadline:
            _, body = _http("GET", url + "/api/jobs")
            jobs = json.loads(body)["jobs"]
            job = next(
                (item for item in jobs if item["id"] == job_id), None
            )
            if job and job["status"] == "done":
                break
            time.sleep(0.5)
        assert job is not None and job["status"] == "done", job
        assert "2 hoạt động" in job["result"], job
    finally:
        cookie_health._default_fetch = orig
        httpd.shutdown()
        httpd.server_close()


def test_http_import_export_excel():
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.append(list(storage.SPREADSHEET_HEADERS))
    ws.append(
        [
            _iso(_now()),
            "imp@fexpost.com",
            "s",
            COOKIE_A,
            "ABC123",
            "",
            _iso(_now()),
        ]
    )
    wb.save(storage.SPREADSHEET_FILE)
    httpd, url = cookie_web.start_server(port=0)
    try:
        status, body = _http("POST", url + "/api/import-excel", {})
        assert status == 200, (status, body)
        data = json.loads(body)
        # Server đã tự nhập file lúc mở (luồng nhập nền) nên
        # lần nhập tay này có thể đếm là "cập nhật" thay vì
        # "mới" — quan trọng là dòng đó có trong kho, đúng 1 dòng.
        assert (
            data["imported"] + data["updated"] == 1
            and data["total"] == 1
        ), data
        status, body = _http("GET", url + "/api/cookies")
        assert json.loads(body)["stats"]["total"] == 1, body
        status, body = _http("GET", url + "/api/export-excel")
        assert status == 200, status
        assert body[:2] == b"PK"  # xlsx là file zip
        import io as _io
        from openpyxl import load_workbook as _load

        values = list(
            _load(_io.BytesIO(body)).active.iter_rows(values_only=True)
        )
        assert values[1][1] == "imp@fexpost.com", values
        assert values[1][3] == COOKIE_A, values
    finally:
        httpd.shutdown()
        httpd.server_close()


def test_dashboard_script_parses_and_handlers_escaped():
    """Hồi quy lỗi 2026-11-17: nút Sao chép/Kiểm tra/Đăng nhập
    lại/Xoá trong bảng dùng `\"` trong chuỗi Python nên JS nhận
    về bị vỡ ngoặc — cả khối <script> không chạy, trang kẹt ở
    "Đang tải…" mãi dù API vẫn trả dữ liệu. Dấu nháy của tham
    số email trong onclick phải là thực thể HTML &quot;."""
    import re as _re

    match = _re.search(
        r"<script>(.*)</script>", cookie_web.DASHBOARD_HTML, _re.S
    )
    assert match, "không tìm thấy khối <script> trong trang"
    script = match.group(1)
    # Mẫu vỡ (dấu nháy trần sau tên hàm) không được xuất hiện.
    assert '""+esc(c.email)' not in script, script
    # Thẻ job không được liệt kê cả nghìn mail (lỗi 2026-11-17:
    # tường email làm xấu giao diện) — chỉ hiện số lượng +
    # tiến độ; trạng thái "Đang kiểm tra…" hiện ở từng dòng.
    assert "j.emails.map(esc).join" not in script, script
    assert "(j.emails||[]).length" in script, script
    assert "Đang kiểm tra…" in script, script
    assert "healthCell" in script, script
    for handler in ("copyCookie", "checkOne", "refreshOne", "deleteCookie"):
        assert (
            f"{handler}(&quot;\"+esc(c.email)+\"&quot;,this)"
            in script
        ), handler
    # Nếu có node thì kiểm tra cú pháp JS thật; không có thì bỏ qua.
    import shutil as _shutil
    import subprocess as _subprocess

    node = _shutil.which("node")
    if node is None:
        return
    tmp_js = TMP / "dashboard.js"
    tmp_js.write_text(script, encoding="utf-8")
    result = _subprocess.run(
        [node, "--check", str(tmp_js)], capture_output=True, text=True
    )
    assert result.returncode == 0, result.stderr


def test_stats_counts_unchecked_separately_from_dead():
    """Die: 0 chỉ đúng khi đã kiểm tra: số chưa từng kiểm tra
    phải hiện riêng (chua-kiem-tra), không được lẫn vào Die."""
    for email in ("a@fexpost.com", "b@fexpost.com", "c@fexpost.com"):
        storage._append_spreadsheet_row(
            email,
            "s",
            cookie=COOKIE_LIVE,
            cookie_taken_at=_iso(_now()),
        )
    stats = cookie_vault.cookie_stats(cookie_vault.list_cookies())
    assert stats["total"] == 3, stats
    assert stats["chua-kiem-tra"] == 3, stats
    assert stats["chet"] == 0 and stats["hoat-dong"] == 0, stats
    cookie_health.check_and_cache(
        "a@fexpost.com",
        COOKIE_LIVE,
        fetcher=_fake_fetch(
            "https://www.meta.ai/",
            '<html><body><textarea placeholder="Ask Meta AI">'
            "</textarea></body></html>",
        ),
    )
    cookie_health.check_and_cache(
        "b@fexpost.com",
        COOKIE_LIVE,
        fetcher=_fake_fetch(
            "https://www.meta.ai/login",
            "<html><body><button>Log in</button></body></html>",
        ),
    )
    stats = cookie_vault.cookie_stats(cookie_vault.list_cookies())
    assert stats["chua-kiem-tra"] == 1, stats
    assert stats["hoat-dong"] == 1 and stats["chet"] == 1, stats


def test_select_refresh_emails_scopes():
    cookies = [
        {"email": "song@x.com", "status": "song", "health": "song"},
        {"email": "sap@x.com", "status": "sap-het", "health": "khong-ro"},
        {"email": "het1@x.com", "status": "het-han", "health": "khong-ro"},
        {"email": "het2@x.com", "status": "het-han", "health": "chet"},
        {"email": "nocookie@x.com", "status": "khong-ro", "health": "khong-ro"},
    ]
    assert cookie_web.select_refresh_emails(
        cookies, "expired"
    ) == ["het1@x.com", "het2@x.com"]
    assert cookie_web.select_refresh_emails(
        cookies, "expiring"
    ) == ["sap@x.com", "het1@x.com", "het2@x.com"]
    assert cookie_web.select_refresh_emails(
        cookies, "dead"
    ) == ["het2@x.com"]
    assert cookie_web.select_refresh_emails(cookies, "all") == [
        item["email"] for item in cookies
    ]
    # Diện lạ rơi về mặc định expiring (sắp hết + đã hết).
    assert cookie_web.select_refresh_emails(
        cookies, "lạ"
    ) == ["sap@x.com", "het1@x.com", "het2@x.com"]


def test_select_check_emails_scopes():
    now = _now()
    fresh = _iso(now - timedelta(minutes=30))
    old = _iso(now - timedelta(hours=25))
    cookies = [
        {"email": "never@x.com", "status": "song", "has_cookie": True, "health_checked_at": ""},
        {"email": "old@x.com", "status": "sap-het", "has_cookie": True, "health_checked_at": old},
        {"email": "fresh@x.com", "status": "song", "has_cookie": True, "health_checked_at": fresh},
        {"email": "het@x.com", "status": "het-han", "has_cookie": True, "health_checked_at": fresh},
        {"email": "nocookie@x.com", "status": "khong-ro", "has_cookie": False, "health_checked_at": ""},
    ]
    # stale: chưa kiểm tra bao giờ hoặc quá 24 giờ trước.
    assert cookie_web.select_check_emails(
        cookies, "stale"
    ) == ["never@x.com", "old@x.com"]
    assert cookie_web.select_check_emails(
        cookies, "expiring"
    ) == ["old@x.com", "het@x.com"]
    assert cookie_web.select_check_emails(cookies, "all") == [
        "never@x.com",
        "old@x.com",
        "fresh@x.com",
        "het@x.com",
    ]


def test_start_check_all_khong_xep_chong_job():
    web = cookie_web.CookieWeb()
    job = web._new_job(["a@fexpost.com"], "check")
    assert web.start_check_all(["b@fexpost.com"]) == job["id"]
    assert len(web.jobs) == 1


def test_check_job_reports_done_count():
    for email in ("a@fexpost.com", "b@fexpost.com"):
        storage._append_spreadsheet_row(
            email,
            "s",
            cookie=COOKIE_LIVE,
            cookie_taken_at=_iso(_now()),
        )
    orig = cookie_health._default_fetch
    cookie_health._default_fetch = _fake_fetch(
        "https://www.meta.ai/",
        '<html><body><textarea placeholder="Ask Meta AI">'
        "</textarea></body></html>",
    )
    try:
        web = cookie_web.CookieWeb()
        job = web._new_job(
            ["a@fexpost.com", "b@fexpost.com"], "check"
        )
        assert job["done"] == 0, job
        web._run_check_job(job)
        assert job["done"] == 2, job
        assert job["status"] == "done", job
        assert "2 hoạt động" in job["result"], job
    finally:
        cookie_health._default_fetch = orig


TESTS = [
    test_status_alive,
    test_status_expiring_soon,
    test_status_expired,
    test_status_boundary_24h,
    test_status_unknown,
    test_mask_cookie,
    test_append_row_records_taken_at,
    test_append_row_without_cookie_leaves_taken_at_empty,
    test_old_six_column_sheet_still_reads,
    test_update_cookie_row_in_place,
    test_clear_cookie_row,
    test_list_cookies_dedupes_and_sorts,
    test_history_append_read,
    test_http_api,
    test_http_refresh_unknown_email_fails_cleanly,
    test_health_alive_when_chat_ui,
    test_health_dead_when_redirected_to_login,
    test_health_dead_when_login_button_without_chat,
    test_health_unknown_when_page_unrecognised,
    test_health_dead_without_hatch_sess,
    test_health_unknown_on_network_error,
    test_health_cache_roundtrip_and_list_enrichment,
    test_http_check_endpoint,
    test_http_check_all_job,
    test_read_many_rows_complete,
    test_last_invite_from_sheet_reads_last_code,
    test_http_check_all_expiring_scope,
    test_http_import_export_excel,
    test_dashboard_script_parses_and_handlers_escaped,
    test_stats_counts_unchecked_separately_from_dead,
    test_select_refresh_emails_scopes,
    test_select_check_emails_scopes,
    test_start_check_all_khong_xep_chong_job,
    test_check_job_reports_done_count,
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
