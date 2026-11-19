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
    # Vòng kiểm tra nền sẽ mở trình duyệt. Test không bật vòng đó.
    cookie_web.TU_DONG_KIEM_TRA = False
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
    storage._append_cookie_history("lam-moi", "b@fexpost.com", "lấy lại")
    storage._append_cookie_history("xoa", "a@fexpost.com", "xoá cookie")
    assert storage._cookie_history_counts() == {
        "lay": 1, "lam-moi": 2, "xoa": 1,
    }


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


def _fake_browser_checker_class(status):
    """Lớp BrowserChecker giả cho test HTTP: trả trạng thái cố
    định, không cần trình duyệt thật."""

    class _Fake:
        def start(self):
            pass

        def smoke_test(self):
            pass

        def check(self, cookie):
            return status, "kiểm tra giả"

        def close(self):
            pass

    return _Fake


def test_http_check_endpoint():
    storage._append_spreadsheet_row(
        "a@fexpost.com",
        "s",
        cookie=COOKIE_LIVE,
        cookie_taken_at=_iso(_now()),
    )
    orig = cookie_health.BrowserChecker
    cookie_health.BrowserChecker = _fake_browser_checker_class("song")
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
        cookie_health.BrowserChecker = orig
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
    orig = cookie_health.BrowserChecker
    cookie_health.BrowserChecker = _fake_browser_checker_class("chet")
    orig_hook = (
        cookie_web.CookieWeb._tu_dong_dang_nhap_lai_cookie_chet
    )
    cookie_web.CookieWeb._tu_dong_dang_nhap_lai_cookie_chet = (
        lambda self, job, dead: None
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
        cookie_health.BrowserChecker = orig
        cookie_web.CookieWeb._tu_dong_dang_nhap_lai_cookie_chet = (
            orig_hook
        )
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
    orig = cookie_health.BrowserChecker
    cookie_health.BrowserChecker = _fake_browser_checker_class("song")
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
        cookie_health.BrowserChecker = orig
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
        cookie_health.save_health_cache({
            "imp@fexpost.com": {
                "status": "song",
                "checked_at": _iso(_now()),
                "detail": "còn sống",
            }
        })
        storage._append_spreadsheet_row(
            "die@fexpost.com",
            "s",
            cookie="hatch_sess=die",
            cookie_taken_at=_iso(_now()),
        )
        cookie_health.save_health_cache({
            "imp@fexpost.com": {
                "status": "song",
                "checked_at": _iso(_now()),
                "detail": "còn sống",
            },
            "die@fexpost.com": {
                "status": "chet",
                "checked_at": _iso(_now()),
                "detail": "die",
            },
        })
        status, body = _http("GET", url + "/api/export-excel")
        assert status == 200, status
        assert body[:2] == b"PK"  # xlsx là file zip
        import io as _io
        from openpyxl import load_workbook as _load

        values = list(
            _load(_io.BytesIO(body)).active.iter_rows(values_only=True)
        )
        emails = [row[1] for row in values[1:]]
        assert emails == ["imp@fexpost.com"], values
        assert values[1][3] == COOKIE_A, values
    finally:
        httpd.shutdown()
        httpd.server_close()


def test_chon_nhieu_the_va_dang_ky_vao_quan_ly():
    """Tick nhiều thẻ thì lượt chạy chỉ lấy nhóm đó. Đăng ký Muse
    và đăng nhập lại ghi đúng kho Quản lý. Meta nằm ở file Meta,
    chưa có cookie Muse thì không hiện là cookie."""
    import web_console

    added, errors, skipped = storage._import_bulk_cards(
        "4000000000000002|12/30|123|10001|Thẻ A\n"
        "4000000000000010|12/30|123|10001|Thẻ B\n"
        "4000000000000028|12/30|123|10001|Thẻ C\n"
    )
    assert errors == [], errors
    assert skipped == 0
    assert added == ["Thẻ A", "Thẻ B", "Thẻ C"], added
    vault = storage._load_vault()
    by_label = {card["label"]: card["id"] for card in vault}
    storage._save_vault_picked([by_label["Thẻ A"], by_label["Thẻ B"]])
    cards, manual = web_console._run_cards()
    assert [card.label for card in cards] == ["Thẻ A", "Thẻ B"]
    assert manual is False
    console = web_console.WebConsole()
    console.cards("limit-many", {"ids": [by_label["Thẻ A"], by_label["Thẻ B"]], "limit": 4})
    limits = {card["label"]: card["limit"] for card in storage._load_vault()}
    assert limits["Thẻ A"] == 4 and limits["Thẻ B"] == 4, limits
    console.cards("delete-many", {"ids": [by_label["Thẻ C"]]})
    assert "Thẻ C" not in {card["label"] for card in storage._load_vault()}
    assert storage._load_vault_picked() == [by_label["Thẻ A"], by_label["Thẻ B"]]

    storage._append_spreadsheet_row(
        "reg@fexpost.com", "pw", cookie="hatch_sess=reg1", invite="INV9"
    )
    listed = {item["email"]: item for item in cookie_vault.list_cookies()}
    assert listed["reg@fexpost.com"]["has_cookie"] is True
    assert listed["reg@fexpost.com"]["invite"] == "INV9"

    storage._append_meta_row(
        "meta@fexpost.com",
        mail_password="secret-mail",
        username="meta-user",
        provider="tempmail.plus",
    )
    pending = storage._pending_meta_logins(storage.META_SPREADSHEET_FILE)
    assert pending and pending[0]["email"] == "meta@fexpost.com", pending
    assert pending[0]["secret"] == "secret-mail"
    assert "meta@fexpost.com" not in listed

    storage._save_relogin_to_accounts(
        "reg@fexpost.com", "hatch_sess=moi", "pw"
    )
    again = {item["email"]: item for item in cookie_vault.list_cookies()}
    assert again["reg@fexpost.com"]["has_cookie"] is True
    assert storage._cookie_history_counts()["lam-moi"] == 1
    page = cookie_web.DASHBOARD_HTML
    assert "Xoá đã chọn" in page and "card-tick" in page
    assert "Chỉ dùng thẻ đã chọn" in page


def test_register_page_and_console_refuse_empty_proxy():
    httpd, url = cookie_web.start_server(port=0)
    try:
        status, body = _http("GET", url + "/dang-ky")
        page = body.decode("utf-8")
        assert status == 200, status
        assert 'id="viewRegister"' in page
        assert "Quản lý" in page and "Đăng ký" in page
        status, body = _http("GET", url + "/api/console")
        data = json.loads(body)
        assert data["running"] is False
        assert isinstance(data["cards"], list)
        status, body = _http(
            "POST", url + "/api/console/muse/register",
            {"total": 1, "proxies": "", "direct": False},
        )
        assert status == 400, body
        assert "không đăng ký muse" in body.decode("utf-8").lower()
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
    assert "healthBadge" in script, script
    assert "cookie-masked" not in script, script
    for handler in ("checkOne", "refreshOne", "deleteCookie"):
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
    assert stats["chua-xac-minh"] == 0, stats
    assert stats["hoat-dong"] == 1 and stats["chet"] == 1, stats
    # Kiểm tra rồi nhưng chưa rõ (lỗi mạng, như 403) là
    # chưa xác minh — không đếm vào còn sống, cũng không
    # đếm vào chưa kiểm tra.
    def _boom(url, cookie, timeout):
        raise urllib.error.URLError("403")

    cookie_health.check_and_cache(
        "c@fexpost.com", COOKIE_LIVE, fetcher=_boom
    )
    stats = cookie_vault.cookie_stats(cookie_vault.list_cookies())
    assert stats["chua-kiem-tra"] == 0, stats
    assert stats["chua-xac-minh"] == 1, stats
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
        {"email": "chua@x.com", "status": "song", "has_cookie": True, "health_checked_at": fresh, "health": "khong-ro"},
        {"email": "ro@x.com", "status": "song", "has_cookie": True, "health_checked_at": fresh, "health": "song"},
        {"email": "song-2ngay@x.com", "status": "song", "has_cookie": True, "health_checked_at": _iso(now - timedelta(days=2)), "health": "song"},
        {"email": "song-4ngay@x.com", "status": "song", "has_cookie": True, "health_checked_at": _iso(now - timedelta(days=4)), "health": "song"},
        {"email": "song-8ngay@x.com", "status": "song", "has_cookie": True, "health_checked_at": _iso(now - timedelta(days=8)), "health": "song"},
        {"email": "chet-moi@x.com", "status": "song", "has_cookie": True, "health_checked_at": fresh, "health": "chet"},
        {"email": "chet-25h@x.com", "status": "song", "has_cookie": True, "health_checked_at": old, "health": "chet"},
    ]
    # stale: chưa kiểm tra, còn sống quá 7 ngày, hoặc vừa kiểm
    # tra nhưng chưa rõ. Còn sống 4 ngày thì thôi. Die thì thôi.
    assert cookie_web.select_check_emails(
        cookies, "stale"
    ) == [
        "never@x.com",
        "old@x.com",
        "chua@x.com",
        "song-8ngay@x.com",
    ]
    assert cookie_web.select_check_emails(
        cookies, "expiring"
    ) == ["old@x.com", "het@x.com"]
    assert cookie_web.select_check_emails(cookies, "all") == [
        "never@x.com",
        "old@x.com",
        "fresh@x.com",
        "het@x.com",
        "chua@x.com",
        "ro@x.com",
        "song-2ngay@x.com",
        "song-4ngay@x.com",
        "song-8ngay@x.com",
        "chet-moi@x.com",
        "chet-25h@x.com",
    ]


def test_vong_nen_tu_bat_dau_va_khong_xep_chong():
    now = _now()
    fresh = _iso(now - timedelta(minutes=30))
    cookies = [
        {"email": "never@x.com", "has_cookie": True, "health_checked_at": ""},
        {
            "email": "vua-song@x.com",
            "has_cookie": True,
            "health_checked_at": fresh,
            "health": "song",
        },
        {
            "email": "vua-chua-ro@x.com",
            "has_cookie": True,
            "health_checked_at": fresh,
            "health": "khong-ro",
        },
        {"email": "trong@x.com", "has_cookie": False, "health_checked_at": ""},
        {
            "email": "song-2ngay@x.com",
            "has_cookie": True,
            "health_checked_at": _iso(now - timedelta(days=2)),
            "health": "song",
        },
        {
            "email": "song-4ngay@x.com",
            "has_cookie": True,
            "health_checked_at": _iso(now - timedelta(days=4)),
            "health": "song",
        },
        {
            "email": "song-8ngay@x.com",
            "has_cookie": True,
            "health_checked_at": _iso(now - timedelta(days=8)),
            "health": "song",
        },
    ]
    # Chưa kiểm tra và vừa ra chưa rõ đều được xếp lại.
    # Còn sống dưới 7 ngày thì thôi. Còn sống quá 7 ngày thì kiểm.
    # Không có cookie thì thôi.
    assert cookie_web.select_mail_vong_nen(cookies) == [
        "never@x.com",
        "vua-chua-ro@x.com",
        "song-8ngay@x.com",
    ]
    web = cookie_web.CookieWeb()
    called = []
    web.start_check_all = lambda emails: called.append(list(emails)) or "job-x"
    assert web.bat_kiem_tra_neu_can(cookies) == "job-x"
    assert called == [["never@x.com", "vua-chua-ro@x.com", "song-8ngay@x.com"]]
    web.jobs["job-1"] = {
        "kind": "check",
        "status": "running",
        "created_at": fresh,
    }
    assert web.bat_kiem_tra_neu_can(cookies) is None
    assert called == [["never@x.com", "vua-chua-ro@x.com", "song-8ngay@x.com"]]


def test_start_check_all_khong_xep_chong_job():
    web = cookie_web.CookieWeb()
    job = web._new_job(["a@fexpost.com"], "check")
    assert web.start_check_all(["b@fexpost.com"]) == job["id"]
    assert len(web.jobs) == 1


class _FakeChecker:
    """BrowserChecker giả: trả trạng thái cố định, không cần
    trình duyệt thật."""

    def __init__(self, status="song"):
        self.status = status
        self.checked = []
        self.closed = False

    def start(self):
        pass

    def smoke_test(self):
        pass

    def check(self, cookie):
        self.checked.append(cookie)
        return self.status, "vào được giao diện chat"

    def close(self):
        self.closed = True


def test_check_job_reports_done_count():
    for email in ("a@fexpost.com", "b@fexpost.com"):
        storage._append_spreadsheet_row(
            email,
            "s",
            cookie=COOKIE_LIVE,
            cookie_taken_at=_iso(_now()),
        )
    created = []

    def factory():
        checker = _FakeChecker()
        created.append(checker)
        return checker

    web = cookie_web.CookieWeb()
    job = web._new_job(["a@fexpost.com", "b@fexpost.com"], "check")
    assert job["done"] == 0, job
    web._run_check_job(job, checker_factory=factory)
    assert job["done"] == 2, job
    assert job["status"] == "done", job
    assert "2 hoạt động" in job["result"], job
    # 1 probe + 2 luồng (2 mail) = 3 checker, mỗi cái đều đóng;
    # chỉ 2 checker của luồng là kiểm tra thật (mỗi cái 1 mail).
    assert len(created) == 3, len(created)
    assert all(c.closed for c in created), created
    assert sum(len(c.checked) for c in created) == 2, created


def test_check_job_dung_ngay_khi_khong_mo_duoc_trinh_duyet():
    class _BrokenChecker:
        def start(self):
            raise RuntimeError("Thiếu Playwright — test")

        def close(self):
            pass

    orig = cookie_health.BrowserChecker
    cookie_health.BrowserChecker = _BrokenChecker
    try:
        web = cookie_web.CookieWeb()
        job = web._new_job(["a@fexpost.com"], "check")
        web._run_check_job(job)
        assert job["done"] == 0, job
        assert "Không kiểm tra được" in job["result"], job
    finally:
        cookie_health.BrowserChecker = orig


def test_parse_cookie_string():
    assert cookie_health.parse_cookie_string(
        "hatch_sess=abc123; c_user=1001; sb=bbb222"
    ) == [
        ("hatch_sess", "abc123"),
        ("c_user", "1001"),
        ("sb", "bbb222"),
    ]
    assert cookie_health.parse_cookie_string("") == []
    assert cookie_health.parse_cookie_string("khong-co-dau-bang") == []


class _FakePage:
    def __init__(self, url, body, fail_content_times=0):
        self.url = url
        self._body = body
        self._fail_content_times = fail_content_times
        self.closed = False

    def goto(self, url, timeout=None, wait_until=None):
        self.goto_wait_until = wait_until

    def wait_for_timeout(self, ms):
        pass

    def content(self):
        if self._fail_content_times > 0:
            self._fail_content_times -= 1
            raise RuntimeError(
                "Page.content: Unable to retrieve content because "
                "the page is navigating and changing the content."
            )
        return self._body

    def close(self):
        self.closed = True


class _FakeContext:
    def __init__(self, page):
        self.page = page
        self.cookies = None
        self.closed = False

    def add_cookies(self, cookies):
        self.cookies = cookies

    def new_page(self):
        return self.page

    def close(self):
        self.closed = True


class _FakeBrowser:
    def __init__(self, page):
        self.page = page
        self.contexts = []
        self.closed = False

    def new_context(self):
        ctx = _FakeContext(self.page)
        self.contexts.append(ctx)
        return ctx

    def close(self):
        self.closed = True


def _browser_checker_with(page, timeout_ms=None):
    browser = _FakeBrowser(page)
    checker = cookie_health.BrowserChecker(
        launcher=lambda: browser, timeout_ms=timeout_ms
    )
    checker.start()
    return checker, browser


def test_browser_checker_chan_anh_font_video():
    # Ảnh, font, video không giúp phân biệt sống/chết — chặn để
    # trang nhẹ hơn. Script và tài liệu vẫn được tải.
    page = _FakePage(
        "https://www.meta.ai/",
        '<html><body><textarea placeholder="Ask Meta AI">'
        "</textarea></body></html>",
    )
    seen = {}

    class _Request:
        def __init__(self, kind):
            self.resource_type = kind

    class _Route:
        def __init__(self, kind):
            self.request = _Request(kind)
            self.aborted = False
            self.continued = False

        def abort(self):
            self.aborted = True

        def continue_(self):
            self.continued = True

    class _Context(_FakeContext):
        def route(self, pattern, handler):
            seen["pattern"] = pattern
            seen["handler"] = handler

    class _Browser(_FakeBrowser):
        def new_context(self):
            ctx = _Context(self.page)
            self.contexts.append(ctx)
            return ctx

    browser = _Browser(page)
    checker = cookie_health.BrowserChecker(launcher=lambda: browser)
    checker.start()
    status, _detail = checker.check(COOKIE_LIVE)
    assert status == "song", status
    assert seen["pattern"] == "**/*"
    image = _Route("image")
    script = _Route("script")
    seen["handler"](image)
    seen["handler"](script)
    assert image.aborted and not image.continued
    assert script.continued and not script.aborted
    font = _Route("font")
    video = _Route("media")
    seen["handler"](font)
    seen["handler"](video)
    assert font.aborted and video.aborted
    checker.close()


def test_browser_checker_song_khi_vao_duoc_chat():
    page = _FakePage(
        "https://www.meta.ai/",
        '<html><body><textarea placeholder="Ask Meta AI">'
        "</textarea></body></html>",
    )
    checker, browser = _browser_checker_with(page)
    status, detail = checker.check(COOKIE_LIVE)
    assert status == "song", (status, detail)
    # Mở trang bằng domcontentloaded (không chờ load hết) để
    # mỗi lần kiểm tra nhanh hơn.
    assert page.goto_wait_until == "domcontentloaded", page.goto_wait_until
    # Cookie được gắn vào context (url meta.ai), rồi context
    # đóng lại sau mỗi lần kiểm tra.
    assert browser.contexts[0].cookies[0]["name"] == "hatch_sess"
    assert (
        browser.contexts[0].cookies[0]["url"]
        == "https://www.meta.ai/"
    )
    assert browser.contexts[0].closed
    assert page.closed
    checker.close()
    assert browser.closed


def test_browser_checker_chet_khi_bi_day_ve_dang_nhap():
    page = _FakePage(
        "https://www.meta.ai/login", "<html><body>Log in</body></html>"
    )
    checker, browser = _browser_checker_with(page)
    status, _detail = checker.check(COOKIE_LIVE)
    assert status == "chet", status
    checker.close()


def test_browser_checker_khong_ro_khi_trang_la():
    page = _FakePage(
        "https://www.meta.ai/", "<html><body>trang lạ</body></html>"
    )
    checker, browser = _browser_checker_with(page, timeout_ms=300)
    status, _detail = checker.check(COOKIE_LIVE)
    assert status == "khong-ro", status
    checker.close()


def test_browser_checker_cho_trang_on_dinh():
    # Trang đang điều hướng 2 lần đầu (content() ném lỗi),
    # lần 3 mới hiện ô chat -> vẫn kết luận sống.
    page = _FakePage(
        "https://www.meta.ai/",
        '<html><body><textarea placeholder="Ask Meta AI">'
        "</textarea></body></html>",
        fail_content_times=2,
    )
    checker, browser = _browser_checker_with(page)
    status, _detail = checker.check(COOKIE_LIVE)
    assert status == "song", status
    checker.close()


def test_browser_checker_chet_khi_trang_hien_sign_in():
    page = _FakePage(
        "https://www.meta.ai/",
        "<html><body><button>Sign in</button></body></html>",
    )
    checker, browser = _browser_checker_with(page)
    status, _detail = checker.check(COOKIE_LIVE)
    assert status == "chet", status
    checker.close()


def test_browser_checker_challenge_page():
    page = _FakePage(
        "https://www.meta.ai/",
        "<html><head><title>Just a moment...</title></head>"
        "<body>Verify you are human</body></html>",
    )
    checker, browser = _browser_checker_with(page, timeout_ms=300)
    status, detail = checker.check(COOKIE_LIVE)
    assert status == "khong-ro", (status, detail)
    assert "challenge" in detail, detail
    checker.close()


def test_browser_checker_khong_can_trinh_duyet_khi_thieu_hatch_sess():
    page = _FakePage("https://www.meta.ai/", "<html></html>")
    checker, browser = _browser_checker_with(page)
    status, _detail = checker.check("c_user=1001; sb=bbb222")
    assert status == "chet", status
    assert not browser.contexts, "không được mở context"
    checker.close()


def test_browser_checker_start_bao_loi_ro_rang():
    def boom():
        raise ImportError("No module named 'playwright'")

    checker = cookie_health.BrowserChecker(launcher=boom)
    try:
        checker.start()
        raise AssertionError("phải báo lỗi")
    except RuntimeError as error:
        assert "Playwright" in str(error), error
        assert "playwright install chromium" in str(error), error


class _FakeAPIResponse:
    def __init__(self, url, body):
        self.url = url
        self._body = body

    def text(self):
        return self._body


class _FakeAPIRequest:
    """Ngăn xếp yêu cầu giả: trả HTML cố định, ghi lại header."""

    def __init__(self, response=None, error=None):
        self._response = response
        self._error = error
        self.get_calls = []
        self.last_headers = None
        self.disposed = False

    def get(self, url, headers=None, timeout=None):
        self.get_calls.append(url)
        self.last_headers = headers or {}
        if self._error is not None:
            raise self._error
        return self._response

    def dispose(self):
        self.disposed = True


def _api_checker_with(url, body, error=None):
    request = _FakeAPIRequest(
        response=_FakeAPIResponse(url, body), error=error
    )
    checker = cookie_health.ApiChecker(launcher=lambda: request)
    checker.start()
    return checker, request


def test_api_checker_song_khi_html_co_ask_meta_ai():
    checker, request = _api_checker_with(
        "https://www.meta.ai/",
        '<html><body><textarea placeholder="Ask Meta AI">'
        "</textarea></body></html>",
    )
    status, detail = checker.check(COOKIE_LIVE)
    assert status == "song", (status, detail)
    assert "yêu cầu nhanh" in detail, detail
    # Cookie được gửi kèm trong header của yêu cầu nhanh.
    assert "hatch_sess=live123" in request.last_headers["Cookie"]
    checker.close()
    assert request.disposed


def test_api_checker_chet_khi_bi_day_ve_dang_nhap():
    checker, _request = _api_checker_with(
        "https://www.meta.ai/login", "<html><body>login</body></html>"
    )
    status, detail = checker.check(COOKIE_LIVE)
    assert status == "chet", (status, detail)
    checker.close()


def test_api_checker_chet_khi_html_hien_sign_in():
    checker, _request = _api_checker_with(
        "https://www.meta.ai/",
        "<html><body><button>Sign in to continue</button></body></html>",
    )
    status, detail = checker.check(COOKIE_LIVE)
    assert status == "chet", (status, detail)
    checker.close()


def test_api_checker_khong_ro_khi_html_la():
    checker, _request = _api_checker_with(
        "https://www.meta.ai/", "<html><body>hello</body></html>"
    )
    status, detail = checker.check(COOKIE_LIVE)
    assert status == "khong-ro", (status, detail)
    checker.close()


def test_api_checker_khong_ro_khi_loi_mang():
    checker, _request = _api_checker_with(
        "https://www.meta.ai/",
        "",
        error=RuntimeError("Connection reset"),
    )
    status, detail = checker.check(COOKIE_LIVE)
    assert status == "khong-ro", (status, detail)
    checker.close()


def test_api_checker_chet_khi_khong_co_hatch_sess():
    checker, request = _api_checker_with(
        "https://www.meta.ai/", "<html></html>"
    )
    status, _detail = checker.check("c_user=1001; sb=bbb222")
    assert status == "chet", status
    # Không có hatch_sess thì không gọi mạng làm gì.
    assert request.get_calls == []
    checker.close()


def test_check_and_cache_dung_trinh_duyet_khi_yeu_cau_nhanh_khong_ro():
    # Yêu cầu nhanh không nhận diện được HTML lạ -> qua trình
    # duyệt, kết quả trình duyệt được giữ.
    fast, _request = _api_checker_with(
        "https://www.meta.ai/", "<html><body>hello</body></html>"
    )
    browser = _FakeChecker(status="song")
    entry = cookie_health.check_and_cache(
        "a@fexpost.com", COOKIE_LIVE, checker=browser,
        fast_checker=fast,
    )
    assert entry["status"] == "song", entry
    assert browser.checked == [COOKIE_LIVE], browser.checked
    fast.close()


def test_check_and_cache_khong_dung_trinh_duyet_khi_yeu_cau_nhanh_ket_luan_duoc():
    # Yêu cầu nhanh đã kết luận sống -> trình duyệt không được
    # gọi (tiết kiệm một lần mở trang).
    fast, _request = _api_checker_with(
        "https://www.meta.ai/",
        '<html><body><textarea placeholder="Ask Meta AI">'
        "</textarea></body></html>",
    )
    browser = _FakeChecker(status="song")
    entry = cookie_health.check_and_cache(
        "a@fexpost.com", COOKIE_LIVE, checker=browser,
        fast_checker=fast,
    )
    assert entry["status"] == "song", entry
    assert browser.checked == [], browser.checked
    fast.close()


def test_check_job_chia_deu_cac_luong():
    emails = [f"u{i}@fexpost.com" for i in range(4)]
    for email in emails:
        storage._append_spreadsheet_row(
            email,
            "s",
            cookie=COOKIE_LIVE,
            cookie_taken_at=_iso(_now()),
        )
    created = []

    def factory():
        checker = _FakeChecker()
        created.append(checker)
        return checker

    web = cookie_web.CookieWeb()
    job = web._new_job(emails, "check")
    web._run_check_job(job, checker_factory=factory)
    assert job["done"] == 4, job
    assert "4 hoạt động" in job["result"], job
    # 1 probe + 4 luồng (4 mail, mỗi luồng 1 mail).
    assert len(created) == 5, len(created)
    workers = [c for c in created if c.checked]
    assert len(workers) == 4, created
    assert all(len(c.checked) == 1 for c in workers), created


def test_check_and_cache_khong_mat_entry_khi_ghi_song_song():
    import threading

    checker = _FakeChecker()
    emails = [f"u{i}@fexpost.com" for i in range(10)]
    threads = [
        threading.Thread(
            target=cookie_health.check_and_cache,
            args=(email, COOKIE_LIVE),
            kwargs={"checker": checker},
            daemon=True,
        )
        for email in emails
    ]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join()
    cache = cookie_health.load_health_cache()
    assert sorted(cache) == sorted(emails), cache
    assert all(
        entry["status"] == "song" for entry in cache.values()
    ), cache


def test_parse_cookie_string_giu_dau_bang_trong_gia_tri():
    assert cookie_health.parse_cookie_string(
        "sb=abc123==; hatch_sess=x=y; c_user=1001"
    ) == [
        ("sb", "abc123=="),
        ("hatch_sess", "x=y"),
        ("c_user", "1001"),
    ]


def test_clean_cookie_pairs_bo_gia_tri_loi_va_trung_ten():
    assert cookie_health.clean_cookie_pairs(
        [
            ("hatch_sess", "live123"),
            ("xau", "co dau cach"),
            ("xau2", "co,dau,phay"),
            ("xau3", 'co"nhay'),
            ("sb", "trung"),
            ("sb", "trung2"),
            ("", "rong"),
        ]
    ) == [("hatch_sess", "live123"), ("sb", "trung2")]


class _FailingCookiesContext(_FakeContext):
    def add_cookies(self, cookies):
        raise RuntimeError(
            "Protocol error (Storage.setCookies): "
            "Invalid cookie fields"
        )


class _FailingCookiesBrowser(_FakeBrowser):
    def new_context(self):
        ctx = _FailingCookiesContext(self.page)
        self.contexts.append(ctx)
        return ctx


def test_browser_checker_bo_qua_cookie_loi():
    page = _FakePage(
        "https://www.meta.ai/",
        '<html><body><textarea placeholder="Ask Meta AI">'
        "</textarea></body></html>",
    )
    browser = _FakeBrowser(page)
    checker = cookie_health.BrowserChecker(launcher=lambda: browser)
    checker.start()
    status, _detail = checker.check(
        "hatch_sess=live123; xau=co dau cach; "
        "c_user=1001; sb=trung; sb=trung2"
    )
    assert status == "song", status
    attached = {
        c["name"]: c["value"]
        for c in browser.contexts[0].cookies
    }
    assert "xau" not in attached, attached
    assert attached["sb"] == "trung2", attached
    assert all(
        c["url"] == "https://www.meta.ai/"
        for c in browser.contexts[0].cookies
    )
    checker.close()


def test_browser_checker_khong_chet_khi_add_cookies_loi():
    page = _FakePage("https://www.meta.ai/", "<html></html>")
    browser = _FailingCookiesBrowser(page)
    checker = cookie_health.BrowserChecker(launcher=lambda: browser)
    checker.start()
    status, detail = checker.check(COOKIE_LIVE)
    assert status == "khong-ro", (status, detail)
    assert "không gắn được cookie" in detail, detail
    assert browser.contexts[0].closed
    checker.close()


class _FlakyChecker(_FakeChecker):
    def check(self, cookie):
        if "RAISE" in cookie:
            raise RuntimeError("trình duyệt sập")
        return super().check(cookie)


class _MixedChecker(_FakeChecker):
    """Checker giả: chết hay sống tùy cookie — để test lượt
    kiểm tra tự đăng nhập lại đúng tài khoản đã chết."""

    def check(self, cookie):
        self.checked.append(cookie)
        if "chet" in cookie:
            return "chet", "bị đẩy về màn đăng nhập"
        return "song", "vào được giao diện chat"


def test_check_job_tu_dong_dang_nhap_lai_cookie_chet():
    storage._append_spreadsheet_row(
        "chet@fexpost.com",
        "s",
        cookie="AAA-chet1-BBB",
        cookie_taken_at=_iso(_now()),
    )
    storage._append_spreadsheet_row(
        "song@fexpost.com",
        "s",
        cookie="AAA-song1-BBB",
        cookie_taken_at=_iso(_now()),
    )
    goi_refresh = []

    def factory():
        return _MixedChecker()

    web = cookie_web.CookieWeb()
    web.start_refresh = lambda emails: goi_refresh.append(
        list(emails)
    ) or "job-99"
    job = web._new_job(
        ["chet@fexpost.com", "song@fexpost.com"], "check"
    )
    web._run_check_job(job, checker_factory=factory)
    assert "1 die" in job["result"], job
    assert goi_refresh == [["chet@fexpost.com"]], goi_refresh
    assert any(
        "tự động đăng nhập lại (job-99)" in line
        for line in job["log"]
    ), job["log"]


def test_check_job_khong_dang_nhap_lai_khi_khong_co_cookie_chet():
    storage._append_spreadsheet_row(
        "song@fexpost.com",
        "s",
        cookie="AAA-song1-BBB",
        cookie_taken_at=_iso(_now()),
    )
    goi_refresh = []

    def factory():
        return _FakeChecker("song")

    web = cookie_web.CookieWeb()
    web.start_refresh = lambda emails: goi_refresh.append(
        list(emails)
    ) or "job-99"
    job = web._new_job(["song@fexpost.com"], "check")
    web._run_check_job(job, checker_factory=factory)
    assert goi_refresh == [], goi_refresh
    assert "1 hoạt động" in job["result"], job


def test_tu_dong_dang_nhap_lai_bo_qua_khi_dang_co_luot_refresh():
    web = cookie_web.CookieWeb()
    web._new_job(["x@fexpost.com"], "refresh")
    goi_refresh = []
    web.start_refresh = lambda emails: goi_refresh.append(
        list(emails)
    ) or "job-99"
    job = web._new_job(["chet@fexpost.com"], "check")
    ket_qua = web._tu_dong_dang_nhap_lai_cookie_chet(
        job, ["chet@fexpost.com"]
    )
    assert ket_qua is None, ket_qua
    assert goi_refresh == [], goi_refresh
    assert any(
        "đã có lượt đăng nhập lại đang chạy" in line
        for line in job["log"]
    ), job["log"]


def test_tu_dong_dang_nhap_lai_ghi_log_khi_khong_khoi_dong_duoc():
    web = cookie_web.CookieWeb()

    def _loi(emails):
        raise RuntimeError("Tool đang chạy một lượt khác")

    web.start_refresh = _loi
    job = web._new_job(["chet@fexpost.com"], "check")
    ket_qua = web._tu_dong_dang_nhap_lai_cookie_chet(
        job, ["chet@fexpost.com"]
    )
    assert ket_qua is None, ket_qua
    assert any(
        "không tự đăng nhập lại được" in line
        for line in job["log"]
    ), job["log"]


def test_check_uu_tien_qua_han_truoc():
    # Ưu tiên quá hạn kiểm tra trước (yêu cầu 2026-11-17):
    # diện stale lấy theo thứ tự của kho — quá hạn lên trước,
    # còn hạn xếp sau. Còn hạn mà sống thì thôi, chết thì lượt
    # kiểm tra tự đăng nhập lại (test riêng ở trên).
    storage._append_spreadsheet_row(
        "conhan@fexpost.com",
        "s",
        cookie="AAA-song1-BBB",
        cookie_taken_at=_iso(_now()),
    )
    storage._append_spreadsheet_row(
        "quahan@fexpost.com",
        "s",
        cookie="AAA-chet1-BBB",
        cookie_taken_at=_iso(_now() - timedelta(days=9)),
    )
    cookies = cookie_vault.list_cookies()
    emails = cookie_web.select_check_emails(cookies, "stale")
    assert emails == [
        "quahan@fexpost.com",
        "conhan@fexpost.com",
    ], emails


def test_kiem_tra_mot_mail_tra_ve_none_khi_ket():
    done = threading.Event()

    class _Ket:
        def check(self, cookie):
            time.sleep(1.2)
            done.set()
            return "song", "xong muộn"

    started = time.monotonic()
    entry = cookie_web._kiem_tra_mot_mail(
        "a@fexpost.com", COOKIE_LIVE, _Ket(), None,
        timeout_seconds=0.5,
    )
    elapsed = time.monotonic() - started
    assert entry is None, entry
    assert elapsed < 3, elapsed
    # Chờ luồng kẹt chạy xong để nó không ghi cache sang test sau.
    assert done.wait(timeout=3), "luồng kẹt không chạy xong"
    time.sleep(0.5)


def test_check_slice_dung_cung_luong_da_mo_trinh_duyet():
    # Playwright sync gắn greenlet vào luồng đã start(). check()
    # và close() phải chạy trên đúng luồng đó — gọi từ luồng chờ
    # timeout sẽ ném 'Cannot switch to a different thread'.
    seen = {}

    class _GhiLuong:
        def start(self):
            seen["start"] = threading.get_ident()

        def check(self, cookie):
            seen["check"] = threading.get_ident()
            return "song", "vào được giao diện chat"

        def close(self):
            seen["close"] = threading.get_ident()

    web = cookie_web.CookieWeb()
    job = web._new_job(["a@fexpost.com"], "check")
    counts = {"song": 0, "chet": 0, "khong-ro": 0}
    web._check_slice(
        job,
        ["a@fexpost.com"],
        {"a@fexpost.com": COOKIE_LIVE},
        lambda: _GhiLuong(),
        counts,
        [],
        threading.Lock(),
        fast_factory=None,
        timeout_seconds=5,
    )
    assert counts["song"] == 1, counts
    assert seen["start"] == seen["check"] == seen["close"], seen
    assert seen["start"] != threading.get_ident(), seen


def test_check_slice_thay_cap_moi_khi_mot_mail_ket():
    # Mail đầu kẹt (check ngủ 1.2 giây), mail sau vẫn được kiểm
    # tra bằng cặp mới — job không đứng cứng (lỗi 2026-11-17:
    # kẹt ở 96/2096).
    calls = {"n": 0}
    done = threading.Event()

    class _KetRoiSong:
        def __init__(self):
            calls["n"] += 1
            self.so = calls["n"]

        def start(self):
            pass

        def check(self, cookie):
            if self.so == 1:
                time.sleep(1.2)
                done.set()
                return "song", "xong muộn"
            return "song", "vào được giao diện chat"

        def close(self):
            pass

    def factory():
        return _KetRoiSong()

    web = cookie_web.CookieWeb()
    job = web._new_job(
        ["ket@fexpost.com", "song@fexpost.com"], "check"
    )
    counts = {"song": 0, "chet": 0, "khong-ro": 0}
    dead = []
    lock = threading.Lock()
    cookies = {
        "ket@fexpost.com": COOKIE_LIVE,
        "song@fexpost.com": COOKIE_LIVE,
    }
    started = time.monotonic()
    web._check_slice(
        job, ["ket@fexpost.com", "song@fexpost.com"], cookies,
        factory, counts, dead, lock, fast_factory=None,
        timeout_seconds=0.5,
    )
    elapsed = time.monotonic() - started
    assert elapsed < 5, elapsed
    assert job["done"] == 2, job
    assert counts == {"song": 1, "chet": 0, "khong-ro": 1}, counts
    assert any("kẹt quá" in line for line in job["log"]), job
    assert calls["n"] == 2, calls
    assert done.wait(timeout=3), "luồng kẹt không chạy xong"
    time.sleep(0.5)


def test_check_job_tu_kiem_tra_lai_mail_chua_ro():
    # Mail chưa rõ được kiểm tra lại một lần khi lượt xong.
    # Lượt kiểm tra lại không tự xếp thêm lượt nữa.
    storage._append_spreadsheet_row(
        "chua@fexpost.com",
        "s",
        cookie="hatch_sess=CHUA",
        cookie_taken_at=_iso(_now()),
    )
    storage._append_spreadsheet_row(
        "song@fexpost.com",
        "s",
        cookie=COOKIE_LIVE,
        cookie_taken_at=_iso(_now()),
    )

    class _Checker:
        def start(self):
            pass

        def smoke_test(self):
            pass

        def check(self, cookie):
            if "CHUA" in cookie:
                return "khong-ro", "trang trả về không nhận diện được"
            return "song", "vào được giao diện chat"

        def close(self):
            pass

    class _NoFast:
        def start(self):
            raise RuntimeError("không dùng yêu cầu nhanh")

        def close(self):
            pass

    web = cookie_web.CookieWeb()
    job = web._new_job(
        ["chua@fexpost.com", "song@fexpost.com"], "check"
    )
    job["kiem_tra_lai"] = True
    web._run_check_job(
        job, checker_factory=_Checker, fast_factory=_NoFast
    )
    extra = [
        thread for thread in threading.enumerate()
        if thread.name.startswith("cookie-check-")
        and thread is not threading.current_thread()
    ]
    for thread in extra:
        thread.join(timeout=5)
        assert not thread.is_alive(), thread.name
    checks = [
        item for item in web.jobs.values() if item["kind"] == "check"
    ]
    assert len(checks) == 2, checks
    retry = next(item for item in checks if item["id"] != job["id"])
    assert retry["emails"] == ["chua@fexpost.com"], retry
    assert retry.get("kiem_tra_lai") is False
    assert "1 chưa rõ" in job["result"], job
    assert any("tự kiểm tra lại" in line for line in job["log"]), job
    assert "1 chưa rõ" in retry["result"], retry


def test_check_slice_khong_chet_khi_mot_mail_loi():
    storage._append_spreadsheet_row(
        "loi@fexpost.com",
        "s",
        cookie="hatch_sess=RAISE123",
        cookie_taken_at=_iso(_now()),
    )
    for email in ("a@fexpost.com", "b@fexpost.com"):
        storage._append_spreadsheet_row(
            email,
            "s",
            cookie=COOKIE_LIVE,
            cookie_taken_at=_iso(_now()),
        )

    def factory():
        return _FlakyChecker()

    web = cookie_web.CookieWeb()
    job = web._new_job(
        ["loi@fexpost.com", "a@fexpost.com", "b@fexpost.com"],
        "check",
    )
    web._run_check_job(job, checker_factory=factory)
    assert job["done"] == 3, job
    assert "2 hoạt động" in job["result"], job
    assert "1 chưa rõ" in job["result"], job
    assert any(
        line.startswith("✘ loi@fexpost.com") for line in job["log"]
    ), job


def test_coordinator_dang_nhap_lai_an_trinh_duyet():
    coord = cookie_web._coordinator_dang_nhap_lai("-2")
    assert coord._headless is True
    assert coord._lane_local.suffix == "-2"
    command = [sys.executable, "browser_helper.py", "a@x.com", "--auto"]
    # Cùng quy ước coordinator._run_helper: suffix "-2" thành
    # đối số "2", và có --headless.
    suffix = coord._lane_local.suffix
    if suffix:
        command.append(suffix.lstrip("-"))
    if coord._headless:
        command.append("--headless")
    assert command[-2:] == ["2", "--headless"]


def test_refresh_job_chay_nhieu_luong():
    web = cookie_web.CookieWeb()
    emails = [f"r{i}@fexpost.com" for i in range(4)]
    job = web._new_job(emails, "refresh")
    seen = []
    seen_lock = threading.Lock()

    def fake_refresh(job, email, suffix="", log_lock=None):
        time.sleep(0.05)
        with seen_lock:
            seen.append((email, suffix, threading.get_ident()))

    web._refresh_one = fake_refresh
    web._run_refresh_job(job)
    assert job["done"] == 4, job
    assert "4 thành công" in job["result"], job
    suffixes = {suffix for _email, suffix, _ident in seen}
    assert suffixes == {"-11", "-12", "-13", "-14"}, suffixes
    idents = {ident for _email, _suffix, ident in seen}
    assert len(idents) == 4, idents


def test_probe_dung_ngay_khi_trinh_duyet_khong_mo_duoc_trang():
    class _NoPageChecker(_FakeChecker):
        def smoke_test(self):
            raise RuntimeError("không tạo được trang")

    def factory():
        return _NoPageChecker()

    web = cookie_web.CookieWeb()
    job = web._new_job(["a@fexpost.com"], "check")
    web._run_check_job(job, checker_factory=factory)
    assert job["done"] == 0, job
    assert "Không kiểm tra được" in job["result"], job
    assert "không tạo được trang" in job["result"], job


def test_dang_nhap_meta_thanh_cong_khong_kiem_tra():
    # Vào được Muse từ Meta là còn sống. Vừa vào thì không kiểm tra.
    # Vào quá 7 ngày thì kiểm lại. Mail không đi từ Meta thì vẫn kiểm.
    now = _iso(_now())
    old = _iso(_now() - timedelta(days=8))
    storage._append_spreadsheet_row(
        "moi@fexpost.com", "pw",
        cookie="hatch_sess=moi", cookie_taken_at=now,
    )
    storage._append_spreadsheet_row(
        "cu@fexpost.com", "pw",
        cookie="hatch_sess=cu", cookie_taken_at=old,
    )
    storage._append_spreadsheet_row(
        "khac@fexpost.com", "pw",
        cookie="hatch_sess=khac", cookie_taken_at=now,
    )
    from openpyxl import Workbook

    book = Workbook()
    sheet = book.active
    sheet.append(list(storage.META_SPREADSHEET_HEADERS))
    sheet.append(["t", "moi@fexpost.com", "pw", "u", "tempmail.plus", now, ""])
    sheet.append(["t", "cu@fexpost.com", "pw", "u", "tempmail.plus", old, ""])
    sheet.append([
        "t", "dang@fexpost.com", "pw", "u", "tempmail.plus",
        "ĐANG ĐĂNG NHẬP " + now, "",
    ])
    path = TMP / "meta_live.xlsx"
    book.save(path)
    da_vao = storage._meta_da_dang_nhap_muse(path)
    assert set(da_vao) == {"moi@fexpost.com", "cu@fexpost.com"}, da_vao
    assert cookie_health.nang_dang_nhap_meta_thanh_song(
        storage._read_cookie_rows(), da_vao
    ) is True
    listed = {
        item["email"]: item for item in cookie_vault.list_cookies()
    }
    assert listed["moi@fexpost.com"]["health"] == "song", listed
    assert listed["cu@fexpost.com"]["health"] == "song", listed
    assert listed["khac@fexpost.com"]["health"] != "song", listed
    due = set(cookie_web.select_mail_vong_nen(list(listed.values())))
    assert "moi@fexpost.com" not in due, due
    assert "cu@fexpost.com" in due, due
    assert "khac@fexpost.com" in due, due
    later = _iso(_now() + timedelta(minutes=5))
    cookie_health.save_health_cache({
        "moi@fexpost.com": {
            "status": "chet",
            "checked_at": later,
            "detail": "die sau khi đã vào",
        }
    })
    cookie_health.nang_dang_nhap_meta_thanh_song(
        storage._read_cookie_rows(), {"moi@fexpost.com": now}
    )
    again = {
        item["email"]: item for item in cookie_vault.list_cookies()
    }
    assert again["moi@fexpost.com"]["health"] == "chet", again


def test_dang_nhap_lai_het_cuu_khong_chon_nua():
    assert cookie_web.con_cuu_duoc(
        "Chờ mãi không thấy form nhập email của Muse.ai."
    )
    assert cookie_web.con_cuu_duoc(
        "Proxy này đẩy sang Facebook. Đã bỏ proxy."
    )
    assert cookie_web.con_cuu_duoc("")
    assert not cookie_web.con_cuu_duoc(
        "Không tìm thấy mã xác minh mới trong hộp thư "
        "tempmail.plus sau 180 giây."
    )
    assert not cookie_web.con_cuu_duoc(
        "Đăng nhập xong Muse đòi nhập thẻ — tab đăng nhập "
        "lại không nộp thẻ."
    )
    assert not cookie_web.con_cuu_duoc(
        "Màn liên kết Meta không khớp mail cần đăng nhập "
        "— không bấm Continue, không lấy cookie."
    )
    now = _iso(_now())
    storage._append_spreadsheet_row(
        "die@fexpost.com", "pw", cookie="hatch_sess=d",
        cookie_taken_at=now,
    )
    storage._append_spreadsheet_row(
        "het@fexpost.com", "pw", cookie="hatch_sess=h",
        cookie_taken_at=now,
    )
    cookie_health.save_health_cache({
        "die@fexpost.com": {
            "status": "chet", "checked_at": now, "detail": "die",
        }
    })
    cookie_health.ghi_het_cuu(
        "het@fexpost.com", "Không tìm thấy mã xác minh mới"
    )
    listed = cookie_vault.list_cookies()
    assert cookie_web.select_refresh_emails(listed, "dead") == [
        "die@fexpost.com"
    ]
    assert "het@fexpost.com" not in cookie_web.select_refresh_emails(
        listed, "all"
    )
    by_email = {item["email"]: item for item in listed}
    assert by_email["het@fexpost.com"]["health"] == "het-cuu"
    stats = cookie_vault.cookie_stats(listed)
    assert stats["het-cuu"] == 1 and stats["chet"] == 1, stats


def test_dang_nhap_lai_thanh_cong_ghi_con_song():
    # Cookie cũ đã Die. Đăng nhập lại lấy cookie mới thì Còn sống
    # phải tăng, không giữ Die của cookie cũ.
    taken_cu = _iso(_now() - timedelta(days=2))
    storage._append_spreadsheet_row(
        "die@fexpost.com",
        "pw",
        cookie="hatch_sess=cu",
        cookie_taken_at=taken_cu,
    )
    cookie_health.save_health_cache({
        "die@fexpost.com": {
            "status": "chet",
            "checked_at": _iso(_now() - timedelta(hours=1)),
            "detail": "bị đẩy về màn đăng nhập",
        }
    })
    storage._update_cookie_row(
        "die@fexpost.com",
        "hatch_sess=moi",
        _iso(_now()),
    )
    assert cookie_health.nang_cookie_moi_thanh_song(
        storage._read_cookie_rows()
    ) is True
    listed = {
        item["email"]: item for item in cookie_vault.list_cookies()
    }
    assert listed["die@fexpost.com"]["health"] == "song", listed
    stats = cookie_vault.cookie_stats(list(listed.values()))
    assert stats["hoat-dong"] == 1 and stats["chet"] == 0, stats

    cookie_health.save_health_cache({
        "die@fexpost.com": {
            "status": "chet",
            "checked_at": _iso(_now() - timedelta(minutes=5)),
            "detail": "die lại",
        }
    })
    storage._save_relogin_to_accounts(
        "die@fexpost.com", "hatch_sess=nua", "pw"
    )
    again = {
        item["email"]: item for item in cookie_vault.list_cookies()
    }
    assert again["die@fexpost.com"]["health"] == "song", again


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
    test_chon_nhieu_the_va_dang_ky_vao_quan_ly,
    test_register_page_and_console_refuse_empty_proxy,
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
    test_vong_nen_tu_bat_dau_va_khong_xep_chong,
    test_start_check_all_khong_xep_chong_job,
    test_check_job_reports_done_count,
    test_check_job_dung_ngay_khi_khong_mo_duoc_trinh_duyet,
    test_check_job_chia_deu_cac_luong,
    test_check_job_tu_dong_dang_nhap_lai_cookie_chet,
    test_check_job_khong_dang_nhap_lai_khi_khong_co_cookie_chet,
    test_tu_dong_dang_nhap_lai_bo_qua_khi_dang_co_luot_refresh,
    test_tu_dong_dang_nhap_lai_ghi_log_khi_khong_khoi_dong_duoc,
    test_check_uu_tien_qua_han_truoc,
    test_check_and_cache_khong_mat_entry_khi_ghi_song_song,
    test_check_job_tu_kiem_tra_lai_mail_chua_ro,
    test_check_slice_khong_chet_khi_mot_mail_loi,
    test_coordinator_dang_nhap_lai_an_trinh_duyet,
    test_refresh_job_chay_nhieu_luong,
    test_dang_nhap_meta_thanh_cong_khong_kiem_tra,
    test_dang_nhap_lai_het_cuu_khong_chon_nua,
    test_dang_nhap_lai_thanh_cong_ghi_con_song,
    test_probe_dung_ngay_khi_trinh_duyet_khong_mo_duoc_trang,
    test_parse_cookie_string,
    test_parse_cookie_string_giu_dau_bang_trong_gia_tri,
    test_clean_cookie_pairs_bo_gia_tri_loi_va_trung_ten,
    test_browser_checker_bo_qua_cookie_loi,
    test_browser_checker_khong_chet_khi_add_cookies_loi,
    test_browser_checker_chan_anh_font_video,
    test_browser_checker_song_khi_vao_duoc_chat,
    test_browser_checker_chet_khi_bi_day_ve_dang_nhap,
    test_browser_checker_khong_ro_khi_trang_la,
    test_browser_checker_cho_trang_on_dinh,
    test_browser_checker_chet_khi_trang_hien_sign_in,
    test_browser_checker_challenge_page,
    test_browser_checker_khong_can_trinh_duyet_khi_thieu_hatch_sess,
    test_browser_checker_start_bao_loi_ro_rang,
    test_api_checker_song_khi_html_co_ask_meta_ai,
    test_api_checker_chet_khi_bi_day_ve_dang_nhap,
    test_api_checker_chet_khi_html_hien_sign_in,
    test_api_checker_khong_ro_khi_html_la,
    test_api_checker_khong_ro_khi_loi_mang,
    test_api_checker_chet_khi_khong_co_hatch_sess,
    test_check_and_cache_dung_trinh_duyet_khi_yeu_cau_nhanh_khong_ro,
    test_check_and_cache_khong_dung_trinh_duyet_khi_yeu_cau_nhanh_ket_luan_duoc,
    test_kiem_tra_mot_mail_tra_ve_none_khi_ket,
    test_check_slice_dung_cung_luong_da_mo_trinh_duyet,
    test_check_slice_thay_cap_moi_khi_mot_mail_ket,
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
