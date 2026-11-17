"""Web quản lý kho cookie — chạy trên máy Windows, mở bằng trình duyệt.

Chạy:  py main.py --cookie-web
  (hoặc bấm nút "Quản lý cookie (web)" ở tab Muse trong tool)

Server chỉ nghe trên 127.0.0.1 (máy này) — không mở ra mạng ngoài.
Kho tài khoản duy nhất là SQLite (accounts.db) — một dòng một
tài khoản, email là khoá duy nhất (xem account_store.py).
accounts.xlsx chỉ còn là đường vào (nút "Nhập từ Excel", nhập
một lần từ file cũ) và đường ra (nút "Xuất ra Excel"). Không đổi
các bước tạo tài khoản hay đăng nhập.

Đăng nhập lại chạy ĐÚNG luồng tab Đăng nhập lại của tool
(coordinator._run_relogin_lane): không proxy (DIRECT), không
nộp thẻ, không mã mời — chỉ lấy OTP từ hộp thư của đúng mail
đó rồi lấy cookie mới qua helper --relogin. Cookie mới ghi đè
vào dòng hiện có trong kho SQLite (và đồng bộ sang file xuất
muse_cookies.xlsx của tab Đăng nhập lại nếu đã đặt đường dẫn).

Mỗi cookie sống 7 ngày kể từ lúc lấy. Dashboard còn kiểm tra
cookie thật sự hoạt động hay đã die (xem cookie_health.py) —
không chỉ "có cookie là được".

Chỉ dùng thư viện chuẩn của Python — không cần cài thêm gì.
"""

import json
import queue
import threading
import time
import webbrowser
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

import mail_providers
import paths
import storage
import cookie_health
import cookie_vault
from coordinator import Coordinator


DEFAULT_PORT = 8765
MAX_LOG_LINES = 200
MAX_HISTORY_SHOWN = 100

_SERVER = None
_SERVER_LOCK = threading.Lock()


def _now_iso():
    return datetime.now().astimezone().isoformat(timespec="seconds")


def _relogin_source_path():
    """File Excel nguồn để NHẬP vào kho (mặc định là
    accounts.xlsx). Kho sống là SQLite — file này chỉ dùng khi
    bấm "Nhập từ Excel"."""
    path = storage._load_relogin_source_path()
    if path:
        return path
    return str(storage.SPREADSHEET_FILE)


def _relogin_output_path():
    """File xuất cookie của tab Đăng nhập lại ("" khi chưa đặt)."""
    return storage._load_relogin_output_path()


def select_refresh_emails(cookies, scope):
    """Chọn mail đăng nhập lại theo diện (dùng cho /api/refresh-all).

    - "expired": chỉ đã hết hạn theo mốc 7 ngày.
    - "dead": chỉ cookie đã kiểm tra thật trên meta.ai là chết.
    - "all": tất cả.
    - "expiring" (mặc định): sắp hết + đã hết hạn.
    """
    if scope == "expired":
        return [
            item["email"]
            for item in cookies
            if item["status"] == cookie_vault.STATUS_EXPIRED
        ]
    if scope == "dead":
        return [
            item["email"]
            for item in cookies
            if item.get("health") == cookie_health.STATUS_DEAD
        ]
    if scope == "all":
        return [item["email"] for item in cookies]
    return [
        item["email"]
        for item in cookies
        if item["status"]
        in (
            cookie_vault.STATUS_EXPIRING,
            cookie_vault.STATUS_EXPIRED,
        )
    ]


CHECK_FRESH_HOURS = 24


def _checked_recently(checked_at, hours=CHECK_FRESH_HOURS):
    """True khi cookie đã kiểm tra trong `hours` giờ qua.

    Mốc kiểm tra hỏng/không đọc được thì tính là chưa kiểm tra
    (phải kiểm tra lại) chứ không được tính là mới.
    """
    if not checked_at:
        return False
    try:
        when = datetime.fromisoformat(str(checked_at))
    except ValueError:
        return False
    if when.tzinfo is None:
        when = when.replace(tzinfo=timezone.utc)
    return (
        datetime.now(timezone.utc) - when
    ).total_seconds() < hours * 3600


def select_check_emails(cookies, scope):
    """Chọn mail kiểm tra sống/chết theo diện (dùng cho /api/check-all).

    - "stale": chưa kiểm tra bao giờ hoặc đã kiểm tra quá 24 giờ
      trước — diện tự kiểm tra khi mở trang: lần đầu mở thì kiểm
      tra hết, lần sau chỉ kiểm tra lại phần đã cũ, không spam
      meta.ai mỗi lần F5.
    - "expiring": có cookie và sắp hết/hết hạn theo mốc 7 ngày.
    - "all": tất cả tài khoản có cookie.
    """
    if scope == "stale":
        return [
            item["email"]
            for item in cookies
            if item["has_cookie"]
            and not _checked_recently(item.get("health_checked_at"))
        ]
    if scope == "expiring":
        return [
            item["email"]
            for item in cookies
            if item["has_cookie"]
            and item["status"]
            in (
                cookie_vault.STATUS_EXPIRING,
                cookie_vault.STATUS_EXPIRED,
            )
        ]
    return [item["email"] for item in cookies if item["has_cookie"]]


class CookieWeb:
    """Trạng thái của web quản lý cookie: jobs + khoá chạy."""

    def __init__(self, desktop_coordinator=None):
        # Coordinator của cửa sổ tkinter (nếu web được mở từ nút
        # trong tool): đăng nhập lại bị từ chối khi tool đang chạy
        # một lượt khác, để hai luồng không giành mailbox.txt.
        self.desktop_coordinator = desktop_coordinator
        self.jobs = {}
        self.jobs_lock = threading.Lock()
        self.refresh_lock = threading.Lock()
        self._job_counter = 0

    # --- Đọc kho ------------------------------------------------

    def list_cookies(self):
        cookies = cookie_vault.list_cookies()
        return {
            "cookies": cookies,
            "stats": cookie_vault.cookie_stats(cookies),
            "server_time": _now_iso(),
        }

    def full_cookie(self, email):
        """Cookie đầy đủ của một email (chỉ để nút Sao chép dùng).

        Không bao giờ render trong bảng danh sách.
        """
        target = (email or "").strip().lower()
        for row in storage._read_cookie_rows():
            if row["email"].strip().lower() == target:
                return row.get("cookie") or ""
        return ""

    def import_excel(self, path=None):
        """Nhập tài khoản từ file Excel vào kho SQLite."""
        return storage.import_accounts_from_excel(path)

    def export_excel_bytes(self):
        """Toàn bộ kho -> bytes file Excel (để tải về)."""
        import account_store

        return account_store.export_to_excel_bytes(
            storage.ACCOUNTS_DB_FILE
        )

    def delete_cookie(self, email):
        """Xoá cookie của một tài khoản (giữ lại dòng + mã mời)."""
        if storage._clear_cookie_row(email):
            storage._append_cookie_history("xoa", email, "xoá cookie")
            return True, "Đã xoá cookie."
        return False, "Không tìm thấy tài khoản này trong kho."

    def history(self):
        return storage._read_cookie_history(MAX_HISTORY_SHOWN)

    def status(self):
        cookies = cookie_vault.list_cookies()
        return {
            "db_path": str(storage.ACCOUNTS_DB_FILE),
            "total": len(cookies),
            "source_path": _relogin_source_path(),
            "output_path": _relogin_output_path(),
            "vault_count": sum(
                1 for item in cookies if item["has_cookie"]
            ),
            "running_jobs": sum(
                1
                for job in self.jobs.values()
                if job["status"] in ("queued", "running")
            ),
            "migrating": storage._accounts_migration_running(),
            "last_import": storage._accounts_migration_result(),
        }

    # --- Jobs ----------------------------------------------------

    def _new_job(self, emails, kind):
        with self.jobs_lock:
            self._job_counter += 1
            job_id = f"job-{self._job_counter}"
            job = {
                "id": job_id,
                "kind": kind,
                "emails": list(emails),
                "status": "queued",
                "result": "",
                "log": [],
                "done": 0,
                "created_at": _now_iso(),
            }
            self.jobs[job_id] = job
        return job

    def start_refresh(self, emails):
        """Tạo một job đăng nhập lại cho các email (chạy nền)."""
        if self.desktop_coordinator is not None and getattr(
            self.desktop_coordinator, "running", False
        ):
            raise RuntimeError(
                "Tool đang chạy một lượt khác — đợi lượt đó xong rồi "
                "đăng nhập lại."
            )
        job = self._new_job(emails, "refresh")
        thread = threading.Thread(
            target=self._run_refresh_job,
            args=(job,),
            name=f"cookie-refresh-{job['id']}",
            daemon=True,
        )
        thread.start()
        return job["id"]

    def start_check_all(self, emails):
        """Tạo một job kiểm tra sống/chết cho các email (chạy nền).

        Đang có job kiểm tra chạy thì trả về job đó thay vì xếp
        job mới — mở trang/F5 nhiều lần không được xếp chồng
        kiểm tra lên nhau.
        """
        with self.jobs_lock:
            for job in self.jobs.values():
                if job["kind"] == "check" and job["status"] in (
                    "queued",
                    "running",
                ):
                    return job["id"]
        job = self._new_job(emails, "check")
        thread = threading.Thread(
            target=self._run_check_job,
            args=(job,),
            name=f"cookie-check-{job['id']}",
            daemon=True,
        )
        thread.start()
        return job["id"]

    def list_jobs(self):
        with self.jobs_lock:
            jobs = sorted(
                self.jobs.values(),
                key=lambda job: job["created_at"],
                reverse=True,
            )
        return [dict(job, log=list(job["log"])) for job in jobs[:20]]

    def _run_refresh_job(self, job):
        """Một job: đăng nhập lại lần lượt từng email."""
        # Một job đăng nhập lại tại một thời điểm (mailbox.txt và
        # proxy.txt là tài nguyên chung của tool).
        with self.refresh_lock:
            job["status"] = "running"
            done = 0
            failed = 0
            for email in job["emails"]:
                try:
                    self._refresh_one(job, email)
                    done += 1
                except Exception as error:
                    failed += 1
                    job["log"].append(f"✘ {email}: {error}")
                    job["log"] = job["log"][-MAX_LOG_LINES:]
                job["done"] = done + failed
            job["status"] = "done"
            job["result"] = (
                f"Xong: {done} thành công, {failed} thất bại "
                f"(trong {len(job['emails'])} tài khoản)."
            )

    def _run_check_job(self, job, checker=None):
        """Một job: kiểm tra sống/chết lần lượt từng email.

        Kiểm tra bằng trình duyệt headless (BrowserChecker — một
        browser dùng chung cho cả job). Không mở được trình duyệt
        thì dừng ngay, không chạy cả nghìn lần vô ích.
        """
        own_checker = checker is None
        if own_checker:
            checker = cookie_health.BrowserChecker()
            try:
                checker.start()
            except RuntimeError as error:
                job["status"] = "done"
                job["result"] = f"Không kiểm tra được: {error}"
                job["log"].append(str(error))
                job["log"] = job["log"][-MAX_LOG_LINES:]
                return
        try:
            job["status"] = "running"
            alive = 0
            dead = 0
            unknown = 0
            for email in job["emails"]:
                cookie = self.full_cookie(email)
                if not cookie:
                    job["log"].append(f"— {email}: chưa có cookie.")
                    unknown += 1
                    job["done"] = alive + dead + unknown
                    continue
                entry = cookie_health.check_and_cache(
                    email, cookie, checker=checker
                )
                status = entry["status"]
                if status == cookie_health.STATUS_ALIVE:
                    alive += 1
                elif status == cookie_health.STATUS_DEAD:
                    dead += 1
                else:
                    unknown += 1
                job["done"] = alive + dead + unknown
                job["log"].append(
                    f"{email}: {cookie_health.STATUS_LABELS[status]} "
                    f"({entry['detail']})"
                )
                job["log"] = job["log"][-MAX_LOG_LINES:]
            job["status"] = "done"
            job["result"] = (
                f"Xong: {alive} hoạt động, {dead} die, {unknown} chưa rõ "
                f"(trong {len(job['emails'])} tài khoản)."
            )
        finally:
            if own_checker:
                checker.close()

    # --- Đăng nhập lại một tài khoản ------------------------------

    def _refresh_one(self, job, email):
        """Đăng nhập lại MỘT tài khoản để lấy cookie mới.

        Chạy ĐÚNG luồng tab Đăng nhập lại: không proxy (ghi
        DIRECT), không nộp thẻ, không mã mời — helper chạy với
        --relogin, OTP lấy từ hộp thư của đúng mail đó. Khác tab
        ở chỗ đích ghi: cookie mới ghi đè vào dòng hiện có trong
        kho SQLite (cột Cookie + cột Thời gian lấy cookie) để
        quản lý một chỗ, đồng thời đồng bộ sang file xuất
        muse_cookies.xlsx của tab (nếu đã đặt đường dẫn) để tab
        không lấy lại mail này.
        """
        log = job["log"]

        def say(message):
            log.append(message)
            if len(log) > MAX_LOG_LINES:
                del log[: -MAX_LOG_LINES]

        rows = storage._relogin_accounts()
        row = next(
            (
                item
                for item in rows
                if item["email"].strip().lower()
                == (email or "").strip().lower()
            ),
            None,
        )
        if row is None:
            raise RuntimeError(
                f"không tìm thấy mail {email} trong kho — đăng nhập "
                "lại cần mail + bí mật hộp thư của tài khoản đó."
            )
        if row["provider"] == "mail.tm" and not row["secret"]:
            raise RuntimeError(
                f"{email} thiếu mật khẩu hộp thư — không đọc được "
                "OTP, bỏ qua."
            )
        coord = Coordinator()
        stop_drain = threading.Event()

        def drain_events():
            while not stop_drain.is_set() or not coord.events.empty():
                try:
                    kind, message = coord.events.get(timeout=0.3)
                except queue.Empty:
                    continue
                if kind == "log":
                    say(str(message))

        drain_thread = threading.Thread(
            target=drain_events, daemon=True
        )
        drain_thread.start()
        try:
            say(
                f"Bắt đầu đăng nhập lại {email} ({row['provider']}) "
                "— không proxy, không thẻ."
            )
            try:
                storage._write_active_proxy("DIRECT", "")
            except OSError as error:
                say(f"Không ghi được proxy.txt: {error}")
            try:
                storage._write_mailbox(
                    {
                        "provider": row["provider"],
                        "email": row["email"],
                        "secret": row["secret"],
                    },
                    "",
                )
            except OSError as error:
                raise RuntimeError(
                    f"không ghi được hộp thư cho {email}: {error}"
                ) from error
            helper = paths.SRC_DIR / "browser_helper.py"
            (
                helper_ok,
                _burned,
                _waitlisted,
                _load_failed,
                cookie,
                _invite,
                _card_skipped,
                _facebook,
            ) = coord._run_helper(
                helper, email, "Web", 1, relogin=True
            )
            if not helper_ok or not cookie:
                raise RuntimeError(
                    "đăng nhập lại thất bại — xem log phía trên."
                )
            now = _now_iso()
            updated = storage._update_cookie_row(email, cookie, now)
            if not updated:
                storage._append_spreadsheet_row(
                    email,
                    row["secret"],
                    cookie=cookie,
                    invite="",
                    otp_link=mail_providers.inbox_url(
                        row["provider"], email
                    ),
                    cookie_taken_at=now,
                )
            output_path = _relogin_output_path()
            if output_path:
                try:
                    storage._write_relogin_cookie(
                        output_path, email, cookie
                    )
                except RuntimeError as error:
                    say(f"Không ghi được file xuất {output_path}: {error}")
            storage._append_cookie_history(
                "lam-moi", email, "đăng nhập lại thành công"
            )
            say(
                f"✔ {email}: đã lấy cookie mới lúc {now} — "
                "còn sống 7 ngày."
            )
        finally:
            stop_drain.set()
            drain_thread.join(timeout=2)


# --- HTTP ----------------------------------------------------------


def _json_bytes(payload, status=200):
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    return status, body


class CookieWebHandler(BaseHTTPRequestHandler):
    server_version = "CookieWeb/1.0"

    def log_message(self, *args):
        # Server chạy ngầm trên máy người dùng: không spam console.
        pass

    # -- Helpers ---------------------------------------------------

    def _send(self, status, body, content_type):
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _send_json(self, payload, status=200):
        code, body = _json_bytes(payload, status)
        self._send(code, body, "application/json; charset=utf-8")

    def _send_html(self, html):
        self._send(
            200, html.encode("utf-8"), "text/html; charset=utf-8"
        )

    def _read_json(self):
        length = int(self.headers.get("Content-Length") or 0)
        if length <= 0:
            return {}
        try:
            return json.loads(self.rfile.read(length).decode("utf-8"))
        except (ValueError, UnicodeDecodeError):
            return {}

    @property
    def web(self):
        return self.server.web

    # -- Routes ----------------------------------------------------

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path
        try:
            if path in ("/", "/index.html"):
                self._send_html(DASHBOARD_HTML)
            elif path == "/api/cookies":
                self._send_json(self.web.list_cookies())
            elif path == "/api/cookie":
                email = parse_qs(parsed.query).get("email", [""])[0]
                self._send_json({"cookie": self.web.full_cookie(email)})
            elif path == "/api/jobs":
                self._send_json({"jobs": self.web.list_jobs()})
            elif path == "/api/history":
                self._send_json({"history": self.web.history()})
            elif path == "/api/status":
                self._send_json(self.web.status())
            elif path == "/api/export-excel":
                body = self.web.export_excel_bytes()
                self.send_response(200)
                self.send_header(
                    "Content-Type",
                    "application/vnd.openxmlformats-officedocument"
                    ".spreadsheetml.sheet",
                )
                self.send_header(
                    "Content-Disposition",
                    'attachment; filename="accounts.xlsx"',
                )
                self.send_header("Content-Length", str(len(body)))
                self.send_header("Cache-Control", "no-store")
                self.end_headers()
                self.wfile.write(body)
            else:
                self._send_json({"error": "không có đường này"}, 404)
        except Exception as error:
            self._send_json({"error": str(error)}, 500)

    def do_POST(self):
        parsed = urlparse(self.path)
        path = parsed.path
        try:
            if path == "/api/refresh":
                data = self._read_json()
                email = (data.get("email") or "").strip()
                if not email:
                    self._send_json({"error": "thiếu email"}, 400)
                    return
                try:
                    job_id = self.web.start_refresh([email])
                except RuntimeError as error:
                    self._send_json({"error": str(error)}, 409)
                    return
                self._send_json({"job_id": job_id}, 202)
            elif path == "/api/refresh-all":
                data = self._read_json()
                scope = (data.get("scope") or "expiring").strip()
                cookies = cookie_vault.list_cookies()
                emails = select_refresh_emails(cookies, scope)
                if not emails:
                    self._send_json(
                        {"error": "không có tài khoản nào thuộc diện này"},
                        400,
                    )
                    return
                try:
                    job_id = self.web.start_refresh(emails)
                except RuntimeError as error:
                    self._send_json({"error": str(error)}, 409)
                    return
                self._send_json(
                    {"job_id": job_id, "count": len(emails)}, 202
                )
            elif path == "/api/check":
                data = self._read_json()
                email = (data.get("email") or "").strip()
                if not email:
                    self._send_json({"error": "thiếu email"}, 400)
                    return
                cookie = self.web.full_cookie(email)
                if not cookie:
                    self._send_json(
                        {"error": "tài khoản này chưa có cookie"}, 400
                    )
                    return
                checker = cookie_health.BrowserChecker()
                try:
                    checker.start()
                except RuntimeError as error:
                    self._send_json({"error": str(error)}, 500)
                    return
                try:
                    entry = cookie_health.check_and_cache(
                        email, cookie, checker=checker
                    )
                finally:
                    checker.close()
                self._send_json({"email": email, **entry})
            elif path == "/api/check-all":
                data = self._read_json()
                scope = (data.get("scope") or "all").strip()
                cookies = cookie_vault.list_cookies()
                emails = select_check_emails(cookies, scope)
                if not emails:
                    self._send_json(
                        {"error": "không có tài khoản nào có cookie"}, 400
                    )
                    return
                job_id = self.web.start_check_all(emails)
                self._send_json(
                    {"job_id": job_id, "count": len(emails)}, 202
                )
            elif path == "/api/delete":
                data = self._read_json()
                email = (data.get("email") or "").strip()
                if not email:
                    self._send_json({"error": "thiếu email"}, 400)
                    return
                ok, message = self.web.delete_cookie(email)
                self._send_json(
                    {"ok": ok, "message": message},
                    200 if ok else 404,
                )
            elif path == "/api/import-excel":
                data = self._read_json()
                source = (data.get("path") or "").strip() or None
                result = self.web.import_excel(source)
                self._send_json(result)
            else:
                self._send_json({"error": "không có đường này"}, 404)
        except Exception as error:
            self._send_json({"error": str(error)}, 500)


# --- Dashboard (một file HTML, không cần mạng ngoài) --------------


DASHBOARD_HTML = """<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Kho Cookie — Muse AI</title>
<style>
:root { --bg:#f3f5f7; --card:#ffffff; --line:#e3e8ee; --text:#1c2333; --dim:#667085;
  --green:#067647; --green-bg:#dcfae6; --amber:#b54708; --amber-bg:#fef0c7;
  --red:#b42318; --red-bg:#fee4e2; --gray:#475467; --gray-bg:#eaecf0;
  --blue:#1570ef; --blue-bg:#eff4ff; }
* { box-sizing:border-box; }
body { margin:0; background:var(--bg); color:var(--text);
  font:15px/1.55 -apple-system,"Segoe UI",Roboto,Arial,sans-serif; }
.wrap { max-width:1400px; margin:0 auto; padding:28px 18px 56px; }
header { margin-bottom:20px; }
h1 { font-size:24px; margin:0 0 4px; letter-spacing:-.2px; }
.sub { color:var(--dim); margin:0; }
.stats { display:grid; grid-template-columns:repeat(auto-fit,minmax(150px,1fr)); gap:12px; margin-bottom:18px; }
.stat { background:var(--card); border:1px solid var(--line); border-radius:14px; padding:14px 16px;
  box-shadow:0 1px 2px rgba(16,24,40,.05); }
.stat b { font-size:26px; display:block; line-height:1.1; }
.stat span { color:var(--dim); font-size:13px; }
.stat.warn b { color:var(--amber); }
.stat.bad b { color:var(--red); }
.stat.good b { color:var(--green); }
.toolbar { display:flex; gap:8px; flex-wrap:wrap; margin-bottom:14px; align-items:center; }
input[type=text], select { background:var(--card); color:var(--text); border:1px solid var(--line);
  border-radius:9px; padding:9px 12px; font:inherit; }
input[type=text] { flex:1; min-width:200px; }
button { background:var(--blue); color:#fff; border:0; border-radius:9px; padding:9px 14px;
  font:inherit; cursor:pointer; font-weight:600; white-space:nowrap; }
button.ghost { background:var(--card); color:var(--text); border:1px solid var(--line); font-weight:500; }
button.small { padding:5px 10px; font-size:13px; border-radius:7px; }
button.warn { background:#b54708; }
button:disabled { opacity:.5; cursor:default; }
.card { background:var(--card); border:1px solid var(--line); border-radius:14px; padding:18px;
  margin-bottom:18px; box-shadow:0 1px 2px rgba(16,24,40,.05); overflow-x:auto; }
.card h2 { font-size:16px; margin:0 0 4px; }
.card .hint { color:var(--dim); font-size:13px; margin:0 0 12px; }
table { width:100%; border-collapse:collapse; }
th, td { text-align:left; padding:9px 10px; border-bottom:1px solid var(--line); vertical-align:middle; }
th { color:var(--dim); font-size:12.5px; text-transform:uppercase; letter-spacing:.4px; }
tr:last-child td { border-bottom:0; }
.mono { font-family:Consolas,Menlo,monospace; font-size:13px; }
.nowrap { white-space:nowrap; }
.cell-flex { display:flex; align-items:center; gap:8px; }
.truncate { overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.cookie-masked { max-width:120px; }
.badge { display:inline-block; padding:3px 11px; border-radius:999px; font-size:13px; font-weight:700; white-space:nowrap; }
.badge.song { background:var(--green-bg); color:var(--green); }
.badge.sap-het { background:var(--amber-bg); color:var(--amber); }
.badge.het-han, .badge.chet { background:var(--red-bg); color:var(--red); }
.badge.khong-ro { background:var(--gray-bg); color:var(--gray); }
.badge.checking { background:#eef4ff; color:#175cd3; animation:pulse 1.2s ease-in-out infinite; }
@keyframes pulse { 0%,100% { opacity:1; } 50% { opacity:.5; } }
.count { font-variant-numeric:tabular-nums; font-weight:700; }
.small-dim { color:var(--dim); font-size:12.5px; }
.job { border:1px solid var(--line); border-radius:10px; padding:11px 13px; margin-bottom:10px; background:#fbfcfd; }
.job .head { display:flex; justify-content:space-between; gap:8px; }
.job pre { background:#101828; color:#e6e8eb; border-radius:8px; padding:10px; max-height:220px;
  overflow:auto; font-size:12.5px; white-space:pre-wrap; margin:10px 0 0; }
.empty { color:var(--dim); }
.errorbox { background:var(--red-bg); color:var(--red); border:1px solid #fecdca; border-radius:12px; padding:12px 14px; margin-bottom:14px; font-weight:600; }
.infobox { background:#eef4ff; color:#175cd3; border:1px solid #b2ccff; border-radius:12px; padding:12px 14px; margin-bottom:14px; font-weight:600; }
.pager { display:flex; gap:8px; align-items:center; margin-top:14px; flex-wrap:wrap; }
.pager .info { color:var(--dim); font-size:13px; margin-right:auto; }
.note { color:var(--dim); font-size:13px; }
.actions { white-space:nowrap; }
@media (max-width:760px) { .stats { grid-template-columns:repeat(2,1fr); } }
</style>
</head>
<body>
<div class="wrap">
  <header>
    <h1>🍪 Kho Cookie — Muse AI</h1>
    <p class="sub">Mỗi cookie sống <b>7 ngày</b> kể từ lúc lấy. Đăng nhập lại chạy đúng luồng tab Đăng nhập lại: không proxy, không thẻ. Cột Hoạt động kiểm tra thật cookie còn vào được Muse hay đã die.</p>
  </header>
  <div class="stats" id="stats"></div>
  <div class="errorbox" id="errorBox" style="display:none"></div>
  <div class="infobox" id="migrateBox" style="display:none"></div>
  <div class="toolbar">
    <input type="text" id="q" placeholder="Tìm email…" oninput="PAGE=1;render()">
    <select id="f" onchange="PAGE=1;render()">
      <option value="">Tất cả hạn</option>
      <option value="song">Còn sống</option>
      <option value="sap-het">Sắp hết</option>
      <option value="het-han">Hết hạn</option>
      <option value="khong-ro">Không rõ hạn</option>
    </select>
    <select id="hf" onchange="PAGE=1;render()">
      <option value="">Tất cả hoạt động</option>
      <option value="song">Đang hoạt động</option>
      <option value="chet">Die</option>
      <option value="khong-ro">Chưa kiểm tra</option>
    </select>
    <button id="btnCheckAll" onclick="checkAll('all')">Kiểm tra tất cả</button>
    <button class="ghost" onclick="checkAll('expiring')">Kiểm tra cookie sắp hết</button>
    <button id="btnAll" onclick="refreshScope('expiring','btnAll')">Đăng nhập lại sắp hết &amp; đã hết</button>
    <button class="ghost" onclick="refreshScope('expired','btnExpired')" id="btnExpired">Đăng nhập lại đã hết</button>
    <button class="ghost" onclick="loadAll()">Làm mới</button>
    <button class="ghost" onclick="importExcel()">Nhập từ Excel</button>
    <button class="ghost" onclick="exportExcel()">Xuất ra Excel</button>
  </div>
  <div class="card">
    <h2>Danh sách cookie</h2>
    <p class="hint">Nguồn: kho SQLite (accounts.db) — một dòng một tài khoản, email là khoá duy nhất. Đăng nhập lại cập nhật đúng dòng của tài khoản đó.</p>
    <div id="tableWrap"><p class="empty">Đang tải…</p></div>
    <div class="pager" id="pager"></div>
  </div>
  <div class="card" id="jobsCard" style="display:none">
    <h2>Lượt chạy</h2>
    <div id="jobs"></div>
  </div>
  <div class="card">
    <h2>Lịch sử</h2>
    <div id="history"><p class="empty">Đang tải…</p></div>
  </div>
  <p class="note">Server chỉ chạy trên máy này (127.0.0.1) — không ai khác mở được. Cookie đầy đủ không hiện trong bảng, chỉ lấy qua nút Sao chép. <span id="metaLine"></span></p>
</div>
<script>
let COOKIES = [];
let PAGE = 1, PER_PAGE = 50;
let CHECKING = {emails:new Set(), since:""};
const LABEL = {song:"Còn sống","sap-het":"Sắp hết","het-han":"Hết hạn","khong-ro":"Không rõ"};
const HEALTH = {song:"Hoạt động","chet":"Die","khong-ro":"Chưa kiểm tra"};
const ACTION = {"lay":"Lấy cookie","lam-moi":"Đăng nhập lại","xoa":"Xoá cookie"};

function fmtRemain(sec){
  if (sec === null || sec === undefined) return "—";
  const neg = sec < 0, a = Math.abs(sec);
  const d = Math.floor(a/86400), h = Math.floor(a%86400/3600), m = Math.floor(a%3600/60);
  // Gọn để vừa một dòng trong cột: có ngày thì bỏ phút,
  // dưới một ngày thì chỉ giờ + phút.
  let s;
  if (d) s = d+" ngày" + (h? " "+h+" giờ":"");
  else if (h) s = h+" giờ" + (m? " "+m+" phút":"");
  else s = m+" phút";
  return neg ? "Hết hạn "+s : "Còn "+s;
}
function fmtTime(iso){
  if (!iso) return "—";
  const d = new Date(iso);
  return d.toLocaleString("vi-VN", {day:"2-digit",month:"2-digit",year:"numeric",hour:"2-digit",minute:"2-digit"});
}
async function checkJson(r){
  const d = await r.json();
  if (!r.ok || d.error) throw new Error(d.error || ("lỗi HTTP " + r.status));
  return d;
}
async function fetchJson(url, timeoutMs){
  // Mỗi API có hạn riêng (15 giây): một đầu chậm/kẹt không
  // được kéo cả trang kẹt theo (lỗi 2026-11-17).
  const ctrl = new AbortController();
  const timer = setTimeout(()=>ctrl.abort(), timeoutMs || 15000);
  try {
    const r = await fetch(url, {signal: ctrl.signal});
    return await checkJson(r);
  } finally { clearTimeout(timer); }
}
async function loadAll(){
  const box = document.getElementById("errorBox");
  const mbox = document.getElementById("migrateBox");
  // Tải từng phần độc lập: phần nào lỗi thì phần đó hiện
  // "Chưa tải được", các phần còn lại vẫn hiện bình thường.
  const [c, h, s] = await Promise.all([
    fetchJson("/api/cookies").catch(()=>null),
    fetchJson("/api/history").catch(()=>null),
    fetchJson("/api/status").catch(()=>null),
  ]);
  if (!c) {
    box.style.display = "block";
    box.textContent = "Không tải được danh sách cookie (quá 15 giây hoặc lỗi mạng) — thử tải lại trang, hoặc kiểm tra cửa sổ đang chạy server.";
    document.getElementById("tableWrap").innerHTML = "<p class='empty'>Chưa tải được dữ liệu.</p>";
    document.getElementById("history").innerHTML = "<p class='empty'>Chưa tải được dữ liệu.</p>";
    return;
  }
  box.style.display = "none"; box.textContent = "";
  COOKIES = c.cookies || [];
  document.getElementById("stats").innerHTML =
    stat(COOKIES.length,"Tổng tài khoản","") +
    stat(c.stats.song,"Còn sống","good") + stat(c.stats["sap-het"],"Sắp hết","warn") +
    stat(c.stats["het-han"],"Hết hạn","bad") + stat(c.stats.chet,"Die","bad") +
    stat(c.stats["chua-kiem-tra"],"Chưa kiểm tra","") +
    stat(c.stats["khong-co-cookie"],"Chưa có cookie","");
  if (s) {
    document.getElementById("metaLine").textContent =
      "Kho: " + s.db_path + " (" + s.total + " tài khoản) · File nhập Excel: " + s.source_path + " · File xuất tab Đăng nhập lại: " + (s.output_path || "chưa đặt") + " · Đang có cookie: " + s.vault_count + "." +
      (s.last_import ? " Lần nhập đầu từ Excel: " + s.last_import.imported + " mới, " + s.last_import.updated + " gộp trùng, " + s.last_import.skipped + " dòng bỏ qua." : "");
    if (s.migrating) {
      mbox.style.display = "block";
      mbox.textContent = "Đang nhập tài khoản từ Excel vào kho lần đầu — bảng sẽ tự đầy khi xong (trang tự tải lại mỗi 30 giây).";
    } else {
      mbox.style.display = "none"; mbox.textContent = "";
    }
  }
  if (h) {
    document.getElementById("history").innerHTML = (h.history||[]).length
      ? "<table><tr><th>Thời gian</th><th>Việc</th><th>Email</th><th>Chi tiết</th></tr>" +
        h.history.map(e=>"<tr><td>"+fmtTime(e.at)+"</td><td>"+(ACTION[e.action]||e.action)+"</td><td>"+esc(e.email)+"</td><td>"+esc(e.detail||"")+"</td></tr>").join("") + "</table>"
      : "<p class='empty'>Chưa có lịch sử.</p>";
  } else {
    document.getElementById("history").innerHTML = "<p class='empty'>Chưa tải được lịch sử.</p>";
  }
  render();
}
function stat(n,label,cls){ return "<div class='stat "+cls+"'><b>"+n+"</b><span>"+label+"</span></div>"; }
function esc(s){ return String(s??"").replace(/[&<>"]/g, c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c])); }
function healthCell(c){
  // Tài khoản đang được job kiểm tra mà chưa có kết quả mới
  // thì hiện "Đang kiểm tra…" ở đúng dòng đó (mốc so sánh là
  // lúc job bắt đầu: đã có kết quả mới hơn thì hiện kết quả).
  const isChecking = CHECKING.emails.has(c.email) &&
    (!c.health_checked_at || new Date(c.health_checked_at) < new Date(CHECKING.since));
  const badge = isChecking
    ? "<span class='badge checking'>Đang kiểm tra…</span>"
    : "<span class='badge "+(c.health==="chet"?"chet":c.health)+"'>"+HEALTH[c.health]+"</span>";
  return "<td><div class='cell-flex'>"+badge+"<button class='ghost small' onclick='checkOne(&quot;"+esc(c.email)+"&quot;,this)'>Kiểm tra</button></div>" +
    (c.health_checked_at? "<div class='small-dim nowrap'>"+fmtTime(c.health_checked_at)+"</div>":"") + "</td>";
}
function render(){
  const q = document.getElementById("q").value.trim().toLowerCase();
  const f = document.getElementById("f").value;
  const hf = document.getElementById("hf").value;
  const rows = COOKIES.filter(c => (!q || c.email.toLowerCase().includes(q)) && (!f || c.status===f) && (!hf || c.health===hf));
  // Phân trang: chỉ vẽ 1 trang (mặc định 50 dòng) — vẽ cả 2k
  // dòng một lúc làm trình duyệt nặng, và đồng hồ đếm ngược
  // mỗi giây phải chạm vào từng ô.
  const totalPages = Math.max(1, Math.ceil(rows.length / PER_PAGE));
  if (PAGE > totalPages) PAGE = totalPages;
  const pageRows = rows.slice((PAGE-1)*PER_PAGE, PAGE*PER_PAGE);
  document.getElementById("pager").innerHTML =
    "<span class='info'>Hiển thị " + (rows.length? (PAGE-1)*PER_PAGE+1 : 0) + "–" + Math.min(rows.length, PAGE*PER_PAGE) + " / " + rows.length + " tài khoản</span>" +
    "<select onchange='PER_PAGE=parseInt(this.value,10);PAGE=1;render()'>" +
    [50,100,200].map(n=>"<option value='"+n+"'"+(n===PER_PAGE?" selected":"")+">"+n+" / trang</option>").join("") + "</select>" +
    "<button class='ghost small' onclick='PAGE=Math.max(1,PAGE-1);render()'"+(PAGE<=1?" disabled":"")+">‹ Trước</button>" +
    "<span>Trang "+PAGE+" / "+totalPages+"</span>" +
    "<button class='ghost small' onclick='PAGE=Math.min("+totalPages+",PAGE+1);render()'"+(PAGE>=totalPages?" disabled":"")+">Sau ›</button>";
  if (!rows.length){ document.getElementById("tableWrap").innerHTML = "<p class='empty'>Không có tài khoản nào khớp.</p>"; return; }
  document.getElementById("tableWrap").innerHTML =
    "<table><tr><th>Email</th><th>Cookie</th><th>Lấy lúc</th><th>Hết hạn</th><th>Còn lại</th><th>Hoạt động</th><th>Hạn</th><th></th></tr>" +
    pageRows.map(c=>"<tr><td class='nowrap'>"+esc(c.email)+"</td>" +
      "<td><div class='cell-flex'><span class='mono truncate cookie-masked'>"+esc(c.cookie_masked||"(trống)")+"</span><button class='ghost small' onclick='copyCookie(&quot;"+esc(c.email)+"&quot;,this)'>Sao chép</button></div></td>" +
      "<td class='nowrap'>"+fmtTime(c.cookie_taken_at)+"</td><td class='nowrap'>"+fmtTime(c.expires_at)+"</td>" +
      "<td class='count nowrap' data-remain='"+(c.remaining_seconds??"")+"'>"+fmtRemain(c.remaining_seconds)+"</td>" +
      healthCell(c) +
      "<td><span class='badge "+c.status+"'>"+LABEL[c.status]+"</span></td>" +
      "<td class='actions'><button class='small' onclick='refreshOne(&quot;"+esc(c.email)+"&quot;,this)'>Đăng nhập lại</button> " +
      "<button class='ghost small' onclick='deleteCookie(&quot;"+esc(c.email)+"&quot;,this)'>Xoá</button></td></tr>").join("") + "</table>";
}
setInterval(()=>{
  document.querySelectorAll("[data-remain]").forEach(el=>{
    const v = el.getAttribute("data-remain");
    if (v === "" || v === null) return;
    const n = parseInt(v,10) - 1;
    el.setAttribute("data-remain", String(n));
    el.textContent = fmtRemain(n);
  });
}, 1000);
async function copyCookie(email, btn){
  const r = await fetch("/api/cookie?email="+encodeURIComponent(email)).then(r=>r.json());
  if (!r.cookie){ alert("Tài khoản này chưa có cookie."); return; }
  try { await navigator.clipboard.writeText(r.cookie); }
  catch(e){
    const ta = document.createElement("textarea"); ta.value = r.cookie;
    document.body.appendChild(ta); ta.select(); document.execCommand("copy"); ta.remove();
  }
  btn.textContent = "Đã sao chép"; setTimeout(()=>btn.textContent="Sao chép", 1500);
}
async function checkOne(email, btn){
  btn.disabled = true; btn.textContent = "Đang kiểm tra…";
  const r = await fetch("/api/check", {method:"POST", headers:{"Content-Type":"application/json"},
    body: JSON.stringify({email})}).then(r=>r.json());
  btn.disabled = false; btn.textContent = "Kiểm tra";
  if (r.error){ alert(r.error); return; }
  loadAll();
}
async function checkAll(scope){
  const btn = document.getElementById("btnCheckAll"); btn.disabled = true;
  const r = await fetch("/api/check-all", {method:"POST", headers:{"Content-Type":"application/json"},
    body: JSON.stringify({scope: scope || "all"})}).then(r=>r.json());
  btn.disabled = false;
  if (r.error){ alert(r.error); return; }
  loadJobs();
}
async function autoCheck(){
  // Mở trang là tự kiểm tra, không cần bấm nút: chỉ kiểm tra
  // diện chưa kiểm tra bao giờ hoặc đã kiểm tra quá 24 giờ
  // trước; server không xếp job mới khi đang có job kiểm tra
  // chạy. Lỗi thì thôi — trang vẫn hiện số theo mốc 7 ngày,
  // lần sau mở lại sẽ tự kiểm tra tiếp.
  try {
    const r = await fetch("/api/check-all", {method:"POST", headers:{"Content-Type":"application/json"},
      body: JSON.stringify({scope:"stale"})}).then(r=>r.json());
    if (!r.error) loadJobs();
  } catch(e) { /* bỏ qua */ }
}
async function refreshOne(email, btn){
  btn.disabled = true;
  const r = await fetch("/api/refresh", {method:"POST", headers:{"Content-Type":"application/json"},
    body: JSON.stringify({email})}).then(r=>r.json());
  btn.disabled = false;
  if (r.error){ alert(r.error); return; }
  loadJobs();
}
async function refreshScope(scope, btnId){
  const btn = document.getElementById(btnId); btn.disabled = true;
  const r = await fetch("/api/refresh-all", {method:"POST", headers:{"Content-Type":"application/json"},
    body: JSON.stringify({scope:scope})}).then(r=>r.json());
  btn.disabled = false;
  if (r.error){ alert(r.error); return; }
  alert("Đã xếp " + r.count + " tài khoản vào hàng đợi đăng nhập lại.");
  loadJobs();
}
async function importExcel(){
  if (!confirm("Nhập tài khoản từ file Excel vào kho? Tài khoản đã có trong kho sẽ được gộp (không tạo dòng trùng).")) return;
  const r = await fetch("/api/import-excel", {method:"POST", headers:{"Content-Type":"application/json"},
    body: JSON.stringify({})}).then(r=>r.json());
  if (r.error){ alert(r.error); return; }
  alert("Nhập xong: " + r.imported + " mới, " + r.updated + " cập nhật, " + r.skipped + " dòng bỏ qua. Kho đang có " + r.total + " tài khoản.");
  loadAll();
}
function exportExcel(){ window.open("/api/export-excel", "_blank"); }
async function deleteCookie(email, btn){
  if (!confirm("Xoá cookie của " + email + "? (Dòng vẫn giữ lại, đăng nhập lại là có cookie mới.)")) return;
  const r = await fetch("/api/delete", {method:"POST", headers:{"Content-Type":"application/json"},
    body: JSON.stringify({email})}).then(r=>r.json());
  alert(r.message || r.error); loadAll();
}
async function loadJobs(){
  const r = await fetch("/api/jobs").then(r=>r.json());
  const jobs = r.jobs || [];
  const card = document.getElementById("jobsCard");
  if (!jobs.length){ card.style.display = "none"; return; }
  card.style.display = "block";
  // Tài khoản nào đang được job kiểm tra thì bảng hiện
  // "Đang kiểm tra…" ở đúng dòng đó — thẻ job chỉ hiện số
  // lượng + tiến độ, không liệt kê mail nữa.
  const checking = new Set();
  let since = "";
  for (const j of jobs){
    if (j.kind === "check" && (j.status === "queued" || j.status === "running")){
      for (const e of (j.emails || [])) checking.add(e);
      if (!since || j.created_at < since) since = j.created_at;
    }
  }
  CHECKING = {emails: checking, since: since};
  document.getElementById("jobs").innerHTML = jobs.map(j=>
    "<div class='job'><div class='head'><b>"+esc(j.id)+" — "+(j.kind==="check"?"Kiểm tra":"Đăng nhập lại")+" — "+(j.emails||[]).length+" tài khoản · đã xong "+(j.done||0)+"</b><span>"+j.status+"</span></div>" +
    (j.result? "<div>"+esc(j.result)+"</div>":"") +
    (j.log && j.log.length? "<pre>"+esc(j.log.slice(-60).join("\\n"))+"</pre>":"") + "</div>").join("");
  render();
  if (jobs.some(j=>j.status==="queued"||j.status==="running")) setTimeout(loadJobs, 2000);
}
loadAll(); loadJobs(); autoCheck(); setInterval(loadAll, 30000);
</script>
</body>
</html>
"""


# --- Khởi động ------------------------------------------------------


class _Server(ThreadingHTTPServer):
    daemon_threads = True
    allow_reuse_address = True

    def __init__(self, address, web):
        self.web = web
        super().__init__(address, CookieWebHandler)


def start_server(port=DEFAULT_PORT, desktop_coordinator=None):
    """Mở server quản lý cookie trên 127.0.0.1.

    Trả về (httpd, url). Thử port từ port trở đi tới khi bind được.
    """
    web = CookieWeb(desktop_coordinator=desktop_coordinator)
    # Nhập kho lần đầu chạy nền: lần đầu mở kho, chuyển 2k dòng
    # từ Excel sang SQLite có thể mất một lúc — trang không được
    # chờ nhập xong mới hiện (lỗi 2026-11-17: trang kẹt ở
    # "Đang tải…" trong lúc nhập).
    threading.Thread(
        target=storage._ensure_accounts_migrated,
        name="cookie-migrate", daemon=True,
    ).start()
    last_error = None
    for candidate in range(port, port + 10):
        try:
            httpd = _Server(("127.0.0.1", candidate), web)
            break
        except OSError as error:
            last_error = error
            httpd = None
    if httpd is None:
        raise RuntimeError(f"Không mở được cổng quản lý cookie: {last_error}")
    thread = threading.Thread(
        target=httpd.serve_forever, name="cookie-web", daemon=True
    )
    thread.start()
    return httpd, f"http://127.0.0.1:{httpd.server_address[1]}"


def open_cookie_web(desktop_coordinator=None):
    """Mở web quản lý cookie (một lần duy nhất mỗi tiến trình).

    Trả về URL. Lần gọi sau chỉ mở lại trình duyệt với URL cũ.
    """
    global _SERVER
    with _SERVER_LOCK:
        if _SERVER is not None:
            _httpd, url = _SERVER
            webbrowser.open(url)
            return url
        httpd, url = start_server(
            desktop_coordinator=desktop_coordinator
        )
        _SERVER = (httpd, url)
    webbrowser.open(url)
    return url


def main():
    """Chạy riêng web quản lý cookie: py main.py --cookie-web."""
    import argparse

    parser = argparse.ArgumentParser(
        description="Web quản lý kho cookie (reg-muse-ai)."
    )
    parser.add_argument(
        "--port", type=int, default=DEFAULT_PORT,
        help="cổng nghe trên 127.0.0.1 (mặc định 8765)",
    )
    args = parser.parse_args()
    httpd, url = start_server(port=args.port)
    print(f"Kho Cookie đang chạy tại {url}")
    print("Mở trình duyệt để quản lý — nhấn Ctrl+C để dừng.")
    webbrowser.open(url)
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nĐã dừng web quản lý cookie.")
