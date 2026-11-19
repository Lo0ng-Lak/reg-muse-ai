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
nộp thẻ, không mã mời, trình duyệt ẩn — nhiều luồng cùng lúc,
mỗi luồng một file hộp thư riêng. OTP lấy từ hộp thư của đúng
mail đó rồi lấy cookie mới qua helper --relogin. Cookie mới ghi đè
vào dòng hiện có trong kho SQLite (và đồng bộ sang file xuất
muse_cookies.xlsx của tab Đăng nhập lại nếu đã đặt đường dẫn).

Mỗi cookie sống 7 ngày kể từ lúc lấy. Dashboard còn kiểm tra
cookie thật sự hoạt động hay đã die (xem cookie_health.py) —
không chỉ "có cookie là được".

Chỉ dùng thư viện chuẩn của Python — không cần cài thêm gì.
"""

import hashlib
import json
import queue
import secrets
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
from web_console import ConsoleError, WebConsole


DEFAULT_PORT = 8765
MAX_LOG_LINES = 200
# Số trình duyệt headless kiểm tra cùng lúc. Mỗi luồng vẫn đi
# lần lượt phần mail của mình. Máy chủ 4 nhân, ~8GB RAM: 10
# trình duyệt là mức vừa. Từng để 5 nên cả nghìn tài khoản nhìn
# như kiểm tra từng cái. 10 luồng từng làm job kẹt (2026-11-17)
# khi một trình duyệt không trả về; giờ mỗi mail có hạn 75 giây
# thì bỏ trình duyệt đó và mở cái mới.
CHECK_WORKERS = 10
# Số luồng đăng nhập lại song song. Mỗi luồng một trình duyệt
# ẩn và một file hộp thư riêng (mailbox.txt, mailbox-2.txt, ...),
# cùng cách tab Đăng nhập lại chia luồng — không giẫm OTP.
REFRESH_WORKERS = 5
# Tối đa cho một cookie: yêu cầu nhanh (10 giây) + trình duyệt
# (mở trang 20 giây + chờ ổn định 20 giây) + ghi cache. Quá hạn
# này thì trình duyệt đã kẹt — bỏ checker đó, mở cái mới, lượt
# kiểm tra tiếp tục thay vì đứng cứng cả job.
CHECK_ONE_TIMEOUT_SECONDS = 75

_SERVER = None
_SERVER_LOCK = threading.Lock()


def _kiem_tra_mot_mail(email, cookie, checker, fast,
                       timeout_seconds=CHECK_ONE_TIMEOUT_SECONDS):
    """Kiểm tra một cookie, chờ tối đa timeout_seconds.

    Chỉ dùng với checker giả trong test — KHÔNG dùng với
    Playwright. sync_playwright gắn greenlet vào đúng luồng đã
    start(); gọi check() từ luồng khác ném 'Cannot switch to a
    different thread'. Lượt kiểm tra thật đi qua _LuongKiemTra:
    start/check/close nằm trên một luồng chủ.

    Trả về entry {status, checked_at, detail}, hoặc None khi
    quá hạn. Luồng kẹt cứ chạy nền (daemon): nếu sau đó nó trả
    kết quả thì kết quả vẫn được ghi vào cache, chỉ là lượt
    kiểm tra không chờ nó nữa.
    """
    result = {}

    def _chay():
        try:
            result["entry"] = cookie_health.check_and_cache(
                email, cookie, checker=checker, fast_checker=fast,
            )
        except Exception as error:
            result["error"] = error

    thread = threading.Thread(target=_chay, daemon=True)
    thread.start()
    thread.join(timeout=timeout_seconds)
    if thread.is_alive():
        return None
    if "error" in result:
        raise result["error"]
    return result.get("entry")


def _mo_cap_kiem_tra(factory, fast_factory):
    """Mở một cặp kiểm tra (trình duyệt + yêu cầu nhanh).

    Trả về (checker, fast, loi) — fast có thể là None khi yêu
    cầu nhanh không mở được; checker None + loi là lý do khi
    trình duyệt không mở được (người gọi dừng luồng).
    """
    checker = factory()
    try:
        checker.start()
    except RuntimeError as error:
        return None, None, str(error)
    fast = None
    if fast_factory is not None:
        try:
            fast = fast_factory()
            fast.start()
        except Exception:
            fast = None
    return checker, fast, None


class _LuongKiemTra(threading.Thread):
    """Một luồng chủ Playwright: chỉ luồng này được start/check/close.

    Playwright sync gắn greenlet vào đúng luồng đã gọi start().
    Gọi check() hay close() từ luồng khác ném 'Cannot switch to
    a different thread' (lỗi 2026-11-17, mọi cookie đều fail khi
    chờ timeout bằng một luồng con). Luồng kiểm tra bên ngoài
    chỉ gửi lệnh qua hàng đợi và chờ có hạn — không đụng tới
    trình duyệt.
    """

    def __init__(self, factory, fast_factory):
        super().__init__(daemon=True, name="cookie-check-owner")
        self._factory = factory
        self._fast_factory = fast_factory
        self._commands = queue.Queue()
        self.ready = threading.Event()
        self.start_error = None
        self._bo = False
        self._da_dong = False

    def run(self):
        checker = None
        fast = None
        try:
            checker, fast, error = _mo_cap_kiem_tra(
                self._factory, self._fast_factory
            )
            if checker is None:
                self.start_error = error
                return
        except Exception as error:
            self.start_error = str(error)
            return
        finally:
            self.ready.set()
        try:
            while True:
                cmd = self._commands.get()
                if cmd is None:
                    break
                email, cookie, reply = cmd
                try:
                    entry = cookie_health.check_and_cache(
                        email,
                        cookie,
                        checker=checker,
                        fast_checker=fast,
                    )
                    # Đăng nhập lại vừa ghi cookie mới trong lúc
                    # lượt này còn cầm cookie cũ: kết quả Die là
                    # của cookie cũ, không được đè Còn sống.
                    hien_tai = _cookie_dang_luu(email)
                    if (
                        hien_tai
                        and hien_tai.strip() != (cookie or "").strip()
                    ):
                        entry = cookie_health.ghi_con_song(email)
                    reply.put(("ok", entry))
                except Exception as error:
                    reply.put(("err", error))
        finally:
            if fast is not None:
                try:
                    fast.close()
                except Exception:
                    pass
            try:
                checker.close()
            except Exception:
                pass

    def kiem_tra(self, email, cookie, timeout_seconds):
        """Trả về entry, hoặc None khi quá hạn (luồng chủ đang kẹt)."""
        reply = queue.Queue(maxsize=1)
        self._commands.put((email, cookie, reply))
        try:
            kind, value = reply.get(timeout=timeout_seconds)
        except queue.Empty:
            return None
        if kind == "err":
            raise value
        return value

    def bo(self):
        """Luồng chủ đang kẹt trong Playwright: không close() từ
        đây (cùng lỗi greenlet). Gửi lệnh đóng để nó tự đóng sau
        khi lần gọi kẹt trả về. Nếu lần gọi không bao giờ trả về
        thì trình duyệt kẹt theo tiến trình (daemon)."""
        self._bo = True
        self._commands.put(None)

    def dong(self):
        if self._bo or self._da_dong:
            return
        self._da_dong = True
        if not self.start_error:
            self._commands.put(None)
        self.join(timeout=30)


def _bat_luong_kiem_tra(job, factory, fast_factory, lock):
    """Mở luồng chủ. None khi trình duyệt không mở được (đã ghi log)."""
    owner = _LuongKiemTra(factory, fast_factory)
    owner.start()
    owner.ready.wait()
    if owner.start_error:
        with lock:
            job["log"].append(
                f"✘ luồng kiểm tra: {owner.start_error}"
            )
            job["log"] = job["log"][-MAX_LOG_LINES:]
        owner.dong()
        return None
    return owner


def _cookie_dang_luu(email):
    """Cookie hiện có trong kho của một email, hoặc ""."""
    import account_store

    row = account_store.get_account(email, storage.ACCOUNTS_DB_FILE)
    if not row:
        return ""
    return row.get("cookie") or ""


def _coordinator_dang_nhap_lai(suffix):
    """Một luồng đăng nhập lại: trình duyệt ẩn, file hộp thư và
    proxy riêng theo suffix ("" hoặc "-2", "-3", ...)."""
    coord = Coordinator()
    coord._headless = True
    coord._lane_local.suffix = suffix
    return coord


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
# Đã kiểm tra còn sống, hoặc vừa đăng nhập được, thì 7 ngày
# mới xem lại.
CHECK_ALIVE_HOURS = 7 * 24


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


def _den_luc_kiem_tra_lai(item):
    """True khi tài khoản này cần được kiểm tra lại.

    Chưa kiểm tra, hoặc lần trước chưa rõ: kiểm tra tiếp.
    Còn sống: 7 ngày mới kiểm lại.
    Die: không kiểm tra lại. Die thì đăng nhập lấy cookie mới.
    """
    health = item.get("health")
    checked_at = item.get("health_checked_at")
    if health == cookie_health.STATUS_DEAD:
        return False
    if health == cookie_health.STATUS_ALIVE:
        return not _checked_recently(checked_at, CHECK_ALIVE_HOURS)
    if health == cookie_health.STATUS_UNKNOWN:
        return True
    return not _checked_recently(checked_at)


def select_check_emails(cookies, scope):
    """Chọn mail kiểm tra sống/chết theo diện (dùng cho /api/check-all).

    - "stale": chưa kiểm tra bao giờ, còn sống đã quá 7 ngày,
      hoặc lần trước ra chưa rõ. Còn sống trong 7 ngày thì
      không kiểm tra lại. Die không kiểm tra lại.
    - "expiring": có cookie và sắp hết/hết hạn theo mốc 7 ngày.
    - "all": tất cả tài khoản có cookie.
    """
    if scope == "stale":
        return [
            item["email"]
            for item in cookies
            if item["has_cookie"] and _den_luc_kiem_tra_lai(item)
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


# Vòng nền hỏi lại kho khi không còn mail cần kiểm tra.
# Còn sống thì 7 ngày mới kiểm lại. Die không kiểm tra lại,
# chỉ đăng nhập để lấy cookie mới.
# Mail chưa xác minh thì kiểm tra lại cho đến khi kết luận được.
CHECK_LOOP_IDLE_SECONDS = 300
CHECK_LOOP_POLL_SECONDS = 15
# Chừa một nhịp sau khi một lượt vừa xong, để lượt kiểm tra
# lại các mail chưa rõ kịp được xếp trước khi vòng nền xếp lượt mới.
CHECK_LOOP_GAP_SECONDS = 10
# Server tự kiểm tra. Bộ test tắt cờ này để không mở trình duyệt.
TU_DONG_KIEM_TRA = True


def select_mail_vong_nen(cookies):
    """Mail vòng nền sẽ kiểm tra.

    Có cookie, và đến lúc kiểm tra lại: chưa kiểm tra, chưa rõ,
    hoặc còn sống quá 7 ngày. Mail còn sống trong 7 ngày thì
    không kiểm tra lại. Die không đưa vào vòng này.
    """
    return [
        item["email"]
        for item in cookies
        if item.get("has_cookie") and _den_luc_kiem_tra_lai(item)
    ]


class CookieWeb:
    """Trạng thái của web quản lý cookie: jobs + khoá chạy."""

    def __init__(self, desktop_coordinator=None):
        # Coordinator của cửa sổ tkinter (nếu web được mở từ nút
        # trong tool): đăng nhập lại bị từ chối khi tool đang chạy
        # một lượt khác, để hai luồng không giành mailbox.txt.
        self.desktop_coordinator = desktop_coordinator
        self.console = WebConsole(coordinator=desktop_coordinator)
        self.token = ""
        self.jobs = {}
        self.jobs_lock = threading.Lock()
        self.refresh_lock = threading.Lock()
        # Ghi file xuất Excel và lịch sử cookie: hai luồng cùng
        # đọc-sửa-ghi một file sẽ mất dòng.
        self._refresh_io_lock = threading.Lock()
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
        """Các tài khoản còn sống -> bytes file Excel."""
        import account_store

        alive = [
            item["email"]
            for item in cookie_vault.list_cookies()
            if item.get("health") == cookie_health.STATUS_ALIVE
        ]
        return account_store.export_to_excel_bytes(
            storage.ACCOUNTS_DB_FILE, only_emails=alive
        )

    def delete_cookie(self, email):
        """Xoá cookie của một tài khoản (giữ lại dòng + mã mời)."""
        if storage._clear_cookie_row(email):
            storage._append_cookie_history("xoa", email, "xoá cookie")
            return True, "Đã xoá cookie."
        return False, "Không tìm thấy tài khoản này trong kho."

    def history_counts(self):
        return storage._cookie_history_counts()

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
        if self.console is not None and getattr(
            getattr(self.console, "coordinator", None), "running", False
        ):
            raise RuntimeError(
                "Đang đăng ký hoặc đăng nhập — đợi lượt đó xong rồi "
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
        # Lượt người dùng bấm hoặc mở trang: mail chưa rõ được
        # kiểm tra lại một lần nữa khi lượt này xong. Lượt kiểm
        # tra lại thì không xếp thêm lượt nữa.
        job["kiem_tra_lai"] = True
        thread = threading.Thread(
            target=self._run_check_job,
            args=(job,),
            name=f"cookie-check-{job['id']}",
            daemon=True,
        )
        thread.start()
        return job["id"]

    def _co_job_kiem_tra(self):
        with self.jobs_lock:
            return any(
                job["kind"] == "check"
                and job["status"] in ("queued", "running")
                for job in self.jobs.values()
            )

    def bat_kiem_tra_neu_can(self, cookies=None):
        """Xếp một lượt kiểm tra nếu kho còn mail cần xem và chưa có lượt nào.

        Trả về id job vừa xếp, hoặc None khi không xếp.
        """
        if storage._accounts_migration_running():
            return None
        if self._co_job_kiem_tra():
            return None
        if cookies is None:
            cookies = cookie_vault.list_cookies()
        emails = select_mail_vong_nen(cookies)
        if not emails:
            return None
        job_id = self.start_check_all(emails)
        print(
            f"Vòng kiểm tra: {len(emails)} tài khoản ({job_id})",
            flush=True,
        )
        return job_id

    def bat_dau_vong_kiem_tra(self):
        """Server tự kiểm tra, không cần mở trang."""
        thread = threading.Thread(
            target=self._vong_kiem_tra,
            name="cookie-check-loop",
            daemon=True,
        )
        thread.start()

    def _vong_kiem_tra(self):
        while True:
            try:
                if (
                    self._co_job_kiem_tra()
                    or storage._accounts_migration_running()
                ):
                    time.sleep(CHECK_LOOP_POLL_SECONDS)
                    continue
                time.sleep(CHECK_LOOP_GAP_SECONDS)
                started = self.bat_kiem_tra_neu_can()
            except Exception as error:
                print(f"Vòng kiểm tra lỗi: {error}", flush=True)
                time.sleep(CHECK_LOOP_IDLE_SECONDS)
                continue
            if not started:
                time.sleep(CHECK_LOOP_IDLE_SECONDS)

    def list_jobs(self):
        with self.jobs_lock:
            jobs = sorted(
                self.jobs.values(),
                key=lambda job: job["created_at"],
                reverse=True,
            )
        return [dict(job, log=list(job["log"])) for job in jobs[:20]]

    def _run_refresh_job(self, job):
        """Một job: đăng nhập lại nhiều luồng, mỗi luồng một
        trình duyệt ẩn.

        Vẫn chỉ một job đăng nhập lại tại một thời điểm. Trong
        job, mail được chia đều cho REFRESH_WORKERS luồng. File
        hộp thư là mailbox-11.txt trở đi, không đụng mailbox.txt
        và mailbox-2.txt … mailbox-10.txt của lượt đăng ký.
        """
        with self.refresh_lock:
            job["status"] = "running"
            emails = list(job["emails"])
            workers = max(1, min(REFRESH_WORKERS, len(emails) or 1))
            counts = {"done": 0, "failed": 0}
            log_lock = threading.Lock()
            threads = []
            for index in range(workers):
                # 11 trở đi: đăng ký dùng mailbox.txt và mailbox-2..10.
                suffix = f"-{index + 11}"
                thread = threading.Thread(
                    target=self._refresh_slice,
                    args=(
                        job,
                        emails[index::workers],
                        suffix,
                        counts,
                        log_lock,
                    ),
                    name=f"cookie-refresh-{job['id']}-{index}",
                    daemon=True,
                )
                thread.start()
                threads.append(thread)
            for thread in threads:
                thread.join()
            job["status"] = "done"
            job["result"] = (
                f"Xong: {counts['done']} thành công, "
                f"{counts['failed']} thất bại "
                f"(trong {len(emails)} tài khoản)."
            )

    def _refresh_slice(self, job, emails, suffix, counts, log_lock):
        """Một luồng đăng nhập lại: lần lượt các mail của mình."""
        for email in emails:
            try:
                self._refresh_one(
                    job, email, suffix=suffix, log_lock=log_lock
                )
            except Exception as error:
                with log_lock:
                    counts["failed"] += 1
                    job["done"] = counts["done"] + counts["failed"]
                    job["log"].append(f"✘ {email}: {error}")
                    job["log"] = job["log"][-MAX_LOG_LINES:]
                continue
            with log_lock:
                counts["done"] += 1
                job["done"] = counts["done"] + counts["failed"]

    def _run_check_job(self, job, checker_factory=None,
                       fast_factory=None):
        """Một job: kiểm tra sống/chết — chia mail cho nhiều luồng
        song song, mỗi luồng một trình duyệt headless riêng.

        Mỗi luồng thử yêu cầu nhanh (ApiChecker) trước: kết luận
        được sống/chết thì khỏi mở trình duyệt; không kết luận
        được thì qua trình duyệt như trước. Mở thử một trình
        duyệt trước: không mở được thì dừng ngay (báo rõ lý do),
        không chạy cả nghìn lần vô ích. Xong lượt kiểm tra,
        cookie nào đã chết thì tự động đăng nhập lại lấy cookie
        mới (yêu cầu 2026-11-17).
        """
        factory = checker_factory or cookie_health.BrowserChecker
        fast_factory = fast_factory or cookie_health.ApiChecker
        probe = factory()
        try:
            probe.start()
            probe.smoke_test()
        except Exception as error:
            job["status"] = "done"
            job["result"] = f"Không kiểm tra được: {error}"
            job["log"].append(str(error))
            job["log"] = job["log"][-MAX_LOG_LINES:]
            return
        finally:
            probe.close()
        job["status"] = "running"
        emails = list(job["emails"])
        # Đọc cookie một lần cho cả job thay vì đọc lại cả kho
        # cho từng mail (2252 mail × 2252 dòng).
        cookies_by_email = {
            (row.get("email") or "").strip().lower(): (
                row.get("cookie") or ""
            )
            for row in storage._read_cookie_rows()
        }
        workers = max(1, min(CHECK_WORKERS, len(emails)))
        job["workers"] = workers
        counts = {"song": 0, "chet": 0, "khong-ro": 0}
        dead = []
        unclear = []
        lock = threading.Lock()
        threads = []
        for index in range(workers):
            thread = threading.Thread(
                target=self._check_slice,
                args=(
                    job,
                    emails[index::workers],
                    cookies_by_email,
                    factory,
                    counts,
                    dead,
                    lock,
                    fast_factory,
                    None,
                    unclear,
                ),
                name=f"cookie-check-{job['id']}-{index}",
                daemon=True,
            )
            thread.start()
            threads.append(thread)
        for thread in threads:
            thread.join()
        job["status"] = "done"
        job["result"] = (
            f"Xong: {counts['song']} hoạt động, {counts['chet']} die, "
            f"{counts['khong-ro']} chưa rõ "
            f"(trong {len(emails)} tài khoản)."
        )
        self._tu_dong_dang_nhap_lai_cookie_chet(job, dead)
        self._tu_dong_kiem_tra_lai_chua_ro(
            job, unclear, factory, fast_factory
        )

    def _tu_dong_kiem_tra_lai_chua_ro(self, job, unclear_emails,
                                      factory, fast_factory):
        """Mail chưa kết luận được sống/chết thì kiểm tra lại ngay
        một lần. Lượt đó không tự xếp thêm lượt ngay. Vòng nền
        vẫn xếp lại các mail chưa xác minh ở vòng sau, cho đến
        khi kết luận được sống hoặc die."""
        if not job.get("kiem_tra_lai"):
            return None
        if not unclear_emails:
            return None
        with self.jobs_lock:
            dang_chay = any(
                item["kind"] == "check"
                and item["status"] in ("queued", "running")
                for item in self.jobs.values()
            )
        if dang_chay:
            job["log"].append(
                f"↻ {len(unclear_emails)} cookie chưa rõ — đã có "
                "lượt kiểm tra đang chạy, để lượt đó làm."
            )
            job["log"] = job["log"][-MAX_LOG_LINES:]
            return None
        retry = self._new_job(list(unclear_emails), "check")
        retry["kiem_tra_lai"] = False
        thread = threading.Thread(
            target=self._run_check_job,
            args=(retry, factory, fast_factory),
            name=f"cookie-check-{retry['id']}",
            daemon=True,
        )
        thread.start()
        job["log"].append(
            f"↻ {len(unclear_emails)} cookie chưa rõ — tự kiểm tra "
            f"lại ({retry['id']})."
        )
        job["log"] = job["log"][-MAX_LOG_LINES:]
        return retry["id"]

    def _tu_dong_dang_nhap_lai_cookie_chet(self, job, dead_emails):
        """Sau một lượt kiểm tra: cookie nào đã chết thì tự động
        đăng nhập lại lấy cookie mới (luồng ReloginTab: không
        proxy, không thẻ). Đăng nhập lại thất bại ở tài khoản
        nào (ví dụ Meta bắt xác minh người) thì ghi vào log,
        không làm hỏng lượt kiểm tra vừa xong."""
        if not dead_emails:
            return None
        with self.jobs_lock:
            refresh_dang_chay = any(
                item["kind"] == "refresh"
                and item["status"] in ("queued", "running")
                for item in self.jobs.values()
            )
        if refresh_dang_chay:
            job["log"].append(
                f"↻ {len(dead_emails)} cookie đã chết — đã có lượt "
                "đăng nhập lại đang chạy, để lượt đó làm."
            )
            job["log"] = job["log"][-MAX_LOG_LINES:]
            return None
        try:
            refresh_id = self.start_refresh(list(dead_emails))
        except Exception as error:
            job["log"].append(
                "⚠ không tự đăng nhập lại được "
                f"({len(dead_emails)} cookie chết): {error}"
            )
            job["log"] = job["log"][-MAX_LOG_LINES:]
            return None
        job["log"].append(
            f"↻ {len(dead_emails)} cookie đã chết — tự động đăng "
            f"nhập lại ({refresh_id})."
        )
        job["log"] = job["log"][-MAX_LOG_LINES:]
        return refresh_id

    def _check_slice(self, job, emails, cookies_by_email, factory,
                     counts, dead, lock, fast_factory=None,
                     timeout_seconds=None, unclear=None):
        """Một luồng kiểm tra: một luồng chủ giữ trình duyệt, kiểm
        tra lần lượt từng mail trong phần của mình.

        Mỗi mail thử yêu cầu nhanh trước (một GET, không mở
        trang): kết luận được thì khỏi mở trình duyệt; yêu cầu
        nhanh không mở được (thiếu Playwright...) thì cả luồng
        chỉ dùng trình duyệt. start/check/close chỉ chạy trên
        luồng chủ — Playwright sync không cho gọi từ luồng khác
        (lỗi 'Cannot switch to a different thread'). Luồng này
        chỉ gửi lệnh và chờ. Một mail kẹt quá
        CHECK_ONE_TIMEOUT_SECONDS thì bỏ luồng chủ đó, mở luồng
        chủ mới, kiểm tra tiếp — job không đứng cứng (lỗi
        2026-11-17: kẹt ở 96/2096). Lỗi ở một mail không được
        giết cả luồng: ghi lại rồi tính là chưa rõ, kiểm tra
        tiếp mail sau.
        """
        timeout = (
            CHECK_ONE_TIMEOUT_SECONDS
            if timeout_seconds is None
            else timeout_seconds
        )
        if unclear is None:
            unclear = []
        owner = _bat_luong_kiem_tra(
            job, factory, fast_factory, lock
        )
        if owner is None:
            with lock:
                unclear.extend(emails)
            return
        try:
            for index, email in enumerate(emails):
                try:
                    cookie = cookies_by_email.get(
                        (email or "").strip().lower(), ""
                    )
                    fresh = cookie_health.load_health_cache().get(
                        (email or "").strip().lower()
                    ) or {}
                    if (
                        fresh.get("status") == cookie_health.STATUS_ALIVE
                        and _checked_recently(
                            fresh.get("checked_at"), CHECK_ALIVE_HOURS
                        )
                    ):
                        with lock:
                            counts["song"] += 1
                            job["done"] = sum(counts.values())
                            job["log"].append(
                                f"{email}: Còn sống (vừa đăng nhập "
                                "Muse từ Meta, chưa cần kiểm tra)."
                            )
                            job["log"] = job["log"][-MAX_LOG_LINES:]
                        continue
                    if fresh.get("status") == cookie_health.STATUS_DEAD:
                        with lock:
                            counts["chet"] += 1
                            job["done"] = sum(counts.values())
                            job["log"].append(
                                f"{email}: Die — không kiểm tra lại. "
                                "Đăng nhập để lấy cookie mới."
                            )
                            job["log"] = job["log"][-MAX_LOG_LINES:]
                        continue
                    if not cookie:
                        with lock:
                            job["log"].append(
                                f"— {email}: chưa có cookie."
                            )
                            counts["khong-ro"] += 1
                            job["done"] = sum(counts.values())
                            job["log"] = job["log"][-MAX_LOG_LINES:]
                        continue
                    entry = owner.kiem_tra(email, cookie, timeout)
                    if entry is None:
                        with lock:
                            counts["khong-ro"] += 1
                            job["done"] = sum(counts.values())
                            job["log"].append(
                                f"✘ {email}: kiểm tra kẹt quá "
                                f"{timeout} giây, bỏ qua."
                            )
                            job["log"] = job["log"][-MAX_LOG_LINES:]
                        owner.bo()
                        with lock:
                            unclear.append(email)
                        owner = _bat_luong_kiem_tra(
                            job, factory, fast_factory, lock
                        )
                        if owner is None:
                            with lock:
                                unclear.extend(emails[index + 1:])
                            return
                        continue
                    status = entry["status"]
                except Exception as error:
                    with lock:
                        counts["khong-ro"] += 1
                        job["done"] = sum(counts.values())
                        job["log"].append(f"✘ {email}: {error}")
                        job["log"] = job["log"][-MAX_LOG_LINES:]
                        unclear.append(email)
                    continue
                with lock:
                    counts[status] += 1
                    if status == cookie_health.STATUS_DEAD:
                        dead.append(email)
                    elif status == cookie_health.STATUS_UNKNOWN:
                        unclear.append(email)
                    job["done"] = sum(counts.values())
                    job["log"].append(
                        f"{email}: "
                        f"{cookie_health.STATUS_LABELS[status]} "
                        f"({entry['detail']})"
                    )
                    job["log"] = job["log"][-MAX_LOG_LINES:]
        finally:
            if owner is not None:
                owner.dong()

    # --- Đăng nhập lại một tài khoản ------------------------------

    def _refresh_one(self, job, email, suffix="", log_lock=None):
        """Đăng nhập lại MỘT tài khoản để lấy cookie mới.

        Chạy ĐÚNG luồng tab Đăng nhập lại: không proxy (ghi
        DIRECT), không nộp thẻ, không mã mời — helper chạy với
        --relogin và --headless (không mở cửa sổ). OTP lấy từ
        hộp thư của đúng mail đó. suffix là file riêng của luồng
        ("" hoặc "-2", "-3", ...). Cookie mới ghi đè vào dòng
        hiện có trong kho SQLite, đồng thời đồng bộ sang file
        xuất muse_cookies.xlsx của tab (nếu đã đặt đường dẫn).
        """
        if log_lock is None:
            log_lock = threading.Lock()
        log = job["log"]

        def say(message):
            with log_lock:
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
        coord = _coordinator_dang_nhap_lai(suffix)
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
                "— không proxy, không thẻ, ẩn trình duyệt."
            )
            try:
                storage._write_active_proxy("DIRECT", suffix)
            except OSError as error:
                say(f"Không ghi được proxy{suffix}.txt: {error}")
            try:
                storage._write_mailbox(
                    {
                        "provider": row["provider"],
                        "email": row["email"],
                        "secret": row["secret"],
                    },
                    suffix,
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
            export_error = None
            with self._refresh_io_lock:
                if output_path:
                    try:
                        storage._write_relogin_cookie(
                            output_path, email, cookie
                        )
                    except RuntimeError as error:
                        export_error = error
                storage._append_cookie_history(
                    "lam-moi", email, "đăng nhập lại thành công"
                )
            if "hatch_sess=" in (cookie or ""):
                cookie_health.ghi_con_song(email)
            if export_error is not None:
                say(
                    f"Không ghi được file xuất {output_path}: "
                    f"{export_error}"
                )
            say(
                f"✔ {email}: đã lấy cookie mới lúc {now} — "
                "đã ghi Còn sống."
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

    def _console_post(self, action):
        data = self._read_json()
        console = self.web.console
        try:
            if action == "cards":
                result = console.cards((data.get("action") or "").strip(), data)
            elif action == "capsolver":
                result = console.save_capsolver(data.get("key") or "")
            elif action == "capsolver-balance":
                result = console.capsolver_balance()
            elif action == "stop":
                result = console.stop()
            elif action == "muse/register":
                result = console.start_muse_register(data)
            elif action == "muse/login":
                result = console.start_muse_login(data)
            elif action == "meta":
                result = console.start_meta(data)
            elif action == "relogin":
                result = console.start_relogin(data)
            else:
                self._send_json({"error": "không có đường này"}, 404)
                return
        except ConsoleError as error:
            self._send_json({"error": str(error)}, 400)
            return
        self._send_json(result)

    # -- Routes ----------------------------------------------------

    def _session_ok(self):
        token = getattr(self.web, "token", "")
        if not token:
            return True
        raw = self.headers.get("Cookie") or ""
        for part in raw.split(";"):
            name, _, value = part.strip().partition("=")
            if name == "muse_session" and value == token:
                return True
        return False

    def _send_login(self):
        self._send_html(LOGIN_HTML)

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path
        try:
            if not self._session_ok():
                self._send_login()
                return
            if path in ("/", "/index.html", "/dang-ky"):
                self._send_html(DASHBOARD_HTML)
            elif path == "/api/console":
                since = parse_qs(parsed.query).get("since", ["0"])[0]
                self._send_json(self.web.console.snapshot(since))
            elif path == "/api/cookies":
                self._send_json(self.web.list_cookies())
            elif path == "/api/cookie":
                email = parse_qs(parsed.query).get("email", [""])[0]
                self._send_json({"cookie": self.web.full_cookie(email)})
            elif path == "/api/jobs":
                self._send_json({"jobs": self.web.list_jobs()})
            elif path == "/api/history":
                self._send_json({"counts": self.web.history_counts()})
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
            if path == "/api/login":
                data = self._read_json()
                digest = hashlib.sha256(
                    (data.get("password") or "").encode("utf-8")
                ).hexdigest()
                if not self.web.token or digest != self.web.token:
                    self._send_json({"error": "sai mật khẩu"}, 401)
                    return
                body = b'{"ok":true}'
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.send_header("Set-Cookie", f"muse_session={digest}; HttpOnly; Path=/; SameSite=Lax")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)
                return
            if not self._session_ok():
                self._send_json({"error": "cần đăng nhập"}, 401)
                return
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
                fast = cookie_health.ApiChecker()
                try:
                    fast.start()
                except Exception:
                    fast = None
                try:
                    entry = cookie_health.check_and_cache(
                        email, cookie, checker=checker,
                        fast_checker=fast,
                    )
                finally:
                    if fast is not None:
                        fast.close()
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
            elif path.startswith("/api/console/"):
                self._console_post(path[len("/api/console/"):])
            else:
                self._send_json({"error": "không có đường này"}, 404)
        except Exception as error:
            self._send_json({"error": str(error)}, 500)


LOGIN_HTML = """<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Muse</title>
<style>
body { margin:0; min-height:100vh; display:grid; place-items:center; background:#f3f5f7;
  font:15px/1.5 "Segoe UI",sans-serif; color:#1c2333; }
form { width:min(360px,92vw); background:#fff; border:1px solid #e3e8ee; border-radius:16px; padding:28px; }
h1 { margin:0 0 8px; font-size:22px; }
p { margin:0 0 16px; color:#667085; }
input { width:100%; box-sizing:border-box; border:1px solid #e3e8ee; border-radius:9px; padding:10px 12px; font:inherit; }
button { margin-top:12px; width:100%; background:#1570ef; color:#fff; border:0; border-radius:9px; padding:10px; font:inherit; font-weight:700; }
.err { color:#b42318; min-height:1.2em; font-size:13px; }
</style>
</head>
<body>
<form id="f">
  <h1>Muse</h1>
  <p>Nhập mật khẩu để vào trang quản lý.</p>
  <input id="p" type="password" autocomplete="current-password" placeholder="Mật khẩu">
  <div class="err" id="e"></div>
  <button type="submit">Vào</button>
</form>
<script>
document.getElementById("f").onsubmit = async function(ev){
  ev.preventDefault();
  const r = await fetch("/api/login", {method:"POST", headers:{"Content-Type":"application/json"},
    body: JSON.stringify({password: document.getElementById("p").value})});
  if (r.ok) location.reload();
  else document.getElementById("e").textContent = "Sai mật khẩu.";
};
</script>
</body>
</html>
"""


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
#viewRegister { display:none; }
html.register #viewRegister { display:block; }
html.register #viewHome { display:none; }
html.home .nav-main, html.register .nav-small { background:var(--blue); color:#fff; }
.topbar { position:sticky; top:0; z-index:20; display:flex; align-items:center; gap:18px;
  height:56px; padding:0 22px; background:#fff; border-bottom:1px solid var(--line); }
.brand { display:flex; align-items:center; gap:10px; color:var(--text); text-decoration:none; font-weight:700; }
.mark { width:28px; height:28px; border-radius:8px; background:var(--blue); color:#fff;
  display:grid; place-items:center; font-size:14px; }
.topbar nav { display:flex; align-items:center; gap:4px; }
.topbar nav a { text-decoration:none; color:var(--dim); font-weight:600; border-radius:8px; }
.nav-main { padding:6px 12px; font-size:14px; }
.nav-small { padding:3px 8px; font-size:12px; }
.run-banner { display:none; position:sticky; top:56px; z-index:19; padding:10px 22px;
  background:var(--blue-bg); color:var(--blue); border-bottom:1px solid #b2ddff;
  font-weight:700; }
.run-banner.show { display:block; }
.run-banner.done { background:var(--green-bg); color:var(--green); border-bottom-color:#abefc6; }
.reg-tabs { display:flex; gap:8px; margin:0 0 14px; }
.reg-tabs button { background:#fff; color:var(--text); border:1px solid var(--line); }
.reg-tabs button.on { background:var(--blue); color:#fff; }
.reg-grid { display:grid; grid-template-columns:minmax(0,1.45fr) minmax(280px,.8fr); gap:14px; }
.field { display:flex; flex-direction:column; gap:4px; margin-bottom:10px; }
.field > span { font-size:12.5px; color:var(--dim); font-weight:600; }
textarea.block, input.block, select.block { width:100%; background:#fff; color:var(--text);
  border:1px solid var(--line); border-radius:9px; padding:8px 10px; font:inherit; }
.rowline { display:flex; gap:8px; flex-wrap:wrap; align-items:flex-end; }
.logbox { background:#101828; color:#e6e8eb; border-radius:10px; min-height:200px; max-height:360px;
  overflow:auto; padding:12px; font:12.5px/1.45 Consolas,Menlo,monospace; white-space:pre-wrap; margin:0; }
.progress { color:var(--dim); font-size:13px; margin:8px 0 0; }
@media (max-width:900px) { .reg-grid { grid-template-columns:1fr; } }
.wrap { max-width:1400px; margin:0 auto; padding:28px 18px 56px; }
.wrap > header { margin-bottom:20px; }
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
th, td { text-align:left; padding:8px 12px; border-bottom:1px solid var(--line); vertical-align:middle; }
th { color:var(--dim); font-size:11px; font-weight:700; text-transform:uppercase; letter-spacing:.4px; }
tr:last-child td { border-bottom:0; }
tbody tr:hover td { background:#f8fafc; }
.email { font-weight:600; }
.when { color:var(--dim); font-size:13px; font-variant-numeric:tabular-nums; white-space:nowrap; }
.status-line { display:flex; gap:6px; align-items:center; flex-wrap:wrap; }
.nowrap { white-space:nowrap; }
.badge { display:inline-block; padding:3px 11px; border-radius:999px; font-size:13px; font-weight:700; white-space:nowrap; }
.badge.song { background:var(--green-bg); color:var(--green); }
.badge.sap-het { background:var(--amber-bg); color:var(--amber); }
.badge.het-han, .badge.chet { background:var(--red-bg); color:var(--red); }
.badge.khong-ro, .badge.chua-kiem-tra { background:var(--gray-bg); color:var(--gray); }
.badge.chua-xac-minh { background:var(--amber-bg); color:var(--amber); }
.badge.checking { background:#eef4ff; color:#175cd3; animation:pulse 1.2s ease-in-out infinite; }
@keyframes pulse { 0%,100% { opacity:1; } 50% { opacity:.5; } }
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
.history-counts { margin:6px 0 0; }
.history-counts b { font-size:22px; font-variant-numeric:tabular-nums; margin-right:4px; }
.history-counts span { color:var(--dim); margin:0 12px; }
.actions { white-space:nowrap; }
@media (max-width:760px) { .stats { grid-template-columns:repeat(2,1fr); } }
</style>
</head>
<body>
<header class="topbar">
  <a class="brand" href="/"><span class="mark">M</span>Muse</a>
  <nav>
    <a class="nav-main" href="/">Quản lý</a>
    <a class="nav-small" href="/dang-ky">Đăng ký</a>
  </nav>
</header>
<div id="runBanner" class="run-banner"></div>
<div class="wrap" id="viewHome">
  <header>
    <h1>🍪 Kho Cookie — Muse AI</h1>
    <p class="sub">Một dòng một tài khoản: lúc lấy cookie, còn sống hay die, và hạn 7 ngày.</p>
  </header>
  <div class="stats" id="stats"></div>
  <div class="errorbox" id="errorBox" style="display:none"></div>
  <div class="infobox" id="migrateBox" style="display:none"></div>
  <div class="toolbar">
    <input type="text" id="q" placeholder="Tìm email…" oninput="PAGE=1;render()">
    <select id="f" onchange="PAGE=1;render()">
      <option value="">Tất cả hạn</option>
      <option value="song">Còn hạn</option>
      <option value="sap-het">Sắp hết</option>
      <option value="het-han">Quá hạn</option>
      <option value="khong-ro">Không rõ hạn</option>
    </select>
    <select id="hf" onchange="PAGE=1;render()">
      <option value="">Tất cả hoạt động</option>
      <option value="song">Còn sống</option>
      <option value="chet">Die</option>
      <option value="chua-kiem-tra">Chưa kiểm tra</option>
      <option value="chua-xac-minh">Chưa xác minh</option>
    </select>
    <button id="btnCheckAll" onclick="checkAll('all')">Kiểm tra tất cả</button>
    <button class="ghost" onclick="checkAll('expiring')">Kiểm tra cookie sắp hết</button>
    <button id="btnAll" onclick="refreshScope('expiring','btnAll')">Đăng nhập lại sắp hết &amp; đã hết</button>
    <button class="ghost" onclick="refreshScope('dead','btnDead')" id="btnDead">Đăng nhập lại Die</button>
    <button class="ghost" onclick="loadAll()">Làm mới</button>
    <button class="ghost" onclick="importExcel()">Nhập từ Excel</button>
    <button class="ghost" onclick="exportExcel()">Xuất còn sống</button>
  </div>
  <div class="card">
    <h2>Danh sách</h2>
    <p class="hint">Cookie không hiện ở bảng. Đăng nhập lại cập nhật đúng dòng của email đó.</p>
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
  <p class="note"><span id="metaLine"></span></p>
</div>
<div class="wrap" id="viewRegister">
  <header>
    <h1>Đăng ký</h1>
    <p class="sub">Đăng ký Meta, rồi đăng nhập Muse bằng những tài khoản Meta đó. Không tạo tài khoản Muse mới. Trình duyệt chạy trên máy đang mở trang này.</p>
  </header>
  <div class="reg-tabs">
    <button type="button" class="on" id="tabMuse" onclick="showReg('muse')">Muse.ai</button>
    <button type="button" id="tabMeta" onclick="showReg('meta')">Meta</button>
    <button type="button" id="tabRelogin" onclick="showReg('relogin')">Đăng nhập lại</button>
  </div>
  <div class="reg-grid">
    <div>
      <div class="card" id="paneMuse">
        <h2>Muse.ai</h2>
        <p class="hint">Đăng nhập Muse bằng tài khoản Meta đã đăng ký. Chọn file Excel Meta, tool lấy mail trong file rồi nộp thẻ trong kho.</p>
        <div id="cardTable" class="small-dim">Đang tải kho thẻ…</div>
        <div class="field"><span>Dán thẻ (số|MM/YY|CVV hoặc thêm ZIP|nhãn)</span><textarea class="block" id="cardPaste" rows="3"></textarea></div>
        <div class="rowline">
          <button type="button" onclick="importCards()">Nhập kho thẻ</button>
          <button type="button" class="ghost" onclick="cardAction('auto')">Tự động xoay</button>
          <input class="block" id="cardLimit" type="number" min="0" max="100" value="0" style="width:80px">
          <button type="button" class="ghost" onclick="limitAll()">Áp dụng cho tất cả thẻ</button>
        </div>
        <div class="rowline">
          <button type="button" class="ghost" onclick="cardsSelected('delete')">Xoá đã chọn</button>
          <button type="button" class="ghost" onclick="cardsSelected('limit')">Giới hạn đã chọn</button>
          <button type="button" class="ghost" onclick="cardsSelected('pick')">Chỉ dùng thẻ đã chọn</button>
        </div>
        <div class="field" style="margin-top:12px"><span>Mã mời</span><input class="block" id="invite" type="text"></div>
        <label class="small-dim"><input type="checkbox" id="singleInvite"> Chỉ dùng 1 mã mời này</label>
        <div class="field" style="margin-top:12px"><span>File Excel Meta trên máy này</span><input class="block" id="metaExcel" type="text" placeholder="C:\\duong\\dan\\meta_accounts.xlsx"></div>
        <button type="button" onclick="startMuseLogin()">Đăng nhập Muse</button>
        <div id="museSetup"></div>
      </div>
      <div class="card" id="paneMeta" style="display:none">
        <h2>Meta</h2>
        <p class="hint">Đăng ký meta.ai, lưu meta_accounts.xlsx. Không nộp thẻ. Cần CapSolver và ảnh selfie.</p>
        <div class="field"><span>CapSolver API key</span><input class="block" id="capKey" type="password" placeholder="Đã lưu thì để trống"></div>
        <div class="rowline">
          <button type="button" class="ghost" onclick="saveKey()">Lưu key</button>
          <button type="button" class="ghost" onclick="checkBalance()">Số dư</button>
          <span class="small-dim" id="capState"></span>
        </div>
        <div class="field" style="margin-top:12px"><span>Ảnh selfie trên máy này</span><input class="block" id="selfie" type="text"></div>
        <div class="field"><span>Temp mail</span><select class="block" id="metaProvider"></select></div>
        <div id="metaSetup"></div>
        <button type="button" onclick="startMeta()">Đăng kí Meta</button>
      </div>
      <div class="card" id="paneRelogin" style="display:none">
        <h2>Đăng nhập lại</h2>
        <p class="hint">Lấy lại cookie của mail đã có trong kho. Không proxy, không thẻ. OTP lấy từ đúng hộp thư.</p>
        <p id="reloginAccounts" class="small-dim"></p>
        <div class="field"><span>File xuất cookie trên máy này</span><input class="block" id="reloginOutput" type="text"></div>
        <div class="rowline">
          <div class="field"><span>Số mail (0 = tất cả)</span><input class="block" id="reloginCount" type="number" min="0" value="0" style="width:120px"></div>
          <div class="field"><span>Số luồng</span><input class="block" id="reloginLanes" type="number" min="1" max="10" value="1" style="width:90px"></div>
          <label class="small-dim"><input type="checkbox" id="reloginHeadless"> Ẩn trình duyệt</label>
        </div>
        <button type="button" onclick="startRelogin()">Đăng nhập lại</button>
      </div>
    </div>
    <div class="card">
      <div class="rowline" style="justify-content:space-between">
        <h2 style="margin:0">Quá trình</h2>
        <button type="button" class="warn" id="btnStop" onclick="stopRun()" disabled>Dừng</button>
      </div>
      <p class="progress" id="runProgress">Sẵn sàng.</p>
      <pre class="logbox" id="runLog"></pre>
    </div>
  </div>
</div>
<script>
document.documentElement.className = location.pathname.indexOf("/dang-ky") === 0 ? "register" : "home";
let COOKIES = [];
let PAGE = 1, PER_PAGE = 50;
let CHECKING = {emails:new Set(), since:""};
const LABEL = {song:"Còn hạn","sap-het":"Sắp hết","het-han":"Quá hạn","khong-ro":"Không rõ"};
const HEALTH = {song:"Còn sống",chet:"Die","chua-kiem-tra":"Chưa kiểm tra","chua-xac-minh":"Chưa xác minh"};
function healthKind(c){
  if (c.health === "song" || c.health === "chet") return c.health;
  if (!c.has_cookie) return "";
  return c.health_checked_at ? "chua-xac-minh" : "chua-kiem-tra";
}
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
  return neg ? "Quá hạn "+s : "Còn "+s;
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
    stat(c.stats["hoat-dong"],"Còn sống","good") +
    stat(c.stats.chet,"Die","bad") +
    stat(c.stats["chua-kiem-tra"],"Chưa kiểm tra","") +
    stat(c.stats["chua-xac-minh"],"Chưa xác minh","warn") +
    stat(c.stats["sap-het"],"Sắp hết","warn") +
    stat(c.stats["het-han"],"Quá hạn","bad") +
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
    const n = h.counts || {};
    document.getElementById("history").innerHTML =
      "<p class='history-counts'><b>"+(n["lam-moi"]||0)+"</b> đăng nhập lại"+
      "<span>·</span><b>"+(n.lay||0)+"</b> lấy cookie"+
      "<span>·</span><b>"+(n.xoa||0)+"</b> xoá cookie</p>";
  } else {
    document.getElementById("history").innerHTML = "<p class='empty'>Chưa tải được lịch sử.</p>";
  }
  render();
}
function stat(n,label,cls){ return "<div class='stat "+cls+"'><b>"+n+"</b><span>"+label+"</span></div>"; }
function esc(s){ return String(s??"").replace(/[&<>"]/g, c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c])); }
function healthBadge(c){
  // Tài khoản đang được job kiểm tra mà chưa có kết quả mới
  // thì hiện "Đang kiểm tra…" ở đúng dòng đó (mốc so sánh là
  // lúc job bắt đầu: đã có kết quả mới hơn thì hiện kết quả).
  const isChecking = CHECKING.emails.has(c.email) &&
    (!c.health_checked_at || new Date(c.health_checked_at) < new Date(CHECKING.since));
  if (isChecking) return "<span class='badge checking'>Đang kiểm tra…</span>";
  const kind = healthKind(c);
  if (!kind) return "<span class='badge khong-ro'>Chưa có cookie</span>";
  const title = c.health_checked_at ? " title='Kiểm tra "+fmtTime(c.health_checked_at)+"'" : "";
  return "<span class='badge "+kind+"'"+title+">"+HEALTH[kind]+"</span>";
}
function render(){
  const q = document.getElementById("q").value.trim().toLowerCase();
  const f = document.getElementById("f").value;
  const hf = document.getElementById("hf").value;
  const rows = COOKIES.filter(c => (!q || c.email.toLowerCase().includes(q)) && (!f || c.status===f) && (!hf || healthKind(c)===hf));
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
    "<table class='acct'><tr><th>Email</th><th>Lấy lúc</th><th>Trạng thái</th><th>Hạn</th><th></th></tr>" +
    pageRows.map(c=>"<tr><td class='email'>"+esc(c.email)+"</td>" +
      "<td class='when'>"+fmtTime(c.cookie_taken_at)+"</td>" +
      "<td>"+healthBadge(c)+"</td>" +
      "<td><div class='status-line'><span class='badge "+c.status+"'>"+LABEL[c.status]+"</span>"+
        "<span class='small-dim'>"+fmtRemain(c.remaining_seconds)+"</span></div></td>" +
      "<td class='actions'><button class='ghost small' onclick='checkOne(&quot;"+esc(c.email)+"&quot;,this)'>Kiểm tra</button> "+
        "<button class='small' onclick='refreshOne(&quot;"+esc(c.email)+"&quot;,this)'>Đăng nhập lại</button> "+
        "<button class='ghost small' onclick='deleteCookie(&quot;"+esc(c.email)+"&quot;,this)'>Xoá</button></td></tr>").join("") + "</table>";
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
  // Mở trang là tự kiểm tra, không cần bấm nút: chưa kiểm tra
  // bao giờ, còn sống đã quá 7 ngày, hoặc lần trước chưa rõ.
  // Die không kiểm tra lại. Server không xếp job mới khi đang có
  // job kiểm tra chạy.
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
  const live = jobs.filter(j=>j.status==="queued"||j.status==="running");
  const card = document.getElementById("jobsCard");
  // Tài khoản nào đang được job kiểm tra thì bảng hiện
  // "Đang kiểm tra…" ở đúng dòng đó — thẻ job chỉ hiện số
  // lượng + tiến độ, không liệt kê mail nữa.
  // Lượt đã xong thì bỏ khỏi trang, không giữ log cũ.
  const checking = new Set();
  let since = "";
  for (const j of live){
    if (j.kind === "check"){
      for (const e of (j.emails || [])) checking.add(e);
      if (!since || j.created_at < since) since = j.created_at;
    }
  }
  CHECKING = {emails: checking, since: since};
  if (!live.length){
    card.style.display = "none";
    document.getElementById("jobs").innerHTML = "";
    render();
    setTimeout(loadJobs, 5000);
    return;
  }
  card.style.display = "block";
  document.getElementById("jobs").innerHTML = live.map(j=>
    "<div class='job'><div class='head'><b>"+esc(j.id)+" — "+(j.kind==="check"?"Kiểm tra":"Đăng nhập lại")+" — "+(j.emails||[]).length+" tài khoản · đã xong "+(j.done||0)+(j.workers?" · "+j.workers+" luồng":"")+"</b><span>"+j.status+"</span></div>" +
    (j.result? "<div>"+esc(j.result)+"</div>":"") +
    (j.log && j.log.length? "<pre>"+esc(j.log.slice(-60).join("\\n"))+"</pre>":"") + "</div>").join("");
  render();
  setTimeout(loadJobs, 2000);
}
function showReg(name){
  ["muse","meta","relogin"].forEach(function(id){
    document.getElementById("pane"+id.charAt(0).toUpperCase()+id.slice(1)).style.display = id===name ? "block" : "none";
    document.getElementById("tab"+id.charAt(0).toUpperCase()+id.slice(1)).classList.toggle("on", id===name);
  });
}
function setupHtml(prefix){
  return "<div class='field'><span>Proxy — mỗi dòng host:port:user:pass</span><textarea class='block' id='"+prefix+"Proxies' rows='3'></textarea></div>"+
    "<div class='field'><span>Proxy lỗi</span><textarea class='block' id='"+prefix+"Bad' rows='2'></textarea></div>"+
    "<div class='rowline'>"+
    "<div class='field'><span>Số tài khoản</span><input class='block' id='"+prefix+"Total' type='number' min='0' max='100' value='5' style='width:90px'></div>"+
    "<div class='field'><span>Tài khoản / proxy</span><input class='block' id='"+prefix+"Per' type='number' min='1' max='100' value='5' style='width:90px'></div>"+
    "<div class='field'><span>Số luồng</span><input class='block' id='"+prefix+"Lanes' type='number' min='1' max='5' value='1' style='width:80px'></div>"+
    "<div class='field'><span>Kết nối</span><select class='block' id='"+prefix+"Direct'><option value='proxy'>Proxy</option><option value='direct'>IP máy</option></select></div>"+
    "<label class='small-dim'><input type='checkbox' id='"+prefix+"Headless'> Ẩn trình duyệt</label></div>";
}
function setupBody(prefix){
  return {
    proxies: document.getElementById(prefix+"Proxies").value,
    bad: document.getElementById(prefix+"Bad").value,
    total: document.getElementById(prefix+"Total").value,
    per_proxy: document.getElementById(prefix+"Per").value,
    lanes: document.getElementById(prefix+"Lanes").value,
    headless: document.getElementById(prefix+"Headless").checked,
    direct: document.getElementById(prefix+"Direct").value === "direct",
  };
}
async function consolePost(path, body){
  const r = await fetch("/api/console/"+path, {method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify(body||{})});
  const d = await r.json();
  if (!r.ok || d.error) throw new Error(d.error || ("lỗi HTTP "+r.status));
  return d;
}
const cardTicks = new Set();
function toggleCard(id, on){
  if (on) cardTicks.add(id); else cardTicks.delete(id);
}
function toggleAllCards(on){
  document.querySelectorAll(".card-tick").forEach(function(box){
    box.checked = on;
    toggleCard(box.value, on);
  });
}
function renderCards(cards, active, picked){
  const use = {};
  (picked || []).forEach(function(id){ use[id] = true; });
  if (active && !(picked && picked.length)) use[active] = true;
  const alive = {};
  cards.forEach(function(c){ alive[c.id] = true; });
  Array.from(cardTicks).forEach(function(id){ if (!alive[id]) cardTicks.delete(id); });
  if (!cards.length){ document.getElementById("cardTable").innerHTML = "<p class='empty'>Chưa có thẻ.</p>"; return; }
  const allOn = cards.every(function(c){ return cardTicks.has(c.id); });
  document.getElementById("cardTable").innerHTML = "<table><tr><th><input type='checkbox'"+(allOn?" checked":"")+" onchange='toggleAllCards(this.checked)'></th><th>Tên</th><th>Đã dùng</th><th>Giới hạn</th><th></th></tr>"+
    cards.map(function(c){
      const mark = use[c.id] ? " · sẽ dùng" : "";
      const ticked = cardTicks.has(c.id) ? " checked" : "";
      return "<tr><td><input class='card-tick' type='checkbox' value='"+esc(c.id)+"'"+ticked+" onchange='toggleCard(this.value, this.checked)'></td><td>"+esc(c.label)+mark+"</td><td>"+c.used+"</td><td>"+(c.limit||"∞")+"</td><td class='actions'>"+
        "<button type='button' class='ghost small' onclick='cardAction(&quot;select&quot;,&quot;"+esc(c.id)+"&quot;)'>Chọn</button> "+
        "<button type='button' class='ghost small' onclick='cardAction(&quot;delete&quot;,&quot;"+esc(c.id)+"&quot;)'>Xoá</button></td></tr>";
    }).join("")+"</table>";
}
async function cardsSelected(action){
  const ids = Array.from(cardTicks);
  if (!ids.length){ alert("Tick ít nhất một thẻ."); return; }
  if (action === "delete" && !confirm("Xoá "+ids.length+" thẻ đã chọn khỏi kho?")) return;
  const body = {action: action === "delete" ? "delete-many" : (action === "limit" ? "limit-many" : "pick"), ids: ids};
  if (action === "limit") body.limit = document.getElementById("cardLimit").value;
  try { await consolePost("cards", body); }
  catch (error){ alert(error.message); return; }
  if (action === "delete") ids.forEach(function(id){ cardTicks.delete(id); });
  consoleReady = false; loadConsole();
}
async function importCards(){
  const text = document.getElementById("cardPaste").value;
  try { await consolePost("cards", {action:"import", text:text}); }
  catch (error){ alert(error.message); return; }
  document.getElementById("cardPaste").value = "";
  consoleReady = false; loadConsole();
}
async function cardAction(action, id){
  try { await consolePost("cards", {action:action, id:id||""}); }
  catch (error){ alert(error.message); return; }
  consoleReady = false; loadConsole();
}
async function limitAll(){
  try { await consolePost("cards", {action:"limit-all", limit:document.getElementById("cardLimit").value}); }
  catch (error){ alert(error.message); return; }
  consoleReady = false; loadConsole();
}
async function saveKey(){
  try { await consolePost("capsolver", {key:document.getElementById("capKey").value}); }
  catch (error){ alert(error.message); return; }
  document.getElementById("capKey").value = "";
  document.getElementById("capState").textContent = "Đã lưu key.";
}
async function checkBalance(){
  try { const d = await consolePost("capsolver-balance", {}); document.getElementById("capState").textContent = "Số dư $"+d.balance; }
  catch (error){ alert(error.message); }
}
async function startMuseLogin(){
  const body = Object.assign(setupBody("muse"), {excel:document.getElementById("metaExcel").value, invite:document.getElementById("invite").value, single_invite:document.getElementById("singleInvite").checked});
  try { await consolePost("muse/login", body); }
  catch (error){
    if (error.message.indexOf("Chưa có thẻ") === 0 && confirm(error.message)){
      body.allow_manual_card = true;
      try { await consolePost("muse/login", body); } catch (again){ alert(again.message); return; }
      paintBanner("Đang chạy…", true);
      return;
    }
    alert(error.message);
    return;
  }
  paintBanner("Đang chạy…", true);
}
async function startMeta(){
  try { await consolePost("meta", Object.assign(setupBody("meta"), {selfie:document.getElementById("selfie").value, provider:document.getElementById("metaProvider").value, capsolver:document.getElementById("capKey").value})); }
  catch (error){ alert(error.message); return; }
  paintBanner("Đang chạy…", true);
}
async function startRelogin(){
  try { await consolePost("relogin", {output:document.getElementById("reloginOutput").value, count:document.getElementById("reloginCount").value, lanes:document.getElementById("reloginLanes").value, headless:document.getElementById("reloginHeadless").checked}); }
  catch (error){ alert(error.message); return; }
  paintBanner("Đang chạy…", true);
}
async function stopRun(){
  try { await consolePost("stop", {}); } catch (error){ alert(error.message); }
}
let LOG_AT = 0, consoleReady = false;
function paintBanner(progress, running){
  const banner = document.getElementById("runBanner");
  if (!banner) return;
  const text = progress || "";
  const finished = text.indexOf("Hoàn tất") >= 0 || text.indexOf("Đã dừng") >= 0;
  if (running){
    banner.className = "run-banner show";
    banner.textContent = "Đang chạy — " + (text && text !== "Sẵn sàng." ? text : "đang xử lý…");
  } else if (finished){
    banner.className = "run-banner show done";
    banner.textContent = text;
  } else {
    banner.className = "run-banner";
    banner.textContent = "";
  }
}
function fillOnce(id, value){
  const el = document.getElementById(id);
  if (el && !el.dataset.filled){ el.value = value || ""; el.dataset.filled = "1"; }
}
async function loadConsole(){
  let d;
  try { d = await fetchJson("/api/console?since="+LOG_AT); }
  catch (error){ return; }
  LOG_AT = d.since || LOG_AT;
  if (d.log && d.log.length){
    const box = document.getElementById("runLog");
    box.textContent += (box.textContent ? "\\n" : "") + d.log.join("\\n");
    box.scrollTop = box.scrollHeight;
  }
  document.getElementById("runProgress").textContent = d.progress || "";
  paintBanner(d.progress, d.running);
  document.getElementById("btnStop").disabled = !d.running;
  renderCards(d.cards || [], d.active_card || "", d.picked || []);
  if (!consoleReady){
    fillOnce("museProxies", d.muse && d.muse.proxies);
    fillOnce("museBad", d.muse && d.muse.bad);
    fillOnce("metaProxies", d.meta && d.meta.proxies);
    fillOnce("metaBad", d.meta && d.meta.bad);
    fillOnce("metaExcel", d.meta_excel);
    fillOnce("selfie", d.selfie);
    fillOnce("invite", d.invite);
    fillOnce("reloginOutput", d.relogin_output);
    const sel = document.getElementById("metaProvider");
    if (sel && !sel.dataset.filled){
      sel.innerHTML = (d.providers||[]).map(function(p){ return "<option value='"+esc(p.id)+"'>"+esc(p.label)+"</option>"; }).join("");
      sel.dataset.filled = "1";
    }
    document.getElementById("capState").textContent = d.capsolver_set ? "Đã có key." : "Chưa lưu key.";
    consoleReady = true;
  }
  document.getElementById("reloginAccounts").textContent = "Kho: "+(d.accounts||0)+" tài khoản";
}
document.getElementById("museSetup").innerHTML = setupHtml("muse");
document.getElementById("metaSetup").innerHTML = setupHtml("meta");
loadConsole();
setInterval(loadConsole, 2000);
if (document.documentElement.className === "home"){
  loadAll(); loadJobs(); autoCheck(); setInterval(loadAll, 30000);
}
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


def _public_token(host, password):
    """Trang mở ra mạng thì cần mật khẩu. Máy này thì không.

    Có --password thì luôn khóa, kể cả khi nginx chuyển tiếp
    vào 127.0.0.1.
    """
    if password:
        return hashlib.sha256(password.encode("utf-8")).hexdigest()
    if host in ("127.0.0.1", "localhost", "::1"):
        return ""
    if not password:
        path = paths.data_file("web_password.txt")
        try:
            password = path.read_text(encoding="utf-8").strip()
        except OSError:
            password = ""
        if not password:
            password = secrets.token_urlsafe(9)
            path.write_text(password + "\n", encoding="utf-8")
        print(f"Mật khẩu trang: {password}")
        print(f"Lưu ở {path}")
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


def start_server(port=DEFAULT_PORT, desktop_coordinator=None, host="127.0.0.1", password=None):
    """Mở server trên host:port.

    Mặc định chỉ nghe 127.0.0.1. Truyền host 0.0.0.0 để máy khác
    mở được. Trả về (httpd, url). Thử port từ port trở đi tới
    khi bind được. port 0 để hệ điều hành chọn cổng trống.
    """
    web = CookieWeb(desktop_coordinator=desktop_coordinator)
    web.token = _public_token(host, password)
    try:
        rows = storage._read_cookie_rows()
        cookie_health.nang_cookie_moi_thanh_song(rows)
        excel = storage._load_meta_excel_path_text()
        if excel:
            cookie_health.nang_dang_nhap_meta_thanh_song(
                rows, storage._meta_da_dang_nhap_muse(excel)
            )
    except Exception:
        pass
    # Nhập kho lần đầu chạy nền: lần đầu mở kho, chuyển 2k dòng
    # từ Excel sang SQLite có thể mất một lúc — trang không được
    # chờ nhập xong mới hiện (lỗi 2026-11-17: trang kẹt ở
    # "Đang tải…" trong lúc nhập).
    threading.Thread(
        target=storage._ensure_accounts_migrated,
        name="cookie-migrate", daemon=True,
    ).start()
    last_error = None
    candidates = [0] if port == 0 else range(port, port + 10)
    for candidate in candidates:
        try:
            httpd = _Server((host, candidate), web)
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
    if TU_DONG_KIEM_TRA:
        web.bat_dau_vong_kiem_tra()
    shown_host = "127.0.0.1" if host in ("0.0.0.0", "::") else host
    return httpd, f"http://{shown_host}:{httpd.server_address[1]}"


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
        help="cổng nghe (mặc định 8765)",
    )
    parser.add_argument(
        "--host", default="127.0.0.1",
        help="địa chỉ nghe, 0.0.0.0 để mở ra mạng",
    )
    parser.add_argument(
        "--password", default="",
        help="mật khẩu khi mở ra mạng (bỏ trống thì tự tạo)",
    )
    args = parser.parse_args()
    httpd, url = start_server(
        port=args.port, host=args.host, password=args.password or None
    )
    print(f"Kho Cookie đang chạy tại {url}")
    print("Mở trình duyệt để quản lý — nhấn Ctrl+C để dừng.")
    webbrowser.open(url)
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nĐã dừng web quản lý cookie.")
