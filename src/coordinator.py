"""Động cơ chạy: điều phối các luồng đăng ký, không có tkinter.

MainWindow / các tab chỉ đọc setup từ giao diện, đóng gói thành
RunSettings rồi gọi Coordinator.start(). Mọi tiến trình chạy đều
báo về qua hàng đợi `events` — giao diện poll hàng đợi đó để ghi
log và cập nhật tiến độ.
"""

import queue
import subprocess
import sys
import threading
import time
from dataclasses import dataclass, field
from pathlib import Path

import mail_providers
import paths
import proxy_pool
import storage


# Giữ lại cho tương thích (test cũ đọc danh sách này): danh
# sách domain tempmail.plus sống ở mail_providers.
TEMPMAIL_DOMAINS = mail_providers.TEMPMAIL_PLUS_DOMAINS
HELPER_TIMEOUT_SECONDS = 900
PAUSE_BETWEEN_ACCOUNTS_SECONDS = 1
# Chế độ kết nối của lượt chạy: qua kho proxy, hoặc chạy trực
# tiếp qua IP của máy (người dùng tự bật VPN trước khi chạy).
# "DIRECT" là câu đánh dấu ghi vào proxy.txt để helper biết
# là không dùng proxy.
DIRECT_PROXY_LINE = "DIRECT"
CONN_MODES = (
    ("proxy", "Proxy (pool)"),
    ("direct", "IP máy (VPN)"),
)


def _create_mailbox_account(provider_id=mail_providers.DEFAULT_PROVIDER):
    """Tạo một hộp thư qua dịch vụ temp mail đã chọn.

    Toàn bộ logic theo từng dịch vụ nằm ở mail_providers.
    Trả về dict {"provider", "email", "secret"} — secret là
    mật khẩu hộp thư (mail.tm) hoặc token (guerrillamail);
    tempmail.plus và maildrop.cc để trống vì
    ở đó địa chỉ tự nó là khoá hộp thư.
    """
    return mail_providers.create_mailbox(provider_id)


def _session_cookie_from_line(line):
    """Tách cookie phiên từ dòng tín hiệu của helper (None nếu không phải).

    Cookie là thông tin đăng nhập — dòng này bị chặn lại, không
    cho hiện vào log của ứng dụng.
    """
    prefix = "SESSION_COOKIE:"
    if line.startswith(prefix):
        return line[len(prefix):].strip()
    return None


def _invite_from_line(line):
    """Tách mã mời của tài khoản từ dòng tín hiệu (None nếu không phải)."""
    prefix = "INVITE_CODE:"
    if line.startswith(prefix):
        return line[len(prefix):].strip()
    return None


@dataclass
class RunSettings:
    """Toàn bộ setup của một lượt chạy, đọc từ tab đã bấm Bắt đầu."""

    total: int = 0
    pool: list = field(default_factory=list)
    per_proxy: int = 5
    lanes: int = 1
    provider_id: str = mail_providers.DEFAULT_PROVIDER
    direct: bool = False
    meta: bool = False
    invite: str = ""
    single_invite: bool = False
    card_limit: int = 0
    card_used: int = 0
    card_tracked: bool = False
    # Kho proxy lỗi của tab đã bấm chạy: file + mapping đã đồng bộ
    # từ ô Proxy lỗi lúc bấm Bắt đầu.
    bad_file: Path | None = None
    bad_mapping: dict = field(default_factory=dict)
    # Panel nhận kết quả tra IP lúc bắt đầu lượt chạy VPN (None =
    # giao diện tự chọn panel đang chạy).
    ip_target: object = None


class Coordinator:
    """Chạy các lượt đăng ký trên luồng riêng, báo về qua events."""

    def __init__(self):
        self.events = queue.Queue()
        self.running = False
        self.stop_event = threading.Event()
        self.helper_process = None
        self.helper_processes = set()
        self._process_lock = threading.Lock()
        self._bad_lock = threading.Lock()
        self._bad_proxies = {}
        self._bad_file = None
        self._ip_target = None
        self._lane_local = threading.local()

    def start(self, settings):
        """Bắt đầu một lượt chạy trên luồng riêng.

        Trả về False nếu đang có lượt chạy khác (giao diện đã
        chặn trước, đây chỉ là chốt chặn cuối).
        """
        if self.running:
            return False
        self._bad_proxies = dict(settings.bad_mapping)
        self._bad_file = settings.bad_file
        self._ip_target = settings.ip_target
        self.running = True
        self.stop_event.clear()
        threading.Thread(
            target=self._run_loop, args=(settings,), daemon=True
        ).start()
        return True

    def stop(self):
        """Yêu cầu dừng lượt chạy (các luồng tự ngắt ở điểm an toàn)."""
        self.stop_event.set()

    def terminate_helpers(self):
        """Đóng mọi trình duyệt helper đang chạy (lúc đóng app)."""
        with self._process_lock:
            processes = list(self.helper_processes)
        if self.helper_process is not None:
            processes.append(self.helper_process)
        for process in processes:
            try:
                process.terminate()
            except OSError:
                pass

    def report_ip(self, info, target=None):
        """Ghi kết quả tra IP vào log + nhãn IP của panel đích.

        Chỉ hiển thị và cảnh báo — không chặn chạy. target là
        panel nhận kết quả (None = panel của lượt chạy đang chạy,
        giao diện tự chọn).
        """
        if target is None:
            target = self._ip_target
        if info is None:
            self.events.put(("ip", (target, "không kiểm tra được")))
            self.events.put(
                (
                    "log",
                    "Không kiểm tra được IP hiện tại của máy "
                    "(mạng hoặc dịch vụ tra IP lỗi) — chạy tiếp "
                    "mà không biết IP thoát ở đâu.",
                )
            )
            return
        status = proxy_pool._format_ip_status(info)
        self.events.put(("ip", (target, status)))
        line = f"IP hiện tại của máy: {status}"
        if info["org"]:
            line += f" — {info['org']}"
        self.events.put(("log", line))
        if info["country_code"] and info["country_code"] != "US":
            self.events.put(
                (
                    "log",
                    f"CẢNH BÁO: IP đang thoát ở "
                    f"{info['country']} "
                    f"({info['country_code']}), không phải US — "
                    "kiểm tra lại VPN trước khi chạy.",
                )
            )

    def _mark_proxy_bad(self, line):
        """Đưa một proxy vào kho lỗi (thread-safe, gọi từ các luồng).

        Ghi ngay xuống file kho lỗi của tab đã bấm chạy và báo cho
        giao diện thêm dòng vào ô Proxy lỗi của tab đó. Trả về
        True nếu đây là lần đầu proxy này bị đánh dấu.
        """
        key = proxy_pool._proxy_key(line)
        with self._bad_lock:
            if key in self._bad_proxies:
                return False
            self._bad_proxies[key] = line
            proxy_pool._save_bad_proxies(self._bad_proxies, self._bad_file)
        self.events.put(("bad_proxy", line))
        return True

    def _run_loop(self, settings):
        # Điều phối các luồng chạy song song. Kho proxy được
        # CHIA theo bước nhảy: luồng k dùng các proxy ở vị trí
        # k, k+lanes, k+2*lanes, ... — không hai luồng nào dùng
        # chung một proxy tại cùng thời điểm, và mỗi luồng xoay
        # vòng trong phần kho của riêng nó (proxy bị chặn cũng
        # chỉ ảnh hưởng luồng đó). Mỗi luồng có CHUỖI MÃ MỜI
        # RIÊNG, nhưng tài khoản ĐẦU TIÊN của mọi luồng đều
        # quy đổi mã trong ô Mã mời (mã chủ tool cài đặt); từ
        # tài khoản thứ hai, mỗi luồng nối tiếp mã do tài khoản
        # trước trong chính luồng đó lấy được. Nếu
        # single_invite bật thì KHÔNG nối chuỗi: MỌI tài khoản
        # của MỌI luồng đều quy đổi đúng mã trong ô Mã mời (mã
        # hạt giống); mã từng tài khoản tự lấy được chỉ ghi vào
        # sheet. Thẻ
        # thì cả các luồng dùng CHUNG một thẻ đã lưu; bộ đếm
        # giới hạn thẻ tính gộp toàn cục (đếm cả số tài khoản
        # đang chạy dở để không vượt giới hạn). Ở chế độ IP máy
        # (direct), pool chỉ có câu đánh dấu DIRECT: mọi luồng
        # đều dùng chung 1 IP của máy — số luồng chạy đúng như
        # đã đặt (chủ tool chấp nhận rủi ro IP bị chặn); mỗi
        # luồng vẫn có file proxy riêng (proxy.txt, proxy-2.txt,
        # ...) đều chứa câu DIRECT.
        total = settings.total
        pool = settings.pool
        direct = settings.direct
        meta = settings.meta
        requested_lanes = max(1, min(int(settings.lanes), 5))
        if direct:
            lanes = requested_lanes
            # Chế độ IP máy: kiểm tra IP thoát trước khi chạy —
            # chỉ hiển thị + cảnh báo, không chặn (tra lỗi hay IP
            # không phải US thì lượt chạy vẫn tiếp tục).
            self.report_ip(proxy_pool._check_public_ip())
        else:
            lanes = max(1, min(requested_lanes, len(pool)))
        if total > 0:
            lanes = min(lanes, total)
        if direct and lanes > 1:
            self.events.put(
                (
                    "log",
                    f"Chế độ IP máy (VPN): {lanes} luồng dùng chung "
                    "1 IP của máy — IP bị Meta chặn là dừng cả lượt "
                    "chạy.",
                )
            )
        shared = {
            "success": 0,
            "fail": 0,
            "card_used": settings.card_used,
            "in_flight": 0,
            "card_stop_logged": False,
        }
        shared_lock = threading.Lock()
        if lanes > 1:
            if meta:
                self.events.put(
                    (
                        "log",
                        f"Chạy {lanes} luồng song song (ĐĂNG KÍ "
                        "META) — mỗi luồng một phần kho proxy "
                        "riêng; tài khoản Meta tạo xong là lưu, "
                        "không qua Muse.",
                    )
                )
            elif direct:
                self.events.put(
                    (
                        "log",
                        f"Chạy {lanes} luồng song song ở chế độ IP "
                        "máy (VPN) — mọi luồng dùng chung 1 IP của "
                        "máy; mỗi luồng một chuỗi mã mời riêng "
                        "(tài khoản đầu của mọi luồng đều quy đổi "
                        "mã trong ô Mã mời); dùng chung thẻ đã "
                        "lưu, giới hạn thẻ tính gộp trên tất cả "
                        "các luồng.",
                    )
                )
            else:
                self.events.put(
                    (
                        "log",
                        f"Chạy {lanes} luồng song song — mỗi luồng "
                        "một phần kho proxy riêng, một chuỗi mã mời "
                        "riêng (tài khoản đầu của mọi luồng đều quy "
                        "đổi mã trong ô Mã mời); dùng chung thẻ đã "
                        "lưu, giới hạn thẻ tính gộp trên tất cả các "
                        "luồng.",
                    )
                )
        threads = []
        for lane_no in range(1, lanes + 1):
            if direct:
                subset = [DIRECT_PROXY_LINE]
            else:
                subset = pool[lane_no - 1::lanes]
            if total > 0:
                share = total // lanes + (
                    1 if lane_no <= total % lanes else 0
                )
            else:
                share = 0
            lane_args = (
                settings,
                lane_no,
                lanes,
                share,
                subset,
                shared,
                shared_lock,
                settings.invite,
            )
            if lanes == 1:
                self._run_lane(*lane_args)
            else:
                thread = threading.Thread(
                    target=self._run_lane,
                    args=lane_args,
                    name=f"lane-{lane_no}",
                    daemon=True,
                )
                thread.start()
                threads.append(thread)
        for thread in threads:
            thread.join()
        self.running = False
        self.events.put(
            (
                "finished",
                (
                    shared["success"],
                    shared["fail"],
                    self.stop_event.is_set(),
                ),
            )
        )

    def _run_lane(
        self, settings, lane_no, lanes, share, subset,
        shared, shared_lock, seed,
    ):
        helper = paths.SRC_DIR / "browser_helper.py"
        suffix = "" if lane_no == 1 else f"-{lane_no}"
        self._lane_local.suffix = suffix
        multi = lanes > 1
        lane_tag = f"[Luồng {lane_no}] " if multi else ""
        rotation = proxy_pool._ProxyRotation(subset, settings.per_proxy)
        written_proxy_index = None
        # Chuỗi mã mời CỦA LUỒNG NÀY: tài khoản sắp chạy sẽ quy
        # đổi pending_code (mã của tài khoản trước cùng luồng);
        # mã nó tự lấy được thành pending_code cho tài khoản sau
        # cùng luồng và được ghi đè vào invite.txt để lượt chạy
        # sau có một đầu chuỗi mà nối tiếp. Ở chế độ 1 mã mời
        # (single_invite), pending_code bị ĐÓNG BĂNG ở mã hạt
        # giống này cho mọi tài khoản — không bao giờ tiến lên.
        pending_code = (seed or "").strip().upper()
        # share = 0 là chế độ không giới hạn: luồng chạy tới khi
        # phần kho proxy của nó dùng hết một vòng thì dừng.
        label_total = share if share > 0 else "∞"
        index = 0
        while not self.stop_event.is_set():
            if share > 0 and index >= share:
                break
            if share == 0 and index > 0 and rotation.pool_exhausted():
                if multi:
                    message = (
                        f"{lane_tag}Đã dùng hết lượt của phần "
                        "proxy của luồng — luồng dừng."
                    )
                else:
                    message = (
                        "Đã dùng hết lượt của toàn bộ proxy trong "
                        "danh sách — dừng lượt chạy."
                    )
                self.events.put(("log", message))
                break
            if rotation.all_retired():
                self.events.put(
                    (
                        "log",
                        f"{lane_tag}Toàn bộ proxy của luồng đã bị "
                        "đưa vào kho proxy lỗi — luồng dừng.",
                    )
                )
                break
            with shared_lock:
                card_hit = (
                    settings.card_tracked
                    and settings.card_limit > 0
                    and shared["card_used"] + shared["in_flight"]
                    >= settings.card_limit
                )
                log_card_stop = (
                    card_hit and not shared["card_stop_logged"]
                )
                if log_card_stop:
                    shared["card_stop_logged"] = True
                if not card_hit:
                    # Giữ chỗ ngay trong cùng một khóa với phép
                    # kiểm tra: hai luồng không thể cùng thấy
                    # "còn chỗ" rồi cùng chạy vượt giới hạn thẻ.
                    # Chỗ được trả lại ở finally cuối vòng lặp.
                    shared["in_flight"] += 1
            if card_hit:
                if log_card_stop:
                    self.events.put(
                        (
                            "log",
                            f"Thẻ đang lưu đã đủ {settings.card_limit} tài "
                            "khoản thành công — dừng lượt chạy. "
                            "Muốn chạy tiếp thì đổi thẻ ở hàng Thẻ "
                            "rồi bấm Lưu thẻ (bộ đếm reset về 0).",
                        )
                    )
                break
            try:
                index += 1
                display = f"L{lane_no} · {index}" if multi else index
                proxy_index, line = rotation.current()
                if proxy_index != written_proxy_index:
                    written_proxy_index = proxy_index
                    try:
                        storage._write_active_proxy(line, suffix)
                        note = ""
                    except OSError as error:
                        note = f" (không ghi được proxy{suffix}.txt: {error})"
                    if settings.direct:
                        self.events.put(
                            (
                                "log",
                                "Chế độ IP máy (VPN): không dùng "
                                "proxy — mọi tài khoản chạy qua IP "
                                f"của máy này.{note}",
                            )
                        )
                    else:
                        self.events.put(
                            (
                                "log",
                                f"{lane_tag}Proxy #{proxy_index + 1}/{len(subset)}: "
                                f"{proxy_pool._proxy_display(line)} — dùng cho các tài "
                                f"khoản tiếp theo.{note}",
                            )
                        )
                with shared_lock:
                    s_now = shared["success"]
                    f_now = shared["fail"]
                    c_now = shared["card_used"]
                card_note = (
                    f" · Thẻ: {c_now}/{settings.card_limit}"
                    if settings.card_tracked
                    else ""
                )
                self.events.put(
                    (
                        "progress",
                        (
                            lane_no,
                            index,
                            label_total,
                            s_now,
                            f_now,
                            card_note,
                            multi,
                        ),
                    )
                )
                try:
                    created = _create_mailbox_account(settings.provider_id)
                except Exception as error:
                    with shared_lock:
                        shared["fail"] += 1
                    self.events.put(
                        ("log", f"[{display}/{label_total}] Tạo hộp thư thất bại: {error}")
                    )
                    rotation.advance_after_account()
                    continue
                if isinstance(created, dict):
                    mailbox = created
                else:
                    # Tương thích ngược (bản vá kiểm thử trả về
                    # chuỗi địa chỉ trần): coi như hộp thư không
                    # có bí mật ở dịch vụ đang chọn.
                    mailbox = {
                        "provider": settings.provider_id,
                        "email": str(created),
                        "secret": "",
                    }
                address = mailbox["email"]
                self.events.put(
                    (
                        "log",
                        f"[{display}/{label_total}] Đã tạo mail {address} — bắt đầu đăng ký.",
                    )
                )
                if not settings.meta:
                    try:
                        storage._write_redeem_code(pending_code, suffix)
                    except OSError:
                        pass
                try:
                    storage._write_mailbox(mailbox, suffix)
                except OSError:
                    pass
                if settings.meta:
                    # Xoá username cũ (nếu có) để tài khoản này
                    # không đọc nhầm username của tài khoản trước.
                    try:
                        paths.data_file(
                            f"meta_username{suffix}.txt"
                        ).unlink()
                    except OSError:
                        pass
                helper_ok, burned, waitlisted, cookie, own_invite = self._run_helper(
                    helper, address, display, label_total, settings.meta
                )
                if helper_ok:
                    saved = False
                    save_error = None
                    with shared_lock:
                        try:
                            # Cột "Mật khẩu hộp thư" của sheet:
                            # bí mật của hộp thư theo dịch vụ —
                            # mật khẩu mail.tm, token tempmail.lol
                            # / Guerrillamail; tempmail.plus và
                            # maildrop.cc để trống vì địa chỉ tự
                            # nó là khoá hộp thư.
                            mail_password = mailbox.get("secret") or ""
                            if settings.meta:
                                username = storage._read_meta_username(suffix)
                                storage._append_meta_row(
                                    address,
                                    mail_password=mail_password,
                                    username=username,
                                )
                            else:
                                storage._append_spreadsheet_row(
                                    address,
                                    mail_password=mail_password,
                                    cookie=cookie,
                                    invite=own_invite,
                                )
                            saved = True
                            shared["success"] += 1
                            if settings.card_tracked:
                                # Tài khoản chỉ được lưu sau khi qua
                                # bước thẻ, nên đây đúng là 1 lượt
                                # dùng thẻ đang lưu.
                                shared["card_used"] += 1
                                storage._save_card_success_count(shared["card_used"])
                        except RuntimeError as error:
                            shared["fail"] += 1
                            save_error = error
                    if saved:
                        if settings.meta:
                            self.events.put(
                                (
                                    "log",
                                    f"[{display}/{label_total}] ✔ Đăng kí Meta thành công {address} "
                                    "— đã lưu vào meta_accounts.xlsx.",
                                )
                            )
                        else:
                            self.events.put(
                                (
                                    "log",
                                    f"[{display}/{label_total}] ✔ Đăng ký thành công {address} "
                                    "— đã lưu vào sheet Excel.",
                                )
                            )
                    else:
                        self.events.put(
                            (
                                "log",
                                f"[{display}/{label_total}] {address} đã đăng ký xong nhưng "
                                f"không ghi được sheet Excel: {save_error}",
                            )
                        )
                else:
                    with shared_lock:
                        shared["fail"] += 1
                    if self.stop_event.is_set():
                        self.events.put(
                            ("log", f"[{display}/{label_total}] Đã dừng theo yêu cầu.")
                        )
                    else:
                        self.events.put(
                            (
                                "log",
                                f"[{display}/{label_total}] ✘ Đăng ký thất bại cho {address} "
                                "— lý do ở dòng RESULT phía trên.",
                            )
                        )
                if helper_ok and own_invite and not settings.single_invite:
                    pending_code = own_invite
                    try:
                        paths.data_file("invite.txt").write_text(
                            own_invite + "\n", encoding="utf-8"
                        )
                    except OSError:
                        pass
                    self.events.put(("invite", own_invite))
                    self.events.put(
                        (
                            "log",
                            f"[{display}/{label_total}] Chuỗi mã mời: tài khoản "
                            f"tiếp theo sẽ quy đổi mã {own_invite} — mã "
                            "của tài khoản này.",
                        )
                    )
                elif (
                    helper_ok
                    and own_invite
                    and settings.single_invite
                    and pending_code
                ):
                    # Chế độ 1 mã mời: mã riêng của tài khoản này
                    # đã nằm trong sheet; KHÔNG nối chuỗi — tài
                    # khoản sau vẫn quy đổi đúng mã hạt giống như
                    # mọi tài khoản khác trong lượt chạy.
                    self.events.put(
                        (
                            "log",
                            f"[{display}/{label_total}] Chế độ 1 mã "
                            f"mời: tài khoản tiếp theo vẫn quy đổi "
                            f"mã {pending_code} (mã riêng của tài "
                            "khoản này chỉ lưu vào sheet).",
                        )
                    )
                if burned and not settings.direct:
                    self.events.put(
                        (
                            "log",
                            f"{lane_tag}Proxy #{proxy_index + 1} bị Meta chặn — "
                            "đổi sang proxy kế từ tài khoản sau.",
                        )
                    )
                if burned or waitlisted:
                    if settings.direct:
                        # IP máy là IP duy nhất của cả lượt chạy:
                        # không có proxy khác để đổi, chạy tiếp
                        # chỉ thất bại nối nhau — dừng ngay để
                        # người dùng đổi server VPN.
                        reason = (
                            "bị Meta chặn tạo tài khoản"
                            if burned
                            else "rơi vào 'Join the waitlist' "
                            "(khu vực không hỗ trợ)"
                        )
                        self.events.put(
                            (
                                "log",
                                f"[{display}/{label_total}] IP máy "
                                f"(VPN) {reason} — dừng lượt chạy. "
                                "Đổi server VPN (ưu tiên US) rồi "
                                "chạy lại.",
                            )
                        )
                        self.stop_event.set()
                    else:
                        # Proxy này hỏng ở cấp IP: Meta chặn tạo tài
                        # khoản, hoặc IP thoát rơi vào khu vực chưa
                        # hỗ trợ (waitlist). Đưa vào kho proxy lỗi để
                        # loại khỏi mọi đợt chạy sau, và loại ngay
                        # trong đợt này (kể cả khi xoay vòng quay lại).
                        reason = (
                            "bị Meta chặn tạo tài khoản"
                            if burned
                            else "bị 'Join the waitlist' (khu vực "
                            "không hỗ trợ)"
                        )
                        if self._mark_proxy_bad(line):
                            self.events.put(
                                (
                                    "log",
                                    f"[{display}/{label_total}] Proxy "
                                    f"{proxy_pool._proxy_display(line)} {reason} — "
                                    "đã đưa vào ô Proxy lỗi, không dùng "
                                    "lại ở các đợt sau.",
                                )
                            )
                        rotation.retire_current()
                else:
                    rotation.advance_after_account(burned=burned)
                if (share == 0 or index < share) and not self.stop_event.is_set():
                    time.sleep(PAUSE_BETWEEN_ACCOUNTS_SECONDS)
            finally:
                with shared_lock:
                    shared["in_flight"] -= 1

    def _run_helper(self, helper, address, index, total, meta=False):
        suffix = getattr(self._lane_local, "suffix", "")
        command = [sys.executable, str(helper), address, "--auto"]
        if meta:
            command.append("--meta")
        if suffix:
            command.append(suffix.lstrip("-"))
        try:
            process = subprocess.Popen(
                command,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
        except OSError as error:
            self.events.put(
                ("log", f"[{index}/{total}] Không mở được trình duyệt: {error}")
            )
            return False, False, False, "", ""
        self.helper_process = process
        with self._process_lock:
            self.helper_processes.add(process)
        signals = {"burn": False, "waitlist": False, "cookie": "", "invite": ""}

        def forward_output():
            for line in process.stdout:
                line = line.strip()
                if line:
                    if proxy_pool._line_burns_proxy(line):
                        signals["burn"] = True
                    if proxy_pool._line_flags_waitlist(line):
                        signals["waitlist"] = True
                    cookie = _session_cookie_from_line(line)
                    if cookie is not None:
                        # Cookie phiên chỉ đi vào sheet, không vào log.
                        signals["cookie"] = cookie
                        continue
                    invite = _invite_from_line(line)
                    if invite is not None:
                        signals["invite"] = invite
                        continue
                    self.events.put(("log", f"[{index}/{total}] {line}"))

        reader = threading.Thread(target=forward_output, daemon=True)
        reader.start()
        deadline = time.monotonic() + HELPER_TIMEOUT_SECONDS
        while process.poll() is None:
            if self.stop_event.is_set():
                process.terminate()
                break
            if time.monotonic() > deadline:
                self.events.put(
                    (
                        "log",
                        f"[{index}/{total}] Quá "
                        f"{HELPER_TIMEOUT_SECONDS // 60} phút cho một tài "
                        "khoản — đóng trình duyệt và bỏ qua.",
                    )
                )
                process.terminate()
                break
            time.sleep(0.3)
        try:
            process.wait(timeout=10)
        except subprocess.TimeoutExpired:
            process.kill()
        reader.join(timeout=2)
        self.helper_process = None
        with self._process_lock:
            self.helper_processes.discard(process)
        return (
            process.returncode == 0 and not self.stop_event.is_set()
        ), signals["burn"], signals["waitlist"], signals["cookie"], signals["invite"]
