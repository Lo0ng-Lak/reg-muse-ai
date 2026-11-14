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
import traceback
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
class CardRunState:
    """Một thẻ trong kho, với giới hạn + số lần dùng của lượt chạy.

    limit = 0 là không giới hạn. used đếm tài khoản thành công
    đã dùng thẻ này (đọc từ kho lúc bắt đầu lượt chạy, ghi lại
    vào kho sau mỗi lần thành công). in_flight là số tài khoản
    đang chạy dở giữ chỗ trên thẻ này — giữ chỗ ngay lúc kiểm
    tra để các luồng không cùng vượt giới hạn.
    """

    card_id: str
    label: str
    limit: int = 0
    used: int = 0
    in_flight: int = 0

    def exhausted(self):
        return (
            self.limit > 0
            and (self.used + self.in_flight) >= self.limit
        )

    def mark_full(self):
        """Meta từ chối thẻ này (quá nhiều tài khoản quảng cáo).

        used >= limit > 0 để exhausted() đúng và lượt sau không
        giữ chỗ thẻ này nữa. Giới hạn 0 được kéo lên ít nhất 1.
        """
        cap = self.limit if self.limit > 0 else max(self.used, 1)
        cap = max(1, min(int(cap), 100))
        self.limit = cap
        if self.used < cap:
            self.used = cap


class _CardVaultRun:
    """Kho thẻ của một lượt chạy, theo thứ tự dùng.

    Mọi phương thức đều phải gọi khi đang giữ khóa chung của
    lượt chạy (các vòng lặp đã giữ khóa ở bước giữ chỗ).
    Chế độ tay (manual=True) chỉ dùng thẻ đầu tiên trong danh
    sách; chế độ tự động dùng thẻ đầu tiên còn giới hạn, theo
    thứ tự kho — hết giới hạn thì tự xoay sang thẻ kế.
    """

    def __init__(self, cards, manual=False):
        self.cards = list(cards)
        self.manual = manual
        self.current = None

    def reserve(self):
        """Giữ chỗ cho tài khoản sắp chạy.

        Trả về (card, switched_from, stop):
        - card: thẻ sẽ dùng cho tài khoản này, hoặc None khi lượt
          chạy không có thẻ trong kho (bước thẻ chờ nhập tay).
        - switched_from: thẻ trước đó khi vừa xoay sang thẻ khác.
        - stop: True khi mọi thẻ (thẻ tay chọn, hoặc toàn kho ở
          chế độ tự động) đều đã đủ giới hạn — lượt chạy phải dừng.
        """
        if not self.cards:
            return None, None, False
        self.refresh_full_from_disk()
        candidates = self.cards[:1] if self.manual else self.cards
        for card in candidates:
            if not card.exhausted():
                switched_from = (
                    self.current
                    if self.current is not None and self.current is not card
                    else None
                )
                self.current = card
                card.in_flight += 1
                return card, switched_from, False
        return None, None, True

    def release(self, card):
        if card is not None:
            card.in_flight = max(0, card.in_flight - 1)

    def refresh_full_from_disk(self):
        """Luồng khác vừa đánh dấu thẻ full thì kho trên đĩa
        đã đổi. Kéo used/limit vào bộ nhớ trước khi giữ chỗ,
        để luồng này không phát lại đúng thẻ đó."""
        try:
            entries = storage._load_vault()
        except Exception:
            return
        by_id = {entry["id"]: entry for entry in entries}
        for item in self.cards:
            entry = by_id.get(item.card_id)
            if not entry:
                continue
            try:
                limit = int(entry.get("limit") or 0)
                used = int(entry.get("used") or 0)
            except (TypeError, ValueError):
                continue
            if limit > 0 and used >= limit:
                item.limit = limit
                if item.used < used:
                    item.used = used

    def mark_card_full(self, card_id):
        """Đánh dấu thẻ full trong bộ nhớ của lượt chạy.

        Chế độ chọn tay chỉ nạp một thẻ: thẻ đó full thì nạp
        phần còn lại của kho và xoay tiếp, để tài khoản sau
        không dừng khi vẫn còn thẻ khác.
        """
        found = False
        for item in self.cards:
            if item.card_id == card_id:
                item.mark_full()
                found = True
                break
        if not found or not self.manual:
            return
        known = {item.card_id for item in self.cards}
        try:
            extras = storage._load_vault()
        except Exception:
            extras = []
        for entry in extras:
            if entry["id"] in known:
                continue
            self.cards.append(
                CardRunState(
                    card_id=entry["id"],
                    label=entry["label"],
                    limit=int(entry.get("limit") or 0),
                    used=int(entry.get("used") or 0),
                )
            )
        self.manual = False

    def take_over_card(self, holder, card_id):
        """Chuyển chỗ đang giữ từ thẻ cũ sang thẻ vừa được nhập.

        Thẻ bị từ chối không được tính một lần dùng. Lượt thành
        công sau đó ghi vào thẻ mới.
        """
        if holder is None or not card_id:
            return
        new = None
        for item in self.cards:
            if item.card_id == card_id:
                new = item
                break
        if new is None:
            try:
                extras = storage._load_vault()
            except Exception:
                extras = []
            for entry in extras:
                if entry["id"] != card_id:
                    continue
                new = CardRunState(
                    card_id=entry["id"],
                    label=entry["label"],
                    limit=int(entry.get("limit") or 0),
                    used=int(entry.get("used") or 0),
                )
                self.cards.append(new)
                break
        if new is None:
            return
        old = holder.get("card")
        if old is not new:
            if old is not None:
                old.in_flight = max(0, old.in_flight - 1)
            new.in_flight += 1
        holder["card"] = new

    def record_success(self, card):
        """Một tài khoản vừa thành công bằng thẻ này: tăng bộ đếm
        và ghi lại vào kho (khóa theo id thẻ, không theo số thẻ)."""
        if card is None:
            return
        card.used += 1
        storage._save_vault_card_success_count(card.card_id, card.used)

    def note_for(self, card):
        """Ghi chú tiến độ cho thẻ đang dùng ('' khi không có thẻ)."""
        if card is None:
            return ""
        limit = card.limit if card.limit > 0 else "∞"
        return f" · Thẻ {card.label}: {card.used}/{limit}"

    def stop_message(self):
        """Thông báo dừng khi mọi thẻ đã đủ giới hạn."""
        if self.manual and self.cards:
            card = self.cards[0]
            return (
                f"Thẻ {card.label} đã đủ giới hạn {card.limit} tài khoản "
                "thành công — dừng lượt chạy. Chọn thẻ khác ở kho thẻ "
                "(hoặc tăng giới hạn) rồi chạy lại."
            )
        return (
            "Tất cả thẻ trong kho đã đủ giới hạn — dừng lượt chạy. "
            "Tăng giới hạn từng thẻ hoặc nhập thêm thẻ mới vào kho "
            "thẻ rồi chạy lại."
        )


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
    # Đăng nhập Muse bằng tài khoản Meta đã đăng ký (đọc từ file
    # Excel người dùng chọn ở tab Meta, không tạo hộp thư mới).
    meta_login: bool = False
    meta_excel_path: str = ""
    # Đăng nhập lại mail Muse đã có (file Excel cột Email
    # Muse.ai): không proxy, không thẻ, chỉ OTP + cookie ra file
    # xuất gồm hai cột Email và Cookie đăng nhập.
    relogin: bool = False
    relogin_excel_path: str = ""
    relogin_output_path: str = ""
    invite: str = ""
    single_invite: bool = False
    # Kho thẻ của lượt chạy: list[CardRunState] theo thứ tự kho.
    # Rỗng là không có thẻ lưu — bước thẻ chờ nhập tay, không
    # tính giới hạn. manual_card=True là chỉ dùng thẻ đầu tiên
    # trong danh sách (thẻ đang chọn tay ở kho); False là tự
    # động xoay sang thẻ kế khi thẻ đang dùng đủ giới hạn.
    cards: list = field(default_factory=list)
    manual_card: bool = False
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
        try:
            threading.Thread(
                target=self._run_loop, args=(settings,), daemon=True
            ).start()
        except Exception:
            # Không tạo được luồng: trả running về False để lần bấm
            # sau không bị chặn lặng ở `if self.coordinator.running`.
            self.running = False
            raise
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
        """Điểm vào của luồng chạy: bọc vòng chạy thật để running
        luôn được reset và sự kiện "finished" luôn được phát, kể cả
        khi vòng chạy chết vì lỗi nội bộ.

        Trước đây running chỉ reset ở cuối _run_loop /
        _run_login_loop, không có try/finally: bất kỳ exception nào
        thoát khỏi _run_lane (luồng đơn chạy trực tiếp trong luồng
        này) cũng giết chết luồng chạy lặng lẽ, running kẹt True
        mãi, không có "finished", giao diện khoá cứng và mọi lần
        bấm sau đều rơi vào `if self.coordinator.running: return`.
        """
        success = fail = 0
        try:
            if settings.relogin:
                success, fail = self._run_relogin_loop(settings)
            elif settings.meta_login:
                success, fail = self._run_login_loop(settings)
            else:
                success, fail = self._run_registration_loop(settings)
        except Exception as error:
            try:
                detail = "".join(
                    traceback.format_exception(error)
                ).rstrip()
            except Exception:
                detail = f"{type(error).__name__}: {error}"
            self.events.put(
                (
                    "log",
                    "⛔ Lượt chạy dừng đột ngột do lỗi nội bộ — "
                    "bấm lại để chạy tiếp (đã ghi lỗi để sửa):\n"
                    f"{detail}",
                )
            )
        finally:
            self.running = False
            self.events.put(
                (
                    "finished",
                    (
                        success,
                        fail,
                        self.stop_event.is_set(),
                    ),
                )
            )

    def _run_registration_loop(self, settings):
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
        # sheet. Thẻ thì các luồng dùng CHUNG kho thẻ: thẻ đang
        # chọn tay thì mọi luồng dùng đúng thẻ đó, còn tự động thì
        # các luồng xoay chung qua các thẻ trong kho theo thứ tự;
        # bộ đếm giới hạn từng thẻ tính gộp toàn cục (đếm cả số
        # tài khoản đang chạy dở để không vượt giới hạn). Ở chế độ
        # IP máy (direct), pool chỉ có câu đánh dấu DIRECT: mọi
        # luồng đều dùng chung 1 IP của máy — số luồng chạy đúng
        # như đã đặt (chủ tool chấp nhận rủi ro IP bị chặn); mỗi
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
            "vault": _CardVaultRun(settings.cards, settings.manual_card),
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
        # Trả số đếm về cho _run_loop: chính nó reset running và
        # phát "finished" (luôn luôn, kể cả khi vòng này chết giữa
        # chừng vì lỗi nội bộ).
        return shared["success"], shared["fail"]

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
            card = None
            with shared_lock:
                vault = shared["vault"]
                # Giữ chỗ ngay trong cùng một khóa với phép kiểm
                # tra: hai luồng không thể cùng thấy "còn chỗ"
                # rồi cùng chạy vượt giới hạn thẻ. Khi thẻ đang
                # dùng đủ giới hạn, reserve() tự xoay sang thẻ
                # kế còn giới hạn; hết kho thì trả stop=True.
                # Chỗ được trả lại ở finally cuối vòng lặp.
                card, switched_from, card_hit = vault.reserve()
                log_card_stop = (
                    card_hit and not shared["card_stop_logged"]
                )
                if log_card_stop:
                    shared["card_stop_logged"] = True
            if card_hit:
                if log_card_stop:
                    self.events.put(("log", vault.stop_message()))
                break
            if switched_from is not None:
                self.events.put(
                    (
                        "log",
                        f"{lane_tag}Thẻ {switched_from.label} đã đủ "
                        f"giới hạn — tự động chuyển sang thẻ "
                        f"{card.label}.",
                    )
                )
            try:
                card_holder = {"card": card}
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
                card_note = shared["vault"].note_for(card)
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
                helper_ok, burned, waitlisted, load_failed, cookie, own_invite, card_skipped, facebook = self._run_helper(
                    helper,
                    address,
                    display,
                    label_total,
                    settings.meta,
                    card_id=card.card_id if card is not None else "",
                    card_holder=card_holder,
                    vault=shared["vault"],
                    vault_lock=shared_lock,
                )
                card = card_holder["card"]
                if settings.meta and not load_failed:
                    # Mở được meta.ai (dù tài khoản này thành hay
                    # bại ở bước sau): proxy đã chứng minh còn tải
                    # được trang — xoá strike cũ, strike chỉ tính
                    # các lần hỏng LIÊN TIẾP.
                    rotation.clear_load_fail_strikes(line)
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
                                otp_link = mail_providers.inbox_url(
                                    settings.provider_id, address,
                                    proxy=line,
                                )
                                storage._append_meta_row(
                                    address,
                                    mail_password=mail_password,
                                    username=username,
                                    provider=settings.provider_id,
                                    otp_link=otp_link,
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
                            # Tài khoản chỉ được lưu sau khi qua bước
                            # thẻ, nên đây đúng là 1 lượt dùng thẻ
                            # đang giữ chỗ. Không có thẻ trong kho
                            # (nhập tay ở bước thẻ) thì không đếm.
                            # Ngoại lệ: tài khoản đã đăng nhập sẵn
                            # qua màn liên kết nên bỏ qua bước thẻ
                            # (CARD_SKIPPED) — thẻ chưa hề dùng,
                            # không được tính lượt.
                            if not card_skipped:
                                shared["vault"].record_success(card)
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
                if facebook and not settings.direct:
                    self.events.put(
                        (
                            "log",
                            f"{lane_tag}Proxy #{proxy_index + 1} đẩy sang "
                            "Facebook — bỏ proxy này, đổi sang proxy kế.",
                        )
                    )
                elif burned and not settings.direct:
                    self.events.put(
                        (
                            "log",
                            f"{lane_tag}Proxy #{proxy_index + 1} bị Meta chặn — "
                            "đổi sang proxy kế từ tài khoản sau.",
                        )
                    )
                # Strike lỗi tải trang (2026-11-13): một lượt chạy
                # helper mở meta.ai lỗi hết 5 lần thử = 1 strike.
                # Đủ 2 strike (2 lượt chạy liên tiếp đều hỏng) mới
                # bỏ proxy; strike 1 thì lần sau vẫn dùng đúng
                # proxy này (không retire, không xoay theo hạn
                # mức) để lấy strike 2 nếu vẫn hỏng.
                load_fail_strikes = 0
                if load_failed and not settings.direct:
                    load_fail_strikes = rotation.record_load_fail_strike(line)
                    if load_fail_strikes >= 2:
                        strike_note = "hỏng lần 2/2, bỏ proxy này."
                    else:
                        strike_note = (
                            "hỏng lần 1/2, lần sau vẫn dùng proxy này."
                        )
                    self.events.put(
                        (
                            "log",
                            f"{lane_tag}Proxy #{proxy_index + 1} "
                            f"({proxy_pool._proxy_display(line)}) mở meta.ai "
                            "lỗi tải trang hết 5 lần thử — "
                            f"{strike_note}",
                        )
                    )
                drop_proxy = facebook or burned or waitlisted or (
                    load_failed and load_fail_strikes >= 2
                )
                hold_proxy = (
                    load_failed
                    and not settings.direct
                    and load_fail_strikes < 2
                    and not facebook
                )
                if drop_proxy or (settings.direct and (load_failed or facebook)):
                    if settings.direct:
                        # IP máy là IP duy nhất của cả lượt chạy:
                        # không có proxy khác để đổi, chạy tiếp
                        # chỉ thất bại nối nhau — dừng ngay để
                        # người dùng đổi server VPN.
                        reason = (
                            "đẩy sang Facebook"
                            if facebook
                            else "bị Meta chặn tạo tài khoản"
                            if burned
                            else "rơi vào 'Join the waitlist' "
                            "(khu vực không hỗ trợ)"
                            if waitlisted
                            else "mở meta.ai lỗi tải trang (hết 5 lần thử)"
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
                        # hỗ trợ (waitlist), hoặc không tải nổi
                        # meta.ai hết 5 lần thử ở 2 lượt chạy liên
                        # tiếp. Đưa vào kho proxy lỗi để loại khỏi
                        # mọi đợt chạy sau, và loại ngay trong đợt
                        # này (kể cả khi xoay vòng quay lại).
                        reason = (
                            "đẩy sang Facebook"
                            if facebook
                            else "bị Meta chặn tạo tài khoản"
                            if burned
                            else "bị 'Join the waitlist' (khu vực "
                            "không hỗ trợ)"
                            if waitlisted
                            else "mở meta.ai lỗi tải trang hết 5 lần thử "
                            "(hỏng lần 2/2)"
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
                elif not hold_proxy:
                    rotation.advance_after_account(burned=burned)
                if (share == 0 or index < share) and not self.stop_event.is_set():
                    time.sleep(PAUSE_BETWEEN_ACCOUNTS_SECONDS)
            finally:
                with shared_lock:
                    shared["vault"].release(card_holder["card"])

    def _run_login_loop(self, settings):
        """Lượt đăng nhập Muse bằng tài khoản Meta đã đăng ký.

        Đọc các dòng chưa đăng nhập trong file Excel người dùng
        chọn ở tab Meta (mỗi dòng: email + bí mật hộp thư + dịch
        vụ), chia theo bước nhảy cho các luồng như kho proxy.
        Mỗi tài khoản chạy helper --muse-login: đăng nhập Muse
        bằng OTP rồi nộp thẻ đã lưu. Thành công thì đánh dấu
        dòng đó trong file Excel + ghi dòng vào accounts.xlsx.
        total = 0 nghĩa là đăng nhập HẾT các dòng chưa đăng nhập.
        """
        rows = storage._pending_meta_logins(settings.meta_excel_path)
        if settings.total > 0:
            rows = rows[: settings.total]
        if not rows:
            self.events.put(
                (
                    "log",
                    "Không còn tài khoản Meta nào chưa đăng nhập Muse "
                    "trong file Excel đã chọn.",
                )
            )
            return 0, 0
        pool = settings.pool
        direct = settings.direct
        requested_lanes = max(1, min(int(settings.lanes), 5))
        lanes = max(1, min(requested_lanes, len(rows)))
        if direct:
            self.report_ip(proxy_pool._check_public_ip())
        shared = {
            "success": 0,
            "fail": 0,
            "vault": _CardVaultRun(settings.cards, settings.manual_card),
            "card_stop_logged": False,
        }
        shared_lock = threading.Lock()
        self.events.put(
            (
                "log",
                f"Đăng nhập Muse cho {len(rows)} tài khoản Meta "
                f"({lanes} luồng) — mỗi tài khoản nộp thẻ đã lưu.",
            )
        )
        threads = []
        for lane_no in range(1, lanes + 1):
            if direct:
                subset = [DIRECT_PROXY_LINE]
            else:
                subset = pool[lane_no - 1::lanes]
            lane_rows = rows[lane_no - 1::lanes]
            lane_args = (
                settings,
                lane_no,
                lanes,
                lane_rows,
                subset,
                shared,
                shared_lock,
                settings.invite,
            )
            if lanes == 1:
                self._run_login_lane(*lane_args)
            else:
                thread = threading.Thread(
                    target=self._run_login_lane,
                    args=lane_args,
                    name=f"login-lane-{lane_no}",
                    daemon=True,
                )
                thread.start()
                threads.append(thread)
        for thread in threads:
            thread.join()
        # Trả số đếm về cho _run_loop: chính nó reset running và
        # phát "finished" (luôn luôn, kể cả khi vòng này chết giữa
        # chừng vì lỗi nội bộ).
        return shared["success"], shared["fail"]

    def _run_login_lane(
        self, settings, lane_no, lanes, rows, subset, shared, shared_lock,
        seed="",
    ):
        """Một luồng đăng nhập: lần lượt các dòng Meta của luồng.

        Cùng nhịp xoay proxy / giữ chỗ thẻ / nghỉ giữa các tài
        khoản như _run_lane, nhưng hộp thư lấy từ dòng Excel
        (không tạo mới). Có mã mời của tài khoản trước thì nhập
        mã đó, rồi lấy mã của tài khoản này cho tài khoản sau.
        Thành công thì đánh dấu dòng đó đã đăng nhập trong file Excel.
        """
        helper = paths.SRC_DIR / "browser_helper.py"
        suffix = "" if lane_no == 1 else f"-{lane_no}"
        self._lane_local.suffix = suffix
        multi = lanes > 1
        lane_tag = f"[Luồng {lane_no}] " if multi else ""
        rotation = proxy_pool._ProxyRotation(subset, settings.per_proxy)
        written_proxy_index = None
        label_total = len(rows)
        # Chuỗi mã mời của luồng này, giống đăng ký: tài khoản
        # sắp chạy nhập pending_code (mã tài khoản trước). Mã nó
        # lấy được thành mã cho tài khoản sau. Chế độ 1 mã thì
        # mọi tài khoản nhập đúng mã hạt giống.
        pending_code = (seed or "").strip().upper()
        for index, row in enumerate(rows, start=1):
            if self.stop_event.is_set():
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
            card = None
            with shared_lock:
                vault = shared["vault"]
                # Giữ chỗ ngay trong cùng một khóa với phép kiểm
                # tra: hai luồng không thể cùng thấy "còn chỗ"
                # rồi cùng chạy vượt giới hạn thẻ. Khi thẻ đang
                # dùng đủ giới hạn, reserve() tự xoay sang thẻ
                # kế còn giới hạn; hết kho thì trả stop=True.
                # Chỗ được trả lại ở finally cuối vòng lặp.
                card, switched_from, card_hit = vault.reserve()
                log_card_stop = (
                    card_hit and not shared["card_stop_logged"]
                )
                if log_card_stop:
                    shared["card_stop_logged"] = True
            if card_hit:
                if log_card_stop:
                    self.events.put(("log", vault.stop_message()))
                break
            if switched_from is not None:
                self.events.put(
                    (
                        "log",
                        f"{lane_tag}Thẻ {switched_from.label} đã đủ "
                        f"giới hạn — tự động chuyển sang thẻ "
                        f"{card.label}.",
                    )
                )
            try:
                card_holder = {"card": card}
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
                card_note = shared["vault"].note_for(card)
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
                address = row["email"]
                mailbox = {
                    "provider": row["provider"],
                    "email": address,
                    "secret": row["secret"],
                }
                self.events.put(
                    (
                        "log",
                        f"[{display}/{label_total}] Đăng nhập Muse bằng "
                        f"tài khoản Meta {address} ({row['provider']}).",
                    )
                )
                try:
                    storage._write_mailbox(mailbox, suffix)
                except OSError:
                    pass
                try:
                    storage._write_redeem_code(pending_code, suffix)
                except OSError:
                    pass
                # Đánh dấu ĐANG đăng nhập TRƯỚC khi chạy helper, dưới
                # khóa chung — luồng khác (hay lượt chạy sau) không
                # lấy trùng mail này. Ghi không được thì vẫn chạy
                # tiếp, chỉ ghi log.
                with shared_lock:
                    try:
                        storage._mark_meta_logging_in(
                            settings.meta_excel_path, address
                        )
                    except RuntimeError as error:
                        self.events.put(
                            ("log", f"[{display}/{label_total}] {error}")
                        )
                helper_ok, burned, waitlisted, load_failed, cookie, own_invite, card_skipped, facebook = (
                    self._run_helper(
                        helper, address, display, label_total,
                        muse_login=True,
                        card_id=card.card_id if card is not None else "",
                        card_holder=card_holder,
                        vault=shared["vault"],
                        vault_lock=shared_lock,
                    )
                )
                card = card_holder["card"]
                if not load_failed:
                    # Đăng nhập Muse luôn mở meta.ai: mở được thì
                    # proxy còn tốt — xoá strike cũ, strike chỉ
                    # tính các lần hỏng LIÊN TIẾP.
                    rotation.clear_load_fail_strikes(line)
                if helper_ok:
                    saved = False
                    save_error = None
                    mark_error = None
                    with shared_lock:
                        try:
                            storage._append_spreadsheet_row(
                                address,
                                row["secret"],
                                cookie=cookie,
                                invite=own_invite,
                                otp_link=mail_providers.inbox_url(
                                    row["provider"], address, proxy=line
                                ),
                            )
                            saved = True
                            shared["success"] += 1
                            # Tài khoản chỉ được lưu sau khi qua bước
                            # thẻ, nên đây đúng là 1 lượt dùng thẻ
                            # đang giữ chỗ. Không có thẻ trong kho
                            # (nhập tay ở bước thẻ) thì không đếm.
                            # Ngoại lệ: tài khoản đã đăng nhập sẵn
                            # qua màn liên kết nên bỏ qua bước thẻ
                            # (CARD_SKIPPED) — thẻ chưa hề dùng,
                            # không được tính lượt.
                            if not card_skipped:
                                shared["vault"].record_success(card)
                        except RuntimeError as error:
                            shared["fail"] += 1
                            save_error = error
                        if saved:
                            # Đánh dấu DƯỚI KHÓA CHUNG: ghi sheet là
                            # đọc-cả-file rồi ghi-lại, hai luồng cùng
                            # ghi là luồng sau đè mất dấu của luồng
                            # trước, lượt sau lấy trùng mail (live
                            # 2026-11-13: 5 luồng chạy song song).
                            try:
                                storage._mark_meta_logged_in(
                                    settings.meta_excel_path, address
                                )
                            except RuntimeError as error:
                                mark_error = error
                    if saved:
                        if mark_error is not None:
                            self.events.put(
                                (
                                    "log",
                                    f"[{display}/{label_total}] {mark_error}",
                                )
                            )
                        self.events.put(
                            (
                                "log",
                                f"[{display}/{label_total}] ✔ Đăng nhập Muse thành công "
                                f"{address} — "
                                + (
                                    "đã đăng nhập sẵn qua màn liên kết, "
                                    "bỏ qua bước thẻ, đã lưu vào sheet Excel."
                                    if card_skipped
                                    else "đã nộp thẻ và lưu vào sheet Excel."
                                ),
                            )
                        )
                    else:
                        self.events.put(
                            (
                                "log",
                                f"[{display}/{label_total}] {address} đã đăng nhập xong nhưng "
                                f"không ghi được sheet Excel: {save_error}",
                            )
                        )
                    if own_invite and not settings.single_invite:
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
                                f"[{display}/{label_total}] Chuỗi mã mời: "
                                f"tài khoản tiếp theo sẽ nhập mã "
                                f"{own_invite}.",
                            )
                        )
                    elif own_invite and settings.single_invite and pending_code:
                        self.events.put(
                            (
                                "log",
                                f"[{display}/{label_total}] Chế độ 1 mã "
                                f"mời: tài khoản tiếp theo vẫn nhập "
                                f"mã {pending_code}.",
                            )
                        )
                else:
                    with shared_lock:
                        shared["fail"] += 1
                        # Đăng nhập thất bại: xoá dấu đang-đăng-nhập
                        # để lượt sau lấy lại mail này được ngay.
                        try:
                            storage._clear_meta_login_mark(
                                settings.meta_excel_path, address
                            )
                        except RuntimeError as error:
                            self.events.put(
                                ("log", f"[{display}/{label_total}] {error}")
                            )
                    if self.stop_event.is_set():
                        self.events.put(
                            ("log", f"[{display}/{label_total}] Đã dừng theo yêu cầu.")
                        )
                    else:
                        self.events.put(
                            (
                                "log",
                                f"[{display}/{label_total}] ✘ Đăng nhập thất bại cho {address} "
                                "— lý do ở dòng RESULT phía trên.",
                            )
                        )
                if facebook and not settings.direct:
                    self.events.put(
                        (
                            "log",
                            f"{lane_tag}Proxy #{proxy_index + 1} đẩy sang "
                            "Facebook — bỏ proxy này, đổi sang proxy kế.",
                        )
                    )
                elif burned and not settings.direct:
                    self.events.put(
                        (
                            "log",
                            f"{lane_tag}Proxy #{proxy_index + 1} bị Meta chặn — "
                            "đổi sang proxy kế từ tài khoản sau.",
                        )
                    )
                # Strike lỗi tải trang (2026-11-13): một lượt chạy
                # helper mở meta.ai lỗi hết 5 lần thử = 1 strike.
                # Đủ 2 strike (2 lượt chạy liên tiếp đều hỏng) mới
                # bỏ proxy; strike 1 thì lần sau vẫn dùng đúng
                # proxy này (không retire, không xoay theo hạn
                # mức) để lấy strike 2 nếu vẫn hỏng.
                load_fail_strikes = 0
                if load_failed and not settings.direct:
                    load_fail_strikes = rotation.record_load_fail_strike(line)
                    if load_fail_strikes >= 2:
                        strike_note = "hỏng lần 2/2, bỏ proxy này."
                    else:
                        strike_note = (
                            "hỏng lần 1/2, lần sau vẫn dùng proxy này."
                        )
                    self.events.put(
                        (
                            "log",
                            f"{lane_tag}Proxy #{proxy_index + 1} "
                            f"({proxy_pool._proxy_display(line)}) mở meta.ai "
                            "lỗi tải trang hết 5 lần thử — "
                            f"{strike_note}",
                        )
                    )
                drop_proxy = facebook or burned or waitlisted or (
                    load_failed and load_fail_strikes >= 2
                )
                hold_proxy = (
                    load_failed
                    and not settings.direct
                    and load_fail_strikes < 2
                    and not facebook
                )
                if drop_proxy or (settings.direct and (load_failed or facebook)):
                    if settings.direct:
                        reason = (
                            "đẩy sang Facebook"
                            if facebook
                            else "bị Meta chặn đăng nhập"
                            if burned
                            else "rơi vào 'Join the waitlist' "
                            "(khu vực không hỗ trợ)"
                            if waitlisted
                            else "mở meta.ai lỗi tải trang (hết 5 lần thử)"
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
                        reason = (
                            "đẩy sang Facebook"
                            if facebook
                            else "bị Meta chặn đăng nhập"
                            if burned
                            else "bị 'Join the waitlist' (khu vực "
                            "không hỗ trợ)"
                            if waitlisted
                            else "mở meta.ai lỗi tải trang hết 5 lần thử "
                            "(hỏng lần 2/2)"
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
                elif not hold_proxy:
                    rotation.advance_after_account(burned=burned)
                if index < len(rows) and not self.stop_event.is_set():
                    time.sleep(PAUSE_BETWEEN_ACCOUNTS_SECONDS)
            finally:
                with shared_lock:
                    shared["vault"].release(card_holder["card"])

    def _run_relogin_loop(self, settings):
        """Đăng nhập lại từng mail trong file Excel, lấy cookie mới.

        Không proxy (ghi DIRECT), không thẻ, không mã mời. Tối đa
        10 luồng: mỗi luồng một file hộp thư riêng, OTP đọc từ
        đúng mail đó. Có cookie thì ghi một dòng Email + Cookie
        vào file xuất. Không cookie thì bỏ qua mail đó.
        """
        rows = storage._relogin_accounts(settings.relogin_excel_path)
        done = storage._relogin_done_emails(settings.relogin_output_path)
        pending = [
            row for row in rows if row["email"].lower() not in done
        ]
        skipped = len(rows) - len(pending)
        if skipped:
            self.events.put(
                (
                    "log",
                    f"Bỏ qua {skipped} mail đã có cookie trong file "
                    f"xuất. Còn {len(pending)} mail chưa lấy.",
                )
            )
        rows = pending
        if settings.total:
            rows = rows[: max(0, int(settings.total))]
        if not rows:
            self.events.put(
                (
                    "log",
                    "Không còn mail nào chưa lấy cookie."
                    if skipped
                    else "Không có mail nào để đăng nhập lại.",
                )
            )
            return 0, 0
        requested = max(1, min(int(settings.lanes or 1), 10))
        lanes = max(1, min(requested, len(rows)))
        shared = {"success": 0, "fail": 0}
        shared_lock = threading.Lock()
        self.events.put(
            (
                "log",
                f"Đăng nhập lại {len(rows)} mail ({lanes} luồng).",
            )
        )
        threads = []
        for lane_no in range(1, lanes + 1):
            lane_rows = rows[lane_no - 1::lanes]
            lane_args = (
                settings, lane_no, lanes, lane_rows, shared, shared_lock
            )
            if lanes == 1:
                self._run_relogin_lane(*lane_args)
            else:
                thread = threading.Thread(
                    target=self._run_relogin_lane,
                    args=lane_args,
                    name=f"relogin-lane-{lane_no}",
                    daemon=True,
                )
                thread.start()
                threads.append(thread)
        for thread in threads:
            thread.join()
        return shared["success"], shared["fail"]

    def _run_relogin_lane(
        self, settings, lane_no, lanes, rows, shared, shared_lock
    ):
        """Một luồng đăng nhập lại: lần lượt các mail được chia cho luồng."""
        helper = paths.SRC_DIR / "browser_helper.py"
        suffix = "" if lane_no == 1 else f"-{lane_no}"
        self._lane_local.suffix = suffix
        multi = lanes > 1
        lane_tag = f"[Luồng {lane_no}] " if multi else ""
        label_total = len(rows)
        try:
            storage._write_active_proxy("DIRECT", suffix)
        except OSError as error:
            self.events.put(
                (
                    "log",
                    f"{lane_tag}Không ghi được proxy{suffix}.txt: {error}",
                )
            )
        for index, row in enumerate(rows, start=1):
            if self.stop_event.is_set():
                break
            address = row["email"]
            display = f"L{lane_no} · {index}" if multi else index
            with shared_lock:
                s_now = shared["success"]
                f_now = shared["fail"]
            self.events.put(
                (
                    "progress",
                    (
                        lane_no,
                        index,
                        label_total,
                        s_now,
                        f_now,
                        "",
                        multi,
                    ),
                )
            )
            self.events.put(
                (
                    "log",
                    f"{lane_tag}[{display}/{label_total}] "
                    f"Đăng nhập lại {address}.",
                )
            )
            if row["provider"] == "mail.tm" and not row["secret"]:
                with shared_lock:
                    shared["fail"] += 1
                self.events.put(
                    (
                        "log",
                        f"{lane_tag}[{display}/{label_total}] {address} "
                        "thiếu mật khẩu hộp thư — không đọc được OTP, "
                        "bỏ qua.",
                    )
                )
                continue
            try:
                storage._write_mailbox(
                    {
                        "provider": row["provider"],
                        "email": address,
                        "secret": row["secret"],
                    },
                    suffix,
                )
            except OSError as error:
                with shared_lock:
                    shared["fail"] += 1
                self.events.put(
                    (
                        "log",
                        f"{lane_tag}[{display}/{label_total}] Không ghi "
                        f"được hộp thư cho {address}: {error}",
                    )
                )
                continue
            (
                helper_ok,
                _burned,
                _waitlisted,
                _load_failed,
                cookie,
                _invite,
                _card_skipped,
                _facebook,
            ) = self._run_helper(
                helper, address, index, label_total, relogin=True
            )
            if helper_ok and cookie:
                try:
                    with shared_lock:
                        storage._write_relogin_cookie(
                            settings.relogin_output_path, address, cookie
                        )
                        shared["success"] += 1
                except RuntimeError as error:
                    with shared_lock:
                        shared["fail"] += 1
                    self.events.put(
                        (
                            "log",
                            f"{lane_tag}[{display}/{label_total}] {error}",
                        )
                    )
                    continue
                self.events.put(
                    (
                        "log",
                        f"{lane_tag}[{display}/{label_total}] Đã lấy "
                        f"cookie của {address} và ghi vào file xuất.",
                    )
                )
            else:
                with shared_lock:
                    shared["fail"] += 1
                self.events.put(
                    (
                        "log",
                        f"{lane_tag}[{display}/{label_total}] Bỏ "
                        f"{address}, chuyển sang mail kế.",
                    )
                )
            if index < label_total and not self.stop_event.is_set():
                time.sleep(PAUSE_BETWEEN_ACCOUNTS_SECONDS)

    def _run_helper(self, helper, address, index, total, meta=False,
                      muse_login=False, card_id="", card_holder=None,
                      vault=None, vault_lock=None, relogin=False):
        suffix = getattr(self._lane_local, "suffix", "")
        command = [sys.executable, str(helper), address, "--auto"]
        if meta:
            command.append("--meta")
        if muse_login:
            command.append("--muse-login")
        if relogin:
            command.append("--relogin")
        if card_id:
            # Thẻ kho đã giữ chỗ cho tài khoản này: helper chỉ
            # đọc đúng thẻ đó (id kho, không phải số thẻ).
            command.extend(["--card", card_id])
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
            return False, False, False, False, "", "", False, False
        self.helper_process = process
        with self._process_lock:
            self.helper_processes.add(process)
        signals = {
            "burn": False,
            "waitlist": False,
            "load_fail": False,
            "facebook": False,
            "cookie": "",
            "invite": "",
            "card_skipped": False,
        }

        def forward_output():
            for line in process.stdout:
                line = line.strip()
                if line:
                    if proxy_pool._line_burns_proxy(line):
                        signals["burn"] = True
                    if proxy_pool._line_flags_waitlist(line):
                        signals["waitlist"] = True
                    if proxy_pool._line_flags_load_fail(line):
                        signals["load_fail"] = True
                    if proxy_pool._line_flags_facebook(line):
                        signals["facebook"] = True
                    cookie = _session_cookie_from_line(line)
                    if cookie is not None:
                        # Cookie phiên chỉ đi vào sheet, không vào log.
                        signals["cookie"] = cookie
                        continue
                    invite = _invite_from_line(line)
                    if invite is not None:
                        signals["invite"] = invite
                        continue
                    if line == "CARD_SKIPPED":
                        # Tài khoản đã đăng nhập sẵn qua màn liên
                        # kết nên bỏ qua bước thẻ — thẻ giữ chỗ
                        # cho tài khoản này KHÔNG bị tính lượt
                        # dùng (xem hai chỗ gọi _run_helper).
                        signals["card_skipped"] = True
                        continue
                    if line.startswith("CARD_FULL:"):
                        # Thẻ bị "too many ad accounts": helper đã
                        # ghi full xuống kho. Cập nhật bộ nhớ ngay
                        # để luồng khác không giữ chỗ thẻ này.
                        full_id = line.split(":", 1)[1].strip()
                        if (
                            vault is not None
                            and vault_lock is not None
                            and full_id
                        ):
                            with vault_lock:
                                vault.mark_card_full(full_id)
                        continue
                    if line.startswith("CARD_SWITCH:"):
                        # Chỗ đang giữ chuyển sang thẻ vừa nhập,
                        # để lần thành công tính cho thẻ đó.
                        switch_id = line.split(":", 1)[1].strip()
                        if (
                            vault is not None
                            and vault_lock is not None
                            and switch_id
                        ):
                            with vault_lock:
                                vault.take_over_card(
                                    card_holder, switch_id
                                )
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
        ), signals["burn"], signals["waitlist"], signals["load_fail"], signals["cookie"], signals["invite"], signals["card_skipped"], signals["facebook"]
