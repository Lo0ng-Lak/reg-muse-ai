"""Lượt đăng ký / đăng nhập Meta và Muse cho giao diện web.

Cùng coordinator và cùng file kho với cửa sổ desktop: proxy,
thẻ, mã mời, CapSolver, selfie, file Excel Meta. Web chỉ gửi
nhãn thẻ và số lần dùng — không gửi số thẻ.
"""

import queue
import threading
from datetime import datetime
from pathlib import Path

import mail_providers
import paths
import proxy_pool
import storage
from coordinator import (
    DIRECT_PROXY_LINE,
    CardRunState,
    Coordinator,
    RunSettings,
)


class ConsoleError(Exception):
    """Setup chưa hợp lệ — web trả câu này cho người dùng."""


class WebConsole:
    """Một coordinator + log cho trang Đăng ký."""

    def __init__(self, coordinator=None):
        self.coordinator = coordinator or Coordinator()
        self._owns_events = coordinator is None
        self.log = []
        self.progress = "Sẵn sàng."
        self.invite = ""
        self._lock = threading.Lock()
        if self._owns_events:
            threading.Thread(
                target=self._drain, name="web-console", daemon=True
            ).start()

    def _log(self, text):
        stamp = datetime.now().strftime("%H:%M:%S")
        with self._lock:
            self.log.append(f"[{stamp}] {text}")
            self.log = self.log[-400:]

    def _drain(self):
        while True:
            try:
                kind, payload = self.coordinator.events.get(timeout=0.4)
            except queue.Empty:
                continue
            try:
                if kind == "log":
                    self._log(str(payload))
                elif kind == "progress":
                    self.progress = _format_progress(payload)
                elif kind == "finished":
                    success, fail, stopped = payload
                    note = "Đã dừng. " if stopped else ""
                    self.progress = (
                        f"{note}Hoàn tất: {success} thành công, "
                        f"{fail} thất bại."
                    )
                    self._log(
                        f"--- Kết thúc: {success} thành công, "
                        f"{fail} thất bại ---"
                    )
                elif kind == "bad_proxy":
                    self._log(
                        "Proxy bị chặn — đã đưa vào kho lỗi: "
                        + proxy_pool._proxy_display(payload)
                    )
                elif kind == "invite":
                    self.invite = str(payload or "")
                    self._log(f"Mã mời: {self.invite}")
            except Exception as error:
                self._log(
                    f"Lỗi xử lý sự kiện: {type(error).__name__}: {error}"
                )

    def snapshot(self, since=0):
        since = max(0, int(since or 0))
        with self._lock:
            lines = self.log[since:]
            next_since = since + len(lines)
            progress = self.progress
            invite = self.invite or _read_text(paths.data_file("invite.txt")).strip()
        active = storage._load_vault_active()
        try:
            accounts = len(storage._relogin_accounts())
        except Exception:
            accounts = 0
        return {
            "running": bool(self.coordinator.running),
            "progress": progress,
            "since": next_since,
            "log": lines,
            "invite": invite,
            "cards": storage._load_vault(),
            "active_card": active or "",
            "picked": storage._load_vault_picked(),
            "muse": _pool_view("proxies.txt", "bad-proxies.txt"),
            "meta": _pool_view("proxies_meta.txt", "bad-proxies-meta.txt"),
            "providers": [
                {"id": pid, "label": label}
                for pid, label in mail_providers.PROVIDERS
            ],
            "capsolver_set": bool(storage._load_capsolver_key()),
            "selfie": storage._load_selfie_path_text(),
            "meta_excel": storage._load_meta_excel_path_text(),
            "relogin_output": storage._load_relogin_output_path(),
            "accounts": accounts,
        }

    def cards(self, action, data):
        if self.coordinator.running and action != "list":
            raise ConsoleError("Đang có lượt chạy — đợi xong rồi sửa kho thẻ.")
        if action == "import":
            added, errors, skipped = storage._import_bulk_cards(
                data.get("text") or ""
            )
            self._log(
                f"Nhập kho thẻ: thêm {len(added)}, bỏ qua {skipped}, "
                f"lỗi {len(errors)} dòng."
            )
            return {
                "added": len(added),
                "errors": [
                    {"line": line, "reason": reason}
                    for line, reason in errors
                ],
                "skipped": skipped,
            }
        if action == "select":
            card_id = (data.get("id") or "").strip()
            storage._save_vault_picked([card_id] if card_id else [])
            self._log("Đã chọn một thẻ cho lượt chạy sau.")
            return {"ok": True}
        if action == "pick":
            ids = _card_ids(data)
            if not ids:
                raise ConsoleError("Chọn ít nhất một thẻ.")
            storage._save_vault_picked(ids)
            self._log(f"Lượt chạy sau chỉ dùng {len(ids)} thẻ đã chọn.")
            return {"ok": True, "count": len(ids)}
        if action == "auto":
            storage._save_vault_picked([])
            self._log("Kho thẻ tự xoay theo thứ tự.")
            return {"ok": True}
        if action == "delete":
            card_id = (data.get("id") or "").strip()
            storage._delete_vault_card(card_id)
            self._log("Đã xoá một thẻ khỏi kho.")
            return {"ok": True}
        if action == "delete-many":
            ids = _card_ids(data)
            if not ids:
                raise ConsoleError("Chọn ít nhất một thẻ.")
            for card_id in ids:
                storage._delete_vault_card(card_id)
            self._log(f"Đã xoá {len(ids)} thẻ khỏi kho.")
            return {"ok": True, "count": len(ids)}
        if action == "limit":
            try:
                limit = int(data.get("limit") or 0)
            except (TypeError, ValueError):
                limit = 0
            if not storage._set_vault_card_limit(
                (data.get("id") or "").strip(), limit
            ):
                raise ConsoleError("Không còn thẻ này trong kho.")
            self._log("Đã lưu giới hạn thẻ.")
            return {"ok": True}
        if action == "limit-many":
            ids = _card_ids(data)
            if not ids:
                raise ConsoleError("Chọn ít nhất một thẻ.")
            try:
                limit = int(data.get("limit") or 0)
            except (TypeError, ValueError):
                limit = 0
            count = storage._set_vault_card_limits(ids, limit)
            self._log(f"Đã đặt giới hạn cho {count} thẻ đã chọn.")
            return {"ok": True, "count": count}
        if action == "limit-all":
            try:
                limit = int(data.get("limit") or 0)
            except (TypeError, ValueError):
                limit = 0
            count = storage._set_all_vault_card_limits(limit)
            self._log(f"Đã đặt giới hạn chung cho {count} thẻ.")
            return {"ok": True, "count": count}
        raise ConsoleError("Không rõ việc cần làm với kho thẻ.")

    def save_capsolver(self, key):
        key = (key or "").strip()
        if not key:
            raise ConsoleError("Nhập CapSolver API key rồi bấm Lưu.")
        storage._save_capsolver_key(key)
        self._log("Đã lưu CapSolver API key.")
        return {"ok": True}

    def capsolver_balance(self):
        key = storage._load_capsolver_key()
        if not key:
            raise ConsoleError("Chưa lưu CapSolver API key.")
        import capsolver

        balance = capsolver.get_balance(key)
        self._log(f"Số dư CapSolver: ${balance:.2f}.")
        return {"balance": balance}

    def stop(self):
        self.coordinator.stop()
        self._log("Đã yêu cầu dừng lượt chạy.")
        return {"ok": True}

    def start_muse_register(self, data):
        raise ConsoleError(
            "Không đăng ký Muse mới. Dùng Đăng nhập Muse với file "
            "Excel tài khoản Meta đã đăng ký."
        )

    def start_muse_login(self, data):
        excel = (data.get("excel") or "").strip()
        if excel:
            storage._save_meta_excel_path_text(excel)
        if not excel or not Path(excel).is_file():
            raise ConsoleError(
                "Chọn file Excel Meta trên máy đang chạy tool."
            )
        pending = storage._pending_meta_logins(excel)
        if not pending:
            raise ConsoleError(
                "File không còn tài khoản Meta nào chưa đăng nhập Muse."
            )
        setup = _read_setup(data, "proxies.txt", "bad-proxies.txt", mail=False)
        cards, manual = _run_cards()
        if not cards and not data.get("allow_manual_card"):
            raise ConsoleError(
                "Chưa có thẻ trong kho. Gửi lại với xác nhận nếu "
                "muốn nhập thẻ tay trên trình duyệt."
            )
        _check_cards(cards, manual)
        invite = (data.get("invite") or "").strip()
        if invite:
            paths.data_file("invite.txt").write_text(
                invite + "\n", encoding="utf-8"
            )
        else:
            invite = storage._last_invite_from_sheet() or ""
        settings = RunSettings(
            total=setup["total"],
            pool=setup["pool"],
            per_proxy=setup["per_proxy"],
            lanes=setup["lanes"],
            provider_id=setup["provider_id"],
            direct=setup["direct"],
            meta_login=True,
            meta_excel_path=excel,
            invite=invite,
            single_invite=bool(data.get("single_invite")),
            cards=cards,
            manual_card=manual,
            bad_file=setup["bad_file"],
            bad_mapping=setup["bad_mapping"],
            headless=setup["headless"],
        )
        self._begin(
            settings,
            f"đăng nhập Muse từ {len(pending)} tài khoản Meta",
        )
        return {"ok": True}

    def start_meta(self, data):
        key = storage._load_capsolver_key()
        if not key:
            entered = (data.get("capsolver") or "").strip()
            if entered:
                storage._save_capsolver_key(entered)
                key = entered
        if not key:
            raise ConsoleError(
                "Nhập CapSolver API key và bấm Lưu trước khi đăng ký Meta."
            )
        selfie = (data.get("selfie") or "").strip()
        if selfie:
            storage._save_selfie_path_text(selfie)
        if not selfie or not Path(selfie).is_file():
            raise ConsoleError(
                "Chọn file ảnh selfie trên máy đang chạy tool."
            )
        setup = _read_setup(
            data, "proxies_meta.txt", "bad-proxies-meta.txt", mail=True
        )
        settings = RunSettings(
            total=setup["total"],
            pool=setup["pool"],
            per_proxy=setup["per_proxy"],
            lanes=setup["lanes"],
            provider_id=setup["provider_id"],
            direct=setup["direct"],
            meta=True,
            bad_file=setup["bad_file"],
            bad_mapping=setup["bad_mapping"],
            headless=setup["headless"],
        )
        self._begin(settings, f"đăng ký Meta, mục tiêu {setup['total']}")
        return {"ok": True}

    def start_relogin(self, data):
        output = (data.get("output") or "").strip()
        if not output:
            raise ConsoleError("Chọn file Excel để ghi cookie.")
        try:
            count = int(data.get("count") or 0)
        except (TypeError, ValueError):
            raise ConsoleError("Số mail phải là số nguyên.")
        if count < 0:
            raise ConsoleError("Số mail phải từ 0 trở lên.")
        try:
            lanes = int(data.get("lanes") or 1)
        except (TypeError, ValueError):
            lanes = 1
        lanes = max(1, min(lanes, 10))
        rows = storage._relogin_accounts()
        if not rows:
            raise ConsoleError("Kho chưa có mail. Hãy nhập từ Excel trước.")
        done = storage._relogin_done_emails(output)
        pending = [
            row for row in rows if row["email"].lower() not in done
        ]
        if not pending:
            raise ConsoleError(
                "Mọi mail đã có cookie trong file xuất."
            )
        storage._save_relogin_output_path(output)
        planned = len(pending) if count == 0 else min(count, len(pending))
        settings = RunSettings(
            total=count,
            direct=True,
            lanes=lanes,
            relogin=True,
            relogin_output_path=output,
            headless=bool(data.get("headless")),
        )
        self._begin(settings, f"đăng nhập lại {planned} mail")
        return {"ok": True}

    def _begin(self, settings, label):
        if self.coordinator.running:
            raise ConsoleError("Đã có một lượt chạy. Bấm Dừng hoặc đợi xong.")
        self.progress = "Đang chạy…"
        self._log(f"--- Bắt đầu {label} ---")
        if not self.coordinator.start(settings):
            raise ConsoleError("Đã có một lượt chạy khác.")


def _format_progress(payload):
    (
        lane_no,
        done,
        label_total,
        success,
        fail,
        card_note,
        multi,
    ) = payload
    if not multi:
        return (
            f"Tài khoản {done}/{label_total} · Thành công {success} · "
            f"Thất bại {fail}{card_note}"
        )
    return (
        f"Luồng {lane_no}: {done} · Tổng thành công {success} · "
        f"Thất bại {fail}{card_note}"
    )


def _pool_view(pool_name, bad_name):
    return {
        "proxies": _read_text(paths.data_file(pool_name)),
        "bad": "\n".join(
            proxy_pool._load_bad_proxies(paths.data_file(bad_name)).values()
        ),
    }


def _read_text(path):
    try:
        return path.read_text(encoding="utf-8")
    except OSError:
        return ""


def _read_setup(data, pool_name, bad_name, mail):
    try:
        total = int(data.get("total") if data.get("total") is not None else 5)
    except (TypeError, ValueError):
        total = 0
    if total < 0 or total > 100:
        raise ConsoleError("Số tài khoản phải từ 0 đến 100.")
    try:
        per_proxy = int(data.get("per_proxy") or 5)
    except (TypeError, ValueError):
        per_proxy = 5
    per_proxy = max(1, min(per_proxy, 100))
    try:
        lanes = int(data.get("lanes") or 1)
    except (TypeError, ValueError):
        lanes = 1
    lanes = max(1, min(lanes, 5))
    direct = bool(data.get("direct"))
    pool_file = paths.data_file(pool_name)
    bad_file = paths.data_file(bad_name)
    bad_mapping = {}
    for line in (data.get("bad") or "").splitlines():
        line = line.strip()
        if line:
            bad_mapping[proxy_pool._proxy_key(line)] = line
    proxy_pool._save_bad_proxies(bad_mapping, bad_file)
    if direct:
        pool = [DIRECT_PROXY_LINE]
        per_proxy = 10**9
    else:
        pool = [
            line.strip()
            for line in (data.get("proxies") or "").splitlines()
            if line.strip()
        ]
        if not pool:
            raise ConsoleError(
                "Chế độ proxy cần ít nhất một dòng host:port:user:pass, "
                "hoặc chọn IP máy."
            )
        pool_file.write_text("\n".join(pool) + "\n", encoding="utf-8")
        paths.data_file("proxy.txt").write_text(pool[0] + "\n", encoding="utf-8")
        pool = proxy_pool._prepare_pool(pool, set(bad_mapping))
        if not pool:
            raise ConsoleError(
                "Mọi proxy đang nằm trong kho lỗi. Xoá bớt dòng lỗi "
                "hoặc dán proxy mới."
            )
    if mail:
        provider_id = (data.get("provider") or mail_providers.DEFAULT_PROVIDER).strip()
        known = {pid for pid, _label in mail_providers.PROVIDERS}
        if provider_id not in known:
            provider_id = mail_providers.DEFAULT_PROVIDER
    else:
        provider_id = mail_providers.DEFAULT_PROVIDER
    return {
        "total": total,
        "per_proxy": per_proxy,
        "lanes": lanes,
        "direct": direct,
        "pool": pool,
        "bad_file": bad_file,
        "bad_mapping": bad_mapping,
        "provider_id": provider_id,
        "headless": bool(data.get("headless")),
    }


def _card_ids(data):
    raw = data.get("ids")
    if not isinstance(raw, list):
        raw = [data.get("id")] if data.get("id") else []
    return storage._clean_picked_ids(raw)


def _run_cards():
    chosen, manual = storage.vault_cards_for_run()
    return (
        [
            CardRunState(
                card_id=card["id"],
                label=card["label"],
                limit=card["limit"],
                used=card["used"],
            )
            for card in chosen
        ],
        manual,
    )


def _check_cards(cards, manual):
    if not cards:
        return
    if manual:
        card = cards[0]
        if card.limit > 0 and card.used >= card.limit:
            raise ConsoleError(
                f"Thẻ {card.label} đã đủ giới hạn {card.limit}. "
                "Chọn thẻ khác hoặc tăng giới hạn."
            )
        return
    if all(card.limit > 0 and card.used >= card.limit for card in cards):
        raise ConsoleError(
            "Mọi thẻ đã đủ giới hạn. Tăng giới hạn hoặc nhập thêm thẻ."
        )
