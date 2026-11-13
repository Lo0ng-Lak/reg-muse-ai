"""Kiểm tra kho thẻ (card vault) với keyring giả lưu trong bộ nhớ.

Chạy:  python test_card_vault.py
Không cần Windows hay Credential Manager thật: keyring được thay
bằng một bản giả lưu trong dict, nạp vào sys.modules TRƯỚC khi
import storage / coordinator / browser_helper; win32cred và
win32crypt cũng được thay bằng bản giả (chỉ trong test cần) để
kiểm tra dọn thẻ mồ côi và mã hoá DPAPI. Không dùng số thẻ
thật — chỉ dùng các số thẻ TEST chuẩn (4111…, 4222…, 4000…).
Giao diện tkinter (muse_tab) không được kiểm tra ở đây — cần chạy
trên máy Windows của Long.
"""

import json
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

import paths

# Danh mục kho nằm ở file vault.json (xem storage): trong test
# này paths.data_file được chuyển sang thư mục tạm để test
# không chạm tới thư mục dữ liệu thật.
TMP = Path(tempfile.mkdtemp(prefix="reg-test-vault-"))
paths.data_file = lambda name: TMP / name


class _FakeKeyring:
    """Bản giả tối thiểu của keyring: get/set/delete_password
    trên một dict {(service, field): value}."""

    def __init__(self):
        self.store = {}

    def get_password(self, service, field):
        return self.store.get((service, field))

    def set_password(self, service, field, value):
        self.store[(service, field)] = value

    def delete_password(self, service, field):
        self.store.pop((service, field), None)


_FAKE_KEYRING = _FakeKeyring()
sys.modules["keyring"] = _FAKE_KEYRING

import storage
from coordinator import CardRunState, RunSettings, _CardVaultRun
import browser_helper

# Số thẻ TEST chuẩn (không phải thẻ thật).
CARD_A = "4111111111111111"
CARD_B = "4222222222222222"
CARD_C = "4000056655665556"


def reset():
    _FAKE_KEYRING.store.clear()
    for name in (
        "vault.json",
        "vault.pending.json",
        "vault.json.tmp",
        "vault.pending.json.tmp",
        "cards.json",
        "cards.json.tmp",
    ):
        try:
            (TMP / name).unlink()
        except OSError:
            pass
    # Đặt lại bộ nhớ đệm của storage để mỗi test chạy như một
    # tiến trình mới: chuyển đổi CM -> file, backend DPAPI,
    # module win32cred, lỗi đọc file thẻ.
    storage._MIGRATION_DONE = False
    storage._LAST_MIGRATION = None
    storage._DPAPI_BACKEND = storage._UNSET
    storage._WIN32CRED = storage._UNSET
    storage._LAST_CARDS_READ_ERROR = ""
    sys.modules.pop("win32cred", None)
    sys.modules.pop("win32crypt", None)


def _seed_cm_vault_card(card_id, number, count=0):
    """Ghi thẳng một thẻ vào keyring giả theo định dạng CM cũ
    (4 trường + bộ đếm), như bản trước khi chuyển sang file."""
    service = f"MuseAIHelper:card:{card_id}"
    _FAKE_KEYRING.store[(service, "number")] = number
    _FAKE_KEYRING.store[(service, "expiry")] = "12/28"
    _FAKE_KEYRING.store[(service, "cvv")] = "123"
    _FAKE_KEYRING.store[(service, "zip")] = "10001"
    if count:
        _FAKE_KEYRING.store[(service, "card_success_count")] = str(count)


class _FakeWin32Crypt:
    """DPAPI giả: mã hoá = đảo ngược bytes + tiền tố, để test
    chứng minh file thẻ không chứa chữ số thẻ ở dạng đọc được."""

    @staticmethod
    def CryptProtectData(data, description, *args):
        return b"ENC:" + bytes(data)[::-1]

    @staticmethod
    def CryptUnprotectData(blob, *args):
        assert blob.startswith(b"ENC:")
        return ("", blob[4:][::-1])


class _FakeWin32Cred:
    """win32cred giả: liệt kê các TargetName đã cho."""

    CRED_TYPE_GENERIC = 1

    def __init__(self, targets):
        self.targets = list(targets)

    def CredEnumerate(self, _filter, _flags):
        return [{"TargetName": target} for target in self.targets]


def _add_test_card(number, label, limit=0, used=0):
    """Thêm một thẻ test vào kho qua đường nhập hàng loạt, trả về id."""
    added, errors, skipped = storage._import_bulk_cards(
        f"{number}|12/28|123|10001|{label}"
    )
    assert not errors, errors
    assert skipped == 0, skipped
    assert added == [label], added
    entry = next(
        e for e in storage._load_vault_index() if e["label"] == label
    )
    if limit:
        storage._set_vault_card_limit(entry["id"], limit)
    if used:
        storage._save_vault_card_success_count(entry["id"], used)
    return entry["id"]


def _vault_states():
    """CardRunState list từ kho hiện tại (như tab Muse dựng lúc chạy)."""
    return [
        CardRunState(
            card_id=card["id"],
            label=card["label"],
            limit=card["limit"],
            used=card["used"],
        )
        for card in storage._load_vault()
    ]


# --- Kho lưu trữ -------------------------------------------------------

def test_index_round_trip_and_clamp():
    storage._save_vault_index(
        [
            {"id": "card-aaaa0001", "label": "Thẻ A", "limit": 5},
            {"id": "card-aaaa0002", "label": "Thẻ B", "limit": 150},
            {"id": "card-aaaa0003", "label": "Thẻ C", "limit": -3},
        ]
    )
    index = storage._load_vault_index()
    assert [e["label"] for e in index] == ["Thẻ A", "Thẻ B", "Thẻ C"]
    assert index[0]["limit"] == 5
    assert index[1]["limit"] == 100, "giới hạn kẹp tối đa 100"
    assert index[2]["limit"] == 0, "giới hạn kẹp tối thiểu 0"


def test_save_load_card_and_count():
    card_id = _add_test_card(CARD_A, "Thẻ A", limit=3)
    card = storage._load_vault_card(card_id)
    assert card == {
        "number": CARD_A,
        "expiry": "12/28",
        "cvv": "123",
        "zip": "10001",
    }
    assert storage._load_vault_card_success_count(card_id) == 0
    storage._save_vault_card_success_count(card_id, 3)
    assert storage._load_vault_card_success_count(card_id) == 3
    assert storage._load_vault_card("card-khong-co") == {}
    # Bộ đếm nằm trong danh mục vault.json (không ở file thẻ),
    # danh mục không chứa chữ số thẻ nào.
    raw = (TMP / "vault.json").read_text(encoding="utf-8")
    assert CARD_A not in raw, "danh mục không được chứa chữ số thẻ"
    assert '"used": 3' in raw, "bộ đếm nằm trong danh mục"


def test_delete_card_removes_index_and_active():
    id_a = _add_test_card(CARD_A, "Thẻ A")
    id_b = _add_test_card(CARD_B, "Thẻ B")
    storage._save_vault_active(id_a)
    storage._delete_vault_card(id_a)
    assert storage._load_vault_card(id_a) == {}
    cards, _ = storage._read_cards_file()
    assert id_a not in cards, "thẻ xoá phải biến khỏi file thẻ"
    assert storage._load_vault_card_success_count(id_a) == 0
    assert [e["id"] for e in storage._load_vault_index()] == [id_b]
    assert storage._load_vault_active() == "", "xoá thẻ đang chọn phải bỏ chọn"


def test_active_save_load_clear():
    id_a = _add_test_card(CARD_A, "Thẻ A")
    assert storage._load_vault_active() == ""
    storage._save_vault_active(id_a)
    assert storage._load_vault_active() == id_a
    storage._save_vault_active("")
    assert storage._load_vault_active() == ""


def test_set_limit():
    card_id = _add_test_card(CARD_A, "Thẻ A")
    assert storage._set_vault_card_limit(card_id, 7) is True
    assert storage._load_vault_index()[0]["limit"] == 7
    assert storage._set_vault_card_limit(card_id, 500) is True
    assert storage._load_vault_index()[0]["limit"] == 100
    assert storage._set_vault_card_limit("card-khong-co", 5) is False


def test_migrate_legacy_card():
    _FAKE_KEYRING.store[("MuseAIHelper:card", "number")] = CARD_A
    _FAKE_KEYRING.store[("MuseAIHelper:card", "expiry")] = "12/28"
    _FAKE_KEYRING.store[("MuseAIHelper:card", "cvv")] = "123"
    _FAKE_KEYRING.store[("MuseAIHelper:card", "zip")] = "10001"
    _FAKE_KEYRING.store[("MuseAIHelper:card", "card_success_count")] = "3"
    vault = storage._load_vault()
    assert len(vault) == 1
    assert vault[0]["label"] == "Thẻ 1"
    assert vault[0]["limit"] == 0
    assert vault[0]["used"] == 3, "bộ đếm cũ được mang sang"
    assert storage._load_card_from_keyring() == {}, "thẻ đơn cũ đã bị xoá"
    assert storage._load_card_success_count() == 0
    assert storage._load_vault_card(vault[0]["id"])["number"] == CARD_A, (
        "thẻ đơn cũ phải nằm trong file thẻ"
    )
    assert len(storage._load_vault()) == 1, "không nhập trùng lần hai"


def test_migrate_no_legacy():
    assert storage._load_vault() == []
    assert storage._migrate_legacy_card() is False


# --- Nhập hàng loạt ----------------------------------------------------

def test_bulk_import_valid():
    text = "\n".join(
        [
            f"{CARD_A}|12/28|123|10001|Thẻ chính",
            f"4222 2222 2222 2222|1228|1234",
            f"{CARD_C}|01/30|999|20000|",
        ]
    )
    added, errors, skipped = storage._import_bulk_cards(text)
    assert errors == [], errors
    assert skipped == 0
    assert added == ["Thẻ chính", "Thẻ 1", "Thẻ 2"], added
    vault = storage._load_vault()
    assert [c["label"] for c in vault] == added
    assert vault[0]["limit"] == 0
    card_b = storage._load_vault_card(vault[1]["id"])
    assert card_b["number"] == CARD_B, "khoảng trắng trong số thẻ được bỏ"
    assert card_b["expiry"] == "12/28", "1228 chuẩn hoá thành 12/28"
    assert card_b["cvv"] == "1234"
    assert card_b["zip"] == "10001", "ZIP trống mặc định 10001"


def test_bulk_import_invalid_lines_report_line_numbers_only():
    lines = [
        f"{CARD_A}|12/28|123|10001|Thẻ chính",  # dòng 1: hợp lệ
        "",  # dòng 2: trống, bỏ qua lặng lẽ
        "12345|12/28|123",  # dòng 3: số thẻ quá ngắn
        f"{CARD_B}|13/28|123",  # dòng 4: tháng 13
        f"{CARD_C}|12/28|12",  # dòng 5: CVV quá ngắn
        "4111111111111111|12/28",  # dòng 6: thiếu trường
        "4111111111111111|12/28|123|10001|Nhãn|thừa",  # dòng 7: thừa trường
        "4111111111111111|12/28|123|1234|Thẻ lẻ",  # dòng 8: ZIP 4 chữ số
    ]
    added, errors, skipped = storage._import_bulk_cards("\n".join(lines))
    assert added == ["Thẻ chính"]
    assert skipped == 0
    assert [line_no for line_no, _ in errors] == [3, 4, 5, 6, 7, 8]
    # Lý do lỗi không chứa chữ số nào của thẻ trong dòng lỗi.
    for line_no, reason in errors:
        number = lines[line_no - 1].split("|")[0].replace(" ", "")
        assert number not in reason, (line_no, reason)
        assert lines[line_no - 1] not in reason


def test_bulk_import_duplicates():
    text = "\n".join(
        [
            f"{CARD_A}|12/28|123|10001|Thẻ A",
            f"{CARD_A}|12/28|123|10001|Thẻ B",  # trùng số thẻ
            f"{CARD_B}|12/28|123|10001|Thẻ A",  # trùng nhãn
            f"{CARD_B}|12/28|123|10001|Thẻ B",  # hợp lệ
        ]
    )
    added, errors, skipped = storage._import_bulk_cards(text)
    assert added == ["Thẻ A", "Thẻ B"]
    assert skipped == 0
    assert [line_no for line_no, _ in errors] == [2, 3]
    # Trùng với thẻ đã có trong kho từ trước cũng bị chặn.
    added2, errors2, skipped2 = storage._import_bulk_cards(
        f"{CARD_A}|12/28|123|10001|Thẻ mới"
    )
    assert added2 == []
    assert skipped2 == 0
    assert [line_no for line_no, _ in errors2] == [1]


def test_bulk_import_empty_text():
    assert storage._import_bulk_cards("") == ([], [], 0)
    assert storage._import_bulk_cards("   \n\n  \n") == ([], [], 0)


# --- Danh mục kho ở file + nhập hàng loạt lớn --------------------------

def test_bulk_import_100_cards():
    """Kho mở rộng tới 100 thẻ: thẻ nằm ở file cards.json (không
    còn qua Credential Manager) nên nhập một đợt lớn vẫn ghi
    được, và không còn giới hạn blob 2560 byte nào."""
    lines = [f"4111111111{i:06d}|12/28|123" for i in range(100)]
    added, errors, skipped = storage._import_bulk_cards("\n".join(lines))
    assert errors == [], errors
    assert skipped == 0
    assert added == [f"Thẻ {i}" for i in range(1, 101)]
    assert len(storage._load_vault()) == 100
    assert (TMP / "vault.json").exists(), "danh mục kho phải nằm ở file vault.json"
    assert (TMP / "cards.json").exists(), "thẻ phải nằm ở file cards.json"
    assert ("MuseAIHelper:cards", "index") not in _FAKE_KEYRING.store, (
        "danh mục không còn ghi vào Credential Manager"
    )
    assert not any(
        service.startswith("MuseAIHelper:card:")
        for service, _field in _FAKE_KEYRING.store
    ), "thẻ không còn ghi vào Credential Manager"
    # Nhập lại đúng 100 dòng đó: toàn bộ trùng, không thêm thẻ nào.
    added2, errors2, skipped2 = storage._import_bulk_cards("\n".join(lines))
    assert added2 == []
    assert len(errors2) == 100
    assert skipped2 == 0


def test_index_migrates_from_keyring_to_file():
    """Nâng cấp: danh mục kho cũ còn trong Credential Manager
    được chuyển sang file vault.json một lần duy nhất, rồi mục
    CM cũ bị xoá; helper đọc được thẻ qua file."""
    _FAKE_KEYRING.store[("MuseAIHelper:cards", "index")] = json.dumps(
        [
            {"id": "card-old0001", "label": "Thẻ 1", "limit": 0},
            {"id": "card-old0002", "label": "Thẻ 2", "limit": 5},
        ],
        ensure_ascii=False,
    )
    index = storage._load_vault_index()
    assert [e["label"] for e in index] == ["Thẻ 1", "Thẻ 2"]
    assert index[1]["limit"] == 5
    assert (TMP / "vault.json").exists()
    assert ("MuseAIHelper:cards", "index") not in _FAKE_KEYRING.store, (
        "mục CM cũ phải được xoá sau khi chuyển sang file"
    )
    assert browser_helper._vault_card_ids() == ["card-old0001", "card-old0002"]
    # Đọc lại lần hai: không chuyển trùng, không mất thẻ.
    assert [e["id"] for e in storage._load_vault_index()] == [
        "card-old0001",
        "card-old0002",
    ]


def test_cards_file_write_failure_adds_nothing():
    """Ghi file thẻ lỗi: import raise lỗi rõ ràng bằng tiếng Việt,
    không thẻ nào được thêm, danh mục không đổi."""
    original_write = storage._write_cards_file

    def _failing_write(cards):
        raise OSError("disk full (giả lập)")

    storage._write_cards_file = _failing_write
    try:
        try:
            storage._import_bulk_cards(f"{CARD_A}|12/28|123")
        except RuntimeError as error:
            message = str(error)
        else:
            raise AssertionError("phải raise RuntimeError khi ghi file thẻ lỗi")
        assert "1783" not in message and "CredWrite" not in message
        assert "cards.json" in message
        assert "1 thẻ" in message
    finally:
        storage._write_cards_file = original_write
    assert storage._load_vault() == [], "ghi file lỗi thì không thêm thẻ nào"
    assert not (TMP / "cards.json").exists()


def test_index_save_failure_rolls_back_cards_file():
    """Lưu danh mục lỗi sau khi đã ghi file thẻ: import raise lỗi
    rõ ràng, file thẻ được hoàn lại như cũ (không mất thẻ cũ,
    không thêm thẻ mới)."""
    _add_test_card(CARD_A, "Thẻ A")
    original_save = storage._save_vault_index

    def _failing_save(entries):
        raise OSError("disk full (giả lập)")

    storage._save_vault_index = _failing_save
    try:
        try:
            storage._import_bulk_cards(f"{CARD_B}|12/28|123")
        except RuntimeError as error:
            message = str(error)
        else:
            raise AssertionError("phải raise RuntimeError khi lưu danh mục lỗi")
        assert "1783" not in message and "CredWrite" not in message
        assert "danh mục" in message
        assert "1 thẻ" in message
    finally:
        storage._save_vault_index = original_save
    vault = storage._load_vault()
    assert [c["label"] for c in vault] == ["Thẻ A"], "thẻ mới phải bị hoàn lại"
    cards, _ = storage._read_cards_file()
    assert len(cards) == 1, "file thẻ phải hoàn lại như cũ"
    assert storage._load_vault_card(vault[0]["id"])["number"] == CARD_A


def test_import_writes_no_card_to_keyring():
    """Nhập thẻ không ghi bất cứ mục thẻ nào vào Credential
    Manager — CM không còn là database của thẻ."""
    storage._import_bulk_cards(f"{CARD_A}|12/28|123\n{CARD_B}|12/28|123")
    assert _FAKE_KEYRING.store == {}, (
        "nhập thẻ không được chạm tới Credential Manager"
    )
    cards, protected = storage._read_cards_file()
    assert len(cards) == 2
    assert protected is False, "Linux không có DPAPI nên file lưu dạng thường"


def test_skipped_header_lines():
    """Dòng tiêu đề/phân cách (không có chữ số nào) được bỏ qua
    và báo số lượng, không tính là lỗi; thẻ thật vẫn nhập."""
    text = "\n".join(
        [
            "số thẻ|MM/YY|CVV",
            "---",
            "Danh sách thẻ mới",
            f"{CARD_A}|12/28|123",
        ]
    )
    added, errors, skipped = storage._import_bulk_cards(text)
    assert added == ["Thẻ 1"]
    assert errors == []
    assert skipped == 3


# --- Chuyển sang file + dọn CM mồ côi + DPAPI ------------------------

def test_migrate_cm_vault_cards_to_file_and_cleanup_orphans():
    """Nâng cấp từ bản CM cũ: thẻ trong danh mục được chuyển sang
    file (kèm bộ đếm), mục CM của thẻ đó bị xoá; thẻ mồ côi (có
    mục CM nhưng không trong danh mục) bị dọn; dịch vụ khác
    (capsolver) không bị đụng."""
    _seed_cm_vault_card("card-keep0001", CARD_A, count=2)
    _seed_cm_vault_card("card-orph0001", CARD_B)
    _seed_cm_vault_card("card-orph0002", CARD_C)
    _FAKE_KEYRING.store[("MuseAIHelper:capsolver", "api_key")] = "cap-key-giu-lai"
    storage._save_vault_index(
        [{"id": "card-keep0001", "label": "Thẻ giữ", "limit": 0}]
    )
    sys.modules["win32cred"] = _FakeWin32Cred(
        [
            "MuseAIHelper:card:card-keep0001",
            "number@MuseAIHelper:card:card-orph0001",
            "expiry@MuseAIHelper:card:card-orph0001",
            "cvv@MuseAIHelper:card:card-orph0001",
            "zip@MuseAIHelper:card:card-orph0001",
            "card_success_count@MuseAIHelper:card:card-orph0001",
            "MuseAIHelper:card:card-orph0002",
            "api_key@MuseAIHelper:capsolver",
        ]
    )
    try:
        vault = storage._load_vault()
    finally:
        sys.modules.pop("win32cred", None)
    assert [c["label"] for c in vault] == ["Thẻ giữ"]
    assert vault[0]["used"] == 2, "bộ đếm CM được mang sang danh mục"
    assert storage._load_vault_card("card-keep0001")["number"] == CARD_A
    assert storage._load_vault_card("card-orph0001") == {}
    assert storage._load_vault_card("card-orph0002") == {}
    remaining = set(_FAKE_KEYRING.store)
    assert ("MuseAIHelper:capsolver", "api_key") in remaining, (
        "dịch vụ khác không được đụng"
    )
    assert not any(
        service.startswith("MuseAIHelper:card:")
        for service, _field in remaining
    ), "mục CM của kho thẻ phải được dọn hết"
    report = storage._last_migration_report()
    assert report["migrated"] == 1
    assert report["counts_migrated"] == 1
    assert report["orphans_removed"] == 2
    assert report["enumerated"] is True
    # Đọc lại lần hai: không chuyển trùng, không mất thẻ.
    assert [c["label"] for c in storage._load_vault()] == ["Thẻ giữ"]


def test_migrate_without_enumeration_keeps_orphans_but_moves_catalogued():
    """Không liệt kê được CM (không phải Windows): thẻ trong danh
    mục vẫn chuyển sang file, thẻ mồ côi ở lại CM, báo cáo ghi rõ
    không liệt kê được — không mất thẻ, không lỗi."""
    _seed_cm_vault_card("card-keep0001", CARD_A, count=1)
    _seed_cm_vault_card("card-orph0001", CARD_B)
    storage._save_vault_index(
        [{"id": "card-keep0001", "label": "Thẻ giữ", "limit": 0}]
    )
    vault = storage._load_vault()
    assert [c["label"] for c in vault] == ["Thẻ giữ"]
    assert storage._load_vault_card("card-keep0001")["number"] == CARD_A
    assert ("MuseAIHelper:card:card-orph0001", "number") in _FAKE_KEYRING.store, (
        "không liệt kê được thì thẻ mồ côi ở lại CM, không tự xoá bừa"
    )
    report = storage._last_migration_report()
    assert report["migrated"] == 1
    assert report["enumerated"] is False
    assert report["cleanup_error"] != ""
    assert report["orphans_removed"] == 0


def test_dpapi_fake_round_trip():
    """Có DPAPI (giả): file thẻ mã hoá — không chứa chữ số thẻ ở
    dạng đọc được; đọc lại vẫn đủ 4 trường; không có cảnh báo."""
    sys.modules["win32crypt"] = _FakeWin32Crypt()
    try:
        card_id = _add_test_card(CARD_A, "Thẻ A")
        raw = (TMP / "cards.json").read_text(encoding="utf-8")
        assert CARD_A not in raw, "file mã hoá không được chứa số thẻ"
        assert '"protected": true' in raw
        assert storage._load_vault_card(card_id) == {
            "number": CARD_A,
            "expiry": "12/28",
            "cvv": "123",
            "zip": "10001",
        }
        assert storage._cards_protection_warning() is None
    finally:
        sys.modules.pop("win32crypt", None)


def test_plaintext_fallback_warns():
    """Không có DPAPI (Linux): file thẻ lưu dạng thường + cảnh báo
    rõ; thẻ vẫn dùng được."""
    card_id = _add_test_card(CARD_A, "Thẻ A")
    raw = (TMP / "cards.json").read_text(encoding="utf-8")
    assert '"protected": false' in raw
    warning = storage._cards_protection_warning()
    assert warning is not None
    assert "KHÔNG mã hoá" in warning
    assert storage._load_vault_card(card_id)["number"] == CARD_A


def test_pending_from_old_build_merges_and_migrates():
    """Bản d9800689 để lại vault.pending.json (thẻ đã ghi CM mà
    lưu danh mục lỗi): lần đọc sau hoà mục chờ vào kho, thẻ
    chuyển sang file, file chờ bị xoá."""
    _seed_cm_vault_card("card-pend0001", CARD_A)
    (TMP / "vault.pending.json").write_text(
        json.dumps(
            [{"id": "card-pend0001", "label": "Thẻ chờ", "limit": 0}],
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    vault = storage._load_vault()
    assert [c["label"] for c in vault] == ["Thẻ chờ"]
    assert storage._load_vault_card("card-pend0001")["number"] == CARD_A
    assert not (TMP / "vault.pending.json").exists()


# --- Chạy: giữ chỗ / xoay thẻ / giới hạn -------------------------------

def test_rotation_across_limits():
    _add_test_card(CARD_A, "Thẻ A", limit=2)
    _add_test_card(CARD_B, "Thẻ B", limit=1)
    _add_test_card(CARD_C, "Thẻ C", limit=0)
    vault = _CardVaultRun(_vault_states(), manual=False)
    # Thẻ A: 2 lượt, thẻ B: 1 lượt, rồi xoay sang thẻ C (không giới hạn).
    card, switched, stop = vault.reserve()
    assert (card.label, switched, stop) == ("Thẻ A", None, False)
    vault.record_success(card)
    vault.release(card)
    card, switched, stop = vault.reserve()
    assert (card.label, switched, stop) == ("Thẻ A", None, False)
    # Đang giữ chỗ lượt 2 của A thì A đã đủ (used 1 + in_flight 1 = 2).
    card2, switched2, stop2 = vault.reserve()
    assert (card2.label, switched2.label, stop2) == ("Thẻ B", "Thẻ A", False)
    vault.record_success(card2)
    vault.release(card2)
    vault.record_success(card)
    vault.release(card)
    card3, switched3, stop3 = vault.reserve()
    assert (card3.label, switched3.label, stop3) == ("Thẻ C", "Thẻ B", False)
    vault.release(card3)
    card4, _, _ = vault.reserve()
    assert card4.label == "Thẻ C", "thẻ không giới hạn dùng tiếp được"
    vault.release(card4)
    # Bộ đếm đã ghi lại vào kho theo id thẻ.
    vault_cards = storage._load_vault()
    assert [c["used"] for c in vault_cards] == [2, 1, 0]


def test_rotation_all_exhausted_stops():
    _add_test_card(CARD_A, "Thẻ A", limit=1, used=1)
    _add_test_card(CARD_B, "Thẻ B", limit=1, used=1)
    vault = _CardVaultRun(_vault_states(), manual=False)
    card, switched, stop = vault.reserve()
    assert (card, switched, stop) == (None, None, True)
    assert "Tất cả thẻ" in vault.stop_message()


def test_manual_mode_only_uses_selected_card():
    _add_test_card(CARD_A, "Thẻ A", limit=1)
    _add_test_card(CARD_B, "Thẻ B", limit=5)
    states = _vault_states()
    vault = _CardVaultRun(states[:1], manual=True)
    card, _, stop = vault.reserve()
    assert (card.label, stop) == ("Thẻ A", False)
    vault.record_success(card)
    vault.release(card)
    card, _, stop = vault.reserve()
    assert (card, stop) == (None, True), "thẻ tay đủ giới hạn là dừng dù thẻ khác còn chỗ"
    assert "Thẻ A" in vault.stop_message()
    assert vault.note_for(None) == ""


def test_limit_zero_is_unlimited():
    state = CardRunState(card_id="card-x", label="Thẻ X", limit=0, used=100)
    assert state.exhausted() is False
    vault = _CardVaultRun([state], manual=False)
    card, _, stop = vault.reserve()
    assert (card.label, stop) == ("Thẻ X", False)
    vault.release(card)


def test_empty_vault_means_manual_entry():
    vault = _CardVaultRun([], manual=False)
    card, switched, stop = vault.reserve()
    assert (card, switched, stop) == (None, None, False)
    vault.record_success(None)
    vault.release(None)


def test_run_settings_defaults():
    settings = RunSettings()
    assert settings.cards == []
    assert settings.manual_card is False
    states = [CardRunState(card_id="card-x", label="Thẻ X", limit=2, used=1)]
    settings2 = RunSettings(cards=states, manual_card=True)
    assert settings2.cards == states
    assert settings2.manual_card is True


# --- Helper: chọn thẻ ---------------------------------------------------

def test_helper_saved_card_explicit_id():
    id_a = _add_test_card(CARD_A, "Thẻ A")
    id_b = _add_test_card(CARD_B, "Thẻ B")
    assert browser_helper._saved_card(id_b)["number"] == CARD_B
    assert browser_helper._saved_card(id_a)["number"] == CARD_A
    # Id lạ: rơi xuống các mức dự phòng, không lỗi.
    assert browser_helper._saved_card("card-khong-co")["number"] == CARD_A


def test_helper_saved_card_fallback_chain():
    id_a = _add_test_card(CARD_A, "Thẻ A")
    id_b = _add_test_card(CARD_B, "Thẻ B")
    # Không tham số: thẻ đầu tiên trong kho.
    assert browser_helper._saved_card()["number"] == CARD_A
    # Thẻ chọn tay được ưu tiên hơn thẻ đầu kho.
    storage._save_vault_active(id_b)
    assert browser_helper._saved_card()["number"] == CARD_B
    storage._save_vault_active("")
    # Kho trống: thẻ đơn cũ.
    storage._delete_vault_card(id_a)
    storage._delete_vault_card(id_b)
    _FAKE_KEYRING.store[("MuseAIHelper:card", "number")] = CARD_C
    _FAKE_KEYRING.store[("MuseAIHelper:card", "expiry")] = "12/28"
    _FAKE_KEYRING.store[("MuseAIHelper:card", "cvv")] = "123"
    _FAKE_KEYRING.store[("MuseAIHelper:card", "zip")] = "10001"
    assert browser_helper._saved_card()["number"] == CARD_C
    # Không còn gì: dict rỗng — bước thẻ chờ nhập tay.
    storage._delete_card_from_keyring()
    assert browser_helper._saved_card() == {}


def test_helper_saved_card_stale_active_falls_through():
    id_a = _add_test_card(CARD_A, "Thẻ A")
    id_b = _add_test_card(CARD_B, "Thẻ B")
    storage._save_vault_active(id_b)
    storage._delete_vault_card(id_b)
    assert browser_helper._saved_card()["number"] == CARD_A


TESTS = [
    test_index_round_trip_and_clamp,
    test_save_load_card_and_count,
    test_delete_card_removes_index_and_active,
    test_active_save_load_clear,
    test_set_limit,
    test_migrate_legacy_card,
    test_migrate_no_legacy,
    test_bulk_import_valid,
    test_bulk_import_invalid_lines_report_line_numbers_only,
    test_bulk_import_duplicates,
    test_bulk_import_empty_text,
    test_bulk_import_100_cards,
    test_index_migrates_from_keyring_to_file,
    test_cards_file_write_failure_adds_nothing,
    test_index_save_failure_rolls_back_cards_file,
    test_import_writes_no_card_to_keyring,
    test_skipped_header_lines,
    test_migrate_cm_vault_cards_to_file_and_cleanup_orphans,
    test_migrate_without_enumeration_keeps_orphans_but_moves_catalogued,
    test_dpapi_fake_round_trip,
    test_plaintext_fallback_warns,
    test_pending_from_old_build_merges_and_migrates,
    test_rotation_across_limits,
    test_rotation_all_exhausted_stops,
    test_manual_mode_only_uses_selected_card,
    test_limit_zero_is_unlimited,
    test_empty_vault_means_manual_entry,
    test_run_settings_defaults,
    test_helper_saved_card_explicit_id,
    test_helper_saved_card_fallback_chain,
    test_helper_saved_card_stale_active_falls_through,
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
