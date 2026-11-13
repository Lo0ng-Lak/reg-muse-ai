"""Kiểm tra kho thẻ (card vault) với keyring giả lưu trong bộ nhớ.

Chạy:  python test_card_vault.py
Không cần Windows hay Credential Manager thật: keyring được thay
bằng một bản giả lưu trong dict, nạp vào sys.modules TRƯỚC khi
import storage / coordinator / browser_helper. Không dùng số thẻ
thật — chỉ dùng các số thẻ TEST chuẩn (4111…, 4222…, 4000…).
Giao diện tkinter (muse_tab) không được kiểm tra ở đây — cần chạy
trên máy Windows của Long.
"""

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))


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


def _add_test_card(number, label, limit=0, used=0):
    """Thêm một thẻ test vào kho qua đường nhập hàng loạt, trả về id."""
    added, errors = storage._import_bulk_cards(
        f"{number}|12/28|123|10001|{label}"
    )
    assert not errors, errors
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


def test_delete_card_removes_index_and_active():
    id_a = _add_test_card(CARD_A, "Thẻ A")
    id_b = _add_test_card(CARD_B, "Thẻ B")
    storage._save_vault_active(id_a)
    storage._delete_vault_card(id_a)
    assert storage._load_vault_card(id_a) == {}
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
    storage._save_card_to_keyring(
        {"number": CARD_A, "expiry": "12/28", "cvv": "123", "zip": "10001"}
    )
    storage._save_card_success_count(3)
    vault = storage._load_vault()
    assert len(vault) == 1
    assert vault[0]["label"] == "Thẻ 1"
    assert vault[0]["limit"] == 0
    assert vault[0]["used"] == 3, "bộ đếm cũ được mang sang"
    assert storage._load_card_from_keyring() == {}, "thẻ đơn cũ đã bị xoá"
    assert storage._load_card_success_count() == 0
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
    added, errors = storage._import_bulk_cards(text)
    assert errors == [], errors
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
    added, errors = storage._import_bulk_cards("\n".join(lines))
    assert added == ["Thẻ chính"]
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
    added, errors = storage._import_bulk_cards(text)
    assert added == ["Thẻ A", "Thẻ B"]
    assert [line_no for line_no, _ in errors] == [2, 3]
    # Trùng với thẻ đã có trong kho từ trước cũng bị chặn.
    added2, errors2 = storage._import_bulk_cards(
        f"{CARD_A}|12/28|123|10001|Thẻ mới"
    )
    assert added2 == []
    assert [line_no for line_no, _ in errors2] == [1]


def test_bulk_import_empty_text():
    assert storage._import_bulk_cards("") == ([], [])
    assert storage._import_bulk_cards("   \n\n  \n") == ([], [])


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
    storage._save_card_to_keyring(
        {"number": CARD_C, "expiry": "12/28", "cvv": "123", "zip": "10001"}
    )
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
