"""Kiểm tra kho tài khoản SQLite (account_store).

Chạy:  python test_account_store.py
Không cần Windows: kho SQLite trỏ sang thư mục tạm, file Excel
chỉ dùng làm dữ liệu nhập/xuất mẫu.
"""

import sys
import tempfile
import time
from datetime import datetime, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

TMP = Path(tempfile.mkdtemp(prefix="reg-test-store-"))

import paths

paths.APP_DATA = TMP

import account_store

account_store.DB_PATH = TMP / "accounts.db"

COOKIE_A = "hatch_sess=aaa111; c_user=1001"
COOKIE_B = "hatch_sess=bbb222; c_user=1002"


def _now():
    return datetime.now().astimezone()


def _iso(dt):
    return dt.isoformat(timespec="seconds")


def reset():
    for name in ("accounts.db", "accounts.db-wal", "accounts.db-shm"):
        try:
            (TMP / name).unlink()
        except OSError:
            pass


def _write_excel(path, rows, headers=None):
    from openpyxl import Workbook

    workbook = Workbook()
    sheet = workbook.active
    sheet.append(list(headers or account_store.EXCEL_HEADERS))
    for row in rows:
        sheet.append(list(row))
    workbook.save(path)


def test_upsert_dedupes_case_insensitive():
    account_store.upsert_account(
        "A@fexpost.com", cookie=COOKIE_A, invite="ABC123"
    )
    # Ghi lại cùng mail (khác hoa/thường): không tạo dòng mới.
    account_store.upsert_account(
        "a@fexpost.com", cookie=COOKIE_B, invite="XYZ789"
    )
    rows = account_store.list_accounts()
    assert len(rows) == 1, rows
    assert rows[0]["cookie"] == COOKIE_B
    assert rows[0]["invite"] == "XYZ789"
    assert account_store.get_account("A@FEXPOST.COM") is not None


def test_upsert_empty_cookie_does_not_clobber():
    account_store.upsert_account("a@fexpost.com", cookie=COOKIE_A)
    account_store.upsert_account("a@fexpost.com", mail_password="s2")
    row = account_store.get_account("a@fexpost.com")
    assert row["cookie"] == COOKIE_A  # cookie cũ giữ nguyên
    assert row["mail_password"] == "s2"
    # Cookie mới đi kèm mốc lấy mới của nó.
    taken = _iso(_now())
    account_store.upsert_account(
        "a@fexpost.com", cookie=COOKIE_B, cookie_taken_at=taken
    )
    row = account_store.get_account("a@fexpost.com")
    assert row["cookie"] == COOKIE_B
    assert row["cookie_taken_at"] == taken


def test_update_and_clear_cookie():
    account_store.upsert_account("a@fexpost.com", cookie=COOKIE_A)
    taken = _iso(_now())
    assert account_store.update_cookie("a@fexpost.com", COOKIE_B, taken)
    row = account_store.get_account("a@fexpost.com")
    assert row["cookie"] == COOKIE_B
    assert row["cookie_taken_at"] == taken
    assert not account_store.update_cookie("nope@fexpost.com", "x", taken)
    assert account_store.clear_cookie("a@fexpost.com")
    row = account_store.get_account("a@fexpost.com")
    assert row["cookie"] == "" and row["cookie_taken_at"] == ""
    assert not account_store.clear_cookie("nope@fexpost.com")


def test_last_invite_skips_invalid():
    for email, invite in [
        ("a@fexpost.com", "ABC123"),
        ("b@fexpost.com", ""),
        ("c@fexpost.com", "không phải mã"),
        ("d@fexpost.com", "ZZZ999"),
        ("e@fexpost.com", "ngắn"),
    ]:
        account_store.upsert_account(email, invite=invite)
    assert account_store.last_invite() == "ZZZ999"
    # Kho trống thì không có mã.
    reset()
    assert account_store.last_invite() == ""


def test_import_six_column_excel():
    path = TMP / "old.xlsx"
    _write_excel(
        path,
        [
            [
                "2026-11-10T10:00:00+07:00",
                "old@fexpost.com",
                "s",
                COOKIE_A,
                "ABC123",
                "",
            ]
        ],
        headers=account_store.EXCEL_HEADERS[:6],
    )
    result = account_store.import_from_excel(path)
    assert result["imported"] == 1 and result["total"] == 1, result
    row = account_store.get_account("old@fexpost.com")
    assert row["cookie"] == COOKIE_A
    assert row["invite"] == "ABC123"
    assert row["cookie_taken_at"] == ""  # sheet cũ không có cột này


def test_import_idempotent_and_merges():
    path = TMP / "two.xlsx"
    _write_excel(
        path,
        [
            [_iso(_now()), "a@fexpost.com", "s", COOKIE_A, "ABC123", "", _iso(_now())],
            [_iso(_now()), "b@fexpost.com", "s", "", "", "", ""],
        ],
    )
    first = account_store.import_from_excel(path)
    assert first == {"imported": 2, "updated": 0, "skipped": 0, "total": 2}, first
    second = account_store.import_from_excel(path)
    assert second["imported"] == 0, second
    assert second["updated"] == 2 and second["total"] == 2, second
    # Dòng import thiếu cookie không xoá cookie đã có.
    assert account_store.get_account("a@fexpost.com")["cookie"] == COOKIE_A


def test_import_skips_rows_without_email():
    path = TMP / "bad.xlsx"
    _write_excel(
        path,
        [
            [_iso(_now()), "", "s", COOKIE_A, "", "", ""],
            [_iso(_now()), "không phải mail", "s", COOKIE_A, "", "", ""],
            [_iso(_now()), "ok@fexpost.com", "s", COOKIE_A, "", "", ""],
        ],
    )
    result = account_store.import_from_excel(path)
    assert result["imported"] == 1 and result["skipped"] == 2, result


def test_export_roundtrip():
    account_store.upsert_account(
        "a@fexpost.com",
        mail_password="s",
        cookie=COOKIE_A,
        invite="ABC123",
        cookie_taken_at=_iso(_now()),
    )
    account_store.upsert_account("b@fexpost.com", mail_password="s2")
    data = account_store.export_to_excel_bytes()
    assert len(data) > 0
    # Nhập lại vào kho mới: đủ 2 tài khoản, đúng dữ liệu.
    reset()
    path = TMP / "export.xlsx"
    path.write_bytes(data)
    result = account_store.import_from_excel(path)
    assert result["total"] == 2, result
    row = account_store.get_account("a@fexpost.com")
    assert row["cookie"] == COOKIE_A and row["invite"] == "ABC123"
    assert row["mail_password"] == "s"


def test_relogin_accounts():
    account_store.upsert_account(
        "a@fexpost.com", mail_password="secret-a"
    )
    account_store.upsert_account("b@fexpost.com", mail_password="secret-b")
    rows = account_store.relogin_accounts()
    assert rows == [
        {"email": "a@fexpost.com", "secret": "secret-a"},
        {"email": "b@fexpost.com", "secret": "secret-b"},
    ], rows


def test_list_2000_rows_fast():
    path = TMP / "big.xlsx"
    rows = []
    for i in range(2000):
        rows.append(
            [
                _iso(_now() - timedelta(days=i % 7)),
                f"acc{i:04d}@fexpost.com",
                "s",
                f"hatch_sess=c{i:04d}",
                f"INV{i:04d}",
                "",
                _iso(_now() - timedelta(days=i % 7)),
            ]
        )
    _write_excel(path, rows)
    result = account_store.import_from_excel(path)
    assert result["total"] == 2000, result
    started = time.monotonic()
    listed = account_store.list_accounts()
    elapsed = time.monotonic() - started
    assert len(listed) == 2000, len(listed)
    assert listed[0]["email"] == "acc0000@fexpost.com"
    assert listed[-1]["email"] == "acc1999@fexpost.com"
    assert elapsed < 1.0, f"list 2000 rows took {elapsed:.2f}s"


TESTS = [
    test_upsert_dedupes_case_insensitive,
    test_upsert_empty_cookie_does_not_clobber,
    test_update_and_clear_cookie,
    test_last_invite_skips_invalid,
    test_import_six_column_excel,
    test_import_idempotent_and_merges,
    test_import_skips_rows_without_email,
    test_export_roundtrip,
    test_relogin_accounts,
    test_list_2000_rows_fast,
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
