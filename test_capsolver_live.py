"""Kiểm tra API CapSolver THẬT bằng captcha demo của Google.

Chạy trên máy Windows (nơi đã lưu key CapSolver trong tool):
    py test_capsolver_live.py

Script này:
  1. Lấy key từ keyring của tool (MuseAIHelper:capsolver) — KHÔNG
     cần dán key, KHÔNG in key ra màn hình.
  2. Kiểm tra số dư qua getBalance.
  3. Giải 1 captcha reCAPTCHA v2 trên trang demo công khai của
     Google (https://www.google.com/recaptcha/api2/demo) qua
     createTask + getTaskResult — đúng luồng mà tool dùng khi
     gặp captcha ở Meta.
  4. In độ dài token nhận được (không in toàn bộ token).

Chi phí: 1 lần giải reCAPTCHA v2 (khoảng $0.001 — một phần nghìn
đô, không đáng kể so với số dư $10).

Kết quả này CHỨNG MINH được: key đúng, API gọi được, task tạo
được, token trả về được. Nó KHÔNG chứng minh được phần chèn
token vào trang Meta (callback của Meta) — phần đó chỉ lần chạy
Meta live gặp captcha mới kiểm chứng được.
"""

import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))

import capsolver  # noqa: E402
import storage  # noqa: E402

DEMO_SITEKEY = "6Le-wvkSAAAAAPBMRTvw0Q4Muexq9bi0DJwx_mJ-"
DEMO_URL = "https://www.google.com/recaptcha/api2/demo"


def main():
    api_key = storage._load_capsolver_key()
    if not api_key:
        print("Chưa lưu key CapSolver trong tool.")
        print("Mở tool -> tab Meta -> nhập key vào ô CapSolver -> Lưu key,")
        print("rồi chạy lại script này.")
        return 1

    print("Key đã lưu trong tool: có (không in ra đây).")
    try:
        balance = capsolver.get_balance(api_key)
    except capsolver.CapSolverError as error:
        print(f"Kiểm tra số dư LỖI: {error}")
        return 1
    print(f"Số dư CapSolver: ${balance}")

    task = {
        "type": "ReCaptchaV2TaskProxyLess",
        "websiteURL": DEMO_URL,
        "websiteKey": DEMO_SITEKEY,
    }
    print("Đang giải captcha demo của Google qua CapSolver...")
    started = time.monotonic()
    try:
        solution = capsolver.solve_task_full(api_key, task)
    except capsolver.CapSolverError as error:
        print(f"Giải captcha LỖI: {error}")
        return 1
    except TimeoutError as error:
        print(f"Giải captcha HẾT GIỜ: {error}")
        return 1
    elapsed = time.monotonic() - started
    token = solution.get("gRecaptchaResponse") or solution.get("token") or ""
    if not token:
        print(f"API trả về nhưng KHÔNG có token: {sorted(solution)}")
        return 1
    print(f"OK — nhận được token sau {elapsed:.1f}s.")
    print(f"Độ dài token: {len(token)} ký tự (đầu token: {token[:8]}...).")
    print("Kết luận: API CapSolver hoạt động được với key này.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
