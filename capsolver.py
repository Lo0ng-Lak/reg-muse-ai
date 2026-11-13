"""Client cho CapSolver (https://www.capsolver.com/) — dịch vụ giải
CAPTCHA trả phí theo lượt dùng.

Chỉ dùng cho bước CAPTCHA trong luồng đăng ký tài khoản Meta
(meta.ai): phát hiện loại CAPTCHA trên trang, gửi cho CapSolver
giải, rồi chèn token trả về vào form để trang tự xác nhận.

API key KHÔNG nằm trong file này và KHÔNG được in ra log bao
giờ — app chính lưu key trong Windows Credential Manager và
helper đọc lại qua keyring, giống cách lưu thẻ.

Dùng bản ProxyLess của mọi loại task: CapSolver giải bằng IP của
họ, không cần gửi proxy của mình sang đó (token gắn với sitekey
+ domain, không gắn IP, với các loại token-based này).

Tài liệu API (đối chiếu 2026-11-13):
- POST https://api.capsolver.com/createTask {clientKey, task}
  -> {errorId: 0, taskId} hoặc {errorId: 1, errorCode, ...}
- POST https://api.capsolver.com/getTaskResult {clientKey, taskId}
  -> {errorId: 0, status: "ready"|"processing", solution}
- POST https://api.capsolver.com/getBalance {clientKey}
  -> {errorId: 0, balance}
Loại task: ReCaptchaV2TaskProxyLess, HCaptchaTaskProxyless,
AntiTurnstileTaskProxyLess, FunCaptchaTaskProxyless.
"""

import json
import time
import urllib.error
import urllib.request

CAPSOLVER_BASE = "https://api.capsolver.com"
POLL_INTERVAL_SECONDS = 2
SOLVE_TIMEOUT_SECONDS = 180
REQUEST_TIMEOUT_SECONDS = 30

# kind -> (loại task CapSolver, các khoá token trong solution
# theo thứ tự ưu tiên). hCaptcha trả token dưới khoá
# gRecaptchaResponse giống reCAPTCHA; Turnstile/FunCaptcha trả
# dưới khoá token.
TASK_TYPES = {
    "recaptcha2": ("ReCaptchaV2TaskProxyLess", ("gRecaptchaResponse", "token")),
    "hcaptcha": ("HCaptchaTaskProxyless", ("gRecaptchaResponse", "token")),
    "turnstile": ("AntiTurnstileTaskProxyLess", ("token", "gRecaptchaResponse")),
    "funcaptcha": ("FunCaptchaTaskProxyless", ("token", "gRecaptchaResponse")),
}


class CapSolverError(RuntimeError):
    """Lỗi từ API CapSolver (key sai, hết tiền, task lỗi...)."""


def _post(path, payload):
    """POST JSON tới CapSolver, trả về dict đã parse.

    Lỗi HTTP/mạng -> CapSolverError với thông điệp ngắn (không
    kèm payload — payload chứa API key).
    """
    body = json.dumps(payload).encode("utf-8")
    request = urllib.request.Request(
        CAPSOLVER_BASE + path,
        data=body,
        headers={
            "Content-Type": "application/json",
            "User-Agent": "MuseAIHelper/1.0",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(
            request, timeout=REQUEST_TIMEOUT_SECONDS
        ) as response:
            return json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as error:
        try:
            detail = error.read().decode("utf-8", "replace")[:200]
        except Exception:
            detail = ""
        raise CapSolverError(
            f"CapSolver trả HTTP {error.code} ({detail or error.reason})."
        ) from error
    except (urllib.error.URLError, TimeoutError, ValueError) as error:
        raise CapSolverError(f"Không gọi được CapSolver: {error}.") from error


def _check_error(data, action):
    """Ném lỗi nếu CapSolver trả errorId != 0."""
    if not isinstance(data, dict) or data.get("errorId") != 0:
        code = data.get("errorCode") if isinstance(data, dict) else None
        description = (
            data.get("errorDescription") if isinstance(data, dict) else None
        )
        detail = description or code or "lỗi không rõ"
        raise CapSolverError(f"CapSolver {action} lỗi: {detail}.")


def get_balance(api_key):
    """Số dư tài khoản CapSolver (đơn vị USD)."""
    data = _post("/getBalance", {"clientKey": api_key})
    _check_error(data, "kiểm tra số dư")
    return float(data.get("balance", 0.0))


def create_task(api_key, task):
    """Tạo task giải CAPTCHA, trả về taskId."""
    data = _post("/createTask", {"clientKey": api_key, "task": task})
    _check_error(data, "tạo task")
    task_id = data.get("taskId")
    if not task_id:
        raise CapSolverError("CapSolver tạo task nhưng không trả taskId.")
    return task_id


def get_task_result(api_key, task_id):
    """Hỏi kết quả một task: (status, solution|None).

    status là "ready", "processing" hoặc "failed".
    """
    data = _post("/getTaskResult", {"clientKey": api_key, "taskId": task_id})
    _check_error(data, "hỏi kết quả task")
    return data.get("status"), data.get("solution")


def solve_task(api_key, task, timeout_seconds=SOLVE_TIMEOUT_SECONDS):
    """Tạo task rồi poll tới khi có token (hoặc hết giờ/lỗi).

    Trả về token (chuỗi). Hết giờ -> TimeoutError; API lỗi ->
    CapSolverError; task failed -> CapSolverError.
    """
    task_type, token_keys = TASK_TYPES.get(
        _kind_of_task(task), ("", ("gRecaptchaResponse", "token"))
    )
    task_id = create_task(api_key, task)
    deadline = time.monotonic() + timeout_seconds
    while True:
        status, solution = get_task_result(api_key, task_id)
        if status == "ready":
            if isinstance(solution, dict):
                for key in token_keys:
                    token = solution.get(key)
                    if isinstance(token, str) and token:
                        return token
            raise CapSolverError(
                "CapSolver báo xong nhưng không có token trong kết quả."
            )
        if status == "failed":
            raise CapSolverError("CapSolver giải thất bại (task failed).")
        if time.monotonic() >= deadline:
            raise TimeoutError(
                f"CapSolver chưa giải xong sau {timeout_seconds} giây."
            )
        time.sleep(POLL_INTERVAL_SECONDS)


def _kind_of_task(task):
    """Đoán kind từ loại task (để biết khoá token nào cần đọc)."""
    task_type = task.get("type", "") if isinstance(task, dict) else ""
    for kind, (name, _keys) in TASK_TYPES.items():
        if name == task_type:
            return kind
    return ""


# --- Phát hiện + chèn token trên trang (chạy trong trình duyệt) ---

_DETECT_JS = """() => {
  const pick = (sel, attr) => {
    const el = document.querySelector(sel);
    return el ? (el.getAttribute(attr) || "") : "";
  };
  let key = pick(".g-recaptcha", "data-sitekey");
  if (key) return {kind: "recaptcha2", sitekey: key};
  key = pick(".h-captcha", "data-sitekey");
  if (key) return {kind: "hcaptcha", sitekey: key};
  key = pick(".cf-turnstile", "data-sitekey");
  if (key) return {kind: "turnstile", sitekey: key};
  const frames = Array.from(document.querySelectorAll("iframe"));
  for (const f of frames) {
    const src = f.src || "";
    if (src.includes("recaptcha/api2/anchor")) {
      const m = src.match(/[?&]k=([^&]+)/);
      if (m) return {kind: "recaptcha2", sitekey: decodeURIComponent(m[1])};
    }
    if (src.includes("hcaptcha.com/captcha")) {
      const m = src.match(/[?&]sitekey=([^&]+)/);
      if (m) return {kind: "hcaptcha", sitekey: decodeURIComponent(m[1])};
    }
    if (src.includes("funcaptcha.com") || src.includes("arkoselabs.com")) {
      const m = src.match(/[?&]pk=([^&]+)/);
      if (m) return {kind: "funcaptcha", sitekey: decodeURIComponent(m[1])};
    }
  }
  return null;
}"""


def detect_captcha(page):
    """Phát hiện CAPTCHA token-based trên trang.

    Trả về {"kind", "sitekey", "website_url"} hoặc None nếu
    không thấy loại nào quen thuộc. kind là một trong:
    recaptcha2, hcaptcha, turnstile, funcaptcha.
    """
    try:
        found = page.evaluate(_DETECT_JS)
    except Exception:
        return None
    if not isinstance(found, dict) or not found.get("sitekey"):
        return None
    try:
        website_url = page.url
    except Exception:
        website_url = ""
    return {
        "kind": found["kind"],
        "sitekey": found["sitekey"],
        "website_url": website_url,
    }


def _task_for(detected):
    """Task CapSolver tương ứng với loại CAPTCHA đã phát hiện."""
    kind = detected["kind"]
    task_type, _keys = TASK_TYPES[kind]
    task = {"type": task_type, "websiteURL": detected["website_url"]}
    if kind == "funcaptcha":
        task["websitePublicKey"] = detected["sitekey"]
    else:
        task["websiteKey"] = detected["sitekey"]
    return task


_INJECT_JS = """([kind, token]) => {
  const setField = (el) => {
    if (!el) return false;
    el.value = token;
    if ("innerHTML" in el) el.innerHTML = token;
    el.dispatchEvent(new Event("input", {bubbles: true}));
    el.dispatchEvent(new Event("change", {bubbles: true}));
    return true;
  };
  if (kind === "recaptcha2") {
    const ok = setField(document.getElementById("g-recaptcha-response"));
    const widget = document.querySelector(".g-recaptcha");
    const cb = widget && widget.getAttribute("data-callback");
    if (cb && window[cb]) { try { window[cb](token); } catch (e) {} }
    return ok;
  }
  if (kind === "hcaptcha") {
    return setField(document.getElementById("h-captcha-response"))
      || setField(document.querySelector('[name="h-captcha-response"]'));
  }
  if (kind === "turnstile") {
    return setField(document.querySelector('[name="cf-turnstile-response"]'));
  }
  if (kind === "funcaptcha") {
    return setField(document.querySelector('[name="fc-token"]'));
  }
  return false;
}"""


def inject_token(page, kind, token):
    """Chèn token đã giải vào ô response của CAPTCHA trên trang.

    Trả True khi chèn được vào một ô (trang sẽ đọc ô này lúc
    submit); False khi không tìm thấy ô nào — phía gọi vẫn thử
    bấm nút tiếp theo vì một số trang đọc token qua callback.
    """
    try:
        return bool(page.evaluate(_INJECT_JS, [kind, token]))
    except Exception:
        return False


def solve_page_captcha(page, api_key, timeout_seconds=SOLVE_TIMEOUT_SECONDS):
    """Giải CAPTCHA trên trang (nếu có) qua CapSolver.

    Trả về kind đã giải ("recaptcha2"/"hcaptcha"/"turnstile"/
    "funcaptcha"), hoặc None khi trang không có CAPTCHA quen
    thuộc nào. Lỗi giải (key sai, hết tiền, hết giờ) ném ra
    ngoài để phía gọi quyết định (luồng Meta: tài khoản đó
    thất bại, không chờ tay).
    """
    detected = detect_captcha(page)
    if detected is None:
        return None
    token = solve_task(api_key, _task_for(detected), timeout_seconds)
    inject_token(page, detected["kind"], token)
    return detected["kind"]
