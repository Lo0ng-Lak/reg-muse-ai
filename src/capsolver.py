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

Tài liệu API (đối chiếu 2026-11-13, docs.capsolver.com/en):
- POST https://api.capsolver.com/createTask {clientKey, task}
  -> {errorId: 0, taskId} hoặc {errorId: 1, errorCode, ...}
- POST https://api.capsolver.com/getTaskResult {clientKey, taskId}
  -> {errorId: 0, status: "ready"|"processing", solution}
- POST https://api.capsolver.com/getBalance {clientKey}
  -> {errorId: 0, balance}
- Tham số anchor: URL iframe /anchor có thể mang `s` (và `sa`).
  Enterprise -> enterprisePayload {"s": ...}; v2 thường ->
  recaptchaDataSValue (s) + pageAction (sa). Thiếu thì task có
  thể không giải được — nên phát hiện từ iframe trước, div sau.
- Callback: sau khi chèn token phải gọi callback của trang
  (Meta dùng hàm ẩn danh trong grecaptcha.enterprise.render —
  không có tên global, không có data-callback — nên tìm trong
  window.___grecaptcha_cfg.clients).
- Solution có thể kèm "recaptcha-ca-e" (v2) — đặt làm cookie
  trên domain của trang đang mở.
Loại task: ReCaptchaV2TaskProxyLess,
ReCaptchaV2EnterpriseTaskProxyLess, HCaptchaTaskProxyless,
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
    # reCAPTCHA Enterprise (Meta dùng ở checkpoint "Confirm you're
    # human", live 2026-11-13): ô tick giống v2 nhưng iframe là
    # recaptcha/enterprise/anchor, token phải giải bằng task
    # Enterprise riêng — task v2 thường không qua được.
    "recaptcha2_enterprise": (
        "ReCaptchaV2EnterpriseTaskProxyLess",
        ("gRecaptchaResponse", "token"),
    ),
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


# Mã lỗi CapSolver hay gặp (docs.capsolver.com/en/guide/api-error)
# -> chú thích tiếng Việt để log đọc được ngay.
_ERROR_HINTS = {
    "ERROR_KEY_DENIED_ACCESS": "API key sai — kiểm tra lại key đã lưu",
    "ERROR_ZERO_BALANCE": "tài khoản CapSolver hết tiền — nạp thêm",
    "ERROR_CAPTCHA_UNSOLVABLE": (
        "CAPTCHA không nhận ra (không bị trừ tiền) — thử lại"
    ),
    "ERROR_TASK_NOT_SUPPORTED": (
        "loại task không được hỗ trợ — kiểm tra tên task"
    ),
    "ERROR_TASK_TIMEOUT": "giải quá 120 giây trên server CapSolver",
    "ERROR_RATE_LIMIT": "gọi API quá nhanh — chờ rồi thử lại",
    "ERROR_INVALID_TASK_DATA": "dữ liệu task sai — kiểm tra sitekey/URL",
    "ERROR_KEY_TEMP_BLOCKED": (
        "key bị khoá tạm 5 phút do quá nhiều lỗi — thử lại sau"
    ),
}


def _check_error(data, action):
    """Ném lỗi nếu CapSolver trả errorId != 0."""
    if not isinstance(data, dict) or data.get("errorId") != 0:
        code = data.get("errorCode") if isinstance(data, dict) else None
        description = (
            data.get("errorDescription") if isinstance(data, dict) else None
        )
        detail = description or code or "lỗi không rõ"
        hint = _ERROR_HINTS.get(code or "")
        if hint:
            detail = f"{detail} ({hint})"
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


def solve_task_full(api_key, task, timeout_seconds=SOLVE_TIMEOUT_SECONDS,
                    on_progress=None, progress_every_seconds=30):
    """Tạo task rồi poll tới khi có solution đầy đủ (hoặc lỗi).

    Trả về dict solution (chứa token + có thể kèm recaptcha-ca-e
    / recaptcha-ca-t để đặt cookie). Hết giờ -> TimeoutError;
    API lỗi -> CapSolverError; task failed -> CapSolverError.
    on_progress (nếu có) được gọi mỗi progress_every_seconds
    giây với số giây đã chờ — để phía gọi in nhịp log trong lúc
    chờ, tránh nhìn như tool đứng yên.
    """
    task_id = create_task(api_key, task)
    started = time.monotonic()
    deadline = started + timeout_seconds
    next_progress = started + progress_every_seconds
    while True:
        status, solution = get_task_result(api_key, task_id)
        if status == "ready":
            if isinstance(solution, dict):
                return solution
            raise CapSolverError(
                "CapSolver báo xong nhưng kết quả không phải dict."
            )
        if status == "failed":
            raise CapSolverError("CapSolver giải thất bại (task failed).")
        if on_progress is not None and time.monotonic() >= next_progress:
            on_progress(int(time.monotonic() - started))
            next_progress += progress_every_seconds
        if time.monotonic() >= deadline:
            raise TimeoutError(
                f"CapSolver chưa giải xong sau {timeout_seconds} giây."
            )
        time.sleep(POLL_INTERVAL_SECONDS)


def solve_task(api_key, task, timeout_seconds=SOLVE_TIMEOUT_SECONDS):
    """Tạo task rồi poll tới khi có token (hoặc hết giờ/lỗi).

    Trả về token (chuỗi). Hết giờ -> TimeoutError; API lỗi ->
    CapSolverError; task failed -> CapSolverError.
    """
    _task_type, token_keys = TASK_TYPES.get(
        _kind_of_task(task), ("", ("gRecaptchaResponse", "token"))
    )
    solution = solve_task_full(api_key, task, timeout_seconds)
    for key in token_keys:
        token = solution.get(key)
        if isinstance(token, str) and token:
            return token
    raise CapSolverError(
        "CapSolver báo xong nhưng không có token trong kết quả."
    )


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
  const param = (src, name) => {
    const m = src.match(new RegExp("[?&]" + name + "=([^&]+)"));
    return m ? decodeURIComponent(m[1]) : "";
  };
  // Đọc tham số anchor (k, s, sa) từ URL iframe: s/sa là tham số
  // CapSolver cần để giải (enterprisePayload / recaptchaDataSValue
  // / pageAction) — div .g-recaptcha không có các tham số này nên
  // iframe luôn được ưu tiên trước div.
  const anchor = (src, kind) => {
    const k = param(src, "k");
    if (!k) return null;
    const out = {kind: kind, sitekey: k, source: "anchor"};
    const s = param(src, "s");
    if (s) out.s = s;
    const sa = param(src, "sa");
    if (sa) out.sa = sa;
    return out;
  };
  const frames = Array.from(document.querySelectorAll("iframe"));
  // reCAPTCHA Enterprise trước: ô tick của nó cũng nằm trong
  // div .g-recaptcha, nhưng iframe là enterprise/anchor — phải
  // nhận ra Enterprise để giải bằng task Enterprise riêng.
  for (const f of frames) {
    const src = f.src || "";
    if (src.includes("recaptcha/enterprise/anchor")) {
      const found = anchor(src, "recaptcha2_enterprise");
      if (found) return found;
    }
  }
  for (const f of frames) {
    const src = f.src || "";
    if (src.includes("recaptcha/api2/anchor")) {
      const found = anchor(src, "recaptcha2");
      if (found) return found;
    }
  }
  let key = pick(".g-recaptcha", "data-sitekey");
  if (key) return {kind: "recaptcha2", sitekey: key, source: "div"};
  key = pick(".h-captcha", "data-sitekey");
  if (key) return {kind: "hcaptcha", sitekey: key, source: "div"};
  key = pick(".cf-turnstile", "data-sitekey");
  if (key) return {kind: "turnstile", sitekey: key, source: "div"};
  for (const f of frames) {
    const src = f.src || "";
    if (src.includes("hcaptcha.com/captcha")) {
      const m = src.match(/[?&]sitekey=([^&]+)/);
      if (m) return {kind: "hcaptcha", sitekey: decodeURIComponent(m[1]), source: "anchor"};
    }
    if (src.includes("funcaptcha.com") || src.includes("arkoselabs.com")) {
      const m = src.match(/[?&]pk=([^&]+)/);
      if (m) return {kind: "funcaptcha", sitekey: decodeURIComponent(m[1]), source: "anchor"};
    }
  }
  return null;
}"""


def _frames_of(page):
    """Tất cả các frame của trang (khung chính trước).

    Meta đặt widget CAPTCHA trong một khung con (live 2026-11-13
    14:31: ô hiện rõ mà quét khung chính không thấy gì) —
    detection phải quét từng frame, không chỉ khung chính.
    """
    try:
        frames = page.frames
    except Exception:
        frames = None
    if frames:
        return list(frames)
    return [page]


def _detect_in_frame(frame):
    """Phát hiện CAPTCHA trong một frame.

    Chỉ thấy div của reCAPTCHA thì chờ 2 giây rồi phát hiện lại
    một lần: iframe anchor thường tải ngay sau div, và nó mang
    tham số s mà task giải cần.
    """
    try:
        found = frame.evaluate(_DETECT_JS)
    except Exception:
        return None
    if (
        isinstance(found, dict)
        and found.get("source") == "div"
        and found.get("kind") in ("recaptcha2", "recaptcha2_enterprise")
    ):
        try:
            frame.wait_for_timeout(2000)
            second = frame.evaluate(_DETECT_JS)
        except Exception:
            second = None
        if isinstance(second, dict) and second.get("sitekey"):
            found = second
    if not isinstance(found, dict) or not found.get("sitekey"):
        return None
    return found


def detect_captcha(page):
    """Phát hiện CAPTCHA token-based trên trang.

    Quét từng frame (khung chính trước, rồi các khung con) —
    Meta đặt widget trong khung con. Trả về {"kind", "sitekey",
    "website_url", "source", "frame", ...} hoặc None nếu không
    thấy loại nào quen thuộc. kind là một trong: recaptcha2,
    recaptcha2_enterprise, hcaptcha, turnstile, funcaptcha.
    source là "anchor" (phát hiện từ iframe — kèm tham số s/sa
    nếu có) hoặc "div" (chỉ thấy div chứa sitekey, chưa thấy
    iframe). "frame" là frame chứa widget — token phải chèn vào
    đúng frame đó.
    """
    try:
        website_url = page.url
    except Exception:
        website_url = ""
    for frame in _frames_of(page):
        found = _detect_in_frame(frame)
        if found is None:
            continue
        result = {
            "kind": found["kind"],
            "sitekey": found["sitekey"],
            "website_url": website_url,
            "source": found.get("source") or "",
            "frame": frame,
        }
        for extra in ("s", "sa"):
            if found.get(extra):
                result[extra] = found[extra]
        return result
    return None


def _task_for(detected):
    """Task CapSolver tương ứng với loại CAPTCHA đã phát hiện.

    Tham số anchor (tài liệu CapSolver): Enterprise gửi kèm
    enterprisePayload {"s": ...} khi anchor có tham số s; v2
    thường gửi recaptchaDataSValue (s) + pageAction (sa).
    """
    kind = detected["kind"]
    task_type, _keys = TASK_TYPES[kind]
    task = {"type": task_type, "websiteURL": detected["website_url"]}
    if kind == "funcaptcha":
        task["websitePublicKey"] = detected["sitekey"]
    else:
        task["websiteKey"] = detected["sitekey"]
    if kind == "recaptcha2_enterprise":
        if detected.get("s"):
            task["enterprisePayload"] = {"s": detected["s"]}
    elif kind == "recaptcha2":
        if detected.get("s"):
            task["recaptchaDataSValue"] = detected["s"]
    if kind in ("recaptcha2", "recaptcha2_enterprise") and detected.get("sa"):
        task["pageAction"] = detected["sa"]
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
  if (kind === "recaptcha2" || kind === "recaptcha2_enterprise") {
    const ok = setField(document.getElementById("g-recaptcha-response"));
    // Chèn token xong phải BÁO cho trang biết — không thì nút
    // Continue của trang vẫn khoá. Thử lần lượt: (1) data-callback
    // trên div, (2) callback trong ___grecaptcha_cfg.clients
    // (trường hợp Meta: hàm ẩn danh truyền vào
    // grecaptcha.enterprise.render, không có tên global),
    // (3) tìm sâu tối đa 3 tầng trong client (callback có thể
    // lồng trong object con). Chỉ tìm sâu khi chưa gọi được gì.
    let called = false;
    const callIt = (fn) => {
      try { fn(token); called = true; } catch (e) {}
    };
    const widget = document.querySelector(".g-recaptcha");
    const cbName = widget && widget.getAttribute("data-callback");
    if (cbName && typeof window[cbName] === "function") callIt(window[cbName]);
    try {
      const cfg = window.___grecaptcha_cfg;
      if (cfg && cfg.clients) {
        const ids = Object.keys(cfg.clients);
        for (let i = 0; i < ids.length; i++) {
          const client = cfg.clients[ids[i]];
          if (!client) continue;
          const cb = client.callback;
          if (typeof cb === "function") callIt(cb);
          else if (typeof cb === "string" && typeof window[cb] === "function") callIt(window[cb]);
        }
        if (!called) {
          const seen = [];
          const deep = (obj, depth) => {
            if (!obj || depth > 3 || seen.indexOf(obj) !== -1) return null;
            seen.push(obj);
            if (typeof obj.callback === "function") return obj.callback;
            for (const k in obj) {
              let v;
              try { v = obj[k]; } catch (e) { continue; }
              if (v && typeof v === "object") {
                const found = deep(v, depth + 1);
                if (found) return found;
              }
            }
            return null;
          };
          for (let i = 0; i < ids.length; i++) {
            const found = deep(cfg.clients[ids[i]], 0);
            if (found) { callIt(found); break; }
          }
        }
      }
    } catch (e) {}
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


def inject_token(page, kind, token, frame=None):
    """Chèn token đã giải vào ô response của CAPTCHA.

    Token phải chèn vào đúng frame chứa widget (frame từ
    detect_captcha) — ô response + callback grecaptcha sống ở
    frame đó, không phải khung chính. frame=None thì chèn vào
    trang (tương thích code cũ).
    Trả True khi chèn được vào một ô (trang sẽ đọc ô này lúc
    submit); False khi không tìm thấy ô nào — phía gọi vẫn thử
    bấm nút tiếp theo vì một số trang đọc token qua callback.
    """
    target = frame if frame is not None else page
    try:
        return bool(target.evaluate(_INJECT_JS, [kind, token]))
    except Exception:
        return False


def _set_solution_cookies(page, solution, frame=None):
    """Đặt cookie từ solution (nếu có) lên domain chứa widget.

    Một số trang v2 cần cookie recaptcha-ca-e (CapSolver tự trả
    kèm khi có); v3 session mode trả recaptcha-ca-t. Cookie đặt
    trên host của frame chứa widget (khác host trang chính khi
    widget nằm trong khung con). Lỗi đặt cookie không làm hỏng
    việc giải — chỉ bỏ qua.
    """
    if not isinstance(solution, dict):
        return
    names = [name for name in ("recaptcha-ca-e", "recaptcha-ca-t")
             if solution.get(name)]
    if not names:
        return
    try:
        from urllib.parse import urlparse

        frame_url = ""
        if frame is not None:
            try:
                frame_url = frame.url
            except Exception:
                frame_url = ""
        domain = urlparse(frame_url or page.url).hostname
        cookies = [
            {
                "name": name,
                "value": solution[name],
                "domain": domain,
                "path": "/",
            }
            for name in names
        ]
        page.context.add_cookies(cookies)
    except Exception:
        pass


def solve_detected(page, api_key, detected,
                     timeout_seconds=SOLVE_TIMEOUT_SECONDS,
                     on_progress=None):
    """Giải một CAPTCHA đã phát hiện bằng detect_captcha: tạo
    task, chờ solution đầy đủ, đặt cookie đi kèm (nếu có), chèn
    token vào trang. Trả về kind đã giải. on_progress (nếu có)
    được gọi định kỳ trong lúc chờ CapSolver (xem solve_task_full).
    """
    solution = solve_task_full(api_key, _task_for(detected),
                               timeout_seconds,
                               on_progress=on_progress)
    _task_type, token_keys = TASK_TYPES.get(
        detected["kind"], ("", ("gRecaptchaResponse", "token"))
    )
    token = None
    for key in token_keys:
        value = solution.get(key)
        if isinstance(value, str) and value:
            token = value
            break
    if token is None:
        raise CapSolverError(
            "CapSolver báo xong nhưng không có token trong kết quả."
        )
    frame = detected.get("frame")
    _set_solution_cookies(page, solution, frame=frame)
    inject_token(page, detected["kind"], token, frame=frame)
    return detected["kind"]


def solve_page_captcha(page, api_key, timeout_seconds=SOLVE_TIMEOUT_SECONDS):
    """Giải CAPTCHA trên trang (nếu có) qua CapSolver.

    Trả về kind đã giải ("recaptcha2"/"recaptcha2_enterprise"/
    "hcaptcha"/"turnstile"/"funcaptcha"), hoặc None khi trang
    không có CAPTCHA quen thuộc nào. Lỗi giải (key sai, hết
    tiền, hết giờ) ném ra ngoài để phía gọi quyết định (luồng
    Meta: tài khoản đó thất bại, không chờ tay).
    """
    detected = detect_captcha(page)
    if detected is None:
        return None
    return solve_detected(page, api_key, detected, timeout_seconds)
