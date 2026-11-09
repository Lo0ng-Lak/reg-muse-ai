"""Các dịch vụ temp mail cho tool đăng ký Muse.ai.

Một module dùng chung cho cả main.py (tạo hộp thư) và
browser_helper.py (chờ mã xác minh), để hai phía không bao
giờ lệch nhau về cách nói chuyện với dịch vụ mail.

Mỗi hộp thư là một dict:
    {"provider": <id>, "email": <địa chỉ>, "secret": <bí mật>}

"secret" là thứ cần để đọc lại hộp thư về sau:
- mail.tm: mật khẩu hộp thư (tự sinh 16 ký tự khi tạo).
- tempmail.lol: token của hộp thư.
- guerrillamail: sid_token của phiên.
- tempmail.plus và maildrop.cc: "" — địa chỉ tự nó là khoá
  hộp thư, không có bí mật nào.

list_messages() trả danh sách [{"id": ..., "subject": str}]
mới nhất trước; read_message() trả toàn văn (dạng text) của
một thư. Mọi lỗi HTTP/kết nối đều ném RuntimeError: phía tạo
hộp thư coi đó là "tạo thất bại", phía chờ mã coi đó là "thử
lại ở vòng chờ tiếp theo" cho tới hết giờ — đúng như cách
code tempmail.plus cũ vẫn xử lý.
"""

import html
import http.client
import json
import secrets
import string
import urllib.error
import urllib.parse
import urllib.request
from html.parser import HTMLParser


USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/130.0.0.0 Safari/537.36"
)
HTTP_TIMEOUT_SECONDS = 20

# Thứ tự này là thứ tự trong ô chọn "Temp mail" của app.
PROVIDERS = [
    ("tempmail.plus", "tempmail.plus"),
    ("tempmail.lol", "tempmail.lol"),
    ("maildrop.cc", "maildrop.cc"),
    ("mail.tm", "mail.tm"),
    ("guerrillamail", "Guerrillamail"),
]
DEFAULT_PROVIDER = "tempmail.plus"

TEMPMAIL_PLUS_BASE = "https://tempmail.plus"
# Domain hộp thư của tempmail.plus (đọc từ trang chủ của họ,
# 2026-11-06). Dịch vụ không có API liệt kê domain nên danh
# sách nằm tại đây; "tạo hộp thư" thử lần lượt từng domain.
TEMPMAIL_PLUS_DOMAINS = [
    "mailto.plus",
    "fexpost.com",
    "fexbox.org",
    "fextemp.com",
    "merepost.com",
    "rover.info",
]
TEMPMAIL_LOL_BASE = "https://api.tempmail.lol/v2"
MAILDROP_GRAPHQL_URL = "https://api.maildrop.cc/graphql"
MAILTM_BASE = "https://api.mail.tm"
GUERRILLA_BASE = "https://api.guerrillamail.com/ajax.php"


def _http_json(method, url, payload=None, headers=None, provider="temp mail"):
    """Một cổng HTTP duy nhất cho mọi dịch vụ: trả JSON đã parse.

    Trả về dict/list (hoặc {} khi thân rỗng). Mọi lỗi đều ném
    RuntimeError mang tên dịch vụ để log đọc được ngay.
    """
    request_headers = {
        "Accept": "application/json",
        "User-Agent": USER_AGENT,
    }
    data = None
    if payload is not None:
        request_headers["Content-Type"] = "application/json"
        data = json.dumps(payload).encode("utf-8")
    if headers:
        request_headers.update(headers)
    request = urllib.request.Request(
        url, data=data, headers=request_headers, method=method
    )
    try:
        with urllib.request.urlopen(
            request, timeout=HTTP_TIMEOUT_SECONDS
        ) as response:
            body = response.read()
    except urllib.error.HTTPError as error:
        detail = error.read().decode("utf-8", errors="replace")
        raise RuntimeError(
            f"{provider} trả lỗi HTTP {error.code}: {detail[:200]}"
        ) from error
    except (
        urllib.error.URLError,
        TimeoutError,
        ConnectionError,
        http.client.HTTPException,
    ) as error:
        # ConnectionError/http.client.HTTPException: phía máy
        # chủ ngắt kết nối giữa chừng (reset) — urllib không
        # luôn gói nó vào URLError. Cũng là lỗi kết nối, xử lý
        # như nhau: nơi gọi tự thử lại theo nhịp của nó.
        raise RuntimeError(
            f"Không kết nối được {provider}: {error}"
        ) from error
    if not body:
        return {}
    try:
        return json.loads(body.decode("utf-8"))
    except (json.JSONDecodeError, UnicodeDecodeError) as error:
        raise RuntimeError(
            f"{provider} trả dữ liệu không phải JSON."
        ) from error


class _HtmlTextExtractor(HTMLParser):
    def __init__(self):
        super().__init__()
        self.parts = []

    def handle_data(self, data):
        value = data.strip()
        if value:
            self.parts.append(value)


def _html_to_text(value):
    """Bóc chữ khỏi HTML (một chuỗi hoặc danh sách chuỗi)."""
    if isinstance(value, list):
        value = "\n".join(str(item) for item in value)
    parser = _HtmlTextExtractor()
    parser.feed(str(value or ""))
    return html.unescape("\n".join(parser.parts))


def _message_text(detail):
    """Text của một thư chi tiết dạng mail.tm/tempmail.plus."""
    text = detail.get("text")
    if text:
        return text
    html_body = detail.get("html") or []
    if not html_body:
        return detail.get("intro", "") or ""
    return _html_to_text(html_body)


def _detail_text(detail):
    """Bóc toàn bộ chữ đọc được trong chi tiết thư.

    Ưu tiên các trường quen thuộc (text/html/intro); nếu vẫn
    rỗng thì gom mọi chuỗi trong JSON — bỏ qua các trường id
    để không nhặt nhầm số id làm mã xác minh.
    """
    if not isinstance(detail, dict):
        return str(detail or "")
    text = _message_text(detail)
    if text:
        return text
    parts = []

    def collect(value):
        if isinstance(value, str):
            parts.append(value)
        elif isinstance(value, dict):
            for key, item in value.items():
                if "id" in key.lower():
                    continue
                collect(item)
        elif isinstance(value, list):
            for item in value:
                collect(item)

    collect(detail)
    return "\n".join(parts)


def _members(response):
    """Danh sách thành viên từ API kiểu mail.tm: chấp cả hai
    dạng hydra ({"hydra:member": [...]}) lẫn mảng trần."""
    if isinstance(response, list):
        return response
    if isinstance(response, dict):
        members = response.get("hydra:member")
        if isinstance(members, list):
            return members
    return []


def _quoted(value):
    """Chuỗi cho truy vấn GraphQL của maildrop (trích dẫn an toàn)."""
    return json.dumps(str(value))


# ----------------------------------------------------------------
# tempmail.plus — mặc định, đã kiểm chứng nhận OTP từ Meta
# (2026-11-06). Không đăng ký, không mật khẩu: mọi địa chỉ
# thuộc các domain của họ đều nhận thư ngay, API đọc thư
# bằng chính địa chỉ đó.
# ----------------------------------------------------------------


def _plus_create():
    last_error = None
    # Bắt đầu từ một domain ngẫu nhiên mỗi lần: các tài khoản
    # tản đều trên cả 6 domain thay vì dồn hết vào domain đầu,
    # và nếu Meta chê một domain nào đó thì chỉ phần tài khoản
    # rơi vào domain đó bị ảnh hưởng.
    start = secrets.randbelow(len(TEMPMAIL_PLUS_DOMAINS))
    ordered = TEMPMAIL_PLUS_DOMAINS[start:] + TEMPMAIL_PLUS_DOMAINS[:start]
    for domain in ordered:
        address = f"muse{secrets.token_hex(5)}@{domain}"
        try:
            response = _http_json(
                "GET",
                TEMPMAIL_PLUS_BASE
                + "/api/mails?email="
                + urllib.parse.quote(address)
                + "&limit=1",
                provider="tempmail.plus",
            )
        except RuntimeError as error:
            last_error = error
            continue
        if isinstance(response, dict) and response.get("result") is False:
            last_error = RuntimeError(
                "tempmail.plus từ chối địa chỉ hộp thư."
            )
            continue
        return {"provider": "tempmail.plus", "email": address, "secret": ""}
    raise last_error or RuntimeError(
        "Không tạo được hộp thư tempmail.plus."
    )


def _plus_list(mailbox):
    response = _http_json(
        "GET",
        TEMPMAIL_PLUS_BASE
        + "/api/mails?email="
        + urllib.parse.quote(mailbox.get("email") or "")
        + "&limit=20",
        provider="tempmail.plus",
    )
    items = (
        response.get("mail_list") if isinstance(response, dict) else None
    ) or []
    result = []
    for item in items:
        if not isinstance(item, dict):
            continue
        message_id = item.get("mail_id") or item.get("id")
        if message_id:
            result.append(
                {"id": message_id, "subject": item.get("subject") or ""}
            )
    return result


def _plus_read(mailbox, message_id):
    detail = _http_json(
        "GET",
        TEMPMAIL_PLUS_BASE
        + f"/api/mails/{message_id}?email="
        + urllib.parse.quote(mailbox.get("email") or ""),
        provider="tempmail.plus",
    )
    return _detail_text(detail)


# ----------------------------------------------------------------
# tempmail.lol — tạo hộp thư bằng một POST, nhận về địa chỉ +
# token; đọc thư bằng token đó. Domain do dịch vụ tự gán.
# ----------------------------------------------------------------


def _lol_create():
    response = _http_json(
        "POST",
        TEMPMAIL_LOL_BASE + "/inbox/create",
        payload={},
        provider="tempmail.lol",
    )
    address = response.get("address") if isinstance(response, dict) else None
    token = response.get("token") if isinstance(response, dict) else None
    if not address or not token:
        raise RuntimeError(
            "tempmail.lol không trả về địa chỉ/token hộp thư."
        )
    return {"provider": "tempmail.lol", "email": address, "secret": token}


def _lol_inbox(mailbox):
    response = _http_json(
        "GET",
        TEMPMAIL_LOL_BASE
        + "/inbox?token="
        + urllib.parse.quote(mailbox.get("secret") or ""),
        provider="tempmail.lol",
    )
    if isinstance(response, list):
        return response
    if isinstance(response, dict):
        return response.get("emails") or response.get("messages") or []
    return []


def _lol_item_id(item, index):
    message_id = item.get("id") or item.get("_id")
    if message_id:
        return message_id
    # Phòng khi API không kèm id: ghép từ các trường ổn định
    # để cùng một thư luôn cho cùng một id giữa các vòng chờ.
    return (
        f"lol:{index}:{item.get('date') or ''}"
        f":{item.get('subject') or ''}"
    )


def _lol_list(mailbox):
    result = []
    for index, item in enumerate(_lol_inbox(mailbox)):
        if isinstance(item, dict):
            result.append(
                {
                    "id": _lol_item_id(item, index),
                    "subject": item.get("subject") or "",
                }
            )
    return result


def _lol_read(mailbox, message_id):
    for index, item in enumerate(_lol_inbox(mailbox)):
        if not isinstance(item, dict):
            continue
        if _lol_item_id(item, index) != message_id:
            continue
        # Hộp thư của tempmail.lol trả cả nội dung trong danh
        # sách thư nên đọc là bóc chữ ngay từ mục đó.
        for key in ("text", "body_text", "body"):
            value = item.get(key)
            if value:
                value = str(value)
                if "<" in value and ">" in value:
                    return _html_to_text(value)
                return value
        if item.get("html"):
            return _html_to_text(item["html"])
        return _detail_text(item)
    return ""


# ----------------------------------------------------------------
# maildrop.cc — một domain duy nhất, không tạo tài khoản:
# phần trước @ là của mình tự chọn, API GraphQL đọc thư
# theo đúng phần đó.
# ----------------------------------------------------------------


def _maildrop_query(query):
    response = _http_json(
        "POST",
        MAILDROP_GRAPHQL_URL,
        payload={"query": query},
        provider="maildrop.cc",
    )
    data = response.get("data") if isinstance(response, dict) else None
    if data is None:
        errors = (
            response.get("errors") if isinstance(response, dict) else None
        )
        raise RuntimeError(f"maildrop.cc trả lỗi GraphQL: {errors}")
    return data


def _maildrop_local(mailbox):
    return (mailbox.get("email") or "").split("@")[0]


def _maildrop_create():
    local = f"muse{secrets.token_hex(5)}"
    # Không có bước tạo tài khoản; gọi thử một lần để chắc
    # chắn dịch vụ đang trả lời, giống nhịp thăm dò của
    # tempmail.plus — lỗi ở đây là tài khoản tính thất bại
    # ngay, không phải chờ tới bước mã xác minh mới chết.
    _maildrop_query(
        "{ inbox(mailbox: " + _quoted(local) + ") { id } }"
    )
    return {
        "provider": "maildrop.cc",
        "email": f"{local}@maildrop.cc",
        "secret": "",
    }


def _maildrop_list(mailbox):
    data = _maildrop_query(
        "{ inbox(mailbox: "
        + _quoted(_maildrop_local(mailbox))
        + ") { id subject date } }"
    )
    items = [
        item
        for item in (data.get("inbox") or [])
        if isinstance(item, dict) and item.get("id")
    ]
    # Sắp mới nhất trước theo ngày (chuỗi ISO so được trực tiếp).
    items.sort(key=lambda item: item.get("date") or "", reverse=True)
    return [
        {"id": item["id"], "subject": item.get("subject") or ""}
        for item in items
    ]


def _maildrop_read(mailbox, message_id):
    data = _maildrop_query(
        "{ message(mailbox: "
        + _quoted(_maildrop_local(mailbox))
        + ", id: "
        + _quoted(message_id)
        + ") { subject data html } }"
    )
    message = data.get("message") or {}
    if message.get("html"):
        # Bản HTML là thân thư đã dựng, không lẫn tiêu đề thô
        # (nơi có thể chứa dãy số giống mã) — ưu tiên nó.
        return _html_to_text(message["html"])
    return str(message.get("data") or "")


# ----------------------------------------------------------------
# mail.tm — hộp thư có mật khẩu thật: tạo tài khoản, lấy token
# Bearer rồi đọc thư. Danh sách domain lấy từ API (không ghi
# cứng); tại thời điểm viết chỉ có một domain công khai đang
# hoạt động là fldubai.com. Dịch vụ này từng throttle IP của
# chính người dùng khi gọi dồn dập (reset kết nối).
# ----------------------------------------------------------------

_mailtm_tokens = {}


def _mailtm_create():
    response = _http_json(
        "GET", MAILTM_BASE + "/domains", provider="mail.tm"
    )
    domains = []
    for item in _members(response):
        if not isinstance(item, dict):
            continue
        domain = item.get("domain")
        if (
            domain
            and item.get("isActive") is not False
            and item.get("isPrivate") is not True
        ):
            domains.append(domain)
    if not domains:
        raise RuntimeError(
            "mail.tm không có domain công khai nào đang hoạt động."
        )
    domain = secrets.choice(domains)
    address = f"muse{secrets.token_hex(5)}@{domain}"
    alphabet = string.ascii_letters + string.digits
    password = "".join(secrets.choice(alphabet) for _ in range(16))
    created = _http_json(
        "POST",
        MAILTM_BASE + "/accounts",
        payload={"address": address, "password": password},
        provider="mail.tm",
    )
    if not isinstance(created, dict) or not created.get("address"):
        raise RuntimeError("mail.tm không tạo được tài khoản hộp thư.")
    mailbox = {"provider": "mail.tm", "email": address, "secret": password}
    # Đăng nhập thử ngay: tài khoản tạo xong mà không lấy
    # được token thì vô dụng — tính là tạo thất bại.
    _mailtm_token(mailbox)
    return mailbox


def _mailtm_token(mailbox):
    email = mailbox.get("email") or ""
    token = _mailtm_tokens.get(email)
    if token:
        return token
    response = _http_json(
        "POST",
        MAILTM_BASE + "/token",
        payload={
            "address": email,
            "password": mailbox.get("secret") or "",
        },
        provider="mail.tm",
    )
    token = response.get("token") if isinstance(response, dict) else None
    if not token:
        raise RuntimeError("mail.tm không cấp được token hộp thư.")
    _mailtm_tokens[email] = token
    return token


def _mailtm_get(path, mailbox):
    def attempt():
        return _http_json(
            "GET",
            MAILTM_BASE + path,
            headers={
                "Authorization": f"Bearer {_mailtm_token(mailbox)}"
            },
            provider="mail.tm",
        )

    try:
        return attempt()
    except RuntimeError as error:
        if "HTTP 401" not in str(error):
            raise
        # Token hết hạn/không còn hợp lệ: bỏ cache, lấy token
        # mới và thử lại đúng một lần.
        _mailtm_tokens.pop(mailbox.get("email") or "", None)
        return attempt()


def _mailtm_list(mailbox):
    result = []
    for item in _members(_mailtm_get("/messages", mailbox)):
        if isinstance(item, dict) and item.get("id"):
            result.append(
                {"id": item["id"], "subject": item.get("subject") or ""}
            )
    return result


def _mailtm_read(mailbox, message_id):
    detail = _mailtm_get(
        "/messages/"
        + urllib.parse.quote(str(message_id), safe=""),
        mailbox,
    )
    return _detail_text(detail)


# ----------------------------------------------------------------
# Guerrillamail — dịch vụ tự gán địa chỉ (ví dụ
# @guerrillamailblock.com) và một sid_token cho phiên; đọc
# thư bằng sid_token đó. Hộp thư luôn có sẵn một thư chào
# mừng — phía chờ mã chụp mốc thư cũ trước khi gửi email
# nên thư đó tự bị bỏ qua.
# ----------------------------------------------------------------


def _guerrilla_call(params):
    return _http_json(
        "GET",
        GUERRILLA_BASE + "?" + urllib.parse.urlencode(params),
        provider="Guerrillamail",
    )


def _guerrilla_create():
    response = _guerrilla_call({"f": "get_email_address", "lang": "en"})
    address = (
        response.get("email_addr") if isinstance(response, dict) else None
    )
    sid_token = (
        response.get("sid_token") if isinstance(response, dict) else None
    )
    if not address or not sid_token:
        raise RuntimeError(
            "Guerrillamail không trả về địa chỉ/phiên hộp thư."
        )
    return {
        "provider": "guerrillamail",
        "email": address,
        "secret": sid_token,
    }


def _guerrilla_list(mailbox):
    response = _guerrilla_call(
        {
            "f": "check_email",
            "seq": 0,
            "sid_token": mailbox.get("secret") or "",
        }
    )
    items = (
        response.get("list") if isinstance(response, dict) else None
    ) or []
    result = []
    for item in items:
        if isinstance(item, dict) and item.get("mail_id") is not None:
            result.append(
                {
                    "id": item["mail_id"],
                    "subject": item.get("mail_subject") or "",
                }
            )
    return result


def _guerrilla_read(mailbox, message_id):
    response = _guerrilla_call(
        {
            "f": "fetch_email",
            "email_id": message_id,
            "sid_token": mailbox.get("secret") or "",
        }
    )
    body = (
        response.get("mail_body") if isinstance(response, dict) else None
    )
    return _html_to_text(body) if body else ""


# ----------------------------------------------------------------
# Điều phối theo id dịch vụ.
# ----------------------------------------------------------------

_CREATE = {
    "tempmail.plus": _plus_create,
    "tempmail.lol": _lol_create,
    "maildrop.cc": _maildrop_create,
    "mail.tm": _mailtm_create,
    "guerrillamail": _guerrilla_create,
}
_LIST = {
    "tempmail.plus": _plus_list,
    "tempmail.lol": _lol_list,
    "maildrop.cc": _maildrop_list,
    "mail.tm": _mailtm_list,
    "guerrillamail": _guerrilla_list,
}
_READ = {
    "tempmail.plus": _plus_read,
    "tempmail.lol": _lol_read,
    "maildrop.cc": _maildrop_read,
    "mail.tm": _mailtm_read,
    "guerrillamail": _guerrilla_read,
}


def _dispatch(table, provider_id, action):
    handler = table.get(provider_id or DEFAULT_PROVIDER)
    if handler is None:
        raise RuntimeError(
            f"Dịch vụ temp mail không hỗ trợ cho {action}: {provider_id}"
        )
    return handler


def create_mailbox(provider_id):
    """Tạo một hộp thư mới ở dịch vụ đã chọn.

    Trả về {"provider", "email", "secret"}. Lỗi dịch vụ/mạng
    ném RuntimeError — nơi gọi coi là tạo hộp thư thất bại.
    """
    return _dispatch(_CREATE, provider_id, "tạo hộp thư")()


def list_messages(mailbox):
    """Thư trong hộp thư, mới nhất trước: [{"id", "subject"}]."""
    return _dispatch(_LIST, mailbox.get("provider"), "đọc hộp thư")(
        mailbox
    )


def read_message(mailbox, message_id):
    """Toàn văn (text) của một thư trong hộp thư."""
    return _dispatch(_READ, mailbox.get("provider"), "đọc thư")(
        mailbox, message_id
    )
