# reg-muse-ai

## Muse.ai Account Helper

A small Windows desktop helper that registers Muse.ai accounts in a batch.
For each account it creates a temp-mail inbox (provider selectable —
tempmail.plus by default), opens Muse.ai in a dedicated
browser, retrieves the email verification code, fills the profile form, and —
only when the signup actually completes — saves the account to a local Excel
sheet. The window has two tabs — **Muse.ai** and **Meta (meta.ai)** — each
with its own run setup (proxy pool, lane count, connection, temp mail,
target count); the live process log, progress line and **Dừng** button are
shared below the tabs. CAPTCHA and any
additional registration steps remain manual; an account that stalls on one
times out, is logged as failed, and the batch moves on to the next address.

> tempmail.plus addresses are temporary email addresses, not Gmail. Muse.ai may
> reject temporary email addresses, and losing access to the mailbox can make
> password recovery impossible. Check Muse.ai's terms and use an email address
> you control for any account you need to keep.

## Run on Windows

1. Install Python 3.10 or later.
2. From this folder, install the dependency:

   ```powershell
   py -m pip install -r requirements.txt
   py -m playwright install chromium
   ```

3. Start the app:

   ```powershell
   py main.py
   ```

Paste a proxy into the **Proxy** field (required — see below), set the
number of accounts (1–100, or **0 to run until the proxy pool is used
up** — every proxy in the list serves its per-proxy quota once, burned
proxies included, then the run stops on its own), and click **Bắt đầu
đăng ký**. For each account
the app then, with no further clicks:

1. Creates a fresh inbox at the temp-mail provider chosen in the
   **Temp mail** dropdown (see "Temp mail providers" below). Depending
   on the provider the inbox may have a secret — a mail.tm password or
   a Guerrillamail token — which is kept for reading the
   code and written to the sheet; providers without a secret store
   nothing, because the address itself is the inbox key.
   (There is no Muse.ai password either — signup is email plus a
   one-time code.)
2. Opens a dedicated Chromium window **through the proxy** (or
   direct, in VPN mode), fills the email form and clicks
   **Continue**, waits for the code, fills the code form (it
   auto-submits the moment the code is fully entered, so filling
   it IS submitting it), then fills the profile form and clicks
   **Confirm** — the same button clicks as the manual flow.
   (The helper also has a native-auth API path — send-otp /
   confirm-otp / register-completion, called from inside the
   browser with an `X-Hatch-CSRF-Token` obtained from `POST
   /api/auth/native/restart`, plus a session bootstrap via a
   hidden form POST to `/hatch/session/bootstrap/` — but it is
   currently OFF: `USE_NATIVE_AUTH_API = False` in
   `src/browser_helper.py` (set 2026-11-13, because the API
   path's bootstrap navigates the browser through Meta callback
   URLs). Set it to True to re-enable it. Note the trade-off:
   the API path accepts temp-mail OTP codes, while the
   button-click flow does not — Meta rejects temp mail at the
   code-confirm step in the UI flow, which is why the API path
   was built in the first place. Logging in with an already
   registered Meta account works fine with button clicks, the
   same as doing it by hand.)
3. Polls the chosen temp-mail provider for the code (up to three minutes), which the
   code form then confirms.
4. On the final profile form, fills the first and last name as "Long" and
   "long" when the form asks for them (some regions, such as US IPs, only
   ask for the birthday — the name step is then skipped), and picks a
   random birthday with an age always over 18, then clicks **Confirm**.
   (On the API path, when enabled, the profile is completed with
   register-completion instead and this form never appears.)
5. Watches the screen after the profile form. A "Join the waitlist" screen
   means Muse is unavailable in the proxy exit's region and the attempt
   fails. On the card-entry screen ("Confirm your age"), if you have saved
   your card in the app (see below) the tool fills the card form and
   confirms for you; otherwise the batch pauses for up to four minutes for
   you to type the card and a US ZIP in the Chromium window yourself. Only
   after the card step is accepted are the account's email and mailbox
   secret appended to `accounts.xlsx` in the app's local Windows data folder;
   failed attempts are only written to the on-screen log, never to the
   sheet.
6. After the card step there is one last screen with a **Get started**
   button (sometimes Confirm/Continue): the tool clicks it — repeatedly
   if several such screens chain together — waits to land inside the
   Muse interface, and reads the session cookies. A successful row in `accounts.xlsx`
   therefore carries: creation time, email, the mailbox secret (the
   mail.tm password or the Guerrillamail session token; empty for providers
   where the address itself is the inbox key), the session cookie in the
   browser's header form (`name=value; name=value; …`, covering the
   muse.ai, meta.com and facebook.com cookies — passed from the helper
   to the app out of band and never printed in the log), and the
   account's own invite code (below). Sheets written by older versions
   are migrated on the next successful row: a fictional "Muse.ai
   password" column, if present, is dropped and the remaining columns
   shifted left; missing headers are filled in. If `accounts.xlsx`
   cannot be written because it is locked (usually because it is open
   in Excel), the row is kept in `accounts.pending.xlsx` next to it
   and merged into the main sheet on the next successful write — close
   the Excel window to unblock it.
7. Invite codes form a chain. Once inside the Muse interface, each
   account first redeems the *previous* account's invite code — the app
   passes it to the helper in `redeem.txt`, and the helper follows the
   UI path from the screenshots: **Cài đặt** (opened with its Ctrl+,
   shortcut, falling back to the bottom-left menu) →
   **Quy đổi mã lời mời** → type the 6 characters → **Xác nhận**. A
   failed redemption is logged but does not fail the account (the code
   can still be redeemed manually within 48 hours). The account then
   opens the **Mời** panel and reads its *own* invite code, which is
   saved in its sheet row and becomes the code the next account
   redeems. (Reading the code follows the real panel: the code
   sits in the panel's DOM in a span whose `title` is the code
   itself, under a **Scratch to reveal** canvas — the tool first
   reads it straight from the DOM. If it is not there yet, it
   paints the cover away with the held mouse — hovering turns
   the cursor into a round brush, and painting reveals the code
   and turns the copy button active — then clicks **Copy
   referral link** and parses the code out of what was copied.)
   The **Mã mời** field in the
   window is the chain's current
   head: whatever is there when a run starts (typically your own
   account's code) is redeemed by the run's first account, and the
   field is overwritten with each new account's code as the run
   proceeds; it persists between runs in a local `invite.txt`. You
   never have to type it: if the field is empty at Start, the app
   seeds the chain from the newest invite code already in
   `accounts.xlsx`, and if there is none, the first account simply
   redeems nothing and starts the chain with its own code.

   The **Chỉ dùng 1 mã mời này** ("use only this one invite
   code") checkbox next to the field switches the run to
   single-code mode: when it is ticked, **every** account of
   **every** lane redeems exactly the code in the Mã mời field
   (after the same empty-field sheet seeding above) and the chain
   never advances — each account's own harvested code is still
   saved in its sheet row, but it is never redeemed by the next
   account, and the field keeps showing your chosen code during
   and after the run. Unticked (the default), the chaining
   behavior described above applies unchanged.

The batch continues with the next address until the target count is reached
or you click **Dừng**. Click **Mở sheet Excel** to open the sheet with the
successfully registered accounts.

### Two tabs, two setups

Only one run happens at a time — whichever tab you press Start in runs
with **that tab's own setup**, and the other tab is locked until the run
ends. Each tab has its own proxy pool, bad-proxy list, lane count,
connection mode, temp-mail provider and target count:

- **Muse.ai** tab: proxy pool in `proxies.txt`, bad proxies in
  `bad-proxies.txt` (the files from before — your existing pool carries
  over untouched), plus the card row, per-card limit and invite-code
  chain. Accounts go to `accounts.xlsx`.
- **Meta (meta.ai)** tab: proxy pool in `proxies_meta.txt`, bad proxies
  in `bad-proxies-meta.txt`, plus the CapSolver key and selfie image.
  Accounts go to `meta_accounts.xlsx`. See "Meta account registration"
  below.

### Code layout

```
main.py                  # thin launcher: py main.py
src/
  paths.py               # repo root + app-data paths
  coordinator.py         # run engine: lanes, helper subprocess, events
  storage.py             # Excel sheets, keyring, run data files
  proxy_pool.py          # proxy rotation, bad-proxy lists, IP check
  browser_helper.py      # Playwright helper (Muse + Meta flows)
  capsolver.py           # CapSolver client (Meta CAPTCHAs)
  mail_providers.py      # temp-mail providers
  ui/
    main_window.py       # window: notebook + shared log/progress/stop
    setup_panel.py       # shared run-setup panel (one per tab)
    muse_tab.py          # Muse.ai tab (card + invite)
    meta_tab.py          # Meta tab (CapSolver + selfie)
```

Run data files (`proxies.txt`, `proxies_meta.txt`, `proxy.txt`,
`mailbox.txt`, `redeem.txt`, `selfie.txt`, ...) stay at the repo root
next to `main.py`, as before — only the code moved into `src/`.

### Temp mail providers

The **Temp mail** dropdown picks which disposable-mail service each run
uses; all four go through one shared module (`mail_providers.py`), and
the choice is per run — pick a provider, then start the batch.

- **tempmail.plus** — the default, and the only provider verified
  end-to-end with Meta so far: on 2026-11-06 Meta accepted its domains
  at signup and the verification email arrived and was auto-read.
  Six domains, a random one per account. Inboxes are public (no
  password); the random address itself is the key.
- **maildrop.cc** — a single domain, no account creation: the tool picks
  the local part and reads via their GraphQL API. Nothing is stored.
- **mail.tm** — password-protected mailboxes; the generated password is
  stored in the sheet's mailbox-password column. It currently has only
  **one** active public domain (fldubai.com), and it throttled the
  owner's own IP under bursts in the past (connection resets) — the
  service itself works, just don't hammer it.
- **Guerrillamail** — the service assigns the address (e.g.
  @guerrillamailblock.com) and a session token, which is stored in the
  sheet. A fresh inbox always contains a welcome message; the tool
  snapshots existing mail before requesting a code, so it is ignored
  automatically.

On 2026-11-09 every provider was put through a real-delivery test: a
real email was sent to a fresh inbox at each provider and read back
with this tool's own code. tempmail.plus (2s), mail.tm (5s),
maildrop.cc (19s — its API occasionally drops a connection; the
tool's polling retries on its own) and Guerrillamail (20s) all
received and read the mail. **tempmail.lol failed the same test** —
its API timed out and dropped connections continuously, and a live
run never retrieved a code with it — so it was **removed from the
choices**. Whether Meta accepts the domains of mail.tm, maildrop.cc
and Guerrillamail at signup can still only be proven by a real run.

When a signup fails, *where* it fails tells you what to fix. If the
attempt is rejected at the email step, or the verification code never
arrives, the mail domain/provider is the suspect — switch providers.
The post-profile block ("You can't create a new account right now") is
a different verdict: that is Meta refusing the proxy's exit IP, the
tool already retires that proxy on its own, and changing the mail
provider will not help there.

Temp-mail traffic always stays on the machine's direct connection,
whichever provider is chosen — only the Muse.ai browser traffic goes
through the proxy. Per-account mailbox state is passed to the helper in
`mailbox.txt` / `mailbox-<n>.txt` next to the scripts (gitignored, like
the proxy and redeem files).

### Connection: proxy pool or machine IP (VPN)

The **Kết nối** (Connection) dropdown picks how registration traffic
goes out: **Proxy (pool)** (default) or **IP máy (VPN)**. Each tab has
its own dropdown, proxy box and bad-proxy box — the Muse tab reads
`proxies.txt` / `bad-proxies.txt`, the Meta tab reads
`proxies_meta.txt` / `bad-proxies-meta.txt`. In proxy mode,
paste your proxies into the **Proxy** box in the tab, **one
proxy per line**: `host:port:username:password` (or a full
`http://username:password@host:port` URL). The pool is stored in the
tab's pool file next to the scripts and pre-fills the box on the next launch.
**Tài khoản / proxy** sets how many accounts each proxy serves: the count
advances by account position, whether the accounts pass or fail, and when
a proxy's quota is used up the run moves to the next proxy in the pool,
wrapping back to the first when the pool runs out. One exception: if Meta
answers an attempt with the "Something went wrong — You can't create a new
account right now" dialog, that IP is treated as blocked — the account
fails at once and the rest of that proxy's quota is skipped, so the next
account starts on the next proxy. The pool is expected to
be US proxies supplied by you: **the tool does not verify exit IPs** (that
check cost several page loads per account and was removed for speed at the
owner's direction). The backstop is the post-profile classification — if
an exit lands in an unsupported region, the "Join the waitlist" screen
still fails that account. Temp-mail traffic stays on the machine's direct
connection. The log only ever shows each proxy's host:port, never its
password. `proxies.txt` and `proxy.txt` are listed in `.gitignore`, so
they are never committed — keep it that way.

**Số luồng** (1–5, default 1) runs that many registration lanes in
parallel. The proxy pool is split by stride — lane k uses the proxies at
positions k, k+lanes, k+2·lanes, … — so no two lanes ever share a proxy
at the same time, and each lane rotates (and burns) within its own
slice. Each lane keeps its **own invite chain** with one shared start:
the **first account of every lane** redeems the code in the Mã mời
field (the code you set in the tool); from the second account on, each
lane continues with the codes its own accounts harvest, exactly as in
single-lane mode. All lanes share the **one
saved card**, and the per-card success limit counts globally across
lanes (slots are reserved atomically, so the limit is never overshot).
Lanes need at least as many proxies as lanes; with fewer proxies the
lane count is reduced to the pool size. Per-lane state lives in
`proxy-<n>.txt` / `redeem-<n>.txt` / `mailbox-<n>.txt` next to the scripts (also gitignored);
the spreadsheet is written under a lock so lanes cannot corrupt it.

The **Proxy lỗi** (bad proxy) box under the Proxy box is the exclusion
list. When a proxy's exit IP lands on the "Join the waitlist" screen or
Meta refuses account creation on it ("Something went wrong"), the app
moves that proxy into this box, stores it in `bad-proxies.txt`
(gitignored), retires it immediately for the rest of the run — it is
never rotated back to, in any lane — and filters it out of the pool at
the start of every later run. A 403 `registrationBlocked` from the
register-completion API counts as the same refusal (same "Something
went wrong" verdict, seen live 2026-11-12 on 4 of 5 lanes at once).
A proxy is identified by its full line
(credentials included): pool lines may share one host:port gateway
with a different user/password per exit session, and those count as
different proxies — only lines that are identical collapse into one.
The box is editable: delete a line there and that proxy is allowed
again from the next run.

**IP máy (VPN)** runs Chromium straight through the machine's own
connection — turn your VPN on before pressing **Bắt đầu**, and make
sure it exits in the US (the tool does not verify exit IPs; the
post-profile waitlist screen is the only backstop, same as in proxy
mode). Differences from proxy mode: the **Tài khoản / proxy** quota is
disabled, and all lanes share the machine's single IP — the lane count
runs exactly as set (the owner's decision; concurrent accounts from one
IP are more likely to be flagged), and if Meta blocks account creation
on this IP ("Something went wrong") or the run lands on "Join the
waitlist", the **whole run stops at once** and the log tells you to
switch VPN servers (US preferred) and run again. All accounts in a
direct-mode run come from one IP, so they are more visibly linked than
accounts spread across a proxy pool. Temp-mail traffic stays on the
machine's direct connection in both modes.

**IP check (direct mode):** the tool looks up the machine's current
public IP (ipinfo.io, falling back to ip-api.com — both free, no key)
and shows it next to **IP hiện tại** with its country and city. The
check runs automatically when a direct-mode run starts, and on demand
via the **Kiểm tra IP** button (usable even mid-run, e.g. if you
suspect the VPN dropped). If the IP exits outside the US, the log
warns you — but the run is **never blocked** by the check: a failed
lookup or a non-US IP only logs a warning and the run continues.

### Your card: entered once, kept in Windows Credential Manager

The **Thẻ** row in the window takes a card number, expiry (MM/YY), CVV and
ZIP. **Lưu thẻ** stores them in Windows Credential Manager — the same
OS-protected store the app already uses for its passwords. The card is
never written to a file, never added to the Excel sheet, and no digit of
it ever appears in the on-screen log. At the card step the signup browser
reads it back from Credential Manager, fills the form, verifies the number
landed intact, and clicks Confirm; if the fill fails for any reason, the
run falls back to waiting for manual entry. **Xoá thẻ** removes the saved
card. Leave the ZIP empty to use 10001.

Passwords are stored by Python `keyring` in the Windows Credential Manager.
The Excel sheet contains mailbox passwords and live session cookies in
readable form, so protect the file and do not share or sync it to an
untrusted location.

### Meta account registration (Đăng kí Meta)

The **Đăng kí Meta** button runs a separate, Meta-only registration: it
creates Meta accounts at meta.ai (temp mail + OTP), saves each one to
`meta_accounts.xlsx`, and **stops there — it never goes to Muse**.
Logging into Muse with a saved Meta account is a separate step. The
point of the split: a Meta account arriving at Muse goes through
*login*, not registration, so it skips the register-completion step
where temp-mail signups get blocked (`registrationBlocked`).

The flow, as observed by hand (2026-11-13): meta.ai → **Log in** →
**Use mobile number or email** → enter the temp-mail address →
birthday (over 18, random) + a random password → keep the
auto-generated username, **Confirm** → enter the OTP code from the
mailbox → CAPTCHA + selfie upload → account created. After the OTP
step there are two variants: the page may return to the meta.ai home
screen (click **Log in** again to reach the CAPTCHA) or go straight
to the CAPTCHA — the tool handles both.

There are **no manual steps** in the Meta flow. That needs two things
set up first, in the **Meta** tab:

- **CapSolver API key** — CAPTCHAs are solved automatically through
  [CapSolver](https://www.capsolver.com/) (reCAPTCHA v2, hCaptcha,
  Cloudflare Turnstile, FunCaptcha; ProxyLess tasks, so your proxies
  are not shared with the service). Create an account there, add
  balance, paste the key and press **Lưu key** — it is stored in
  Windows Credential Manager, never in a file or in the log.
  **Số dư** checks the remaining balance. CapSolver is pay-per-use
  (roughly $0.5–$2 per 1,000 simple solves; check their current
  pricing). If a CAPTCHA cannot be solved (bad key, no balance,
  timeout), that account fails — the run never waits for manual input.
- **Ảnh selfie** — one selfie image file, uploaded automatically at
  the selfie step for every account in the run. Pick it with
  **Chọn ảnh…**; the path is kept in gitignored `selfie.txt`.

The Meta password is generated randomly for each account's form and
is **not saved anywhere** (Muse login later uses an email OTP code,
not the password). Each saved row also carries a `Link lấy OTP`
column holding the inbox link for the account's mailbox (a direct
inbox link for tempmail.plus; the inbox page for the other
providers). Saved accounts go to
`%LOCALAPPDATA%\MuseAIHelper\meta_accounts.xlsx` with columns
Thời gian tạo, Email Meta, Mật khẩu hộp thư, Username Meta,
Nhà cung cấp, Muse, Link lấy OTP. The Meta
run reuses the proxy pool / VPN mode, lane count, temp-mail provider
and account target of the main window; it does not use the saved card
or invite codes.

Note on CAPTCHAs: the Muse signup flow does not bypass CAPTCHA — a
challenge there waits for you to solve it in the browser. The Meta
flow solves its CAPTCHAs through CapSolver, a paid third-party
solving service, using your own API key and balance.

### Muse login with a Meta account (Đăng nhập Muse)

The **Đăng nhập Muse** button (Muse.ai tab) logs into Muse with Meta
accounts you already registered: pick the Meta Excel file with
**Chọn file…** (the path is kept in gitignored `meta_excel.txt`),
and the run reads each row's Email Meta / Mật khẩu hộp thư /
Nhà cung cấp columns — **no new mailbox is created**. Each account
logs in with an email OTP, then submits the saved card at the card
step, exactly like a signup from there on. After the OTP, Meta may
show a "Confirm you're human" checkpoint
(auth.meta.com/checkpoints/...) — the helper clicks Continue and
solves the CAPTCHA through CapSolver (using the key saved on the
Meta tab), or waits up to 4 minutes for you to solve it by hand if
no key is saved; the flow then continues to the card step (or the
profile form).
A successful login is saved to `accounts.xlsx` — email, mailbox
secret, session cookie and invite code as usual, plus the mailbox's
`Link lấy OTP` — and marked in the Meta file's **Muse** column, so
the next run skips it. (Muse accounts sign in with an email OTP, so
there is no Muse password to save.) No invite
code is redeemed in login mode. The run uses the Muse.ai tab's proxy
pool / VPN mode, lane count and account target (0 = log in all
pending rows). The card cap is not enforced in login runs. The
Muse.ai tab has no Temp mail row — the mail comes from the Meta file
(the old direct signup, **Bắt đầu đăng ký**, still creates fresh
tempmail.plus mailboxes).
