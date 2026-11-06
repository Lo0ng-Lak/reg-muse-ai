# reg-muse-ai

## Muse.ai Account Helper

A small Windows desktop helper that registers Muse.ai accounts in a batch.
For each account it creates a tempmail.plus inbox, opens Muse.ai in a dedicated
browser, retrieves the email verification code, fills the profile form, and —
only when the signup actually completes — saves the account to a local Excel
sheet. The window shows nothing but the live process log. CAPTCHA and any
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

1. Creates a random tempmail.plus address. tempmail.plus inboxes have no
   password — the address itself is the inbox key, so nothing is stored.
   (There is no Muse.ai password either — signup is email plus a
   one-time code.)
2. Opens a dedicated Chromium window **through the proxy**, first checks the
   exit IP it actually got, and only continues if that IP is in the US.
   It then fills the email into the visible "Mobile number or email" field
   and submits **Continue** to request a verification code.
3. Polls tempmail.plus for the code (up to three minutes), fills it, and submits
   **Confirm** on the code form.
4. On the final profile form, fills the first and last name as "Long" and
   "long" when the form asks for them (some regions, such as US IPs, only
   ask for the birthday — the name step is then skipped), and picks a
   random birthday with an age always over 18, then clicks **Confirm**.
5. Watches the screen after the profile form. A "Join the waitlist" screen
   means Muse is unavailable in the proxy exit's region and the attempt
   fails. On the card-entry screen ("Confirm your age"), if you have saved
   your card in the app (see below) the tool fills the card form and
   confirms for you; otherwise the batch pauses for up to four minutes for
   you to type the card and a US ZIP in the Chromium window yourself. Only
   after the card step is accepted are the account's email and password
   appended to `accounts.xlsx` in the app's local Windows data folder;
   failed attempts are only written to the on-screen log, never to the
   sheet.
6. After the card step there is one last screen with a **Get started**
   button (sometimes Confirm/Continue): the tool clicks it — repeatedly
   if several such screens chain together — waits to land inside the
   Muse interface, and reads the session cookies. A successful row in `accounts.xlsx`
   therefore carries: creation time, email, the tempmail.plus mailbox
   password (the only password there is), the session cookie in the
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

The batch continues with the next address until the target count is reached
or you click **Dừng**. Click **Mở sheet Excel** to open the sheet with the
successfully registered accounts.

### Proxy pool: required, US exit only

Registration only ever happens through a proxy — there is no VPN/direct
mode. Paste your proxies into the **Proxy** box in the app window, **one
proxy per line**: `host:port:username:password` (or a full
`http://username:password@host:port` URL). The pool is stored in a local
`proxies.txt` next to the scripts and pre-fills the box on the next launch.
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
still fails that account. tempmail.plus traffic stays on the machine's direct
connection. The log only ever shows each proxy's host:port, never its
password. `proxies.txt` and `proxy.txt` are listed in `.gitignore`, so
they are never committed — keep it that way.

**Số luồng** (1–5, default 1) runs that many registration lanes in
parallel. The proxy pool is split by stride — lane k uses the proxies at
positions k, k+lanes, k+2·lanes, … — so no two lanes ever share a proxy
at the same time, and each lane rotates (and burns) within its own
slice. Each lane keeps its **own invite chain**: lane 1 continues the
code in the Mã mời field, the other lanes start fresh chains from their
first account, and within a lane each account redeems the previous
account's code exactly as in single-lane mode. All lanes share the **one
saved card**, and the per-card success limit counts globally across
lanes (slots are reserved atomically, so the limit is never overshot).
Lanes need at least as many proxies as lanes; with fewer proxies the
lane count is reduced to the pool size. Per-lane state lives in
`proxy-<n>.txt` / `redeem-<n>.txt` next to the scripts (also gitignored);
the spreadsheet is written under a lock so lanes cannot corrupt it.

The **Proxy lỗi** (bad proxy) box under the Proxy box is the exclusion
list. When a proxy's exit IP lands on the "Join the waitlist" screen or
Meta refuses account creation on it ("Something went wrong"), the app
moves that proxy into this box, stores it in `bad-proxies.txt`
(gitignored), retires it immediately for the rest of the run — it is
never rotated back to, in any lane — and filters it out of the pool at
the start of every later run. Matching is by host:port, so a re-pasted
line with fresh credentials is still excluded. The box is editable:
delete a line there and that proxy is allowed again from the next run.

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
untrusted location. The app does not bypass CAPTCHA or other
site controls.
