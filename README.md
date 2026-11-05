# reg-muse-ai

## Muse.ai Account Helper

A small Windows desktop helper that registers Muse.ai accounts in a batch.
For each account it creates a mail.tm inbox, opens Muse.ai in a dedicated
browser, retrieves the email verification code, fills the profile form, and —
only when the signup actually completes — saves the account to a local Excel
sheet. The window shows nothing but the live process log. CAPTCHA and any
additional registration steps remain manual; an account that stalls on one
times out, is logged as failed, and the batch moves on to the next address.

> mail.tm addresses are temporary email addresses, not Gmail. Muse.ai may
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
number of accounts (1–100), and click **Bắt đầu đăng ký**. For each account
the app then, with no further clicks:

1. Creates a random mail.tm address and a random Muse.ai password, and saves
   both passwords to Windows Credential Manager.
2. Opens a dedicated Chromium window **through the proxy**, first checks the
   exit IP it actually got, and only continues if that IP is in the US.
   It then fills the email into the visible "Mobile number or email" field
   and submits **Continue** to request a verification code.
3. Polls mail.tm for the code (up to three minutes), fills it, and submits
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
6. After the card step there is one last confirmation screen: the tool
   clicks its Confirm button, waits to land inside the Muse interface,
   and reads the session cookies. A successful row in `accounts.xlsx`
   therefore carries: creation time, email, the Muse.ai password, the
   mail.tm mailbox password, the session cookie (JSON, captured from the
   browser — it is passed from the helper to the app out of band and is
   never printed in the log), and the account's own invite code (below).
   Older sheets with only the first three columns gain the new headers
   automatically on the next successful row.
7. Invite codes form a chain. Once inside the Muse interface, each
   account first redeems the *previous* account's invite code — the app
   passes it to the helper in `redeem.txt`, and the helper follows the
   UI path from the screenshots: bottom-left menu → **Cài đặt** →
   **Quy đổi mã lời mời** → type the 6 characters → **Xác nhận**. A
   failed redemption is logged but does not fail the account (the code
   can still be redeemed manually within 48 hours). The account then
   opens the **Mời** panel and reads its *own* invite code, which is
   saved in its sheet row and becomes the code the next account
   redeems. The **Mã mời** field in the window is the chain's current
   head: whatever is there when a run starts (typically your own
   account's code) is redeemed by the run's first account, and the
   field is overwritten with each new account's code as the run
   proceeds; it persists between runs in a local `invite.txt`.

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
still fails that account. mail.tm traffic stays on the machine's direct
connection. The log only ever shows each proxy's host:port, never its
password. `proxies.txt` and `proxy.txt` are listed in `.gitignore`, so
they are never committed — keep it that way.

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
The Excel sheet contains the Muse.ai passwords in readable form, so protect
the file and do not share or sync it to an untrusted location. The Muse.ai
password is not automatically entered because the current Muse.ai form only
asks for a phone number or email. The app does not bypass CAPTCHA or other
site controls.
