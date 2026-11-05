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

Set the number of accounts (1–100) and click **Bắt đầu đăng ký**. Before the
run, the app asks you to confirm that Surfshark VPN is connected (once per app
session); the app does not verify VPN status or control network routing, so
make sure Surfshark is connected and that split tunneling does not exclude
Chromium. For each account the app then, with no further clicks:

1. Creates a random mail.tm address and a random Muse.ai password, and saves
   both passwords to Windows Credential Manager.
2. Opens a dedicated Chromium window, fills the email into the visible
   "Mobile number or email" field, and submits **Continue** to request a
   verification code.
3. Polls mail.tm for the code (up to three minutes), fills it, and submits
   **Confirm** on the code form.
4. On the final profile form, fills the first and last name as "Long" and
   "long" and the birthday 21 November 2001, then clicks **Confirm**.
5. Counts the account as registered only when the profile form disappears
   afterwards. Only then are its email and password appended to
   `accounts.xlsx` in the app's local Windows data folder; failed attempts
   are only written to the on-screen log, never to the sheet.

The batch continues with the next address until the target count is reached
or you click **Dừng**. Click **Mở sheet Excel** to open the sheet with the
successfully registered accounts.

Passwords are stored by Python `keyring` in the Windows Credential Manager.
The Excel sheet contains the Muse.ai passwords in readable form, so protect
the file and do not share or sync it to an untrusted location. The Muse.ai
password is not automatically entered because the current Muse.ai form only
asks for a phone number or email. The app does not bypass CAPTCHA or other
site controls.
