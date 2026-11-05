# reg-muse-ai

## Muse.ai Account Helper

A small Windows desktop helper for a single personal account. It creates a
mail.tm inbox, opens Muse.ai in a dedicated browser, retrieves email
verification codes, and stores mailbox and Muse.ai passwords using the Windows
keyring backend. CAPTCHA and any additional registration steps remain manual.

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

Click **Tạo mail & đăng ký** to generate a random mail.tm address and a random
Muse.ai password in a single step. The app saves the mailbox password and
Muse.ai password to
Windows Credential Manager before showing the credentials. It then asks you to
confirm that Surfshark VPN is connected (once per app session), and immediately
opens a dedicated Chromium window and fills the email into the visible
"Mobile number or email" field, then submits **Continue** to request a
verification code. It checks mail.tm for a new message, fills the detected
code, and submits **Confirm** on the code form. Leave the
Chromium window open while it waits (up to three minutes); it waits up to
90 seconds for a form button to become ready on a slow connection. CAPTCHA and any
additional steps are left to you. On the final profile form, it fills the
first and last name as "Long" and "long", and the supplied birthday
21 November 2001, then clicks **Confirm** when the button is enabled. If an
error occurs, Chromium stays
open so you can inspect the form; close its window when finished. The app does not
verify VPN status or
control network routing, so make sure Surfshark is connected and that split
tunneling does not exclude Chromium. Click
**Lấy thư** to read the verification email; double-click a message to view it
and copy a detected verification code. **Sao chép email + mật khẩu** copies
the generated account credentials to the clipboard. Each generated Muse.ai
email and password is also appended to `accounts.xlsx` in the app's local
Windows data folder; click **Mở sheet Excel** to open it.

Passwords are stored by Python `keyring` in the Windows Credential Manager.
The Excel sheet contains the Muse.ai password in readable form, so protect
the file and do not share or sync it to an untrusted location. The app's
roaming data file contains email addresses only. The Muse.ai password is not
automatically entered because the current Muse.ai form only asks for a phone
number or email. The app does not bypass CAPTCHA or other site controls.
