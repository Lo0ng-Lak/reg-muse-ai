"""Muse.ai Account Helper — điểm vào duy nhất.

Chạy:  py main.py
      py main.py --cookie-web   (chỉ mở web quản lý kho cookie,
                                 không mở cửa sổ tkinter)

Code thật nằm trong src/ (coordinator, storage, proxy_pool,
browser_helper, capsolver, mail_providers) và src/ui/ (cửa sổ
chính + hai tab). File dữ liệu chạy (proxies.txt, proxy.txt,
mailbox.txt, ...) vẫn nằm ở gốc repo như trước.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))

if __name__ == "__main__":
    if "--cookie-web" in sys.argv:
        # Bỏ cờ này trước khi vào cookie_web.main (argparse của
        # nó không biết cờ này).
        sys.argv.remove("--cookie-web")
        from cookie_web import main as cookie_web_main

        cookie_web_main()
    else:
        from ui.main_window import main

        main()
