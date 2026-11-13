"""Muse.ai Account Helper — điểm vào duy nhất.

Chạy:  py main.py

Code thật nằm trong src/ (coordinator, storage, proxy_pool,
browser_helper, capsolver, mail_providers) và src/ui/ (cửa sổ
chính + hai tab). File dữ liệu chạy (proxies.txt, proxy.txt,
mailbox.txt, ...) vẫn nằm ở gốc repo như trước.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))

from ui.main_window import main

if __name__ == "__main__":
    main()
