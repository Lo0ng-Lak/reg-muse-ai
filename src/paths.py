"""Đường dẫn dùng chung cho cả app.

Code nằm trong src/, còn file dữ liệu chạy (proxy.txt, mailbox.txt,
redeem.txt, selfie.txt, ...) nằm ở GỐC repo — cùng chỗ với bản cũ
để kho proxy và các file của người dùng không phải dời đi.
"""

from pathlib import Path

# Gốc repo = thư mục cha của src/.
ROOT = Path(__file__).resolve().parent.parent
SRC_DIR = Path(__file__).resolve().parent
# Thư mục dữ liệu của app trên Windows (%LOCALAPPDATA%\MuseAIHelper).
APP_DATA = Path.home() / "AppData" / "Local" / "MuseAIHelper"


def data_file(name):
    """Đường dẫn tuyệt đối tới một file dữ liệu chạy ở gốc repo."""
    return ROOT / name
