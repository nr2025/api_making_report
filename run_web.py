# Запуск Streamlit-приложения веб.py через интерпретатор текущего процесса.
# Имя файла — латиница, чтобы .bat на Windows не ломал кириллицу в командной строке.

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
WEB_APP = ROOT / "веб.py"


def main() -> int:
    if not WEB_APP.exists():
        print(f"Не найден файл приложения: {WEB_APP}")
        return 1
    cmd = [
        sys.executable,
        "-m",
        "streamlit",
        "run",
        str(WEB_APP),
        "--server.address",
        "0.0.0.0",
        "--server.port",
        "8501",
    ]
    return subprocess.call(cmd)


if __name__ == "__main__":
    raise SystemExit(main())
