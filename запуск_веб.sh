#!/bin/bash
# Запуск веб-интерфейса из venv проекта (Git Bash / Windows).

cd "$(dirname "$0")"

if [ ! -f "./.venv/Scripts/python.exe" ]; then
  echo "Не найден .venv. Создайте окружение: python -m venv .venv, затем установите пакеты: .venv\\Scripts\\python.exe -m pip install -r requirements.txt"
  exit 1
fi

./.venv/Scripts/python.exe -m streamlit run веб.py --server.address 0.0.0.0 --server.port 8501
