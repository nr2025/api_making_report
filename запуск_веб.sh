#!/bin/bash
# Запуск веб-интерфейса из venv проекта (Git Bash / Windows).

cd "$(dirname "$0")"

if [ ! -f "./.venv/Scripts/python.exe" ]; then
  echo "Не найден .venv. Создайте окружение: python -m venv .venv, затем установите пакеты: .venv\\Scripts\\python.exe -m pip install -r requirements.txt"
  exit 1
fi

if [ ! -f "./run_web.py" ]; then
  echo "Не найден run_web.py в папке проекта."
  exit 1
fi

./.venv/Scripts/python.exe ./run_web.py
