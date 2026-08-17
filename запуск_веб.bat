@echo off
REM ============================================================================
REM Установка на сервере:
REM   1. Скопируйте папку проекта на сервер.
REM   2. Откройте командную строку в папке проекта и создайте окружение:
REM        python -m venv .venv
REM   3. Установите пакеты:
REM        .venv\Scripts\python.exe -m pip install -r requirements.txt
REM   4. Планировщик задач Windows: создайте задачу "при запуске компьютера",
REM      действие — запуск этого файла запуск_веб.bat (рабочая папка = папка проекта).
REM ============================================================================

cd /d "%~dp0"

if not exist "%~dp0.venv\Scripts\python.exe" (
    echo Не найден .venv. Создайте окружение: python -m venv .venv, затем установите пакеты: .venv\Scripts\python.exe -m pip install -r requirements.txt
    pause
    exit /b 1
)

if not exist "%~dp0run_web.py" (
    echo Не найден run_web.py в папке проекта.
    pause
    exit /b 1
)

"%~dp0.venv\Scripts\python.exe" "%~dp0run_web.py"
pause
