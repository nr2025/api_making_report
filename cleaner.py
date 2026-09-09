from __future__ import annotations

import shutil
import time
from datetime import date, datetime, timedelta
from pathlib import Path

# Час удаления по правилу «конец следующего рабочего дня».
DELETE_HOUR = 20

# Папка без итогового .docx старше этого срока считается зависшей.
STUCK_HOURS = 72

# Периодичность проверки в минутах.
CHECK_EVERY_MINUTES = 30

# Не удалять, если промежуточные/ менялись за этот интервал (задача ещё идёт).
RECENT_ACTIVITY_MINUTES = 30

OUTPUT_DIR = Path("выход")
INPUT_DIR = Path("вход")
LOG_FILE = Path("уборщик.log")
INTERMEDIATE_DIRNAME = "промежуточные"


def log(message: str) -> None:
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{timestamp}] {message}"
    print(line, flush=True)
    with LOG_FILE.open("a", encoding="utf-8") as fh:
        fh.write(line + "\n")


def folder_code(folder_name: str) -> str:
    """Из имени папки FIO__TIMESTAMP__CODE вернуть только CODE."""
    parts = folder_name.split("__")
    if len(parts) >= 3 and parts[-1]:
        return parts[-1]
    return folder_name


def next_working_day(from_date: date) -> date:
    candidate = from_date + timedelta(days=1)
    while candidate.weekday() >= 5:  # 5=суббота, 6=воскресенье
        candidate += timedelta(days=1)
    return candidate


def retention_deadline(docx_mtime: datetime) -> datetime:
    """Момент удаления: следующий рабочий день после даты .docx, в DELETE_HOUR:00."""
    target = next_working_day(docx_mtime.date())
    return datetime(target.year, target.month, target.day, DELETE_HOUR, 0, 0)


def is_recently_active(person_output_dir: Path, now: datetime) -> bool:
    intermediate = person_output_dir / INTERMEDIATE_DIRNAME
    if not intermediate.exists():
        return False
    modified_at = datetime.fromtimestamp(intermediate.stat().st_mtime)
    return now - modified_at < timedelta(minutes=RECENT_ACTIVITY_MINUTES)


def delete_task_folders(person_output_dir: Path, reason: str) -> None:
    input_dir_to_delete = INPUT_DIR / person_output_dir.name
    code = folder_code(person_output_dir.name)
    log(f"Удаляю папку {code} ({reason}).")
    shutil.rmtree(person_output_dir, ignore_errors=False)

    if input_dir_to_delete.exists() and input_dir_to_delete.is_dir():
        log(f"Удаляю связанную папку вход/{code} ({reason}).")
        shutil.rmtree(input_dir_to_delete, ignore_errors=False)


def process_old_folders() -> None:
    if not OUTPUT_DIR.exists():
        log(f'Папка "{OUTPUT_DIR}" не найдена, пропуск проверки.')
        return

    now = datetime.now()
    stuck_delta = timedelta(hours=STUCK_HOURS)

    for person_output_dir in OUTPUT_DIR.iterdir():
        if not person_output_dir.is_dir():
            continue

        if is_recently_active(person_output_dir, now):
            log(
                f"Пропуск {folder_code(person_output_dir.name)}: "
                f"{INTERMEDIATE_DIRNAME}/ изменялась менее {RECENT_ACTIVITY_MINUTES} мин. назад."
            )
            continue

        final_word_file = person_output_dir / f"{person_output_dir.name}.docx"
        if final_word_file.exists() and final_word_file.is_file():
            docx_mtime = datetime.fromtimestamp(final_word_file.stat().st_mtime)
            deadline = retention_deadline(docx_mtime)
            if now < deadline:
                continue
            delete_task_folders(person_output_dir, "срок хранения")
            continue

        folder_mtime = datetime.fromtimestamp(person_output_dir.stat().st_mtime)
        if now - folder_mtime <= stuck_delta:
            continue

        delete_task_folders(person_output_dir, "зависшая задача")


def main() -> None:
    log(
        "Уборщик запущен. "
        f"Проверка раз в {CHECK_EVERY_MINUTES} мин., "
        f"удаление в {DELETE_HOUR}:00 следующего рабочего дня, "
        f"зависшие задачи — {STUCK_HOURS} ч."
    )
    while True:
        try:
            process_old_folders()
        except Exception as exc:
            log(f"Ошибка во время проверки: {exc}")
        time.sleep(CHECK_EVERY_MINUTES * 60)


if __name__ == "__main__":
    main()
