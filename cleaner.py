from __future__ import annotations

import shutil
import time
from datetime import datetime, timedelta
from pathlib import Path

# Время хранения готового отчета в минутах.
RETENTION_MINUTES = 10

# Периодичность проверки папки "выход" в минутах.
CHECK_EVERY_MINUTES = 1

OUTPUT_DIR = Path("выход")
INPUT_DIR = Path("вход")
SUBJECT_FILENAME = "субъект.txt"


def log(message: str) -> None:
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] {message}", flush=True)


def process_old_folders() -> None:
    if not OUTPUT_DIR.exists():
        log(f'Папка "{OUTPUT_DIR}" не найдена, пропуск проверки.')
        return

    now = datetime.now()
    retention_delta = timedelta(minutes=RETENTION_MINUTES)

    for person_output_dir in OUTPUT_DIR.iterdir():
        if not person_output_dir.is_dir():
            continue

        final_word_file = person_output_dir / f"{person_output_dir.name}.docx"
        if not final_word_file.exists() or not final_word_file.is_file():
            log(
                f'В {person_output_dir} нет итогового файла "{person_output_dir.name}.docx", пропуск.'
            )
            continue

        created_at = datetime.fromtimestamp(final_word_file.stat().st_ctime)
        age = now - created_at

        if age <= retention_delta:
            continue

        input_dir_to_delete = INPUT_DIR / person_output_dir.name

        log(
            f"Удаляю устаревшую папку {person_output_dir} "
            f"(файлу {final_word_file.name} больше {RETENTION_MINUTES} минут)."
        )
        shutil.rmtree(person_output_dir, ignore_errors=False)

        if (
            input_dir_to_delete.exists()
            and input_dir_to_delete.is_dir()
            and input_dir_to_delete.name != SUBJECT_FILENAME
        ):
            log(f"Удаляю связанную папку {input_dir_to_delete}.")
            shutil.rmtree(input_dir_to_delete, ignore_errors=False)


def main() -> None:
    log(
        "Уборщик запущен. "
        f"Проверка раз в {CHECK_EVERY_MINUTES} минут, хранение {RETENTION_MINUTES} минут."
    )
    while True:
        try:
            process_old_folders()
        except Exception as exc:
            log(f"Ошибка во время проверки: {exc}")
        time.sleep(CHECK_EVERY_MINUTES * 60)


if __name__ == "__main__":
    main()
