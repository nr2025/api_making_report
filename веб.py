# Запуск: streamlit run веб.py --server.address 0.0.0.0 --server.port 8501

from __future__ import annotations

import secrets
import threading
import time
import uuid
from datetime import datetime, timedelta
from pathlib import Path

import streamlit as st

from run_report_mvp import SETTINGS_PATH, load_settings, process_subject_folder, проверить_окружение

st.set_page_config(page_title="Обработка документов", layout="wide")

ALLOWED_EXTENSIONS = {".pdf", ".doc", ".docx", ".jpg", ".jpeg", ".png"}
INVALID_FOLDER_CHARS = '<>:"/\\|?*'
# Без похожих символов: 0/O, 1/I/L.
CODE_ALPHABET = "23456789ABCDEFGHJKMNPQRSTUVWXYZ"
CODE_LENGTH = 4
WORD_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


def generate_task_code() -> str:
    return "".join(secrets.choice(CODE_ALPHABET) for _ in range(CODE_LENGTH))


def normalize_lookup_code(raw: str) -> str:
    return (raw or "").strip().upper()


def folder_code_from_name(folder_name: str) -> str | None:
    """Код — третий блок имени папки (разделитель __)."""
    parts = folder_name.split("__")
    if len(parts) < 3:
        return None
    return parts[2].upper()


def folder_fio_from_name(folder_name: str) -> str:
    return folder_name.split("__", 1)[0]


class JobQueue:
    """Общая очередь задач: один фоновый обработчик на всё приложение."""

    def __init__(self) -> None:
        self.tasks: list[dict] = []
        self.lock = threading.Lock()
        self._worker = threading.Thread(target=self._worker_loop, daemon=True)
        self._worker.start()

    def _existing_codes(self) -> set[str]:
        return {str(task.get("код", "")).upper() for task in self.tasks if task.get("код")}

    def allocate_code(self) -> str:
        with self.lock:
            existing = self._existing_codes()
            for _ in range(100):
                code = generate_task_code()
                if code not in existing:
                    return code
        return generate_task_code()

    def add_task(self, fio: str, folder: Path, code: str) -> str:
        task_id = str(uuid.uuid4())
        task = {
            "id": task_id,
            "ФИО": fio,
            "папка": str(folder),
            "код": code,
            "статус": "в очереди",
            "время постановки": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "путь к результату": None,
            "ошибка": None,
            "шаг": None,
            "всего_шагов": None,
            "нечитаемые": [],
        }
        with self.lock:
            self.tasks.append(task)
        return task_id

    def get_task(self, task_id: str | None) -> dict | None:
        if not task_id:
            return None
        with self.lock:
            for task in self.tasks:
                if task["id"] == task_id:
                    return dict(task)
        return None

    def find_task_by_code(self, code: str) -> dict | None:
        normalized = normalize_lookup_code(code)
        if not normalized:
            return None
        with self.lock:
            for task in self.tasks:
                if str(task.get("код", "")).upper() == normalized:
                    return dict(task)
        return None

    def snapshot_tasks(self) -> list[dict]:
        with self.lock:
            return [dict(task) for task in self.tasks]

    def queue_position(self, task_id: str) -> int | None:
        with self.lock:
            waiting = [task for task in self.tasks if task["статус"] == "в очереди"]
            for index, task in enumerate(waiting, start=1):
                if task["id"] == task_id:
                    return index
        return None

    def _worker_loop(self) -> None:
        settings = load_settings(SETTINGS_PATH)
        while True:
            task: dict | None = None
            with self.lock:
                for candidate in self.tasks:
                    if candidate["статус"] == "в очереди":
                        candidate["статус"] = "обрабатывается"
                        candidate["ошибка"] = None
                        candidate["шаг"] = None
                        candidate["всего_шагов"] = None
                        task = candidate
                        break
            if task is None:
                time.sleep(0.5)
                continue

            def progress_callback(step: int, total: int) -> None:
                with self.lock:
                    task["шаг"] = step
                    task["всего_шагов"] = total

            try:
                _, report_path, unreadable = process_subject_folder(
                    settings,
                    Path(task["папка"]),
                    task["ФИО"],
                    progress_callback=progress_callback,
                )
                with self.lock:
                    task["статус"] = "готово"
                    task["путь к результату"] = str(report_path)
                    task["ошибка"] = None
                    task["нечитаемые"] = list(unreadable)
            except Exception as exc:  # noqa: BLE001 — ошибка пишется в задачу
                with self.lock:
                    task["статус"] = "ошибка"
                    task["ошибка"] = str(exc)
                    task["путь к результату"] = None
                    task["нечитаемые"] = []


@st.cache_resource
def get_job_queue() -> JobQueue:
    return JobQueue()


def sanitize_folder_name(name: str) -> str:
    cleaned = "".join("_" if ch in INVALID_FOLDER_CHARS else ch for ch in name.strip())
    cleaned = cleaned.strip(" .")
    return cleaned or "субъект"


def save_uploaded_files(folder: Path, uploaded_files) -> None:
    folder.mkdir(parents=True, exist_ok=True)
    for uploaded in uploaded_files or []:
        target = folder / Path(uploaded.name).name
        target.write_bytes(uploaded.getbuffer())


def save_pasted_text(folder: Path, text: str) -> Path:
    folder.mkdir(parents=True, exist_ok=True)
    target = folder / "вставленный_текст.txt"
    target.write_text(text, encoding="utf-8")
    return target


def format_materials_summary(file_count: int, has_pasted_text: bool) -> str:
    if file_count and has_pasted_text:
        return f"файлов: {file_count} + вставленный текст"
    if has_pasted_text:
        return "вставленный текст"
    return f"файлов: {file_count}"


def format_task_status_line(queue: JobQueue, task: dict) -> str:
    fio = task["ФИО"]
    code = task.get("код", "")
    status = task["статус"]
    if status == "в очереди":
        position = queue.queue_position(task["id"])
        if position is None:
            middle = "в очереди"
        else:
            middle = f"в очереди, позиция {position}"
    elif status == "обрабатывается":
        step = task.get("шаг")
        total = task.get("всего_шагов")
        if step is not None and total is not None:
            middle = f"обрабатывается (шаг {step} из {total})"
        else:
            middle = "обрабатывается"
    elif status == "готово":
        middle = "готово"
    elif status == "ошибка":
        middle = f"ошибка: {task.get('ошибка') or 'неизвестная ошибка'}"
    else:
        middle = status
    line = f"{fio} — {middle} — код: {code}"
    unreadable = task.get("нечитаемые") or []
    if unreadable:
        line = f"{line} — ⚠ не прочитано файлов: {len(unreadable)}"
    return line


def find_output_docx_by_code(output_dir: Path, code: str) -> tuple[str, Path] | None:
    normalized = normalize_lookup_code(code)
    if not normalized or not output_dir.exists():
        return None
    for folder in sorted(output_dir.iterdir()):
        if not folder.is_dir():
            continue
        folder_code = folder_code_from_name(folder.name)
        if folder_code != normalized:
            continue
        docx_path = folder / f"{folder.name}.docx"
        if docx_path.exists() and docx_path.is_file():
            return folder_fio_from_name(folder.name), docx_path
    return None


def render_own_tasks(queue: JobQueue, task_ids: list[str]) -> None:
    st.subheader("Ваши задачи")
    if not task_ids:
        st.info("Задачи ещё не поставлены.")
        st.caption(
            "Сохраните код задачи — по нему можно забрать результат позже, "
            "в том числе после закрытия страницы."
        )
        return

    for task_id in task_ids:
        task = queue.get_task(task_id)
        if task is None:
            st.warning(f"Задача {task_id} не найдена в общей очереди.")
            continue
        st.write(format_task_status_line(queue, task))
        if task["статус"] == "готово":
            result_path = task.get("путь к результату")
            if result_path:
                path = Path(result_path)
                if path.exists():
                    st.download_button(
                        label="Скачать Word-файл",
                        data=path.read_bytes(),
                        file_name=path.name,
                        mime=WORD_MIME,
                        key=f"download_{task_id}",
                    )
                else:
                    st.error(f"Файл результата не найден: {path}")

    st.caption(
        "Сохраните код задачи — по нему можно забрать результат позже, "
        "в том числе после закрытия страницы."
    )




def render_lookup_by_code(queue: JobQueue) -> None:
    st.subheader("Получить результат по коду")
    code_input = st.text_input("Код задачи", key="lookup_code_input")
    if not st.button("Найти", key="lookup_code_button"):
        return

    code = normalize_lookup_code(code_input)
    if not code:
        st.error("Введите код задачи.")
        return

    task = queue.find_task_by_code(code)
    if task is not None and task["статус"] in {"в очереди", "обрабатывается"}:
        st.info(f"Субъект: {task['ФИО']}")
        status = task["статус"]
        if status == "в очереди":
            position = queue.queue_position(task["id"])
            if position is None:
                st.write("Статус: в очереди")
            else:
                st.write(f"Статус: в очереди, позиция {position}")
        else:
            step = task.get("шаг")
            total = task.get("всего_шагов")
            if step is not None and total is not None:
                st.write(f"Статус: обрабатывается, шаг {step} из {total}")
            else:
                st.write("Статус: обрабатывается")
        return

    settings = load_settings(SETTINGS_PATH)
    found = find_output_docx_by_code(Path(settings.output_dir), code)
    if found is None and task is not None and task["статус"] == "готово":
        result_path = task.get("путь к результату")
        if result_path:
            path = Path(result_path)
            if path.exists():
                found = (task["ФИО"], path)

    if found is not None:
        fio, docx_path = found
        st.success(f"Субъект: {fio}")
        st.download_button(
            label="Скачать Word-файл",
            data=docx_path.read_bytes(),
            file_name=docx_path.name,
            mime=WORD_MIME,
            key=f"dl_code_{code}",
        )
        return

    if task is not None and task["статус"] == "ошибка":
        st.error(f"Субъект: {task['ФИО']}. Ошибка: {task.get('ошибка') or 'неизвестная ошибка'}")
        return

    st.error("Задача не найдена: код неверен или результат удалён по сроку хранения")


@st.fragment(run_every=timedelta(seconds=5))
def status_and_queue_block(queue: JobQueue) -> None:
    render_own_tasks(queue, st.session_state.get("task_ids", []))


@st.cache_resource
def check_environment_once() -> str | None:
    """Один раз при старте приложения. None — ок, иначе текст ошибки."""
    try:
        проверить_окружение()
    except RuntimeError as exc:
        return str(exc)
    return None


def main() -> None:
    st.markdown(
        """
        <style>
        .block-container {
            padding-top: 2rem;
            padding-bottom: 1rem;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
    st.header("Обработка документов")

    env_error = check_environment_once()
    if env_error:
        st.error(env_error)
        st.stop()

    queue = get_job_queue()
    if "task_ids" not in st.session_state:
        st.session_state.task_ids = []
    # Совместимость со старым ключом session_state.
    if st.session_state.get("task_id") and st.session_state.task_id not in st.session_state.task_ids:
        st.session_state.task_ids.append(st.session_state.task_id)

    left_col, right_col = st.columns([1, 1])

    with left_col:
        fio = st.text_input("ФИО субъекта")
        uploaded_files = st.file_uploader(
            "Файлы документов",
            type=["pdf", "doc", "docx", "jpg", "jpeg", "png"],
            accept_multiple_files=True,
        )
        pasted_text = st.text_area(
            "Текст (вставьте, если данные не в файле)",
            height=200,
            placeholder="Можно оставить пустым.",
        )

        if st.button("Отправить в обработку", type="primary"):
            fio_clean = (fio or "").strip()
            files = list(uploaded_files or [])
            text_clean = (pasted_text or "").strip()
            if not fio_clean:
                st.error("Укажите ФИО субъекта.")
            elif not files and not text_clean:
                st.error("Загрузите файлы или вставьте текст")
            else:
                unsupported = [
                    file.name
                    for file in files
                    if Path(file.name).suffix.lower() not in ALLOWED_EXTENSIONS
                ]
                if unsupported:
                    st.error(f"Неподдерживаемые файлы: {', '.join(unsupported)}")
                else:
                    settings = load_settings(SETTINGS_PATH)
                    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    code = queue.allocate_code()
                    # Одинаковое имя для вход/ и выход/ (уборщик сопоставляет по имени папки).
                    folder_name = f"{sanitize_folder_name(fio_clean)}__{stamp}__{code}"
                    subject_dir = Path(settings.input_dir) / folder_name
                    save_uploaded_files(subject_dir, files)
                    if text_clean:
                        save_pasted_text(subject_dir, text_clean)
                    task_id = queue.add_task(fio_clean, subject_dir, code)
                    st.session_state.task_ids.append(task_id)
                    st.session_state.task_id = task_id
                    materials = format_materials_summary(len(files), bool(text_clean))
                    st.success(
                        f"Задача поставлена в очередь. Код: {code}. "
                        f"{materials}. Папка: {subject_dir.name}"
                    )

    with right_col:
        status_and_queue_block(queue)
        st.divider()
        render_lookup_by_code(queue)


if __name__ == "__main__":
    main()
