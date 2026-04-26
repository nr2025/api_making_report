from __future__ import annotations

import base64
import json
import re
import urllib.error
import urllib.request
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterable

try:
    import fitz
except ImportError:
    fitz = None

try:
    import docx
except ImportError:
    docx = None


ROOT_DIR = Path(__file__).resolve().parent
SETTINGS_PATH = ROOT_DIR / "настройки.yaml"
INPUT_DIR_DEFAULT = ROOT_DIR / "вход"
PROMPTS_DIR_DEFAULT = ROOT_DIR / "промпты"
OUTPUT_DIR_DEFAULT = ROOT_DIR / "выход"
TEMPLATE_PATH_DEFAULT = ROOT_DIR / "шаблон.docx"
SUPPORTED_EXTENSIONS = {".docx", ".doc", ".pdf", ".jpg", ".jpeg", ".png"}
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png"}
PDF_TEXT_MIN_CHARS_PER_PAGE = 50


@dataclass
class Settings:
    api_url: str
    model: str
    input_dir: Path
    prompts_dir: Path
    output_dir: Path
    template_path: Path
    num_predict: int
    request_timeout_sec: int
    max_json_retries: int


@dataclass
class PreparedInput:
    text_documents: list[dict[str, str]]
    images_base64: list[str]
    all_documents: list[str]


@dataclass
class PromptRunResult:
    prompt_file: Path
    output_json: Path
    parsed_json: dict | list | None
    raw_response_text: str


def parse_simple_yaml(path: Path) -> dict[str, str]:
    settings: dict[str, str] = {}
    raw_content = path.read_text(encoding="utf-8-sig")
    for raw_line in raw_content.splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or ":" not in line:
            continue
        key, value = line.split(":", 1)
        settings[key.strip()] = value.strip().strip("'\"")
    return settings


def load_settings(path: Path) -> Settings:
    if not path.exists():
        raise FileNotFoundError(f"Не найден файл настроек: {path}")

    raw = parse_simple_yaml(path)
    api_url = raw.get("api_url")
    model = raw.get("model")
    if not api_url or not model:
        raise ValueError("В настройках должны быть указаны api_url и model")

    input_dir = Path(raw.get("input_dir", str(INPUT_DIR_DEFAULT)))
    prompts_dir = Path(raw.get("prompts_dir", str(PROMPTS_DIR_DEFAULT)))
    output_dir = Path(raw.get("output_dir", str(OUTPUT_DIR_DEFAULT)))
    template_path = Path(raw.get("template_path", str(TEMPLATE_PATH_DEFAULT)))
    num_predict = int(raw.get("num_predict", "16384"))
    request_timeout_sec = int(raw.get("request_timeout_sec", "300"))
    max_json_retries = int(raw.get("max_json_retries", "3"))

    return Settings(
        api_url=api_url.rstrip("/"),
        model=model,
        input_dir=input_dir,
        prompts_dir=prompts_dir,
        output_dir=output_dir,
        template_path=template_path,
        num_predict=max(num_predict, 16384),
        request_timeout_sec=request_timeout_sec,
        max_json_retries=max(1, max_json_retries),
    )


def read_docx_text(path: Path) -> str:
    with zipfile.ZipFile(path) as archive:
        xml_bytes = archive.read("word/document.xml")
    root = ET.fromstring(xml_bytes)
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    chunks: list[str] = []
    for para in root.findall(".//w:p", ns):
        texts = [node.text for node in para.findall(".//w:t", ns) if node.text]
        paragraph = "".join(texts).strip()
        if paragraph:
            chunks.append(paragraph)
    return "\n".join(chunks)


def read_text_file(path: Path) -> str:
    encodings = ("utf-8", "utf-8-sig", "cp1251")
    for enc in encodings:
        try:
            return path.read_text(encoding=enc)
        except UnicodeDecodeError:
            continue
    return path.read_text(encoding="utf-8", errors="replace")


def is_supported_document_file(file_path: Path) -> bool:
    return file_path.suffix.lower() in SUPPORTED_EXTENSIONS


def read_image_as_base64(path: Path) -> str:
    return base64.b64encode(path.read_bytes()).decode("ascii")


def convert_doc_to_docx(path: Path) -> Path:
    try:
        import win32com.client
    except ImportError as exc:
        raise RuntimeError("Для обработки .doc требуется pywin32 (win32com).") from exc

    output_path = path.with_suffix(".converted.docx")
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc_handle = None
    try:
        doc_handle = word.Documents.Open(str(path))
        doc_handle.SaveAs(str(output_path), FileFormat=16)
    finally:
        if doc_handle is not None:
            doc_handle.Close(False)
        word.Quit()
    return output_path


def read_pdf(path: Path) -> tuple[str, list[str]]:
    if fitz is None:
        raise RuntimeError("Для обработки PDF требуется PyMuPDF (fitz).")

    text_parts: list[str] = []
    pages_text_len: list[int] = []
    image_payloads: list[str] = []
    with fitz.open(path) as pdf:
        for page in pdf:
            page_text = page.get_text("text").strip()
            text_parts.append(page_text)
            pages_text_len.append(len(page_text))

        avg_page_chars = (sum(pages_text_len) / len(pages_text_len)) if pages_text_len else 0
        is_text_pdf = avg_page_chars >= PDF_TEXT_MIN_CHARS_PER_PAGE
        if is_text_pdf:
            full_text = "\n".join(part for part in text_parts if part).strip()
            return full_text, []

        for page in pdf:
            pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0))
            image_payloads.append(base64.b64encode(pix.tobytes("png")).decode("ascii"))

    return "", image_payloads


def list_subject_folders(input_dir: Path) -> list[Path]:
    if not input_dir.exists():
        raise FileNotFoundError(f"Папка входа не найдена: {input_dir}")
    subdirs = sorted(p for p in input_dir.iterdir() if p.is_dir())
    if subdirs:
        return subdirs
    return [input_dir]


def read_subject_name(subject_dir: Path) -> str:
    subject_path = subject_dir / "субъект.txt"
    if subject_path.exists():
        value = read_text_file(subject_path).strip()
        if value:
            return value
    return subject_dir.name


def collect_documents(subject_dir: Path) -> PreparedInput:
    text_docs: list[dict[str, str]] = []
    images_base64: list[str] = []
    all_docs: list[str] = []

    for file_path in sorted(p for p in subject_dir.rglob("*") if p.is_file()):
        if file_path.name.lower() == "субъект.txt":
            continue
        if not is_supported_document_file(file_path):
            continue

        rel_path = str(file_path.relative_to(subject_dir))
        all_docs.append(rel_path)
        ext = file_path.suffix.lower()
        try:
            if ext == ".docx":
                content = read_docx_text(file_path).strip()
                if content:
                    text_docs.append({"path": rel_path, "content": content})
            elif ext == ".doc":
                converted = convert_doc_to_docx(file_path)
                content = read_docx_text(converted).strip()
                if content:
                    text_docs.append({"path": rel_path, "content": content})
            elif ext == ".pdf":
                pdf_text, pdf_images = read_pdf(file_path)
                if pdf_text:
                    text_docs.append({"path": rel_path, "content": pdf_text})
                images_base64.extend(pdf_images)
            elif ext in IMAGE_EXTENSIONS:
                images_base64.append(read_image_as_base64(file_path))
        except Exception as exc:  # noqa: BLE001
            text_docs.append({"path": rel_path, "content": f"[Ошибка чтения файла: {exc}]"})
    return PreparedInput(text_documents=text_docs, images_base64=images_base64, all_documents=all_docs)


def prompt_sort_key(path: Path) -> tuple[int, str]:
    match = re.match(r"^(\d+)", path.stem)
    if match:
        return (int(match.group(1)), path.name)
    return (10**9, path.name)


def list_prompt_files(prompts_dir: Path) -> list[Path]:
    if not prompts_dir.exists():
        raise FileNotFoundError(f"Папка промптов не найдена: {prompts_dir}")
    files = sorted((p for p in prompts_dir.rglob("*.txt") if p.is_file()), key=prompt_sort_key)
    if not files:
        raise FileNotFoundError(f"В папке промптов нет .txt файлов: {prompts_dir}")
    return files


def subject_prefix(subject_name: str) -> str:
    return (
        f"Субъект исследования: {subject_name}.\n"
        "Если в задании не указано иное, данные извлекаются в отношении Субъекта."
    )


def build_prompt_with_documents(
    subject_name: str,
    base_prompt: str,
    documents: Iterable[dict[str, str]],
    image_count: int,
) -> str:
    parts = [subject_prefix(subject_name), "", base_prompt.strip(), "", "Документы для анализа:"]
    for idx, doc in enumerate(documents, start=1):
        parts.append("")
        parts.append(f"=== Документ {idx}: {doc['path']} ===")
        parts.append(doc["content"])
    if image_count:
        parts.append("")
        parts.append(f"Приложено изображений для анализа: {image_count}")
    return "\n".join(parts).strip()


def build_prompt_with_previous_steps(
    subject_name: str,
    base_prompt: str,
    previous_steps: list[PromptRunResult],
) -> str:
    parts = [subject_prefix(subject_name), "", base_prompt.strip(), "", "Результаты предыдущих шагов:"]
    for step in previous_steps:
        parts.append("")
        parts.append(f"=== {step.prompt_file.stem} ===")
        parts.append(step.output_json.read_text(encoding="utf-8"))
    return "\n".join(parts).strip()


def normalize_api_generate_url(api_url: str) -> str:
    if re.search(r"/api/generate/?$", api_url):
        return api_url
    return f"{api_url}/api/generate"


def call_ollama(
    api_url: str,
    model: str,
    prompt: str,
    num_predict: int,
    timeout_sec: int,
    images_base64: list[str] | None = None,
) -> dict:
    url = normalize_api_generate_url(api_url)
    payload = {
        "model": model,
        "prompt": prompt,
        "stream": False,
        "options": {"num_predict": num_predict},
    }
    if images_base64:
        payload["images"] = images_base64

    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    request = urllib.request.Request(
        url=url,
        data=data,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout_sec) as response:
            body = response.read().decode("utf-8")
    except urllib.error.HTTPError as exc:
        details = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Ollama HTTP {exc.code}: {details}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Не удалось подключиться к Ollama: {exc}") from exc

    return json.loads(body)


def extract_json_from_text(text: str) -> dict | list | None:
    text = text.strip()
    if not text:
        return None
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    first_obj = text.find("{")
    first_arr = text.find("[")
    starts = [i for i in (first_obj, first_arr) if i != -1]
    if not starts:
        return None
    start = min(starts)
    last_obj = text.rfind("}")
    last_arr = text.rfind("]")
    end = max(last_obj, last_arr)
    if end < start:
        return None

    try:
        return json.loads(text[start : end + 1])
    except json.JSONDecodeError:
        return None


def call_until_valid_json(
    settings: Settings,
    prompt: str,
    images_base64: list[str] | None = None,
) -> tuple[dict, dict | list | None]:
    last_response: dict = {}
    parsed_json: dict | list | None = None
    for _ in range(settings.max_json_retries):
        last_response = call_ollama(
            api_url=settings.api_url,
            model=settings.model,
            prompt=prompt,
            num_predict=settings.num_predict,
            timeout_sec=settings.request_timeout_sec,
            images_base64=images_base64,
        )
        parsed_json = extract_json_from_text(last_response.get("response", ""))
        if parsed_json is not None:
            break
    return last_response, parsed_json


def ensure_dict_payload(value: dict | list | None) -> dict:
    if isinstance(value, dict):
        return value
    if isinstance(value, list):
        return {"данные": value}
    return {"данные": []}


def save_prompt_result_json(
    intermediate_dir: Path,
    prompt_file: Path,
    subject_name: str,
    docs: PreparedInput,
    ollama_response: dict,
    parsed_json: dict | list | None,
) -> PromptRunResult:
    intermediate_dir.mkdir(parents=True, exist_ok=True)
    output_json = intermediate_dir / f"{prompt_file.stem}.json"
    payload = {
        "meta": {
            "created_at_utc": datetime.now(timezone.utc).isoformat(),
            "subject_name": subject_name,
            "prompt_file": str(prompt_file),
            "documents_count": len(docs.all_documents),
            "text_documents_count": len(docs.text_documents),
            "images_count": len(docs.images_base64),
            "documents": docs.all_documents,
        },
        "данные": ensure_dict_payload(parsed_json).get("данные", []),
        "parsed_json_response": parsed_json,
        "raw_response_text": ollama_response.get("response", ""),
    }
    output_json.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return PromptRunResult(
        prompt_file=prompt_file,
        output_json=output_json,
        parsed_json=parsed_json,
        raw_response_text=ollama_response.get("response", ""),
    )


def set_paragraph_text_with_breaks(paragraph, text: str) -> None:
    paragraph.clear()
    parts = str(text).split("\n")
    run = paragraph.add_run(parts[0] if parts else "")
    for chunk in parts[1:]:
        run.add_break()
        run = paragraph.add_run(chunk)


def set_cell_text_with_breaks(cell, text: str) -> None:
    if not cell.paragraphs:
        paragraph = cell.add_paragraph("")
    else:
        paragraph = cell.paragraphs[0]
    set_paragraph_text_with_breaks(paragraph, text)
    for extra in cell.paragraphs[1:]:
        extra.clear()


def replace_placeholder_with_lines(paragraph, placeholder: str, lines: list[str]) -> bool:
    if placeholder not in paragraph.text:
        return False
    safe_lines = lines or ["данные не обнаружены"]
    prefix, suffix = paragraph.text.split(placeholder, 1)
    first_line = f"{prefix}{safe_lines[0]}{suffix}"
    set_paragraph_text_with_breaks(paragraph, first_line)

    anchor = paragraph
    for extra_line in safe_lines[1:]:
        new_paragraph = paragraph.insert_paragraph_before("")
        anchor._p.addnext(new_paragraph._p)  # noqa: SLF001
        set_paragraph_text_with_breaks(new_paragraph, extra_line)
        anchor = new_paragraph
    return True


def stringify_cell_value(value) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return value
    if isinstance(value, (int, float, bool)):
        return str(value)
    return json.dumps(value, ensure_ascii=False)


def find_placeholder_row(table, placeholder: str):
    for row in table.rows:
        for cell in row.cells:
            if placeholder in cell.text:
                return row
    return None


def fill_table_for_objects(table, placeholder: str, values: list[dict]) -> bool:
    row = find_placeholder_row(table, placeholder)
    if row is None:
        return False

    table._tbl.remove(row._tr)  # noqa: SLF001
    rows_to_add = values if values else [{"данные": "данные не обнаружены"}]
    for item in rows_to_add:
        new_row = table.add_row()
        for idx, (_, value) in enumerate(item.items()):
            if idx >= len(new_row.cells):
                break
            set_cell_text_with_breaks(new_row.cells[idx], stringify_cell_value(value))
    return True


def replace_placeholder_in_paragraphs(document, placeholder: str, lines: list[str]) -> bool:
    changed = False
    for paragraph in document.paragraphs:
        if replace_placeholder_with_lines(paragraph, placeholder, lines):
            changed = True
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if replace_placeholder_with_lines(paragraph, placeholder, lines):
                        changed = True
    return changed


def build_word_report(template_path: Path, output_docx: Path, prompt_results: list[PromptRunResult]) -> Path:
    if docx is None:
        raise RuntimeError("Для сборки Word нужен пакет python-docx.")
    if not template_path.exists():
        raise FileNotFoundError(f"Не найден шаблон Word: {template_path}")

    document = docx.Document(str(template_path))
    payload_by_tag: dict[str, dict] = {}
    for step in prompt_results:
        try:
            payload_by_tag[step.prompt_file.stem] = json.loads(step.output_json.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            payload_by_tag[step.prompt_file.stem] = {"данные": []}

    for tag, payload in payload_by_tag.items():
        placeholder = f"{{{{{tag}}}}}"
        raw_values = payload.get("данные", [])
        values = raw_values if isinstance(raw_values, list) else []
        is_object_values = values and all(isinstance(item, dict) for item in values)

        if is_object_values:
            table_handled = False
            for table in document.tables:
                if fill_table_for_objects(table, placeholder, values):
                    table_handled = True
                    break
            if not table_handled:
                fallback_lines = [json.dumps(values, ensure_ascii=False, indent=2)]
                replace_placeholder_in_paragraphs(document, placeholder, fallback_lines)
            continue

        string_lines = [str(item) for item in values if isinstance(item, str)]
        if not string_lines:
            string_lines = ["данные не обнаружены"]
        replace_placeholder_in_paragraphs(document, placeholder, string_lines)

    output_docx.parent.mkdir(parents=True, exist_ok=True)
    document.save(str(output_docx))
    return output_docx


def process_subject_folder(settings: Settings, subject_dir: Path) -> tuple[Path, Path]:
    subject_name = read_subject_name(subject_dir)
    prepared = collect_documents(subject_dir)
    if not prepared.text_documents and not prepared.images_base64:
        raise RuntimeError(f"Нет читаемых документов в папке: {subject_dir}")

    prompt_files = list_prompt_files(settings.prompts_dir)
    output_subject_dir = settings.output_dir / subject_dir.name
    intermediate_dir = output_subject_dir / "промежуточные"
    prompt_runs: list[PromptRunResult] = []

    for index, prompt_file in enumerate(prompt_files):
        prompt_text = read_text_file(prompt_file)
        is_last = index == len(prompt_files) - 1
        if is_last:
            full_prompt = build_prompt_with_previous_steps(subject_name, prompt_text, prompt_runs)
            images = None
        else:
            full_prompt = build_prompt_with_documents(
                subject_name=subject_name,
                base_prompt=prompt_text,
                documents=prepared.text_documents,
                image_count=len(prepared.images_base64),
            )
            images = prepared.images_base64

        response, parsed_json = call_until_valid_json(settings, full_prompt, images)
        prompt_runs.append(
            save_prompt_result_json(
                intermediate_dir=intermediate_dir,
                prompt_file=prompt_file,
                subject_name=subject_name,
                docs=prepared,
                ollama_response=response,
                parsed_json=parsed_json,
            )
        )

    report_docx_path = output_subject_dir / f"{subject_dir.name}.docx"
    build_word_report(settings.template_path, report_docx_path, prompt_runs)
    return intermediate_dir, report_docx_path


def main() -> None:
    settings = load_settings(SETTINGS_PATH)
    subject_dirs = list_subject_folders(settings.input_dir)
    for subject_dir in subject_dirs:
        intermediate_dir, report_docx_path = process_subject_folder(settings, subject_dir)
        print(f"Промежуточные JSON: {intermediate_dir}")
        print(f"Итоговый Word: {report_docx_path}")


if __name__ == "__main__":
    main()
