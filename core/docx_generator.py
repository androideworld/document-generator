"""Замена маркеров {{...}} в DOCX-шаблонах через python-docx."""

import re
from pathlib import Path
from typing import Iterable

from docx import Document
from docx.document import Document as DocumentType
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

from .excel_reader import Client


# ---------------------------------------------------------------------------
# Форматирование значений
# ---------------------------------------------------------------------------

def format_amount(value: str) -> str:
    """
    Приводит сумму к виду с разделителем тысяч (пробел, как принято в РФ).
    '150000'      → '150 000'
    '150000.50'   → '150 000.50'
    '150 000,50'  → '150 000,50'
    '' / None     → ''
    """
    if value is None:
        return ""
    text = str(value).strip()
    if not text:
        return ""

    # Отделяем целую часть от дробной (по точке или запятой)
    m = re.match(r"^(\d+)([.,]\d+)?$", text.replace(" ", ""))
    if not m:
        # Нестандартный формат — возвращаем как есть
        return text

    integer_part, frac_part = m.group(1), m.group(2) or ""
    # Вставляем пробел каждые 3 цифры справа
    grouped = f"{int(integer_part):,}".replace(",", " ")
    return grouped + frac_part


def format_date(value: str) -> str:
    """Оставляет дату как есть (уже нормализована в excel_reader)."""
    return (value or "").strip()


# ---------------------------------------------------------------------------
# Карта маркеров → значений
# ---------------------------------------------------------------------------

def build_marker_map(client: Client) -> dict:
    """
    Возвращает словарь {маркер: значение} для замены в шаблоне.
    Ключи — без фигурных скобок, чтобы удобно собирать регуляркой.
    """
    return {
        "ClientName":   client.fio,
        "Amount":       format_amount(client.sum),
        "ContractDate": format_date(client.date),
        "ContractNum":  client.contract_num,
        "Address":      client.address,
        "ID":           client.id,
        "Phone":        client.phone,
    }


# Регулярка: {{ MarkerName }} с любыми пробелами внутри
MARKER_RE = re.compile(r"\{\{\s*([A-Za-z_][A-Za-z0-9_]*)\s*\}\}")


def replace_in_text(text: str, marker_map: dict) -> str:
    """Заменяет все {{...}} в строке согласно карте. Неизвестные маркеры не трогает."""
    def repl(match: re.Match) -> str:
        key = match.group(1)
        return marker_map.get(key, match.group(0))
    return MARKER_RE.sub(repl, text)


# ---------------------------------------------------------------------------
# Обход всех текстовых контейнеров документа
# ---------------------------------------------------------------------------

def _iter_paragraphs(container) -> Iterable[Paragraph]:
    """Итерирует параграфы в документе или в ячейке таблицы."""
    for p in container.paragraphs:
        yield p
    for table in getattr(container, "tables", []):
        yield from _iter_table(table)


def _iter_table(table: Table) -> Iterable[Paragraph]:
    """Итерирует параграфы в таблице, включая вложенные таблицы."""
    for row in table.rows:
        for cell in row.cells:
            yield from _iter_paragraphs(cell)


def _replace_in_paragraph(paragraph: Paragraph, marker_map: dict) -> bool:
    """
    Заменяет маркеры в параграфе.
    Возвращает True, если что-то заменили.

    Важно: в Word один маркер может быть разбит на несколько runs
    (кусочков текста). Поэтому если после простой замены по runs
    маркер не найден, а в объединённом тексте он есть — пересобираем
    параграф заново.
    """
    changed = False

    # 1) Пробуем заменить внутри каждого run по отдельности (быстрый путь)
    for run in paragraph.runs:
        new_text = replace_in_text(run.text, marker_map)
        if new_text != run.text:
            run.text = new_text
            changed = True

    # 2) Если в runs остались незакрытые маркеры, значит маркер разбит
    full = paragraph.text
    if "{{" in full and MARKER_RE.search(full):
        new_full = replace_in_text(full, marker_map)
        if new_full != full:
            # Стираем все runs и пишем один run с новым текстом
            # (сохраняя формат первого run)
            for i, run in enumerate(paragraph.runs):
                if i == 0:
                    run.text = new_full
                else:
                    run.text = ""
            changed = True

    return changed


# ---------------------------------------------------------------------------
# Публичное API
# ---------------------------------------------------------------------------

def replace_markers(doc: DocumentType, client: Client) -> int:
    """
    Заменяет все маркеры в документе (параграфы, таблицы, колонтитулы).
    Возвращает количество обработанных параграфов.
    """
    marker_map = build_marker_map(client)
    counter = 0

    # Основной текст (включая таблицы)
    for paragraph in _iter_paragraphs(doc):
        if _replace_in_paragraph(paragraph, marker_map):
            counter += 1

    # Колонтитулы (верхний и нижний) во всех секциях
    for section in doc.sections:
        for header_footer in (section.header, section.footer):
            for paragraph in _iter_paragraphs(header_footer):
                if _replace_in_paragraph(paragraph, marker_map):
                    counter += 1

    return counter


def generate_docx(template_path: str, output_path: str, client: Client) -> None:
    """Открывает шаблон, заменяет маркеры, сохраняет результат."""
    tpl = Path(template_path)
    if not tpl.exists():
        raise FileNotFoundError(f"Шаблон не найден: {template_path}")

    doc = Document(str(tpl))
    replace_markers(doc, client)

    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
