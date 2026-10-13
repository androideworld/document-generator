"""Список поддерживаемых форматов вывода и их свойства."""

from dataclasses import dataclass
from typing import Dict, List


@dataclass(frozen=True)
class FormatInfo:
    """Описание одного формата вывода."""
    code: str            # внутренний код: 'docx', 'pdf', 'pdfa', ...
    title: str           # человекочитаемое имя для GUI
    extension: str       # расширение файла: '.docx', '.pdf', ...
    suffix: str          # суффикс имени до расширения (например, '_pdfa')
    needs_libreoffice: bool
    lo_target: str       # что передать в soffice --convert-to (для non-LO пусто)
    lo_filter: str       # опциональный фильтр (например, 'writer_pdf_Export')


# Порядок важен: он используется в GUI для Combobox
FORMATS: Dict[str, FormatInfo] = {
    "docx": FormatInfo(
        code="docx", title="Word (.docx)", extension=".docx",
        suffix="", needs_libreoffice=False,
        lo_target="", lo_filter="",
    ),
    "pdf": FormatInfo(
        code="pdf", title="PDF (.pdf)", extension=".pdf",
        suffix="", needs_libreoffice=True,
        lo_target="pdf", lo_filter="",
    ),
    "pdfa": FormatInfo(
        code="pdfa", title="PDF/A (_pdfa.pdf)", extension=".pdf",
        suffix="_pdfa", needs_libreoffice=True,
        lo_target="pdf", lo_filter="writer_pdf_Export",
    ),
    "doc": FormatInfo(
        code="doc", title="Word 97-2003 (.doc)", extension=".doc",
        suffix="", needs_libreoffice=True,
        lo_target="doc", lo_filter="",
    ),
    "rtf": FormatInfo(
        code="rtf", title="Rich Text (.rtf)", extension=".rtf",
        suffix="", needs_libreoffice=True,
        lo_target="rtf", lo_filter="",
    ),
    "odt": FormatInfo(
        code="odt", title="OpenDocument (.odt)", extension=".odt",
        suffix="", needs_libreoffice=True,
        lo_target="odt", lo_filter="",
    ),
    "txt": FormatInfo(
        code="txt", title="Текст (.txt)", extension=".txt",
        suffix="", needs_libreoffice=False,
        lo_target="", lo_filter="",
    ),
    "xps": FormatInfo(
        code="xps", title="XPS (.xps)", extension=".xps",
        suffix="", needs_libreoffice=True,
        lo_target="xps", lo_filter="",
    ),
}


def get_format(code: str) -> FormatInfo:
    """Возвращает FormatInfo по коду или кидает KeyError."""
    if code not in FORMATS:
        raise KeyError(f"Неизвестный формат: {code!r}. "
                       f"Доступные: {sorted(FORMATS)}")
    return FORMATS[code]


def list_titles() -> List[str]:
    """Список человекочитаемых названий для GUI."""
    return [f.title for f in FORMATS.values()]


def by_title(title: str) -> FormatInfo:
    """Обратный поиск FormatInfo по человекочитаемому названию."""
    for f in FORMATS.values():
        if f.title == title:
            return f
    raise KeyError(f"Неизвестный заголовок формата: {title!r}")
