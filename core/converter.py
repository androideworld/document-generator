"""Конвертация DOCX в PDF/DOC/RTF/ODT/XPS через LibreOffice headless
и в TXT — напрямую через python-docx."""

import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Optional

from docx import Document


# ---------------------------------------------------------------------------
# Поиск soffice
# ---------------------------------------------------------------------------

_SOFFICE_CANDIDATES = [
    "/usr/bin/soffice",
    "/usr/lib64/libreoffice/program/soffice",
    "/usr/lib/libreoffice/program/soffice",
    "/snap/bin/libreoffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
]


def find_soffice() -> Optional[str]:
    """Ищет soffice в PATH и в известных путях."""
    p = shutil.which("soffice") or shutil.which("libreoffice")
    if p:
        return p
    for c in _SOFFICE_CANDIDATES:
        if Path(c).exists():
            return c
    return None


# ---------------------------------------------------------------------------
# Конвертация через LibreOffice
# ---------------------------------------------------------------------------

def _build_convert_args(
    soffice: str,
    input_path: Path,
    outdir: Path,
    target: str,
    filter_name: str = "",
    filter_options: str = "",
) -> list:
    """
    Собирает argv для soffice.

    Синтаксис --convert-to: <target>[:<filter>[:<options>]]
    """
    convert_to = target
    if filter_name:
        convert_to += f":{filter_name}"
        if filter_options:
            convert_to += f":{filter_options}"

    return [
        soffice,
        "--headless",
        "--norestore",
        "--nolockcheck",
        "--nodefault",
        "--nofirststartwizard",
        "--convert-to", convert_to,
        "--outdir", str(outdir),
        str(input_path),
    ]


def convert_with_libreoffice(
    input_path: str,
    output_dir: str,
    target_format: str,
    filter_name: str = "",
    filter_options: str = "",
    timeout: int = 120,
) -> str:
    """Конвертирует файл через LibreOffice headless."""
    soffice = find_soffice()
    if not soffice:
        raise FileNotFoundError(
            "LibreOffice (soffice) не найден. Установите LibreOffice."
        )

    src = Path(input_path).resolve()
    if not src.exists():
        raise FileNotFoundError(f"Исходный файл не найден: {src}")

    outdir = Path(output_dir).resolve()
    outdir.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(prefix="lo_profile_") as profile_dir:
        args = _build_convert_args(
            soffice, src, outdir, target_format,
            filter_name, filter_options,
        )
        # На Windows URI профиля должен быть file:///C:/...
        # (прямые слэши, три после file:). Иначе LibreOffice падает.
        profile_uri = Path(profile_dir).as_uri()
        args.insert(2, f"-env:UserInstallation={profile_uri}")

        try:
            proc = subprocess.run(
                args,
                capture_output=True,
                text=True,
                timeout=timeout,
                check=False,
            )
        except subprocess.TimeoutExpired as e:
            raise RuntimeError(
                f"LibreOffice не завершился за {timeout} с."
            ) from e

    expected = outdir / (src.stem + "." + target_format)
    if not expected.exists():
        raise RuntimeError(
            f"LibreOffice не создал файл {expected}.\n"
            f"stdout: {proc.stdout}\n"
            f"stderr: {proc.stderr}\n"
            f"returncode: {proc.returncode}"
        )

    return str(expected)


# ---------------------------------------------------------------------------
# TXT — прямая запись текста из DOCX
# ---------------------------------------------------------------------------

def generate_txt(input_docx: str, output_path: str) -> str:
    """Извлекает текст из DOCX и сохраняет как .txt."""
    from docx.oxml.ns import qn
    from docx.text.paragraph import Paragraph
    from docx.table import Table

    src = Path(input_docx)
    if not src.exists():
        raise FileNotFoundError(f"DOCX не найден: {src}")

    doc = Document(str(src))
    lines = []

    body = doc.element.body
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            p = Paragraph(child, doc)
            lines.append(p.text)
        elif child.tag == qn("w:tbl"):
            t = Table(child, doc)
            for row in t.rows:
                cells = [c.text.replace("\n", " ").strip() for c in row.cells]
                lines.append("\t".join(cells))

    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text("\n".join(lines), encoding="utf-8")

    return str(out)


# ---------------------------------------------------------------------------
# Универсальная функция: конвертировать в целевой формат
# ---------------------------------------------------------------------------

def convert_to_format(input_docx: str, output_dir: str, fmt) -> str:
    """
    Конвертирует DOCX в указанный формат.

    :param input_docx: путь к .docx
    :param output_dir: папка для результата
    :param fmt: FormatInfo (из core.formats.get_format)
    :return: путь к созданному файлу
    """
    src = Path(input_docx)
    if not src.exists():
        raise FileNotFoundError(f"DOCX не найден: {src}")

    outdir = Path(output_dir)
    outdir.mkdir(parents=True, exist_ok=True)

    base_name = src.stem + fmt.suffix + fmt.extension
    target_path = outdir / base_name

    # TXT — без LibreOffice
    if fmt.code == "txt":
        return generate_txt(str(src), str(target_path))

    # DOCX — копия (или no-op, если src и target совпадают)
    if fmt.code == "docx":
        if src.resolve() == target_path.resolve():
            return str(target_path)
        shutil.copyfile(src, target_path)
        return str(target_path)

    # Остальные — через LibreOffice
    if not fmt.needs_libreoffice:
        raise ValueError(f"Формат {fmt.code!r} не обработан")

    with tempfile.TemporaryDirectory(prefix="lo_conv_") as tmpdir:
        tmp_input = Path(tmpdir) / (src.stem + ".docx")
        shutil.copyfile(src, tmp_input)

        try:
            result_path = convert_with_libreoffice(
                str(tmp_input), tmpdir,
                target_format=fmt.lo_target,
                filter_name=fmt.lo_filter,
                filter_options=(
                    '{"SelectPdfVersion":{"type":"long","value":1}}'
                    if fmt.code == "pdfa" else ""
                ),
            )
        except RuntimeError as e:
            if fmt.code == "xps":
                raise RuntimeError(
                    "XPS-фильтр не установлен в вашем LibreOffice. "
                    "Установите пакет libreoffice-xps (Fedora) "
                    "или libreoffice-impress (Debian/Ubuntu), "
                    "либо отключите формат XPS."
                ) from e
            raise

        shutil.copyfile(result_path, target_path)

    return str(target_path)

# ---------------------------------------------------------------------------
# CLI для быстрой проверки
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import sys

    print("soffice:", find_soffice() or "НЕ НАЙДЕН")

    if len(sys.argv) < 2:
        print("\nИспользование: python core/converter.py <input.docx> [outdir]")
        sys.exit(0)

    input_docx = sys.argv[1]
    outdir = sys.argv[2] if len(sys.argv) > 2 else "out"

    Path(outdir).mkdir(parents=True, exist_ok=True)

    txt = generate_txt(input_docx, str(Path(outdir) / (Path(input_docx).stem + ".txt")))
    print(f"TXT: {txt}")

    if find_soffice():
        for target in ("pdf", "doc", "rtf", "odt"):
            try:
                result = convert_with_libreoffice(input_docx, outdir, target)
                print(f"{target.upper()}: {result}")
            except Exception as e:
                print(f"{target.upper()}: ОШИБКА — {e}")
    else:
        print("LibreOffice не найден — пропускаем PDF/DOC/RTF/ODT")
