"""Сборка «DocForge» в один исполняемый файл через PyInstaller.

Запуск:
    python build.py

Требования:
    pip install pyinstaller

Результат:
    Linux:   dist/DocForge
    Windows: dist/DocForge.exe
    macOS:   dist/DocForge
"""

import shutil
import subprocess
import sys
from pathlib import Path

# --- Принудительный UTF-8 для stdout/stderr (Windows + CI) ---
# На Windows консоль по умолчанию использует CP1252, из-за чего print()
# с русскими буквами падает с UnicodeEncodeError. Переводим на UTF-8.
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


# Имя файла .exe / бинарника (без расширения)
APP_NAME = "DocForge"

# Точка входа
ENTRY_POINT = "main.py"

# Иконка (если есть)
ICON_FILE = "resources/icon.ico"


def check_pyinstaller() -> bool:
    try:
        import PyInstaller  # noqa: F401
        return True
    except ImportError:
        return False


def _make_windows_version_file(project_root: Path) -> str | None:
    """Создаёт version_info.txt для PyInstaller (Windows).

    Возвращает путь к файлу или None, если не Windows.
    """
    if sys.platform != "win32":
        return None

    try:
        from core.version import (
            VERSION, VERSION_TUPLE, APP_NAME as V_APP_NAME,
            APP_DESCRIPTION, APP_AUTHOR, APP_COPYRIGHT,
        )
    except ImportError:
        print("Не удалось импортировать core.version — пропускаю version-file.")
        return None

    # Дополняем до 4 компонентов (PyInstaller требует)
    v = list(VERSION_TUPLE) + [0] * (4 - len(VERSION_TUPLE))
    v_str = ".".join(str(x) for x in v)

    content = f"""VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=({v[0]}, {v[1]}, {v[2]}, {v[3]}),
    prodvers=({v[0]}, {v[1]}, {v[2]}, {v[3]}),
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo([
      StringTable(
        '040904B0',
        [StringStruct('CompanyName', '{APP_AUTHOR}'),
         StringStruct('FileDescription', '{APP_DESCRIPTION}'),
         StringStruct('FileVersion', '{v_str}'),
         StringStruct('InternalName', '{V_APP_NAME}'),
         StringStruct('LegalCopyright', '{APP_COPYRIGHT}'),
         StringStruct('OriginalFilename', '{V_APP_NAME}.exe'),
         StringStruct('ProductName', '{V_APP_NAME}'),
         StringStruct('ProductVersion', '{v_str}')])
    ]),
    VarFileInfo([VarStruct('Translation', [1033, 1200])])
  ]
)"""

    path = project_root / "version_info.txt"
    path.write_text(content, encoding="utf-8")
    return str(path)


def build() -> int:
    project_root = Path(__file__).parent.resolve()
    dist_dir = project_root / "dist"
    build_dir = project_root / "build"

    # Чистим предыдущие сборки
    for d in (dist_dir, build_dir):
        if d.exists():
            print(f"Удаляю {d}...")
            shutil.rmtree(d)

    # Формируем аргументы
    args = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--windowed",              # без консольного окна
        "--name", APP_NAME,
        "--collect-data", "docx",        # ← шаблоны для python-docx
        "--collect-data", "openpyxl",
        # Явные hidden-imports (иногда PyInstaller их не подхватывает)
        "--hidden-import", "openpyxl",
        "--hidden-import", "docx",
        "--hidden-import", "lxml",
        "--hidden-import", "tkinter",
        "--hidden-import", "tkinter.ttk",
        "--hidden-import", "tkinter.filedialog",
        "--hidden-import", "tkinter.messagebox",
    ]

    # Метаданные для .exe (только Windows)
    vf = _make_windows_version_file(project_root)
    if vf:
        args += ["--version-file", vf]
        print(f"Добавлены метаданные версии: {vf}")

    # Иконка — только если файл существует
    icon_path = project_root / ICON_FILE
    if icon_path.exists():
        args += ["--icon", str(icon_path)]
    else:
        print(f"Иконка не найдена ({ICON_FILE}) — сборка без иконки.")

    # Дополнительные данные: если в resources/ есть что-то ещё, кладём рядом
    resources_dir = project_root / "resources"
    if resources_dir.exists():
        sep = ";" if sys.platform == "win32" else ":"
        args += ["--add-data", f"{resources_dir}{sep}resources"]

    # --- Данные python-docx: templates/ и parts/ ---
    # --collect-data не всегда подхватывает вложенные папки
    try:
        import docx as _docx_module
        from pathlib import Path as _Path
        docx_root = _Path(_docx_module.__file__).parent
        sep = ";" if sys.platform == "win32" else ":"

        # Подпапки, которые нужны python-docx при создании документов
        for sub in ("templates", "parts"):
            sub_path = docx_root / sub
            if sub_path.exists():
                args += [
                    "--add-data",
                    f"{sub_path}{sep}docx/{sub}",
                ]
                print(f"Добавлены данные docx/{sub}: {sub_path}")
    except Exception as e:
        print(f"Не удалось собрать данные docx: {e}")

    args.append(ENTRY_POINT)

    print("Запускаю PyInstaller:")
    print("  " + " ".join(args))
    print()

    result = subprocess.run(args, cwd=str(project_root))

    # Чистим временный version_info.txt
    vf_path = project_root / "version_info.txt"
    if vf_path.exists():
        vf_path.unlink()

    return result.returncode


def main():
    if not check_pyinstaller():
        print("PyInstaller не установлен.")
        print("Установите: pip install pyinstaller")
        sys.exit(1)

    rc = build()

    if rc == 0:
        print("\n✅ Сборка завершена успешно.")
        ext = ".exe" if sys.platform == "win32" else ""
        print(f"   Исполняемый файл: dist/{APP_NAME}{ext}")
        print()
        print("Напоминание: для PDF/DOC/RTF/ODT требуется LibreOffice,")
        print("установленный в системе (или рядом с утилитой).")
    else:
        print(f"\n❌ Ошибка сборки. Код возврата: {rc}")
        sys.exit(rc)


if __name__ == "__main__":
    main()
