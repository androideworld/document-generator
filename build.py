"""Сборка «Генератора документов» в один исполняемый файл через PyInstaller.

Запуск:
    python build.py

Требования:
    pip install pyinstaller

Результат:
    Linux:   dist/ГенераторДокументов
    Windows: dist/ГенераторДокументов.exe
    macOS:   dist/ГенераторДокументов
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


APP_NAME = "ГенераторДокументов"
ENTRY_POINT = "main.py"
ICON_FILE = "resources/icon.ico"   # на Linux/macOS может отсутствовать


def check_pyinstaller() -> bool:
    try:
        import PyInstaller  # noqa: F401
        return True
    except ImportError:
        return False


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
        # Явные hidden-imports (иногда PyInstaller их не подхватывает)
        "--hidden-import", "openpyxl",
        "--hidden-import", "docx",
        "--hidden-import", "lxml",
        "--hidden-import", "tkinter",
        "--hidden-import", "tkinter.ttk",
        "--hidden-import", "tkinter.filedialog",
        "--hidden-import", "tkinter.messagebox",
    ]

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

    args.append(ENTRY_POINT)

    print("Запускаю PyInstaller:")
    print("  " + " ".join(args))
    print()

    result = subprocess.run(args, cwd=str(project_root))
    return result.returncode


def main():
    if not check_pyinstaller():
        print("PyInstaller не установлен.")
        print("Установите: pip install pyinstaller")
        sys.exit(1)

    rc = build()

    if rc == 0:
        print("\n✅ Сборка завершена успешно.")
        print(f"   Исполняемый файл: dist/{APP_NAME}" +
              (".exe" if sys.platform == "win32" else ""))
        print()
        print("Напоминание: для PDF/DOC/RTF/ODT требуется LibreOffice,")
        print("установленный в системе (или рядом с утилитой).")
    else:
        print(f"\n❌ Ошибка сборки. Код возврата: {rc}")
        sys.exit(rc)


if __name__ == "__main__":
    main()
