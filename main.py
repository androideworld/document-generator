"""Генератор документов — точка входа (GUI).

При запуске из .exe (PyInstaller) проверяет и при необходимости создаёт
портабельную структуру рядом с .exe:
    clients.xlsx
    Шаблоны/
        Пример_Договор.docx
    Готовые документы/
    config.json
"""

import sys
from pathlib import Path


def _setup_portable_structure():
    """Создаёт структуру папок при запуске из .exe.

    При запуске из исходников — только если рядом с main.py уже есть
    config.json (пользователь сознательно выбрал портабельный режим).
    """
    from core import config

    is_frozen = getattr(sys, "frozen", False)
    install_dir = config.get_install_dir()

    # Запуск из исходников и нет config.json рядом с main.py —
    # не трогаем файловую систему, используем классический конфиг.
    if not is_frozen and not (install_dir / "config.json").exists():
        return

    try:
        paths = config.ensure_portable_structure()
        print(f"[STARTUP] Портабельная структура готова: {paths}")
    except Exception as e:
        # Не валим запуск из-за структуры — просто предупреждение
        print(f"[STARTUP] Не удалось создать структуру: {e}", file=sys.stderr)


def main():
    _setup_portable_structure()

    from ui.main_window import MainWindow
    app = MainWindow()
    app.run()


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"[ОШИБКА] {e}", file=sys.stderr)
        sys.exit(1)
