"""Генератор документов — точка входа (GUI)."""

import sys

from ui.main_window import MainWindow


def main():
    app = MainWindow()
    app.run()


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        # Fallback: сообщение в консоль, если GUI упал на старте
        print(f"[ОШИБКА] {e}", file=sys.stderr)
        sys.exit(1)
