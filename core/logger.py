"""Логирование процесса генерации в generation_log.txt."""

import logging
from datetime import datetime
from pathlib import Path


LOG_FILENAME = "generation_log.txt"


def setup_logger(folder: str, level: int = logging.INFO) -> logging.Logger:
    """
    Настраивает логгер с записью в <folder>/generation_log.txt.
    Файл ПЕРЕЗАПИСЫВАЕТСЯ при каждом вызове (mode='w').

    :param folder: папка для лог-файла
    :param level: уровень логирования (по умолчанию INFO)
    :return: настроенный logging.Logger
    """
    folder_path = Path(folder)
    folder_path.mkdir(parents=True, exist_ok=True)

    log_path = folder_path / LOG_FILENAME

    # Уникальное имя логгера, чтобы не смешивать с другими вызовами
    logger = logging.getLogger(f"generator.{id(log_path)}")
    logger.setLevel(level)
    logger.handlers.clear()
    logger.propagate = False

    # Формат: [YYYY-MM-DD HH:MM:SS] LEVEL: сообщение
    formatter = logging.Formatter(
        fmt="[%(asctime)s] %(levelname)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    # Файл — перезапись
    file_handler = logging.FileHandler(log_path, mode="w", encoding="utf-8")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    # Консоль — тоже полезно видеть
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    return logger


def log_run_header(logger: logging.Logger, **params) -> None:
    """Пишет в лог шапку запуска с параметрами."""
    logger.info("=" * 70)
    logger.info("Генератор документов — запуск")
    logger.info("Время: %s", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    logger.info("-" * 70)
    for key, value in params.items():
        logger.info("%s: %s", key, value)
    logger.info("=" * 70)


def log_run_footer(logger: logging.Logger, ok: int, fail: int) -> None:
    """Пишет в лог итог: сколько успешно, сколько с ошибками."""
    logger.info("-" * 70)
    logger.info("Итог: успешно %d, с ошибками %d", ok, fail)
    logger.info("=" * 70)
