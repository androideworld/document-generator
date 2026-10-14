"""Сохранение и загрузка настроек в config.json.

Расположение:
    Windows: %APPDATA%\\ГенераторДокументов\\config.json
    Linux/macOS: ~/.config/ГенераторДокументов/config.json
"""

import json
import os
import sys
from pathlib import Path
from typing import Any, Dict


APP_DIR_NAME = "ГенераторДокументов"


def get_config_dir() -> Path:
    """Возвращает папку для конфига в зависимости от ОС."""
    if sys.platform == "win32":
        base = os.environ.get("APPDATA") or str(Path.home() / "AppData" / "Roaming")
        return Path(base) / APP_DIR_NAME
    xdg = os.environ.get("XDG_CONFIG_HOME")
    if xdg:
        return Path(xdg) / APP_DIR_NAME
    return Path.home() / ".config" / APP_DIR_NAME


def get_config_path() -> Path:
    """Полный путь к config.json."""
    return get_config_dir() / "config.json"


def load_config() -> Dict[str, Any]:
    """Загружает config.json. При ошибке возвращает пустой словарь."""
    path = get_config_path()
    if not path.exists():
        return {}
    try:
        with path.open("r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            return {}
        return data
    except (json.JSONDecodeError, OSError):
        return {}


def save_config(config: Dict[str, Any]) -> None:
    """Сохраняет config.json атомарно."""
    path = get_config_path()
    path.parent.mkdir(parents=True, exist_ok=True)

    tmp = path.with_suffix(".json.tmp")
    with tmp.open("w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)

    tmp.replace(path)


def get(key: str, default: Any = None) -> Any:
    return load_config().get(key, default)


def set_value(key: str, value: Any) -> None:
    cfg = load_config()
    cfg[key] = value
    save_config(cfg)


def update(**kwargs) -> None:
    cfg = load_config()
    cfg.update(kwargs)
    save_config(cfg)

# ---------------------------------------------------------------------------
# Шаблоны путей для генерации
# ---------------------------------------------------------------------------

DEFAULT_FOLDER_PATTERN = "{year}/{month}"
DEFAULT_CLIENT_PATTERN = "{short_fio}_{date}_{contract}"


def get_folder_pattern() -> str:
    """Шаблон верхних папок.

    По умолчанию: {year}/{month}
    Пример: 2024/01

    Доступные подстановки:
        {year}, {month}, {first_letter}
    """
    return get("folder_pattern", DEFAULT_FOLDER_PATTERN)


def get_client_pattern() -> str:
    """Шаблон папки клиента.

    По умолчанию: {short_fio}_{date}_{contract}
    Пример: Иванов И.И._01.01.2024_101

    Доступные подстановки:
        {short_fio}, {fio}, {date}, {contract}, {id}, {phone}
    """
    return get("client_pattern", DEFAULT_CLIENT_PATTERN)

if __name__ == "__main__":
    print("Платформа:", sys.platform)
    print("Папка конфига:", get_config_dir())
    print("Путь к config.json:", get_config_path())

    print("\nТекущий конфиг:", load_config())
    print("\nДемонстрация get():")
    print("  по ключу 'excel_path':", get("excel_path"))
    print("  по ключу 'несуществующий':", get("несуществующий", "default"))
