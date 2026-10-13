# Генератор документов

Кроссплатформенная утилита для генерации документов Word/PDF из шаблонов
на основе данных из Excel-файла. Не требует Microsoft Office.

## Возможности

- Чтение клиентов из `clients.xlsx` (первый лист, первая строка — шапка).
- Фильтрация по 4 полям: ФИО, ID, НомерДоговора, Телефон — без учёта регистра.
- Замена маркеров `{{...}}` в Word-шаблонах: параграфы, таблицы, колонтитулы.
- Вывод в форматах: DOCX, PDF, PDF/A, DOC, RTF, ODT, TXT, XPS (при поддержке).
- Логирование в `generation_log.txt` в папке вывода.
- Сохранение настроек в `~/.config/ГенераторДокументов/config.json` (Linux/macOS)
  или `%APPDATA%\ГенераторДокументов\config.json` (Windows).

## Маркеры в шаблонах

| Маркер | Значение |
|---|---|
| `{{ClientName}}` | ФИО |
| `{{Amount}}` | Сумма (с разделителем тысяч) |
| `{{ContractDate}}` | Дата (DD.MM.YYYY) |
| `{{ContractNum}}` | Номер договора |
| `{{Address}}` | Адрес |
| `{{ID}}` | ID клиента |
| `{{Phone}}` | Телефон |

## Требования

- **Python 3.11+** (проверено на 3.14).
- **LibreOffice** — для форматов, кроме DOCX и TXT.
  - Fedora: `sudo dnf install libreoffice-writer`
  - Ubuntu/Debian: `sudo apt install libreoffice-writer`
  - macOS: `brew install --cask libreoffice`
  - Windows: [libreoffice.org](https://www.libreoffice.org/download/)

### Python-зависимости

    pip install -r requirements.txt

## Запуск

    python main.py

## Сборка

    pip install pyinstaller
    python build.py

Результат:

- Linux: `dist/ГенераторДокументов`
- Windows: `dist/ГенераторДокументов.exe`
- macOS: `dist/ГенераторДокументов`

> Сборка `.exe` возможна **только в Windows** — PyInstaller не умеет
> кросс-компиляцию.

## Структура проекта

    generator/
      main.py
      ui/main_window.py
      core/excel_reader.py
      core/docx_generator.py
      core/converter.py
      core/formats.py
      core/config.py
      core/logger.py
      resources/
      requirements.txt
      build.py
      README.md
