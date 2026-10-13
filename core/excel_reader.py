"""Чтение клиентов из clients.xlsx через openpyxl."""

from dataclasses import dataclass, asdict
from datetime import datetime, date
from typing import List
import openpyxl


@dataclass
class Client:
    id: str
    fio: str
    sum: str
    date: str
    contract_num: str
    address: str
    phone: str

    def to_dict(self) -> dict:
        return asdict(self)


def _normalize_id(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _normalize_text(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _normalize_date(value) -> str:
    if value is None or value == "":
        return ""
    if isinstance(value, (datetime, date)):
        return value.strftime("%d.%m.%Y")
    text = str(value).strip()
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(text, fmt).strftime("%d.%m.%Y")
        except ValueError:
            continue
    return text


def _cell(row: tuple, index: int) -> str:
    if index < len(row):
        return row[index]
    return ""


def read_clients(path: str) -> List[Client]:
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.active

    clients: List[Client] = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        if row is None:
            continue

        client_id = _normalize_id(_cell(row, 0))
        if not client_id:
            continue

        clients.append(Client(
            id=client_id,
            fio=_normalize_text(_cell(row, 1)),
            sum=_normalize_text(_cell(row, 2)),
            date=_normalize_date(_cell(row, 3)),
            contract_num=_normalize_text(_cell(row, 4)),
            address=_normalize_text(_cell(row, 5)),
            phone=_normalize_text(_cell(row, 6)),
        ))

    wb.close()
    return clients

def filter_clients(
    clients: List[Client],
    fio: str = "",
    id: str = "",
    contract: str = "",
    phone: str = "",
) -> List[Client]:
    """
    Фильтрует клиентов по 4 полям без учёта регистра.
    Пустые поля фильтра игнорируются.
    Совпадение — подстрока (contains), а не точное равенство.
    """
    # Нормализуем фильтры: убираем пробелы и приводим к нижнему регистру
    q_fio = fio.strip().lower()
    q_id = id.strip().lower()
    q_contract = contract.strip().lower()
    q_phone = phone.strip().lower()

    # Нормализуем телефон для сравнения:
    # оставляем только цифры — чтобы "+7 900 123-45-67" и "79001234567"
    # совпадали
    def phone_digits(value: str) -> str:
        return "".join(ch for ch in value if ch.isdigit())

    q_phone_digits = phone_digits(q_phone)

    result: List[Client] = []
    for c in clients:
        # ФИО — подстрока без учёта регистра
        if q_fio and q_fio not in c.fio.lower():
            continue

        # ID — подстрока без учёта регистра (работает и для чисел, и для строк)
        if q_id and q_id not in c.id.lower():
            continue

        # Номер договора — подстрока без учёта регистра
        if q_contract and q_contract not in c.contract_num.lower():
            continue

        # Телефон — сравнение по цифрам (без +, пробелов, дефисов)
        if q_phone_digits:
            if q_phone_digits not in phone_digits(c.phone):
                continue

        result.append(c)

    return result
if __name__ == "__main__":
    import sys

    excel_path = sys.argv[1] if len(sys.argv) > 1 else "clients.xlsx"

    try:
        clients = read_clients(excel_path)
    except FileNotFoundError:
        print(f"[ОШИБКА] Файл не найден: {excel_path}")
        sys.exit(1)
    except Exception as e:
        print(f"[ОШИБКА] Не удалось прочитать файл: {e}")
        sys.exit(1)

    print(f"Прочитано клиентов: {len(clients)}")
    print("=" * 70)

    for i, c in enumerate(clients, 1):
        print(f"{i:>3}. ID={c.id} | {c.fio} | дог. {c.contract_num} | {c.phone}")
        print(f"     сумма={c.sum} | дата={c.date} | адрес={c.address}")

    # --- Демонстрация фильтрации ---
    print("\n" + "=" * 70)
    print("Демонстрация фильтрации:")
    print("=" * 70)

    cases = [
        ("ФИО: 'иван' (регистр не важен)", dict(fio="иван")),
        ("ФИО: 'ПЕТРОВ' (верхний регистр)", dict(fio="ПЕТРОВ")),
        ("ID: '1' (подстрока — совпадёт с 1 и ... не с 2)", dict(id="1")),
        ("Договор: 'д-00'", dict(contract="д-00")),
        ("Телефон: '900123' (формат с + и дефисами)", dict(phone="900123")),
        ("Телефон: '+7 900 765' (с пробелами)", dict(phone="+7 900 765")),
        ("ФИО + Договор: 'иван' + '001'", dict(fio="иван", contract="001")),
        ("Ничего не найдено", dict(fio="КтотоНесуществующий")),
    ]

    for label, kwargs in cases:
        found = filter_clients(clients, **kwargs)
        print(f"\n[{label}] → найдено: {len(found)}")
        for c in found:
            print(f"    ID={c.id} | {c.fio} | дог. {c.contract_num} | {c.phone}")
