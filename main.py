"""
Генератор PDF-чеков.

Скрипт читает данные о чеках из папки data/ (CSV и JSON),
шаблоны из папки templates/ (HTML + Jinja2) и создаёт PDF
в папке output/ с помощью WeasyPrint.

Запуск:  python main.py
"""

import json
import os
import subprocess
import sys
from pathlib import Path

# ---------------------------------------------------------------------------
# Пути проекта. Всё считаем относительно папки, где лежит main.py,
# поэтому скрипт работает, из какой бы папки его ни запустили.
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
TEMPLATES_DIR = BASE_DIR / "templates"
OUTPUT_DIR = BASE_DIR / "output"
FONTS_DIR = BASE_DIR / "fonts"

# Папка с DLL-библиотеками GTK3, которые нужны WeasyPrint на Windows
GTK_BIN_DIR = Path(r"C:\Program Files\GTK3-Runtime Win64\bin")

# Поля, которые обязательно должны быть в каждой строке данных
REQUIRED_FIELDS = ["invoice_id", "date", "customer", "product", "price", "qty"]


# ---------------------------------------------------------------------------
# Подготовка GTK и импорт WeasyPrint.
# ВАЖНО: путь к DLL нужно добавить ДО "import weasyprint",
# иначе WeasyPrint не найдёт библиотеки и упадёт при импорте.
# ---------------------------------------------------------------------------
if sys.platform == "win32" and GTK_BIN_DIR.is_dir():
    os.add_dll_directory(str(GTK_BIN_DIR))

try:
    from weasyprint import HTML
except (ImportError, OSError) as error:
    print("Не удалось загрузить WeasyPrint.")
    print(f"Причина: {error}")
    print("Проверьте, что пакет установлен (pip install -r requirements.txt)")
    print(f"и что GTK3 Runtime установлен в папку: {GTK_BIN_DIR}")
    sys.exit(1)

import pandas as pd
from jinja2 import Environment, FileSystemLoader, TemplateError, select_autoescape


class DataError(Exception):
    """Ошибка в файле с данными (битый файл, нет нужных полей и т.п.)."""


# ---------------------------------------------------------------------------
# Чтение данных
# ---------------------------------------------------------------------------
def read_csv_file(path: Path) -> list[dict]:
    """Читает CSV через pandas и возвращает список строк-словарей."""
    try:
        # utf-8-sig корректно читает и обычный UTF-8, и UTF-8 с BOM (так сохраняет Excel)
        df = pd.read_csv(path, encoding="utf-8-sig", dtype=str)
    except (pd.errors.ParserError, pd.errors.EmptyDataError, UnicodeDecodeError) as error:
        raise DataError(f"не удалось прочитать CSV: {error}") from error
    # Пустые ячейки pandas превращает в NaN — заменим их на пустую строку
    df = df.fillna("")
    return df.to_dict(orient="records")


def read_json_file(path: Path) -> list[dict]:
    """Читает JSON стандартным модулем json. Ожидается список объектов."""
    try:
        with path.open(encoding="utf-8") as file:
            data = json.load(file)
    except (json.JSONDecodeError, UnicodeDecodeError) as error:
        raise DataError(f"не удалось прочитать JSON: {error}") from error

    if not isinstance(data, list) or not all(isinstance(row, dict) for row in data):
        raise DataError("JSON должен содержать список объектов вида [{...}, {...}]")
    return data


def normalize_rows(rows: list[dict]) -> list[dict]:
    """
    Проверяет, что в каждой строке есть все нужные поля,
    и приводит типы: price -> float, qty -> int, остальное -> str.
    """
    if not rows:
        raise DataError("в файле нет ни одной строки с данными")

    result = []
    for number, row in enumerate(rows, start=1):
        missing = [field for field in REQUIRED_FIELDS if field not in row]
        if missing:
            raise DataError(f"строка {number}: нет полей {', '.join(missing)}")
        try:
            price = float(str(row["price"]).replace(",", "."))
            qty = int(float(str(row["qty"]).replace(",", ".")))
        except ValueError:
            raise DataError(
                f"строка {number}: price и qty должны быть числами "
                f"(сейчас price={row['price']!r}, qty={row['qty']!r})"
            ) from None

        result.append({
            "invoice_id": str(row["invoice_id"]).strip(),
            "date": str(row["date"]).strip(),
            "customer": str(row["customer"]).strip(),
            "product": str(row["product"]).strip(),
            "price": price,
            "qty": qty,
        })
    return result


def load_rows(path: Path) -> list[dict]:
    """Читает файл данных (CSV или JSON) и возвращает проверенные строки."""
    try:
        if path.suffix.lower() == ".csv":
            rows = read_csv_file(path)
        else:
            rows = read_json_file(path)
    except OSError as error:
        raise DataError(f"не удалось открыть файл: {error}") from error
    return normalize_rows(rows)


def group_invoices(rows: list[dict]) -> dict[str, list[dict]]:
    """
    Группирует строки-товары по invoice_id.
    Возвращает словарь {invoice_id: [строки этого чека]} в порядке появления.
    """
    invoices: dict[str, list[dict]] = {}
    for row in rows:
        invoices.setdefault(row["invoice_id"], []).append(row)
    return invoices


def build_context(invoice_id: str, rows: list[dict]) -> dict:
    """Собирает данные для шаблона: шапку чека, товары и итоговую сумму."""
    items = []
    for row in rows:
        items.append({
            "product": row["product"],
            "price": row["price"],
            "qty": row["qty"],
            "line_total": row["price"] * row["qty"],  # сумма по строке
        })

    first = rows[0]  # дата и клиент одинаковые для всех строк чека
    return {
        "invoice_id": invoice_id,
        "date": first["date"],
        "customer": first["customer"],
        "items": items,
        "total": sum(item["line_total"] for item in items),  # общий итог
    }


# ---------------------------------------------------------------------------
# Консольное меню
# ---------------------------------------------------------------------------
def print_menu(title: str, options: list[str]) -> None:
    """Печатает заголовок и пронумерованный список."""
    print()
    print(title)
    print("-" * len(title))
    for number, option in enumerate(options, start=1):
        print(f"  {number}. {option}")


def ask_number(prompt: str, count: int) -> int:
    """
    Просит ввести номер от 1 до count. Повторяет вопрос, пока ввод неверный.
    Возвращает индекс (от 0). Ввод 0 или q — выход из программы.
    """
    while True:
        answer = input(f"{prompt} (1–{count}, 0 — выход): ").strip().lower()
        if answer in ("0", "q"):
            print("Выход.")
            sys.exit(0)
        if answer.isdigit() and 1 <= int(answer) <= count:
            return int(answer) - 1
        print(f"  ⚠ Введите число от 1 до {count}.")


def list_files(folder: Path, extensions: tuple[str, ...]) -> list[Path]:
    """Возвращает отсортированный список файлов с нужными расширениями."""
    if not folder.is_dir():
        return []
    return sorted(
        path for path in folder.iterdir()
        if path.is_file() and path.suffix.lower() in extensions
    )


# ---------------------------------------------------------------------------
# Генерация PDF и открытие файла
# ---------------------------------------------------------------------------
def money(value: float) -> str:
    """Фильтр для шаблона: 1234.5 -> '1 234,50'."""
    return f"{value:,.2f}".replace(",", " ").replace(".", ",")


def render_html(template_name: str, context: dict) -> str:
    """Подставляет данные в HTML-шаблон через Jinja2."""
    env = Environment(
        loader=FileSystemLoader(TEMPLATES_DIR, encoding="utf-8"),
        autoescape=select_autoescape(["html"]),  # защита от "поломки" HTML спецсимволами
    )
    env.filters["money"] = money
    template = env.get_template(template_name)
    return template.render(**context)


def generate_pdf(html: str, output_path: Path) -> None:
    """Превращает HTML в PDF через WeasyPrint."""
    # base_url нужен, чтобы относительные пути в шаблоне (fonts/DejaVuSans.ttf)
    # искались от папки проекта. Слэш в конце обязателен — это папка.
    base_url = BASE_DIR.as_uri() + "/"
    HTML(string=html, base_url=base_url).write_pdf(output_path)


def open_file(path: Path) -> None:
    """Открывает файл программой по умолчанию в системе."""
    try:
        if sys.platform == "win32":
            os.startfile(path)  # есть только на Windows
        elif sys.platform == "darwin":
            subprocess.run(["open", str(path)], check=False)
        else:
            subprocess.run(["xdg-open", str(path)], check=False)
    except OSError as error:
        print(f"Не удалось открыть PDF автоматически: {error}")
        print(f"Откройте его вручную: {path}")


# ---------------------------------------------------------------------------
# Главная функция
# ---------------------------------------------------------------------------
def main() -> None:
    print("=" * 40)
    print("   ГЕНЕРАТОР PDF-ЧЕКОВ")
    print("=" * 40)

    # 1. Файлы с данными
    data_files = list_files(DATA_DIR, (".csv", ".json"))
    if not data_files:
        print(f"В папке {DATA_DIR} нет файлов .csv или .json. Добавьте данные и запустите снова.")
        return

    # 2. Шаблоны
    templates = list_files(TEMPLATES_DIR, (".html",))
    if not templates:
        print(f"В папке {TEMPLATES_DIR} нет HTML-шаблонов. Добавьте шаблон и запустите снова.")
        return

    # Предупредим, если шрифта нет: PDF создастся, но кириллица может отобразиться криво
    if not (FONTS_DIR / "DejaVuSans.ttf").is_file():
        print(f"⚠ Не найден шрифт {FONTS_DIR / 'DejaVuSans.ttf'} — см. README/инструкцию.")

    # 3. Показываем списки
    print_menu("Файлы с данными", [path.name for path in data_files])
    print_menu("Шаблоны", [path.name for path in templates])
    print()

    # 4. Выбор файла данных. Если файл битый — сообщаем и просим выбрать другой.
    while True:
        data_path = data_files[ask_number("Выберите файл данных", len(data_files))]
        try:
            invoices = group_invoices(load_rows(data_path))
            break
        except DataError as error:
            print(f"  ⚠ Ошибка в файле {data_path.name}: {error}")
            print("  Выберите другой файл.")

    template_path = templates[ask_number("Выберите шаблон", len(templates))]

    # 5. Список чеков из выбранного файла
    invoice_ids = list(invoices)
    options = []
    for invoice_id in invoice_ids:
        ctx = build_context(invoice_id, invoices[invoice_id])
        options.append(
            f"№ {invoice_id} — {ctx['customer']}, {ctx['date']}, "
            f"товаров: {len(ctx['items'])}, итого {money(ctx['total'])} ₽"
        )
    print_menu(f"Чеки в файле {data_path.name}", options)
    print()

    # 6. Выбор чека и генерация PDF
    invoice_id = invoice_ids[ask_number("Выберите чек", len(invoice_ids))]
    context = build_context(invoice_id, invoices[invoice_id])

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)  # создаём output/, если её нет
    output_path = OUTPUT_DIR / f"invoice_{invoice_id}.pdf"

    try:
        html = render_html(template_path.name, context)
        generate_pdf(html, output_path)
    except TemplateError as error:
        print(f"Ошибка в шаблоне {template_path.name}: {error}")
        return
    except PermissionError:
        print(f"Не удалось записать {output_path.name}: файл занят.")
        print("Закройте этот PDF в программе просмотра и попробуйте снова.")
        return
    except Exception as error:  # любые другие ошибки WeasyPrint
        print(f"Не удалось создать PDF: {error}")
        return

    print(f"\n✅ Готово! PDF сохранён: {output_path}")

    # 7. Открываем PDF системной программой
    open_file(output_path)


if __name__ == "__main__":
    try:
        main()
    except (KeyboardInterrupt, EOFError):
        # Ctrl+C или закрытый ввод — выходим без длинного traceback
        print("\nВыход.")
