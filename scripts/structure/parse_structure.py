# parse_structure.py
import argparse
import os
import re
import sys
from datetime import datetime
from typing import List, Dict, Tuple

import pandas as pd

try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None


SCRIPT_DIR = os.path.abspath(os.path.dirname(__file__))
BASE_DIR = os.path.abspath(os.path.join(SCRIPT_DIR, ".."))

# По умолчанию пишем Structure.xlsx в папку "Результаты_EXCEL"
DEFAULT_OUT = os.path.join(BASE_DIR, "Результаты_EXCEL", "Structure.xlsx")


# ---------------- Служебка ----------------

def archive_old(path: str) -> None:
    """
    Если файл существует, переносит его в подкаталог 'Архив' рядом с ним.
    """
    if not os.path.isfile(path):
        return

    folder = os.path.dirname(path)
    archive_dir = os.path.join(folder, "Архив")
    os.makedirs(archive_dir, exist_ok=True)

    base = os.path.basename(path)
    name, ext = os.path.splitext(base)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    new_name = f"{name}_{ts}{ext}"
    new_path = os.path.join(archive_dir, new_name)

    try:
        os.replace(path, new_path)
    except OSError as e:
        print(f"Не удалось переместить старый файл '{path}': {e}")


def normalize_cell(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


# ---------------- PDF: таблица "Система / Подсистема / Наименование" ----------------

def _extract_system_rows_from_pdf(pdf_path: str) -> List[Tuple[str, str, str]]:
    """
    Парсим PDF, вытаскиваем строки вида (код_системы, код_подсистемы, имя).
    Работает только внутри таблицы с заголовком:
      Система
      Подсистема
      Наименование
    """
    if fitz is None:
        raise RuntimeError(
            "Для парсинга PDF нужна библиотека PyMuPDF (модуль 'fitz')."
        )

    rows: List[Tuple[str, str, str]] = []

    doc = fitz.open(pdf_path)

    for page in doc:
        text = page.get_text("text")
        lines = [ln.strip() for ln in text.splitlines()]

        in_table = False
        current_system: str | None = None
        i = 0

        while i < len(lines):
            line = lines[i]
            low = line.lower()

            if not in_table:
                # Ищем заголовок "Система / Подсистема / Наименование"
                if (
                    low.startswith("система")
                    and i + 2 < len(lines)
                    and lines[i + 1].strip().lower().startswith("подсистема")
                    and lines[i + 2].strip().lower().startswith("наименование")
                ):
                    in_table = True
                    i += 3
                    continue
            else:
                # Выход из таблицы
                if low.startswith("перечень функций самолета"):
                    in_table = False
                    break

                if not line:
                    i += 1
                    continue

                m = re.fullmatch(r"\d{1,3}", line)
                if m:
                    num = int(line)

                    # Код системы: 21, 22, 23, ... (НЕ кратен 10)
                    if num % 10 != 0:
                        current_system = line
                        i += 1
                        continue

                    # Код подсистемы: 00, 10, 20, ...
                    sub_code = line

                    # Собираем название подсистемы (может быть многострочным)
                    i += 1
                    name_parts: List[str] = []
                    while i < len(lines):
                        s = lines[i].strip()
                        if not s:
                            i += 1
                            continue

                        if re.fullmatch(r"\d{1,3}", s):
                            break
                        if s.lower().startswith("перечень функций самолета"):
                            break

                        name_parts.append(s)
                        i += 1

                    name = " ".join(name_parts).strip()
                    if not name:
                        continue
                    if current_system is None:
                        # Странный случай – подсистема без системы, пропускаем
                        continue

                    rows.append((current_system, sub_code, name))
                    continue

            i += 1

    return rows


# ---------------- Excel: таблица "Система / Подсистема / Наименование" ----------------

def _extract_system_rows_from_excel(path: str) -> List[Tuple[str, str, str]]:
    """
    Ожидается таблица типа HB-17:

        ... | Система | Подсистема | Наименование/Функции | ...

    Ищем строку с такими заголовками, ниже собираем строки (Система, Подсистема, Имя).
    """
    all_sheets = pd.read_excel(path, sheet_name=None, header=None)
    rows: List[Tuple[str, str, str]] = []

    for sheet_name, df in all_sheets.items():
        header_row_idx: int | None = None
        col_sys = col_sub = col_name = None

        # Поиск строки заголовков
        for ridx in range(len(df)):
            row = df.iloc[ridx]
            for cidx in range(len(row) - 2):
                c0 = normalize_cell(row.iloc[cidx]).lower()
                c1 = normalize_cell(row.iloc[cidx + 1]).lower()
                c2 = normalize_cell(row.iloc[cidx + 2]).lower()

                if (
                    c0.startswith("система")
                    and c1.startswith("подсистема")
                    and (c2.startswith("наимен") or c2.startswith("функц"))
                ):
                    header_row_idx = ridx
                    col_sys = cidx
                    col_sub = cidx + 1
                    col_name = cidx + 2
                    break
            if header_row_idx is not None:
                break

        if header_row_idx is None:
            # На этом листе нет нужной таблицы
            continue

        current_system: str | None = None

        # Все строки ниже заголовка
        for ridx in range(header_row_idx + 1, len(df)):
            row = df.iloc[ridx]
            sys_raw = normalize_cell(row.iloc[col_sys])
            sub_raw = normalize_cell(row.iloc[col_sub])
            name_raw = normalize_cell(row.iloc[col_name])

            # Пустая строка
            if not sys_raw and not sub_raw and not name_raw:
                continue

            # Если в строке указан код системы – обновляем "текущую систему"
            if sys_raw:
                current_system = sys_raw

            if not current_system:
                continue

            if not sub_raw or not name_raw:
                continue

            rows.append((current_system, sub_raw, name_raw))

    return rows


# ---------------- Сборка Structure.xlsx ----------------

def build_items_from_rows(rows: List[Tuple[str, str, str]]) -> pd.DataFrame:
    """
    На вход: список (код_системы, код_подсистемы, имя).
    На выход: DataFrame с колонками шаблона Pragmatica.
    """
    items: Dict[str, Dict[str, str]] = {}

    for sys_code, sub_code, name in rows:
        sys_code = str(sys_code).strip()
        sub_code = str(sub_code).strip()
        name = str(name).strip()

        if not sys_code or not sub_code or not name:
            continue

        item_id = f"{sys_code}.{sub_code}"
        parent_id = "" if sub_code == "00" else f"{sys_code}.00"

        if item_id not in items:
            items[item_id] = {
                "Item_ID": item_id,
                "Parent_ID": parent_id,
                "Name": name,
                "Description": "",
                "Quantity": "1",
                "UOM": "Н",
            }
        else:
            existing = items[item_id]
            if name and name != existing["Name"]:
                print(
                    f"Предупреждение: повторяющийся код структуры '{item_id}' "
                    f"('{existing['Name']}' / '{name}') — используется только первое вхождение."
                )
                alt = name
                if alt not in existing["Description"]:
                    if existing["Description"]:
                        existing["Description"] += "\nАльтернативное название: " + alt
                    else:
                        existing["Description"] = "Альтернативное название: " + alt

    df = pd.DataFrame(items.values())
    df = df[["Item_ID", "Parent_ID", "Name", "Description", "Quantity", "UOM"]]
    df.sort_values("Item_ID", inplace=True)
    df.reset_index(drop=True, inplace=True)
    return df


# ---------------- main ----------------

def main():
    parser = argparse.ArgumentParser(
        description="Парсер структуры изделия в шаблон Structure.xlsx (Pragmatica)."
    )
    parser.add_argument(
        "input_file",
        help="Источник структуры (.pdf, .xlsx, .xls)",
    )
    parser.add_argument(
        "--out",
        default=DEFAULT_OUT,
        help=(
            "Имя выходного Excel-файла "
            "(по умолчанию Structure.xlsx в папке 'Результаты_EXCEL')."
        ),
    )

    args = parser.parse_args()
    src = os.path.abspath(args.input_file)
    out_path = os.path.abspath(args.out)

    print("================================")
    print(" Парсер структуры изделия")
    print(f" Входной файл : {src}")
    print(f" Выходной файл: {out_path}")
    print("================================")

    if not os.path.isfile(src):
        print(f"Ошибка: файл '{src}' не найден.")
        sys.exit(1)

    ext = os.path.splitext(src)[1].lower()

    if ext == ".pdf":
        print("Источник: PDF, поиск таблицы 'Система / Подсистема / Наименование'.")
        rows = _extract_system_rows_from_pdf(src)
    elif ext in (".xlsx", ".xls"):
        print("Источник: Excel, поиск таблицы 'Система / Подсистема / Наименование'.")
        rows = _extract_system_rows_from_excel(src)
    else:
        print("Ошибка: для структуры сейчас поддерживаются только PDF и Excel.")
        sys.exit(1)

    if not rows:
        print("Не найдено ни одной строки вида 'Система / Подсистема / Наименование'.")
        sys.exit(1)

    df = build_items_from_rows(rows)

    print(f"Элементов всего : {len(df)}")
    roots = (df["Parent_ID"] == "").sum()
    print(f"Корневых узлов  : {roots}")
    print("--------------------------------")

    # Перед записью нового Structure.xlsx старый уезжает в Архив
    archive_old(out_path)

    out_dir = os.path.dirname(out_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    try:
        df.to_excel(out_path, index=False, sheet_name="Структура")
    except Exception as e:
        print(f"Ошибка при сохранении Excel: {e}")
        sys.exit(1)

    print("Статус          : УСПЕХ")
    print(f"Выходной Excel  : {out_path}")
    print("================================")


if __name__ == "__main__":
    main()