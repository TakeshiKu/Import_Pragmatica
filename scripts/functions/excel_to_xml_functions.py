import sys
import os
import pandas as pd
import xml.etree.ElementTree as ET
from datetime import datetime

# Папка, где лежит этот скрипт (Служебные файлы)
SCRIPT_DIR = os.path.abspath(os.path.dirname(__file__))

# Общая папка "Функции" на уровень выше
BASE_DIR = os.path.abspath(os.path.join(SCRIPT_DIR, ".."))

# Папки с результатами
EXCEL_DIR = os.path.join(BASE_DIR, "Результаты_EXCEL")
XML_DIR = os.path.join(BASE_DIR, "Результаты_XML")

# Имена файлов
INPUT_FILE = os.path.join(EXCEL_DIR, "Functions.xlsx")
OUTPUT_FILE = os.path.join(XML_DIR, "functions_output.xml")
SHEET_NAME = "Функции"

# Если хочешь жёстко задать код ФИ – раскомментируй и укажи строку:
# TARGET_FI = "ТЕСТ"
TARGET_FI = None

def archive_old(path: str):
    """
    Если файл path существует, переносит его в подкаталог 'Архив'
    с добавлением штампа даты/времени к имени.
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
        # В GUI это уйдёт в лог консоли, не убьёт процесс
        print(f"Не удалось переместить старый файл '{path}': {e}")
        
def normalize_cell(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def validate_and_build_functions(df):
    required_columns = [
        "FI_Обозначение",
        "Func_LCN",
        "Parent_LCN",
        "Name",
        "Description",
    ]

    for col in required_columns:
        if col not in df.columns:
            raise ValueError(
                f"Ошибка: в листе '{SHEET_NAME}' не найден столбец '{col}'"
            )

    functions = {}
    order = []

    for idx, row in df.iterrows():
        row_num = idx + 2  # с учётом заголовка

        fi = normalize_cell(row.get("FI_Обозначение"))
        lcn = normalize_cell(row.get("Func_LCN"))
        parent_lcn = normalize_cell(row.get("Parent_LCN"))
        name = normalize_cell(row.get("Name"))
        description = normalize_cell(row.get("Description"))

        if not lcn:
            raise ValueError(f"Ошибка в строке {row_num}: пустое значение Func_LCN")
        if not name:
            raise ValueError(f"Ошибка в строке {row_num}: пустое значение Name")

        if lcn in functions:
            raise ValueError(
                f"Ошибка: дубликат Func_LCN '{lcn}' (строка {row_num})"
            )

        functions[lcn] = {
            "fi": fi,
            "lcn": lcn,
            "parent_lcn": parent_lcn,
            "name": name,
            "description": description,
        }
        order.append(lcn)

    # проверка, что все Parent_LCN существуют
    for func in functions.values():
        parent_lcn = func["parent_lcn"]
        if parent_lcn and parent_lcn not in functions:
            raise ValueError(
                f"Ошибка: для функции '{func['lcn']}' указан несуществующий "
                f"Parent_LCN '{parent_lcn}'"
            )

    return functions, order


def build_children_map(functions, order):
    children = {lcn: [] for lcn in functions.keys()}
    roots = []

    for lcn in order:
        func = functions[lcn]
        parent_lcn = func["parent_lcn"]
        if parent_lcn:
            children[parent_lcn].append(lcn)
        else:
            roots.append(lcn)

    return roots, children


def get_parentfi_for_root(func):
    if TARGET_FI:
        return TARGET_FI
    return func["fi"]


def add_function_xml(parent_element, func, children_map, functions):
    # корневая или нет
    if func["parent_lcn"]:
        parent = func["parent_lcn"]
        parentfi = ""
    else:
        parent = ""
        parentfi = get_parentfi_for_root(func)

    attrs = {
        "guid": "",
        "description": func["description"],
        "parent": parent,
        "parentfi": parentfi,
        "lcn": func["lcn"],
        "name": func["name"],
    }

    func_el = ET.SubElement(parent_element, "Function", attrs)

    for child_lcn in children_map[func["lcn"]]:
        child_func = functions[child_lcn]
        add_function_xml(func_el, child_func, children_map, functions)


def main():
    print("================================")
    print(" Конвертация функций в XML")
    print(f" Входной файл : {INPUT_FILE}")
    print(f" Лист         : {SHEET_NAME}")
    print("================================")

    # гарантируем, что папка для XML существует
    os.makedirs(XML_DIR, exist_ok=True)

    # ВАЖНО: проверяем, что Excel действительно есть в Результаты_EXCEL
    if not os.path.isfile(INPUT_FILE):
        print("Ошибка: входной Excel не найден.")
        print(f"Ожидаемый файл: {INPUT_FILE}")
        print("Запустите импорт через GUI, чтобы сначала создать Functions.xlsx.")
        sys.exit(1)

    try:
        df = pd.read_excel(INPUT_FILE, sheet_name=SHEET_NAME, engine="openpyxl")
    except Exception as e:
        print(f"Ошибка при чтении файла {INPUT_FILE}: {e}")
        sys.exit(1)

    try:
        functions, order = validate_and_build_functions(df)
        roots, children_map = build_children_map(functions, order)
    except ValueError as ve:
        print("Ошибка при обработке данных Excel:")
        print(ve)
        sys.exit(1)

    dataset = ET.Element("Dataset", {"GUID": "urn:placeholder"})

    for root_lcn in roots:
        func = functions[root_lcn]
        add_function_xml(dataset, func, children_map, functions)

    tree = ET.ElementTree(dataset)

    try:
        ET.indent(tree, space="    ", level=0)
    except Exception:
        # старые версии Python без ET.indent
        pass

    # Перед записью нового XML отправим старый в Архив
    archive_old(OUTPUT_FILE)
        
    try:
        tree.write(OUTPUT_FILE, encoding="utf-8", xml_declaration=True)
    except Exception as e:
        print(f"Ошибка при сохранении XML: {e}")
        sys.exit(1)

    print("Строк обработано :", len(functions))
    print("Корневых функций :", len(roots))
    print("--------------------------------")
    print("Статус           : УСПЕХ")
    print(f"Выходной файл    : {OUTPUT_FILE}")
    print("================================")


if __name__ == "__main__":
    main()