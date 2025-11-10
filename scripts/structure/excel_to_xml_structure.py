import sys
import os
import pandas as pd
import xml.etree.ElementTree as ET
from datetime import datetime

# ---------------- БАЗОВЫЕ ПУТИ ----------------

# Папка, где лежит этот скрипт (Служебные файлы)
SCRIPT_DIR = os.path.abspath(os.path.dirname(__file__))

# Общая папка "Структура изделия" на уровень выше
BASE_DIR = os.path.abspath(os.path.join(SCRIPT_DIR, ".."))

# Папки с результатами
EXCEL_DIR = os.path.join(BASE_DIR, "Результаты_EXCEL")
XML_DIR = os.path.join(BASE_DIR, "Результаты_XML")

# Имена файлов
INPUT_FILE = os.path.join(EXCEL_DIR, "Structure.xlsx")
OUTPUT_FILE = os.path.join(XML_DIR, "structure_output.xml")
SHEET_NAME = "Структура"


# ---------------- СЛУЖЕБКА ----------------

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
        # Просто пишем в консоль, чтобы не ронять весь процесс
        print(f"Не удалось переместить старый файл '{path}': {e}")


def normalize_cell(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


# ---------------- ЛОГИКА СТРУКТУРЫ ----------------

def validate_and_build_items(df):
    """
    Строим словарь элементов структуры и порядок обхода.

    Поведение:
      - Обязательные колонки валидируем жёстко.
      - Пустой Item_ID -> ошибка.
      - Дубликаты Item_ID:
          * если данные полностью совпадают – просто предупреждение, берём первую строку;
          * если данные отличаются – объединяем, стараясь дополнить пустые поля,
            и выводим предупреждение.
      - Parent_ID, который ни на кого не указывает:
          * вместо фатальной ошибки делаем элемент корневым и пишем предупреждение.
    """
    required_columns = [
        "Item_ID",
        "Parent_ID",
        "Name",
        "Description",
        "Quantity",
        "UOM",
    ]

    for col in required_columns:
        if col not in df.columns:
            raise ValueError(
                f"Ошибка: в листе '{SHEET_NAME}' не найден столбец '{col}'"
            )

    items = {}
    order = []

    for idx, row in df.iterrows():
        row_num = idx + 2  # с учетом заголовка

        item_id = normalize_cell(row.get("Item_ID"))
        parent_id = normalize_cell(row.get("Parent_ID"))
        name = normalize_cell(row.get("Name"))
        description = normalize_cell(row.get("Description"))
        quantity_raw = normalize_cell(row.get("Quantity"))
        uom = normalize_cell(row.get("UOM")) or "Н"

        if not item_id:
            raise ValueError(
                f"Ошибка в строке {row_num}: пустое значение Item_ID"
            )

        if not quantity_raw:
            quantity = "1"
        else:
            quantity = quantity_raw

        data = {
            "id": item_id,
            "parent_id": parent_id,
            "name": name,
            "description": description,
            "quantity": quantity,
            "uom": uom,
        }

        if item_id not in items:
            # первая встреча этого элемента
            items[item_id] = data
            order.append(item_id)
        else:
            # Дубликат: пытаемся аккуратно склеить
            existing = items[item_id]

            same_all = (
                existing["parent_id"] == data["parent_id"]
                and existing["name"] == data["name"]
                and existing["description"] == data["description"]
                and existing["quantity"] == data["quantity"]
                and existing["uom"] == data["uom"]
            )

            if same_all:
                print(
                    f"Предупреждение: дубликат Item_ID '{item_id}' "
                    f"(строка {row_num}) с теми же данными — использована первая строка."
                )
            else:
                print(
                    f"Предупреждение: дубликат Item_ID '{item_id}' "
                    f"(строка {row_num}) с отличающимися данными — выполнено объединение."
                )
                # parent_id: оставляем первый встретившийся, чтобы не шашлык из дерева
                # name: если старое пустое, берём новое
                if not existing["name"] and data["name"]:
                    existing["name"] = data["name"]
                # description: дополняем, если раньше было пусто
                if not existing["description"] and data["description"]:
                    existing["description"] = data["description"]
                # quantity: если было "1", а новое что-то осмысленное, берём новое
                if existing["quantity"] == "1" and data["quantity"] != "1":
                    existing["quantity"] = data["quantity"]
                # uom: если была заглушка "Н", а новое что-то другое — берём новое
                if existing["uom"] == "Н" and data["uom"] != "Н":
                    existing["uom"] = data["uom"]
                # parent_id, даже если отличается, не меняем, чтобы не порвать структуру

    # проверка, что все Parent_ID существуют
    for item in items.values():
        parent_id = item["parent_id"]
        if parent_id and parent_id not in items:
            # вместо падения делаем элемент корневым
            print(
                f"Предупреждение: для элемента '{item['id']}' "
                f"указан несуществующий Parent_ID '{parent_id}'. "
                f"Элемент будет считаться корневым."
            )
            item["parent_id"] = ""

    return items, order


def build_children_map(items, order):
    children = {item_id: [] for item_id in items.keys()}
    roots = []

    for item_id in order:
        item = items[item_id]
        parent_id = item["parent_id"]
        if parent_id:
            children[parent_id].append(item_id)
        else:
            roots.append(item_id)

    return roots, children


def add_cube_xml(parent_element, item, children_map, items, is_root=False):
    final_item = "1" if is_root else "0"

    cube_attrs = {
        "is_MSI": "0",
        "final_item": final_item,
        "description": item["description"],
        "name": item["name"],
        "id": item["id"],
        "uom": item["uom"],
    }

    cube_el = ET.SubElement(parent_element, "Cube", cube_attrs)

    for child_id in children_map[item["id"]]:
        child = items[child_id]
        quantity = child["quantity"] or "1"

        link_el = ET.SubElement(cube_el, "CubeLink", {"quantity": quantity})
        add_cube_xml(link_el, child, children_map, items, is_root=False)


# ---------------- MAIN ----------------

def main():
    print("================================")
    print(" Конвертация структуры изделия в XML")
    print(f" Входной файл : {INPUT_FILE}")
    print(f" Лист         : {SHEET_NAME}")
    print("================================")

    # гарантируем, что папка для XML существует
    os.makedirs(XML_DIR, exist_ok=True)

    if not os.path.exists(INPUT_FILE):
        print(f"Ошибка: файл '{INPUT_FILE}' не найден.")
        print("Положи Structure.xlsx в папку 'Результаты_EXCEL' рядом с этой папкой.")
        sys.exit(1)

    try:
        df = pd.read_excel(INPUT_FILE, sheet_name=SHEET_NAME, engine="openpyxl")
    except Exception as e:
        print(f"Ошибка при чтении файла {INPUT_FILE}: {e}")
        sys.exit(1)

    try:
        items, order = validate_and_build_items(df)
        roots, children_map = build_children_map(items, order)
    except ValueError as ve:
        print("Ошибка при обработке данных Excel:")
        print(ve)
        sys.exit(1)

    if not roots:
        print("Ошибка: не найден ни один корневой элемент (строка с пустым Parent_ID).")
        sys.exit(1)

    dataset = ET.Element("Dataset", {"GUID": "urn:placeholder"})

    for root_id in roots:
        root_item = items[root_id]
        add_cube_xml(dataset, root_item, children_map, items, is_root=True)

    tree = ET.ElementTree(dataset)

    try:
        ET.indent(tree, space="    ", level=0)
    except Exception:
        # старые версии Python без ET.indent — просто живём дальше
        pass

    # Перед записью нового XML отправим старый в Архив
    archive_old(OUTPUT_FILE)

    try:
        tree.write(OUTPUT_FILE, encoding="utf-8", xml_declaration=True)
    except Exception as e:
        print(f"Ошибка при сохранении XML: {e}")
        sys.exit(1)

    print("Элементов всего :", len(items))
    print("Корневых узлов  :", len(roots))
    print("--------------------------------")
    print("Статус          : УСПЕХ")
    print(f"Выходной файл   : {OUTPUT_FILE}")
    print("================================")


if __name__ == "__main__":
    main()