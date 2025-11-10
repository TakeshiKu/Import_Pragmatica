import os
from datetime import datetime

def main():
    print("=== Импорт данных в систему Pragmatica ===")
    print(f"Дата запуска: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("Проверка структуры проекта...")

    required_dirs = ["data/input", "data/generated", "docs", "tools"]
    missing = [d for d in required_dirs if not os.path.exists(d)]

    if missing:
        print("Отсутствуют директории:", ", ".join(missing))
    else:
        print("Все директории на месте. Можно загружать данные.")

if __name__ == "__main__":
    main()