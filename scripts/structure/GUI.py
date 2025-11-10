# GUI.py  (для "Структура изделия")
import os
import socket
import subprocess
import tkinter as tk
import tkinter.filedialog as filedialog
import tkinter.messagebox as messagebox

import customtkinter as ctk

# Попробуем подключить Pillow для логотипа
try:
    from PIL import Image
except ImportError:
    Image = None

# ===== Пути =====
SCRIPT_DIR = os.path.abspath(os.path.dirname(__file__))      # ...\Структура изделия\Служебные файлы
BASE_DIR = os.path.abspath(os.path.join(SCRIPT_DIR, ".."))   # ...\Структура изделия

# Папка с логотипами: ImportPragmatica\logo
LOGO_DIR = os.path.abspath(os.path.join(BASE_DIR, "..", "logo"))
LOGO_PNG = os.path.join(LOGO_DIR, "logo.png")
LOGO_ICO = os.path.join(LOGO_DIR, "logo.ico")

# Папки результатов
RESULT_EXCEL_DIR = os.path.join(BASE_DIR, "Результаты_EXCEL")
RESULT_XML_DIR = os.path.join(BASE_DIR, "Результаты_XML")

os.makedirs(RESULT_EXCEL_DIR, exist_ok=True)
os.makedirs(RESULT_XML_DIR, exist_ok=True)

_single_instance_socket = None


def ensure_single_instance():
    """Блокируем запуск второго такого же окна."""
    global _single_instance_socket
    _single_instance_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        _single_instance_socket.bind(("127.0.0.1", 54322))
    except OSError:
        messagebox.showerror(
            "Импорт структуры изделия",
            "Окно импорта структуры уже запущено.\nИспользуйте его, а не открывайте новое.",
        )
        raise SystemExit


def log_box_insert(text: str, bold: bool = False):
    log_box.configure(state="normal")
    if bold:
        log_box.insert("end", text, ("bold",))
    else:
        log_box.insert("end", text)
    log_box.see("end")
    log_box.configure(state="disabled")


def run_import():
    file_path = file_var.get().strip()

    # очистка лога
    log_box.configure(state="normal")
    log_box.delete("1.0", "end")
    log_box.configure(state="disabled")

    if not file_path:
        messagebox.showerror("Ошибка", "Не указан файл структуры.")
        return
    if not os.path.isfile(file_path):
        messagebox.showerror("Ошибка", f"Файл не найден:\n{file_path}")
        return

    # Скрипты
    parse_script = os.path.join(SCRIPT_DIR, "parse_structure.py")
    xml_script = os.path.join(SCRIPT_DIR, "excel_to_xml_structure.py")

    if not os.path.isfile(parse_script):
        messagebox.showerror(
            "Ошибка",
            f"Не найден parse_structure.py в {SCRIPT_DIR}",
        )
        return
    if not os.path.isfile(xml_script):
        messagebox.showerror(
            "Ошибка",
            f"Не найден excel_to_xml_structure.py в {SCRIPT_DIR}",
        )
        return

    out_excel = os.path.join(RESULT_EXCEL_DIR, "Structure.xlsx")

    # ---------- Шаг 1: парсинг исходного файла в Structure.xlsx ----------
    cmd_parse = [
        "python",
        parse_script,
        file_path,
        "--out",
        out_excel,
    ]

    log_box_insert("Шаг 1. Анализ файла и создание Structure.xlsx\n", bold=True)
    log_box_insert(f"Файл структуры: {file_path}\n")
    log_box_insert(f"Выходной Excel: {out_excel}\n")
    log_box_insert("\n> " + " ".join(cmd_parse) + "\n\n")

    btn_run.configure(state="disabled")
    try:
        result = subprocess.run(
            cmd_parse,
            capture_output=True,
            text=True,
            cwd=SCRIPT_DIR,
        )
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось запустить Python:\n{e}")
        btn_run.configure(state="normal")
        return
    finally:
        btn_run.configure(state="normal")

    if result.stdout:
        log_box_insert(result.stdout + "\n")
    if result.stderr:
        log_box_insert("STDERR:\n" + result.stderr + "\n")

    if result.returncode != 0:
        messagebox.showerror(
            "Ошибка",
            "parse_structure.py завершился с ошибкой.\nСм. лог выше.",
        )
        return

    if not os.path.isfile(out_excel):
        messagebox.showerror(
            "Ошибка",
            "Файл Structure.xlsx не создан. См. лог выше.",
        )
        return

    # ---------- Шаг 2: конвертация Structure.xlsx в XML ----------
    cmd_xml = ["python", xml_script]

    log_box_insert("\nШаг 2. Конвертация Structure.xlsx в XML\n", bold=True)
    log_box_insert("> " + " ".join(cmd_xml) + "\n\n")

    result2 = subprocess.run(
        cmd_xml,
        capture_output=True,
        text=True,
        cwd=SCRIPT_DIR,
    )

    if result2.stdout:
        log_box_insert(result2.stdout + "\n")
    if result2.stderr:
        log_box_insert("STDERR:\n" + result2.stderr + "\n")

    xml_path = os.path.join(RESULT_XML_DIR, "structure_output.xml")
    if os.path.isfile(xml_path):
        messagebox.showinfo("Готово", f"Создан файл:\n{xml_path}")
    else:
        messagebox.showwarning(
            "Предупреждение",
            "XML не найден. Проверьте лог.",
        )


def choose_file():
    path = filedialog.askopenfilename(
        title="Выбор файла структуры",
        filetypes=[
            ("Все поддерживаемые", "*.pdf *.docx *.xlsx *.xls *.txt"),
            ("PDF файлы", "*.pdf"),
            ("Word файлы", "*.docx"),
            ("Excel файлы", "*.xlsx *.xls"),
            ("Текстовые файлы", "*.txt"),
            ("Все файлы", "*.*"),
        ],
    )
    if path:
        file_var.set(path)


# ================ GUI =====================

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("green")

ensure_single_instance()

# Цвета
APP_BG = "#135263"        # внешний фон
PANEL_BG = "#0F3C47"      # панели
LOG_BG = "#FFFFFF"

BUTTON_MAIN = "#007874"
BUTTON_MAIN_HOVER = "#009184"

DEFAULT_FONT = ("Segoe UI", 13)
SMALL_FONT = ("Segoe UI", 12, "normal")

app = ctk.CTk()
app.title("Импорт структуры изделия в Pragmatica")
app.geometry("950x700")
app.minsize(850, 550)
app.configure(fg_color=APP_BG)

# иконка окна
try:
    if os.path.isfile(LOGO_ICO):
        app.iconbitmap(LOGO_ICO)
    elif os.path.isfile(LOGO_PNG):
        _app_icon = tk.PhotoImage(file=LOGO_PNG)
        app.iconphoto(False, _app_icon)
except Exception as e:
    print("Не удалось установить иконку окна:", e)

app.grid_rowconfigure(2, weight=2)
app.grid_columnconfigure(0, weight=1)

file_var = ctk.StringVar()

top_frame = ctk.CTkFrame(app, corner_radius=10, fg_color=PANEL_BG)
top_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
top_frame.grid_columnconfigure(1, weight=1)
top_frame.grid_columnconfigure(2, weight=0)

# логотип под кнопкой "Обзор..."
if Image is not None and os.path.isfile(LOGO_PNG):
    try:
        logo_image = ctk.CTkImage(Image.open(LOGO_PNG), size=(90, 90))
        logo_label = ctk.CTkLabel(
            top_frame,
            image=logo_image,
            text="",
            fg_color="transparent",
        )
        logo_label.grid(
            row=1,
            column=2,
            padx=(10, 20),
            pady=(10, 0),
            sticky="n",
        )
    except Exception as e:
        print("Не удалось загрузить логотип (структура):", e)

ENTRY_KWARGS = dict(
    fg_color="white",
    text_color="black",
    border_color="#1E4C55",
    font=DEFAULT_FONT,
)

label_file = ctk.CTkLabel(
    top_frame,
    text="Файл структуры:",
    font=DEFAULT_FONT,
    text_color="white",
)
label_file.grid(row=0, column=0, padx=10, pady=(8, 3), sticky="w")

entry_file = ctk.CTkEntry(
    top_frame,
    textvariable=file_var,
    **ENTRY_KWARGS,
)
entry_file.grid(row=0, column=1, padx=5, pady=(8, 3), sticky="we")

btn_browse = ctk.CTkButton(
    top_frame,
    text="Обзор...",
    width=110,
    command=choose_file,
    fg_color=BUTTON_MAIN,
    hover_color=BUTTON_MAIN_HOVER,
    font=DEFAULT_FONT,
)
btn_browse.grid(row=0, column=2, padx=10, pady=(8, 3), sticky="e")

btn_run = ctk.CTkButton(
    top_frame,
    text="▶ Запустить импорт",
    height=38,
    command=run_import,
    fg_color=BUTTON_MAIN,
    hover_color=BUTTON_MAIN_HOVER,
    font=("Segoe UI", 13, "bold"),
)
btn_run.grid(row=2, column=0, columnspan=3, padx=10, pady=(10, 8), sticky="we")

# --- лог ---

log_frame = ctk.CTkFrame(app, corner_radius=10, fg_color=PANEL_BG)
log_frame.grid(row=2, column=0, padx=10, pady=(0, 10), sticky="nsew")

label_log = ctk.CTkLabel(
    log_frame,
    text="Результат импорта структуры",
    font=("Segoe UI", 13, "bold"),
    text_color="white",
)
label_log.pack(anchor="w", padx=10, pady=(10, 0))

log_box = ctk.CTkTextbox(
    log_frame,
    wrap="word",
    fg_color=LOG_BG,
    text_color="black",
    font=("Consolas", 12),
)
log_box.pack(fill="both", expand=True, padx=10, pady=(5, 10))
log_box.tag_config("bold")
log_box.configure(state="disabled")

app.mainloop()