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
SCRIPT_DIR = os.path.abspath(os.path.dirname(__file__))           # ...\Функции\Служебные файлы
BASE_DIR = os.path.abspath(os.path.join(SCRIPT_DIR, ".."))        # ...\Функции

# Папка с логотипами: ImportPragmatica\logo
LOGO_DIR = os.path.abspath(os.path.join(BASE_DIR, "..", "logo"))
LOGO_PNG = os.path.join(LOGO_DIR, "logo.png")
LOGO_ICO = os.path.join(LOGO_DIR, "logo.ico")

# Папки результатов
RESULT_EXCEL_DIR = os.path.join(BASE_DIR, "Результаты_EXCEL")
RESULT_XML_DIR = os.path.join(BASE_DIR, "Результаты_XML")

os.makedirs(RESULT_EXCEL_DIR, exist_ok=True)
os.makedirs(RESULT_XML_DIR, exist_ok=True)

FIELD_WIDTH = 260  # ширина всех полей слева (кроме "Файл")

_single_instance_socket = None


def ensure_single_instance():
    """Блокируем запуск второго такого же окна."""
    global _single_instance_socket
    _single_instance_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        _single_instance_socket.bind(("127.0.0.1", 54321))
    except OSError:
        messagebox.showerror(
            "Импорт функций",
            "Окно импорта уже запущено.\nИспользуйте его, а не открывайте новое.",
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
    fi = fi_var.get().strip()
    max_depth = depth_var.get().strip()
    mode = mode_var.get().strip() or "fi"

    # буква кода из переключателя
    code_letter_display = code_letter_var.get().strip()
    if code_letter_display == "Любая":
        code_letter = ""
    else:
        code_letter = code_letter_display

    code_prefix = code_prefix_var.get().strip()

    # очистка лога
    log_box.configure(state="normal")
    log_box.delete("1.0", "end")
    log_box.configure(state="disabled")

    if not file_path:
        messagebox.showerror("Ошибка", "Не указан файл.")
        return
    if not os.path.isfile(file_path):
        messagebox.showerror("Ошибка", f"Файл не найден:\n{file_path}")
        return
    if not fi:
        messagebox.showerror("Ошибка", "Не указано обозначение ФИ.")
        return
    if not max_depth:
        max_depth = "0"

    parse_script = os.path.join(SCRIPT_DIR, "parse_functions.py")
    if not os.path.isfile(parse_script):
        messagebox.showerror("Ошибка", f"Не найден parse_functions.py в {SCRIPT_DIR}")
        return

    # Excel складываем в папку Результаты_EXCEL
    out_excel = os.path.join(RESULT_EXCEL_DIR, "Functions.xlsx")

    cmd_parse = [
        "python",
        parse_script,
        file_path,
        "--fi",
        fi,
        "--out",
        out_excel,
        "--max-depth",
        max_depth,
        "--mode",
        mode,
    ]

    if code_letter:
        cmd_parse += ["--code-letter", code_letter]
    if code_prefix:
        cmd_parse += ["--code-prefix", code_prefix]

    log_box_insert("Шаг 1. Анализ файла и создание Functions.xlsx\n", bold=True)
    log_box_insert(f"Файл: {file_path}\n")
    log_box_insert(f"ФИ: {fi}\n")
    log_box_insert(f"Требуемый уровень вложенности: {max_depth}\n")
    log_box_insert(
        f"Тип функций: {'ФИ изделия' if mode == 'fi' else 'Функции систем (ФС)'}\n"
    )
    if code_letter:
        log_box_insert(f"Буква кода: {code_letter}\n")
    if code_prefix:
        log_box_insert(f"Начало числовой части: {code_prefix}\n")
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

    # если парсер отвалился или не нашёл функций – не продолжаем
    if result.returncode != 0:
        messagebox.showerror(
            "Ошибка",
            "parse_functions.py завершился с ошибкой или не нашёл ни одной функции.\nСм. лог.",
        )
        return

    if not os.path.isfile(out_excel):
        messagebox.showerror(
            "Ошибка",
            "Файл Functions.xlsx не создан. См. лог выше.",
        )
        return

    # ---------- Шаг 2 ----------
    xml_script = os.path.join(SCRIPT_DIR, "excel_to_xml_functions.py")
    if not os.path.isfile(xml_script):
        messagebox.showwarning(
            "Предупреждение",
            "Не найден excel_to_xml_functions.py.\nСоздан только Functions.xlsx.",
        )
        return

    cmd_xml = ["python", xml_script]

    log_box_insert("\nШаг 2. Конвертация Functions.xlsx в XML\n", bold=True)
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

    # XML ищем в папке Результаты_XML
    xml_path = os.path.join(RESULT_XML_DIR, "functions_output.xml")
    if os.path.isfile(xml_path):
        messagebox.showinfo("Готово", f"Создан файл:\n{xml_path}")
    else:
        messagebox.showwarning(
            "Предупреждение",
            "XML не найден. Проверьте лог.",
        )


def choose_file():
    path = filedialog.askopenfilename(
        title="Выбор файла с функциями",
        filetypes=[
            ("Все поддерживаемые", "*.docx *.pdf *.xlsx *.xls *.txt"),
            ("Word документы", "*.docx"),
            ("PDF файлы", "*.pdf"),
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

# блокируем второй экземпляр
ensure_single_instance()

# Цвета
APP_BG = "#135263"        # ВНЕШНИЙ фон окна
PANEL_BG = "#0F3C47"      # ВНУТРЕННИЕ панели (где текст/поля/лог)
LOG_BG = "#FFFFFF"        # фон текстбокса (белый)

BUTTON_BROWSE = "#007874"
BUTTON_BROWSE_HOVER = "#009184"

BUTTON_MAIN = BUTTON_BROWSE
BUTTON_MAIN_HOVER = BUTTON_BROWSE_HOVER

SEGMENT_SELECTED = BUTTON_BROWSE
SEGMENT_SELECTED_HOVER = BUTTON_BROWSE_HOVER
SEGMENT_UNSELECTED = "#155260"
SEGMENT_UNSELECTED_HOVER = "#1D5F6E"

# Увеличенный шрифт
DEFAULT_FONT = ("Segoe UI", 13)
SMALL_FONT = ("Segoe UI", 12, "normal")

app = ctk.CTk()
app.title("Импорт функций в Pragmatica")
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
fi_var = ctk.StringVar()
depth_var = ctk.StringVar(value="0")
mode_var = ctk.StringVar(value="fi")
code_prefix_var = ctk.StringVar()
code_letter_var = ctk.StringVar(value="Любая")

top_frame = ctk.CTkFrame(app, corner_radius=10, fg_color=PANEL_BG)
top_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
top_frame.grid_columnconfigure(1, weight=1)
top_frame.grid_columnconfigure(2, weight=0)

SEGMENT_KWARGS = dict(
    width=FIELD_WIDTH,
    height=32,
    corner_radius=6,
    selected_color=SEGMENT_SELECTED,
    selected_hover_color=SEGMENT_SELECTED_HOVER,
    unselected_color=SEGMENT_UNSELECTED,
    unselected_hover_color=SEGMENT_UNSELECTED_HOVER,
    text_color="white",
    font=DEFAULT_FONT,
)

# логотип справа под кнопкой "Обзор..."
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
            rowspan=3,
            padx=(10, 20),
            pady=(10, 0),
            sticky="ne",
        )
    except Exception as e:
        print("Не удалось загрузить логотип (справа):", e)

ENTRY_KWARGS = dict(
    fg_color="white",
    text_color="black",
    border_color="#1E4C55",
    font=DEFAULT_FONT,
)

label_file = ctk.CTkLabel(top_frame, text="Файл:", font=DEFAULT_FONT, text_color="white")
label_file.grid(row=0, column=0, padx=10, pady=(8, 3), sticky="w")

entry_file = ctk.CTkEntry(top_frame, textvariable=file_var, **ENTRY_KWARGS)
entry_file.grid(row=0, column=1, padx=5, pady=(8, 3), sticky="we")

btn_browse = ctk.CTkButton(
    top_frame,
    text="Обзор...",
    width=110,
    command=choose_file,
    fg_color=BUTTON_BROWSE,
    hover_color=BUTTON_BROWSE_HOVER,
    font=DEFAULT_FONT,
)
btn_browse.grid(row=0, column=2, padx=10, pady=(8, 3), sticky="e")

label_fi = ctk.CTkLabel(top_frame, text="ФИ:", font=DEFAULT_FONT, text_color="white")
label_fi.grid(row=1, column=0, padx=10, pady=3, sticky="w")

entry_fi = ctk.CTkEntry(
    top_frame,
    width=FIELD_WIDTH,
    textvariable=fi_var,
    **ENTRY_KWARGS,
)
entry_fi.grid(row=1, column=1, padx=5, pady=3, sticky="w")

label_fi_hint = ctk.CTkLabel(
    top_frame,
    text="Как в Проекте АН",
    font=SMALL_FONT,
    text_color="white",
)
label_fi_hint.grid(row=1, column=1, padx=(FIELD_WIDTH + 20, 5), pady=3, sticky="w")

label_depth = ctk.CTkLabel(
    top_frame,
    text="Требуемый уровень вложенности:",
    font=DEFAULT_FONT,
    text_color="white",
)
label_depth.grid(row=2, column=0, padx=10, pady=3, sticky="w")


def on_depth_change(value: str):
    depth_var.set(value)


depth_segment = ctk.CTkSegmentedButton(
    top_frame,
    values=["0", "1", "2", "3", "4", "5"],
    command=on_depth_change,
    **SEGMENT_KWARGS,
)
depth_segment.grid(row=2, column=1, padx=5, pady=3, sticky="w")
depth_segment.set(depth_var.get())

label_hint = ctk.CTkLabel(
    top_frame,
    text="0 = вся глубина",
    font=SMALL_FONT,
    text_color="white",
)
label_hint.grid(row=2, column=1, padx=(FIELD_WIDTH + 20, 5), pady=3, sticky="w")

label_mode = ctk.CTkLabel(
    top_frame,
    text="Тип функций:",
    font=DEFAULT_FONT,
    text_color="white",
)
label_mode.grid(row=3, column=0, padx=10, pady=3, sticky="w")


def on_mode_change(value: str):
    if "ФИ" in value:
        mode_var.set("fi")
    else:
        mode_var.set("fs")


mode_segment = ctk.CTkSegmentedButton(
    top_frame,
    values=["Функции ФИ", "Функции систем"],
    command=on_mode_change,
    **SEGMENT_KWARGS,
)
mode_segment.grid(row=3, column=1, padx=5, pady=3, sticky="w")
mode_segment.set("Функции ФИ")

label_letter = ctk.CTkLabel(
    top_frame,
    text="Буква кода:",
    font=DEFAULT_FONT,
    text_color="white",
)
label_letter.grid(row=4, column=0, padx=10, pady=3, sticky="w")


def on_letter_change(value: str):
    code_letter_var.set(value)


letter_segment = ctk.CTkSegmentedButton(
    top_frame,
    values=["Любая", "F", "Ф"],
    command=on_letter_change,
    **SEGMENT_KWARGS,
)
letter_segment.grid(row=4, column=1, padx=5, pady=3, sticky="w")
letter_segment.set("Любая")

label_prefix = ctk.CTkLabel(
    top_frame,
    text="Числовая часть начинается с:",
    font=DEFAULT_FONT,
    text_color="white",
)
label_prefix.grid(row=5, column=0, padx=10, pady=3, sticky="w")

entry_prefix = ctk.CTkEntry(
    top_frame,
    width=FIELD_WIDTH,
    textvariable=code_prefix_var,
    **ENTRY_KWARGS,
)
entry_prefix.grid(row=5, column=1, padx=5, pady=3, sticky="w")

btn_run = ctk.CTkButton(
    top_frame,
    text="▶ Запустить импорт",
    height=38,
    command=run_import,
    fg_color=BUTTON_MAIN,
    hover_color=BUTTON_MAIN_HOVER,
    font=("Segoe UI", 13, "bold"),
)
btn_run.grid(row=6, column=0, columnspan=3, padx=10, pady=(10, 8), sticky="we")

# --- лог ---

log_frame = ctk.CTkFrame(app, corner_radius=10, fg_color=PANEL_BG)
log_frame.grid(row=2, column=0, padx=10, pady=(0, 10), sticky="nsew")

label_log = ctk.CTkLabel(
    log_frame,
    text="Результат импорта",
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