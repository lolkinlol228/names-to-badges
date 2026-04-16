import tkinter as tk
from tkinter import filedialog, messagebox, ttk, colorchooser
from PIL import Image, ImageDraw, ImageFont, ImageTk
import os
import json
import math

class BadgeGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("Генератор бейджиков для печати на А4")
        self.root.geometry("1400x900")
        self.settings_file = "badge_settings.json"

        # --- Переменные для шаблона и областей ---
        self.template_image = None
        self.template_path = ""
        self.name_area = None          # Область для ФИО (x1, y1, x2, y2)
        self.position_area = None      # Область для должности (x1, y1, x2, y2)
        self.student_font_path = None
        self.student_text_color = "#2b3182"

        # Студенты: [(name, position), ...]
        self.students = []

        # Печать
        self.a4_width = 2480  # 300 DPI
        self.a4_height = 3508
        self.badge_width = 400
        self.badge_height = 300
        self.margin_x = 100
        self.margin_y = 100
        self.spacing_x = 50
        self.spacing_y = 50
        self.cut_margin = 20

        self.available_fonts = [
            "arial.ttf", "Arial.ttf", "calibri.ttf", "Calibri.ttf",
            "times.ttf", "Times.ttf", "comic.ttf", "Comic.ttf",
            "/System/Library/Fonts/Arial.ttf",  # macOS
            "/System/Library/Fonts/Times.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",  # Linux
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        ]

        self.load_settings()
        self.setup_ui()

    def load_settings(self):
        """Загружает сохраненные настройки из JSON файла."""
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)

                # Загружаем настройки печати
                self.badge_width_default = settings.get('badge_width', '93')
                self.badge_height_default = settings.get('badge_height', '50')
                self.margin_x_default = settings.get('margin_x', '1')
                self.margin_y_default = settings.get('margin_y', '1')
                self.spacing_x_default = settings.get('spacing_x', '1')
                self.spacing_y_default = settings.get('spacing_y', '1')
                self.cut_margin_default = settings.get('cut_margin', '3')
                self.student_text_color = settings.get('student_color', '#2b3182')
                self.student_font_path = settings.get('student_font_path', None)

                # Загружаем области текста
                if 'name_area' in settings:
                    self.name_area = tuple(settings['name_area'])
                if 'position_area' in settings:
                    self.position_area = tuple(settings['position_area'])
            else:
                # Настройки по умолчанию
                self.badge_width_default = '93'
                self.badge_height_default = '50'
                self.margin_x_default = '1'
                self.margin_y_default = '1'
                self.spacing_x_default = '1'
                self.spacing_y_default = '1'
                self.cut_margin_default = '3'
                self.student_font_path = r"C:\Users\user\Desktop\Новая папка (2)\шрифт\IntroDemoCond-BlackCAPS.otf"
        except Exception as e:
            print(f"Ошибка загрузки настроек: {e}")
            # Устанавливаем значения по умолчанию
            self.badge_width_default = '93'
            self.badge_height_default = '50'
            self.margin_x_default = '1'
            self.margin_y_default = '1'
            self.spacing_x_default = '1'
            self.spacing_y_default = '1'
            self.cut_margin_default = '3'
            self.student_font_path = r"C:\Users\user\Desktop\Новая папка (2)\шрифт\IntroDemoCond-BlackCAPS.otf"

    def save_settings(self):
        """Сохраняет текущие настройки в JSON файл."""
        try:
            settings = {
                'badge_width': self.badge_width_var.get() if hasattr(self, 'badge_width_var') else self.badge_width_default,
                'badge_height': self.badge_height_var.get() if hasattr(self, 'badge_height_var') else self.badge_height_default,
                'margin_x': self.margin_x_var.get() if hasattr(self, 'margin_x_var') else self.margin_x_default,
                'margin_y': self.margin_y_var.get() if hasattr(self, 'margin_y_var') else self.margin_y_default,
                'spacing_x': self.spacing_x_var.get() if hasattr(self, 'spacing_x_var') else self.spacing_x_default,
                'spacing_y': self.spacing_y_var.get() if hasattr(self, 'spacing_y_var') else self.spacing_y_default,
                'cut_margin': self.cut_margin_var.get() if hasattr(self, 'cut_margin_var') else self.cut_margin_default,
                'student_color': self.student_text_color,
                'student_font_path': self.student_font_path,
                'name_area': list(self.name_area) if self.name_area else None,
                'position_area': list(self.position_area) if self.position_area else None
            }
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Ошибка сохранения: {e}")

    def setup_ui(self):
        """Настройка пользовательского интерфейса."""
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True)

        template_frame = ttk.Frame(notebook)
        notebook.add(template_frame, text="Шаблон")

        print_frame = ttk.Frame(notebook)
        notebook.add(print_frame, text="Печать")

        preview_frame = ttk.Frame(notebook)
        notebook.add(preview_frame, text="Предпросмотр")

        self.setup_template_tab(template_frame)
        self.setup_print_tab(print_frame)
        self.setup_preview_tab(preview_frame)

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def on_closing(self):
        """Обработчик закрытия окна."""
        self.save_settings()
        self.root.destroy()

    def setup_template_tab(self, parent):
        """Настройка вкладки 'Шаблон'."""
        top_frame = ttk.Frame(parent)
        top_frame.pack(fill=tk.X, padx=5, pady=5)

        ttk.Button(top_frame, text="Загрузить шаблон", command=self.load_template).pack(pady=5)

        self.path_var = tk.StringVar()
        path_label = ttk.Label(top_frame, textvariable=self.path_var, foreground="gray", wraplength=400)
        path_label.pack(pady=2)

        text_frame = ttk.LabelFrame(top_frame, text="Настройки текста")
        text_frame.pack(fill=tk.X, pady=5)

        font_frame = ttk.Frame(text_frame)
        font_frame.pack(fill=tk.X, padx=5, pady=2)
        ttk.Label(font_frame, text="Шрифт:").pack(anchor=tk.W)
        self.font_var = tk.StringVar(value="arial.ttf")
        if self.student_font_path and os.path.exists(self.student_font_path):
            self.font_var.set(os.path.basename(self.student_font_path))

        font_combo = ttk.Combobox(font_frame, textvariable=self.font_var,
                                 values=["arial.ttf", "calibri.ttf", "times.ttf", "Загрузить свой..."],
                                 state="readonly", width=20)
        font_combo.pack(fill=tk.X, pady=2)
        font_combo.bind("<<ComboboxSelected>>", lambda e: self.on_font_change())

        self.font_info = tk.StringVar()
        if self.student_font_path:
            self.font_info.set(f"Загружен: {os.path.basename(self.student_font_path)}")
        font_info_label = ttk.Label(font_frame, textvariable=self.font_info, foreground="gray", font=("Arial", 8))
        font_info_label.pack(anchor=tk.W)

        color_frame = ttk.Frame(text_frame)
        color_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(color_frame, text="Цвет:").pack(side=tk.LEFT)
        self.color_var = tk.StringVar(value=self.student_text_color)
        color_label = tk.Label(color_frame, text="  ████  ", bg=self.color_var.get())
        color_label.pack(side=tk.LEFT, padx=(5, 0))

        def choose_color():
            current = self.color_var.get()
            color = colorchooser.askcolor(title="Цвет текста", color=current)[1]
            if color:
                self.color_var.set(color)
                self.student_text_color = color
                color_label.config(bg=color)

        ttk.Button(color_frame, text="Выбрать", command=choose_color).pack(side=tk.LEFT, padx=(5, 0))

        area_frame = ttk.LabelFrame(top_frame, text="Области текста")
        area_frame.pack(fill=tk.X, pady=5)

        self.name_area_info = tk.StringVar(value="ФИО: не выбрана")
        self.position_area_info = tk.StringVar(value="Должность: не выбрана")

        ttk.Label(area_frame, textvariable=self.name_area_info, font=("Courier", 8)).pack(anchor=tk.W)
        ttk.Label(area_frame, textvariable=self.position_area_info, font=("Courier", 8)).pack(anchor=tk.W)

        mode_frame = ttk.Frame(top_frame)
        mode_frame.pack(fill=tk.X, pady=5)
        self.area_mode = tk.StringVar(value="name")
        ttk.Radiobutton(mode_frame, text="Выделять область ФИО", variable=self.area_mode, value="name").pack(side=tk.LEFT)
        ttk.Radiobutton(mode_frame, text="Выделять область должности", variable=self.area_mode, value="position").pack(side=tk.LEFT)

        manual_frame = ttk.LabelFrame(top_frame, text="Ручной ввод координат")
        manual_frame.pack(fill=tk.X, pady=5)

        fio_frame = ttk.Frame(manual_frame)
        fio_frame.pack(fill=tk.X, padx=5, pady=2)
        ttk.Label(fio_frame, text="ФИО (X1 Y1 X2 Y2):").pack(side=tk.LEFT)
        self.nx1 = tk.StringVar(value="0"); self.ny1 = tk.StringVar(value="0")
        self.nx2 = tk.StringVar(value="300"); self.ny2 = tk.StringVar(value="80")
        [ttk.Entry(fio_frame, textvariable=v, width=6).pack(side=tk.LEFT, padx=2) for v in [self.nx1, self.ny1, self.nx2, self.ny2]]

        pos_frame = ttk.Frame(manual_frame)
        pos_frame.pack(fill=tk.X, padx=5, pady=2)
        ttk.Label(pos_frame, text="Должн (X1 Y1 X2 Y2):").pack(side=tk.LEFT)
        self.px1 = tk.StringVar(value="0"); self.py1 = tk.StringVar(value="90")
        self.px2 = tk.StringVar(value="300"); self.py2 = tk.StringVar(value="140")
        [ttk.Entry(pos_frame, textvariable=v, width=6).pack(side=tk.LEFT, padx=2) for v in [self.px1, self.py1, self.px2, self.py2]]

        def apply_manual():
            try:
                self.name_area = (int(self.nx1.get()), int(self.ny1.get()), int(self.nx2.get()), int(self.ny2.get()))
                self.position_area = (int(self.px1.get()), int(self.py1.get()), int(self.px2.get()), int(self.py2.get()))
                self.update_area_labels()
                messagebox.showinfo("Успех", "Координаты областей успешно применены!")
            except ValueError:
                messagebox.showerror("Ошибка", "Некорректные координаты")

        ttk.Button(manual_frame, text="Применить", command=apply_manual).pack(pady=5)

        test_frame = ttk.LabelFrame(top_frame, text="Тест")
        test_frame.pack(fill=tk.X, pady=5)

        self.test_name = tk.StringVar(value="Иванов И.И.")
        self.test_position = tk.StringVar(value="Преподаватель")
        ttk.Entry(test_frame, textvariable=self.test_name, width=25).pack(pady=2)
        ttk.Entry(test_frame, textvariable=self.test_position, width=25).pack(pady=2)
        ttk.Button(test_frame, text="Предпросмотр", command=self.show_preview).pack(pady=2)

        canvas_frame = ttk.Frame(parent)
        canvas_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        ttk.Label(canvas_frame, text="Выделите области текста", foreground="blue").pack()
        self.canvas = tk.Canvas(canvas_frame, bg="white", width=400, height=300)
        self.canvas.pack(fill=tk.BOTH, expand=True)

        self.canvas.bind("<Button-1>", self.start_selection)
        self.canvas.bind("<B1-Motion>", self.update_selection)
        self.canvas.bind("<ButtonRelease-1>", self.end_selection)

        control_frame = ttk.Frame(parent)
        control_frame.pack(fill=tk.X, pady=(10, 0))
        ttk.Button(control_frame, text="Загрузить ФИО", command=self.load_names).pack(side=tk.LEFT, padx=5)
        ttk.Button(control_frame, text="Загрузить должности", command=self.load_positions).pack(side=tk.LEFT, padx=5)

    def on_font_change(self):
        """Обработчик изменения шрифта."""
        if self.font_var.get() == "Загрузить свой...":
            self.load_custom_font()

    def load_custom_font(self):
        """Загрузка пользовательского шрифта."""
        file_path = filedialog.askopenfilename(title="Шрифт", filetypes=[("Шрифты", "*.ttf *.otf")])
        if file_path:
            try:
                ImageFont.truetype(file_path, 20)
                self.student_font_path = file_path
                self.font_var.set(os.path.basename(file_path))
                self.font_info.set(f"Загружен: {os.path.basename(file_path)}")
            except:
                messagebox.showerror("Ошибка", "Не удалось загрузить шрифт")
                self.font_var.set("arial.ttf")

    def update_area_labels(self):
        """Обновляет метки с информацией об областях."""
        if self.name_area:
            x1, y1, x2, y2 = self.name_area
            self.name_area_info.set(f"ФИО: X{x1}-{x2}, Y{y1}-{y2} (Ш:{x2-x1} В:{y2-y1})")
            for v, val in zip([self.nx1, self.ny1, self.nx2, self.ny2], [x1, y1, x2, y2]):
                v.set(str(val))
        if self.position_area:
            x1, y1, x2, y2 = self.position_area
            self.position_area_info.set(f"Должность: X{x1}-{x2}, Y{y1}-{y2} (Ш:{x2-x1} В:{y2-y1})")
            for v, val in zip([self.px1, self.py1, self.px2, self.py2], [x1, y1, x2, y2]):
                v.set(str(val))

    def setup_print_tab(self, parent):
        """Настройка вкладки 'Печать'."""
        main = ttk.Frame(parent); main.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        settings = ttk.LabelFrame(main, text="Печать"); settings.pack(side=tk.LEFT, fill=tk.Y, padx=(0,10))
        info = ttk.LabelFrame(main, text="Размещение"); info.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        size = ttk.LabelFrame(settings, text="Размер (мм)"); size.pack(fill=tk.X, pady=5)
        ttk.Label(size, text="Ширина:").pack(anchor=tk.W)
        self.badge_width_var = tk.StringVar(value=self.badge_width_default)
        tk.Spinbox(size, from_=20, to=200, width=8, textvariable=self.badge_width_var).pack()
        ttk.Label(size, text="Высота:").pack(anchor=tk.W)
        self.badge_height_var = tk.StringVar(value=self.badge_height_default)
        tk.Spinbox(size, from_=20, to=200, width=8, textvariable=self.badge_height_var).pack()

        space = ttk.LabelFrame(settings, text="Отступы и расстояния (мм)"); space.pack(fill=tk.X, pady=5)
        ttk.Label(space, text="Поля (X):").pack(anchor=tk.W)
        self.margin_x_var = tk.StringVar(value=self.margin_x_default)
        tk.Spinbox(space, from_=0, to=50, width=8, textvariable=self.margin_x_var).pack()
        ttk.Label(space, text="Поля (Y):").pack(anchor=tk.W)
        self.margin_y_var = tk.StringVar(value=self.margin_y_default)
        tk.Spinbox(space, from_=0, to=50, width=8, textvariable=self.margin_y_var).pack()
        ttk.Label(space, text="Расстояние X:").pack(anchor=tk.W)
        self.spacing_x_var = tk.StringVar(value=self.spacing_x_default)
        tk.Spinbox(space, from_=0, to=30, width=8, textvariable=self.spacing_x_var).pack()
        ttk.Label(space, text="Расстояние Y:").pack(anchor=tk.W)
        self.spacing_y_var = tk.StringVar(value=self.spacing_y_default)
        tk.Spinbox(space, from_=0, to=30, width=8, textvariable=self.spacing_y_var).pack()

        ttk.Label(settings, text="Запас обрезки (мм):").pack(anchor=tk.W)
        self.cut_margin_var = tk.StringVar(value=self.cut_margin_default)
        tk.Spinbox(settings, from_=0, to=10, width=8, textvariable=self.cut_margin_var).pack(pady=2)

        ttk.Button(settings, text="Рассчитать", command=self.calculate_layout).pack(fill=tk.X, pady=5)
        ttk.Button(settings, text="Создать листы", command=self.generate_a4_pages).pack(fill=tk.X, pady=5)
        ttk.Button(settings, text="Сохранить настройки", command=self.save_settings).pack(fill=tk.X, pady=5)

        self.layout_info = tk.Text(info, wrap=tk.WORD)
        scroll = ttk.Scrollbar(info, orient=tk.VERTICAL, command=self.layout_info.yview)
        self.layout_info.configure(yscrollcommand=scroll.set)
        self.layout_info.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)

    def setup_preview_tab(self, parent):
        """Настройка вкладки 'Предпросмотр'."""
        control = ttk.Frame(parent); control.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(control, text="Предпросмотр листов А4").pack(side=tk.LEFT)
        nav = ttk.Frame(control); nav.pack(side=tk.RIGHT)
        ttk.Button(nav, text="◀", command=self.prev_page).pack(side=tk.LEFT)
        self.page_info_var = tk.StringVar(value="Страница: 0 из 0")
        ttk.Label(nav, textvariable=self.page_info_var).pack(side=tk.LEFT, padx=10)
        ttk.Button(nav, text="▶", command=self.next_page).pack(side=tk.LEFT)

        container = ttk.Frame(parent); container.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        h = ttk.Scrollbar(container, orient=tk.HORIZONTAL)
        v = ttk.Scrollbar(container, orient=tk.VERTICAL)
        self.preview_canvas = tk.Canvas(container, bg="white", xscrollcommand=h.set, yscrollcommand=v.set)
        h.config(command=self.preview_canvas.xview)
        v.config(command=self.preview_canvas.yview)
        self.preview_canvas.grid(row=0, column=0, sticky="nsew")
        h.grid(row=1, column=0, sticky="ew")
        v.grid(row=0, column=1, sticky="ns")
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self.current_page = 0
        self.preview_pages = []

    def mm_to_pixels(self, mm, dpi=300):
        """Конвертирует мм в пиксели."""
        return int(mm * dpi / 25.4)

    def calculate_layout(self):
        """Рассчитывает размещение бейджиков на листе А4."""
        try:
            bw = float(self.badge_width_var.get())
            bh = float(self.badge_height_var.get())
            mx = float(self.margin_x_var.get())
            my = float(self.margin_y_var.get())
            sx = float(self.spacing_x_var.get())
            sy = float(self.spacing_y_var.get())
            cm = float(self.cut_margin_var.get())

            self.badge_width = self.mm_to_pixels(bw)
            self.badge_height = self.mm_to_pixels(bh)
            self.margin_x = self.mm_to_pixels(mx)
            self.margin_y = self.mm_to_pixels(my)
            self.spacing_x = self.mm_to_pixels(sx)
            self.spacing_y = self.mm_to_pixels(sy)
            self.cut_margin = self.mm_to_pixels(cm)

            avail_w = self.a4_width - 2 * self.margin_x
            avail_h = self.a4_height - 2 * self.margin_y

            badge_w = self.badge_width + 2 * self.cut_margin
            badge_h = self.badge_height + 2 * self.cut_margin

            cols = max(1, (avail_w + self.spacing_x) // (badge_w + self.spacing_x))
            rows = max(1, (avail_h + self.spacing_y) // (badge_h + self.spacing_y))

            self.badges_per_page = int(cols * rows)
            self.grid_cols = int(cols)
            self.grid_rows = int(rows)

            total = len(self.students)
            pages = math.ceil(total / self.badges_per_page) if total else 0

            info = f"""ПАРАМЕТРЫ РАЗМЕЩЕНИЯ:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📄 Лист А4: 210 × 297 мм
🏷️ Размер бейджика: {bw} × {bh} мм
📏 Поля: {mx} мм (X), {my} мм (Y)
📐 Расстояние: {sx} мм (X), {sy} мм (Y)
✂️ Запас для обрезки: {cm} мм
РАЗМЕЩЕНИЕ НА ЛИСТЕ:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📊 Бейджиков в ряд: {cols}
📊 Количество рядов: {rows}
🔢 Всего на листе: {self.badges_per_page}
СТАТИСТИКА:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
👥 Общее количество студентов: {total}
📄 Потребуется листов: {pages}
"""
            self.layout_info.delete(1.0, tk.END)
            self.layout_info.insert(tk.END, info)
        except ValueError:
            messagebox.showerror("Ошибка", "Некорректные значения")

    def load_names(self):
        """Загружает список ФИО из файла."""
        path = filedialog.askopenfilename(title="ФИО", filetypes=[("TXT", "*.txt")])
        if path:
            with open(path, 'r', encoding='utf-8') as f:
                self.names = [line.strip() for line in f if line.strip()]
            if hasattr(self, 'positions'):
                self.merge_data()

    def load_positions(self):
        """Загружает список должностей из файла."""
        path = filedialog.askopenfilename(title="Должности", filetypes=[("TXT", "*.txt")])
        if path:
            with open(path, 'r', encoding='utf-8') as f:
                self.positions = [line.strip() for line in f if line.strip()]
            if hasattr(self, 'names'):
                self.merge_data()

    def merge_data(self):
        """Объединяет ФИО и должности."""
        n = max(len(self.names), len(self.positions))
        while len(self.names) < n: self.names.append("")
        while len(self.positions) < n: self.positions.append("none")
        self.students = [(name, pos) for name, pos in zip(self.names, self.positions) if name]
        messagebox.showinfo("Успех", f"Загружено {len(self.students)} записей")

    def start_selection(self, event):
        """Начало выделения области."""
        self.start_x = self.canvas.canvasx(event.x)
        self.start_y = self.canvas.canvasy(event.y)
        self.selecting = True
        if hasattr(self, 'rect_id') and self.rect_id:
            self.canvas.delete(self.rect_id)
        self.rect_id = None

    def update_selection(self, event):
        """Обновление выделения области."""
        if not self.selecting: return
        if hasattr(self, 'rect_id') and self.rect_id:
            self.canvas.delete(self.rect_id)
        x, y = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
        self.rect_id = self.canvas.create_rectangle(self.start_x, self.start_y, x, y, outline="red", dash=(5,5), width=2)

    def end_selection(self, event):
        """Конец выделения области."""
        if not self.selecting: return
        self.selecting = False

        if hasattr(self, 'rect_id') and self.rect_id:
            self.canvas.delete(self.rect_id)
            self.rect_id = None

        if not self.template_image:
            messagebox.showwarning("Ошибка", "Сначала загрузите шаблон изображения")
            return

        # --- Исправленная логика пересчета координат ---
        end_x = self.canvas.canvasx(event.x)
        end_y = self.canvas.canvasy(event.y)

        # Получаем размеры Canvas
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()

        # Если размеры Canvas еще не определены, используем стандартные
        if canvas_width <= 1 or canvas_height <= 1:
            canvas_width, canvas_height = 400, 300

        # Получаем размеры изображения
        img_width = self.template_image.width
        img_height = self.template_image.height

        # Рассчитываем, как изображение масштабировалось и позиционировалось в Canvas
        img_ratio = img_width / img_height
        canvas_ratio = canvas_width / canvas_height

        if img_ratio > canvas_ratio:
            disp_width = canvas_width
            disp_height = int(canvas_width / img_ratio)
        else:
            disp_height = canvas_height
            disp_width = int(canvas_height * img_ratio)

        # Положение изображения в Canvas (центрировано)
        img_left = (canvas_width - disp_width) // 2
        img_top = (canvas_height - disp_height) // 2

        # Масштаб
        scale_x = img_width / disp_width
        scale_y = img_height / disp_height

        # Переводим координаты мыши в координаты исходного изображения
        x1 = max(0, min(img_width, int((min(self.start_x, end_x) - img_left) * scale_x)))
        y1 = max(0, min(img_height, int((min(self.start_y, end_y) - img_top) * scale_y)))
        x2 = max(0, min(img_width, int((max(self.start_x, end_x) - img_left) * scale_x)))
        y2 = max(0, min(img_height, int((max(self.start_y, end_y) - img_top) * scale_y)))

        if x2 <= x1 or y2 <= y1:
            messagebox.showwarning("Ошибка", "Выделена некорректная область")
            return

        mode = self.area_mode.get()
        if mode == "name":
            self.name_area = (x1, y1, x2, y2)
            messagebox.showinfo("Успех", f"Область ФИО установлена:\nX: {x1}-{x2}, Y: {y1}-{y2}\nРазмер: {x2-x1}×{y2-y1} пикселей")
        else:
            self.position_area = (x1, y1, x2, y2)
            messagebox.showinfo("Успех", f"Область должности установлена:\nX: {x1}-{x2}, Y: {y1}-{y2}\nРазмер: {x2-x1}×{y2-y1} пикселей")

        self.update_area_labels()

    def get_font(self, size):
        """Получает объект жирного шрифта."""
        # 1. Попробовать загруженный пользовательский шрифт (предполагаем, что он жирный или имеет жирное начертание)
        if self.student_font_path and os.path.exists(self.student_font_path):
            try:
                # Попытка загрузить с указанием жирности не всегда работает с .ttf/.otf
                # Лучше загрузить исходный файл, если он, например, 'bold' версия
                return ImageFont.truetype(self.student_font_path, size)
            except:
                pass # Если не получилось, идём дальше

        # 2. Попробовать стандартные жирные версии из списка
        bold_fonts = [
            "arialbd.ttf", "Arial Bold.ttf", "Arial_Bold.ttf", # Arial Bold
            "calibrib.ttf", "Calibri Bold.ttf", "Calibri_Bold.ttf", # Calibri Bold
            "timesbd.ttf", "Times New Roman Bold.ttf", "Times_Bold.ttf", # Times Bold
            "comicbd.ttf", "Comic Sans MS Bold.ttf", "Comic_Bold.ttf", # Comic Sans Bold
            # macOS
            "/System/Library/Fonts/Arial Bold.ttf",
            "/System/Library/Fonts/Times New Roman Bold.ttf",
            # Linux (часто это просто ttf-шрифты, но с суффиксом или в отдельной папке)
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
            "/usr/share/fonts/truetype/droid/DroidSans-Bold.ttf",
        ]

        # Попробуем найти жирную версию, соответствующую выбранному шрифту
        selected_font_name = self.font_var.get().lower()
        for bf in bold_fonts:
            if selected_font_name.replace(" ", "").replace(".", "") in bf.replace(" ", "").replace(".", "") and os.path.exists(bf):
                try:
                    return ImageFont.truetype(bf, size)
                except: pass

        # 3. Если жирная версия не найдена, пробуем обычную, но ищем её жирную разновидность
        # или используем общий подход (например, ищем в системных папках)
        # Это менее надёжно, но оставим как резерв
        for fp in self.available_fonts:
            if self.font_var.get().replace(" ", "").replace(".", "").lower() in fp.replace(" ", "").replace(".", "").lower():
                # Попробовать найти "bold" версию этого же шрифта
                base_name = os.path.splitext(os.path.basename(fp))[0]
                # Простая попытка найти версию с "bold" в названии
                possible_bold_names = [
                    base_name + "bd",
                    base_name + "bold",
                    base_name + "-Bold",
                    base_name + "_Bold",
                ]
                for pb_name in possible_bold_names:
                    # Ищем в той же папке, что и обычный шрифт
                    dir_path = os.path.dirname(fp)
                    for ext in [".ttf", ".TTF", ".otf", ".OTF"]:
                        bold_path = os.path.join(dir_path, pb_name + ext)
                        if os.path.exists(bold_path):
                            try:
                                return ImageFont.truetype(bold_path, size)
                            except: pass
                # Если не нашли рядом, пробуем стандартные
                for bf in bold_fonts:
                    if base_name.replace(" ", "").replace(".", "") in bf.replace(" ", "").replace(".", "") and os.path.exists(bf):
                        try:
                            return ImageFont.truetype(bf, size)
                        except: pass

        # 4. Если ничего не нашли, пробуем стандартные жирные
        for bf in bold_fonts:
            if os.path.exists(bf):
                try:
                    return ImageFont.truetype(bf, size)
                except: pass

        # 5. Если совсем не получилось, возвращаем стандартный (Arial Bold часто доступен)
        try:
            # Попытка загрузить Arial Bold на Windows
            return ImageFont.truetype("arialbd.ttf", size)
        except:
            # Если и это не удалось, используем загрузку по умолчанию (может быть не жирный)
            # Или можно попробовать 'DejaVuSans-Bold.ttf' если он установлен
            for fallback_bold in ["/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"]:
                if os.path.exists(fallback_bold):
                    try:
                        return ImageFont.truetype(fallback_bold, size)
                    except: pass
            # В крайнем случае
            print("Предупреждение: Не удалось загрузить жирный шрифт, используется стандартный.")
            return ImageFont.load_default()


    def draw_centered_text(self, draw, text, area, font, color):
        """Рисует центрированный текст в заданной области, масштабируя шрифт под размеры области."""
        if not text or text.lower() == "none":
            return

        x1, y1, x2, y2 = area
        area_width = x2 - x1
        area_height = y2 - y1

        if area_width <= 0 or area_height <= 0:
            return

        # Используем начальный шрифт, переданный в функцию, и его размер
        current_font = font
        font_size = font.size

        # Получаем bbox для начального шрифта
        try:
            bbox = draw.textbbox((0, 0), text, font=current_font)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]
        except:
            # Если bbox не работает, используем приблизительные оценки
            text_width = len(text) * font_size * 0.6
            text_height = font_size

        # --- Новый алгоритм масштабирования ---
        scale_factor_w = area_width / text_width if text_width > 0 else 1
        scale_factor_h = area_height / text_height if text_height > 0 else 1
        scale_factor = min(scale_factor_w, scale_factor_h)

        # Если текст не помещается (scale_factor < 1), уменьшаем шрифт
        if scale_factor < 1:
            new_font_size = max(1, int(font_size * scale_factor * 0.9)) # *0.9 для небольшого запаса
            current_font = self.get_font(new_font_size)
        # Если текст помещается и scale_factor > 1, можно попробовать увеличить шрифт
        # Но с осторожностью, чтобы не выйти за границы
        elif scale_factor > 1.1: # Порог для увеличения
            # Попробуем увеличить шрифт на 10% от текущего scale_factor
            # Это не идеальный способ, но простой
            # Лучше было бы итеративно увеличивать размер шрифта до границы
            # Для простоты оставим так
            # Начнем итеративное увеличение
            temp_font_size = font_size
            temp_font = font
            while True:
                try:
                    temp_bbox = draw.textbbox((0, 0), text, font=temp_font)
                    tw = temp_bbox[2] - temp_bbox[0]
                    th = temp_bbox[3] - temp_bbox[1]
                    if tw <= area_width * 0.95 and th <= area_height * 0.95: # 95% от площади, чтобы не вплотную
                        new_font_size = temp_font_size
                        current_font = temp_font
                        temp_font_size += 1
                        temp_font = self.get_font(temp_font_size)
                    else:
                        break
                except:
                    break

        # Пересчитываем размеры с новым шрифтом (если он изменился)
        try:
            bbox = draw.textbbox((0, 0), text, font=current_font)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]
        except:
            # Если не получилось с новым шрифтом, используем оценки
            text_width = len(text) * current_font.size * 0.6
            text_height = current_font.size

        # Центрируем текст
        text_x = x1 + (area_width - text_width) // 2
        text_y = y1 + (area_height - text_height) // 2

        # Рисуем текст
        draw.text((text_x, text_y), text, fill=color, font=current_font)


    def show_preview(self):
        """Показывает предварительный просмотр одного бейджика."""
        if not self.template_image:
            messagebox.showwarning("Ошибка", "Загрузите шаблон")
            return

        name = self.test_name.get()
        pos = self.test_position.get()

        try:
            badge = self.template_image.copy()
            draw = ImageDraw.Draw(badge)
            if self.name_area:
                self.draw_centered_text(draw, name, self.name_area, self.get_font(40), self.student_text_color)
            if self.position_area:
                self.draw_centered_text(draw, pos, self.position_area, self.get_font(30), self.student_text_color)

            self.show_preview_window(badge, "Предпросмотр")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Предпросмотр не удался:\n{e}")

    def show_preview_window(self, image, title):
        """Создает окно предварительного просмотра."""
        preview_window = tk.Toplevel(self.root)
        preview_window.title(title)
        preview_window.geometry("500x400")
        display_image = image.copy()
        if display_image.width > 400 or display_image.height > 300:
            display_image.thumbnail((400, 300), Image.Resampling.LANCZOS)
        photo = ImageTk.PhotoImage(display_image)
        canvas = tk.Canvas(preview_window, width=450, height=350)
        canvas.pack(expand=True)
        canvas.create_image(225, 175, image=photo)
        canvas.image = photo

    def load_template(self):
        """Загружает шаблон изображения."""
        path = filedialog.askopenfilename(filetypes=[("Изображения", "*.png *.jpg *.jpeg *.bmp")])
        if path:
            try:
                self.template_image = Image.open(path)
                self.path_var.set(os.path.basename(path))
                # Отображаем изображение в Canvas с центровкой и масштабированием
                self.display_image(self.canvas, self.template_image)
                messagebox.showinfo("Успех", f"Шаблон загружен: {os.path.basename(path)}\nРазмер: {self.template_image.width}×{self.template_image.height}")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось загрузить: {e}")

    def display_image(self, canvas, image):
        """Исправленная функция отображения изображения с правильным расчетом координат"""
        # Получаем размеры Canvas
        canvas_width = canvas.winfo_width()
        canvas_height = canvas.winfo_height()

        # Если canvas еще не отображается, используем стандартные размеры
        if canvas_width <= 1 or canvas_height <= 1:
            canvas_width, canvas_height = 400, 300

        # Сохраняем размеры изображения
        img_width = image.width
        img_height = image.height

        # Рассчитываем, как изображение масштабировалось и позиционировалось в Canvas
        img_ratio = img_width / img_height
        canvas_ratio = canvas_width / canvas_height

        if img_ratio > canvas_ratio:
            disp_width = canvas_width
            disp_height = int(canvas_width / img_ratio)
        else:
            disp_height = canvas_height
            disp_width = int(canvas_height * img_ratio)

        disp_image = image.resize((disp_width, disp_height), Image.Resampling.LANCZOS)
        photo = ImageTk.PhotoImage(disp_image)

        canvas.delete("all")
        canvas.create_image(canvas_width//2, canvas_height//2, image=photo)
        canvas.image = photo # Сохраняем ссылку

        # --- Сохраняем правильные координаты изображения для выделения областей ---
        canvas.img_left = (canvas_width - disp_width) // 2
        canvas.img_top = (canvas_height - disp_height) // 2
        canvas.scale_x = img_width / disp_width if disp_width > 0 else 1.0
        canvas.scale_y = img_height / disp_height if disp_height > 0 else 1.0
        canvas.image_width = img_width
        canvas.image_height = img_height


    def generate_a4_pages(self):
        """Генерирует листы А4 с бейджиками."""
        if not self.students:
            messagebox.showwarning("Ошибка", "Загрузите ФИО и должности")
            return
        if not self.template_image:
            messagebox.showwarning("Ошибка", "Загрузите шаблон")
            return
        if not hasattr(self, 'badges_per_page'):
            messagebox.showwarning("Ошибка", "Рассчитайте размещение")
            return

        out_dir = filedialog.askdirectory()
        if not out_dir: return

        total = len(self.students)
        pages_total = math.ceil(total / self.badges_per_page)

        prog = tk.Toplevel(self.root)
        prog.title("Генерация...")
        prog.geometry("500x150")
        prog.transient(self.root)
        prog.grab_set()
        pb = ttk.Progressbar(prog, maximum=pages_total)
        pb.pack(fill=tk.X, padx=20, pady=10)
        status = ttk.Label(prog, text="Начало...")
        status.pack()

        self.preview_pages = []
        created = 0

        for p in range(pages_total):
            status.config(text=f"Страница {p+1} из {pages_total}")
            page_img = Image.new('RGB', (self.a4_width, self.a4_height), 'white')
            start = p * self.badges_per_page
            page_studs = self.students[start:start+self.badges_per_page]

            for i, (name, pos) in enumerate(page_studs):
                row, col = i // self.grid_cols, i % self.grid_cols
                x = self.margin_x + col*(self.badge_width+2*self.cut_margin+self.spacing_x) + self.cut_margin
                y = self.margin_y + row*(self.badge_height+2*self.cut_margin+self.spacing_y) + self.cut_margin

                try:
                    badge_img = self.template_image.copy()
                    draw = ImageDraw.Draw(badge_img)
                    if self.name_area:
                        self.draw_centered_text(draw, name, self.name_area, self.get_font(40), self.student_text_color)
                    if self.position_area:
                        self.draw_centered_text(draw, pos, self.position_area, self.get_font(30), self.student_text_color)

                    badge_resized = badge_img.resize((self.badge_width, self.badge_height), Image.Resampling.LANCZOS)
                    page_img.paste(badge_resized, (x, y))

                    # Направляющие для обрезки
                    c = "#CCCCCC"
                    d = ImageDraw.Draw(page_img)
                    d.rectangle([x-self.cut_margin, y-self.cut_margin,
                                x+self.badge_width+self.cut_margin,
                                y+self.badge_height+self.cut_margin], outline=c, width=1)
                except Exception as e:
                    print(f"Ошибка при создании бейджика для {name}: {e}")

            fname = f"страница_{p+1:02d}.png"
            page_img.save(os.path.join(out_dir, fname), "PNG", dpi=(300,300))

            preview = page_img.copy()
            preview.thumbnail((800, 1131), Image.Resampling.LANCZOS) # Пропорционально А4
            self.preview_pages.append(preview)

            created += 1
            pb['value'] = created
            prog.update()

        prog.destroy()
        self.current_page = 0
        self.update_preview()
        messagebox.showinfo("Готово", f"Создано {created} страниц")

    def update_preview(self):
        """Обновляет предварительный просмотр страницы."""
        if not self.preview_pages:
            self.page_info_var.set("Страница: 0 из 0")
            self.preview_canvas.delete("all")
            return

        total = len(self.preview_pages)
        self.page_info_var.set(f"Страница: {self.current_page+1} из {total}")

        if 0 <= self.current_page < total:
            img = self.preview_pages[self.current_page]
            self.preview_photo = ImageTk.PhotoImage(img)
            self.preview_canvas.delete("all")
            self.preview_canvas.create_image(400, 565, image=self.preview_photo)
            self.preview_canvas.configure(scrollregion=self.preview_canvas.bbox("all"))

    def prev_page(self):
        """Предыдущая страница предварительного просмотра."""
        if self.preview_pages and self.current_page > 0:
            self.current_page -= 1
            self.update_preview()

    def next_page(self):
        """Следующая страница предварительного просмотра."""
        if self.preview_pages and self.current_page < len(self.preview_pages)-1:
            self.current_page += 1
            self.update_preview()

if __name__ == "__main__":
    root = tk.Tk()
    app = BadgeGenerator(root)
    root.mainloop()
