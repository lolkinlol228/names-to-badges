import tkinter as tk
from tkinter import filedialog, messagebox, ttk, colorchooser
from PIL import Image, ImageDraw, ImageFont, ImageTk
import os
import json
import re
import math

class BadgeGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("Генератор бейджиков для печати на А4")
        self.root.geometry("1400x900")
        self.settings_file = "badge_settings.json"

        # --- ОСТАВЛЕН ОДИН ШАБЛОН ДЛЯ УПРОЩЕНИЯ ---
        self.template_image = None
        self.template_path = ""
        self.name_area = None          # Область для ФИО
        self.position_area = None      # Область для должности
        self.student_font_path = None
        self.student_text_color = "#2b3182"

        # Студенты: [(name, position), ...]
        self.students = []

        # Печать
        self.a4_width = 2480
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
            "/System/Library/Fonts/Arial.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        ]

        self.load_settings()
        self.setup_ui()

    def load_settings(self):
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                self.badge_width_default = settings.get('badge_width', '93')
                self.badge_height_default = settings.get('badge_height', '50')
                self.margin_x_default = settings.get('margin_x', '1')
                self.margin_y_default = settings.get('margin_y', '1')
                self.spacing_x_default = settings.get('spacing_x', '1')
                self.spacing_y_default = settings.get('spacing_y', '1')
                self.cut_margin_default = settings.get('cut_margin', '3')
                self.student_text_color = settings.get('student_color', '#2b3182')
                self.student_font_path = settings.get('student_font_path', None)
                if 'name_area' in settings:
                    self.name_area = tuple(settings['name_area'])
                if 'position_area' in settings:
                    self.position_area = tuple(settings['position_area'])
            else:
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
            self.badge_width_default = '93'
            self.badge_height_default = '50'
            self.margin_x_default = '1'
            self.margin_y_default = '1'
            self.spacing_x_default = '1'
            self.spacing_y_default = '1'
            self.cut_margin_default = '3'
            self.student_font_path = r"C:\Users\user\Desktop\Новая папка (2)\шрифт\IntroDemoCond-BlackCAPS.otf"

    def save_settings(self):
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
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True)

        template_frame = ttk.Frame(notebook)
        notebook.add(template_frame, text="Шаблон")
        self.setup_template_tab(template_frame)

        print_frame = ttk.Frame(notebook)
        notebook.add(print_frame, text="Печать")
        self.setup_print_tab(print_frame)

        preview_frame = ttk.Frame(notebook)
        notebook.add(preview_frame, text="Предпросмотр")
        self.setup_preview_tab(preview_frame)

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def on_closing(self):
        self.save_settings()
        self.root.destroy()

    def setup_template_tab(self, parent):
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
                                 state="readonly")
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
        ttk.Label(fio_frame, text="ФИО:").pack(side=tk.LEFT)
        self.nx1 = tk.StringVar(value="0"); self.ny1 = tk.StringVar(value="0")
        self.nx2 = tk.StringVar(value="300"); self.ny2 = tk.StringVar(value="80")
        [ttk.Entry(fio_frame, textvariable=v, width=6).pack(side=tk.LEFT, padx=2) for v in [self.nx1, self.ny1, self.nx2, self.ny2]]

        pos_frame = ttk.Frame(manual_frame)
        pos_frame.pack(fill=tk.X, padx=5, pady=2)
        ttk.Label(pos_frame, text="Должн:").pack(side=tk.LEFT)
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

        # Добавляем подсказки для работы с размерами
        help_frame = ttk.LabelFrame(top_frame, text="Подсказка")
        help_frame.pack(fill=tk.X, pady=5)
        ttk.Label(help_frame, text="Перетащите мышью на изображении для выделения области.\nРазмеры будут автоматически рассчитаны из координат выделения.",
                  font=("Arial", 8), justify=tk.LEFT).pack(anchor=tk.W, padx=5, pady=2)

    def on_font_change(self):
        if self.font_var.get() == "Загрузить свой...":
            self.load_custom_font()

    def load_custom_font(self):
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
        if self.name_area:
            x1, y1, x2, y2 = self.name_area
            self.name_area_info.set(f"ФИО: X{x1}-{x2}, Y{y1}-{y2} (размер: {x2-x1}×{y2-y1})")
            for v, val in zip([self.nx1, self.ny1, self.nx2, self.ny2], [x1, y1, x2, y2]):
                v.set(str(val))
        if self.position_area:
            x1, y1, x2, y2 = self.position_area
            self.position_area_info.set(f"Должность: X{x1}-{x2}, Y{y1}-{y2} (размер: {x2-x1}×{y2-y1})")
            for v, val in zip([self.px1, self.py1, self.px2, self.py2], [x1, y1, x2, y2]):
                v.set(str(val))

    def setup_print_tab(self, parent):
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
        return int(mm * dpi / 25.4)

    def calculate_layout(self):
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

            info = f"Бейджик: {bw}×{bh} мм\nПоля: {mx}×{my} мм\nРасстояние: {sx}×{sy} мм\nОбрезка: {cm} мм\nНа листе: {cols}×{rows} = {self.badges_per_page}\nСтудентов: {total}\nСтраниц: {pages}"
            self.layout_info.delete(1.0, tk.END)
            self.layout_info.insert(tk.END, info)
        except ValueError:
            messagebox.showerror("Ошибка", "Некорректные значения")

    def load_names(self):
        path = filedialog.askopenfilename(title="ФИО", filetypes=[("TXT", "*.txt")])
        if path:
            with open(path, 'r', encoding='utf-8') as f:
                self.names = [line.strip() for line in f if line.strip()]
            if hasattr(self, 'positions'):
                self.merge_data()

    def load_positions(self):
        path = filedialog.askopenfilename(title="Должности", filetypes=[("TXT", "*.txt")])
        if path:
            with open(path, 'r', encoding='utf-8') as f:
                self.positions = [line.strip() for line in f if line.strip()]
            if hasattr(self, 'names'):
                self.merge_data()

    def merge_data(self):
        n = max(len(self.names), len(self.positions))
        while len(self.names) < n: self.names.append("")
        while len(self.positions) < n: self.positions.append("none")
        self.students = [(name, pos) for name, pos in zip(self.names, self.positions) if name]
        messagebox.showinfo("Успех", f"Загружено {len(self.students)} записей")

    def start_selection(self, event):
        self.start_x = self.canvas.canvasx(event.x)
        self.start_y = self.canvas.canvasy(event.y)
        self.selecting = True
        self.rect_id = None  # Сбрасываем предыдущий прямоугольник

    def update_selection(self, event):
        if not self.selecting: return
        
        # Удаляем предыдущий прямоугольник
        if hasattr(self, 'rect_id') and self.rect_id:
            self.canvas.delete(self.rect_id)
        
        x, y = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
        self.rect_id = self.canvas.create_rectangle(self.start_x, self.start_y, x, y, outline="red", dash=(5,5), width=2)

    def end_selection(self, event):
        if not self.selecting: return
        self.selecting = False
        
        # Получаем координаты окончания выделения
        end_x = self.canvas.canvasx(event.x)
        end_y = self.canvas.canvasy(event.y)

        # Удаляем временный прямоугольник
        if hasattr(self, 'rect_id') and self.rect_id:
            self.canvas.delete(self.rect_id)
            self.rect_id = None

        # Проверяем, что изображение загружено
        if not self.template_image:
            messagebox.showwarning("Ошибка", "Сначала загрузите шаблон изображения")
            return

        # Получаем реальные координаты изображения из canvas
        img_width = getattr(self.canvas, 'image_width', 400)
        img_height = getattr(self.canvas, 'image_height', 300)
        img_left = getattr(self.canvas, 'img_left', 0)
        img_top = getattr(self.canvas, 'img_top', 0)
        
        # Рассчитываем масштаб более точно
        scale_x = getattr(self.canvas, 'scale_x', 1.0)
        scale_y = getattr(self.canvas, 'scale_y', 1.0)
        
        # Вычисляем координаты в системе координат изображения
        x1 = max(0, int((min(self.start_x, end_x) - img_left) * scale_x))
        y1 = max(0, int((min(self.start_y, end_y) - img_top) * scale_y))
        x2 = min(img_width, int((max(self.start_x, end_x) - img_left) * scale_x))
        y2 = min(img_height, int((max(self.start_y, end_y) - img_top) * scale_y))

        # Проверяем, что область корректная
        if x2 <= x1 or y2 <= y1:
            messagebox.showwarning("Ошибка", "Выделена некорректная область")
            return

        # Сохраняем область в зависимости от режима
        mode = self.area_mode.get()
        if mode == "name":
            self.name_area = (x1, y1, x2, y2)
            messagebox.showinfo("Успех", f"Область ФИО установлена:\nX: {x1}-{x2}, Y: {y1}-{y2}\nРазмер: {x2-x1}×{y2-y1} пикселей")
        else:
            self.position_area = (x1, y1, x2, y2)
            messagebox.showinfo("Успех", f"Область должности установлена:\nX: {x1}-{x2}, Y: {y1}-{y2}\nРазмер: {x2-x1}×{y2-y1} пикселей")

        self.update_area_labels()

    def get_font(self, size):
        if self.student_font_path and os.path.exists(self.student_font_path):
            return ImageFont.truetype(self.student_font_path, size)
        for fp in self.available_fonts:
            if self.font_var.get() in fp:
                try:
                    return ImageFont.truetype(fp, size)
                except: pass
        try:
            return ImageFont.truetype(self.font_var.get(), size)
        except:
            return ImageFont.load_default()

    def draw_centered_text(self, draw, text, area, font, color):
        x1, y1, x2, y2 = area
        w, h = x2 - x1, y2 - y1
        try:
            bbox = draw.textbbox((0,0), text, font=font)
            tw, th = bbox[2]-bbox[0], bbox[3]-bbox[1]
        except:
            tw, th = len(text) * font.size * 0.6, font.size
        tx, ty = x1 + (w - tw) // 2, y1 + (h - th) // 2
        draw.text((tx, ty), text, fill=color, font=font)

    def show_preview(self):
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
            if pos and pos.lower() != "none" and self.position_area:
                self.draw_centered_text(draw, pos, self.position_area, self.get_font(30), self.student_text_color)
            self.show_preview_window(badge, "Предпросмотр")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Предпросмотр не удался:\n{e}")

    def show_preview_window(self, image, title):
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

    def display_image(self, canvas, image):
        """Исправленная функция отображения изображения с правильным расчетом координат"""
        cw, ch = canvas.winfo_width(), canvas.winfo_height()
        
        # Если canvas еще не отображается, используем стандартные размеры
        if cw <= 1 or ch <= 1:
            cw, ch = 400, 300
        
        ratio = image.width / image.height
        if ratio > cw/ch:
            dw, dh = cw, int(cw / ratio)
        else:
            dh, dw = ch, int(ch * ratio)
        
        disp = image.resize((dw, dh), Image.Resampling.LANCZOS)
        photo = ImageTk.PhotoImage(disp)
        
        canvas.delete("all")
        canvas.create_image(cw//2, ch//2, image=photo)
        canvas.image = photo
        
        # Сохраняем правильные координаты изображения для выделения областей
        canvas.img_left = (cw - dw) // 2
        canvas.img_top = (ch - dh) // 2
        canvas.scale_x = image.width / dw if dw > 0 else 1.0
        canvas.scale_y = image.height / dh if dh > 0 else 1.0
        canvas.image_width = image.width
        canvas.image_height = image.height
        
        # Обновляем атрибуты для совместимости
        canvas.display_scale = dw / image.width
        canvas.display_width = dw
        canvas.display_height = dh

    def load_template(self):
        path = filedialog.askopenfilename(filetypes=[("Изображения", "*.png *.jpg *.jpeg *.bmp")])
        if path:
            try:
                self.template_image = Image.open(path)
                self.path_var.set(os.path.basename(path))
                # Обновляем размер canvas под реальные размеры окна
                self.canvas.update()
                self.display_image(self.canvas, self.template_image)
                messagebox.showinfo("Успех", f"Шаблон загружен: {os.path.basename(path)}\nРазмер: {self.template_image.width}×{self.template_image.height}")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось загрузить: {e}")

    def generate_a4_pages(self):
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
                    if pos and pos.lower() != "none" and self.position_area:
                        self.draw_centered_text(draw, pos, self.position_area, self.get_font(30), self.student_text_color)
                    badge_resized = badge_img.resize((self.badge_width, self.badge_height), Image.Resampling.LANCZOS)
                    page_img.paste(badge_resized, (x, y))
                    c = "#CCCCCC"
                    d = ImageDraw.Draw(page_img)
                    d.rectangle([x-self.cut_margin, y-self.cut_margin,
                                x+self.badge_width+self.cut_margin,
                                y+self.badge_height+self.cut_margin], outline=c, width=1)
                except Exception as e:
                    print(f"Ошибка: {e}")
            fname = f"страница_{p+1:02d}.png"
            page_img.save(os.path.join(out_dir, fname), "PNG", dpi=(300,300))
            preview = page_img.copy()
            preview.thumbnail((800, 1131), Image.Resampling.LANCZOS)
            self.preview_pages.append(preview)
            created += 1
            pb['value'] = created
            prog.update()
        prog.destroy()
        self.current_page = 0
        self.update_preview()
        messagebox.showinfo("Готово", f"Создано {created} страниц")

    def update_preview(self):
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
        if self.preview_pages and self.current_page > 0:
            self.current_page -= 1
            self.update_preview()

    def next_page(self):
        if self.preview_pages and self.current_page < len(self.preview_pages)-1:
            self.current_page += 1
            self.update_preview()

if __name__ == "__main__":
    root = tk.Tk()
    app = BadgeGenerator(root)
    root.mainloop()
