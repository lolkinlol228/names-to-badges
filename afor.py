import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk
import os
import json
import math

class EmptyBadgeGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("Генератор пустых листов А4")
        self.root.geometry("800x600")
        
        # Файл для сохранения настроек
        self.settings_file = "empty_badge_settings.json"
        
        # Переменные для шаблонов
        self.student_template = None
        self.leader_template = None
        self.student_template_path = ""
        self.leader_template_path = ""
        
        # Настройки печати (размеры в пикселях для 300 DPI)
        self.a4_width = 2480  # 210mm при 300 DPI
        self.a4_height = 3508  # 297mm при 300 DPI
        self.badge_width = 400  # пиксели
        self.badge_height = 300  # пиксели
        self.margin_x = 100  # отступы от края
        self.margin_y = 100
        self.spacing_x = 50  # расстояние между бейджиками
        self.spacing_y = 50
        self.cut_margin = 20  # запас для обрезки
        
        # Загружаем сохраненные настройки
        self.load_settings()
        
        self.setup_ui()
        
    def load_settings(self):
        """Загружает сохраненные настройки"""
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
                
                # Загружаем пути к шаблонам
                self.student_template_path = settings.get('student_template_path', '')
                self.leader_template_path = settings.get('leader_template_path', '')
                
            else:
                # Настройки по умолчанию
                self.badge_width_default = '93'
                self.badge_height_default = '50'
                self.margin_x_default = '1'
                self.margin_y_default = '1'
                self.spacing_x_default = '1'
                self.spacing_y_default = '1'
                self.cut_margin_default = '3'
                self.student_template_path = ''
                self.leader_template_path = ''
                
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
            self.student_template_path = ''
            self.leader_template_path = ''
            
    def save_settings(self):
        """Сохраняет текущие настройки"""
        try:
            settings = {
                'badge_width': self.badge_width_var.get(),
                'badge_height': self.badge_height_var.get(),
                'margin_x': self.margin_x_var.get(),
                'margin_y': self.margin_y_var.get(),
                'spacing_x': self.spacing_x_var.get(),
                'spacing_y': self.spacing_y_var.get(),
                'cut_margin': self.cut_margin_var.get(),
                'student_template_path': self.student_template_path,
                'leader_template_path': self.leader_template_path
            }
            
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
                
        except Exception as e:
            print(f"Ошибка сохранения настроек: {e}")
        
    def setup_ui(self):
        # Главное окно разделено на две части
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Левая часть - шаблоны
        templates_frame = ttk.LabelFrame(main_frame, text="Шаблоны")
        templates_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
        # Шаблон для студентов
        student_frame = ttk.LabelFrame(templates_frame, text="Шаблон для обычных студентов")
        student_frame.pack(fill=tk.X, pady=5)
        
        self.student_path_var = tk.StringVar(value=os.path.basename(self.student_template_path) if self.student_template_path else "Не выбран")
        ttk.Label(student_frame, textvariable=self.student_path_var, foreground="gray").pack(pady=2)
        
        ttk.Button(student_frame, text="Выбрать шаблон студентов", 
                  command=self.load_student_template).pack(pady=5)
        
        # Предпросмотр студенческого шаблона
        self.student_canvas = tk.Canvas(student_frame, bg="white", width=200, height=150)
        self.student_canvas.pack(pady=5)
        
        # Шаблон для лидеров
        leader_frame = ttk.LabelFrame(templates_frame, text="Шаблон для лидеров")
        leader_frame.pack(fill=tk.X, pady=5)
        
        self.leader_path_var = tk.StringVar(value=os.path.basename(self.leader_template_path) if self.leader_template_path else "Не выбран")
        ttk.Label(leader_frame, textvariable=self.leader_path_var, foreground="gray").pack(pady=2)
        
        ttk.Button(leader_frame, text="Выбрать шаблон лидеров", 
                  command=self.load_leader_template).pack(pady=5)
        
        # Предпросмотр лидерского шаблона
        self.leader_canvas = tk.Canvas(leader_frame, bg="white", width=200, height=150)
        self.leader_canvas.pack(pady=5)
        
        # Правая часть - настройки и генерация
        control_frame = ttk.LabelFrame(main_frame, text="Настройки и генерация")
        control_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=(10, 0))
        
        # Настройки печати
        print_settings_frame = ttk.LabelFrame(control_frame, text="Настройки печати")
        print_settings_frame.pack(fill=tk.X, pady=5)
        
        # Размер бейджика
        size_frame = ttk.LabelFrame(print_settings_frame, text="Размер бейджика (мм)")
        size_frame.pack(fill=tk.X, pady=5)
        
        size_row1 = ttk.Frame(size_frame)
        size_row1.pack(fill=tk.X, padx=5, pady=2)
        ttk.Label(size_row1, text="Ширина:").pack(side=tk.LEFT)
        self.badge_width_var = tk.StringVar(value=self.badge_width_default)
        tk.Spinbox(size_row1, from_=20, to=200, width=8, textvariable=self.badge_width_var).pack(side=tk.RIGHT)
        
        size_row2 = ttk.Frame(size_frame)
        size_row2.pack(fill=tk.X, padx=5, pady=2)
        ttk.Label(size_row2, text="Высота:").pack(side=tk.LEFT)
        self.badge_height_var = tk.StringVar(value=self.badge_height_default)
        tk.Spinbox(size_row2, from_=20, to=200, width=8, textvariable=self.badge_height_var).pack(side=tk.RIGHT)
        
        # Отступы
        margin_frame = ttk.LabelFrame(print_settings_frame, text="Отступы (мм)")
        margin_frame.pack(fill=tk.X, pady=5)
        
        margin_row1 = ttk.Frame(margin_frame)
        margin_row1.pack(fill=tk.X, padx=5, pady=2)
        ttk.Label(margin_row1, text="По горизонтали:").pack(side=tk.LEFT)
        self.margin_x_var = tk.StringVar(value=self.margin_x_default)
        tk.Spinbox(margin_row1, from_=0, to=50, width=8, textvariable=self.margin_x_var).pack(side=tk.RIGHT)
        
        margin_row2 = ttk.Frame(margin_frame)
        margin_row2.pack(fill=tk.X, padx=5, pady=2)
        ttk.Label(margin_row2, text="По вертикали:").pack(side=tk.LEFT)
        self.margin_y_var = tk.StringVar(value=self.margin_y_default)
        tk.Spinbox(margin_row2, from_=0, to=50, width=8, textvariable=self.margin_y_var).pack(side=tk.RIGHT)
        
        # Расстояние между бейджиками
        spacing_frame = ttk.LabelFrame(print_settings_frame, text="Расстояние между бейджиками (мм)")
        spacing_frame.pack(fill=tk.X, pady=5)
        
        spacing_row1 = ttk.Frame(spacing_frame)
        spacing_row1.pack(fill=tk.X, padx=5, pady=2)
        ttk.Label(spacing_row1, text="По горизонтали:").pack(side=tk.LEFT)
        self.spacing_x_var = tk.StringVar(value=self.spacing_x_default)
        tk.Spinbox(spacing_row1, from_=0, to=30, width=8, textvariable=self.spacing_x_var).pack(side=tk.RIGHT)
        
        spacing_row2 = ttk.Frame(spacing_frame)
        spacing_row2.pack(fill=tk.X, padx=5, pady=2)
        ttk.Label(spacing_row2, text="По вертикали:").pack(side=tk.LEFT)
        self.spacing_y_var = tk.StringVar(value=self.spacing_y_default)
        tk.Spinbox(spacing_row2, from_=0, to=30, width=8, textvariable=self.spacing_y_var).pack(side=tk.RIGHT)
        
        # Запас для обрезки
        cut_frame = ttk.LabelFrame(print_settings_frame, text="Запас для обрезки (мм)")
        cut_frame.pack(fill=tk.X, pady=5)
        
        cut_row = ttk.Frame(cut_frame)
        cut_row.pack(fill=tk.X, padx=5, pady=2)
        ttk.Label(cut_row, text="Запас:").pack(side=tk.LEFT)
        self.cut_margin_var = tk.StringVar(value=self.cut_margin_default)
        tk.Spinbox(cut_row, from_=0, to=10, width=8, textvariable=self.cut_margin_var).pack(side=tk.RIGHT)
        
        # Количество листов
        count_frame = ttk.LabelFrame(control_frame, text="Количество листов")
        count_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(count_frame, text="Сколько листов создать:").pack(pady=2)
        self.pages_count_var = tk.StringVar(value="1")
        tk.Spinbox(count_frame, from_=1, to=100, width=10, textvariable=self.pages_count_var).pack(pady=2)
        
        # Кнопки генерации
        buttons_frame = ttk.LabelFrame(control_frame, text="Генерация")
        buttons_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(buttons_frame, text="Создать пустые листы А4 (Студенты)", 
                  command=lambda: self.generate_empty_pages("student")).pack(fill=tk.X, pady=2)
        
        ttk.Button(buttons_frame, text="Создать пустые листы А4 (Лидеры)", 
                  command=lambda: self.generate_empty_pages("leader")).pack(fill=tk.X, pady=2)
        
        ttk.Button(buttons_frame, text="Рассчитать размещение", 
                  command=self.calculate_layout).pack(fill=tk.X, pady=2)
        
        ttk.Button(buttons_frame, text="Сохранить настройки", 
                  command=self.save_settings).pack(fill=tk.X, pady=2)
        
        # Информация о размещении
        info_frame = ttk.LabelFrame(control_frame, text="Информация")
        info_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.info_text = tk.Text(info_frame, wrap=tk.WORD, height=8)
        info_scroll = ttk.Scrollbar(info_frame, orient=tk.VERTICAL, command=self.info_text.yview)
        self.info_text.configure(yscrollcommand=info_scroll.set)
        
        self.info_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        info_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Загружаем шаблоны при старте, если пути сохранены
        if self.student_template_path and os.path.exists(self.student_template_path):
            self.load_template_from_path(self.student_template_path, "student")
        if self.leader_template_path and os.path.exists(self.leader_template_path):
            self.load_template_from_path(self.leader_template_path, "leader")
            
        # Сохраняем настройки при закрытии
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
    def on_closing(self):
        """Обработчик закрытия окна"""
        self.save_settings()
        self.root.destroy()
        
    def load_student_template(self):
        """Загружает шаблон для студентов"""
        self.load_template("student")
        
    def load_leader_template(self):
        """Загружает шаблон для лидеров"""
        self.load_template("leader")
        
    def load_template(self, template_type):
        """Загружает шаблон указанного типа"""
        file_path = filedialog.askopenfilename(
            title=f"Выберите шаблон для {'лидеров' if template_type == 'leader' else 'студентов'}",
            filetypes=[("Изображения", "*.png *.jpg *.jpeg *.bmp *.tiff")]
        )
        
        if file_path:
            self.load_template_from_path(file_path, template_type)
            
    def load_template_from_path(self, file_path, template_type):
        """Загружает шаблон из указанного пути"""
        try:
            image = Image.open(file_path)
            
            if template_type == "leader":
                self.leader_template = image
                self.leader_template_path = file_path
                self.leader_path_var.set(os.path.basename(file_path))
                self.display_image(self.leader_canvas, image)
            else:
                self.student_template = image
                self.student_template_path = file_path
                self.student_path_var.set(os.path.basename(file_path))
                self.display_image(self.student_canvas, image)
                
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить изображение: {str(e)}")
            
    def display_image(self, canvas, image):
        """Отображает изображение на canvas с масштабированием"""
        canvas_width = 200
        canvas_height = 150
        
        img_ratio = image.width / image.height
        canvas_ratio = canvas_width / canvas_height
        
        if img_ratio > canvas_ratio:
            display_width = canvas_width
            display_height = int(canvas_width / img_ratio)
        else:
            display_height = canvas_height
            display_width = int(canvas_height * img_ratio)
        
        display_image = image.resize((display_width, display_height), Image.Resampling.LANCZOS)
        photo = ImageTk.PhotoImage(display_image)
        
        canvas.delete("all")
        canvas.create_image(canvas_width//2, canvas_height//2, image=photo)
        canvas.image = photo
        
    def mm_to_pixels(self, mm, dpi=300):
        """Конвертирует мм в пиксели при заданном DPI"""
        return int(mm * dpi / 25.4)
    
    def calculate_layout(self):
        """Рассчитывает как будут размещены бейджики на листе А4"""
        try:
            # Получаем параметры в мм и конвертируем в пиксели
            badge_w_mm = float(self.badge_width_var.get())
            badge_h_mm = float(self.badge_height_var.get())
            margin_x_mm = float(self.margin_x_var.get())
            margin_y_mm = float(self.margin_y_var.get())
            spacing_x_mm = float(self.spacing_x_var.get())
            spacing_y_mm = float(self.spacing_y_var.get())
            cut_margin_mm = float(self.cut_margin_var.get())
            
            # Конвертируем в пиксели (300 DPI)
            self.badge_width = self.mm_to_pixels(badge_w_mm)
            self.badge_height = self.mm_to_pixels(badge_h_mm)
            self.margin_x = self.mm_to_pixels(margin_x_mm)
            self.margin_y = self.mm_to_pixels(margin_y_mm)
            self.spacing_x = self.mm_to_pixels(spacing_x_mm)
            self.spacing_y = self.mm_to_pixels(spacing_y_mm)
            self.cut_margin = self.mm_to_pixels(cut_margin_mm)
            
            # Рассчитываем сколько поместится на лист
            available_width = self.a4_width - 2 * self.margin_x
            available_height = self.a4_height - 2 * self.margin_y
            
            # Учитываем размер бейджика плюс запас для обрезки
            badge_total_width = self.badge_width + 2 * self.cut_margin
            badge_total_height = self.badge_height + 2 * self.cut_margin
            
            # Количество бейджиков по горизонтали и вертикали
            cols = max(1, (available_width + self.spacing_x) // (badge_total_width + self.spacing_x))
            rows = max(1, (available_height + self.spacing_y) // (badge_total_height + self.spacing_y))
            
            badges_per_page = cols * rows
            
            # Информация для пользователя
            info_text = f"""ПАРАМЕТРЫ РАЗМЕЩЕНИЯ:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📄 Лист А4: 210 × 297 мм
🏷️ Размер бейджика: {badge_w_mm} × {badge_h_mm} мм
📏 Поля: {margin_x_mm} мм (X), {margin_y_mm} мм (Y)
📐 Расстояние: {spacing_x_mm} мм (X), {spacing_y_mm} мм (Y)
✂️ Запас для обрезки: {cut_margin_mm} мм

РАЗМЕЩЕНИЕ НА ЛИСТЕ:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📊 Бейджиков в ряд: {cols}
📊 Количество рядов: {rows}
🔢 Всего на листе: {badges_per_page}

На каждом листе будет размещено {badges_per_page} пустых шаблонов бейджиков.
"""
            
            self.info_text.delete(1.0, tk.END)
            self.info_text.insert(tk.END, info_text)
            
            # Сохраняем параметры сетки
            self.grid_cols = cols
            self.grid_rows = rows
            self.badges_per_page = badges_per_page
            
        except ValueError as e:
            messagebox.showerror("Ошибка", "Проверьте правильность введенных значений")
    
    def generate_empty_pages(self, template_type):
        """Генерирует пустые листы А4 с шаблонами"""
        # Проверяем наличие нужного шаблона
        if template_type == "leader":
            if not self.leader_template:
                messagebox.showwarning("Предупреждение", "Сначала выберите шаблон для лидеров")
                return
            template = self.leader_template
            type_name = "лидеров"
        else:
            if not self.student_template:
                messagebox.showwarning("Предупреждение", "Сначала выберите шаблон для студентов")
                return
            template = self.student_template
            type_name = "студентов"
        
        # Проверяем рассчитанные параметры
        if not hasattr(self, 'badges_per_page'):
            messagebox.showwarning("Предупреждение", "Сначала рассчитайте размещение")
            return
            
        # Получаем количество листов
        try:
            pages_count = int(self.pages_count_var.get())
            if pages_count <= 0:
                messagebox.showerror("Ошибка", "Количество листов должно быть больше 0")
                return
        except ValueError:
            messagebox.showerror("Ошибка", "Введите корректное количество листов")
            return
            
        # Выбираем папку для сохранения
        output_dir = filedialog.askdirectory(title="Выберите папку для сохранения пустых листов А4")
        if not output_dir:
            return
            
        # Создаем прогресс-бар
        progress_window = tk.Toplevel(self.root)
        progress_window.title("Создание пустых листов А4...")
        progress_window.geometry("400x150")
        progress_window.transient(self.root)
        progress_window.grab_set()
        
        progress_var = tk.DoubleVar()
        progress_bar = ttk.Progressbar(progress_window, variable=progress_var, maximum=pages_count)
        progress_bar.pack(expand=True, fill=tk.X, padx=20, pady=10)
        
        status_label = ttk.Label(progress_window, text=f"Создание пустых листов для {type_name}...")
        status_label.pack()
        
        page_info_label = ttk.Label(progress_window, text="", foreground="blue")
        page_info_label.pack()
        
        progress_window.update()
        
        # Создаем листы
        for page_num in range(pages_count):
            page_info_label.config(text=f"Создание листа {page_num + 1} из {pages_count}")
            progress_window.update()
            
            # Создаем чистый лист А4
            page_image = Image.new('RGB', (self.a4_width, self.a4_height), 'white')
            
            # Размещаем пустые шаблоны на странице
            for idx in range(self.badges_per_page):
                # Определяем позицию в сетке
                row = idx // self.grid_cols
                col = idx % self.grid_cols
                
                # Рассчитываем координаты
                x = self.margin_x + col * (self.badge_width + 2 * self.cut_margin + self.spacing_x) + self.cut_margin
                y = self.margin_y + row * (self.badge_height + 2 * self.cut_margin + self.spacing_y) + self.cut_margin
                
                try:
                    # Масштабируем шаблон к нужному размеру
                    template_resized = template.resize((self.badge_width, self.badge_height), Image.Resampling.LANCZOS)
                    
                    # Вставляем на страницу
                    page_image.paste(template_resized, (x, y))
                    
                    # Добавляем направляющие линии для обрезки (тонкие серые линии)
                    from PIL import ImageDraw
                    draw = ImageDraw.Draw(page_image)
                    cut_color = "#CCCCCC"
                    
                    # Горизонтальные линии
                    draw.line([(x - self.cut_margin, y - self.cut_margin), 
                              (x + self.badge_width + self.cut_margin, y - self.cut_margin)], 
                             fill=cut_color, width=1)
                    draw.line([(x - self.cut_margin, y + self.badge_height + self.cut_margin), 
                              (x + self.badge_width + self.cut_margin, y + self.badge_height + self.cut_margin)], 
                             fill=cut_color, width=1)
                    
                    # Вертикальные линии
                    draw.line([(x - self.cut_margin, y - self.cut_margin), 
                              (x - self.cut_margin, y + self.badge_height + self.cut_margin)], 
                             fill=cut_color, width=1)
                    draw.line([(x + self.badge_width + self.cut_margin, y - self.cut_margin), 
                              (x + self.badge_width + self.cut_margin, y + self.badge_height + self.cut_margin)], 
                             fill=cut_color, width=1)
                    
                except Exception as e:
                    print(f"Ошибка создания шаблона в позиции {idx}: {e}")
            
            # Сохраняем страницу
            page_filename = f"Пустой_лист_{type_name}_{page_num + 1:02d}.png"
            page_path = os.path.join(output_dir, page_filename)
            page_image.save(page_path, "PNG", dpi=(300, 300))
            
            progress_var.set(page_num + 1)
            progress_window.update()
        
        progress_window.destroy()
        
        # Показываем результат
        result_message = f"""Успешно создано {pages_count} пустых листов А4 для {type_name}!

Параметры:
• Бейджиков на листе: {self.badges_per_page}
• Общее количество пустых шаблонов: {pages_count * self.badges_per_page}
• Сохранено в: {output_dir}

Файлы готовы для печати на принтере с разрешением 300 DPI."""
        
        messagebox.showinfo("Готово!", result_message)
        
        # Предложение открыть папку
        if messagebox.askyesno("Открыть папку?", "Открыть папку с созданными файлами?"):
            try:
                if os.name == 'nt':  # Windows
                    os.startfile(output_dir)
                elif os.name == 'posix':  # macOS и Linux
                    os.system(f'open "{output_dir}"' if 'darwin' in os.sys.platform else f'xdg-open "{output_dir}"')
            except:
                pass

if __name__ == "__main__":
    root = tk.Tk()
    app = EmptyBadgeGenerator(root)
    root.mainloop()
