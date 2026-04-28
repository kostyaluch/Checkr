"""checkr_native_gui.py — Native desktop GUI для Checkr.

Нативний графічний інтерфейс користувача для програми Checkr.
Використовує tkinter для створення вікна програми без потреби у браузері.

Використання:
    python checkr_native_gui.py
"""

import sys
import os
import tempfile
import threading
from pathlib import Path
from tkinter import (
    Tk, Frame, Label, Button, filedialog, Text, Scrollbar,
    messagebox, ttk, StringVar, BOTH, RIGHT, LEFT, Y, END, WORD
)
from tkinter.font import Font

# Імпортуємо основну функцію валідації з checkr.py
try:
    from checkr import validate_feed
except ImportError:
    print("Помилка: не вдалося імпортувати модуль checkr.py", file=sys.stderr)
    sys.exit(1)


class CheckrNativeGUI:
    """Нативний GUI для Checkr на базі tkinter."""
    
    def __init__(self, root):
        """Ініціалізує GUI.
        
        Аргументи:
            root: Головне вікно tkinter (Tk instance).
        """
        self.root = root
        self.root.title("Checkr — Валідація товарного фіду")
        self.root.geometry("800x600")
        self.root.minsize(600, 400)
        
        # Змінні
        self.selected_file = None
        self.output_file = None
        self.is_processing = False
        
        # Налаштування стилів
        self.setup_styles()
        
        # Створення інтерфейсу
        self.create_widgets()
        
        # Центрування вікна
        self.center_window()
    
    def setup_styles(self):
        """Налаштовує стилі для віджетів."""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Стиль для кнопок
        style.configure('Main.TButton',
                       font=('Segoe UI', 10),
                       padding=10)
        
        # Стиль для прогрес-бару
        style.configure('Main.Horizontal.TProgressbar',
                       troughcolor='#f0f0f0',
                       background='#667eea',
                       borderwidth=0,
                       thickness=25)
    
    def create_widgets(self):
        """Створює всі віджети інтерфейсу."""
        # Заголовок
        header_frame = Frame(self.root, bg='#667eea', height=80)
        header_frame.pack(fill='x')
        header_frame.pack_propagate(False)
        
        title_font = Font(family='Segoe UI', size=20, weight='bold')
        title_label = Label(header_frame,
                          text="🔍 Checkr",
                          font=title_font,
                          bg='#667eea',
                          fg='white')
        title_label.pack(pady=10)
        
        subtitle_label = Label(header_frame,
                             text="Автоматична валідація товарного фіду e-commerce",
                             font=('Segoe UI', 10),
                             bg='#667eea',
                             fg='white')
        subtitle_label.pack()
        
        # Основна область
        main_frame = Frame(self.root, bg='white', padx=20, pady=20)
        main_frame.pack(fill=BOTH, expand=True)
        
        # Область вибору файлу
        file_frame = Frame(main_frame, bg='white')
        file_frame.pack(fill='x', pady=(0, 10))
        
        self.file_label = Label(file_frame,
                               text="Файл не обрано",
                               font=('Segoe UI', 10),
                               bg='white',
                               fg='#666',
                               anchor='w')
        self.file_label.pack(side=LEFT, fill='x', expand=True, padx=(0, 10))
        
        select_button = ttk.Button(file_frame,
                                  text="Вибрати файл",
                                  style='Main.TButton',
                                  command=self.select_file)
        select_button.pack(side=RIGHT)
        
        # Кнопка перевірки
        self.check_button = ttk.Button(main_frame,
                                      text="Перевірити файл",
                                      style='Main.TButton',
                                      command=self.validate_file,
                                      state='disabled')
        self.check_button.pack(fill='x', pady=(0, 10))
        
        # Прогрес-бар
        self.progress_var = StringVar(value="Готово до роботи")
        progress_label = Label(main_frame,
                             textvariable=self.progress_var,
                             font=('Segoe UI', 9),
                             bg='white',
                             fg='#666')
        progress_label.pack(anchor='w', pady=(0, 5))
        
        self.progress_bar = ttk.Progressbar(main_frame,
                                           style='Main.Horizontal.TProgressbar',
                                           mode='indeterminate')
        self.progress_bar.pack(fill='x', pady=(0, 10))
        
        # Область виводу результатів
        output_label = Label(main_frame,
                           text="Результат перевірки:",
                           font=('Segoe UI', 10, 'bold'),
                           bg='white',
                           fg='#333',
                           anchor='w')
        output_label.pack(anchor='w', pady=(0, 5))
        
        # Текстове поле з прокруткою
        output_frame = Frame(main_frame, bg='white')
        output_frame.pack(fill=BOTH, expand=True)
        
        scrollbar = Scrollbar(output_frame)
        scrollbar.pack(side=RIGHT, fill=Y)
        
        self.output_text = Text(output_frame,
                               font=('Courier New', 9),
                               bg='#f8f9fa',
                               fg='#333',
                               wrap=WORD,
                               yscrollcommand=scrollbar.set,
                               state='disabled',
                               relief='flat',
                               borderwidth=1)
        self.output_text.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.config(command=self.output_text.yview)
        
        # Теги для кольорового форматування
        self.output_text.tag_config('success', foreground='#28a745', font=('Courier New', 9, 'bold'))
        self.output_text.tag_config('error', foreground='#dc3545', font=('Courier New', 9, 'bold'))
        self.output_text.tag_config('info', foreground='#667eea', font=('Courier New', 9, 'bold'))
        
        # Кнопка відкриття результату
        self.open_button = ttk.Button(main_frame,
                                     text="Відкрити файл результату",
                                     style='Main.TButton',
                                     command=self.open_result_file,
                                     state='disabled')
        self.open_button.pack(fill='x', pady=(10, 0))
    
    def center_window(self):
        """Центрує вікно на екрані."""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def select_file(self):
        """Відкриває діалог вибору файлу."""
        filetypes = (
            ('CSV файли', '*.csv'),
            ('Excel файли', '*.xlsx *.xls'),
            ('Всі файли', '*.*')
        )
        
        filename = filedialog.askopenfilename(
            title='Виберіть файл для перевірки',
            filetypes=filetypes
        )
        
        if filename:
            self.selected_file = filename
            # Показуємо ім'я файлу
            file_path = Path(filename)
            self.file_label.config(text=f"Обрано: {file_path.name}", fg='#333')
            self.check_button.config(state='normal')
            
            # Очищаємо попередній результат
            self.output_text.config(state='normal')
            self.output_text.delete(1.0, END)
            self.output_text.config(state='disabled')
            self.open_button.config(state='disabled')
            self.output_file = None
    
    def append_output(self, text, tag=None):
        """Додає текст до області виводу.
        
        Аргументи:
            text: Текст для виводу.
            tag: Тег для форматування ('success', 'error', 'info' або None).
        """
        self.output_text.config(state='normal')
        if tag:
            self.output_text.insert(END, text, tag)
        else:
            self.output_text.insert(END, text)
        self.output_text.see(END)
        self.output_text.config(state='disabled')
    
    def validate_file(self):
        """Запускає валідацію вибраного файлу."""
        if not self.selected_file or self.is_processing:
            return
        
        # Перевіряємо існування файлу
        if not Path(self.selected_file).exists():
            messagebox.showerror("Помилка", "Обраний файл не існує")
            return
        
        # Створюємо ім'я вихідного файлу
        input_path = Path(self.selected_file)
        output_path = input_path.parent / f"{input_path.stem}_result.xlsx"
        self.output_file = str(output_path)
        
        # Очищаємо вивід
        self.output_text.config(state='normal')
        self.output_text.delete(1.0, END)
        self.output_text.config(state='disabled')
        
        # Вимикаємо кнопки та показуємо прогрес
        self.check_button.config(state='disabled')
        self.open_button.config(state='disabled')
        self.progress_var.set("Обробка файлу...")
        self.progress_bar.start(10)
        self.is_processing = True
        
        # Запускаємо валідацію в окремому потоці
        thread = threading.Thread(target=self.run_validation)
        thread.daemon = True
        thread.start()
    
    def run_validation(self):
        """Виконує валідацію у фоновому потоці."""
        try:
            # Перенаправляємо вивід
            import io
            from contextlib import redirect_stdout
            
            output_buffer = io.StringIO()
            with redirect_stdout(output_buffer):
                validate_feed(self.selected_file, self.output_file)
            
            output = output_buffer.getvalue()
            
            # Оновлюємо GUI з результатами (безпечно з потоку)
            self.root.after(0, self.on_validation_success, output)
            
        except Exception as e:
            # Оновлюємо GUI з помилкою
            self.root.after(0, self.on_validation_error, str(e))
    
    def on_validation_success(self, output):
        """Обробляє успішне завершення валідації.
        
        Аргументи:
            output: Текст виводу валідації.
        """
        self.progress_bar.stop()
        self.progress_var.set("Перевірка завершена успішно!")
        
        # Виводимо результат
        self.append_output("✅ Перевірка завершена успішно!\n\n", 'success')
        self.append_output(output)
        self.append_output(f"\n\n📁 Результат збережено: {Path(self.output_file).name}\n", 'info')
        
        # Активуємо кнопки
        self.check_button.config(state='normal')
        self.open_button.config(state='normal')
        self.is_processing = False
        
        messagebox.showinfo("Готово", f"Перевірка завершена!\nРезультат збережено у:\n{self.output_file}")
    
    def on_validation_error(self, error):
        """Обробляє помилку валідації.
        
        Аргументи:
            error: Текст помилки.
        """
        self.progress_bar.stop()
        self.progress_var.set("Виникла помилка під час обробки")
        
        # Виводимо помилку
        self.append_output("❌ Помилка під час обробки файлу:\n\n", 'error')
        self.append_output(error)
        
        # Активуємо кнопку перевірки
        self.check_button.config(state='normal')
        self.is_processing = False
        
        messagebox.showerror("Помилка", f"Виникла помилка під час обробки:\n{error}")
    
    def open_result_file(self):
        """Відкриває файл результату у стандартній програмі."""
        if not self.output_file or not Path(self.output_file).exists():
            messagebox.showwarning("Увага", "Файл результату не знайдено")
            return
        
        try:
            # Відкриваємо файл у стандартній програмі ОС
            import subprocess
            import platform
            
            if platform.system() == 'Windows':
                os.startfile(self.output_file)
            elif platform.system() == 'Darwin':  # macOS
                subprocess.call(['open', self.output_file])
            else:  # Linux
                subprocess.call(['xdg-open', self.output_file])
        except Exception as e:
            messagebox.showerror("Помилка", f"Не вдалося відкрити файл:\n{e}")


def main():
    """Точка входу програми."""
    root = Tk()
    app = CheckrNativeGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
