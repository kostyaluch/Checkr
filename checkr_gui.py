"""checkr_gui.py — GUI для автоматичної валідації товарного фіду e-commerce.

Графічний інтерфейс користувача для програми Checkr.
Дозволяє вибрати файл через діалогове вікно та зберігає результат поруч із вихідним файлом.

Використання:
    python checkr_gui.py
"""

import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
from threading import Thread

# Імпортуємо основну функцію валідації з checkr.py
try:
    from checkr import validate_feed
except ImportError:
    print("Помилка: не вдалося імпортувати модуль checkr.py", file=sys.stderr)
    sys.exit(1)


class CheckrGUI:
    """Головний клас графічного інтерфейсу для Checkr."""

    def __init__(self, root: tk.Tk):
        """Ініціалізація GUI.
        
        Args:
            root: Кореневе вікно tkinter
        """
        self.root = root
        self.root.title("Checkr — Валідація товарного фіду")
        self.root.geometry("600x400")
        self.root.resizable(False, False)
        
        self.selected_file = None
        self._setup_ui()

    def _setup_ui(self):
        """Налаштування елементів інтерфейсу."""
        # Заголовок
        title_label = tk.Label(
            self.root,
            text="Checkr — Валідація товарного фіду e-commerce",
            font=("Arial", 14, "bold"),
            pady=20
        )
        title_label.pack()

        # Опис
        desc_label = tk.Label(
            self.root,
            text="Програма перевіряє логічні розбіжності між назвою,\nописами та характеристиками товарів.",
            font=("Arial", 10),
            justify=tk.CENTER,
            pady=10
        )
        desc_label.pack()

        # Рамка для вибору файлу
        file_frame = tk.Frame(self.root, pady=20)
        file_frame.pack()

        self.file_label = tk.Label(
            file_frame,
            text="Файл не обрано",
            font=("Arial", 10),
            fg="gray",
            width=50,
            anchor="w"
        )
        self.file_label.pack(side=tk.LEFT, padx=10)

        select_btn = tk.Button(
            file_frame,
            text="Обрати файл",
            command=self._select_file,
            font=("Arial", 10),
            bg="#4CAF50",
            fg="white",
            padx=20,
            pady=5,
            cursor="hand2"
        )
        select_btn.pack(side=tk.LEFT)

        # Кнопка перевірки
        self.check_btn = tk.Button(
            self.root,
            text="Перевірити файл",
            command=self._validate_file,
            font=("Arial", 12, "bold"),
            bg="#2196F3",
            fg="white",
            padx=40,
            pady=10,
            state=tk.DISABLED,
            cursor="hand2"
        )
        self.check_btn.pack(pady=20)

        # Прогрес бар
        self.progress = ttk.Progressbar(
            self.root,
            mode='indeterminate',
            length=400
        )
        self.progress.pack(pady=10)

        # Текстове поле для виводу результатів
        output_frame = tk.Frame(self.root)
        output_frame.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)

        scrollbar = tk.Scrollbar(output_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.output_text = tk.Text(
            output_frame,
            height=8,
            width=70,
            font=("Courier", 9),
            yscrollcommand=scrollbar.set,
            state=tk.DISABLED,
            bg="#f5f5f5"
        )
        self.output_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.output_text.yview)

    def _select_file(self):
        """Відкриває діалог вибору файлу."""
        filetypes = [
            ("Підтримувані файли", "*.csv *.xlsx *.xls"),
            ("CSV файли", "*.csv"),
            ("Excel файли", "*.xlsx *.xls"),
            ("Всі файли", "*.*")
        ]
        
        filename = filedialog.askopenfilename(
            title="Оберіть файл для перевірки",
            filetypes=filetypes
        )
        
        if filename:
            self.selected_file = filename
            # Показуємо лише ім'я файлу, якщо шлях довгий
            display_name = Path(filename).name
            if len(display_name) > 45:
                display_name = display_name[:42] + "..."
            self.file_label.config(text=display_name, fg="black")
            self.check_btn.config(state=tk.NORMAL)
            self._log(f"Обрано файл: {filename}\n")

    def _validate_file(self):
        """Запускає валідацію вибраного файлу."""
        if not self.selected_file:
            messagebox.showwarning("Попередження", "Будь ласка, оберіть файл для перевірки.")
            return

        # Створюємо ім'я вихідного файлу поруч із вихідним
        input_path = Path(self.selected_file)
        output_file = input_path.parent / f"{input_path.stem}_result.xlsx"

        # Відключаємо кнопки під час обробки
        self.check_btn.config(state=tk.DISABLED)
        self.progress.start()
        self._log(f"Початок перевірки файлу: {input_path.name}\n")
        self._log(f"Результат буде збережено як: {output_file.name}\n")
        self._log("-" * 60 + "\n")

        # Запускаємо валідацію в окремому потоці, щоб не блокувати GUI
        thread = Thread(
            target=self._run_validation,
            args=(str(self.selected_file), str(output_file)),
            daemon=True
        )
        thread.start()

    def _run_validation(self, input_file: str, output_file: str):
        """Виконує валідацію у фоновому потоці.
        
        Args:
            input_file: Шлях до вхідного файлу
            output_file: Шлях до вихідного файлу
        """
        try:
            # Перенаправляємо stdout у GUI
            import io
            from contextlib import redirect_stdout

            output_buffer = io.StringIO()
            with redirect_stdout(output_buffer):
                validate_feed(input_file, output_file)
            
            # Отримуємо вивід з буфера
            output = output_buffer.getvalue()
            
            # Оновлюємо GUI в головному потоці
            self.root.after(0, self._validation_complete, output, output_file, None)
            
        except Exception as exc:
            # Передаємо помилку в головний потік
            self.root.after(0, self._validation_complete, "", output_file, exc)

    def _validation_complete(self, output: str, output_file: str, error: Exception = None):
        """Викликається після завершення валідації.
        
        Args:
            output: Текст виводу з валідації
            output_file: Шлях до вихідного файлу
            error: Помилка, якщо виникла
        """
        self.progress.stop()
        self.check_btn.config(state=tk.NORMAL)

        if error:
            self._log(f"\n❌ ПОМИЛКА: {error}\n")
            messagebox.showerror("Помилка", f"Виникла помилка під час перевірки:\n{error}")
        else:
            self._log(output + "\n")
            self._log("=" * 60 + "\n")
            self._log(f"✅ Перевірка завершена успішно!\n")
            self._log(f"📁 Результат збережено: {output_file}\n")
            
            messagebox.showinfo(
                "Успіх",
                f"Перевірка завершена!\n\nРезультат збережено:\n{output_file}"
            )

    def _log(self, message: str):
        """Додає повідомлення у текстове поле виводу.
        
        Args:
            message: Повідомлення для виводу
        """
        self.output_text.config(state=tk.NORMAL)
        self.output_text.insert(tk.END, message)
        self.output_text.see(tk.END)
        self.output_text.config(state=tk.DISABLED)


def main():
    """Точка входу GUI-програми."""
    root = tk.Tk()
    app = CheckrGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
