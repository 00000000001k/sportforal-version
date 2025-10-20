# custom_widgets.py
import customtkinter as ctk
import tkinter as tk
import re

class CustomEntry(ctk.CTkEntry):
    def __init__(self, master, field_name=None, examples_dict=None, **kwargs):
        self.field_name = field_name
        self.examples_dict = examples_dict if examples_dict else {}
        self.example_text = self.examples_dict.get(field_name, "Введіть значення")
        self.is_placeholder_visible = True

        # Определяем числовые поля для проверки ввода
        self.is_numeric_field = field_name in ["кількість", "ціна за одиницю", "загальна сума", "разом","загальна сума", "сума"] # Добавил "сума"

        # Инициализация без текста в placeholder_text
        super().__init__(master, placeholder_text="", **kwargs)

        # Вставляем пример как обычный текст серого цвета
        self._entry.insert(0, self.example_text)
        self._entry.config(foreground='gray')

        # Обработчики фокуса для автоматического обновления содержимого
        self._entry.bind("<FocusIn>", self._on_focus_in)
        self._entry.bind("<FocusOut>", self._on_focus_out)

        # Для числовых полей добавляем обработчик ввода
        if self.is_numeric_field:
            self._entry.bind("<KeyRelease>", self._check_numeric_input)

    def _check_numeric_input(self, event):
        """Проверяет числовой ввод после нажатия клавиши"""
        if self.is_placeholder_visible:
            return

        current_text = self._entry.get()
        if not current_text:
            return

        # Разрешаем пробелы, цифры, одну точку или одну запятую
        # Не разрешаем начинать с точки или запятой, но разрешаем пробелы в начале (которые потом уберутся)
        if not re.match(r'^\s*[\d\s]*([.,]\d*)?$', current_text):
            # Удаляем последний введенный символ, если он не соответствует шаблону
            cursor_pos = self._entry.index(tk.INSERT)
            # Чтобы избежать зацикливания, если первый символ невалиден
            if len(current_text) > 0:
                valid_text = current_text[:-1]
                # Дополнительная проверка, чтобы не удалить все, если вставка невалидна
                while len(valid_text) > 0 and not re.match(r'^\s*[\d\s]*([.,]\d*)?$', valid_text):
                    valid_text = valid_text[:-1]

                self._entry.delete(0, tk.END)
                self._entry.insert(0, valid_text)
                # Восстанавливаем позицию курсора
                try:
                    self._entry.icursor(min(cursor_pos -1, len(valid_text)))
                except:
                     self._entry.icursor(len(valid_text))


    def _on_focus_in(self, event):
        # При получении фокуса убираем плейсхолдер, если он отображается
        if self.is_placeholder_visible:
            self._entry.delete(0, "end")
            # Используем стандартный цвет текста для темы
            # self._entry.config(foreground='black') # Для светлой темы
            # Для темной темы может быть другой цвет, лучше использовать опции CTkinter
            text_color = self._apply_appearance_mode(self.cget("text_color"))
            self._entry.configure(foreground=text_color)
            self.is_placeholder_visible = False

    def _on_focus_out(self, event):
        # При потере фокуса, если поле пустое, показываем плейсхолдер
        current_text = self._entry.get().strip()
        if not current_text:
            self._entry.delete(0, "end")  # Очищаем поле перед вставкой плейсхолдера
            self._entry.insert(0, self.example_text)
            self._entry.config(foreground='gray')  # Серый цвет для плейсхолдера
            self.is_placeholder_visible = True

    def get(self):
        # Переопределяем метод get, чтобы не возвращал текст плейсхолдера
        if self.is_placeholder_visible:
            return ""
        # Удаляем пробелы в начале и конце, заменяем запятую на точку для числовых полей
        text_value = self._entry.get().strip()
        if self.is_numeric_field:
            return text_value.replace(",", ".") # Всегда возвращаем с точкой для единообразия
        return text_value

    def set_text(self, text):
        """Устанавливает текст в поле, убирая плейсхолдер, если нужно."""
        if self.is_placeholder_visible:
            self._entry.delete(0, "end")
            text_color = self._apply_appearance_mode(self.cget("text_color"))
            self._entry.configure(foreground=text_color)
            self.is_placeholder_visible = False
        else:
            self._entry.delete(0, "end")
        self._entry.insert(0, str(text) if text is not None else "")
        if not self._entry.get().strip(): # Если после вставки текста он пуст (например, вставили None или "")
            self._on_focus_out(None) # Отобразить плейсхолдер


    def insert(self, index, string):
        # Переопределяем метод insert, чтобы правильно вставлять текст
        if self.is_placeholder_visible:
            self._entry.delete(0, "end")
            text_color = self._apply_appearance_mode(self.cget("text_color"))
            self._entry.configure(foreground=text_color)
            self.is_placeholder_visible = False
        self._entry.insert(index, string)

    def delete(self, first_index, last_index=None):
        # Переопределяем метод delete для правильной работы с плейсхолдером
        if self.is_placeholder_visible:
            return
        self._entry.delete(first_index, last_index)

    def clear(self):
        """Очищает поле и отображает плейсхолдер."""
        self.set_text("") # Это вызовет _on_focus_out, если поле станет пустым