# sportforall/gui/custom_widgets.py

import customtkinter as ctk
import tkinter as tk
import re
from typing import Callable, Any
# Додано Callable для анотації типу функції зворотного виклику

# Імпортуємо константи з верхнього рівня пакета sportforall
from sportforall import constants


# Определяем цвет фона для активного поля (когда текст вводится)
# Можно использовать 'white', или подобрать цвет, который лучше сочетается с темой.
# Для простоты начнем с 'white'.
ACTIVE_BG_COLOR = 'white'


# Расширенный класс Entry с настоящими плейсхолдерами и проверкой числовых полей
# Наследуется от CTkEntry
class CustomEntry(ctk.CTkEntry):
    """
    Расширенный виджет ввода текста CustomTkinter (CTkEntry) с поддержкой:
    - "Настоящего" плейсхолдера (текст-пример, который исчезает при фокусе).
    - Автоматической проверки и форматирования числового ввода для определенных полей.
    - Вызова внешнего колбека при изменении содержимого поля.
    - Управления цветом фона поля в зависимости от состояния (активно/неактивно).
    """
    def __init__(self, master, field_name: str | None = None, **kwargs):
        """
        Инициализирует CustomEntry.

        Args:
            master: Батьківський віджет.
            field_name: Опциональное имя поля, используемое для получения текста примера
                        из constants.EXAMPLES и определения, является ли поле числовым.
            **kwargs: Дополнительные аргументы для CustomTkinter.CTkEntry.
        """
        # print(f"DEBUG CustomEntry: Initializing field '{field_name}'") # For debug
        # Сохраняем имя поля
        self.field_name = field_name
        # Получаем текст примера из констант по имени поля, если оно задано
        self.example_text = constants.EXAMPLES.get(field_name, "Введіть значення")
        # Флаг, показывающий, отображается ли сейчас плейсхолдер
        self.is_placeholder_visible = True

        # Определяем, является ли это поле, требующее проверки числового ввода.
        self.is_numeric_field = field_name in ["кількість", "ціна за одиницю", "загальна сума", "разом"]

        # Атрибут для хранения внешнего колбека
        self._change_callback: Callable[[str, str | None], None] | None = None


        # Инициализируем базовый класс CTkEntry.
        # Убираем встроенный плейсхолдер CustomTkinter, т.к. реализуем свой.
        super().__init__(master, placeholder_text="", **kwargs)

        # --- Получаем стандартный цвет фона темы для неактивного состояния ---
        self._default_fg_color = self.cget("fg_color")

        # --- Получаем внутренний виджет Tkinter Entry ---
        self._entry = None # Инициализируем как None
        try:
             # Попытка получить доступ к _entry. Это может не сработать в вашей версии CTk.
             if hasattr(self, '_entry') and isinstance(self._entry, tk.Entry):
                 self._entry = self._entry
             else:
                 self._entry = None

        except AttributeError:
             self._entry = None

        # Печатаем сообщение об ошибке только один раз, если _entry не был получен
        if not self._entry:
             print("Помилка: Не вдалося отримати внутрішній віджет Tkinter Entry з CTkEntry.")
             # TODO: Логировать эту ошибку более правильно


        # --- Вставляем текст примера как обычный текст серого цвета при инициализации ---
        # Это делается только если _entry доступен
        if self._entry:
            self._entry.insert(0, self.example_text)
            # !!! УДАЛЕНЫ вызовы self.configure(...) отсюда, чтобы избежать ошибки !!!
            self.is_placeholder_visible = True
        else:
            # Если _entry не получен, плейсхолдер и цвет могут работать иначе в CTkEntry
            # Также не вызываем configure здесь
            print("Попередження: CustomEntry инициализирован без доступа к внутреннему Entry. Функции плейсхолдера и привязок могут работать некорректно.") # Печатается только если _entry None
            pass # Не вызываем configure здесь


        # --- Привязываем обработчики событий ... (используем self._entry с проверкой) ---
        # Это делается только если _entry доступен
        if self._entry: # Проверяем, существует ли _entry перед привязкой
            self._entry.bind("<FocusIn>", self._on_focus_in)
            self._entry.bind("<FocusOut>", self._on_focus_out)
            self._entry.bind("<KeyRelease>", self._on_change, add="+")
            self._entry.bind("<FocusOut>", self._on_change, add="+")

            # Привязываем обработчик ввода для числовых полей (только если _entry доступен)
            if self.is_numeric_field:
                self._entry.bind("<KeyRelease>", self._check_numeric_input, add="+")

        else:
             print("Попередження: CustomEntry не может привязать события из-за недоступности внутреннего Entry.") # Для дебагу


    # !!! ДОБАВЛЕНО/ПРОВЕРЕНО: Этот метод должен быть внутри класса CustomEntry !!!
    def set_change_callback(self, callback: Callable[[str, str | None], None]):
        # print(f"DEBUG CustomEntry: Устанавливаем колбек для поля '{self.field_name}'.") # Для дебагу
        self._change_callback = callback


    def _on_change(self, event):
        # print(f"DEBUG CustomEntry: _on_change called for field '{self.field_name}'. is_placeholder_visible: {self.is_placeholder_visible}") # For debug
        # Если сейчас отображается плейсхолдер, не вызываем колбек с ним как значением
        if self.is_placeholder_visible:
             return

        # Получаем текущее значение поля (НЕ используя self.get(), чтобы избежать рекурсии или проблем с плейсхолдером)
        # Используем self._entry.get(), но только если _entry доступен
        current_value = self._entry.get() if self._entry else ""


        # Если установлен внешний колбек, вызываем его
        if self._change_callback and callable(self._change_callback):
            try:
                # print(f"DEBUG CustomEntry: Вызываем внешний колбек для поля '{self.field_name}' со значением '{current_value}'") # Для дебагу
                # Передаем текущее значение и имя поля в колбек
                self._change_callback(current_value, self.field_name)
            except Exception as e:
                # TODO: Логировать ошибку
                print(f"ERROR in CustomEntry change callback for field '{self.field_name}': {e}") # Для дебагу
                # traceback.print_exc() # Для дебагу


    def _check_numeric_input(self, event):
        # print(f"DEBUG CustomEntry: _check_numeric_input called for field '{self.field_name}', key: {event.keysym}") # Для дебагу
        if not self._entry: # Проверяем, существует ли _entry
            # print("DEBUG CustomEntry: _check_numeric_input - _entry не доступен.") # Для дебагу
            return # Выходим, если _entry не существует

        # Разрешенные символы: цифры, точка, запятая, Backspace, Delete, стрелки, Home, End, Tab
        allowed_keys = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '.', ',', 'BackSpace', 'Delete', 'Left', 'Right', 'Home', 'End', 'Tab']

        # Если это один из разрешенных управляющих символов (Backspace, Delete, стрелки и т.д.)
        if event.keysym in allowed_keys[8:]: # Начиная с 'BackSpace'
             # print(f"DEBUG CustomEntry: Управляющий символ '{event.keysym}' разрешен.") # Для дебагу
             return # Разрешаем ввод управляющего символа

        # Получаем текущий текст в поле entry (с учетом только что введенного символа, т.к. KeyRelease)
        text_after_input = self._entry.get()
        input_char = event.char # Введенный символ (пусто для управляющих клавиш)

        # Проверяем, является ли введенный символ допустимым числовым символом
        # Если input_char не пустой и не является числом, точкой или запятой
        if input_char and input_char not in ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '.', ',']:
             # Если символ не является числом, точкой или запятой, отменяем его
             # print(f"DEBUG CustomEntry: Недопустимый символ '{input_char}'. Отменяем ввод.") # Для дебагу
             # Удаляем только что введенный символ
             # Это нужно делать осторожно, чтобы не нарушить работу Entry
             try:
                 # Попробуем удалить последний введенный символ
                 self._entry.delete("insert - 1 char", "insert")
             except tk.TclError:
                 # print("DEBUG CustomEntry: Не удалось удалить последний символ Tkinter.") # Для дебагу
                 pass # Игнорируем ошибку Tkinter при удалении

             return 'break' # Прерываем стандартную обработку события


        # Проверяем всю строку после ввода на соответствие формату числа (с одной точкой/запятой)
        # Заменяем запятую на точку для проверки float
        numeric_text = text_after_input.replace(',', '.')

        if numeric_text: # Проверяем, если строка не пустая после замены
             try:
                 float(numeric_text)
                 # Если преобразование в float удалось, значит, формат числа корректен.
                 # print(f"DEBUG CustomEntry: Ввод '{text_after_input}' допустим.") # Для дебагу
             except ValueError:
                 # Если преобразование в float не удалось (напр., две точки/запятые),
                 # отменяем ввод последнего символа.
                 # print(f"DEBUG CustomEntry: Ввод '{text_after_input}' недопустим (ValueError). Отменяем.") # Для дебагу
                 try:
                     # Попробуем удалить последний введенный символ, который привел к ошибке
                     self._entry.delete("insert - 1 char", "insert")
                 except tk.TclError:
                     # print("DEBUG CustomEntry: Не удалось удалить последний символ Tkinter.") # Для дебагу
                     pass # Игнорируем ошибку Tkinter при удалении

                 return 'break' # Прерываем стандартную обработку события

        # Если символ допустимый и формат строки корректен, позволяем ввод.
        # print(f"DEBUG CustomEntry: Ввод '{text_after_input}' разрешен.") # Для дебагу


    def _on_focus_in(self, event):
        """Обработчик события получения фокуса."""
        # print(f"DEBUG CustomEntry: Поле '{self.field_name}' получило фокус. is_placeholder_visible: {self.is_placeholder_visible}") # Для дебагу
        if self.is_placeholder_visible and self._entry: # Проверяем флаг и _entry
            self._entry.delete(0, tk.END) # Удаляем текст примера из внутреннего entry
            # ИСПРАВЛЕНО: Используем CustomTkinter опцию text_color на self
            self.configure(text_color='black') # Ставим черный цвет для обычного текста
            # ИСПРАВЛЕНО: Устанавливаем цвет фона на светлый при фокусе
            self.configure(fg_color=ACTIVE_BG_COLOR) # Ставим светлый фон
            self.is_placeholder_visible = False
        # else: Handle if _entry is not available - colors might still be set on self, but logic is partial
        # if not self._entry:
        #     self.configure(text_color='black')
        #     self.configure(fg_color=ACTIVE_BG_COLOR)


    def _on_focus_out(self, event):
        """Обработчик события потери фокуса."""
        # print(f"DEBUG CustomEntry: Поле '{self.field_name}' потеряло фокус.") # Для дебагу
        if self._entry and not self._entry.get():
            # Поле пустое после потери фокуса, отображаем placeholder
            # print(f"DEBUG CustomEntry: Поле '{self.field_name}' пустое после потери фокуса. Отображаем placeholder.") # Для дебагу
            self._entry.insert(0, self.example_text) # Вставляем текст примера
            # ИСПРАВЛЕНО: Используем CustomTkinter опцию text_color на self
            self.configure(text_color='gray') # Ставим серый цвет для плейсхолдера
            # ИСПРАВЛЕНО: Устанавливаем цвет фона обратно на стандартный (темный)
            self.configure(fg_color=self._default_fg_color) # Ставим стандартный фон
            self.is_placeholder_visible = True
        # else: Handle if _entry is not available or field is not empty
        # if not self._entry: # If _entry is None
        #     # Placeholder won't be inserted, but set colors
        #     if not self.get(): # Check using CustomEntry's get() (will be empty if _entry is None)
        #         self.configure(text_color='gray')
        #         self.configure(fg_color=self._default_fg_color)
        #     # else: field is not empty (impossible if _entry is None)
        elif self._entry and self._entry.get():
             # Поле не пустое, оставляем активные цвета (черный текст, светлый фон)
             # print(f"DEBUG CustomEntry: Поле '{self.field_name}' не пустое после потери фокуса.") # Для дебагу
             self.configure(text_color='black')
             self.configure(fg_color=ACTIVE_BG_COLOR)


    def get(self) -> str:
        # print(f"DEBUG CustomEntry: Вызван get() для поля '{self.field_name}'. is_placeholder_visible: {self.is_placeholder_visible}") # Для дебагу
        if self.is_placeholder_visible:
            return ""
        else:
            # Возвращаем актуальное содержимое внутреннего entry, если _entry существует
            return self._entry.get() if self._entry else ""


    def set(self, string: str):
        """
        Программно устанавливает значение поля.
        """
        # print(f"DEBUG CustomEntry: Вызван set('{string}') для поля '{self.field_name}'.") # Для дебагу
        if self._entry: # Проверяем, существует ли _entry
            self._entry.delete(0, tk.END)
            value_str = str(string) # Убедимся, что входные данные - строка
            self._entry.insert(0, value_str)
            self.is_placeholder_visible = False

            self.configure(text_color='black') # Ставим черный цвет
            # ИСПРАВЛЕНО: Устанавливаем цвет фона на светлый при установке значения
            self.configure(fg_color=ACTIVE_BG_COLOR) # Ставим светлый фон


            # Если поле оказалось пустым после установки значения (например, передали ""),
            # отображаем плейсхолдер и устанавливаем серый цвет
            if not self._entry.get(): # Проверяем актуальное содержимое после вставки
                # print(f"DEBUG CustomEntry: Поле '{self.field_name}' стало пустым после set. Отображаем placeholder.") # Для дебагу
                # Очищаем снова, чтобы гарантировать удаление текста примера, если он был случайно вставлен
                self._entry.delete(0, tk.END)
                self._entry.insert(0, self.example_text) # Вставляем текст примера
                self.configure(text_color='gray') # Ставим серый цвет
                # ИСПРАВЛЕНО: Устанавливаем цвет фона обратно на стандартный (темный)
                self.configure(fg_color=self._default_fg_color) # Ставим стандартный фон
                self.is_placeholder_visible = True
        # else: Handle if _entry is not available
        # print(f"Попередження: CustomEntry не может установить значение '{string}' из-за недоступности внутреннего Entry.") # Для дебагу
        # Try to set colors anyway, might affect border/widget appearance
        # self.configure(text_color='black')
        # self.configure(fg_color=ACTIVE_BG_COLOR)
        # if not string:
        #      self.configure(text_color='gray')
        #      self.configure(fg_color=self._default_fg_color)
        #      self.is_placeholder_visible = True


    def insert(self, index, string):
        """Вставляет текст в поле, управляя плейсхолдером."""
        # print(f"DEBUG CustomEntry: Вызван insert({index}, '{string}') для поля '{self.field_name}'.") # Для дебагу
        if self._entry: # Проверяем, существует ли _entry
            if self.is_placeholder_visible:
                 # Если виден плейсхолдер, сначала удаляем его
                 self._entry.delete(0, tk.END)
                 # Устанавливаем цвет текста на черный
                 self.configure(text_color='black') # Ставим черный
                 # ИСПРАВЛЕНО: Устанавливаем цвет фона на светлый при вставке (если плейсхолдер был виден) !!!
                 self.configure(fg_color=ACTIVE_BG_COLOR) # Ставим светлый фон
                 self.is_placeholder_visible = False

            # Вставляем новый текст
            self._entry.insert(index, string)
        # else: handle _entry not available
        # print(f"DEBUG CustomEntry: _entry не доступен, не удалось вставить текст '{string}'.") # Для дебагу


    def delete(self, first_index, last_index=None):
        """Удаляет текст из поля, управляя плейсхолдером."""
        # print(f"DEBUG CustomEntry: Вызван delete({first_index}, {last_index}) для поля '{self.field_name}'.") # Для дебагу
        if self._entry: # Проверяем, существует ли _entry
            if not self.is_placeholder_visible: # Удаляем только если не плейсхолдер
                 self._entry.delete(first_index, last_index)
                 # После удаления, если поле стало пустым, отображаем плейсхолдер
                 if not self._entry.get():
                     # print(f"DEBUG CustomEntry: Поле '{self.field_name}' стало пустым после delete. Отображаем placeholder.") # Для дебагу
                     self._entry.insert(0, self.example_text) # Вставляем текст примера
                     # Устанавливаем цвет текста на серый
                     self.configure(text_color='gray') # Ставим серый
                     # ИСПРАВЛЕНО: Устанавливаем цвет фона обратно на стандартный (темный) !!!
                     self.configure(fg_color=self._default_fg_color) # Ставим стандартный фон
                     self.is_placeholder_visible = True
            # else: Удаление при видимом плейсхолдере не меняет его состояние
        # else: handle _entry not available
        # print("DEBUG CustomEntry: _entry не доступен, не удалось удалить текст.") # Для дебагу


    def clear(self):
        """Очищает поле и возвращает плейсхолдер."""
        # print(f"DEBUG CustomEntry: Вызван clear() для поля '{self.field_name}'.") # Для дебагу
        if self._entry: # Проверяем, существует ли _entry
            self._entry.delete(0, tk.END) # Удаляем все содержимое внутреннего entry
            # Отображаем плейсхолдер
            self._entry.insert(0, self.example_text)
            # Устанавливаем цвет текста на серый
            self.configure(text_color='gray') # Ставим серый
            # ИСПРАВЛЕНО: Устанавливаем цвет фона обратно на стандартный (темный) !!!
            self.configure(fg_color=self._default_fg_color) # Ставим стандартный фон
            self.is_placeholder_visible = True
        # else: handle _entry not available
        # print("DEBUG CustomEntry: _entry не доступен, не удалось очистить поле.") # Для дебагу


    def get_internal_entry(self) -> tk.Entry | None:
        """
        Возвращает внутренний виджет Tkinter Entry.
        Используется для привязки событий (напр., контекстного меню).
        Возвращает None, если внутренний виджет недоступен.
        """
        # print(f"DEBUG CustomEntry: Вызван get_internal_entry() для поля '{self.field_name}'. Возвращает {'Entry' if self._entry else 'None'}.") # Для дебагу
        return self._entry # Возвращаем _entry (может быть None)