# sportforall/gui/contract_details_view.py

import tkinter as tk
import tkinter.ttk as ttk
import tkinter.messagebox as messagebox
import os
import sys
import traceback
from typing import Callable, Any

# Імпортуємо наші модулі з пакета sportforall
from sportforall import constants
from sportforall.models import AppData, Event, Contract, Item
# from sportforall.gui.custom_widgets import CustomEntry # !!! УДАЛЕНО: Переходим на ttk.Entry
from sportforall import error_handling
from sportforall.utils import number_to_currency_text

# Определяем цвета для стандартных ttk.Entry
# Стандартный фон ttk.Entry обычно светлый, но можно явно задать
ENTRY_BG_COLOR_ACTIVE = 'white' # Светлый фон для ввода
ENTRY_FG_COLOR_ACTIVE = 'black' # Черный текст при вводе
# Для неактивного состояния (placeholder убран) можно оставить стандартные цвета темы ttk

class ContractDetailsView(ttk.Frame):
    """
    Права панель для відображення та редагування деталей обраного договору.
    Теперь использует стандартные ttk.Entry виджеты.
    """
    def __init__(self, master, callbacks: dict):
        """
        Ініціалізує панель деталей договору.

        Args:
            master: Батьківський віджет (зазвичай вкладка Notebook).
            callbacks: Словник функцій зворотного виклику до головного додатку (MainApp).
                       {
                           'add_item': function(contract_id),        # Натискання "Додати товар" на деталях
                           'remove_item': function(contract_id, item_id), # Видалення товару
                           'calculate_total_sum': function(contract_id), # Перерахунок загальної суми
                           'contract_field_changed': function(contract_id, field_name, value), # Зміна поля договору
                           'get_context_menu': function() -> tk.Menu, # Запит контекстного меню
                       }
        """
        super().__init__(master)
        self._callbacks = callbacks # Зберігаємо колбеки
        self._current_contract: Contract | None = None # Об'єкт поточного обраного договору
        # Атрибут для зберігання ID поточного обраного товару (временный)
        self._current_item_id: str | None = None


        # --- Створюємо віджети ---
        self.create_widgets()

        # Привязка контекстного меню и горячих клавиш (теперь к ttk.Entry)
        self._bind_entry_shortcuts_and_context_menu()

        # При старте деактивируем кнопки товаров
        self._update_item_buttons_state()


    def create_widgets(self):
        """Створює всі віджети панелі деталей договору."""
        # Налаштовуємо сітку головного фрейму DetailsView
        self.columnconfigure(0, weight=1) # Єдина колонка розтягується
        self.rowconfigure(0, weight=0)    # Заголовок (не розтягується)
        self.rowconfigure(1, weight=1)    # Фрейм з полями договору (розтягується)
        self.rowconfigure(2, weight=1)    # Фрейм зі списком товарів (розтягується)
        self.rowconfigure(3, weight=0)    # Рядок для кнопок товарів (не розтягується)
        self.rowconfigure(4, weight=0)    # Мітка загальної суми (не розтягується)


        # --- Заголовок панелі ---
        self.title_label = ttk.Label(self, text="Деталі Договору: Не обрано", font=("Arial", 14, "bold"))
        self.title_label.grid(row=0, column=0, sticky="nw", padx=10, pady=5)


        # --- Фрейм для полів договору (с прокруткой) ---
        self.fields_canvas = tk.Canvas(self, borderwidth=0)
        self.fields_scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.fields_canvas.yview)
        self.fields_canvas.configure(yscrollcommand=self.fields_scrollbar.set)

        self.fields_scrollbar.grid(row=1, column=1, sticky="ns")
        self.fields_canvas.grid(row=1, column=0, sticky="nsew")

        # Создаем фрейм внутри Canvas, куда будут помещены поля
        self.fields_frame = ttk.Frame(self.fields_canvas)
        # Помещаем fields_frame в окно Canvas
        self.fields_canvas.create_window((0, 0), window=self.fields_frame, anchor="nw")

        # Настраиваем растяжение внутри fields_frame
        self.fields_frame.columnconfigure(0, weight=0) # Мітка (не розтягується)
        self.fields_frame.columnconfigure(1, weight=1) # Поле введення (розтягується)

        # Привязываем событие изменения размера фрейма полей для обновления скроллбара
        self.fields_frame.bind("<Configure>", lambda e: self.fields_canvas.configure(scrollregion = self.fields_canvas.bbox("all")))
        # Привязываем событие скролла мышью (для Canvas) - биндим на master
        self.master.bind_all("<MouseWheel>", self._on_mousewheel_fields_canvas) # Для Windows/macOS
        self.master.bind_all("<Button-4>", self._on_mousewheel_fields_canvas) # Для Linux up
        self.master.bind_all("<Button-5>", self._on_mousewheel_fields_canvas) # Для Linux down


        # --- Создаем поля ввода для стандартных полей договора из constants.FIELDS ---
        # !!! ИЗМЕНЕНО: Теперь используем ttk.Entry вместо CustomEntry !!!
        self._field_widgets = {} # Словарь для хранения ссылок на виджеты полей {field_name: ttk.Entry}
        for i, field_name in enumerate(constants.FIELDS):
            row_frame = ttk.Frame(self.fields_frame) # Фрейм для строки: Метка + Поле
            row_frame.grid(row=i, column=0, sticky="ew", padx=5, pady=2)
            row_frame.columnconfigure(1, weight=1) # Поле ввода растягивается в рамках строки

            # Метка поля
            label = ttk.Label(row_frame, text=f"{field_name.capitalize()}:")
            label.grid(row=0, column=0, sticky="w", padx=5)

            # Поле ввода (ttk.Entry)
            # У ttk.Entry нет встроенных плейсхолдеров или продвинутых проверок из CustomEntry
            # Плейсхолдер можно добавить вручную, но пока для простоты опустим
            entry = ttk.Entry(row_frame, width=400) # Используем стандартный ttk.Entry
            entry.grid(row=0, column=1, sticky="ew", padx=5)

            # !!! ИЗМЕНЕНО: Привязываем событие изменения содержимого (для ttk.Entry используем <KeyRelease>) !!!
            # Command колбек CustomEntry заменяем на bind <KeyRelease> и <FocusOut>
            # Передаем имя поля в лямбду
            entry.bind("<KeyRelease>", lambda e, name=field_name: self._on_field_change_callback(name, entry.get()))
            entry.bind("<FocusOut>", lambda e, name=field_name: self._on_field_change_callback(name, entry.get()), add="+") # add="+", чтобы не перезатереть привязку к <KeyRelease>

            # !!! ИЗМЕНЕНО: Устанавливаем цвета для ttk.Entry (светлый фон, черный текст) !!!
            # Можно установить цвета по умолчанию или при фокусе
            entry.configure(background=ENTRY_BG_COLOR_ACTIVE, foreground=ENTRY_FG_COLOR_ACTIVE)


            self._field_widgets[field_name] = entry # Сохраняем ссылку на виджет поля


        # --- Фрейм для списка товаров (с прокруткой) ---
        self.items_canvas = tk.Canvas(self, borderwidth=0)
        self.items_scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.items_canvas.yview)
        self.items_canvas.configure(yscrollcommand=self.items_scrollbar.set)

        self.items_scrollbar.grid(row=2, column=1, sticky="ns")
        self.items_canvas.grid(row=2, column=0, sticky="nsew")

        # Создаем фрейм внутри Canvas, куда будут помещены товары
        self.items_frame = ttk.Frame(self.items_canvas)
        # Помещаем items_frame в окно Canvas
        self.items_canvas.create_window((0, 0), window=self.items_frame, anchor="nw")

        # Настраиваем растяжение внутри items_frame
        # Первая колонка для самих виджетов товара (растягивается)
        self.items_frame.columnconfigure(0, weight=1)

        # Привязываем событие изменения размера фрейма товаров для обновления скроллбара
        self.items_frame.bind("<Configure>", lambda e: self.items_canvas.configure(scrollregion = self.items_canvas.bbox("all")))
        # Привязываем событие скролла мышью (для Canvas) - биндим на master
        self.master.bind_all("<MouseWheel>", self._on_mousewheel_fields_canvas) # Для Windows/macOS
        self.master.bind_all("<Button-4>", self._on_mousewheel_items_canvas) # Для Linux up
        self.master.bind_all("<Button-5>", self._on_mousewheel_items_canvas) # Для Linux down


        # --- Кнопки для работы с товарами ---
        self.item_buttons_frame = ttk.Frame(self)
        # Размещаем фрейм кнопок товаров отдельно (под canvas), не в canvas
        self.item_buttons_frame.grid(row=3, column=0, sticky="ew", padx=5, pady=5)
        # Настраиваем растяжение кнопок товаров
        self.item_buttons_frame.columnconfigure(0, weight=1) # Добавить товар
        self.item_buttons_frame.columnconfigure(1, weight=1) # Удалить товар

        # Кнопка "Додати товар"
        self._add_item_button = ttk.Button(self.item_buttons_frame, text="➕ Додати Товар", command=self._add_item)
        self._add_item_button.grid(row=0, column=0, sticky="ew", padx=5)

        # Кнопка "Видалити товар"
        # TODO: Сделать кнопку удаления товара активной только при выбранном товаре
        self.remove_item_button = ttk.Button(self.item_buttons_frame, text="🗑 Видалити Товар", command=self._remove_selected_item, state=tk.DISABLED)
        self.remove_item_button.grid(row=0, column=1, sticky="ew", padx=5)


        # --- Мітка загальної суми товарів ---
        self.total_sum_label = ttk.Label(self, text=f"{constants.ITEM_TOTAL_ROW_TEXT} 0.00")
        # Размещаем под items_canvas или item_buttons_frame, но не внутри них
        self.total_sum_label.grid(row=4, column=0, sticky="e", padx=10, pady=5) # Отдельный ряд для метки


        # --- Привязка контекстного меню (выполняется после создания всех полей) ---
        # Вызываем в __init__ после create_widgets()


    def _on_mousewheel_fields_canvas(self, event):
        # ... (ваш существующий код) ...
        if not self.winfo_exists():
             return

        widget_under_mouse = self.winfo_containing(event.x_root, event.y_root)

        if widget_under_mouse == self.fields_canvas or \
           (widget_under_mouse and self.fields_canvas.winfo_exists() and self.fields_canvas.winfo_children() and widget_under_mouse in self.fields_canvas.winfo_children()):

             if event.num == 0 or event.delta != 0:
                  self.fields_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
             elif event.num == 4:
                  self.fields_canvas.yview_scroll(-1, "units")
             elif event.num == 5:
                  self.fields_canvas.yview_scroll(1, "units")
             return "break"


    def _on_mousewheel_items_canvas(self, event):
        # ... (ваш существующий код) ...
        if not self.winfo_exists():
             return

        widget_under_mouse = self.winfo_containing(event.x_root, event.y_root)

        if widget_under_mouse == self.items_canvas or \
           (widget_under_mouse and self.items_canvas.winfo_exists() and self.items_canvas.winfo_children() and widget_under_mouse in self.items_canvas.winfo_children()):

             if event.num == 0 or event.delta != 0:
                  self.items_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
             elif event.num == 4:
                  self.items_canvas.yview_scroll(-1, "units")
             elif event.num == 5:
                  self.items_canvas.yview_scroll(1, "units")
             return "break"


    def _bind_entry_shortcuts_and_context_menu(self):
        """Привязывает горячие клавиши и контекстное меню ко всем полям ввода (ttk.Entry)."""
        # print("DEBUG: _bind_entry_shortcuts_and_context_menu called.")
        context_menu = self._callbacks.get('get_context_menu', lambda: None)()

        # Проходим по всем виджетам полей (теперь это ttk.Entry)
        for field_name, entry_widget in self._field_widgets.items():
            # Для ttk.Entry нет внутреннего виджета, работаем напрямую с entry_widget

            # Привязываем контекстное меню к правому клику мыши
            if context_menu:
                 # !!! ИСПРАВЛЕНО: Привязываем напрямую к ttk.Entry виджету !!!
                 entry_widget.bind("<Button-3>", lambda e, menu=context_menu: menu.post(e.x_root, e.y_root), add="+")

            # Привязываем стандартные горячие клавиши (Ctrl+C, Ctrl+V, Ctrl+X, Ctrl+A)
            # Стандартные ttk.Entry обычно обрабатывают их по умолчанию.
            # Явная привязка виртуальных событий <<Copy>>, <<Paste>>, <<Cut>>, <<SelectAll>>
            # гарантирует, что они сработают.
            # !!! ИСПРАВЛЕНО: Привязываем напрямую к ttk.Entry виджету !!!
            entry_widget.bind("<Control-c>", lambda e: e.widget.event_generate('<<Copy>>'), add="+")
            entry_widget.bind("<Control-v>", lambda e: e.widget.event_generate('<<Paste>>'), add="+")
            entry_widget.bind("<Control-x>", lambda e: e.widget.event_generate('<<Cut>>'), add="+")
            entry_widget.bind("<Control-a>", lambda e: e.widget.event_generate('<<SelectAll>>'), add="+")


    def display_contract_details(self, contract: Contract | None):
        # ... (ваш существующий код) ...
        self._current_contract = contract

        if contract:
            self.title_label.config(text=f"Деталі Договору: {contract.name}")
            self.update_contract_fields_gui(contract)
            self.update_items_gui(contract)
            # Обновляем общую сумму (вызывается внутри update_contract_fields_gui через autofill)

            # Управляем состоянием отдельных кнопок товаров
            if hasattr(self, '_add_item_button') and self._add_item_button:
                 self._add_item_button.config(state=tk.NORMAL)
            self._current_item_id = None
            self._update_item_buttons_state()

        else:
            self.title_label.config(text="Деталі Договору: Не обрано")
            self.clear_contract_fields_gui()
            self.clear_items_gui()
            self.update_total_sum_label(0.0)

            # Деактивировать кнопки добавления/удаления товаров
            if hasattr(self, '_add_item_button') and self._add_item_button:
                 self._add_item_button.config(state=tk.DISABLED)
            # Кнопка удаления уже деактивирована после clear_items_gui -> _update_item_buttons_state


    def update_contract_fields_gui(self, contract: Contract):
        """
        Обновляет значения полей ввода договора на основе данных из объекта Contract.
        Теперь для ttk.Entry.
        """
        # print(f"DEBUG: update_contract_fields_gui called for contract {contract.id if contract else 'None'}")
        if not contract:
             self.clear_contract_fields_gui()
             return

        # Проходим по всем виджетам полей (теперь ttk.Entry)
        for field_name, entry_widget in self._field_widgets.items():
            field_value = contract.fields.get(field_name, "")

            # !!! ИЗМЕНЕНО: Обновление значения для ttk.Entry !!!
            # Удаляем текущее содержимое
            entry_widget.delete(0, tk.END)
            # Вставляем новое значение
            entry_widget.insert(0, str(field_value)) # Вставляем как строку

            # Вызываем автозаполнение для пересчета зависимых полей (например, сума прописом)
            self._autofill_related_fields(field_name, field_value) # Передаем имя поля и его значение

        # После обновления всех полей, вызываем пересчет общей суммы на всякий случай.
        # Этот вызов больше не вызывает рекурсию, т.к. удален из конца этого метода.
        # Пересчет теперь инициируется через _autofill_related_fields
        pass # Вызов calculate_total_sum удален из конца этого метода


    def clear_contract_fields_gui(self):
        """Очищает все поля ввода деталей договора (теперь ttk.Entry)."""
        # print("DEBUG: clear_contract_fields_gui called.")
        for field_name, entry_widget in self._field_widgets.items():
             # !!! ИЗМЕНЕНО: Очистка для ttk.Entry !!!
             entry_widget.delete(0, tk.END)

        self._current_item_id = None
        self._update_item_buttons_state()


    def update_items_gui(self, contract: Contract):
        # ... (ваш существующий код, отображает товары как метки) ...
        self.clear_items_gui()

        if not contract:
             return

        self._item_widgets = {}
        for item in contract.items:
             item_label = ttk.Label(self.items_frame, text=f"Товар: {item.name}, К-во: {item.quantity}, Цена: {item.price}, Сумма: {item.get_total_sum()}")
             item_label.pack(fill="x", padx=5, pady=2)
             self._item_widgets[item.id] = item_label

             # Временная привязка клика к метке товара (для выбора)
             item_label.bind("<Button-1>", lambda e, i_id=item.id: self._on_item_selected(i_id), add="+")


    def clear_items_gui(self):
        # ... (ваш существующий код) ...
        for widget in self.items_frame.winfo_children():
             widget.destroy()
        self._item_widgets = {}

        self._current_item_id = None
        self._update_item_buttons_state()


    def update_total_sum_label(self, total_sum: float | int | str):
        # ... (ваш существующий код) ...
        numeric_sum = 0.0
        if total_sum is not None and str(total_sum).strip() != "":
             try:
                  numeric_sum = float(str(total_sum).replace(",", ".").strip())
             except (ValueError, TypeError):
                  print(f"Попередження: update_total_sum_label отримала некоректне значення суми: '{total_sum}'")
                  error_handling.log_error(type(ValueError), ValueError(f"Некоректне значення суми: '{total_sum}'"), sys.exc_info()[2], level="WARNING", message=f"Некоректне значення суми в update_total_sum_label: '{total_sum}'")
                  numeric_sum = 0.0

        formatted_sum = f"{numeric_sum:,.2f}".replace(",", " ").replace(".", ",")

        self.total_sum_label.config(text=f"{constants.ITEM_TOTAL_ROW_TEXT} {formatted_sum}")


    def _on_field_change_callback(self, field_name: str, value: str):
        """
        Обработчик изменения содержимого поля ввода договора (ttk.Entry).
        Вызывается при <KeyRelease> или <FocusOut>.
        """
        # print(f"DEBUG: _on_field_change_callback: Поле '{field_name}' изменено на '{value}'.")
        if not self._current_contract:
            return

        # Обновляем данные в объекте договора
        self._current_contract.update_field(field_name, value)

        # Инициируем автозаполнение или пересчет для связанных полей
        self._autofill_related_fields(field_name, value)

        # Вызываем колбек к MainApp для сохранения данных после изменения поля
        if self._callbacks.get('contract_field_changed'):
             self._callbacks['contract_field_changed'](self._current_contract.id, field_name, value)


    def _autofill_related_fields(self, changed_field: str, new_value_str: str):
        """
        Автоматично заповнює пов'язані поля при зміні числових полів договору,
        или инициирует пересчет суммы товаров при изменении полей товара.
        """
        # print(f"DEBUG: _autofill_related_fields called for field '{changed_field}' with value '{new_value_str}'.")

        if not self._current_contract:
            return

        # Поля товара, при изменении которых нужно пересчитывать сумму товара и договора
        # TODO: Когда будут редактируемые поля товара, эта логика будет там
        # item_numeric_fields_to_watch = ["кількість", "ціна за одиницю"] # Пример


        # Поля договора, которые являются числовыми и при изменении которых нужно пересчитывать "сума прописом"
        contract_numeric_fields = ["загальна сума", "разом"] # "сума прописом" - текстовое поле

        try:
            # --- Обработка изменения полей товара ---
            # Пока редактирование товаров не реализовано, этот блок не актуален для прямого ввода в GUI
            # Если бы редактирование было, здесь бы вызывался пересчет суммы договора.
            # if changed_field in item_numeric_fields_to_watch:
            #     if self._callbacks.get('calculate_total_sum'):
            #          self._callbacks['calculate_total_sum'](self._current_contract.id)
            #     pass


            # --- Обработка изменения числовых полей договора ("загальна сума" или "разом") ---
            if changed_field in contract_numeric_fields:
                 # print(f"DEBUG: _autofill_related_fields: Изменено числовое поле договора '{changed_field}'.")
                 new_contract_numeric_value_str = self._current_contract.fields.get(changed_field, "")

                 contract_numeric_value = 0.0
                 if new_contract_numeric_value_str is not None and str(new_contract_numeric_value_str).strip() != "":
                      try:
                           contract_numeric_value = float(str(new_contract_numeric_value_str).replace(",", ".").strip())
                      except (ValueError, TypeError) as num_conv_error:
                           print(f"Попередження: Не вдалося перетворити '{new_contract_numeric_value_str}' на число для поля договору '{changed_field}'. Помилка: {num_conv_error}")
                           error_handling.log_error(type(num_conv_error), num_conv_error, sys.exc_info()[2], level="WARNING", message=f"Не вдалося перетворити '{new_contract_numeric_value_str}' на число для поля договору '{changed_field}'.")
                           contract_numeric_value = 0.0

                 text_sum = number_to_currency_text(contract_numeric_value)
                 # print(f"DEBUG: Сумма прописью: '{text_sum}'")
                 self._current_contract.update_field("сума прописом", text_sum) # Обновляем поле в данных

                 # Оновлюємо відображення поля "сума прописом" в GUI (теперь ttk.Entry)
                 self.update_field_gui("сума прописом", text_sum) # Вызываем метод для обновления GUI

                 # Обновляем метку общей суммы, чтобы она соответствовала введенной вручную сумме договора
                 # print(f"DEBUG: Обновляем метку общей суммы с {contract_numeric_value}")
                 self.update_total_sum_label(contract_numeric_value)

            # Если изменилось нечисловое поле договора, просто сохраняем его (уже сделано в _on_field_change_callback)
            # и не вызываем пересчет или автозаполнение других полей (кроме сума прописом для числовых)


        except Exception as e:
             error_handling.log_error(type(e), e, sys.exc_info()[2], level="WARNING", message=f"Помилка автоматичного заповнення поля {changed_field} в договорі {self._current_contract.id if self._current_contract else 'Немає договору'}")
             print(f"Помилка в _autofill_related_fields для поля {changed_field}: {e}")


    # --- Методы работы с товарами (вызываются из кнопок на этой панели) ---

    def _add_item(self):
        # ... (ваш существующий код) ...
        if not self._current_contract:
             messagebox.showwarning("Увага", "Оберіть договір, до якого потрібно додати товар.")
             return

        if self._callbacks.get('add_item'):
             self._callbacks['add_item'](self._current_contract.id)


    def _add_item_to_contract_callback(self, item_data: dict):
        # ... (ваш существующий код) ...
        if not self._current_contract:
             return

        try:
            new_item = Item.from_dict(item_data)
            self._current_contract.add_item(new_item)
            self.update_items_gui(self._current_contract)

            if self._callbacks.get('calculate_total_sum'):
                 self._callbacks['calculate_total_sum'](self._current_contract.id)

        except Exception as e:
            messagebox.showerror("Помилка додавання товару", f"Не вдалося додати товар: {e}")
            exc_type, exc_value, exc_traceback = sys.exc_info()
            error_handling.log_error(exc_type, exc_value, exc_traceback, message=f"Помилка при додаванні товару до договору {self._current_contract.id}: {e}")


    def _remove_selected_item(self):
        # ... (ваш существующий код) ...
        if not self._current_contract:
             messagebox.showwarning("Увага", "Немає обраного договору.")
             return

        selected_item_id = self.get_selected_item_id()

        if not selected_item_id:
             messagebox.showwarning("Увага", "Оберіть товар, який потрібно видалити.")
             return

        item_to_delete = self._current_contract.find_item_by_id(selected_item_id)
        if not item_to_delete:
             messagebox.showwarning("Увага", "Товар не знайдено в даних.")
             return

        confirm = messagebox.askyesno("Підтвердження видалення", f"Видалити товар '{item_to_delete.name}'?")
        if confirm:
             try:
                 if self._callbacks.get('remove_item'):
                      self._callbacks['remove_item'](self._current_contract.id, selected_item_id)
             except Exception as e:
                 messagebox.showerror("Помилка видалення товару", f"Не вдалося видалити товар: {e}")
                 exc_type, exc_value, exc_traceback = sys.exc_info()
                 error_handling.log_error(exc_type, exc_value, exc_traceback, message=f"Помилка при видаленні товару {selected_item_id} з договору {self._current_contract.id}")

        self._current_item_id = None
        self._update_item_buttons_state()


    def get_selected_item_id(self) -> str | None:
        """
        Временный метод для получения ID выбранного товара из GUI.
        TODO: Реализовать выбор товара и получение ID из ItemView.
        Сейчас всегда возвращает None.
        """
        return self._current_item_id


    def _on_item_selected(self, item_id: str | None):
        """
        Временный обработчик выбора товара (клика по метке).
        TODO: Реализовать выбор товара в ItemView.
        """
        self._current_item_id = item_id
        self._update_item_buttons_state()


    def _update_item_buttons_state(self):
        # ... (ваш существующий код) ...
        if hasattr(self, 'remove_item_button') and self.remove_item_button:
             if self._current_item_id is not None:
                  self.remove_item_button.config(state=tk.NORMAL)
             else:
                  self.remove_item_button.config(state=tk.DISABLED)


    # --- Методы обновления GUI (теперь для ttk.Entry) ---

    # update_contract_fields_gui(self, contract: Contract) - реализован выше

    def update_field_gui(self, field_name: str, value: str):
        """
        Обновляет отображение значения конкретного поля договора в GUI (ttk.Entry).
        """
        # print(f"DEBUG: update_field_gui called for field '{field_name}' with value '{value}'")
        if field_name in self._field_widgets:
             entry_widget = self._field_widgets[field_name]
             # !!! ИЗМЕНЕНО: Обновление значения для ttk.Entry !!!
             entry_widget.delete(0, tk.END)
             entry_widget.insert(0, str(value)) # Вставляем как строку
             # Цвета для ttk.Entry уже заданы в create_widgets


    # clear_contract_fields_gui(self) - реализован выше
    # update_items_gui(self, contract: Contract) - реализован выше
    # clear_items_gui(self) - реализован выше
    # update_total_sum_label(self, total_sum: float | int | str) - реализован выше