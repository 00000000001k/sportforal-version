# sportforall/gui/event_contract_views.py

import tkinter as tk
import tkinter.ttk as ttk
import tkinter.messagebox as messagebox
import tkinter.filedialog as filedialog # Додано для діалогу вибору файлу
import os # Додано для роботи зі шляхами файлів
import sys # Додано для sys.exc_info

# Імпортуємо наші модулі з пакета sportforall
from sportforall import constants # Константы
from sportforall.models import AppData, Event, Contract, Item # Моделі даних
from sportforall import error_handling # Логирование ошибок

# TODO: Возможно, document_processor больше не нужен, если find_placeholders не используется для старого метода?
# from sportforall import document_processor # Для find_placeholders


class EventContractViews(ttk.Frame): # Успадковуємо від ttk.Frame
    """
    Ліва панель з списками заходів та договорів.
    Відображає Treeview для заходів та Treeview для договорів
    обраного заходу.
    """
    # !!! ИЗМЕНЕНО: Убран аргумент contracts_frame_container из __init__
    def __init__(self, master, app_data: AppData, callbacks: dict):
        """
        Ініціалізує панель списків.

        Args:
            master: Батьківський віджет (зазвичай Frame).
            app_data: Об'єкт AppData з усіма даними програми.
            callbacks: Словник функцій зворотного виклику до головного додатку (MainApp).
                       {
                           'event_selected': function(event_id), # Вибір заходу
                           'add_event': function(),              # Натискання "Додати Захід"
                           'delete_event': function(event_id),   # Натискання "Видалити Захід"
                           'contract_selected': function(contract_id), # Вибір договору
                           'add_contract': function(event_id),   # Натискання "Додати Договір"
                           'delete_contract': function(contract_id), # Натискання "Видалити Договір"
                           'generate_event_contracts': function(), # Натискання "Генерувати Захід"
                           'generate_single_contract': function(), # Натискання "Генерувати Договір" (можливо, в MainApp)
                           'contract_field_changed': function(contract_id, field_name, value), # Повідомлення про зміну поля
                       }
        """
        super().__init__(master)
        self._app_data = app_data # Зберігаємо посилання на дані
        self._callbacks = callbacks # Зберігаємо колбеки

        # !!! УДАЛЕНО: Нет необходимости хранить контейнер для фреймов договоров
        # self._contracts_frame_container = contracts_frame_container

        # Словники для зберігання IID в Treeview за ID об'єктів
        self._event_iids = {} # {event_id: event_iid}
        self._contract_iids = {} # {contract_id: contract_iid} (для поточного обраного заходу)

        # Словники для зберігання посилань на віджети (якщо потрібно)
        self._event_widgets = {}    # {event_id: widget}
        self._contract_widgets = {} # {contract_id: widget} (для поточного обраного заходу)

        self._current_event_id = None # ID поточного обраного заходу
        self._current_contract_id = None # ID поточного обраного договору (добавлено для отслеживания выбранного договора)


        # --- Створюємо віджети ---
        self.create_widgets()

    def create_widgets(self):
        """Створює всі віджети лівої панелі."""
        # Налаштовуємо сітку головного фрейму EventContractViews
        self.columnconfigure(0, weight=1) # Єдина колонка розтягується
        self.rowconfigure(0, weight=1)    # Рядок для списку заходів розтягується
        self.rowconfigure(1, weight=1)    # Рядок для списку договорів розтягується
        self.rowconfigure(2, weight=0)    # Рядок для кнопок заходів (не розтягується)
        self.rowconfigure(3, weight=0)    # Рядок для кнопок договорів (не розтягується)


        # --- Фрейм для списку заходів ---
        self.events_frame = ttk.LabelFrame(self, text="Заходи")
        self.events_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        self.events_frame.columnconfigure(0, weight=1)
        self.events_frame.rowconfigure(0, weight=1)

        # Таблиця заходів (Treeview)
        self.events_tree = ttk.Treeview(self.events_frame, columns=["Назва Заходу"], show="headings")
        self.events_tree.heading("Назва Заходу", text="Назва Заходу")
        self.events_tree.column("Назва Заходу", width=250, anchor="w") # Фіксована ширина
        self.events_tree.grid(row=0, column=0, sticky="nsew", padx=5, pady=2)

        # Скролбар для таблиці заходів
        events_scrollbar = ttk.Scrollbar(self.events_frame, orient="vertical", command=self.events_tree.yview)
        self.events_tree.configure(yscrollcommand=events_scrollbar.set)
        events_scrollbar.grid(row=0, column=1, sticky="ns")

        # Прив'язка події вибору елемента в таблиці заходів
        self.events_tree.bind("<<TreeviewSelect>>", self._on_event_tree_select)


        # --- Фрейм для списку договорів ---
        self.contracts_frame = ttk.LabelFrame(self, text="Договори")
        self.contracts_frame.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        self.contracts_frame.columnconfigure(0, weight=1)
        self.contracts_frame.rowconfigure(0, weight=1)

        # Таблиця договорів (Treeview)
        self.contracts_tree = ttk.Treeview(self.contracts_frame, columns=["Назва Договору", "Шаблон"], show="headings")
        self.contracts_tree.heading("Назва Договору", text="Назва Договору")
        self.contracts_tree.heading("Шаблон", text="Шаблон")
        self.contracts_tree.column("Назва Договору", width=150, anchor="w") # Фіксована ширина
        self.contracts_tree.column("Шаблон", width=100, anchor="w") # Фіксована ширина
        self.contracts_tree.grid(row=0, column=0, sticky="nsew", padx=5, pady=2)

        # Скролбар для таблиці договорів
        contracts_scrollbar = ttk.Scrollbar(self.contracts_frame, orient="vertical", command=self.contracts_tree.yview)
        self.contracts_tree.configure(yscrollcommand=contracts_scrollbar.set)
        contracts_scrollbar.grid(row=0, column=1, sticky="ns")

        # Прив'язка події вибору елемента в таблиці договорів
        self.contracts_tree.bind("<<TreeviewSelect>>", self._on_contract_tree_select)


        # --- Фрейм для кнопок заходів ---
        self.event_buttons_frame = ttk.Frame(self)
        self.event_buttons_frame.grid(row=2, column=0, sticky="ew", padx=5, pady=5)
        self.event_buttons_frame.columnconfigure(0, weight=1) # Кнопки розтягуємо рівномірно
        self.event_buttons_frame.columnconfigure(1, weight=1)

        # Кнопки для роботи з заходами
        # Кнопка "Додати Захід" (перенесено в MainApp top_buttons_frame)
        # Кнопка "Видалити Захід"
        ttk.Button(self.event_buttons_frame, text="🗑 Видалити Захід", command=self._delete_selected_event).grid(row=0, column=0, sticky="ew", padx=5)
        # Кнопка "Генерувати Захід" (перенесено в MainApp top_buttons_frame)


        # --- Фрейм для кнопок договорів ---
        self.contract_buttons_frame = ttk.Frame(self)
        self.contract_buttons_frame.grid(row=3, column=0, sticky="ew", padx=5, pady=5)
        self.contract_buttons_frame.columnconfigure(0, weight=1) # Кнопки розтягуємо рівномірно
        self.contract_buttons_frame.columnconfigure(1, weight=1)
        self.contract_buttons_frame.columnconfigure(2, weight=1)

        # Кнопки для роботи з договорами
        # Кнопка "Додати Договір"
        ttk.Button(self.contract_buttons_frame, text="➕ Додати Договір", command=self._add_contract_to_selected_event).grid(row=0, column=0, sticky="ew", padx=5)
        # Кнопка "Видалити Договір"
        ttk.Button(self.contract_buttons_frame, text="🗑 Видалити Договір", command=self._delete_selected_contract).grid(row=0, column=1, sticky="ew", padx=5)
        # Кнопка "Обрати Шаблон"
        # !!! Убедитесь, что ссылка на эту кнопку сохраняется здесь
        self.select_template_button = ttk.Button(self.contract_buttons_frame,
                                                 text="📁 Обрати Шаблон",
                                                 command=self._select_template_for_selected_contract) # <-- Устанавливаем команду
        self.select_template_button.grid(row=0, column=2, sticky="ew", padx=5)
        # При старте кнопка неактивна (состояние управляется из MainApp)
        self.select_template_button.config(state=tk.DISABLED)

        # Кнопка "Генерувати Договір" (перенесено в MainApp top_buttons_frame)


    # --- Методы загрузки данных в Treeviews ---

    # !!! ИЗМЕНЕНО: Теперь метод load_events принимает список объектов Event
    def load_events(self, events: list[Event]):
        """
        Завантажує список заходів у Treeview заходів.
        Приймає список об'єктів Event.
        """
        # Очищаємо поточний список заходів та словники відповідностей
        self._clear_events_tree()
        self._event_iids = {} # {event_id: event_iid}
        self._event_widgets = {} # Очищаємо словник віджетів заходів

        # Додаємо заходи зі списку об'єктів Event в Treeview
        # !!! ИЗМЕНЕНО: Итерируем напрямую по списку
        for event in events: # Итерируем по каждому объекту Event в списке
            event_id = event.id # Получаем ID из объекта Event
            # Додаємо захід як новий рядок у Treeview
            event_iid = self.events_tree.insert(
                "", tk.END,          # Вставляємо в корінь ("") в кінець (tk.END)
                iid=event_id,        # Використовуємо ID заходу як IID елемента Treeview
                values=(event.name,) # Значення для колонок (назва заходу)
            )
            self._event_iids[event_id] = event_iid # Зберігаємо відповідність ID -> IID
            self._event_widgets[event_id] = self.events_tree # Можливо, зберігати сам Treeview або IID достатньо

        # После загрузки, очищаем список договоров
        self.load_contracts({}) # load_contracts ожидает словарь {contract_id: Contract}
        # Примечание: load_contracts вызывается при выборе Event.
        # Здесь, при загрузке всех Events, мы просто очищаем список Contracts,
        # потому что ни один Event еще не выбран.


    def _clear_events_tree(self):
        """Очищає вміст Treeview заходів."""
        for item_iid in self.events_tree.get_children():
            self.events_tree.delete(item_iid)


    # Метод load_contracts ожидает словарь {contract_id: Contract}
    def load_contracts(self, contracts: dict[str, Contract]):
        """
        Завантажує список договорів у Treeview договорів.
        Приймає словник об'єктів Contract {contract_id: Contract}.
        Викликається при виборі нового заходу.
        """
        # Очищаємо поточний список договорів та словники відповідностей
        self._clear_contracts_tree()
        self._contract_iids = {}
        self._contract_widgets = {} # Очищаємо словник віджетів договорів

        # Додаємо договори з об'єкта Event в Treeview
        # Итерируем по словарю
        for contract_id, contract in contracts.items():
            # Додаємо договір як новий рядок у Treeview
            contract_iid = self.contracts_tree.insert(
                "", tk.END,              # Вставляємо в корінь ("") в кінець (tk.END)
                iid=contract_id,        # Використовуємо ID договору як IID елемента Treeview
                values=(contract.name, os.path.basename(contract.template_path) if contract.template_path else "Не обрано") # Значення для колонок
            )
            self._contract_iids[contract_id] = contract_iid # Зберігаємо відповідність ID -> IID
            self._contract_widgets[contract_id] = self.contracts_tree # Можливо, зберігати сам Treeview або IID достатньо


    def _clear_contracts_tree(self):
        """Очищає вміст Treeview договорів."""
        for item_iid in self.contracts_tree.get_children():
            self.contracts_tree.delete(item_iid)

    # --- Методы добавления/удаления элементов GUI ---

    def add_event_to_gui(self, event: Event):
        """
        Додає новий захід у Treeview заходів.
        Викликається з MainApp після додавання заходу до даних.
        """
        if event.id not in self._event_iids:
            event_iid = self.events_tree.insert("", tk.END, iid=event.id, values=(event.name,))
            self._event_iids[event.id] = event_iid
            self._event_widgets[event.id] = self.events_tree # Зберігаємо посилання


    def remove_event_from_gui(self, event_id: str):
        """
        Видаляє захід з Treeview заходів.
        Викликається з MainApp після видалення заходу з даних.
        """
        if event_id in self._event_iids:
            event_iid = self._event_iids[event_id]
            try:
                self.events_tree.delete(event_iid) # Видаляємо елемент з Treeview
            except tk.TclError:
                # Елемент міг бути вже видалений, ігноруємо помилку
                pass
            del self._event_iids[event_id] # Видаляємо зі словника відповідностей
            if event_id in self._event_widgets:
                del self._event_widgets[event_id] # Видаляємо зі словника віджетів


    def select_event(self, event_id: str):
        """
        Програмно обирає захід у Treeview заходів.
        """
        if event_id in self._event_iids:
            event_iid = self._event_iids[event_id]
            self.events_tree.selection_set(event_iid) # Встановлюємо виділення
            self.events_tree.focus(event_iid)       # Встановлюємо фокус


    def add_contract_to_gui(self, event_id: str, contract: Contract):
        """
        Додає новий договір у Treeview договорів (якщо цей договір належить
        до поточного обраного заходу).
        Викликається з MainApp після додавання договору до даних.
        """
        # Перевіряємо, чи доданий договір належить до поточного обраного заходу
        if self._current_event_id == event_id:
            # Проверяем, не добавлен ли уже договор
            if contract.id not in self._contract_iids:
                contract_iid = self.contracts_tree.insert(
                    "", tk.END,
                    iid=contract.id,
                    values=(contract.name, os.path.basename(contract.template_path) if contract.template_path else "Не обрано")
                )
                self._contract_iids[contract.id] = contract_iid
                self._contract_widgets[contract.id] = self.contracts_tree # Зберігаємо посилання
                # print(f"GUI: Додано договір '{contract.name}' ({contract.id}) до Treeview.") # Для дебагу
            # else:
                # print(f"GUI: Договір '{contract.name}' ({contract.id}) вже існує в Treeview.") # Для дебагу


    def remove_contract_from_gui(self, event_id: str, contract_id: str):
        """
        Видаляє договір з Treeview договорів (якщо цей договір належить
        до поточного обраного заходу).
        Викликається з MainApp після видалення договору з даних.
        """
        # Перевіряємо, чи видалений договір належить до поточного обраного заходу
        if self._current_event_id == event_id:
            if contract_id in self._contract_iids:
                contract_iid = self._contract_iids[contract_id]
                try:
                    self.contracts_tree.delete(contract_iid) # Видаляємо елемент з Treeview
                    # print(f"GUI: Видалено договір з ID {contract_id} з Treeview.") # Для дебагу
                except tk.TclError:
                    # Элемент мог быть уже удален или не найден в Treeview
                    # print(f"GUI: Помилка видалення договору з ID {contract_id} з Treeview (можливо, вже видалено).") # Для дебагу
                    pass # Елемент міг бути вже видалений, ігноруємо помилку
                del self._contract_iids[contract_id] # Видаляємо зі словника відповідностей
                if contract_id in self._contract_widgets:
                    del self._contract_widgets[contract_id] # Видаляємо зі словника віджетів
            # else:
                # print(f"GUI: Договір з ID {contract_id} не знайдено в словнику відповідностей для видалення.") # Для дебагу


    def select_contract(self, contract_id: str):
        """
        Програмно обирає договір у Treeview договорів.
        """
        # Проверяем, что текущее мероприятие выбрано и содержит этот договор
        if self._current_event_id is not None and contract_id in self._contract_iids:
            contract_iid = self._contract_iids[contract_id]
            self.contracts_tree.selection_set(contract_iid) # Встановлюємо виділення
            self.contracts_tree.focus(contract_iid)       # Встановлюємо фокус
            # print(f"GUI: Програмно обрано договір з ID {contract_id} в Treeview.") # Для дебагу
        # else:
            # print(f"GUI: Не вдалося програмно обрати договір з ID {contract_id} в Treeview (можливо, не в поточному заході).") # Для дебагу



    def update_contract_in_tree(self, contract: Contract): # <-- Отступ в 4 пробела
         """
         Оновлює інформацію про договір у Treeview договорів (напр., назву або шаблон).
         Викликається з MainApp, коли деталі договору змінюються.
         """ # <-- Отступ в 4 пробела для всей документации
         # Проверяем, что текущее мероприятие выбрано и содержит этот договор
         # Эта строка должна иметь отступ в 8 пробелов
         if self._current_event_id is not None and contract.id in self._contract_iids: # <-- Отступ в 8 пробелов
              # Код внутри этого if имеет дополнительный отступ (12 пробелов)
              contract_iid = self._contract_iids[contract.id] # <-- Отступ в 12 пробелов
              # Оновлюємо значення елемента в Treeview
              # Эта строка и следующие внутри вызова item должны иметь отступ в 12 пробелов
              self.contracts_tree.item( # <-- Отступ в 12 пробелов
                  contract_iid, # <-- Отступ в 16 пробелов для аргументов, перенесенных на новую строку
                  values=(contract.name, os.path.basename(contract.template_path) if contract.template_path else "Не обрано") # <-- Отступ в 16 пробелов
              ) # <-- Отступ в 12 пробелов для закрывающей скобки
              # print(f"GUI: Оновлено відображення договору '{contract.name}' ({contract.id}) в Treeview.") # Для дебагу # <-- Отступ в 12 пробелов

    # ... другие методы с отступом в 4 пробела ...


    # --- Обробники подій вибору в Treeview ---

    def _on_event_tree_select(self, event):
        """Обробник вибору елемента в Treeview заходів."""
        selected_iids = self.events_tree.selection()
        # print(f"GUI: Вибір заходу в Treeview. selected_iids: {selected_iids}") # Для дебагу

        if selected_iids:
            # Отримуємо ID обраного заходу (це IID елемента Treeview)
            selected_event_id = selected_iids[0]
            # print(f"GUI: Обрано захід з ID: {selected_event_id}") # Для дебагу

            # Проверяем, изменился ли выбранный захід
            if selected_event_id != self._current_event_id:
                # print(f"GUI: Обрано новий захід.") # Для дебагу
                # Очищаємо вибір в дереві договорів при зміні заходу
                self.contracts_tree.selection_set() # Знімаємо виділення з договорів
                self._current_contract_id = None # Скидаємо обраний договір

                # Викликаємо колбек до головного додатку, передаючи ID обраного заходу
                # MainApp._on_event_selected обработает смену мероприятия, загрузит договоры и обновит кнопки
                if self._callbacks.get('event_selected'):
                    self._callbacks['event_selected'](selected_event_id)

            # else: Захід не змінився, ничего не делаем
        else:
            # Если выбор сброшен (например, клик мимо элементов)
            # print(f"GUI: Вибір заходу знято.") # Для дебагу
            if self._current_event_id is not None:
                # print(f"GUI: Повідомляємо MainApp про зняття вибору заходу.") # Для дебагу
                # Вызываем колбек, передавая None, чтобы сообщить о сбросе выбора
                # MainApp._on_event_selected(None) обработает сброс выбора
                if self._callbacks.get('event_selected'):
                    self._callbacks['event_selected'](None)


    def _on_contract_tree_select(self, event):
        """Обробник вибору елемента в Treeview договорів."""
        selected_iids = self.contracts_tree.selection()
        # print(f"GUI: Вибір договору в Treeview. selected_iids: {selected_iids}") # Для дебагу

        if selected_iids:
            # Отримуємо ID обраного договору (це IID елемента Treeview)
            selected_contract_id = selected_iids[0]
            # print(f"GUI: Обрано договір з ID: {selected_contract_id}") # Для дебагу

            # Проверяем, изменился ли выбранный договор
            if selected_contract_id != self._current_contract_id:
                # print(f"GUI: Обрано новий договір.") # Для дебагу
                # Вызываем колбек до головного додатку, передавая ID обраного договора
                # MainApp._on_contract_selected обработает смену договора, отобразит детали и обновит кнопки
                if self._callbacks.get('contract_selected'):
                    self._callbacks['contract_selected'](selected_contract_id)

            # else: Договір не змінився, ничего не делаем
        else:
            # Если выбор сброшен
            # print(f"GUI: Вибір договору знято.") # Для дебагу
            if self._current_contract_id is not None:
                # print(f"GUI: Повідомляємо MainApp про зняття вибору договору.") # Для дебагу
                # Вызываем колбек, передавая None, чтобы сообщить о сбросе выбора
                # MainApp._on_contract_selected(None) обработает сброс выбора
                if self._callbacks.get('contract_selected'):
                    self._callbacks['contract_selected'](None)


    # --- Обробники натискань кнопок ---

    def _delete_selected_event(self):
        """Обробник натискання кнопки "Видалити Захід". Викликає колбек."""
        selected_iids = self.events_tree.selection()
        if selected_iids:
            selected_event_id = selected_iids[0]
            # Викликаємо колбек до головного додатку для видалення заходу
            # MainApp._delete_event обработает удаление из данных и вызовет remove_event_from_gui
            if self._callbacks.get('delete_event'):
                self._callbacks['delete_event'](selected_event_id)
        else:
            messagebox.showwarning("Увага", "Оберіть захід, який потрібно видалити.")


    def _add_contract_to_selected_event(self):
        """Обробник натискання кнопки "Додати Договір". Викликає колбек."""
        selected_iids = self.events_tree.selection()
        if selected_iids:
            selected_event_id = selected_iids[0]
            # Викликаємо колбек до головного додатку, передаючи ID заходу,
            # до якого потрібно додати договір.
            # MainApp._add_contract обработает добавление в данные и вызовет add_contract_to_gui
            if self._callbacks.get('add_contract'):
                self._callbacks['add_contract'](selected_event_id)
        else:
            messagebox.showwarning("Увага", "Оберіть захід, до якого потрібно додати договір.")


    def _delete_selected_contract(self):
        """Обробник натискання кнопки "Видалити Договір". Викликає колбек."""
        selected_iids = self.contracts_tree.selection()
        if selected_iids:
            selected_contract_id = selected_iids[0]
            # Викликаємо колбек до головного додатку для видалення договору.
            # MainApp._delete_contract обработает удаление из данных и вызовет remove_contract_from_gui
            if self._callbacks.get('delete_contract'):
                self._callbacks['delete_contract'](selected_contract_id)
        else:
            messagebox.showwarning("Увага", "Оберіть договір, який потрібно видалити.")


    # !!! Убедитесь, что этот метод присутствует и имеет правильные отступы
    def _select_template_for_selected_contract(self):
        """Обробник натискання кнопки "Обрати Шаблон"."""
        selected_iids = self.contracts_tree.selection()
        # print(f"GUI: Натиснуто Обрати Шаблон. selected_iids: {selected_iids}") # Для дебагу
        if selected_iids:
            selected_contract_id = selected_iids[0]
            # print(f"GUI: Обрано договір з ID: {selected_contract_id}") # Для дебагу
            contract = self._app_data.get_contract(selected_contract_id) # Используем _app_data из EventContractViews (ссылка на данные)
            if contract:
                # print("GUI: Об'єкт договору знайдено. Відкриваємо діалог вибору файлу.") # Для дебагу
                # !!! Убедитесь, что filedialog импортирован: import tkinter.filedialog as filedialog
                # import tkinter.filedialog as filedialog # Импорт должен быть в начале файла

                filepath = filedialog.askopenfilename(
                    title="Оберіть шаблон договору",
                    filetypes=[("Word Documents", "*.docm"), ("All files", "*.*")] # Дозволяємо .docm
                )
                # print(f"GUI: Шлях до обраного файлу: {filepath}") # Для дебагу

                if filepath:
                    # Зберігаємо шлях до шаблону в об'єкті договору
                    contract.template_path = filepath
                    # print(f"GUI: Шлях до шаблону в договорі оновлено на: {contract.template_path}") # Для дебагу

                    # !!! ДОБАВЛЕНО: Обновляем имя договора на основе имени файла шаблона
                    try:
                        # Получаем имя файла без пути
                       filename = os.path.basename(filepath) # !!! Убедитесь, что os импортирован: import os
                       # Удаляем расширение .docm (или другое)
                       contract_name_from_template = os.path.splitext(filename)[0]
                       # Обновляем атрибут name объекта договора
                       contract.name = contract_name_from_template
                       print(f"Ім'я договору оновлено з шаблону: {contract.name}") # Для дебагу
                    except Exception as e:
                       # Логируем ошибку, но не останавливаем процесс
                       error_handling.log_error(type(e), e, sys.exc_info()[2], level="WARNING", message=f"Помилка при отриманні ім'я договору з файлу шаблону '{filepath}'")
                       print(f"Помилка при отриманні ім'я з шаблону: {e}") # Для дебагу
                    # !!! КОНЕЦ ДОБАВЛЕНО


                    # Оновлюємо відображення в Treeview
                    self.update_contract_in_tree(contract) # Метод уже существует в этом классе

                    # Зберігаємо дані програми через колбек до MainApp
                    # Используем колбек contract_field_changed как триггер сохранения
                    if self._callbacks.get('contract_field_changed'):
                        # Повідомляємо MainApp про зміну (хоча це не поле, але викликає збереження)
                        # Передаем фиктивные данные поля, так как изменилось имя и шаблон
                        # print("GUI: Вызываем колбек contract_field_changed для сохранения.") # Для дебагу
                        self._callbacks['contract_field_changed'](contract.id, "template_path", filepath) # Используем template_path как триггер


        else:
            messagebox.showwarning("Увага", "Оберіть договір, для якого потрібно обрати шаблон.")

    # --- Додайте інші методи, якщо потрібно ---
    # Наприклад, метод для редагування товару