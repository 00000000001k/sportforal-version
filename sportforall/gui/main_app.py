# sportforall/gui/main_app.py

import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox
import os
import sys
# Добавлено для работы с датой
# from datetime import datetime # Возможно потребуется, если будем работать с датами в полях

# Импортируем классы из наших модулей
from sportforall.app_data import AppData
from sportforall.models import Contract, Event, Field  # Убедитесь, что Field импортирован
from sportforall.word_generator_win32 import generate_document_win32 # Генератор Word через win32com
from sportforall import error_handling # Модуль для логирования ошибок


# Определение констант для имен полей
# TODO: Перенести эти константы в отдельный файл constants.py
class FieldConstants:
    CONTRACT_NAME = "Назва договору"
    TEMPLATE_PATH = "Шлях до шаблону"
    OUTPUT_DIR = "Папка для збереження"
    EVENT_NAME = "Назва заходу"
    # Определите здесь имена всех ваших стандартных полей, которые могут быть в договорах
    # Эти имена должны совпадать с ключами в словаре fields объекта Contract
    FIELD_ТОВАР = "товар" # Пример: поле для названия товара/услуги
    FIELD_ДАТА_ДОГОВОРУ = "дата_договору" # Пример: поле для даты договора
    FIELD_ЦІНА = "ціна" # Пример: поле для цены
    # Добавьте другие поля по необходимости...


class MainApplication(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Sport For All Document Generator")
        self.geometry("1000x700")

        # Инициализация данных приложения
        self.app_data = AppData("app_data.json")
        self.app_data.load_data()

        # Переменные для хранения текущего выбранного договора или мероприятия
        self.selected_contract: Contract | None = None
        self.selected_event: Event | None = None

        # Словарь для хранения значений полей текущего выбранного договора,
        # введенных или измененных пользователем в GUI.
        # Ключи - имена полей, значения - их текущие значения из Entry виджетов.
        self.selected_contract_fields: dict[str, str] = {}


        # --- Создание элементов интерфейса ---

        # Конфигурация сетки главного окна
        self.grid_columnconfigure(0, weight=0) # Левая панель (списки) не растягивается
        self.grid_columnconfigure(1, weight=1) # Правая панель (детали) растягивается
        self.grid_rowconfigure(0, weight=1) # Единственная строка растягивается

        # Левая панель (списки мероприятий и договоров)
        self.left_frame = ctk.CTkFrame(self, width=250, corner_radius=0)
        self.left_frame.grid(row=0, column=0, sticky="nsew")
        self.left_frame.grid_rowconfigure(0, weight=0) # Заголовок не растягивается
        self.left_frame.grid_rowconfigure(1, weight=1) # Список мероприятий растягивается
        self.left_frame.grid_rowconfigure(2, weight=0) # Разделитель не растягивается
        self.left_frame.grid_rowconfigure(3, weight=1) # Список договоров растягивается
        self.left_frame.grid_rowconfigure(4, weight=0) # Кнопки не растягиваются


        # Заголовок для мероприятий
        self.event_label = ctk.CTkLabel(self.left_frame, text="Мероприятия", font=ctk.CTkFont(size=12, weight="bold"))
        self.event_label.grid(row=0, column=0, padx=10, pady=5, sticky="ew")

        # Список мероприятий
        self.event_listbox = tk.Listbox(self.left_frame, width=40, height=10, font=("Arial", 10),
                                        bg=self._apply_appearance_mode(ctk.ThemeManager.theme["CTkTextbox"]["fg_color"], ctk.ThemeManager.theme["CTkTextbox"]["_dark_fg_color"]),
                                        fg=self._apply_appearance_mode(ctk.ThemeManager.theme["CTkTextbox"]["text_color"], ctk.ThemeManager.theme["CTkTextbox"]["_dark_text_color"]),
                                        selectbackground=self._apply_appearance_mode(ctk.ThemeManager.theme["CTkOptionmenu"]["button_color"][0], ctk.ThemeManager.theme["CTkOptionmenu"]["_dark_button_color"][0]),
                                        selectforeground=self._apply_appearance_mode(ctk.ThemeManager.theme["CTkOptionmenu"]["text_color"], ctk.ThemeManager.theme["CTkOptionmenu"]["_dark_text_color"]),
                                        borderwidth=0, highlightthickness=0 # Убираем стандартные границы Tk
                                        )
        self.event_listbox.grid(row=1, column=0, padx=10, pady=0, sticky="nsew")
        self.event_listbox.bind("<<ListboxSelect>>", self._on_event_select)


        # Разделитель
        self.separator = ctk.CTkFrame(self.left_frame, height=2, fg_color="gray", corner_radius=0)
        self.separator.grid(row=2, column=0, padx=10, pady=5, sticky="ew")


        # Заголовок для договоров
        self.contract_label = ctk.CTkLabel(self.left_frame, text="Договоры", font=ctk.CTkFont(size=12, weight="bold"))
        self.contract_label.grid(row=3, column=0, padx=10, pady=5, sticky="ew")


        # Список договоров
        self.contract_listbox = tk.Listbox(self.left_frame, width=40, height=15, font=("Arial", 10),
                                           bg=self._apply_appearance_mode(ctk.ThemeManager.theme["CTkTextbox"]["fg_color"], ctk.ThemeManager.theme["CTkTextbox"]["_dark_fg_color"]),
                                           fg=self._apply_appearance_mode(ctk.ThemeManager.theme["CTkTextbox"]["text_color"], ctk.ThemeManager.theme["CTkTextbox"]["_dark_text_color"]),
                                           selectbackground=self._apply_appearance_mode(ctk.ThemeManager.theme["CTkOptionmenu"]["button_color"][0], ctk.ThemeManager.theme["CTkOptionmenu"]["_dark_button_color"][0]),
                                           selectforeground=self._apply_appearance_mode(ctk.ThemeManager.theme["CTkOptionmenu"]["text_color"], ctk.ThemeManager.theme["CTkOptionmenu"]["_dark_text_color"]),
                                           borderwidth=0, highlightthickness=0 # Убираем стандартные границы Tk
                                           )
        self.contract_listbox.grid(row=4, column=0, padx=10, pady=0, sticky="nsew")
        self.contract_listbox.bind("<<ListboxSelect>>", self._on_contract_select)


        # Кнопки на левой панели (добавить, удалить и т.д.)
        self.left_buttons_frame = ctk.CTkFrame(self.left_frame, fg_color="transparent")
        self.left_buttons_frame.grid(row=5, column=0, padx=10, pady=10, sticky="ew")
        self.left_buttons_frame.grid_columnconfigure(0, weight=1) # Кнопки растягиваются


        self.add_contract_button = ctk.CTkButton(self.left_buttons_frame, text="Добавить договор", command=self._add_contract)
        self.add_contract_button.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

        self.delete_contract_button = ctk.CTkButton(self.left_buttons_frame, text="Удалить договор", command=self._delete_contract, fg_color="red")
        self.delete_contract_button.grid(row=1, column=0, padx=5, pady=5, sticky="ew")

        # TODO: Добавить кнопки для мероприятий (добавить/удалить)


        # Правая панель (детали выбранного элемента)
        self.right_frame = ctk.CTkFrame(self, corner_radius=0)
        self.right_frame.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
        self.right_frame.grid_columnconfigure(0, weight=1) # Основное содержимое растягивается
        self.right_frame.grid_rowconfigure(1, weight=1) # Секция полей растягивается


        # Заголовок правой панели
        self.right_panel_title = ctk.CTkLabel(self.right_frame, text="Выберите договор или мероприятие", font=ctk.CTkFont(size=14, weight="bold"))
        self.right_panel_title.grid(row=0, column=0, padx=10, pady=10, sticky="ew")


        # Фрейм для отображения деталей (переключается между договором и мероприятием)
        self.detail_frame = ctk.CTkFrame(self.right_frame, fg_color="transparent")
        self.detail_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=0)
        self.detail_frame.grid_columnconfigure(0, weight=1)
        self.detail_frame.grid_rowconfigure(0, weight=1) # Сделать так, чтобы вложенные фреймы могли растягиваться


        # Фрейм для деталей договора
        self.contract_detail_frame = ctk.CTkFrame(self.detail_frame)
        self.contract_detail_frame.grid(row=0, column=0, sticky="nsew") # Используем grid для размещения
        self.contract_detail_frame.grid_columnconfigure(0, weight=0) # Метки не растягиваются
        self.contract_detail_frame.grid_columnconfigure(1, weight=1) # Поля ввода растягиваются
        self.contract_detail_frame.grid_columnconfigure(2, weight=0) # Кнопки (если есть) не растягиваются


        # Стандартные поля договора (отображаются всегда при выборе договора)
        self._add_detail_row(self.contract_detail_frame, FieldConstants.CONTRACT_NAME, self.app_data.contract_default_template_name) # Поле для названия договора
        self._add_detail_row(self.contract_detail_frame, FieldConstants.TEMPLATE_PATH, self.app_data.template_path, is_file_path=True) # Поле для пути к шаблону
        self._add_detail_row(self.contract_detail_frame, FieldConstants.OUTPUT_DIR, self.app_data.output_dir, is_folder_path=True) # Поле для папки сохранения


        # Фрейм для динамических полей договора (добавляются из FieldConstants)
        self.dynamic_fields_frame = ctk.CTkFrame(self.contract_detail_frame, fg_color="transparent")
        self.dynamic_fields_frame.grid(row=3, column=0, columnspan=2, sticky="ew") # Располагаем ниже стандартных полей
        self.dynamic_fields_frame.grid_columnconfigure(0, weight=0) # Метки не растягиваются
        self.dynamic_fields_frame.grid_columnconfigure(1, weight=1) # Поля ввода растягиваются

        # TODO: Добавить кнопки для управления полями (добавить/удалить динамические поля)


        # Фрейм для деталей мероприятия (пока заглушка)
        self.event_detail_frame = ctk.CTkFrame(self.detail_frame)
        self.event_detail_label = ctk.CTkLabel(self.event_detail_frame, text="Детали мероприятия (в разработке)")
        self.event_detail_label.pack(pady=20)
         # TODO: Добавить здесь UI для деталей мероприятия, списка договоров в мероприятии и т.д.


        # Кнопки действий (Генерировать)
        self.action_buttons_frame = ctk.CTkFrame(self.right_frame, fg_color="transparent")
        self.action_buttons_frame.grid(row=2, column=0, padx=10, pady=10, sticky="ew")
        self.action_buttons_frame.grid_columnconfigure(0, weight=1)
        self.action_buttons_frame.grid_columnconfigure(1, weight=1)


        self.generate_single_button = ctk.CTkButton(self.action_buttons_frame, text="Сгенерировать (Текущий договор)", command=self._generate_single_contract_from_gui)
        self.generate_single_button.grid(row=0, column=0, padx=5, sticky="ew")

        self.generate_event_button = ctk.CTkButton(self.action_buttons_frame, text="Сгенерировать (Захід)", command=self._generate_event_contracts_from_gui, state="disabled") # Пока отключено
        self.generate_event_button.grid(row=0, column=1, padx=5, sticky="ew")


        # Загрузка данных при старте
        self._populate_lists()
        self._update_detail_frame() # Изначально скрываем фреймы деталей


    def _apply_appearance_mode(self, light_color, dark_color):
         """Помощник для получения цвета в зависимости от темы."""
         if ctk.get_appearance_mode() == "Light":
             return light_color
         else:
             return dark_color

    def _add_detail_row(self, parent_frame, label_text, initial_value="", is_file_path=False, is_folder_path=False):
        """Добавляет строку с меткой и полем ввода в указанный фрейм."""
        row_num = parent_frame.grid_size()[1] # Определяем номер следующей строки

        label = ctk.CTkLabel(parent_frame, text=label_text + ":")
        label.grid(row=row_num, column=0, padx=5, pady=5, sticky="w")

        entry = ctk.CTkEntry(parent_frame)
        entry.insert(0, str(initial_value)) # Вставляем начальное значение
        entry.grid(row=row_num, column=1, padx=5, pady=5, sticky="ew")

        # Привязываем событие изменения текста к полю ввода
        # Используем <KeyRelease> для реакции на ввод пользователя
        # Или <FocusOut> для реакции на потерю фокуса
        # Или trace_add для более точного отслеживания изменений
        # Проще использовать <KeyRelease> или <FocusOut> для начала.
        # Используем FocusOut, чтобы не сохранять на каждое нажатие клавиши
        entry.bind("<FocusOut>", lambda event, name=label_text, entry_widget=entry: self._on_field_value_change(name, entry_widget.get()))


        if is_file_path:
             # Добавляем кнопку для выбора файла
             button = ctk.CTkButton(parent_frame, text="Выбрать файл...", width=30, command=lambda entry_widget=entry, name=label_text: self._select_file_path(entry_widget, name))
             button.grid(row=row_num, column=2, padx=5, pady=5, sticky="w")
        elif is_folder_path:
             # Добавляем кнопку для выбора папки
             button = ctk.CTkButton(parent_frame, text="Выбрать папку...", width=30, command=lambda entry_widget=entry, name=label_text: self._select_folder_path(entry_widget, name))
             button.grid(row=row_num, column=2, padx=5, pady=5, sticky="w")
        else:
             # Если это динамическое поле (не путь), сохраняем ссылку на Entry
             # для доступа к его значению при генерации
             # Только динамические поля (которые не стандартные)
             if label_text not in [FieldConstants.CONTRACT_NAME, FieldConstants.TEMPLATE_PATH, FieldConstants.OUTPUT_DIR]:
                  # Привязываем к событию изменения значения, но сохраняем в отдельный словарь
                  entry.bind("<FocusOut>", lambda event, name=label_text, entry_widget=entry: self._on_field_value_change(name, entry_widget.get()))
                  # Сохраняем ссылку на виджет Entry для этого поля
                  # self.dynamic_field_entries[label_text] = entry # Нам больше не нужен этот словарь напрямую, т.к. сохраняем в app_data через on_field_value_change


        return entry # Возвращаем ссылку на Entry виджет, если нужно


    def _select_file_path(self, entry_widget, field_name):
        """Открывает диалог выбора файла и обновляет поле ввода."""
        file_path = tk.filedialog.askopenfilename(
            title=f"Выберите файл для '{field_name}'",
            filetypes=(("Word Documents", "*.docm *.docx"), ("All files", "*.*"))
        )
        if file_path:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, file_path)
            self._on_field_value_change(field_name, file_path) # Сохраняем выбранный путь


    def _select_folder_path(self, entry_widget, field_name):
        """Открывает диалог выбора папки и обновляет поле ввода."""
        folder_path = tk.filedialog.askdirectory(
            title=f"Выберите папку для '{field_name}'"
        )
        if folder_path:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, folder_path)
            self._on_field_value_change(field_name, folder_path) # Сохраняем выбранный путь


    def _populate_lists(self):
        """Заполняет Listbox'ы данными из app_data."""
        self.event_listbox.delete(0, tk.END)
        self.contract_listbox.delete(0, tk.END)

        # Сначала мероприятия
        event_names = list(self.app_data.events.keys())
        for name in event_names:
            self.event_listbox.insert(tk.END, name)

        # Затем договоры
        contract_names = [c.name for c in self.app_data.contracts.values()]
        for name in contract_names:
            self.contract_listbox.insert(tk.END, name)

        # Сбрасываем выбранные элементы
        self.event_listbox.selection_clear(0, tk.END)
        self.contract_listbox.selection_clear(0, tk.END)


    def _update_detail_frame(self):
        """Показывает нужный фрейм деталей в зависимости от выбора."""
        # Сначала скрываем все фреймы деталей
        self.contract_detail_frame.grid_forget()
        self.event_detail_frame.grid_forget()
        self.generate_single_button.grid_forget()
        self.generate_event_button.grid_forget()


        if self.selected_contract:
            self.right_panel_title.configure(text=f"Детали договора: {self.selected_contract.name}")
            self.contract_detail_frame.grid(row=0, column=0, sticky="nsew")
            self.generate_single_button.grid(row=0, column=0, padx=5, sticky="ew") # Показываем кнопку для одного договора
            self.generate_event_button.grid_forget() # Скрываем кнопку для мероприятия
            self.generate_single_button.configure(state="normal")

            # Обновляем стандартные поля
            # Находим виджет Entry по его метке и обновляем
            for widget in self.contract_detail_frame.winfo_children():
                if isinstance(widget, ctk.CTkLabel):
                    label_text = widget.cget("text").replace(":", "") # Убираем двоеточие
                    entry_widget = widget.grid_slaves(row=widget.grid_info()["row"], column=1)[0] # Находим Entry в той же строке

                    if label_text == FieldConstants.CONTRACT_NAME:
                        entry_widget.delete(0, tk.END)
                        entry_widget.insert(0, self.selected_contract.name)
                    elif label_text == FieldConstants.TEMPLATE_PATH:
                        entry_widget.delete(0, tk.END)
                        entry_widget.insert(0, self.app_data.template_path) # Используем глобальный путь из app_data
                    elif label_text == FieldConstants.OUTPUT_DIR:
                        entry_widget.delete(0, tk.END)
                        entry_widget.insert(0, self.app_data.output_dir) # Используем глобальный путь из app_data

            # Очищаем и добавляем динамические поля
            for widget in self.dynamic_fields_frame.winfo_children():
                 widget.destroy() # Удаляем все предыдущие динамические поля

            # Находим определение полей для этого типа договора (по имени шаблона или по ID?)
            # Сейчас предполагаем, что поля определяются на основе того,
            # какие поля были сохранены для конкретного объекта договора.
            # TODO: Возможно, нужно будет определять поля на основе выбранного ШАБЛОНА

            # Заполняем dynamic_fields_frame полями из self.selected_contract.fields
            # self.selected_contract_fields уже должен быть синхронизирован с self.selected_contract.fields
            # при выборе контракта в _on_contract_select
            for field_name, field_value in self.selected_contract.fields.items():
                 # Не добавляем стандартные поля как динамические
                 if field_name not in [FieldConstants.CONTRACT_NAME, FieldConstants.TEMPLATE_PATH, FieldConstants.OUTPUT_DIR]:
                      self._add_detail_row(self.dynamic_fields_frame, field_name, field_value) # Добавляем динамическое поле
            # TODO: Добавить кнопку "Добавить поле" для добавления новых полей


        elif self.selected_event:
            self.right_panel_title.configure(text=f"Детали мероприятия: {self.selected_event.name}")
            self.event_detail_frame.grid(row=0, column=0, sticky="nsew")
            self.generate_single_button.grid_forget() # Скрываем кнопку для одного договора
            self.generate_event_button.grid(row=0, column=1, padx=5, sticky="ew") # Показываем кнопку для мероприятия
            # TODO: Включить кнопку генерации мероприятия, когда логика будет готова
            self.generate_event_button.configure(state="disabled") # Пока всегда отключено


        else:
            self.right_panel_title.configure(text="Выберите договор или мероприятие")


    def _on_event_select(self, event=None):
        """Обработчик выбора мероприятия в списке."""
        selected_indices = self.event_listbox.curselection()
        if selected_indices:
            index = selected_indices[0]
            event_name = self.event_listbox.get(index)
            self.selected_event = self.app_data.get_event_by_name(event_name)
            self.selected_contract = None # Сбрасываем выбранный договор

            # Сбрасываем выбор в списке договоров
            self.contract_listbox.selection_clear(0, tk.END)

            # Обновляем правую панель
            self._update_detail_frame()
        # else:
             # print("DEBUG: _on_event_select - Nothing selected") # Для дебагу


    def _on_contract_select(self, event=None):
        """Обработчик выбора договора в списке."""
        selected_indices = self.contract_listbox.curselection()
        if selected_indices:
            index = selected_indices[0]
            contract_name = self.contract_listbox.get(index)
            # Находим объект договора по имени
            contract_id = None
            for c_id, contract in self.app_data.contracts.items():
                 if contract.name == contract_name:
                      contract_id = c_id
                      break

            if contract_id:
                 self.selected_contract = self.app_data.get_contract(contract_id)
                 self.selected_event = None # Сбрасываем выбранное мероприятие

                 # Сбрасываем выбор в списке мероприятий
                 self.event_listbox.selection_clear(0, tk.END)

                 # !!! ДОБАВЛЕНО: Копируем поля выбранного договора в temporary dict для GUI !!!
                 # Это нужно, чтобы GUI поля были заполнены данными из выбранного договора
                 self.selected_contract_fields = self.selected_contract.fields.copy()
                 print(f"DEBUG MAIN_APP: Loaded {len(self.selected_contract_fields)} fields from contract '{contract_name}' into GUI buffer.")
                 # End of ADDED section

                 # Обновляем правую панель
                 self._update_detail_frame()
            # else:
                 # print(f"DEBUG: _on_contract_select - Contract with name '{contract_name}' not found in data.") # Для дебагу
        # else:
             # print("DEBUG: _on_contract_select - Nothing selected") # Для дебагу


    def _on_field_value_change(self, field_name, value):
        """Обробник зміни значення поля договору в GUI."""
        # Этот метод вызывается при изменении значения в Entry виджетах
        # Он должен обновить временный словарь и сохранить данные в app_data
        if self.selected_contract:
            print(f"DEBUG MAIN_APP: Field value changed for '{field_name}' = '{value}'")
            # Обновляем значение в словаре, который используется для GUI
            self.selected_contract_fields[field_name] = value

            # !!! ДОБАВЛЕНО: Обновляем значение в самом объекте договора в app_data !!!
            # Это гарантирует, что при сохранении или генерации (особенно для мероприятий)
            # используются актуальные данные.
            contract_in_data = self.app_data.get_contract(self.selected_contract.id)
            if contract_in_data:
                 contract_in_data.fields[field_name] = value
                 print(f"DEBUG MAIN_APP: Updated field '{field_name}' on contract object in app_data.")
            # End of ADDED section

            # Сохраняем все данные приложения, чтобы изменения не потерялись
            self.app_data.save_data()
            print("DEBUG MAIN_APP: Application data saved after field change.")
        # else:
            # print(f"DEBUG MAIN_APP: Field value changed for '{field_name}' but no contract selected.") # For debug


    def _add_contract(self):
        """Добавляет новый пустой договор."""
        new_contract_name = ctk.CTkInputDialog(text="Введите название нового договора:", title="Новый договор").get_input()
        if new_contract_name:
            # Создаем новый объект договора. Поля пока будут пустыми.
            # TODO: Возможно, при создании договора предлагать выбрать шаблон?
            new_contract = self.app_data.create_contract(new_contract_name)
            # self.app_data.add_contract(new_contract) # create_contract уже добавляет его

            # Сохраняем данные после добавления
            self.app_data.save_data()
            print(f"DEBUG MAIN_APP: Added new contract: {new_contract_name}")

            # Обновляем список договоров в GUI
            self._populate_lists()

            # TODO: Возможно, сразу выбрать новый договор после добавления
            # Находим индекс нового договора и выбираем его
            # try:
            #     index = list(self.app_data.contracts.keys()).index(new_contract.id)
            #     self.contract_listbox.selection_set(index)
            #     self._on_contract_select() # Вызываем обработчик выбора
            # except ValueError:
            #     pass # Если не нашли (не должно происходить)

    def _delete_contract(self):
        """Удаляет выбранный договор."""
        selected_indices = self.contract_listbox.curselection()
        if selected_indices:
            index = selected_indices[0]
            contract_name_to_delete = self.contract_listbox.get(index)

            confirm = messagebox.askyesno("Удалить договор", f"Вы уверены, что хотите удалить договор '{contract_name_to_delete}'?")
            if confirm:
                # Находим ID договора по имени
                contract_id_to_delete = None
                for c_id, contract in self.app_data.contracts.items():
                     if contract.name == contract_name_to_delete:
                          contract_id_to_delete = c_id
                          break

                if contract_id_to_delete:
                    self.app_data.delete_contract(contract_id_to_delete)
                    self.selected_contract = None # Сбрасываем выбранный договор
                    self.selected_contract_fields = {} # Очищаем поля

                    # Сохраняем данные после удаления
                    self.app_data.save_data()
                    print(f"DEBUG MAIN_APP: Deleted contract: {contract_name_to_delete}")

                    # Обновляем список и детали
                    self._populate_lists()
                    self._update_detail_frame()
                # else:
                    # print(f"DEBUG MAIN_APP: Attempted to delete contract '{contract_name_to_delete}' but ID not found.") # Для дебагу


    # def _add_event(self): # TODO: Реализовать добавление мероприятий
    #     pass
    #
    # def _delete_event(self): # TODO: Реализовать удаление мероприятий
    #      pass


    def _generate_single_contract_from_gui(self):
        """
        Запускает генерацию документа для текущего выбранного договора,
        используя данные из полей GUI и глобальные настройки путей.
        """
        if self.selected_contract:
            # Получаем актуальный объект договора.
            # Поля уже должны быть обновлены через _on_field_value_change
            # Но на всякий случай, можно еще раз обновить из GUI перед вызовом генератора
            # Хотя лучше полагаться на _on_field_value_change для поддержания актуальности
            # Копирование из self.selected_contract_fields в contract.fields делается
            # внутри generate_document_win32, но лучше делать это здесь или
            # убедиться, что selected_contract_fields = selected_contract.fields
            # при выборе. Сейчас выбранный контракт напрямую ссылается на объект в app_data,
            # и _on_field_value_change обновляет этот объект.

            # Получаем актуальные пути из GUI (они сохраняются в app_data через on_field_value_change)
            template_path_from_gui = self.selected_contract_fields.get(FieldConstants.TEMPLATE_PATH, self.app_data.template_path)
            output_dir_from_gui = self.selected_contract_fields.get(FieldConstants.OUTPUT_DIR, self.app_data.output_dir)

            # Используем объект договора из self.selected_contract, который уже должен содержать
            # актуальные данные полей благодаря _on_field_value_change
            contract_for_generation = self.selected_contract

            # !!! ДОБАВЛЕНО: Отладочный вывод объекта договора и ключей полей ПЕРЕД вызовом генератора !!!
            print(f"DEBUG MAIN_APP: Preparing to generate single contract '{contract_for_generation.name}'")
            print(f"DEBUG MAIN_APP: Template path for generation: '{template_path_from_gui}'")
            print(f"DEBUG MAIN_APP: Output directory for generation: '{output_dir_from_gui}'")
            print(f"DEBUG MAIN_APP: Contract object fields keys BEFORE generator call: {list(contract_for_generation.fields.keys())}")
            print(f"DEBUG MAIN_APP: Contract object fields BEFORE generator call: {contract_for_generation.fields}") # Print the whole fields dict for full debug
            # End of ADDED section


            # Вызываем функцию генерации документа
            generated_file_path = generate_document_win32(
                 contract_for_generation, # Передаем объект договора с данными
                 template_path_from_gui,
                 output_dir_from_gui
            )

            if generated_file_path:
                 messagebox.showinfo("Генерація завершена", f"Документ успішно сгенеровано:\n{generated_file_path}")
            # else:
                 # Ошибка уже логируется и показывается в generate_document_win32


        else:
            messagebox.showwarning("Виберіть договір", "Будь ласка, виберіть договір для генерації.")


    def _generate_event_contracts_from_gui(self):
        """Запускает генерацию документов для всех договоров в выбранном мероприятии."""
        if self.selected_event:
            # Получаем актуальные пути из GUI (они сохраняются в app_data)
            template_path_from_gui = self.selected_contract_fields.get(FieldConstants.TEMPLATE_PATH, self.app_data.template_path) # Поля путей берутся из общих настроек GUI
            output_dir_from_gui = self.selected_contract_fields.get(FieldConstants.OUTPUT_DIR, self.app_data.output_dir)

            # Перебираем договоры, связанные с этим мероприятием
            contracts_to_generate = list(self.selected_event.contracts.values()) # Получаем список объектов договоров

            if not contracts_to_generate:
                 messagebox.showinfo("Генерація документів", f"У заході '{self.selected_event.name}' немає договорів для генерації.")
                 return

            # TODO: Добавить индикатор прогресса для генерации по мероприятию

            generated_count = 0
            failed_count = 0

            # Перебираем все договоры в рамках мероприятия
            for contract_in_event in contracts_to_generate:
                 # Важно! Для генерации по мероприятию нужно брать актуальные данные полей
                 # для КАЖДОГО договора. Эти данные должны быть сохранены в app_data.
                 # _on_field_value_change должен сохранять поля в конкретный объект Contract
                 # внутри app_data.events[event_id].contracts[contract_id].fields

                 # Получаем актуальный объект договора из app_data перед генерацией
                 # Это гарантирует, что мы используем последние сохраненные данные
                 actual_contract_from_data = self.app_data.get_contract(contract_in_event.id)

                 if not actual_contract_from_data:
                      print(f"WARNING MAIN_APP: Contract with ID {contract_in_event.id} not found in app_data for event generation. Skipping.")
                      error_handling.log_error(None, None, None, level="WARNING", message=f"Договір ID {contract_in_event.id} не знайдено в даних для генерації заходу.")
                      failed_count += 1
                      continue # Пропускаем этот договор

                 # !!! Теперь actual_contract_from_data.fields должен содержать сохраненные данные !!!

                 # !!! ДОБАВЛЕНО: Отладочный вывод объекта договора и ключей полей ПЕРЕД вызовом генератора (для мероприятия) !!!
                 print(f"DEBUG MAIN_APP: Preparing to generate event contract '{actual_contract_from_data.name}' (ID: {actual_contract_from_data.id})")
                 print(f"DEBUG MAIN_APP: Template path for generation: '{template_path_from_gui}'")
                 print(f"DEBUG MAIN_APP: Output directory for generation: '{output_dir_from_gui}'")
                 print(f"DEBUG MAIN_APP: Contract object fields keys BEFORE generator call (Event): {list(actual_contract_from_data.fields.keys())}")
                 print(f"DEBUG MAIN_APP: Contract object fields BEFORE generator call (Event): {actual_contract_from_data.fields}") # Print the whole fields dict for full debug
                 # End of ADDED section


                 # Вызываем функцию генерации документа
                 # Передаем актуальный объект договора actual_contract_from_data
                 generated_file_path = generate_document_win32(
                      actual_contract_from_data, # Передаем объект договора с данными
                      template_path_from_gui,
                      output_dir_from_gui
                 )

                 if generated_file_path:
                     generated_count += 1
                 else:
                     failed_count += 1
                     # Ошибка уже логируется и показывается в generate_document_win32


            # Показываем сводку по завершении генерации всех договоров
            messagebox.showinfo("Генерація завершена", f"Завершено генерацію для заходу '{self.selected_event.name}'.\nУспішно згенеровано: {generated_count}\nПомилки генерації: {failed_count}")


        else:
            messagebox.showwarning("Виберіть захід", "Будь ласка, виберіть захід для генерації договорів.")


    # Метод для сброса выбранных элементов в GUI списках (нужен после добавления/удаления)
    def _reset_selection(self):
        self.selected_contract = None
        self.selected_event = None
        self.selected_contract_fields = {} # Очищаем буфер полей
        self.event_listbox.selection_clear(0, tk.END)
        self.contract_listbox.selection_clear(0, tk.END)
        self._update_detail_frame() # Обновляем правую панель


    # Helper method to get the correct color based on appearance mode (already exists)
    # def _apply_appearance_mode(self, light_color, dark_color):
    #      if ctk.get_appearance_mode() == "Light":
    #          return light_color
    #      else:
    #          return dark_color


if __name__ == "__main__":
    # Устанавливаем тему customtkinter
    ctk.set_appearance_mode("System") # Modes: "System" (default), "Light", "Dark"
    ctk.set_default_color_theme("blue") # Themes: "blue" (default), "green", "dark-blue"

    app = MainApplication()
    app.mainloop()