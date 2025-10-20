# sportforall/gui.py

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import sys
import re
import uuid
import datetime # Додаємо для можливого використання дат

# Імпортуємо власні модулі
from sportforall import constants
from sportforall import error_handling
from sportforall.models import AppData, Event, Contract, Item
from sportforall import document_processor
from sportforall import excel_logger
from sportforall import number_to_text_ua


# --- Допоміжний клас для скрольованого фрейму ---
class ScrolledFrame(ttk.Frame):
    """
    Фрейм з вертикальним скролбаром.
    Розміщуйте вміст всередині 'self.inner_frame'.
    """
    def __init__(self, master=None, **kwargs):
        super().__init__(master, **kwargs)

        self.canvas = tk.Canvas(self)
        self.vscrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vscrollbar.set)

        self.vscrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        # Створюємо внутрішній фрейм, який буде скролитись
        self.inner_frame = ttk.Frame(self.canvas)
        # Розміщуємо внутрішній фрейм у вікні канвасу
        self._canvas_window_id = self.canvas.create_window((0, 0), window=self.inner_frame, anchor="nw")

        # Прив'язуємо події зміни розміру фрейму та канвасу для оновлення області прокрутки
        self.inner_frame.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        # Підтримка скролінгу колесом миші
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind_all("<Button-4>", self._on_mousewheel) # Scroll up
        self.canvas.bind_all("<Button-5>", self._on_mousewheel) # Scroll down


    def _on_frame_configure(self, event=None):
        """Оновлює область прокрутки канвасу при зміні розміру внутрішнього фрейму."""
        # Додаємо невеликий відступ знизу, щоб останній елемент не був перекритий
        bbox = self.canvas.bbox("all")
        if bbox:
             x1, y1, x2, y2 = bbox
             self.canvas.config(scrollregion=(x1, y1, x2, y2 + 10)) # Додаємо 10 пікселів відступу


    def _on_canvas_configure(self, event=None):
        """Оновлює ширину внутрішнього фрейму при зміні ширини канвасу."""
        canvas_width = event.width
        # Налаштовуємо ширину вікна в канвасі, щоб внутрішній фрейм розтягувався
        self.canvas.itemconfigure(self._canvas_window_id, width=canvas_width)


    def _on_mousewheel(self, event):
        """Обробляє скролінг колесом миші."""
        widget_at_mouse = event.widget.winfo_containing(event.x_root, event.y_root)
        if widget_at_mouse is not None:
             is_within_frame = False
             current = widget_at_mouse
             while current is not None:
                  if current == self.canvas or current == self.inner_frame:
                       is_within_frame = True
                       break
                  current = current.master

             if current == self.vscrollbar:
                  is_within_frame = False


             if is_within_frame:
                if sys.platform.startswith('win'):
                    self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
                elif sys.platform.startswith('linux'):
                    if event.num == 4:
                        self.canvas.yview_scroll(-1, "units")
                    elif event.num == 5:
                        self.canvas.yview_scroll(1, "units")


# --- Контекстне меню для полів введення ---
class TextEntryMenu(tk.Menu):
    """
    Контекстне меню (ПКМ) для стандартних полів введення тексту (Entry, Text, ScrolledText).
    """
    def __init__(self, master=None, **kwargs):
        super().__init__(master, tearoff=0, **kwargs)
        self.add_command(label="Вирізати", command=self.cut_text)
        self.add_command(label="Копіювати", command=self.copy_text)
        self.add_command(label="Вставити", command=self.paste_text)
        self.add_separator()
        self.add_command(label="Виділити все", command=self.select_all_text)

        self.active_widget = None

    def show(self, event):
        widget_at_mouse = event.widget.winfo_containing(event.x_root, event.y_root)
        if widget_at_mouse is not None and isinstance(widget_at_mouse, (ttk.Entry, tk.Entry, tk.Text, scrolledtext.ScrolledText)):
            self.active_widget = widget_at_mouse
            try:
                 state_cut_paste = "normal" if self.active_widget.cget('state') != 'readonly' else "disabled"
                 self.entryconfig("Вирізати", state=state_cut_paste)
                 self.entryconfig("Вставити", state=state_cut_paste)
                 self.entryconfig("Копіювати", state="normal")
                 self.entryconfig("Виділити все", state="normal")

                 self.tk_popup(event.x_root, event.y_root)
            finally:
                self.grab_release()
        else:
             self.active_widget = None


    def cut_text(self):
        if self.active_widget:
             try:
                 if self.active_widget.cget('state') != 'readonly':
                      self.active_widget.event_generate("<<Cut>>")
             except tk.TclError: pass

    def copy_text(self):
         if self.active_widget:
             try:
                 self.active_widget.event_generate("<<Copy>>")
             except tk.TclError: pass

    def paste_text(self):
        if self.active_widget:
            try:
                 if self.active_widget.cget('state') != 'readonly':
                      self.active_widget.event_generate("<<Paste>>")
            except tk.TclError: pass

    def select_all_text(self):
        if self.active_widget:
            if isinstance(self.active_widget, (ttk.Entry, tk.Entry)):
                 self.active_widget.selection_range(0, tk.END)
                 self.active_widget.icursor(tk.END)
            elif isinstance(self.active_widget, (tk.Text, scrolledtext.ScrolledText)):
                 self.active_widget.tag_add(tk.SEL, "1.0", tk.END)
                 self.active_widget.mark_set(tk.INSERT, tk.END)
            self.active_widget.focus_set()


class AppGUI(ttk.Frame):
    """
    Головний клас графічного інтерфейсу користувача.
    Наслідується від ttk.Frame і містить всі елементи управління.
    """
    def __init__(self, master=None, app_data: AppData = None):
        """
        Ініціалізує головний інтерфейс.

        Args:
            master: Батьківський віджет (зазвичай вікно tk.Tk).
            app_data: Об'єкт AppData з даними додатка.
        """
        # Цей рядок викликає конструктор батьківського класу (ttk.Frame)
        # Передаємо йому головне вікно (master) як батьківський віджет.
        # НЕ потрібно передавати app_data або інші аргументи сюди.
        super().__init__(master)

        # Зберігаємо посилання на батьківське вікно та дані додатка в атрибутах об'єкта
        self.master = master
        self.app_data = app_data
        self.current_event: Event = None # Змінна для відстеження поточного обраного заходу
        self.current_contract: Contract = None # Змінна для відстеження поточного обраного договору

        # Створюємо контекстне меню для полів введення
        self.text_menu = TextEntryMenu(self.master)


        # Словники для зберігання посилань на Tkinter змінні та віджети
        self._field_vars = {} # {placeholder_key: tk.StringVar}
        self._item_vars = {} # {item_id: {field_name: tk.Variable}}
        self._item_frames = {} # {item_id: ttk.Frame} # Зберігаємо фрейми рядків товарів для їх видалення

        # Tkinter змінні для полів суми (ініціалізуються тут, оновлюються в calculate_total_sum)
        self.total_sum_var = tk.StringVar()
        self.text_sum_var = tk.StringVar()


        self.create_widgets()
        self.update_gui_from_data()


    def create_widgets(self):
        """
        Створює всі віджети головного вікна GUI.
        """
        # Налаштовуємо сітку головного фрейму
        self.columnconfigure(0, weight=1) # Колонка для лівої панелі (вибір файлів, заходи)
        self.columnconfigure(1, weight=3) # Колонка для правої панелі (деталі заходу/договору)
        self.rowconfigure(1, weight=1)    # Рядок для основного контенту (списки та деталі)

        # --- Панель верхніх налаштувань (вибір файлів) ---
        settings_frame = ttk.LabelFrame(self, text="Налаштування файлів")
        settings_frame.grid(row=0, column=0, columnspan=2, sticky="nsew", padx=5, pady=5)
        settings_frame.columnconfigure(0, weight=1) # Для шляху шаблону
        settings_frame.columnconfigure(1, weight=0) # Для кнопки шаблону
        settings_frame.columnconfigure(2, weight=1) # Для шляху папки збереження
        settings_frame.columnconfigure(3, weight=0) # Для кнопки папки збереження

        # Вибір файлу шаблону
        ttk.Label(settings_frame, text="Шлях до шаблону (.docm):").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.template_path_var = tk.StringVar()
        # Використовуємо Entry, але робимо його тільки для читання (state='readonly')
        self.template_path_entry = ttk.Entry(settings_frame, textvariable=self.template_path_var, state='readonly')
        self.template_path_entry.grid(row=1, column=0, sticky="ew", padx=5, pady=2)
        ttk.Button(settings_frame, text="Обрати шаблон", command=self.select_template_file).grid(row=1, column=1, sticky="e", padx=5, pady=2)

        # Вибір папки для збереження
        ttk.Label(settings_frame, text="Папка для збереження договорів:").grid(row=0, column=2, sticky="w", padx=5, pady=2)
        self.output_dir_var = tk.StringVar()
         # Використовуємо Entry, але робимо його тільки для читання (state='readonly')
        self.output_dir_entry = ttk.Entry(settings_frame, textvariable=self.output_dir_var, state='readonly')
        self.output_dir_entry.grid(row=1, column=2, sticky="ew", padx=5, pady=2)
        ttk.Button(settings_frame, text="Обрати папку", command=self.select_output_directory).grid(row=1, column=3, sticky="e", padx=5, pady=2)

        # --- Ліва панель: Список заходів ---
        events_panel = ttk.Frame(self)
        events_panel.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        events_panel.rowconfigure(1, weight=1) # Список заходів
        events_panel.columnconfigure(0, weight=1)

        ttk.Label(events_panel, text="Заходи:").grid(row=0, column=0, sticky="w", padx=5, pady=2)

        # Список заходів (TreeView)
        # Використовуємо TreeView як список
        # ### ПОПРАВКА ТУТ: Використовуємо внутрішній ідентифікатор колонки 'event_name_col'
        self.events_tree = ttk.Treeview(events_panel, columns=("event_name_col",), show="headings")
        # ### ПОПРАВКА ТУТ: Налаштовуємо заголовок для колонки з ідентифікатором 'event_name_col'
        self.events_tree.heading("event_name_col", text="Назва Заходу")
        # ### ПОПРАВКА ТУТ: Налаштовуємо параметри колонки з ідентифікатором 'event_name_col'
        self.events_tree.column("event_name_col", width=150, anchor="w")
        self.events_tree.grid(row=1, column=0, sticky="nsew", padx=5, pady=2)

        # Скролбар для списку заходів
        events_scrollbar = ttk.Scrollbar(events_panel, orient="vertical", command=self.events_tree.yview)
        self.events_tree.configure(yscrollcommand=events_scrollbar.set)
        events_scrollbar.grid(row=1, column=1, sticky="ns")

        # Кнопка для додавання нового заходу
        ttk.Button(events_panel, text="Додати Захід", command=self.add_event).grid(row=2, column=0, sticky="ew", padx=5, pady=5)

        # Прив'язка події вибору елемента у списку заходів
        self.events_tree.bind("<<TreeviewSelect>>", self.on_event_select)

        # --- Права панель: Деталі заходу та список договорів ---
        # Використовуємо Notebook для перемикання між виглядом списку договорів
        # та виглядом деталей конкретного договору.
        self.details_notebook = ttk.Notebook(self)
        self.details_notebook.grid(row=1, column=1, sticky="nsew", padx=5, pady=5)

        # Вкладка для списку договорів обраного заходу
        self.contracts_list_frame = ttk.Frame(self.details_notebook)
        self.details_notebook.add(self.contracts_list_frame, text="Договори Заходу")
        self.contracts_list_frame.columnconfigure(0, weight=1)
        self.contracts_list_frame.rowconfigure(1, weight=1)

        ttk.Label(self.contracts_list_frame, text="Договори для обраного заходу:").grid(row=0, column=0, sticky="w", padx=5, pady=2)

        # Список договорів для обраного заходу (TreeView)
        # ### ПОПРАВКА ТУТ: Використовуємо внутрішній ідентифікатор колонки 'contract_name_col'
        self.contracts_tree = ttk.Treeview(self.contracts_list_frame, columns=("contract_name_col",), show="headings")
        # ### ПОПРАВКА ТУТ: Налаштовуємо заголовок для колонки з ідентифікатором 'contract_name_col'
        self.contracts_tree.heading("contract_name_col", text="Назва Договору")
        # ### ПОПРАВКА ТУТ: Налаштовуємо параметри колонки з ідентифікатором 'contract_name_col'
        self.contracts_tree.column("contract_name_col", width=200, anchor="w")
        self.contracts_tree.grid(row=1, column=0, sticky="nsew", padx=5, pady=2)

        # Скролбар для списку договорів
        contracts_scrollbar = ttk.Scrollbar(self.contracts_list_frame, orient="vertical", command=self.contracts_tree.yview)
        self.contracts_tree.configure(yscrollcommand=contracts_scrollbar.set)
        contracts_scrollbar.grid(row=1, column=1, sticky="ns")

        # Кнопки для роботи зі списком договорів
        contracts_buttons_frame = ttk.Frame(self.contracts_list_frame)
        contracts_buttons_frame.grid(row=2, column=0, sticky="ew", padx=5, pady=5)
        contracts_buttons_frame.columnconfigure(0, weight=1)
        contracts_buttons_frame.columnconfigure(1, weight=1)
        contracts_buttons_frame.columnconfigure(2, weight=1)
        contracts_buttons_frame.columnconfigure(3, weight=1)

        ttk.Button(contracts_buttons_frame, text="Додати Договір", command=self.add_contract).grid(row=0, column=0, sticky="ew", padx=2)
        ttk.Button(contracts_buttons_frame, text="Видалити Договір", command=self.delete_contract).grid(row=0, column=1, sticky="ew", padx=2)
        ttk.Button(contracts_buttons_frame, text="Замінити Шаблон", command=self.replace_contract_template).grid(row=0, column=2, sticky="ew", padx=2)
        ttk.Button(contracts_buttons_frame, text="Генерувати Захід", command=self.generate_event_contracts).grid(row=0, column=3, sticky="ew", padx=2)

        # Прив'язка події вибору елемента у списку договорів
        # ### ПОПРАВКА ТУТ: Прив'язуємо подію до методу on_contract_select
        self.contracts_tree.bind("<<TreeviewSelect>>", self.on_contract_select)


        # Вкладка для деталей конкретного договору
        self.contract_details_frame = ttk.Frame(self.details_notebook)
        self.details_notebook.add(self.contract_details_frame, text="Деталі Договору")
        # Ховаємо вкладку деталей при старті
        self.details_notebook.hide(self.details_notebook.index(self.contract_details_frame))


        # --- Додаємо скрольований фрейм для полів деталей договору ---
        self.contract_details_scrolled_frame = ScrolledFrame(self.contract_details_frame)
        self.contract_details_scrolled_frame.pack(expand=True, fill="both", padx=5, pady=5)
        self.contract_details_scrolled_frame.inner_frame.columnconfigure(0, weight=0) # Метки (не розтягуємо)
        self.contract_details_scrolled_frame.inner_frame.columnconfigure(1, weight=1) # Поля введення (розтягуємо)


        # --- Панель стану (версія) ---
        status_bar = ttk.Frame(self)
        status_bar.grid(row=2, column=0, columnspan=2, sticky="ew", padx=5, pady=2)
        status_bar.columnconfigure(0, weight=1) # Для версії

        ttk.Label(status_bar, text=f"Версія: {constants.APP_VERSION}").grid(row=0, column=0, sticky="w")

        if self.app_data is None:
             self.app_data = AppData(version=constants.APP_VERSION)

        # Прив'язуємо контекстне меню до головного вікна
        # Це призведе до спрацьовування show() на будь-якому віджеті під курсором ПКМ
        self.master.bind("<Button-3>", self.text_menu.show)


    def update_gui_from_data(self):
        """
        Оновлює елементи GUI відповідно до даних, завантажених з self.app_data.
        """
        if self.app_data:
            self.template_path_var.set(self.app_data.template_path)
            self.output_dir_var.set(self.app_data.output_dir)

            # Очищаємо дерево заходів
            for item in self.events_tree.get_children():
                self.events_tree.delete(item)

            # Додаємо заходи з даних
            for event in self.app_data.events:
                # Використовуємо внутрішній ідентифікатор колонки для вставки значення
                self.events_tree.insert("", "end", iid=event.id, values=(event.name,))

            # Очищаємо панель деталей та список договорів
            self.clear_contract_details_frame_content()
            self.clear_contracts_list()

            # Перемикаємось на вкладку списку договорів
            if self.details_notebook.index(self.contract_details_frame) in self.details_notebook.tabs():
                 self.details_notebook.hide(self.details_notebook.index(self.contract_details_frame))
                 self.details_notebook.select(self.contracts_list_frame)


    def select_template_file(self):
        """
        Відкриває діалог вибору файлу шаблону .docm.
        Оновлює шлях у даних та GUI.
        """
        filepath = filedialog.askopenfilename(
            title="Оберіть файл шаблону .docm або .docx",
            filetypes=(("Word Documents", "*.docm *.docx"), ("All files", "*.*"))
        )
        if filepath:
            self.template_path_var.set(filepath)
            self.app_data.template_path = filepath
            print(f"Обрано файл шаблону: {filepath}")

            # TODO: Оновити плейсхолдери для існуючих договорів? Це складно.
            # Наразі, плейсхолдери визначаються при створенні нового договору.


    def select_output_directory(self):
        """
        Відкриває діалог вибору папки для збереження згенерованих договорів.
        """
        directory = filedialog.askdirectory(
            title="Оберіть папку для збереження договорів"
        )
        if directory:
            self.output_dir_var.set(directory)
            self.app_data.output_dir = directory
            print(f"Обрано папку для збереження: {directory}")

    def add_event(self):
        """
        Додає новий захід.
        """
        dialog = tk.Toplevel(self.master)
        dialog.title("Новий Захід")
        dialog.transient(self.master)
        dialog.grab_set()
        dialog.resizable(False, False)

        label = ttk.Label(dialog, text="Введіть назву нового заходу:")
        label.pack(padx=10, pady=10)

        entry = ttk.Entry(dialog, width=40)
        entry.pack(padx=10, pady=5)
        entry.focus_set()

        def on_ok():
            event_name = entry.get().strip()
            if event_name:
                new_event = Event(name=event_name)
                self.app_data.events.append(new_event)
                # Використовуємо внутрішній ідентифікатор колонки для вставки значення
                self.events_tree.insert("", "end", iid=new_event.id, values=(new_event.name,))
            dialog.destroy()

        def on_cancel():
            dialog.destroy()

        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)
        ttk.Button(button_frame, text="OK", command=on_ok).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Скасувати", command=on_cancel).pack(side="left", padx=5)

        dialog.bind('<Return>', lambda e=None: on_ok())
        dialog.bind('<Escape>', lambda e=None: on_cancel())

        self.master.wait_window(dialog)


    def on_event_select(self, event=None):
        """
        Обробник вибору заходу.
        """
        selected_items = self.events_tree.selection()
        if not selected_items:
            self.current_event = None
            self.clear_contracts_list()
            self.clear_contract_details_frame_content()
            # При виборі пустого рядка або знятті виділення, ховаємо вкладку деталей
            if self.details_notebook.index(self.contract_details_frame) in self.details_notebook.tabs():
                 self.details_notebook.hide(self.details_notebook.index(self.contract_details_frame))
                 self.details_notebook.select(self.contracts_list_frame)

            print("Нічого не обрано у списку заходів.")
            return

        selected_event_id = selected_items[0]
        self.current_event = next((e for e in self.app_data.events if e.id == selected_event_id), None)

        if self.current_event:
            print(f"Обрано захід: {self.current_event.name}")
            self.update_contracts_list(self.current_event.contracts)
            # При виборі нового заходу, очищаємо деталі попереднього договору
            self.clear_contract_details_frame_content()
            # При виборі нового заходу, перемикаємось на вкладку списку договорів
            if self.details_notebook.index(self.contract_details_frame) in self.details_notebook.tabs():
                 self.details_notebook.hide(self.details_notebook.index(self.contract_details_frame))
                 self.details_notebook.select(self.contracts_list_frame)

        else:
             error_logger.log_error(ValueError("Не знайдено об'єкт Захід для обраного ID у Treeview"), f"Не знайдено об'єкт Захід з ID: {selected_event_id}")
             self.current_event = None
             self.clear_contracts_list()
             self.clear_contract_details_frame_content()
             if self.details_notebook.index(self.contract_details_frame) in self.details_notebook.tabs():
                 self.details_notebook.hide(self.details_notebook.index(self.contract_details_frame))
                 self.details_notebook.select(self.contracts_list_frame)


    def update_contracts_list(self, contracts: list):
        """
        Оновлює список договорів у Treeview.
        """
        self.clear_contracts_list()
        for contract in contracts:
             # Використовуємо внутрішній ідентифікатор колонки для вставки значення
            self.contracts_tree.insert("", "end", iid=contract.id, values=(contract.name,))

    def clear_contracts_list(self):
        """
        Очищає Treeview зі списком договорів.
        """
        for item in self.contracts_tree.get_children():
            self.contracts_tree.delete(item)
        # Не скидаємо current_contract тут, це робиться в on_contract_select або clear_contract_details_frame_content


    def add_contract(self):
        """
        Додає новий договір до обраного заходу.
        Читає плейсхолдери з обраного шаблону та ініціалізує поля.
        """
        if not self.current_event:
            messagebox.showwarning("Попередження", "Будь ласка, спочатку оберіть захід, до якого потрібно додати договір.")
            return

        # Визначаємо шаблон для нового договору. Використовуємо головний шаблон додатка.
        template_to_use = self.app_data.template_path
        if not template_to_use or not os.path.exists(template_to_use):
            messagebox.showwarning("Попередження", "Будь ласка, оберіть файл шаблону в налаштуваннях файлів перед додаванням договору.")
            # Пропонуємо обрати шаблон, якщо він не обраний
            self.select_template_file()
            template_to_use = self.app_data.template_path # Перечитуємо після спроби вибору

        if not template_to_use or not os.path.exists(template_to_use):
             messagebox.showerror("Помилка шаблону", "Не вдалося визначити або знайти файл шаблону. Додавання договору скасовано.")
             return

        # Діалог для назви договору
        dialog = tk.Toplevel(self.master)
        dialog.title(f"Новий Договір для '{self.current_event.name}'")
        dialog.transient(self.master)
        dialog.grab_set()
        dialog.resizable(False, False)

        tk.Label(dialog, text="Введіть назву нового договору:").pack(padx=10, pady=5)
        contract_name_entry = ttk.Entry(dialog, width=50)
        contract_name_entry.pack(padx=10, pady=5)
        contract_name_entry.focus_set()

        def on_ok():
            contract_name = contract_name_entry.get().strip()
            if not contract_name:
                messagebox.showwarning("Попередження", "Назва договору не може бути порожньою.")
                return

            new_contract = Contract(name=contract_name)
            new_contract.template_path = "" # Новий договір за замовчуванням використовує головний шаблон

            try:
                 # Читання плейсхолдерів з обраного шаблону за допомогою document_processor
                 # Може виникнути помилка, якщо шаблон некоректний
                 placeholders = document_processor.find_placeholders_in_template(template_to_use)

                 # Ініціалізуємо поля договору з пустими значеннями для знайдених плейсхолдерів
                 new_contract.fields = {ph: "" for ph in placeholders}
                 # Автоматично додаємо <разом> та <сума прописью>, якщо вони не були знайдені в шаблоні
                 # (це потрібно, якщо шаблон не містить цих плейсхолдерів, але ми хочемо їх розраховувати)
                 new_contract.fields.setdefault(constants.TOTAL_SUM_PLACEHOLDER, "")
                 new_contract.fields.setdefault(constants.TEXT_SUM_PLACEHOLDER, "")


                 # Автоматично заповнюємо спільні поля з першого договору, якщо він є
                 if self.current_event.contracts:
                      first_contract = self.current_event.contracts[0]
                      for placeholder in constants.COMMON_PLACEHOLDERS:
                          # Копіюємо поле, тільки якщо воно існує в першому договорі І в новому (знайдене в шаблоні)
                          if placeholder in first_contract.fields and placeholder in new_contract.fields:
                               new_contract.fields[placeholder] = first_contract.fields[placeholder]

                 # Заповнюємо назву заходу, якщо плейсхолдер існує в новому договорі
                 if "<назва заходу>" in new_contract.fields:
                     new_contract.fields["<назва заходу>"] = self.current_event.name


                 self.current_event.contracts.append(new_contract)
                 self.update_contracts_list(self.current_event.contracts)
                 print(f"Додано новий договір '{new_contract.name}' до заходу '{self.current_event.name}'")
                 # Автоматично обираємо новостворений договір
                 self.contracts_tree.selection_set(new_contract.id)
                 # on_contract_select має викликатись автоматично при зміні вибору Treeview


            except Exception as e:
                 # Якщо сталася помилка при читанні шаблону або ініціалізації полів
                 error_logger.log_error(e, f"Помилка при створенні нового договору '{contract_name}' або читанні шаблону '{template_to_use}'")
                 messagebox.showerror("Помилка", f"Не вдалося створити договір або прочитати шаблон. Перевірте error.txt")
                 # Видаляємо створений об'єкт договору, якщо він був частково створений і доданий до списку
                 if 'new_contract' in locals() and new_contract in self.current_event.contracts:
                      self.current_event.contracts.remove(new_contract)
                      self.update_contracts_list(self.current_event.contracts)


            dialog.destroy()

        def on_cancel():
            dialog.destroy()

        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)
        ttk.Button(button_frame, text="OK", command=on_ok).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Скасувати", command=on_cancel).pack(side="left", padx=5)

        dialog.bind('<Return>', lambda e=None: on_ok())
        dialog.bind('<Escape>', lambda e=None: on_cancel())

        self.master.wait_window(dialog)


    def delete_contract(self):
        """
        Видаляє обраний договір.
        """
        if not self.current_event:
             messagebox.showwarning("Попередження", "Спочатку оберіть захід.")
             return

        selected_items = self.contracts_tree.selection()
        if not selected_items:
            messagebox.showwarning("Попередження", "Будь ласка, оберіть договір для видалення.")
            return

        selected_contract_id = selected_items[0]

        contract_to_delete = next((c for c in self.current_event.contracts if c.id == selected_contract_id), None)

        if contract_to_delete:
            if messagebox.askyesno("Підтвердження видалення", f"Ви впевнені, що хочете видалити договір '{contract_to_delete.name}'?"):
                # Видаляємо договір зі списку даних
                self.current_event.contracts.remove(contract_to_delete)
                # Оновлюємо список договорів у GUI
                self.update_contracts_list(self.current_event.contracts)
                # Очищаємо деталі договору, якщо видалили саме обраний
                if self.current_contract and self.current_contract.id == selected_contract_id:
                    self.clear_contract_details_frame_content()
                    # Перемикаємось на вкладку списку договорів після видалення обраного
                    if self.details_notebook.index(self.contract_details_frame) in self.details_notebook.tabs():
                         self.details_notebook.hide(self.details_notebook.index(self.contract_details_frame))
                         self.details_notebook.select(self.contracts_list_frame)

                print(f"Видалено договір '{contract_to_delete.name}' з заходу '{self.current_event.name}'")
        else:
             error_logger.log_error(ValueError("Не знайдено об'єкт Договір для обраного ID у Treeview"), f"Не знайдено об'єкт Договір з ID: {selected_contract_id} у заході {self.current_event.name if self.current_event else 'Немає обраного заходу'}")


    def replace_contract_template(self):
        """
        Дозволяє замінити файл шаблону для обраного договору.
        """
        # TODO: Реалізувати
        messagebox.showinfo("Функціонал у розробці", "Функція заміни шаблону договору ще не реалізована.")
        # Це складно, оскільки при зміні шаблону потрібно перечитати плейсхолдери
        # і вирішити, як поводитися з існуючими значеннями полів.


    def generate_event_contracts(self):
        """
        Генерує документи для всіх договорів у поточному заході.
        """
        if not self.current_event:
            messagebox.showwarning("Попередження", "Будь ласка, спочатку оберіть захід, договори якого потрібно згенерувати.")
            return

        if not self.current_event.contracts:
            messagebox.showinfo("Інформація", "У обраному заході немає договорів для генерації.")
            return

        # Перевіряємо наявність загального шаблону та папки збереження
        template_to_use_default = self.app_data.template_path
        if not template_to_use_default or not os.path.exists(template_to_use_default):
             messagebox.showwarning("Попередження", "Будь ласка, оберіть загальний файл шаблону в налаштуваннях файлів.")
             self.select_template_file() # Пропонуємо обрати шаблон
             template_to_use_default = self.app_data.template_path
             if not template_to_use_default or not os.path.exists(template_to_use_default):
                  messagebox.showerror("Помилка шаблону", "Не вдалося визначити або знайти файл шаблону. Генерація скасована.")
                  return

        output_dir = self.app_data.output_dir
        if not output_dir or not os.path.isdir(output_dir):
             messagebox.showwarning("Попередження", "Будь ласка, оберіть папку для збереження згенерованих договорів.")
             self.select_output_directory() # Пропонуємо обрати папку
             output_dir = self.app_data.output_dir
             if not output_dir or not os.path.isdir(output_dir):
                  messagebox.showerror("Помилка збереження", "Не вдалося визначити або знайти папку для збереження. Генерація скасована.")
                  return

        if messagebox.askyesno("Підтвердження генерації", f"Ви впевнені, що хочете згенерувати всі {len(self.current_event.contracts)} договор(и/ів) для заходу '{self.current_event.name}' у папку '{output_dir}'?"):
             print(f"Початок генерації договорів для заходу: {self.current_event.name}")

             generated_count = 0
             failed_count = 0
             failed_contracts = []

             # Ітеруємо по всіх договорах у поточному заході
             for contract in self.current_event.contracts:
                 # Визначаємо шлях до шаблону для цього конкретного договору
                 # Якщо у договору вказано власний шаблон, використовуємо його, інакше - загальний
                 current_contract_template = contract.template_path if contract.template_path and os.path.exists(contract.template_path) else template_to_use_default

                 if not os.path.exists(current_contract_template):
                      error_message = f"Пропущено договір '{contract.name}': Файл шаблону не знайдено за шляхом: {current_contract_template}"
                      print(error_message)
                      failed_count += 1
                      failed_contracts.append(contract.name)
                      error_logger.log_error(FileNotFoundError(f"Шаблон не знайдено для договору {contract.name}: {current_contract_template}"), error_message)
                      continue # Переходимо до наступного договору

                 try:
                     # Генеруємо документ за допомогою document_processor
                     print(f"  Генерація договору: '{contract.name}'...")
                     generated_filepath = document_processor.generate_document(
                         contract,
                         current_contract_template,
                         output_dir
                     )

                     if generated_filepath:
                         print(f"  Успішно згенеровано: {generated_filepath}")
                         # Логуємо інформацію про згенерований договір у Excel журнал
                         # Переконаємося, що excel_logger коректно обробляє None event/contract, хоча тут вони мають бути
                         excel_logger.log_contract(self.current_event, contract, generated_filepath)
                         generated_count += 1
                     else:
                         # Якщо generate_document повернув None (означає помилку, яка вже залогована)
                         print(f"  Не вдалося згенерувати договір: '{contract.name}'. Дивіться error.txt")
                         failed_count += 1
                         failed_contracts.append(contract.name)

                 except Exception as e:
                     # Перехоплюємо будь-які непередбачені помилки під час генерації конкретного договору
                     error_logger.log_error(e, f"Непередбачена помилка при генерації договору '{contract.name}' для заходу '{self.current_event.name}'")
                     print(f"  Непередбачена помилка при генерації договору '{contract.name}'. Дивіться error.txt")
                     failed_count += 1
                     failed_contracts.append(contract.name)


             # Повідомлення про завершення генерації
             summary_message = f"Генерація договорів для заходу '{self.current_event.name}' завершена.\n\n"
             summary_message += f"Успішно згенеровано: {generated_count}\n"
             summary_message += f"Не вдалося згенерувати: {failed_count}"

             if failed_contracts:
                 summary_message += "\n\nДоговори, що не вдалося згенерувати:\n" + "\n".join(failed_contracts)
                 messagebox.showerror("Генерація завершена з помилками", summary_message)
             else:
                 messagebox.showinfo("Генерація завершена успішно", summary_message)

             print(f"Генерація для заходу завершена. Успішно: {generated_count}, Помилок: {failed_count}")


    def on_contract_select(self, event=None):
        """
        Обробник вибору договору.
        """
        selected_items = self.contracts_tree.selection()
        if not selected_items:
            # Якщо нічого не обрано, скидаємо поточний договір
            self.current_contract = None
            # Очищаємо деталі договору, якщо нічого не обрано
            self.clear_contract_details_frame_content()
            # Перемикаємось на вкладку списку договорів
            # Перевіряємо, чи вкладка деталей існує (може бути невидимою)
            if self.details_notebook.index(self.contract_details_frame) in self.details_notebook.tabs():
                 self.details_notebook.hide(self.details_notebook.index(self.contract_details_frame))
                 self.details_notebook.select(self.contracts_list_frame)

            print("Нічого не обрано у списку договорів.")
            return

        selected_contract_id = selected_items[0]
        # Шукаємо обраний договір у списку договорів поточного заходу
        if self.current_event:
             self.current_contract = next((c for c in self.current_event.contracts if c.id == selected_contract_id), None)
        else:
             # Це не повинно трапитись, якщо on_event_select працює коректно
             # Але для безпеки можна спробувати знайти по всіх заходах, якщо current_event None
             # Цей сценарій малоймовірний, але додамо його для надійності.
             self.current_contract = None # Спочатку скидаємо
             for event in self.app_data.events:
                  self.current_contract = next((c for c in event.contracts if c.id == selected_contract_id), None)
                  if self.current_contract:
                       # Якщо знайшли договір, також оновлюємо current_event на відповідний захід
                       # TODO: Можливо, потрібно також оновити виділення в Treeview заходів self.events_tree.selection_set(event.id)
                       # Але це може викликати on_event_select рекурсивно. Поки без цього.
                       break # Знайшли договір, виходимо з циклу по заходах


        if self.current_contract:
            print(f"Обрано договір: {self.current_contract.name}")
            # Оновлюємо деталі договору в GUI
            self.update_contract_details(self.current_contract)
            # Перемикаємось на вкладку деталей договору
            # Перевіряємо, чи вкладка деталей існує (може бути невидимою)
            if self.details_notebook.index(self.contract_details_frame) in self.details_notebook.tabs():
                 # Показуємо вкладку деталей, якщо вона була прихована
                 self.details_notebook.enable(self.details_notebook.index(self.contract_details_frame))
                 self.details_notebook.select(self.details_notebook.index(self.contract_details_frame))

        else:
             # Якщо об'єкт договору не знайдено за ID (напр., був видалений поза GUI)
             error_logger.log_error(ValueError("Не знайдено об'єкт Договір для обраного ID у Treeview"), f"Не знайдено об'єкт Договір з ID: {selected_contract_id}")
             self.current_contract = None
             # Очищаємо деталі договору, якщо об'єкт не знайдено
             self.clear_contract_details_frame_content()
             # Ховаємо вкладку деталей, якщо об'єкт не знайдено
             if self.details_notebook.index(self.contract_details_frame) in self.details_notebook.tabs():
                  self.details_notebook.hide(self.details_notebook.index(self.contract_details_frame))
                  self.details_notebook.select(self.contracts_list_frame) # Повертаємось до списку


    def update_contract_details(self, contract: Contract):
        """
        Заповнює скрольований фрейм деталей договору полями введення
        відповідно до плейсхолдерів у contract.fields.
        Реалізує скрольовану область та динамічні поля.
        Використовує GUI_FIELD_DISPLAY_ORDER для сортування полів.
        """
        # Очищаємо попередній вміст фрейму деталей
        self.clear_contract_details_frame_content()

        # Встановлюємо поточний договір (вже зроблено в on_contract_select)
        self.current_contract = contract

        # --- ДОБАВЬТЕ ЭТУ СТРОКУ ДЛЯ ОТЛАДКИ ---
        print(f"DEBUG: Поля договора {contract.name}: {list(contract.fields.keys())}")
        # --------------------------------------

        # 2. Отримуємо внутрішній фрейм скрольованої області, куди будемо додавати віджети
        inner_frame = self.contract_details_scrolled_frame.inner_frame

        # Додаємо заголовок
        ttk.Label(inner_frame, text=f"Поля договору: {contract.name}", font='TkDefaultFont 10 bold').grid(row=0, column=0, columnspan=2, sticky="w", padx=5, pady=5)

        # 3. Додаємо поля введення для плейсхолдерів
        row_idx = 1

        # Визначаємо всі плейсхолдери з договору, які не є спеціальними для таблиці чи сум
        all_contract_placeholders = list(contract.fields.keys())
        special_placeholders = [constants.ITEM_LIST_PLACEHOLDER, constants.TOTAL_SUM_PLACEHOLDER, constants.TEXT_SUM_PLACEHOLDER]
        placeholders_to_create_entries_for = [
            ph for ph in all_contract_placeholders if ph not in special_placeholders
        ]

        # ### ПОПРАВКА ТУТ: Сортуємо плейсхолдери для відображення
        # Список плейсхолдерів для відображення, відсортований за GUI_FIELD_DISPLAY_ORDER
        sorted_placeholders = []

        # Додаємо плейсхолдери зі списку порядку відображення, якщо вони є в договорі
        for preferred_ph in constants.GUI_FIELD_DISPLAY_ORDER:
            if preferred_ph in placeholders_to_create_entries_for:
                sorted_placeholders.append(preferred_ph)
                # Видаляємо зі списку, щоб не додавати двічі
                placeholders_to_create_entries_for.remove(preferred_ph)

        # Додаємо решту плейсхолдерів (які були в договорі, але не в списку порядку) в алфавітному порядку
        sorted_placeholders.extend(sorted(placeholders_to_create_entries_for))


        # Словник для зберігання Tkinter змінних, пов'язаних з полями договору
        self._field_vars = {} # {placeholder: tk.StringVar}

        # ### ПОПРАВКА ТУТ: Використовуємо відсортований список плейсхолдерів для створення полів
        for placeholder in sorted_placeholders:
             # --- ДОБАВЬТЕ ЭТУ СТРОКУ ---
            print(f"DEBUG: Processing placeholder for GUI: {placeholder}")
            # -------------------------

            # Витягуємо "чисте" ім'я поля (без дужок) для підпису
            field_name = placeholder.strip('<>')

            ttk.Label(inner_frame, text=f"{field_name}:").grid(row=row_idx, column=0, sticky="w", padx=5, pady=2)

            # Використовуємо tk.StringVar для зв'язку поля введення з даними
            field_var = tk.StringVar()
            # Встановлюємо початкове значення з моделі
            field_var.set(contract.fields.get(placeholder, ""))
            self._field_vars[placeholder] = field_var


            # Додаємо віджет введення (Entry)
            # TODO: Можна використовувати tk.Text для багаторядкових полів, якщо це потрібно
            field_entry = ttk.Entry(inner_frame, textvariable=field_var, width=50)
            field_entry.grid(row=row_idx, column=1, sticky="ew", padx=5, pady=2)

            # --- ДОБАВЬТЕ ЭТУ СТРОКУ ---
            print(f"DEBUG: Создана Entry для {placeholder} в ряду {row_idx}")
            # -------------------------

            # Прив'язуємо обробник події зміни тексту (trace) для ВСІХ полів
            # Обробник on_field_change тепер включає базову валідацію для числових кандидатів
            field_var.trace_add('write', lambda name, index, mode, ph=placeholder, var=field_var: self.on_field_change(ph, var.get()))


            # Прив'язуємо контекстне меню до поля введення
            self._bind_text_context_menu(field_entry)

            # TODO: Можна додати візуальну індикацію (напр., колір рамки), якщо введення некоректне


            row_idx += 1 # <--- Увеличиваем номер ряда

        # --- ДОБАВЬТЕ ЭТИ ДВЕ СТРОКИ СРАЗУ ПОСЛЕ ОКОНЧАНИЯ ЦИКЛА ---
        print("DEBUG: Завершено створення полів для плейсхолдерів.") # Эта строка уже есть
        # Принудительно обновляем геометрию и область прокрутки
        inner_frame = self.contract_details_scrolled_frame.inner_frame # Убедимся, что inner_frame доступен
        inner_frame.update_idletasks()
        self.contract_details_scrolled_frame._on_frame_configure()
        # ----------------------------------------------------------

        
        # --- Секція для списку товарів ---
        # TODO: Можливо, додати сюди кнопку "Копіювати товари з іншого договору"?

        ttk.Label(inner_frame, text="Список товарів:", font='TkDefaultFont 10 bold').grid(row=row_idx, column=0, columnspan=2, sticky="w", padx=5, pady=10)
        row_idx += 1

        # Фрейм для кнопок управління товарами
        item_buttons_frame = ttk.Frame(inner_frame)
        item_buttons_frame.grid(row=row_idx, column=0, columnspan=2, sticky="ew", padx=5, pady=2)
        item_buttons_frame.columnconfigure(0, weight=1)
        item_buttons_frame.columnconfigure(1, weight=1)

        ttk.Button(item_buttons_frame, text="Додати Товар", command=self.add_item_to_contract).grid(row=0, column=0, sticky="ew", padx=2)
        # Змінюємо текст кнопки, щоб було зрозуміліше, що вона видаляє останній елемент
        ttk.Button(item_buttons_frame, text="Видалити ОСТАННІЙ Товар", command=self.remove_selected_item).grid(row=0, column=1, sticky="ew", padx=2) # TODO: Реалізувати вибір

        row_idx += 1

        # Фрейм, який буде містити рядки полів для кожного товару
        self.items_list_frame = ttk.Frame(inner_frame)
        self.items_list_frame.grid(row=row_idx, column=0, columnspan=2, sticky="nsew", padx=5, pady=5)
        # Налаштовуємо сітку для items_list_frame
        # Ширина колонок приблизна, можна налаштувати
        self.items_list_frame.columnconfigure(0, weight=4) # Назва
        self.items_list_frame.columnconfigure(1, weight=2) # ДК
        self.items_list_frame.columnconfigure(2, weight=1) # Кількість
        # self.items_list_frame.columnconfigure(3, weight=1) # Ціна за одиницю (якщо є)
        self.items_list_frame.columnconfigure(3, weight=2) # Сума
        self.items_list_frame.columnconfigure(4, weight=0) # Кнопка Видалити

        row_idx += 1 # Збільшуємо row_idx для наступних елементів


        # Додаємо існуючі товари з моделі даних
        # update_items_gui створює поля для товарів і прив'язує on_item_field_change
        self.update_items_gui(contract.items)


        # --- Секція для підрахунку суми ---
        # ПОЛЯ <разом> та <сума прописью> створюються тут, якщо вони є в словнику fields договору.
        # Це вже реалізовано в попередній версії update_contract_details.
        # Вони обробляються спеціально в calculate_total_sum.

        # Перевіряємо, чи потрібно створити поле <разом>
        # Плейсхолдер <разом> має бути доданий до contract.fields при створенні договору,
        # якщо він був знайдений в шаблоні АБО якщо він вказаний як спеціальний плейсхолдер.
        if constants.TOTAL_SUM_PLACEHOLDER in contract.fields: # Перевіряємо наявність плейсхолдера в даних
             ttk.Label(inner_frame, text=f"{constants.TOTAL_SUM_PLACEHOLDER.strip('<>')}:").grid(row=row_idx, column=0, sticky="w", padx=5, pady=2)
             # total_sum_var вже ініціалізовано в __init__
             self.total_sum_entry = ttk.Entry(inner_frame, textvariable=self.total_sum_var, state='readonly')
             self.total_sum_entry.grid(row=row_idx, column=1, sticky="ew", padx=5, pady=2)
             self._bind_text_context_menu(self.total_sum_entry) # Прив'язуємо ПКМ (тільки Копіювати/Виділити все)
             row_idx += 1

        # Перевіряємо, чи потрібно створити поле <сума прописью>
        # Плейсхолдер <сума прописью> має бути доданий до contract.fields при створенні договору,
        # якщо він був знайдений в шаблоні АБО якщо він вказаний як спеціальний плейсхолдер.
        if constants.TEXT_SUM_PLACEHOLDER in contract.fields: # Перевіряємо наявність плейсхолдера в даних
             ttk.Label(inner_frame, text=f"{constants.TEXT_SUM_PLACEHOLDER.strip('<>')}:").grid(row=row_idx, column=0, sticky="w", padx=5, pady=2)
             # text_sum_var вже ініціалізовано в __init__
             self.text_sum_entry = ttk.Entry(inner_frame, textvariable=self.text_sum_var, state='readonly')
             self.text_sum_entry.grid(row=row_idx, column=1, sticky="ew", padx=5, pady=2)
             self._bind_text_context_menu(self.text_sum_entry) # Прив'язуємо ПКМ (тільки Копіювати/Виділити все)
             row_idx += 1


        # Викликаємо функцію для початкового розрахунку суми та суми прописом
        # Це відбувається після заповнення полів товарів в update_items_gui
        self.calculate_total_sum() # Цей метод вже є і оновлює total_sum_var і text_sum_var


        # --- Кнопки дій для договору ---
        contract_actions_frame = ttk.Frame(inner_frame)
        contract_actions_frame.grid(row=row_idx, column=0, columnspan=2, sticky="ew", padx=5, pady=10)
        contract_actions_frame.columnconfigure(0, weight=1)
        contract_actions_frame.columnconfigure(1, weight=1)

        ttk.Button(contract_actions_frame, text="🧹 Очистити всі поля договору", command=self.clear_current_contract_fields).grid(row=0, column=0, sticky="ew", padx=2)
        ttk.Button(contract_actions_frame, text="📄 Генерувати тільки цей договір", command=self.generate_single_contract).grid(row=0, column=1, sticky="ew", padx=2)

        row_idx += 1

        # Оновлюємо область прокрутки після додавання всіх віджетів
        # Викликаємо через невелику затримку
        self.master.after(10, self.contract_details_scrolled_frame._on_frame_configure)


    def on_field_change(self, placeholder: str, value: str):
        """
        Обробник зміни тексту у полі введення плейсхолдера.
        Оновлює значення у словнику self.current_contract.fields.
        Виконує базову очистку введення для числових полів-кандидатів.
        """
        if self.current_contract:
            # ### ПОПРАВКА ТУТ: Додаємо базову очистку введення для числових полів
            field_name = placeholder.strip('<>')
            # Визначаємо, чи це поле, яке потенційно є числовим і потребує очистки
            NUMERIC_PLACEHOLDERS_GUI_CHECK = {
                 "<кількість>", "<ціна за одиницю>", "<загальна сума>", "<разом>" # Зі старого коду
            }
            # Додамо інші, якщо їх назва містить ключові слова (sum, quantity, price, total, грн)
            is_numeric_candidate = placeholder in NUMERIC_PLACEHOLDERS_GUI_CHECK or any(key in field_name.lower() for key in ["сума", "кількість", "ціна", "вартість", "грн", "total", "quantity", "price"])


            if is_numeric_candidate:
                 # Очистимо ввід: дозволяємо цифри, пробіли, кому та крапку
                 cleaned_value = re.sub(r'[^\d\s,\.]', '', value)
                 # Дозволяємо тільки одну крапку або одну кому як десятковий роздільник
                 if cleaned_value.count('.') > 1 or cleaned_value.count(',') > 1:
                      if '.' in cleaned_value and ',' in cleaned_value:
                           # Якщо є обидва, це, ймовірно, помилка. Можна залишити перший знайдений.
                           first_dot = cleaned_value.find('.')
                           first_comma = cleaned_value.find(',')
                           if first_dot < first_comma:
                                # Залишаємо частину до першої коми після першої крапки
                                second_comma_index = cleaned_value.find(',', first_dot+1)
                                if second_comma_index != -1:
                                     cleaned_value = cleaned_value[:second_comma_index]
                           else:
                                # Залишаємо частину до першої крапки після першої коми
                                second_dot_index = cleaned_value.find('.', first_comma+1)
                                if second_dot_index != -1:
                                     cleaned_value = cleaned_value[:second_dot_index]
                      elif cleaned_value.count('.') > 1:
                           # Залишаємо частину до другої крапки
                           second_dot_index = cleaned_value.find('.', cleaned_value.find('.') + 1)
                           if second_dot_index != -1:
                                cleaned_value = cleaned_value[:second_dot_index]
                      elif cleaned_value.count(',') > 1:
                            # Залишаємо частину до другої коми
                           second_comma_index = cleaned_value.find(',', cleaned_value.find(',') + 1)
                           if second_comma_index != -1:
                                cleaned_value = cleaned_value[:second_comma_index]

                 # Спеціальні поля <разом> та <сума прописью> оновлюються в calculate_total_sum,
                 # їх значення в моделі тут не змінюємо вручну, хоча їх StringVar може бути пов'язаний.
                 # Але обробник on_field_change не прив'язується до них, бо вони readonly.
                 # Тому ця перевірка не потрібна для <разом>/<сума прописью>.

                 # Оновлюємо значення поля в моделі (з очищеним значенням як рядком)
                 self.current_contract.fields[placeholder] = cleaned_value

                 # Оновлюємо StringVar поля в GUI, щоб показати очищене значення (якщо воно змінилось)
                 # Це потрібно робити обережно, щоб уникнути нескінченного циклу trace
                 current_gui_value = self._field_vars[placeholder].get()
                 if current_gui_value != cleaned_value:
                      # Відключаємо trace тимчасово
                      trace_info = self._field_vars[placeholder].trace_info()
                      if trace_info: # Перевіряємо, чи є прив'язані trace
                           # trace_info повертає список кортежів: [(trace_id, mode, callback), ...]
                           # Беремо перший ID для режиму 'write'
                           write_trace_id = None
                           for tid, mode, cb in trace_info:
                               if mode == 'write':
                                    write_trace_id = tid
                                    break
                           if write_trace_id:
                                self._field_vars[placeholder].trace_remove('write', write_trace_id)

                      self._field_vars[placeholder].set(cleaned_value)

                      # Включаємо trace знову з тим же callback
                      if write_trace_id: # Якщо trace був відключений
                           self._field_vars[placeholder].trace_add('write', lambda name, index, mode, ph=placeholder, var=self._field_vars[placeholder]: self.on_field_change(ph, var.get()))

                 # Якщо зміна цього числового поля повинна впливати на загальну суму
                 # (напр., якщо це поле суми або кількості, але не в таблиці товарів),
                 # потрібно додати виклик self.calculate_total_sum() тут.
                 # Але calculate_total_sum наразі залежить тільки від item.total_sum товарів.
                 # Тому зміни в цих "текстових" числових полях НЕ БУДУТЬ оновлювати загальну суму.
                 # Якщо це потрібно, логіку розрахунку суми треба переглянути.
                 # Наразі, розрахунок суми з товарів відбувається в calculate_total_sum
                 # і викликається з on_item_field_change.
                 pass # Не викликаємо calculate_total_sum тут, щоб уникнути плутанини.

            else:
                 # Якщо поле не числове, просто оновлюємо значення в моделі як є
                 self.current_contract.fields[placeholder] = value
                 # Значення у StringVar вже оновилось автоматично


            # TODO: Якщо змінилося поле, яке впливає на спільні поля інших договорів (COMMON_PLACEHOLDERS),
            # потрібно оновити і їх. Ця логіка вже описана в коментарі в попередній версії gui.py
            # ... (код для оновлення спільних полів) ...


        else:
             print(f"Помилка: Зміна поля {placeholder} відбулась, але current_contract відсутній.")


    def clear_contract_details_frame_content(self):
        """
        Очищає вміст внутрішнього фрейму деталей договору
        (того, що знаходиться всередині скрольованої області).
        """
        # Видаляємо всі віджети з внутрішнього фрейму скрольованої області
        if hasattr(self, 'contract_details_scrolled_frame') and self.contract_details_scrolled_frame.inner_frame:
            for widget in self.contract_details_scrolled_frame.inner_frame.winfo_children():
                widget.destroy()
            # Скидаємо збережені StringVar та інші словники, пов'язані з деталями договору
            self._field_vars = {}
            self._item_vars = {}
            # _item_widgets не потрібно скидати тут, воно очищається в update_items_gui
            self._item_frames = {} # Також очищаємо словник фреймів товарів
            print("Вміст деталей договору очищено.")
        # else:
             # print("ScrolledFrame для деталей договору ще не створено або очищено.") # Для відладки

        # Скидаємо обраний договір, коли його деталі очищені
        self.current_contract = None

        # Очищаємо також StringVar для полів суми
        # self.total_sum_var.set("") # Це робиться в calculate_total_sum коли current_contract None
        # self.text_sum_var.set("") # Це робиться в calculate_total_sum коли current_contract None
        self.calculate_total_sum() # Викликаємо, щоб очистити поля суми в GUI


    def clear_current_contract_fields(self):
         """
         Очищає значення всіх полів та список товарів для поточного обраного договору.
         """
         if self.current_contract:
              if messagebox.askyesno("Підтвердження", f"Ви впевнені, що хочете очистити всі поля для договору '{self.current_contract.name}'?"):
                 # Зберігаємо значення спільних полів перед очищенням, якщо вони заповнені,
                 # щоб не втратити дані, які мають бути однаковими для заходу.
                 common_field_values_to_restore = {}
                 if self.current_event and self.current_event.contracts:
                     # Простіший підхід: просто зберігаємо значення спільних полів поточного договору
                     for placeholder in constants.COMMON_PLACEHOLDERS:
                          if placeholder in self.current_contract.fields and self.current_contract.fields[placeholder].strip():
                               common_field_values_to_restore[placeholder] = self.current_contract.fields[placeholder]

                 # Очищаємо словник полів та список товарів
                 self.current_contract.fields.clear()
                 self.current_contract.items.clear()

                 # Відновлюємо значення спільних полів після очищення
                 for placeholder, value in common_field_values_to_restore.items():
                     self.current_contract.fields[placeholder] = value
                 # Також відновлюємо назву заходу, якщо вона була заповнена і існує
                 if "<назва заходу>" in self.current_contract.fields and self.current_event:
                     self.current_contract.fields["<назва заходу>"] = self.current_event.name


                 print(f"Поля договору '{self.current_contract.name}' очищено.")
                 # Перестворюємо GUI деталей договору, щоб відобразити порожні поля та порожній список товарів.
                 self.update_contract_details(self.current_contract)
         else:
              messagebox.showwarning("Попередження", "Не обрано договір для очищення полів.")


    # --- Методи для роботи зі списком товарів ---

    def add_item_to_contract(self):
        """
        Додає новий порожній товар до списку поточного договору
        та оновлює GUI.
        """
        if not self.current_contract:
             messagebox.showwarning("Попередження", "Спочатку оберіть договір, до якого потрібно додати товар.")
             return

        # Створюємо новий об'єкт Item
        new_item = Item()
        # Додаємо його до списку товарів поточного договору
        self.current_contract.items.append(new_item)
        print(f"Додано новий товар до договору '{self.current_contract.name}'. Всього товарів: {len(self.current_contract.items)}")

        # Оновлюємо GUI для відображення ВСІХ товарів (перемальовуємо список)
        self.update_items_gui(self.current_contract.items)
        # Перераховуємо суму
        self.calculate_total_sum()

        # TODO: Прокрутити скрольовану область до нового доданого товару


    def remove_selected_item(self):
        """
        Видаляє обраний товар зі списку поточного договору
        та оновлює GUI.
        Наразі видаляє останній товар як тимчасова реалізація.
        """
        if not self.current_contract or not self.current_contract.items:
             messagebox.showwarning("Попередження", "Не обрано договір або немає товарів для видалення.")
             return

        # TODO: Реалізувати механізм вибору товару в GUI.
        # Наразі видаляємо останній товар як тимчасове рішення для тестування
        item_to_remove = self.current_contract.items[-1] # Видаляємо останній елемент

        if messagebox.askyesno("Підтвердження видалення", f"Видалити останній товар зі списку ('{item_to_remove.name if item_to_remove.name else 'Без назви'}')?"):
             # Видаляємо з даних
             self.current_contract.items.remove(item_to_remove)
             print(f"Видалено товар з договору '{self.current_contract.name}'")

             # Оновлюємо GUI для відображення ВСІХ товарів (перемальовуємо список)
             self.update_items_gui(self.current_contract.items)
             # Перераховуємо суму
             self.calculate_total_sum()

             # Оновлюємо область прокрутки після видалення
             self.master.after(10, self.contract_details_scrolled_frame._on_frame_configure)

        else:
              print("Видалення товару скасовано.")


    def update_items_gui(self, items: list):
         """
         Оновлює секцію GUI (self.items_list_frame), що відображає список товарів.
         Створює поля введення для кожного товару в списку.
         """
         # Очищаємо поточні віджети товарів з items_list_frame
         for widget in self.items_list_frame.winfo_children():
             widget.destroy()

         # Словники для зберігання Tkinter змінних та посилань на віджети товарів
         self._item_vars = {} # {item_id: {attribute: tk.StringVar}}
         self._item_widgets = {} # {item_id: {attribute: widget}}

         # Заголовки колонок для списку товарів (розміщуємо їх безпосередньо в items_list_frame)
         ttk.Label(self.items_list_frame, text="Назва Товару").grid(row=0, column=0, sticky="w", padx=5, pady=2)
         ttk.Label(self.items_list_frame, text="ДК 021:2015").grid(row=0, column=1, sticky="w", padx=5, pady=2)
         ttk.Label(self.items_list_frame, text="Кількість").grid(row=0, column=2, sticky="w", padx=5, pady=2)
         # ttk.Label(self.items_list_frame, text="Ціна за одиницю").grid(row=0, column=3, sticky="w", padx=5, pady=2) # Наразі не використовуємо окремо
         ttk.Label(self.items_list_frame, text="Сума").grid(row=0, column=3, sticky="w", padx=5, pady=2)
         ttk.Label(self.items_list_frame, text="Дії").grid(row=0, column=4, sticky="w", padx=5, pady=2) # Заголовок для кнопок

         row_idx = 1 # Починаємо додавати рядки товарів з другого рядка

         for item in items:
             item_id = item.id
             self._item_vars[item_id] = {}
             self._item_widgets[item_id] = {}

             # Поле: Назва Товару
             name_var = tk.StringVar(value=item.name)
             name_entry = ttk.Entry(self.items_list_frame, textvariable=name_var, width=30)
             name_entry.grid(row=row_idx, column=0, sticky="ew", padx=2, pady=2)
             name_var.trace_add('write', lambda name, index, mode, item_id=item_id, attr="name", var=name_var: self.on_item_field_change(item_id, attr, var.get()))
             self._item_vars[item_id]["name"] = name_var
             self._item_widgets[item_id]["name"] = name_entry
             self._bind_text_context_menu(name_entry)

             # Поле: ДК 021:2015
             dk_var = tk.StringVar(value=item.dk)
             dk_entry = ttk.Entry(self.items_list_frame, textvariable=dk_var, width=15)
             dk_entry.grid(row=row_idx, column=1, sticky="ew", padx=2, pady=2)
             dk_var.trace_add('write', lambda name, index, mode, item_id=item_id, attr="dk", var=dk_var: self.on_item_field_change(item_id, attr, var.get()))
             self._item_vars[item_id]["dk"] = dk_var
             self._item_widgets[item_id]["dk"] = dk_entry
             self._bind_text_context_menu(dk_entry)

             # Поле: Кількість (числове)
             # Зберігаємо в StringVar як рядок, але чистимо введення
             quantity_var = tk.StringVar(value=str(item.quantity).replace('.', ',')) # Відображаємо з комою
             quantity_entry = ttk.Entry(self.items_list_frame, textvariable=quantity_var, width=10)
             quantity_entry.grid(row=row_idx, column=2, sticky="ew", padx=2, pady=2)
             # Прив'язуємо обробник зміни, який включає валідацію та оновлення моделі/суми
             quantity_var.trace_add('write', lambda name, index, mode, item_id=item_id, attr="quantity", var=quantity_var: self.on_item_field_change(item_id, attr, var.get()))
             self._item_vars[item_id]["quantity"] = quantity_var
             self._item_widgets[item_id]["quantity"] = quantity_entry
             self._bind_text_context_menu(quantity_entry)

             # Поле: Сума для цього товару (числове)
             # Зберігаємо в StringVar як рядок, але чистимо введення
             total_sum_var = tk.StringVar(value=str(item.total_sum).replace('.', ',')) # Відображаємо з комою
             total_sum_entry = ttk.Entry(self.items_list_frame, textvariable=total_sum_var, width=15)
             total_sum_entry.grid(row=row_idx, column=3, sticky="ew", padx=2, pady=2)
             # Прив'язуємо обробник зміни, який включає валідацію та оновлення моделі/суми
             total_sum_var.trace_add('write', lambda name, index, mode, item_id=item_id, attr="total_sum", var=total_sum_var: self.on_item_field_change(item_id, attr, var.get()))
             self._item_vars[item_id]["total_sum"] = total_sum_var
             self._item_widgets[item_id]["total_sum"] = total_sum_entry
             self._bind_text_context_menu(total_sum_entry)


             # Кнопка "Видалити" для цього товару
             # Передаємо сам об'єкт товару в команду кнопки
             delete_button = ttk.Button(self.items_list_frame, text="Видалити", command=lambda item_obj=item: self.remove_specific_item(item_obj), width=8) # Ширина кнопки
             delete_button.grid(row=row_idx, column=4, sticky="w", padx=2, pady=2)


             row_idx += 1

         # Після додавання всіх товарів, оновлюємо область прокрутки головного скрольованого фрейму
         self.items_list_frame.update_idletasks() # Оновлюємо геометрію фрейму товарів
         self.master.after(10, self.contract_details_scrolled_frame._on_frame_configure) # Оновлюємо скролбар


    def on_item_field_change(self, item_id: str, attribute: str, value: str):
         """
         Обробник зміни тексту у полі введення конкретного товару.
         Оновлює відповідний атрибут об'єкта Item та перераховує загальну суму.
         Додає базову валідацію та очищення для числових полів товару.
         """
         if self.current_contract:
             # Знаходимо об'єкт Item за його ID
             item = next((item for item in self.current_contract.items if item.id == item_id), None)
             if item:
                 # ### ПОПРАВКА ТУТ: Базова валідація та очищення для числових полів товару
                 if attribute in ["quantity", "total_sum"]: # Поля, які мають бути числами
                     # Очистимо ввід: дозволяємо цифри, пробіли, кому та крапку
                     cleaned_value = re.sub(r'[^\d\s,\.]', '', value)
                     # Дозволяємо тільки одну крапку або одну кому як десятковий роздільник
                     if cleaned_value.count('.') > 1 or cleaned_value.count(',') > 1:
                          if '.' in cleaned_value and ',' in cleaned_value:
                               first_dot = cleaned_value.find('.')
                               first_comma = cleaned_value.find(',')
                               if first_dot < first_comma:
                                    second_comma_index = cleaned_value.find(',', first_dot+1)
                                    if second_comma_index != -1:
                                         cleaned_value = cleaned_value[:second_comma_index]
                               else:
                                    second_dot_index = cleaned_value.find('.', first_comma+1)
                                    if second_dot_index != -1:
                                         cleaned_value = cleaned_value[:second_dot_index]
                          elif cleaned_value.count('.') > 1:
                               second_dot_index = cleaned_value.find('.', cleaned_value.find('.') + 1)
                               if second_dot_index != -1:
                                    cleaned_value = cleaned_value[:second_dot_index]
                          elif cleaned_value.count(',') > 1:
                               second_comma_index = cleaned_value.find(',', cleaned_value.find(',') + 1)
                               if second_comma_index != -1:
                                    cleaned_value = cleaned_value[:second_comma_index]


                     # Замінюємо кому на крапку для коректного перетворення у float
                     float_candidate = cleaned_value.replace(',', '.').replace(' ', '') # Прибираємо пробіли перед конвертацією

                     try:
                         # Спробуємо конвертувати в float
                         if not float_candidate or float_candidate == '.':
                             float_value = 0.0
                         else:
                              float_value = float(float_candidate)

                         # Оновлюємо атрибут об'єкта Item як float
                         setattr(item, attribute, float_value)

                     except ValueError:
                         # Якщо після очищення все одно не вдалося конвертувати в число
                         print(f"Попередження: Некоректне числове значення для поля '{attribute}' товару '{item.name}': '{value}'")
                         # Значення в моделі Item (float) залишається попереднім або 0.0.
                         setattr(item, attribute, 0.0) # Встановлюємо 0.0 якщо конвертація не вдалась
                         # Можна додати візуальне повідомлення про помилку введення.
                         pass # Нічого не робимо зі StringVar тут

                     # ### ПОПРАВКА ТУТ: Оновлюємо StringVar поля в GUI з очищеним значенням
                     # Тільки якщо очищене значення відрізняється від поточного у StringVar:
                     current_gui_value = self._item_vars[item_id][attribute].get()
                     if current_gui_value != cleaned_value:
                           trace_info = self._item_vars[item_id][attribute].trace_info()
                           if trace_info:
                                write_trace_id = None
                                for tid, mode, cb in trace_info:
                                     if mode == 'write':
                                          write_trace_id = tid
                                          break
                                if write_trace_id:
                                     self._item_vars[item_id][attribute].trace_remove('write', write_trace_id)

                           self._item_vars[item_id][attribute].set(cleaned_value)

                           if write_trace_id:
                                self._item_vars[item_id][attribute].trace_add('write', lambda name, index, mode, iid=item_id, attr=attribute, var=self._item_vars[item_id][attribute]: self.on_item_field_change(iid, attr, var.get()))


                 else: # Для текстових полів (назва, ДК)
                     # Для текстових полів просто оновлюємо модель
                     setattr(item, attribute, value)
                     # Значення у StringVar вже оновилось автоматично


                 # Після будь-якої зміни поля товару, перераховуємо загальну суму договору
                 self.calculate_total_sum()
             else:
                 print(f"Помилка: Не знайдено об'єкт Item з ID {item_id} при зміні поля '{attribute}'")
         else:
              print(f"Помилка: Зміна поля товару відбулась, але current_contract відсутній.")


    def remove_specific_item(self, item: Item):
         """
         Видаляє конкретний об'єкт Item зі списку поточного договору
         та оновлює GUI. Викликається кнопкою "Видалити" біля кожного товару.
         """
         # Перевірка наявності поточного договору, списку товарів та самого товару
         if not self.current_contract or not self.current_contract.items or item not in self.current_contract.items:
              print("Помилка: Не вдалося знайти товар для видалення або договір не обрано.")
              return

         # Запит підтвердження перед видаленням
         if messagebox.askyesno("Підтвердження видалення", f"Видалити товар '{item.name if item.name else 'Без назви'}'?"):
             # Видаляємо з даних (списку товарів у моделі договору)
             self.current_contract.items.remove(item)
             print(f"Видалено товар з договору '{self.current_contract.name}'")

             # Оновлюємо GUI для відображення ВСІХ товарів (перемальовуємо секцію списку товарів)
             self.update_items_gui(self.current_contract.items)
             # Перераховуємо загальну суму договору
             self.calculate_total_sum()

             # Оновлюємо область прокрутки після видалення (можливо, розмір зменшився)
             self.master.after(10, self.contract_details_scrolled_frame._on_frame_configure)

         else:
              print("Видалення товару скасовано користувачем.")


    def calculate_total_sum(self):
         """
         Перераховує загальну суму для поточного договору на основі списку товарів.
         Оновлює поля <разом> та <сума прописью> в моделі та GUI.
         """
         # Якщо немає обраного договору, очищаємо поля суми в GUI та моделі (якщо вони там були)
         if not self.current_contract:
             self.total_sum_var.set("")
             self.text_sum_var.set("")
             # Очищаємо в моделі, якщо вони там є
             if self.current_contract and constants.TOTAL_SUM_PLACEHOLDER in self.current_contract.fields:
                  self.current_contract.fields[constants.TOTAL_SUM_PLACEHOLDER] = ""
             if self.current_contract and constants.TEXT_SUM_PLACEHOLDER in self.current_contract.fields:
                  self.current_contract.fields[constants.TEXT_SUM_PLACEHOLDER] = ""
             print("Calculate total sum: Немає обраного договору.")
             return

         total = 0.0
         for item in self.current_contract.items:
             # Додаємо до загальної суми тільки валідні числові значення total_sum товару
             try:
                 # total_sum в моделі Item зберігається як float
                 total += float(item.total_sum)
             except (ValueError, TypeError):
                 # Якщо total_sum в Item чомусь не число, ігноруємо цей товар
                 # Це може статися, якщо при завантаженні даних було некоректне значення
                 print(f"Попередження: Некоректне числове значення total_sum в товарі '{item.name}': {item.total_sum}. Ігнорується.")
                 pass

         # Оновлюємо поле <разом> у словнику fields договору (в моделі)
         # Зберігаємо в моделі як рядок з крапкою, 2 знаки після коми
         formatted_total_model = f"{total:.2f}"
         self.current_contract.fields[constants.TOTAL_SUM_PLACEHOLDER] = formatted_total_model

         # Оновлюємо StringVar, пов'язаний з полем <разом> в GUI (для відображення)
         # Відображаємо з комою як десятковим роздільником
         formatted_total_gui = formatted_total_model.replace('.', ',')
         self.total_sum_var.set(formatted_total_gui)


         # Перетворюємо суму прописом і оновлюємо поле <сума прописью>
         try:
             # Викликаємо функцію перетворення суми прописом з модуля number_to_text_ua
             # Передаємо числове значення total
             sum_text_ua = number_to_text_ua.number_to_currency_text(total)

             # Оновлюємо поле <сума прописью> у словнику fields договору (в моделі)
             self.current_contract.fields[constants.TEXT_SUM_PLACEHOLDER] = sum_text_ua

             # Оновлюємо StringVar, пов'язаний з полем <сума прописью> в GUI
             self.text_sum_var.set(sum_text_ua)

         except Exception as e:
             # Якщо виникла помилка при перетворенні суми прописом
             error_logger.log_error(e, f"Помилка при перетворенні суми '{total}' прописом для договору '{self.current_contract.name}'")
             text_sum = "Помилка перетворення суми прописом"
             self.current_contract.fields[constants.TEXT_SUM_PLACEHOLDER] = text_sum
             self.text_sum_var.set(text_sum)
             print(f"Помилка перетворення суми прописом: {e}")


         print(f"Перераховано загальну суму для '{self.current_contract.name}': {formatted_total_model}")


    def generate_single_contract(self):
        """
        Генерує документ тільки для поточного обраного договору.
        """
        if not self.current_contract:
             messagebox.showwarning("Попередження", "Спочатку оберіть договір для генерації.")
             return

        if not self.current_event:
             # Це не повинно трапитись, якщо current_contract обрано з Treeview, але для безпеки
             messagebox.showerror("Помилка", "Не вдалося визначити захід для обраного договору.")
             error_logger.log_error(ValueError("current_event відсутній при спробі генерації single_contract"), f"Спроба генерації договору '{self.current_contract.name}' без обраного заходу.")
             return

        # Перевіряємо наявність шаблону та папки збереження
        # Визначаємо шаблон для цього конкретного договору (спочатку шаблон договору, потім загальний)
        template_to_use = self.current_contract.template_path if self.current_contract.template_path and os.path.exists(self.current_contract.template_path) else self.app_data.template_path

        if not template_to_use or not os.path.exists(template_to_use):
            messagebox.showwarning("Попередження", f"Файл шаблону для договору '{self.current_contract.name}' не обрано або не знайдено за шляхом: {template_to_use}. Генерація скасована.")
            return

        output_dir = self.app_data.output_dir
        if not output_dir or not os.path.isdir(output_dir):
             messagebox.showwarning("Попередження", "Будь ласка, оберіть папку для збереження згенерованих договорів.")
             self.select_output_directory() # Пропонуємо обрати папку
             output_dir = self.app_data.output_dir
             if not output_dir or not os.path.isdir(output_dir):
                  messagebox.showerror("Помилка збереження", "Не вдалося визначити або знайти папку для збереження. Генерація скасована.")
                  return


        if messagebox.askyesno("Підтвердження генерації", f"Ви впевнені, що хочете згенерувати договір '{self.current_contract.name}' для заходу '{self.current_event.name}' у папку '{output_dir}'?"):
            print(f"Початок генерації договору: {self.current_contract.name}")

            try:
                # Генеруємо документ за допомогою document_processor
                generated_filepath = document_processor.generate_document(
                    self.current_contract,
                    template_to_use,
                    output_dir
                )

                if generated_filepath:
                    print(f"Успішно згенеровано: {generated_filepath}")
                    # Логуємо інформацію про згенерований договір у Excel журнал
                    excel_logger.log_contract(self.current_event, self.current_contract, generated_filepath)
                    messagebox.showinfo("Генерація завершена", f"Договір '{self.current_contract.name}' успішно згенеровано.\nФайл збережено за шляхом:\n{generated_filepath}")
                else:
                    # Якщо generate_document повернув None (означає помилку, яка вже залогована)
                    messagebox.showerror("Помилка генерації", f"Не вдалося згенерувати договір '{self.current_contract.name}'. Дивіться error.txt")
                    print(f"Не вдалося згенерувати договір: '{self.current_contract.name}'. Дивіться error.txt")


            except Exception as e:
                # Перехоплюємо будь-які непередбачені помилки під час генерації
                error_logger.log_error(e, f"Непередбачена помилка при генерації договору '{self.current_contract.name}' для заходу '{self.current_event.name}'")
                messagebox.showerror("Критична помилка генерації", f"Виникла непередбачена помилка при генерації договору '{self.current_contract.name}'. Дивіться error.txt")
                print(f"Критична помилка при генерації договору: '{self.current_contract.name}'. Дивіться error.txt")


    def _bind_text_context_menu(self, widget):
         """Прив'язує контекстне меню до заданого віджета введення тексту."""
         # Ми вже прив'язуємо меню до головного вікна, і меню саме перевіряє тип віджета.
         # Але якщо хочемо прив'язати індивідуально:
         # if isinstance(widget, (ttk.Entry, tk.Entry, tk.Text, scrolledtext.ScrolledText)):
         #    widget.bind("<Button-3>", self.text_menu.show, add="+")