# document_block.py 

import os
import tkinter.messagebox as messagebox
import customtkinter as ctk
from tkinter import filedialog
import json

from globals import EXAMPLES, document_blocks
from gui_utils import bind_entry_shortcuts, create_context_menu
from custom_widgets import CustomEntry
from text_utils import number_to_ukrainian_text
from data_persistence import get_template_memory, load_memory
from state_manager import save_current_state
from event_common_fields import fill_common_fields_for_new_contract, COMMON_FIELDS
from generation import enhanced_extract_placeholders_from_word

# Глобальные поля мероприятия (создаются один раз для всего мероприятия)
global_event_fields = {}
global_event_frame = None

# Поля товаров - определяем структуру полей для одного товара
PRODUCT_FIELDS = {
    "товар": "Название товара/услуги",
    "дк": "ДК код товара",
    "кількість": "Количество",
    "одиниця виміру": "Единица измерения",
    "ціна за одиницю": "Цена за единицу",
    "сума": "Сумма по товару",
    "пдв": "НДС"
}


def create_global_event_fields(parent_frame, tabview):
    """Создает глобальные поля мероприятия без визуального интерфейса"""
    global global_event_fields, global_event_frame

    # Если поля уже созданы для этой вкладки, не создаем повторно
    current_tab = tabview.get()
    if current_tab in global_event_fields:
        return global_event_fields[current_tab]

    # Создаем невидимые поля для использования в шаблонах
    event_fields_data = [
        ("захід", "Название мероприятия"),
        ("дата", "Дата проведения мероприятия"),
        ("адреса", "Адрес проведения мероприятия")
    ]

    current_event_entries = {}
    general_memory = load_memory()

    # Создаем невидимые entry-поля (они не будут отображаться)
    for field_key, description in event_fields_data:
        # Создаем скрытое поле ввода
        entry = CustomEntry(parent_frame, field_name=field_key, examples_dict=EXAMPLES)

        # Заполняем сохраненными данными
        saved_value = general_memory.get(field_key)
        if saved_value:
            entry.set_text(saved_value)

        current_event_entries[field_key] = entry

        # Скрываем поле (не добавляем в pack/grid)
        entry.pack_forget()

    # Сохраняем ссылки на глобальные поля
    if current_tab not in global_event_fields:
        global_event_fields[current_tab] = {}

    global_event_fields[current_tab] = current_event_entries
    global_event_frame = None  # Нет видимого фрейма

    # Заполняем общие поля из данных мероприятия
    fill_common_fields_for_new_contract(current_tab, current_event_entries)

    return current_event_entries


def get_global_event_fields(tab_name):
    """Возвращает глобальные поля мероприятия для указанной вкладки"""
    return global_event_fields.get(tab_name, {})


def create_products_table_widget(parent_frame, products_data=None):
    """Створює повноцінну редаговану таблицю товарів для document_block.py"""

    # Головний фрейм для таблиці товарів
    products_frame = ctk.CTkFrame(parent_frame)
    products_frame.pack(fill="both", expand=True, padx=5, pady=5)

    # Заголовок таблиці
    header_label = ctk.CTkLabel(products_frame, text="🛒 Таблиця товарів/послуг",
                                font=("Arial", 16, "bold"), text_color="#FF6B35")
    header_label.pack(pady=(10, 5))

    # Контейнер для скролінгу
    scrollable_frame = ctk.CTkScrollableFrame(products_frame, height=400)
    scrollable_frame.pack(fill="both", expand=True, padx=10, pady=5)

    # Заголовки стовпців
    headers_frame = ctk.CTkFrame(scrollable_frame, fg_color="gray25")
    headers_frame.pack(fill="x", pady=(0, 5))

    headers = ["№", "Товар/Послуга", "ДК-021:2015", "Кіл-ть", "Од.вим.", "Ціна за од.", "Сума", "Дії"]
    header_widths = [40, 250, 120, 80, 100, 100, 100, 80]

    for i, (header, width) in enumerate(zip(headers, header_widths)):
        label = ctk.CTkLabel(headers_frame, text=header, width=width,
                             font=("Arial", 12, "bold"))
        label.grid(row=0, column=i, padx=2, pady=5, sticky="ew")

    # Список для зберігання рядків товарів
    product_rows = []

    # Створюємо контекстне меню для полів
    from gui_utils import create_context_menu, bind_entry_shortcuts
    context_menu = create_context_menu(products_frame)

    def calculate_row_total(row_index):
        """Перераховує суму для рядка товару"""
        try:
            if row_index >= len(product_rows):
                return

            row_data = product_rows[row_index]
            qty_text = row_data["entries"]["кількість"].get().replace(",", ".").strip()
            price_text = row_data["entries"]["ціна за одиницю"].get().replace(",", ".").strip()

            qty = float(qty_text) if qty_text else 0
            price = float(price_text) if price_text else 0
            total = qty * price

            # Оновлюємо поле суми для рядка
            sum_entry = row_data["entries"]["сума"]
            sum_entry.configure(state="normal")
            sum_entry.delete(0, "end")
            sum_entry.insert(0, f"{total:.2f}")
            sum_entry.configure(state="readonly")

            # Перераховуємо загальну суму
            update_total_display()

        except (ValueError, AttributeError):
            pass

    def add_product_row(product_data=None):
        """Додає новий рядок товару"""
        row_index = len(product_rows)

        # Фрейм для рядка товару
        row_frame = ctk.CTkFrame(scrollable_frame, fg_color="transparent")
        row_frame.pack(fill="x", pady=1)

        # Словник для зберігання entry полів рядка
        row_entries = {}

        # Номер рядка
        num_label = ctk.CTkLabel(row_frame, text=str(row_index + 1), width=40)
        num_label.grid(row=0, column=0, padx=2, pady=2)

        # Поля товару з правильними назвами згідно PRODUCT_FIELDS
        fields = [
            ("товар", 250),
            ("дк", 120),
            ("кількість", 80),
            ("одиниця виміру", 100),
            ("ціна за одиницю", 100),
            ("сума", 100)
        ]

        for col, (field_name, width) in enumerate(fields, 1):
            # Використовуємо CustomEntry замість ctk.CTkEntry
            entry = CustomEntry(row_frame, field_name=field_name, examples_dict=EXAMPLES, width=width)
            entry.grid(row=0, column=col, padx=2, pady=2, sticky="ew")
            row_entries[field_name] = entry

            # Заповнюємо даними якщо є
            if product_data and field_name in product_data:
                entry.insert(0, str(product_data[field_name]))

            # Прив'язуємо контекстне меню та клавіатурні скорочення
            bind_entry_shortcuts(entry, context_menu)

            # Прив'язуємо події для перерахунку
            if field_name in ["кількість", "ціна за одиницю"]:
                entry.bind("<KeyRelease>", lambda e, idx=row_index: calculate_row_total(idx))
                entry.bind("<FocusOut>", lambda e, idx=row_index: calculate_row_total(idx))

        # Робимо поле суми readonly
        row_entries["сума"].configure(state="readonly", fg_color=("gray90", "gray20"))

        # Кнопка видалення рядка
        def remove_row():
            if len(product_rows) > 1:  # Залишаємо хоча б один рядок
                row_frame.destroy()
                product_rows.pop(row_index)
                update_row_numbers()
                update_total_display()
            else:
                import tkinter.messagebox as messagebox
                messagebox.showwarning("Попередження", "Повинен залишитися хоча б один рядок товару")

        remove_btn = ctk.CTkButton(row_frame, text="🗑", width=30, height=25,
                                   fg_color="red", hover_color="darkred", command=remove_row)
        remove_btn.grid(row=0, column=len(fields) + 1, padx=2, pady=2)

        # Зберігаємо дані рядка
        row_data = {
            "frame": row_frame,
            "entries": row_entries,
            "remove_btn": remove_btn,
            "num_label": num_label
        }
        product_rows.append(row_data)

        # Первинний розрахунок
        calculate_row_total(row_index)
        return row_data

    def update_row_numbers():
        """Оновлює номери рядків після видалення"""
        for i, row_data in enumerate(product_rows):
            row_data["num_label"].configure(text=str(i + 1))

    def add_new_product():
        """Додає новий порожній рядок товару"""
        add_product_row()

    # Кнопки управління таблицею
    buttons_frame = ctk.CTkFrame(products_frame, fg_color="transparent")
    buttons_frame.pack(fill="x", padx=10, pady=5)

    add_btn = ctk.CTkButton(buttons_frame, text="➕ Додати товар", command=add_new_product,
                            fg_color="#2E8B57", hover_color="#228B22")
    add_btn.pack(side="left", padx=5)

    def clear_all_products():
        """Очищає всі рядки товарів"""
        import tkinter.messagebox as messagebox
        result = messagebox.askyesno("Підтвердження", "Ви впевнені, що хочете очистити всі товари?")
        if result:
            for row_data in product_rows[:]:
                row_data["frame"].destroy()
            product_rows.clear()
            add_product_row()  # Додаємо один порожній рядок

    clear_btn = ctk.CTkButton(buttons_frame, text="🗑 Очистити все", command=clear_all_products,
                              fg_color="#DC3545", hover_color="#C82333")
    clear_btn.pack(side="left", padx=5)

    # Поле загальної суми
    total_frame = ctk.CTkFrame(products_frame, fg_color="transparent")
    total_frame.pack(fill="x", padx=10, pady=5)

    total_label = ctk.CTkLabel(total_frame, text="Загальна сума:", font=("Arial", 14, "bold"))
    total_label.pack(side="left", padx=5)

    # Використовуємо CustomEntry для поля загальної суми
    total_entry = CustomEntry(total_frame, field_name="разом", examples_dict=EXAMPLES, width=120)
    total_entry.pack(side="left", padx=5)
    total_entry.configure(state="readonly", fg_color=("gray90", "gray20"))

    # Прив'язуємо контекстне меню до поля загальної суми
    bind_entry_shortcuts(total_entry, context_menu)

    def update_total_display():
        """Оновлює відображення загальної суми"""
        total_sum = 0
        for row_data in product_rows:
            try:
                sum_text = row_data["entries"]["сума"].get().replace(",", ".").strip()
                if sum_text:
                    total_sum += float(sum_text)
            except (ValueError, AttributeError):
                pass

        total_entry.configure(state="normal")
        total_entry.delete(0, "end")
        total_entry.insert(0, f"{total_sum:.2f}")
        total_entry.configure(state="readonly")
        return total_sum

    def get_products_data():
        """Повертає дані всіх товарів у вигляді списку словників"""
        products = []
        for row_data in product_rows:
            product = {}
            for field_name, entry in row_data["entries"].items():
                value = entry.get().strip()
                product[field_name] = value

            # Додаємо товар тільки якщо є назва або код
            if product.get("товар") or product.get("дк"):
                products.append(product)

        return products

    def set_products_data(products_data):
        """Встановлює дані товарів"""
        # Очищаємо існуючі рядки
        for row_data in product_rows[:]:
            row_data["frame"].destroy()
        product_rows.clear()

        # Додаємо нові рядки
        if products_data and len(products_data) > 0:
            for product in products_data:
                add_product_row(product)
        else:
            add_product_row()  # Додаємо порожній рядок

        update_total_display()

    def get_total_sum():
        """Повертає поточну загальну суму"""
        return update_total_display()

    # Ініціалізуємо рядки товарів
    if products_data and len(products_data) > 0:
        for product in products_data:
            add_product_row(product)
    else:
        add_product_row()  # Додаємо один порожній рядок за замовчуванням

    # Повертаємо словник з усіма потрібними методами та об'єктами
    return {
        "frame": products_frame,
        "get_data": get_products_data,
        "set_data": set_products_data,
        "total_entry": total_entry,
        "update_total": update_total_display,
        "get_total_sum": get_total_sum,
        "product_rows": product_rows,
        "context_menu": context_menu
    }




def create_document_fields_block(parent_frame, tabview=None, template_path=None):
    if not tabview:
        messagebox.showerror("Помилка", "TabView не переданий")
        return

    if not template_path:
        template_path = filedialog.askopenfilename(
            title="Оберіть шаблон договору",
            filetypes=[("Word Documents", "*.docm *.docx")]
        )
        if not template_path:
            return

    try:
        dynamic_fields = enhanced_extract_placeholders_from_word(template_path)

        # Проверяем наличие плейсхолдера <таблиця_товарів>
        has_products_table = "<таблиця_товарів>" in open(template_path, 'rb').read().decode('utf-8', errors='ignore')

        if not dynamic_fields and not has_products_table:
            messagebox.showwarning("Увага",
                                   f"У шаблоні {os.path.basename(template_path)} не знайдено жодного плейсхолдера типу <поле> або <таблиця_товарів>.\n"
                                   "Переконайтеся, що у шаблоні є поля у форматі <назва_поля> або <таблиця_товарів>")
            return

    except Exception as e:
        messagebox.showerror("Помилка", f"Не вдалося прочитати шаблон:\n{e}")
        return

    # Создаем или получаем глобальные поля мероприятия
    global_entries = create_global_event_fields(parent_frame, tabview)

    # Создаем блок для полей конкретного договора
    block_frame = ctk.CTkFrame(parent_frame, border_width=1, border_color="gray70")
    block_frame.pack(pady=10, padx=5, fill="both", expand=True)

    header_frame = ctk.CTkFrame(block_frame, fg_color="transparent")
    header_frame.pack(fill="x", padx=5, pady=(5, 0))

    products_info = f" + таблиця товарів" if has_products_table else ""
    path_label = ctk.CTkLabel(header_frame,
                              text=f"Шаблон: {os.path.basename(template_path)} ({len(dynamic_fields)} полів{products_info})",
                              text_color="blue", anchor="w", wraplength=800)
    path_label.pack(side="left", padx=(0, 5), expand=True, fill="x")

    current_block_entries = {}
    template_specific_memory = get_template_memory(template_path)
    general_memory = load_memory()

    # Определяем поля договора (исключаем глобальные поля мероприятия и товарные поля)
    global_event_field_names = ["захід", "дата", "адреса"]
    product_field_names = list(PRODUCT_FIELDS.keys())

    # Исключаем из обычных полей товарные поля и глобальные поля
    contract_fields = [field for field in dynamic_fields
                       if field not in global_event_field_names
                       and field not in product_field_names]

    # Добавляем остальные общие поля (если есть в шаблоне)
    template_common_fields = [field for field in contract_fields if field in COMMON_FIELDS]

    # Создаем контекстное меню
    main_context_menu = create_context_menu(block_frame)

    # БЛОК ОБЩИХ ПОЛЕЙ ДОГОВОРА (кроме глобальных полей мероприятия)
    if template_common_fields:
        common_data_frame = ctk.CTkFrame(block_frame)
        common_data_frame.pack(fill="x", padx=5, pady=5)

        common_label = ctk.CTkLabel(common_data_frame, text="📋 Общие данные договора",
                                    font=("Arial", 14, "bold"), text_color="#FF6B35")
        common_label.pack(pady=(10, 5))

        common_grid_frame = ctk.CTkFrame(common_data_frame, fg_color="transparent")
        common_grid_frame.pack(fill="x", padx=10, pady=(0, 10))

        for i, field_key in enumerate(template_common_fields):
            label = ctk.CTkLabel(common_grid_frame, text=f"<{field_key}>", anchor="w", width=140,
                                 font=("Arial", 12, "bold"), text_color="#FF6B35")
            label.grid(row=i, column=0, padx=5, pady=3, sticky="w")

            entry = CustomEntry(common_grid_frame, field_name=field_key, examples_dict=EXAMPLES)
            entry.grid(row=i, column=1, padx=5, pady=3, sticky="ew")
            common_grid_frame.columnconfigure(1, weight=1)

            saved_value = template_specific_memory.get(field_key, general_memory.get(field_key))
            if saved_value is not None:
                entry.set_text(saved_value)

            bind_entry_shortcuts(entry, main_context_menu)
            current_block_entries[field_key] = entry

            hint_text = EXAMPLES.get(field_key, f"Общее поле договора: {field_key}")
            hint_button = ctk.CTkButton(common_grid_frame, text="ℹ", width=28, height=28, font=("Arial", 14),
                                        command=lambda h=hint_text, f=field_key:
                                        messagebox.showinfo(f"Підказка для <{f}>", h))
            hint_button.grid(row=i, column=2, padx=(0, 5), pady=3, sticky="e")

    # ТАБЛИЦА ТОВАРОВ (если есть плейсхолдер)
    products_widget = None
    if has_products_table:
        # Пытаемся загрузить сохраненные данные товаров
        saved_products = []
        if template_path in [block.get("path") for block in document_blocks]:
            # Ищем существующий блок с этим шаблоном
            for block in document_blocks:
                if block.get("path") == template_path and "products" in block.get("entries", {}):
                    saved_products = block["entries"]["products"]
                    break

        products_widget = create_products_table_widget(block_frame, saved_products)

    # БЛОК СПЕЦИФИЧЕСКИХ ПОЛЕЙ ДОГОВОРА
    specific_contract_fields = [field for field in contract_fields if field not in COMMON_FIELDS]

    if specific_contract_fields:
        fields_grid_frame = ctk.CTkFrame(block_frame, fg_color="transparent")
        fields_grid_frame.pack(fill="both", expand=True, padx=5, pady=5)

        # Сортируем поля по приоритету
        priority_fields = [
            # Финансовые поля (не товарные)
            "загальна сума", "разом", "всього", "сума прописом",
            "знижка", "аванс", "доплата", "залишок", "пдв", "ндс"
        ]

        sorted_fields = []
        remaining_fields = specific_contract_fields.copy()

        # Сначала добавляем поля по приоритету
        for priority_field in priority_fields:
            if priority_field in remaining_fields:
                sorted_fields.append(priority_field)
                remaining_fields.remove(priority_field)

        # Затем добавляем оставшиеся поля в алфавитном порядке
        sorted_fields.extend(sorted(remaining_fields))

        # Создаем поля в отсортированном порядке
        for i, field_key in enumerate(sorted_fields):
            label = ctk.CTkLabel(fields_grid_frame, text=f"<{field_key}>", anchor="w", width=140,
                                 font=("Arial", 12))
            label.grid(row=i, column=0, padx=5, pady=3, sticky="w")

            entry = CustomEntry(fields_grid_frame, field_name=field_key, examples_dict=EXAMPLES)
            entry.grid(row=i, column=1, padx=5, pady=3, sticky="ew")
            fields_grid_frame.columnconfigure(1, weight=1)

            saved_value = template_specific_memory.get(field_key, general_memory.get(field_key))
            if saved_value is not None:
                entry.set_text(saved_value)

            bind_entry_shortcuts(entry, main_context_menu)
            current_block_entries[field_key] = entry

            hint_text = EXAMPLES.get(field_key, f"Поле для заповнення: {field_key}")
            hint_button = ctk.CTkButton(fields_grid_frame, text="ℹ", width=28, height=28, font=("Arial", 14),
                                        command=lambda h=hint_text, f=field_key:
                                        messagebox.showinfo(f"Підказка для <{f}>", h))
            hint_button.grid(row=i, column=2, padx=(0, 5), pady=3, sticky="e")

        # Функция для обновления общих полей на основе данных таблицы товаров
        def update_summary_fields():
            if not products_widget:
                return

            total_sum = products_widget["update_total"]()

            # Обновляем поля общей суммы
            summary_fields = ["разом", "загальна сума", "всього"]
            for field_name in summary_fields:
                if field_name in current_block_entries:
                    entry = current_block_entries[field_name]
                    entry.configure(state="normal")
                    entry.set_text(f"{total_sum:.2f}")
                    entry.configure(state="readonly")

            # Обновляем поле "сума прописом"
            if "сума прописом" in current_block_entries:
                entry = current_block_entries["сума прописом"]
                entry.configure(state="normal")
                if total_sum > 0:
                    entry.set_text(number_to_ukrainian_text(total_sum).capitalize())
                else:
                    entry.set_text("")
                entry.configure(state="readonly")

        # Привязываем обновление сводных полей к изменениям в таблице товаров
        if products_widget:
            # Перехватываем функцию обновления таблицы
            original_update = products_widget["update_total"]

            def enhanced_update():
                result = original_update()
                update_summary_fields()
                return result

            products_widget["update_total"] = enhanced_update

        # Робимо readonly поля, які автоматично обчислюються
        readonly_fields = ["сума прописом", "разом", "загальна сума", "всього"]
        for key in readonly_fields:
            if key in current_block_entries:
                current_block_entries[key].configure(state="readonly")
                current_block_entries[key].configure(takefocus=False)
                current_block_entries[key].configure(fg_color=("gray90", "gray20"))

    # ПАНЕЛЬ ДЕЙСТВИЙ ДЛЯ БЛОКА
    block_actions_frame = ctk.CTkFrame(block_frame, fg_color="transparent")
    block_actions_frame.pack(fill="x", padx=5, pady=5)

    def clear_block_fields():
        if messagebox.askokcancel("Очистити поля", "Очистити всі поля цього договору?"):
            for field_key, entry in current_block_entries.items():
                if field_key not in COMMON_FIELDS:
                    entry.configure(state="normal")
                    entry.set_text("")

            # Очищаем таблицу товаров
            if products_widget:
                products_widget["set_data"]([])

    def replace_block_template():
        new_path = filedialog.askopenfilename(
            title="Оберіть новий шаблон",
            filetypes=[("Word Documents", "*.docm *.docx")]
        )
        if new_path:
            try:
                new_placeholders = enhanced_extract_placeholders_from_word(new_path)
                has_new_table = "<таблиця_товарів>" in open(new_path, 'rb').read().decode('utf-8', errors='ignore')

                if not new_placeholders and not has_new_table:
                    messagebox.showwarning("Увага", "У новому шаблоні не знайдено плейсхолдерів!")
                    return

                table_info = " + таблиця товарів" if has_new_table else ""
                path_label.configure(
                    text=f"Шаблон: {os.path.basename(new_path)} ({len(new_placeholders)} полів{table_info})")
                messagebox.showinfo("Увага", "Щоб застосувати новий шаблон, видаліть блок і створіть новий.")
            except Exception as e:
                messagebox.showerror("Помилка", f"Не вдалося прочитати новий шаблон:\n{e}")

    def remove_this_block():
        if messagebox.askokcancel("Видалити", "Видалити цей блок договору?"):
            if block_dict in document_blocks:
                document_blocks.remove(block_dict)
            block_frame.destroy()
            save_current_state(document_blocks, tabview)

    # Кнопки действий
    ctk.CTkButton(block_actions_frame, text="🧹 Очистити поля", command=clear_block_fields).pack(side="left", padx=3)
    ctk.CTkButton(block_actions_frame, text="🔄 Замінити шаблон", command=replace_block_template).pack(side="left",
                                                                                                      padx=3)

    # Информационная метка о полях
    fields_list = sorted(list(dynamic_fields))
    products_info_text = " + таблиця товарів" if has_products_table else ""
    info_text = f"Знайдено {len(dynamic_fields)} полів{products_info_text}: " + ", ".join(fields_list[:3])
    if len(dynamic_fields) > 3:
        info_text += f" та ще {len(dynamic_fields) - 3}..."

    info_label = ctk.CTkLabel(block_actions_frame, text=info_text, text_color="gray60", font=("Arial", 10))
    info_label.pack(side="left", padx=10)

    # Кнопка удаления в хедере
    remove_button = ctk.CTkButton(header_frame, text="🗑", width=28, height=28, fg_color="gray50", hover_color="gray40",
                                  command=remove_this_block)
    remove_button.pack(side="right", padx=(5, 0))

    # СОХРАНЕНИЕ БЛОКА с новой структурой данных
    all_entries = {}
    all_entries.update(global_entries)  # Добавляем глобальные поля
    all_entries.update(current_block_entries)  # Добавляем поля договора

    # Добавляем данные товаров в новом формате
    if products_widget:
        all_entries["products"] = products_widget["get_data"]()  # Список товаров
        all_entries["has_products_table"] = True
    else:
        all_entries["products"] = []
        all_entries["has_products_table"] = False

    # Получаем текущую активную вкладку
    current_tab_name = tabview.get()
    event_number = None

    # Получаем номер события из объекта
    try:
        # Извлекаем номер события из названия вкладки (например, "Захід 1" -> 1)
        if current_tab_name.startswith("Захід "):
            event_number = int(current_tab_name.split(" ")[1])
        else:
            event_number = 1  # По умолчанию
    except (ValueError, IndexError):
        event_number = 1

    # Создаем словарь блока документа
    block_dict = {
        "path": template_path,
        "entries": all_entries,
        "fields": list(dynamic_fields),
        "event_number": event_number,
        "tab_name": current_tab_name,
        "has_products_table": has_products_table
    }

    # Добавляем блок в глобальный список
    document_blocks.append(block_dict)

    # Сохраняем текущее состояние
    save_current_state(document_blocks, tabview)

    # Первоначальное обновление сводных полей если есть таблица товаров
    if products_widget:
        products_widget["update_total"]()

    messagebox.showinfo("Успіх",
                        f"Блок договору створено успішно!\n"
                        f"Шаблон: {os.path.basename(template_path)}\n"
                        f"Полів: {len(dynamic_fields)}\n"
                        f"Таблиця товарів: {'Так' if has_products_table else 'Ні'}")

    return block_dict


def get_all_block_data(tab_name):
    """Получает все данные блоков для определенной вкладки"""
    tab_blocks = []
    for block in document_blocks:
        if block.get("tab_name") == tab_name:
            # Обновляем данные товаров если есть products_widget
            current_data = {}
            current_data.update(block["entries"])

            # Если блок содержит таблицу товаров, обновляем данные
            if block.get("has_products_table", False):
                # Здесь можно добавить логику обновления данных таблицы товаров
                pass

            tab_blocks.append({
                "path": block["path"],
                "entries": current_data,
                "fields": block["fields"],
                "has_products_table": block.get("has_products_table", False)
            })

    return tab_blocks


def clear_all_blocks_data(tab_name):
    """Очищает данные всех блоков для определенной вкладки"""
    cleared_count = 0
    for block in document_blocks:
        if block.get("tab_name") == tab_name:
            # Очищаем поля (кроме общих)
            for field_key in block["entries"]:
                if field_key not in COMMON_FIELDS and field_key != "products":
                    block["entries"][field_key] = ""

            # Очищаем данные товаров
            if "products" in block["entries"]:
                block["entries"]["products"] = []

            cleared_count += 1

    return cleared_count


def update_block_products_data(block_dict, products_data):
    """Обновляет данные товаров в блоке"""
    if "entries" not in block_dict:
        block_dict["entries"] = {}

    block_dict["entries"]["products"] = products_data
    block_dict["entries"]["has_products_table"] = len(products_data) > 0


def get_block_products_data(block_dict):
    """Получает данные товаров из блока"""
    if "entries" in block_dict and "products" in block_dict["entries"]:
        return block_dict["entries"]["products"]
    return []


def validate_block_data(block_dict):
    """Проверяет корректность данных блока"""
    errors = []

    if not block_dict.get("path"):
        errors.append("Не вказано шлях до шаблону")

    if not os.path.exists(block_dict.get("path", "")):
        errors.append("Файл шаблону не існує")

    if not block_dict.get("entries"):
        errors.append("Відсутні дані полів")

    # Проверяем обязательные поля
    required_fields = ["контрагент", "послуга", "сума"]  # Пример обязательных полей
    for field in required_fields:
        if field in block_dict.get("fields", []):
            if not block_dict.get("entries", {}).get(field):
                errors.append(f"Не заповнено обов'язкове поле: {field}")

    return errors


def export_block_data_to_json(block_dict, file_path):
    """Экспортирует данные блока в JSON файл"""
    try:
        export_data = {
            "template_path": block_dict.get("path"),
            "fields": block_dict.get("fields", []),
            "entries": block_dict.get("entries", {}),
            "has_products_table": block_dict.get("has_products_table", False),
            "event_number": block_dict.get("event_number"),
            "tab_name": block_dict.get("tab_name"),
            "export_date": str(os.path.getctime(file_path)) if os.path.exists(file_path) else ""
        }

        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(export_data, f, ensure_ascii=False, indent=2)

        return True
    except Exception as e:
        print(f"Помилка експорту: {e}")
        return False


def import_block_data_from_json(file_path):
    """Импортирует данные блока из JSON файла"""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            import_data = json.load(f)

        # Проверяем существование шаблона
        template_path = import_data.get("template_path")
        if not template_path or not os.path.exists(template_path):
            return None, "Шаблон не знайдено"

        block_dict = {
            "path": template_path,
            "entries": import_data.get("entries", {}),
            "fields": import_data.get("fields", []),
            "has_products_table": import_data.get("has_products_table", False),
            "event_number": import_data.get("event_number", 1),
            "tab_name": import_data.get("tab_name", "Захід 1")
        }

        return block_dict, None
    except Exception as e:
        return None, f"Помилка імпорту: {e}"


def duplicate_block(original_block_dict):
    """Создает копию блока с очищенными данными"""
    new_block = {
        "path": original_block_dict["path"],
        "entries": {},
        "fields": original_block_dict["fields"].copy(),
        "has_products_table": original_block_dict.get("has_products_table", False),
        "event_number": original_block_dict.get("event_number", 1),
        "tab_name": original_block_dict.get("tab_name", "Захід 1")
    }

    # Копируем только общие поля
    for field_key, value in original_block_dict["entries"].items():
        if field_key in COMMON_FIELDS:
            new_block["entries"][field_key] = value
        else:
            new_block["entries"][field_key] = ""

    # Очищаем данные товаров
    if "products" in new_block["entries"]:
        new_block["entries"]["products"] = []

    return new_block


def get_blocks_summary(tab_name):
    """Возвращает сводку по блокам для вкладки"""
    tab_blocks = [block for block in document_blocks if block.get("tab_name") == tab_name]

    summary = {
        "total_blocks": len(tab_blocks),
        "templates_used": len(set(block["path"] for block in tab_blocks)),
        "blocks_with_products": sum(1 for block in tab_blocks if block.get("has_products_table", False)),
        "total_fields": sum(len(block.get("fields", [])) for block in tab_blocks)
    }

    return summary